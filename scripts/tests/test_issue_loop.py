"""Tests for the unattended issue loop's usage-limit detector and its prompt.

A false positive here is expensive in a way only a test cheaply guards: the loop
discards the run it just finished and sleeps an hour before retrying, so one
over-broad substring costs an issue's worth of unattended progress. On
2026-09-08 the detector grepped the raw run log and matched the CLI's own
telemetry — a `rate_limit_event` whose status was `allowed_warning`, and the
init event's list of slash commands, which contains `usage-credits`. It threw
away a successful 79-turn run while the seven-day usage window sat at 58%.

The detector is exercised through `--usage-limit-reason`, the same entry point
the loop itself calls, so these tests survive a rewrite of how it is implemented.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPT = Path(__file__).resolve().parents[1] / "issue_loop.sh"


def init_event() -> str:
    """The CLI's session header. Its slash-command list contains `usage-credits`."""
    return json.dumps(
        {
            "type": "system",
            "subtype": "init",
            "model": "claude-opus-5",
            "slash_commands": ["security-review", "usage-credits", "extra-usage", "usage"],
        }
    )


def rate_limit_event(status: str, utilization: float = 0.58) -> str:
    return json.dumps(
        {
            "type": "rate_limit_event",
            "rate_limit_info": {
                "status": status,
                "rateLimitType": "seven_day",
                "utilization": utilization,
                "isUsingOverage": False,
            },
        }
    )


def result_event(text: str, subtype: str = "success") -> str:
    return json.dumps(
        {
            "type": "result",
            "subtype": subtype,
            "is_error": False,
            "num_turns": 79,
            "result": text,
        }
    )


def refused_report(text: str, api_error: str, status: int) -> str:
    """A run the API refused before any work: the CLI's own structured verdict."""
    return json.dumps(
        {
            "type": "result",
            "subtype": "success",
            "is_error": True,
            "num_turns": 1,
            "total_cost_usd": 0,
            "terminal_reason": "api_error",
            "api_error_status": status,
            "api_error": api_error,
            "result": text,
        }
    )


class UsageLimitReasonTest(unittest.TestCase):
    def detect(self, log_text: str) -> subprocess.CompletedProcess:
        with tempfile.TemporaryDirectory() as tmp:
            log = Path(tmp) / "issue-1381-20260908-203328.json"
            log.write_text(log_text, encoding="utf-8")
            return self.run_detector(str(log))

    def run_detector(self, path: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            ["bash", str(SCRIPT), "--usage-limit-reason", path],
            capture_output=True,
            text=True,
        )

    def assertNoLimit(self, done: subprocess.CompletedProcess) -> None:
        self.assertEqual(done.stdout.strip(), "", done.stderr)
        self.assertEqual(done.returncode, 1, done.stderr)

    def assertLimit(self, done: subprocess.CompletedProcess) -> None:
        self.assertNotEqual(done.stdout.strip(), "", done.stderr)
        self.assertEqual(done.returncode, 0, done.stderr)

    def test_telemetry_of_a_successful_verbose_run_is_not_a_limit(self):
        # The regression: every line here is normal --verbose stream-json output.
        log = "\n".join(
            [
                init_event(),
                rate_limit_event("allowed_warning"),
                rate_limit_event("allowed"),
                result_event("Merged PR #1661 and closed the issue."),
            ]
        )
        self.assertNoLimit(self.detect(log))

    def test_a_rejected_rate_limit_event_is_a_limit(self):
        log = "\n".join(
            [
                init_event(),
                rate_limit_event("allowed_warning"),
                rate_limit_event("rejected", utilization=1.0),
                result_event("You've reached your Fable limit."),
            ]
        )
        self.assertLimit(self.detect(log))

    def test_a_refused_run_reports_the_limit_from_its_result_text(self):
        # --output-format json emits one object and no telemetry, so the message
        # itself is the only evidence the loop gets.
        self.assertLimit(
            self.detect(
                result_event(
                    "You've reached your Fable limit. Switch to another model, or "
                    "manage usage credits at claude.ai/settings/usage"
                    "?from=cc_cli_limit_message, to continue."
                )
            )
        )

    def test_the_clis_structured_usage_limit_verdict_is_a_limit_whatever_its_wording(self):
        # 2026-10-19: "You've hit your weekly limit" matched none of the message
        # patterns, so three refused runs counted as failures, two issues were
        # wrongly labelled blocked, and the circuit breaker stopped the loop.
        self.assertLimit(
            self.detect(
                refused_report(
                    "You've hit your weekly limit · resets 2pm (Asia/Seoul)",
                    "usage_limit_reached",
                    429,
                )
            )
        )

    def test_another_api_error_is_not_a_usage_limit(self):
        # Triangulation: an overloaded API also refuses the run, but sleeping an
        # hour for it would waste the window; only the usage verdict counts.
        self.assertNoLimit(
            self.detect(refused_report("API Error: Overloaded", "overloaded_error", 529))
        )

    def test_a_plain_text_cli_failure_is_a_limit(self):
        # A hard stop can kill the CLI before it emits any JSON at all.
        self.assertLimit(self.detect("Claude AI usage limit reached|1789362000\n"))

    def test_an_agents_prose_about_rate_limiting_is_not_a_limit(self):
        # Triangulation: an agent that fixes retry code writes these words into
        # its own summary. Both phrases matched the pre-2026-09-08 pattern.
        self.assertNoLimit(
            self.detect(
                "\n".join(
                    [
                        init_event(),
                        result_event(
                            "Added backoff for the API's rate limit and documented "
                            "/usage-credits in the runbook."
                        ),
                    ]
                )
            )
        )

    def test_codex_quota_failure_is_a_limit(self):
        self.assertLimit(self.detect(json.dumps({"type": "turn.failed", "error": {
            "message": "You've hit your usage limit. Try again later."}})))

    def test_codex_error_event_is_a_limit(self):
        self.assertLimit(self.detect(json.dumps({"type": "error", "message": "rate_limit_exceeded"})))

    def test_codex_agent_and_command_text_cannot_claim_a_quota_failure(self):
        for item in (
            {"type": "agent_message", "text": "You've hit your usage limit."},
            {"type": "command_execution", "aggregated_output": "rate_limit_exceeded"},
        ):
            self.assertNoLimit(self.detect(json.dumps({"type": "item.completed", "item": item})))

    def test_an_empty_log_is_not_a_limit(self):
        self.assertNoLimit(self.detect(""))

    def test_a_missing_log_is_not_a_limit(self):
        # The loop must not die on a run that never opened its log file.
        with tempfile.TemporaryDirectory() as tmp:
            self.assertNoLimit(self.run_detector(str(Path(tmp) / "absent.json")))


class HostNoteTest(unittest.TestCase):
    """The prompt says when the host cannot export a native Office PDF.

    Seven runs on 2026-10-17/18 (#1723, #1728, #1731, #1736, #1738, #1748,
    #1761) each spent turns rediscovering that a Linux host has no native
    Office export of its own: disk-wide `find` for PDFs, precedent drafts, advisor calls. A Mac
    can export one, so the note must depend on the host, not the prompt file.
    """

    NOTE = "cannot run Microsoft Office"

    def prompt_on(self, system: str) -> str:
        with tempfile.TemporaryDirectory() as tmp:
            bin_dir = Path(tmp) / "bin"
            bin_dir.mkdir()
            uname = bin_dir / "uname"
            uname.write_text("#!/bin/sh\necho %s\n" % system, encoding="utf-8")
            uname.chmod(0o755)
            env = dict(os.environ, PATH="%s%s%s" % (bin_dir, os.pathsep, os.environ["PATH"]))
            done = subprocess.run(
                ["bash", str(SCRIPT), "--log-dir", str(Path(tmp) / "logs"), "--print-prompt", "1723"],
                capture_output=True,
                text=True,
                env=env,
            )
        self.assertEqual(done.returncode, 0, done.stderr)
        return done.stdout

    def test_a_linux_host_is_told_it_has_no_native_office(self):
        prompt = self.prompt_on("Linux")
        self.assertIn(self.NOTE, prompt)
        self.assertIn("runs Linux", prompt)

    def test_a_linux_host_may_use_an_external_reference_corpus_the_issue_names(self):
        # docxide-pdf's corpus keeps its Word reference PDFs outside this
        # checkout; DOCX issues filed from it must still be able to use them.
        self.assertIn("reference corpus the issue names", self.prompt_on("Linux"))

    def test_any_host_but_macos_is_told_it_has_no_native_office(self):
        self.assertIn("runs FreeBSD", self.prompt_on("FreeBSD"))

    def test_a_macos_host_is_not_told_it_lacks_office(self):
        self.assertNotIn(self.NOTE, self.prompt_on("Darwin"))


if __name__ == "__main__":
    unittest.main()
