"""Tests for the unattended issue loop's usage-limit detector.

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


if __name__ == "__main__":
    unittest.main()
