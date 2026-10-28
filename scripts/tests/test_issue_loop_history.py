"""Tests for the issue loop's experience store.

Each `claude -p` run starts with an empty context, so a second attempt at an
issue used to repeat the first attempt's dead ends: the loop kept only the final
JSON report, and nothing pointed the next agent at it. Meta-Harness
(arXiv:2603.28052) found that a proposer given the raw traces of earlier
candidates on a filesystem beats one given scores or summaries. These tests pin
that the store keeps each run's full transcript beside its outcome, and that the
next prompt for the same issue points at them.

Everything runs through the command line the loop itself calls.
"""

import json
import os
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
HISTORY = SCRIPTS / "issue_loop_history.py"
LOOP = SCRIPTS / "issue_loop.sh"
FAKE_REPO = "/srv/checkout/office2pdf"


def result_report(session_id: str, text: str, turns: int = 41) -> str:
    """What `claude -p --output-format json` writes when the run ends."""
    return json.dumps(
        {
            "type": "result",
            "subtype": "success",
            "is_error": False,
            "num_turns": turns,
            "duration_ms": 1_834_000,
            "total_cost_usd": 12.5,
            "session_id": session_id,
            "result": text,
        }
    )


class StoreTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        self.home = self.root / "home"
        self.logs = self.root / "logs"
        self.transcripts = self.home / ".claude" / "projects" / FAKE_REPO.replace("/", "-")
        self.transcripts.mkdir(parents=True)

    def tearDown(self):
        self.tmp.cleanup()

    def run_tool(self, *args: str, script: Path = HISTORY) -> subprocess.CompletedProcess:
        command = ["bash", str(script)] if script.suffix == ".sh" else [sys.executable, str(script)]
        return subprocess.run(
            command + list(args),
            capture_output=True,
            text=True,
            env={**os.environ, "HOME": str(self.home)},
        )

    def attempt(self, issue: int, stamp: str, report: str, outcome: str, rc: int = 0) -> Path:
        run_dir = self.logs / "runs" / f"issue-{issue}" / stamp
        run_dir.mkdir(parents=True)
        (run_dir / "prompt.md").write_text(f"Your task is GitHub issue #{issue}.\n")
        (run_dir / "result.json").write_text(report)
        done = self.run_tool(
            "record", "--log-dir", str(self.logs), "--repo-root", FAKE_REPO,
            "--run-dir", str(run_dir), "--issue", str(issue),
            "--outcome", outcome, "--rc", str(rc),
        )
        self.assertEqual(done.returncode, 0, done.stderr)
        return run_dir

    def context(self, issue: int, merged_refs: int = 0) -> str:
        done = self.run_tool(
            "context", "--log-dir", str(self.logs), "--issue", str(issue),
            "--merged-refs", str(merged_refs),
        )
        self.assertEqual(done.returncode, 0, done.stderr)
        return done.stdout

    def test_record_keeps_the_transcript_beside_the_outcome(self):
        lines = [json.dumps({"type": "assistant", "message": {"content": [
            {"type": "text", "text": "The fixture renders the row at 15pt."}]}})]
        (self.transcripts / "5f1c.jsonl").write_text("\n".join(lines) + "\n")

        run_dir = self.attempt(1720, "20261017-220646",
                               result_report("5f1c", "Opened PR #2180; CI still running."),
                               "pr-open")

        outcome = json.loads((run_dir / "outcome.json").read_text())
        self.assertEqual(outcome["issue"], 1720)
        self.assertEqual(outcome["outcome"], "pr-open")
        self.assertEqual(outcome["session_id"], "5f1c")
        self.assertEqual(outcome["num_turns"], 41)
        self.assertEqual(outcome["total_cost_usd"], 12.5)
        self.assertEqual(outcome["final_message"], "Opened PR #2180; CI still running.")
        self.assertEqual((run_dir / "transcript.jsonl").read_text().splitlines(), lines)
        index = [json.loads(line) for line in
                 (self.logs / "runs" / "index.jsonl").read_text().splitlines()]
        self.assertEqual([(row["issue"], row["outcome"]) for row in index], [(1720, "pr-open")])

    def test_codex_record_preserves_exact_session_and_full_events_for_retry(self):
        run_dir = self.logs / "runs/issue-2101/attempt-codex"
        run_dir.mkdir(parents=True)
        report = {"issue": 2101, "status": "incomplete", "pr_number": None,
                  "summary": "The integration needs a history repair.", "remaining": ["Save events."]}
        (run_dir / "result.json").write_text(json.dumps(report))
        events = "\n".join(json.dumps(x) for x in [
            {"type": "thread.started", "thread_id": "exact-codex-session"},
            {"type": "item.completed", "item": {"type": "command_execution", "command": "cargo test --locked"}},
            {"type": "turn.failed", "error": {"message": "Interrupted after testing."}},
        ]) + "\n"
        (run_dir / "events.jsonl").write_text(events)
        done = self.run_tool("record", "--log-dir", str(self.logs), "--repo-root", FAKE_REPO,
                             "--run-dir", str(run_dir), "--issue", "2101", "--agent", "codex",
                             "--outcome", "no-progress", "--rc", "1")
        self.assertEqual(done.returncode, 0, done.stderr)
        outcome = json.loads((run_dir / "outcome.json").read_text())
        self.assertEqual(outcome["agent"], "codex")
        self.assertEqual(outcome["session_id"], "exact-codex-session")
        self.assertEqual(outcome["final_message"], report["summary"])
        self.assertEqual((run_dir / "transcript.jsonl").read_text(), events)
        context = self.context(2101)
        self.assertIn(report["summary"], context)
        self.assertIn(str(run_dir / "transcript.jsonl"), context)

    def test_a_crashed_run_without_json_is_still_recorded(self):
        # A run killed before it printed its report leaves plain text; the attempt
        # still counts, and the next agent should see that it died.
        run_dir = self.attempt(1723, "20261017-230000", "Error: socket hang up\n",
                               "no-progress", rc=1)

        outcome = json.loads((run_dir / "outcome.json").read_text())
        self.assertIsNone(outcome["session_id"])
        self.assertEqual(outcome["rc"], 1)
        self.assertIn("socket hang up", outcome["final_message"])
        self.assertFalse((run_dir / "transcript.jsonl").exists())

    def test_context_lists_every_earlier_attempt_at_the_same_issue_oldest_first(self):
        (self.transcripts / "a1.jsonl").write_text("{}\n")
        first = self.attempt(1721, "20261016-090000",
                             result_report("a1", "Stopped: the probe control failed."), "no-progress")
        second = self.attempt(1721, "20261017-100000",
                              result_report("b2", "PR #2190 open, awaiting CI."), "pr-open")
        self.attempt(1730, "20261017-110000", result_report("c3", "Closed #1730."), "closed")

        text = self.context(1721)

        self.assertLess(text.index(str(first)), text.index(str(second)))
        self.assertIn("Stopped: the probe control failed.", text)
        self.assertIn("no-progress", text)
        self.assertIn(str(first / "transcript.jsonl"), text)
        self.assertNotIn("Closed #1730.", text)

    def test_context_is_empty_for_a_first_attempt_with_no_merged_reference(self):
        self.attempt(1730, "20261017-110000", result_report("c3", "Closed #1730."), "closed")
        self.assertEqual(self.context(1748), "")

    def test_context_flags_an_issue_a_merged_pull_request_already_cites(self):
        # A duplicate's fix often lands first; the agent should check that before
        # planning new work instead of rediscovering it.
        text = self.context(1748, merged_refs=2)
        self.assertIn("2 merged pull requests already cite #1748", text)

    def test_the_loop_prompt_carries_the_context(self):
        self.attempt(1721, "20261016-090000",
                     result_report("a1", "Stopped: the probe control failed."), "no-progress")

        done = self.run_tool("--print-prompt", "1721", "--log-dir", str(self.logs), script=LOOP)

        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn("Your task is GitHub issue #1721", done.stdout)
        self.assertIn("Stopped: the probe control failed.", done.stdout)

    def test_the_loop_prompt_reads_the_issue_through_rest(self):
        # `gh issue view` without --json asks GraphQL for classic project cards,
        # which GitHub now rejects; the #1720 trial lost turns to that error.
        done = self.run_tool("--print-prompt", "1720", "--log-dir", str(self.logs), script=LOOP)

        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertNotIn("gh issue view", done.stdout)
        self.assertIn("gh api repos/{owner}/{repo}/issues/1720/comments", done.stdout)

    def test_list_summarizes_runs_newest_last(self):
        self.attempt(1721, "20261016-090000", result_report("a1", "x"), "no-progress")
        self.attempt(1730, "20261017-110000", result_report("c3", "y", turns=7), "closed")

        done = self.run_tool("list", "--log-dir", str(self.logs))

        self.assertEqual(done.returncode, 0, done.stderr)
        rows = done.stdout.strip().splitlines()[1:]
        self.assertEqual([row.split()[1] for row in rows], ["#1721", "#1730"])
        self.assertIn("closed", rows[1])


if __name__ == "__main__":
    unittest.main()
