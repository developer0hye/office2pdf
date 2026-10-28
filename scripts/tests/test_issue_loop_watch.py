"""Tests for the issue-loop transcript viewer.

The two pure functions are what a stale assumption breaks silently: a changed
transcript-directory naming scheme makes the viewer watch nothing, and a
malformed transcript line must never take the viewer down mid-run.
"""

import io
import json
import os
import sys
import tempfile
import time
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from issue_loop_watch import format_events, newest_transcript, transcript_dir


def assistant_line(*blocks: dict) -> str:
    return json.dumps({"type": "assistant", "message": {"content": list(blocks)}})


class TranscriptDirTest(unittest.TestCase):
    def test_project_path_becomes_a_hyphenated_directory_name(self):
        self.assertEqual(
            transcript_dir("/Users/dev/Documents/office2pdf", home="/Users/dev"),
            "/Users/dev/.claude/projects/-Users-dev-Documents-office2pdf",
        )

    def test_a_nested_checkout_keeps_every_path_segment(self):
        # Worktrees differ only in the trailing segment, so truncating any part
        # of the path would collide two checkouts onto one transcript directory.
        self.assertNotEqual(
            transcript_dir("/src/repo", home="/h"),
            transcript_dir("/src/repo-fix", home="/h"),
        )


class NewestTranscriptTest(unittest.TestCase):
    def test_finds_live_codex_events_in_the_durable_run_layout(self):
        with tempfile.TemporaryDirectory() as directory:
            run = Path(directory) / "runs/issue-2101/attempt-codex"
            run.mkdir(parents=True)
            events = run / "events.jsonl"
            events.write_text('{"type":"thread.started","thread_id":"codex-session"}\n')
            copied = run / "transcript.jsonl"
            copied.write_text(events.read_text())
            self.assertEqual(newest_transcript(directory), str(events))

    def test_picks_the_most_recently_written_session(self):
        with tempfile.TemporaryDirectory() as directory:
            older = os.path.join(directory, "older.jsonl")
            newer = os.path.join(directory, "newer.jsonl")
            Path(older).write_text("{}\n", encoding="utf-8")
            time.sleep(0.01)
            Path(newer).write_text("{}\n", encoding="utf-8")
            self.assertEqual(newest_transcript(directory), newer)

    def test_excluded_sessions_are_never_selected(self):
        with tempfile.TemporaryDirectory() as directory:
            wanted = os.path.join(directory, "agent.jsonl")
            Path(wanted).write_text("{}\n", encoding="utf-8")
            time.sleep(0.01)
            skipped = os.path.join(directory, "mine.jsonl")
            Path(skipped).write_text("{}\n", encoding="utf-8")
            self.assertEqual(newest_transcript(directory, exclude=("mine",)), wanted)

    def test_an_empty_directory_yields_no_transcript(self):
        with tempfile.TemporaryDirectory() as directory:
            self.assertIsNone(newest_transcript(directory))

    def test_excludes_codex_thread_id_even_when_not_in_event_path(self):
        with tempfile.TemporaryDirectory() as directory:
            run = Path(directory) / 'runs/reflect-attempt'
            run.mkdir(parents=True)
            events = run / 'events.jsonl'
            events.write_text('{"type":"thread.started","thread_id":"excluded-session"}\n')
            self.assertIsNone(newest_transcript(directory, exclude=('excluded-session',)))


class FormatEventsTest(unittest.TestCase):
    def test_a_tool_call_reports_its_name_and_leading_argument(self):
        line = assistant_line(
            {"type": "tool_use", "name": "Bash", "input": {"command": "cargo test"}}
        )
        self.assertEqual(format_events(line), ["[TOOL] Bash: cargo test"])

    def test_a_tool_call_without_a_known_argument_still_reports_its_name(self):
        line = assistant_line({"type": "tool_use", "name": "ToolSearch", "input": {}})
        self.assertEqual(format_events(line), ["[TOOL] ToolSearch: "])

    def test_assistant_text_is_flattened_to_one_line(self):
        line = assistant_line({"type": "text", "text": "first\nsecond"})
        self.assertEqual(format_events(line), ["[SAY]  first second"])

    def test_long_text_is_truncated(self):
        line = assistant_line({"type": "text", "text": "x" * 500})
        self.assertEqual(len(format_events(line)[0]), len("[SAY]  ") + 300)

    def test_blank_text_is_not_an_event(self):
        self.assertEqual(format_events(assistant_line({"type": "text", "text": "  "})), [])

    def test_one_line_can_carry_several_events(self):
        line = assistant_line(
            {"type": "text", "text": "running the suite"},
            {"type": "tool_use", "name": "Bash", "input": {"command": "cargo test"}},
        )
        self.assertEqual(len(format_events(line)), 2)

    def test_codex_command_start_and_completion_are_visible(self):
        for event in ('item.started', 'item.completed'):
            line = json.dumps({'type': event, 'item': {'type': 'command_execution',
                'command': 'cargo test --locked', 'exit_code': 0}})
            self.assertIn('cargo test --locked', format_events(line)[0])

    def test_codex_final_message_is_visible(self):
        line = json.dumps({'type': 'item.completed', 'item': {'type': 'agent_message', 'text': 'Fixed #2101'}})
        self.assertEqual(format_events(line), ['[SAY]  Fixed #2101'])

    def test_codex_failure_is_visible(self):
        line = json.dumps({'type': 'turn.failed', 'error': {'message': 'Quota exhausted'}})
        self.assertEqual(format_events(line), ['[ERROR] Quota exhausted'])

    def test_torn_line_is_retried_when_completed(self):
        from issue_loop_watch import read_complete_lines
        stream = io.StringIO('first\n{"type":')
        self.assertEqual(read_complete_lines(stream), ['first\n'])
        position = stream.tell()
        stream.seek(0, 2)
        stream.write('"turn.completed"}\n')
        stream.seek(position)
        self.assertEqual(read_complete_lines(stream), ['{"type":"turn.completed"}\n'])

    def test_non_assistant_records_are_ignored(self):
        self.assertEqual(format_events(json.dumps({"type": "user"})), [])

    def test_a_malformed_line_is_ignored_rather_than_raising(self):
        # A transcript is appended to while being read, so a torn final line is
        # normal; the viewer must survive it.
        self.assertEqual(format_events('{"type": "assist'), [])
        self.assertEqual(format_events(""), [])


if __name__ == "__main__":
    unittest.main()
