"""Tests for the repository-specific instructions in the unattended issue loop's prompt.

The prompt is read through `--print-prompt`, the same text a run receives, so
these tests survive a change to how `issue_loop.sh` assembles it.
"""

import re
import subprocess
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SCRIPT = ROOT / "scripts" / "issue_loop.sh"
CI = ROOT / ".github" / "workflows" / "ci.yml"


def prompt(issue: str) -> str:
    with tempfile.TemporaryDirectory() as tmp:
        done = subprocess.run(
            ["bash", str(SCRIPT), "--log-dir", str(Path(tmp) / "logs"), "--print-prompt", issue],
            capture_output=True,
            text=True,
        )
    if done.returncode != 0:
        raise AssertionError(done.stderr)
    return done.stdout


def ci_job(name: str) -> str:
    """One job's block of ci.yml: from its key to the next job key at the same indent."""
    text = CI.read_text(encoding="utf-8")
    start = re.search(r"^  %s:\n" % re.escape(name), text, re.MULTILINE)
    if start is None:
        raise AssertionError("ci.yml has no %s job" % name)
    end = re.search(r"^  [A-Za-z0-9_-]+:\n", text[start.end():], re.MULTILINE)
    return text[start.start(): start.end() + end.start() if end else len(text)]


class WasmTestBuildNoteTest(unittest.TestCase):
    """The prompt names the local check that compiles the unit tests for wasm32.

    Runs 20261019-210859 (issue 2224) and 20261021-014040 (issue 2236) each
    pushed a test that used a native-only helper. The library build passed on
    wasm32, but CI's WASM job compiles the unit tests too, and both of its legs
    failed. Each run spent 9-11 tool calls and another CI round to find and fix
    it; the first then hit the turn cap before merging, and the second hit a
    usage limit before CI finished. This check reproduces both CI errors
    locally in about 15 seconds, the plain check passes both, and the fix
    commits pass it.
    """

    CHECK = "cargo check --locked --tests --target wasm32-unknown-unknown -p office2pdf --features wasm"

    def test_the_prompt_names_the_wasm32_test_build_check(self):
        self.assertIn(self.CHECK, prompt("2224"))

    def test_ci_still_compiles_the_unit_tests_for_that_feature(self):
        # The note is only true while CI's WASM job builds the tests with the
        # feature the check names; if either changes, the note is stale.
        job = ci_job("wasm-check")
        self.assertIn("wasm-pack test", job)
        matrix = re.search(r"feature:\s*\[([^\]]*)\]", job)
        self.assertIsNotNone(matrix, job)
        self.assertIn("wasm", [leg.strip() for leg in matrix.group(1).split(",")])


if __name__ == "__main__":
    unittest.main()
