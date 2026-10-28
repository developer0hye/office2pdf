"""Tests for the issue loop's red-main rule.

From 2026-10-16 the `cargo-deny` advisories check failed on `main` itself, so
every pull request inherited the failure. Unattended runs kept merging past it
as "pre-existing", and eight pull requests landed with a failing check before
anyone fixed main. The loop now works the issue that tracks a red main before
anything else (filing one when none is open), and its instructions forbid
merging over any failing check, inherited or not.

The loop's issue selection runs through `--print-next-issue` against a fake
`gh` on PATH, so no network or credentials are involved.
"""

import json
import os
import subprocess
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
LOOP = SCRIPTS / "issue_loop.sh"
PROMPT = SCRIPTS / "issue_loop_prompt.md"

FAKE_GH = r'''#!/usr/bin/env python3
import json, os, sys
args = sys.argv[1:]
with open(os.environ["FAKE_GH_LOG"], "a") as log:
    log.write(json.dumps(args) + "\n")
joined = " ".join(args)
if args[:2] == ["auth", "status"]:
    pass
elif args[:2] == ["api", "user"]:
    print(os.environ.get("FAKE_LOGIN", "developer0hye"))
elif args[:2] == ["repo", "view"]:
    print("developer0hye")
elif args[:2] == ["auth", "switch"]:
    sys.exit(1)
elif args[:1] == ["api"] and "commits/main" in joined:
    if os.environ.get("FAKE_MAIN_ERROR"): sys.exit(1)
    print("0123abc")
elif args[:2] == ["run", "list"]:
    print(os.environ.get("FAKE_FAILURES", ""))
elif args[:2] == ["issue", "list"] and "main-red" in joined:
    want_unblocked = "-label:autofix-blocked" in joined
    key = "FAKE_MAIN_RED_UNBLOCKED" if want_unblocked else "FAKE_MAIN_RED_ANY"
    print(os.environ.get(key, ""))
elif args[:2] == ["issue", "list"]:
    print(os.environ.get("FAKE_QUEUE", ""))
elif args[:2] == ["issue", "create"]:
    print("https://github.com/owner/repo/issues/" + os.environ.get("FAKE_CREATED", "9999"))
'''


class MainRedTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        bin_dir = root / "bin"
        bin_dir.mkdir()
        gh = bin_dir / "gh"
        gh.write_text(FAKE_GH, encoding="utf-8")
        gh.chmod(0o755)
        self.log = root / "gh.log"
        self.log.touch()
        self.env = {
            **os.environ,
            "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}",
            "FAKE_GH_LOG": str(self.log),
            "FAKE_QUEUE": "1721",
        }
        self.logs = root / "logs"

    def tearDown(self):
        self.tmp.cleanup()

    def next_issue(self, *extra: str, **fake: str) -> str:
        done = subprocess.run(
            ["bash", str(LOOP), "--log-dir", str(self.logs), *extra, "--print-next-issue"],
            capture_output=True,
            text=True,
            env={**self.env, **fake},
        )
        self.assertEqual(done.returncode, 0, done.stderr)
        return done.stdout.strip()

    def gh_calls(self) -> list[list[str]]:
        return [json.loads(line) for line in self.log.read_text().splitlines()]

    def created(self) -> list[list[str]]:
        return [call for call in self.gh_calls() if call[:2] == ["issue", "create"]]

    def test_a_green_main_leaves_the_queue_order_alone(self):
        self.assertEqual(self.next_issue(), "1721")
        self.assertEqual(self.created(), [])

    def test_an_open_red_main_issue_is_worked_first(self):
        picked = self.next_issue(FAKE_FAILURES="cargo-deny", FAKE_MAIN_RED_ANY="2179",
                                 FAKE_MAIN_RED_UNBLOCKED="2179")
        self.assertEqual(picked, "2179")
        self.assertEqual(self.created(), [])

    def test_a_red_main_without_an_issue_gets_one_filed_and_worked(self):
        picked = self.next_issue(FAKE_FAILURES="cargo-deny, CI", FAKE_CREATED="2300")

        self.assertEqual(picked, "2300")
        (create,) = self.created()
        self.assertIn("main-red", create)
        title = create[create.index("--title") + 1]
        self.assertIn("cargo-deny, CI", title)
        # The label must exist before it is applied, even on a repository that
        # has never had a red main.
        calls = self.gh_calls()
        label = next(i for i, c in enumerate(calls) if c[:3] == ["label", "create", "main-red"])
        self.assertLess(label, calls.index(create))

    def test_a_red_main_outranks_a_queue_label(self):
        picked = self.next_issue("--label", "priority", FAKE_FAILURES="cargo-deny",
                                 FAKE_MAIN_RED_ANY="2179", FAKE_MAIN_RED_UNBLOCKED="2179")
        self.assertEqual(picked, "2179")

    def test_a_blocked_red_main_issue_is_not_refiled_or_retried(self):
        # Two failed attempts already labelled it blocked: a duplicate would just
        # be retried forever, so the loop goes back to its queue.
        picked = self.next_issue(FAKE_FAILURES="cargo-deny", FAKE_MAIN_RED_ANY="2179")
        self.assertEqual(picked, "1721")
        self.assertEqual(self.created(), [])

    def test_diagnostic_checks_identity_before_each_main_red_write(self):
        self.next_issue(FAKE_FAILURES="CI")
        since = []
        for call in self.gh_calls():
            if call[:2] in (["label", "create"], ["issue", "create"]):
                self.assertIn(["auth", "status"], since)
                self.assertIn(["api", "user", "--jq", ".login"], since)
                since = []
            else:
                since.append(call)

    def test_diagnostic_cannot_publish_with_wrong_account(self):
        done = subprocess.run(
            ["bash", str(LOOP), "--log-dir", str(self.logs), "--print-next-issue"],
            capture_output=True, text=True,
            env={**self.env, "FAKE_FAILURES": "CI", "FAKE_LOGIN": "someone-else"},
        )
        self.assertNotEqual(done.returncode, 0)
        self.assertEqual(self.created(), [])
        self.assertFalse(any(c[:2] == ["label", "create"] for c in self.gh_calls()))

    def test_diagnostic_propagates_main_api_failure(self):
        done = subprocess.run(
            ["bash", str(LOOP), "--log-dir", str(self.logs), "--print-next-issue"],
            capture_output=True, text=True, env={**self.env, "FAKE_MAIN_ERROR": "1"},
        )
        self.assertNotEqual(done.returncode, 0)
        self.assertEqual(done.stdout, "")
        self.assertEqual(self.created(), [])

    def test_the_instructions_forbid_merging_over_an_inherited_failure(self):
        text = " ".join(PROMPT.read_text(encoding="utf-8").split())
        self.assertIn("every check", text)
        self.assertIn("including one that also fails on `main`", text)


if __name__ == "__main__":
    unittest.main()
