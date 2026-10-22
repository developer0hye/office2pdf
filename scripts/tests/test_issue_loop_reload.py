"""The issue loop picks up a new version of itself between runs.

Bash parses the loop body and its functions once, so a long-lived loop keeps
running the code it started with. The loop that started on 2026-10-19 14:11
still counted progress over the newest 30 merged pull requests on 2026-10-21,
although cc33dda4 had replaced that count and the main worktree had held the
fixed file since 2026-10-20 02:12. Run issue-2229/20261021-065200-claude-31901-37
merged a pull request that cites its issue, the stale count read 6 before and
after, and the second attempt then labelled the issue blocked as having no
merged pull request.

Each test copies the loop into a throwaway git repository, because the loop
finds its root with `git rev-parse`, and drives it with a fake `gh` and agent on
PATH, so no network, credentials or model tokens are involved.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parents[1]
COPIED = ("issue_loop.sh", "issue_loop_history.py", "issue_loop_watch.py", "issue_loop_prompt.md")

FAKE_GH = r'''
args = sys.argv[1:]
path = Path(os.environ["FAKE_STATE"])
state = json.loads(path.read_text())
with open(os.environ["FAKE_CALLS"], "a") as log:
    log.write(json.dumps({"tool": "gh", "args": args}) + "\n")
if args[:2] in (["auth", "status"], ["label", "create"], ["issue", "comment"]):
    pass
elif args[:2] == ["api", "user"]:
    print("developer0hye")
elif args[:2] == ["repo", "view"]:
    print("developer0hye")
elif args[:2] == ["api", "repos/{owner}/{repo}/commits/main"]:
    print("current-main-sha")
elif args[:2] == ["run", "list"] or args[:2] == ["api", "graphql"]:
    print("")
elif args[:2] == ["issue", "list"]:
    if "--label" in args:
        print("")
    else:
        pending = [n for n in state["issues"] if n not in state["blocked"]]
        print(pending[0] if pending else "null")
elif args[:2] == ["issue", "view"]:
    if "state" in args:
        print("OPEN")
    else:
        print("true" if int(args[2]) in state["blocked"] else "false")
elif args[:2] == ["issue", "edit"] and "--add-label" in args:
    state["blocked"].append(int(args[2]))
    path.write_text(json.dumps(state))
else:
    sys.exit(1)
'''

# Makes no progress, and on its first run replaces the loop's script the way a
# `git pull` in the main worktree does: a new file renamed over the old one, so
# the running bash keeps reading the inode it opened. Writing in place would let
# the old process read the new bytes and pass the test for the wrong reason.
FAKE_AGENT = r'''
with open(os.environ["FAKE_CALLS"], "a") as log:
    log.write(json.dumps({"tool": "claude", "resume_state": os.environ.get("ISSUE_LOOP_RESUME_STATE")}) + "\n")
marker = Path(os.environ["FAKE_EDITED"])
if not marker.exists():
    marker.touch()
    script = Path(os.environ["FAKE_SCRIPT"])
    text = script.read_text()
    old, new = os.environ["FAKE_EDIT_OLD"], os.environ["FAKE_EDIT_NEW"]
    assert old in text, old
    staged = script.with_name(script.name + ".new")
    staged.write_text(text.replace(old, new, 1))
    staged.chmod(0o755)
    os.replace(staged, script)
print(json.dumps({"result": "nothing merged", "num_turns": 1}))
'''


class ReloadTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        repo = root / "repo"
        (repo / "scripts").mkdir(parents=True)
        subprocess.run(["git", "init", "-q", str(repo)], check=True)
        for name in COPIED:
            shutil.copy2(SCRIPTS / name, repo / "scripts" / name)
        self.script = repo / "scripts" / "issue_loop.sh"
        self.logs = root / "logs"
        self.calls = root / "calls.jsonl"
        self.calls.touch()
        state = root / "state.json"
        state.write_text(json.dumps({"issues": [2229, 2230], "blocked": []}))
        bin_dir = root / "bin"
        bin_dir.mkdir()
        for name, body in (("gh", FAKE_GH), ("claude", FAKE_AGENT), ("sleep", "pass\n")):
            tool = bin_dir / name
            tool.write_text(f"#!{sys.executable}\nimport json, os, sys\nfrom pathlib import Path\n" + body)
            tool.chmod(0o755)
        self.env = {
            **os.environ,
            "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}",
            "FAKE_STATE": str(state),
            "FAKE_CALLS": str(self.calls),
            "FAKE_SCRIPT": str(self.script),
            "FAKE_EDITED": str(root / "edited"),
            "FAKE_EDIT_OLD": 'echo "issue-loop: repo=',
            "FAKE_EDIT_NEW": 'echo "issue-loop (new version): repo=',
        }

    def run_loop(self, **env: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            ["bash", str(self.script), "--once", "--min-free-gb", "0", "--max-issues", "2",
             "--log-dir", str(self.logs)],
            capture_output=True, text=True, env={**self.env, **env}, timeout=30,
        )

    def calls_of(self, tool: str) -> list[dict]:
        return [c for c in map(json.loads, self.calls.read_text().splitlines()) if c["tool"] == tool]

    def run_dirs(self) -> list[str]:
        return sorted(path.name for path in (self.logs / "runs").glob("issue-*/*") if path.is_dir())

    def test_a_changed_script_runs_the_next_attempt(self):
        done = self.run_loop()
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn("issue-loop (new version): repo=", done.stdout)
        self.assertIn("restarting", done.stdout)

    def test_counters_survive_the_restart(self):
        done = self.run_loop()
        self.assertEqual(done.returncode, 0, done.stderr)
        # The second attempt at the same issue is still attempt 2 and still
        # exhausts the per-issue cap, and --max-issues counts both runs.
        self.assertIn("issue #2229 attempt 2 ->", done.stdout)
        self.assertIn("failure 2/2", done.stdout)
        self.assertEqual(len(self.calls_of("claude")), 2)
        labelled = [c["args"] for c in self.calls_of("gh") if c["args"][:2] == ["issue", "edit"]]
        self.assertEqual(labelled, [["issue", "edit", "2229", "--add-label", "autofix-blocked"]])
        # exec keeps the process ID, so run directories stay one numbered series.
        first, second = self.run_dirs()
        self.assertEqual(first.rsplit("-", 2)[1], second.rsplit("-", 2)[1])
        self.assertEqual([first.rsplit("-", 1)[1], second.rsplit("-", 1)[1]], ["0", "1"])

    def test_the_carried_state_does_not_reach_the_agent(self):
        self.assertEqual(self.run_loop().returncode, 0)
        self.assertEqual([c["resume_state"] for c in self.calls_of("claude")], [None, None])

    def test_a_replacement_that_does_not_parse_keeps_the_loaded_version(self):
        done = self.run_loop(FAKE_EDIT_NEW='if then echo "issue-loop (new version): repo=')
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn("does not parse", done.stderr)
        self.assertNotIn("new version", done.stdout)
        self.assertIn("issue #2229 attempt 2 ->", done.stdout)
        self.assertEqual(len(self.calls_of("claude")), 2)

    def test_an_unchanged_script_does_not_restart(self):
        done = self.run_loop(FAKE_EDIT_NEW='echo "issue-loop: repo=')
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertNotIn("restarting", done.stdout)
        self.assertEqual(done.stdout.count("issue-loop: repo="), 1)


if __name__ == "__main__":
    unittest.main()
