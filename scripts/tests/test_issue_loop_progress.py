"""Tests for how the issue loop reads an issue's progress from its pull requests.

The loop calls a run a milestone when more merged pull requests cite the issue
after the run than before it, and `pr-open` when an open one cites it. Both
counts used to read only the 30 most recently created pull requests. On
2026-10-19 run `issue-2217/20261019-160601-claude-31901-2` merged #2321
citing #2217, but #2319 and #2321 had pushed the older citing #2219 out of
that window, so the count read 2 before and 2 after. The loop recorded
`no-progress` and, since that was the second attempt, labeled #2217
`autofix-blocked` with a comment saying no pull request had merged. The same
window gave the next prompts an undercounted preflight note: #2224's run was told
nothing about #2225, which already cited it.

Progress is read through `--print-progress` against a fake `gh` on PATH. The
fake serves `gh pr list` like GitHub does (newest first, `--limit`, a body
search) and runs the loop's own `--jq` filters through `jq`.
"""

import json
import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

LOOP = Path(__file__).resolve().parents[1] / "issue_loop.sh"

FAKE_GH = r'''#!/usr/bin/env python3
import json, os, re, subprocess, sys
args = sys.argv[1:]

def option(name, default=None):
    return args[args.index(name) + 1] if name in args else default

if args[:2] != ["pr", "list"]:
    sys.exit("unexpected gh call: %s" % args)
query = option("--search")
if query is not None and os.environ.get("FAKE_SEARCH_ERROR"):
    sys.exit("HTTP 502")
prs = [pr for pr in json.loads(os.environ["FAKE_PRS"]) if pr["state"] == option("--state", "open")]
prs.sort(key=lambda pr: pr["number"], reverse=True)
if query is not None:
    # GitHub's index splits "#2217" into the word 2217, and may not hold a
    # pull request merged moments ago yet.
    word = query.split()[0]
    prs = [pr for pr in prs if pr.get("indexed", True)
           and re.search(r"\b%s\b" % re.escape(word), pr["body"])]
prs = prs[: int(option("--limit", "30"))]
rows = json.dumps([{field: pr[field] for field in option("--json").split(",")} for pr in prs])
jq = option("--jq")
if jq is None:
    print(rows)
else:
    sys.stdout.write(subprocess.run(["jq", "-r", jq], input=rows, capture_output=True,
                                    text=True, check=True).stdout)
'''


def pr(number: int, body: str, state: str = "merged", indexed: bool = True) -> dict:
    return {"number": number, "body": body, "state": state, "indexed": indexed}


def unrelated(first: int, count: int, state: str = "merged") -> list[dict]:
    return [pr(number, "Related: #1874", state) for number in range(first, first + count)]


@unittest.skipUnless(shutil.which("jq"), "the fake gh evaluates --jq filters with jq")
class ProgressTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        bin_dir = root / "bin"
        bin_dir.mkdir()
        gh = bin_dir / "gh"
        gh.write_text(FAKE_GH, encoding="utf-8")
        gh.chmod(0o755)
        self.env = {**os.environ, "PATH": f"{bin_dir}{os.pathsep}{os.environ['PATH']}"}
        self.logs = root / "logs"

    def tearDown(self):
        self.tmp.cleanup()

    def run_progress(self, issue: int, prs: list[dict], **fake: str) -> subprocess.CompletedProcess:
        return subprocess.run(
            ["bash", str(LOOP), "--log-dir", str(self.logs), "--print-progress", str(issue)],
            capture_output=True,
            text=True,
            env={**self.env, "FAKE_PRS": json.dumps(prs), **fake},
        )

    def progress(self, issue: int, prs: list[dict]) -> dict[str, str]:
        done = self.run_progress(issue, prs)
        self.assertEqual(done.returncode, 0, done.stderr)
        return dict(line.split("=", 1) for line in done.stdout.splitlines())

    def test_a_milestone_merged_while_newer_merges_push_out_an_older_citer_counts(self):
        # #2217's second attempt: one citing pull request is the oldest of the
        # newest thirty when the run starts; the run merges another citer
        # while an unrelated one merges too.
        before = [pr(100, "Related: #2217")] + unrelated(101, 29)
        after = before + unrelated(130, 1) + [pr(131, "Related: #2217")]

        self.assertEqual(self.progress(2217, before)["merged"], "1")
        self.assertEqual(self.progress(2217, after)["merged"], "2")

    def test_a_citing_pull_request_older_than_the_newest_thirty_counts(self):
        prs = [pr(1, "Fixes part of it.\n\nRelated: #2224")] + unrelated(2, 40)

        self.assertEqual(self.progress(2224, prs)["merged"], "1")

    def test_a_pull_request_search_has_not_indexed_yet_still_counts(self):
        prs = unrelated(1, 5) + [pr(6, "Related: #2217", indexed=False)]

        self.assertEqual(self.progress(2217, prs)["merged"], "1")

    def test_a_pull_request_both_lists_return_counts_once(self):
        prs = [pr(5, "Related: #2217"), pr(6, "TODO(#2217) stays open")]

        self.assertEqual(self.progress(2217, prs)["merged"], "2")

    def test_a_body_that_only_contains_the_digits_does_not_count(self):
        # Search matches the bare word; the loop still requires a `#N` citation.
        prs = [pr(7, "Measured 2217 glyphs. Related: #22170"), pr(8, "Related: #1874")]

        self.assertEqual(self.progress(2217, prs)["merged"], "0")

    def test_an_open_citing_pull_request_older_than_the_newest_thirty_is_found(self):
        prs = [pr(1, "Related: #2056", "open")] + unrelated(2, 30, "open")

        self.assertEqual(self.progress(2056, prs)["open"], "1")

    def test_the_newest_open_citing_pull_request_is_reported(self):
        prs = [pr(3, "Related: #2056", "open"), pr(9, "Related: #2056", "open")]
        prs += unrelated(10, 3, "open") + [pr(20, "Related: #2056")]

        self.assertEqual(self.progress(2056, prs), {"merged": "1", "open": "9"})

    def test_no_citing_pull_request_reports_none(self):
        self.assertEqual(self.progress(2056, unrelated(1, 3)), {"merged": "0", "open": ""})

    def test_a_failed_search_is_an_error_not_a_count(self):
        # Reading zero would record the run as `no-progress` and count it
        # toward the blocked label; the loop must stop with a state error.
        done = self.run_progress(2217, [pr(5, "Related: #2217")], FAKE_SEARCH_ERROR="1")

        self.assertNotEqual(done.returncode, 0)
        self.assertNotIn("merged=", done.stdout)


if __name__ == "__main__":
    unittest.main()
