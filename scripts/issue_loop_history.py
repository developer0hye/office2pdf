#!/usr/bin/env python3
"""Experience store for `scripts/issue_loop.sh`: every run's prompt, report, outcome
and full transcript, kept where the next agent can read them.

Each loop run is a fresh agent process, so without this a retry starts blind
and repeats the dead end before it. Meta-Harness (arXiv:2603.28052) found that an
agent given raw prior traces on a filesystem, to grep as it needs, beats one given
scores or summaries; this keeps those traces and points the next prompt at them.

    scripts/issue_loop_history.py list                 # every run, oldest first
    scripts/issue_loop_history.py list --issue 1721    # one issue's attempts
    scripts/issue_loop_history.py show RUN_DIR         # outcome plus the run's last events

Layout under the loop's log directory:

    runs/index.jsonl                         one outcome per line, append-only
    runs/issue-<N>/<attempt>/               prompt.md, result.json, outcome.json,
                                             transcript.jsonl (when the session saved one)
    runs/reflect-<attempt>/                 prompt.md and the reflection report/events

Codex attempts save live events.jsonl and events.stderr.log in their run directory;
record also preserves those events as transcript.jsonl and extracts the exact thread ID.

The loop calls `record` after each run and `context` while building a prompt.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

from issue_loop_watch import format_events, transcript_dir

FINAL_MESSAGE_CHARS = 1500
CONTEXT_MESSAGE_CHARS = 600


def runs_dir(log_dir: str) -> Path:
    return Path(log_dir) / "runs"


def read_report(path: Path) -> dict:
    """The CLI's final JSON report, or a stand-in when the run died before writing one."""
    try:
        raw = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return {}
    try:
        report = json.loads(raw)
    except ValueError:
        return {"result": raw.strip()}
    return report if isinstance(report, dict) else {}


def record(args: argparse.Namespace) -> int:
    run_dir = Path(args.run_dir)
    report = read_report(run_dir / "result.json")
    session_id = report.get("session_id")
    transcript = None
    if args.agent == "codex":
        source = run_dir / "events.jsonl"
        if source.is_file():
            transcript = run_dir / "transcript.jsonl"
            shutil.copyfile(source, transcript)
            for line in source.read_text(encoding="utf-8", errors="replace").splitlines():
                try:
                    event = json.loads(line)
                except ValueError:
                    continue
                if isinstance(event, dict) and event.get("type") == "thread.started":
                    session_id = event.get("thread_id")
                    if isinstance(session_id, str) and session_id:
                        break
    elif isinstance(session_id, str) and session_id:
        source = Path(transcript_dir(args.repo_root)) / f"{session_id}.jsonl"
        if source.is_file():
            transcript = run_dir / "transcript.jsonl"
            shutil.copyfile(source, transcript)
    outcome = {
        "issue": args.issue,
        "agent": args.agent,
        "run_dir": str(run_dir),
        "outcome": args.outcome,
        "rc": args.rc,
        "session_id": session_id if isinstance(session_id, str) else None,
        "num_turns": report.get("num_turns"),
        "duration_ms": report.get("duration_ms"),
        "total_cost_usd": report.get("total_cost_usd"),
        "is_error": report.get("is_error"),
        "final_message": str(report.get("result") or report.get("summary") or "")[:FINAL_MESSAGE_CHARS],
        "transcript": str(transcript) if transcript else None,
    }
    (run_dir / "outcome.json").write_text(json.dumps(outcome, indent=2) + "\n", encoding="utf-8")
    index = runs_dir(args.log_dir) / "index.jsonl"
    index.parent.mkdir(parents=True, exist_ok=True)
    with index.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(outcome) + "\n")
    return 0


def load_index(log_dir: str) -> list[dict]:
    index = runs_dir(log_dir) / "index.jsonl"
    rows: list[dict] = []
    try:
        lines = index.read_text(encoding="utf-8").splitlines()
    except OSError:
        return rows
    for line in lines:
        try:
            row = json.loads(line)
        except ValueError:
            continue
        if isinstance(row, dict):
            rows.append(row)
    return rows


def context(args: argparse.Namespace) -> int:
    """Prompt section for the next run at an issue; prints nothing when there is nothing to say."""
    parts: list[str] = []
    if args.merged_refs > 0:
        noun = "pull request" if args.merged_refs == 1 else "pull requests"
        parts.append(
            f"Preflight: {args.merged_refs} merged {noun} already cite #{args.issue}. "
            "Before planning new work, check whether they already resolve it; if so, "
            "verify on current main and close it with that evidence."
        )
    attempts = [row for row in load_index(args.log_dir) if row.get("issue") == args.issue]
    if attempts:
        lines = [
            f"Earlier unattended attempts at #{args.issue} (oldest first). Read their "
            "outcome.json and, before repeating an approach, the transcript.jsonl to see "
            "where it stalled. `scripts/issue_loop_history.py show RUN_DIR` prints a "
            "run's last events.",
        ]
        for row in attempts:
            message = " ".join(str(row.get("final_message") or "").split())
            lines.append(
                f"- {row.get('run_dir')}: outcome={row.get('outcome')} rc={row.get('rc')} "
                f"turns={row.get('num_turns')}\n"
                f"  transcript: {row.get('transcript') or 'not saved'}\n"
                f"  final message: {message[:CONTEXT_MESSAGE_CHARS] or '(none)'}"
            )
        parts.append("\n".join(lines))
    if parts:
        print("\n\n".join(parts))
    return 0


def list_runs(args: argparse.Namespace) -> int:
    print(f"{'run':<16} {'issue':<7} {'outcome':<12} {'rc':>3} {'turns':>6} {'cost':>8}")
    for row in load_index(args.log_dir):
        if args.issue is not None and row.get("issue") != args.issue:
            continue
        cost = row.get("total_cost_usd")
        print(
            f"{Path(str(row.get('run_dir'))).name:<16} {'#' + str(row.get('issue')):<7} "
            f"{str(row.get('outcome')):<12} {str(row.get('rc')):>3} "
            f"{str(row.get('num_turns')):>6} "
            f"{('$%.2f' % cost) if isinstance(cost, (int, float)) else '-':>8}"
        )
    return 0


def show(args: argparse.Namespace) -> int:
    run_dir = Path(args.run_dir)
    try:
        print((run_dir / "outcome.json").read_text(encoding="utf-8"))
    except OSError:
        print(f"no outcome.json in {run_dir}", file=sys.stderr)
        return 1
    events: list[str] = []
    try:
        with (run_dir / "transcript.jsonl").open(encoding="utf-8") as handle:
            for line in handle:
                events.extend(format_events(line))
    except OSError:
        print("(no transcript saved for this run)")
        return 0
    print("\n".join(events[-args.events:]))
    return 0


def default_log_dir() -> str:
    root = subprocess.run(["git", "rev-parse", "--show-toplevel"], capture_output=True,
                          text=True, check=False).stdout.strip() or os.getcwd()
    return os.path.join(root, "target", "issue-loop-logs")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)

    record_parser = commands.add_parser("record", help="store one finished run (called by the loop)")
    record_parser.add_argument("--log-dir", required=True)
    record_parser.add_argument("--repo-root", required=True,
                               help="repository root; names Claude's transcript folder")
    record_parser.add_argument("--agent", choices=("claude", "codex"), default="claude")
    record_parser.add_argument("--run-dir", required=True)
    record_parser.add_argument("--issue", type=int, required=True)
    record_parser.add_argument("--outcome", required=True)
    record_parser.add_argument("--rc", type=int, required=True)
    record_parser.set_defaults(handler=record)

    context_parser = commands.add_parser("context", help="prompt section for an issue's next run")
    context_parser.add_argument("--log-dir", required=True)
    context_parser.add_argument("--issue", type=int, required=True)
    context_parser.add_argument("--merged-refs", type=int, default=0)
    context_parser.set_defaults(handler=context)

    list_parser = commands.add_parser("list", help="one line per recorded run")
    list_parser.add_argument("--log-dir", default=None)
    list_parser.add_argument("--issue", type=int, default=None)
    list_parser.set_defaults(handler=list_runs)

    show_parser = commands.add_parser("show", help="a run's outcome and its last transcript events")
    show_parser.add_argument("run_dir")
    show_parser.add_argument("--events", type=int, default=30)
    show_parser.set_defaults(handler=show)

    args = parser.parse_args()
    if getattr(args, "log_dir", "") is None:
        args.log_dir = default_log_dir()
    return args.handler(args)


if __name__ == "__main__":
    sys.exit(main())
