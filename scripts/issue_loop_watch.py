#!/usr/bin/env python3
"""Show what the agent started by `scripts/issue_loop.sh` is doing right now.

Claude runs are read from its live session transcripts; Codex runs are read from
JSONL logs written by the loop. Both follow new attempts as the queue advances.

    scripts/issue_loop_watch.py --agent codex            # follow live
    scripts/issue_loop_watch.py --agent codex --dump 20  # last 20 events
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import subprocess
import sys
import time
from typing import TextIO

POLL_SECONDS = 2.0
TAIL_LINES_ON_ATTACH = 40
TOOL_ARG_KEYS = ("command", "file_path", "pattern", "description")


def repo_root() -> str:
    return subprocess.run(
        ["git", "rev-parse", "--show-toplevel"],
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()


def transcript_dir(repo_path: str, home: str | None = None) -> str:
    """Directory holding this repository's Claude Code session transcripts.

    Claude Code names it after the absolute project path with every separator
    replaced by a hyphen, so `/Users/x/src/repo` becomes `-Users-x-src-repo`.
    """
    home = home if home is not None else os.path.expanduser("~")
    return os.path.join(home, ".claude", "projects", repo_path.replace(os.sep, "-"))


def newest_transcript(directory: str, exclude: tuple[str, ...] = ()) -> str | None:
    candidates = [
        path
        for path in glob.glob(os.path.join(directory, "*.jsonl"))
        if not any(token in path for token in exclude)
    ]
    return max(candidates, key=os.path.getmtime) if candidates else None


def format_events(raw_line: str, max_text: int = 300) -> list[str]:
    """One display line per tool call or assistant message in a transcript line.

    Anything else (user turns, tool results, malformed lines) yields nothing: the
    point is to see what the agent decided, not to replay its inputs.
    """
    try:
        record = json.loads(raw_line)
    except (json.JSONDecodeError, TypeError):
        return []
    if not isinstance(record, dict):
        return []
    kind = record.get("type")
    if kind in ("error", "turn.failed"):
        error = record.get("error") or {}
        message = error.get("message", "") if isinstance(error, dict) else str(error)
        return [f"[ERROR] {str(record.get('message') or message)[:max_text]}".replace("\n", " ")]
    if kind in ("item.started", "item.completed"):
        item = record.get("item") or {}
        if not isinstance(item, dict):
            return []
        if item.get("type") == "command_execution":
            status = "start" if kind == "item.started" else f"exit {item.get('exit_code')}"
            return [f"[TOOL] {status}: {str(item.get('command', ''))[:150]}".replace("\n", " ")]
        if kind == "item.completed" and item.get("type") == "agent_message":
            return [f"[SAY]  {str(item.get('text', ''))[:max_text]}".replace("\n", " ")]
        if kind == "item.completed" and item.get("type") == "file_change":
            paths = ", ".join(str(change.get("path", "")) for change in item.get("changes", []))
            return [f"[EDIT] {paths[:max_text]}"]
        if item.get("type") == "mcp_tool_call":
            return [f"[TOOL] {item.get('server')}/{item.get('tool')}: {item.get('status', kind)}"]
        return []
    if kind != "assistant":
        return []
    message = record.get("message") or {}
    events: list[str] = []
    for block in message.get("content") or []:
        if not isinstance(block, dict):
            continue
        if block.get("type") == "tool_use":
            payload = block.get("input") or {}
            argument = next(
                (str(payload[key]) for key in TOOL_ARG_KEYS if payload.get(key)), ""
            )
            events.append(
                f"[TOOL] {block.get('name')}: {argument[:150]}".replace("\n", " ")
            )
        elif block.get("type") == "text" and (block.get("text") or "").strip():
            events.append(f"[SAY]  {block['text'][:max_text]}".replace("\n", " "))
    return events


def dump(directory: str, count: int, exclude: tuple[str, ...]) -> int:
    path = newest_transcript(directory, exclude)
    if path is None:
        print(f"no session transcript under {directory}", file=sys.stderr)
        return 1
    print(f">>> {os.path.basename(path)}")
    events: list[str] = []
    with open(path, encoding="utf-8") as handle:
        for line in handle:
            events.extend(format_events(line))
    print("\n".join(events[-count:]))
    return 0


def read_complete_lines(handle: TextIO) -> list[str]:
    """Keep the cursor before a torn JSONL record so the next poll retries it."""
    lines: list[str] = []
    while True:
        position = handle.tell()
        line = handle.readline()
        if not line:
            break
        if not line.endswith("\n"):
            handle.seek(position)
            break
        lines.append(line)
    return lines


def follow(directory: str, exclude: tuple[str, ...]) -> int:
    current: str | None = None
    handle = None
    try:
        while True:
            newest = newest_transcript(directory, exclude)
            if newest is not None and newest != current:
                if handle is not None:
                    handle.close()
                current = newest
                handle = open(current, encoding="utf-8")
                recent = read_complete_lines(handle)[-TAIL_LINES_ON_ATTACH:]
                print(f"\n>>> now watching {os.path.basename(current)}", flush=True)
                for line in recent:
                    for event in format_events(line):
                        print(event, flush=True)
            if handle is not None:
                for line in read_complete_lines(handle):
                    for event in format_events(line):
                        print(event, flush=True)
            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        return 0
    finally:
        if handle is not None:
            handle.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--dump",
        type=int,
        metavar="N",
        help="print the last N events of the newest session and exit",
    )
    parser.add_argument(
        "--exclude",
        default="",
        help="comma-separated session ids to ignore (e.g. the session you are typing in)",
    )
    parser.add_argument("--agent", choices=("claude", "codex"), default="claude")
    parser.add_argument("--log-dir", help="Codex JSONL directory (default target/issue-loop-logs)")
    args = parser.parse_args()

    exclude = tuple(token for token in args.exclude.split(",") if token)
    if args.agent == "codex":
        directory = args.log_dir or os.path.join(repo_root(), "target", "issue-loop-logs")
    else:
        directory = transcript_dir(repo_root())
    if args.dump is not None:
        return dump(directory, args.dump, exclude)
    return follow(directory, exclude)


if __name__ == "__main__":
    sys.exit(main())
