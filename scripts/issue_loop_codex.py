#!/usr/bin/env python3
"""Run one fresh Codex attempt; preserve events, diagnostics, and its final report.

Uses only the standard library. GitHub, not the report, determines queue progress.
"""
from __future__ import annotations

import argparse
import json
import os
import re
from pathlib import Path
import signal
import subprocess
import sys

STATUSES = ('resolved', 'milestone', 'blocked', 'incomplete')


def valid_report(path: Path, issue: int) -> bool:
    try:
        report = json.loads(path.read_text(encoding='utf-8'))
    except (OSError, ValueError):
        return False
    if not isinstance(report, dict):
        return False
    pr_number = report.get('pr_number')
    return (
        set(report) == {'issue', 'status', 'pr_number', 'summary', 'remaining'}
        and type(report['issue']) is int and report['issue'] == issue
        and report['status'] in STATUSES
        and (pr_number is None or (type(pr_number) is int and pr_number > 0))
        and isinstance(report['summary'], str)
        and isinstance(report['remaining'], list)
        and all(isinstance(item, str) for item in report['remaining'])
    )


def configuration_failure(events_path: Path, stderr_path: Path) -> bool:
    """Detect CLI/auth/model setup failures without scanning agent/tool prose."""
    messages: list[str] = []
    for line in events_path.read_text(encoding='utf-8', errors='replace').splitlines():
        try:
            event = json.loads(line)
        except ValueError:
            continue
        if not isinstance(event, dict) or event.get('type') not in ('error', 'turn.failed'):
            continue
        error = event.get('error') or {}
        messages.append(str(event.get('message') or (error.get('message') if isinstance(error, dict) else error)))
    messages.append(stderr_path.read_text(encoding='utf-8', errors='replace'))
    pattern = (
        r"model.{0,100}(?:not supported|does not exist)|not logged in|invalid_api_key|"
        r"incorrect api key|invalid (?:output )?schema|unexpected argument|"
        r"unrecognized argument|not authenticated|authentication required|"
        r"authorizationrequired|unauthorized|token.{0,40}expired"
    )
    return any(re.search(pattern, message, re.IGNORECASE) for message in messages)


def stop_process_group(process: subprocess.Popen[str]) -> None:
    # Include tools spawned by Codex so a timed-out run cannot keep editing.
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        pass
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        pass
    process.wait()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--log', type=Path, required=True)
    parser.add_argument('--issue', type=int, required=True)
    parser.add_argument('--model', default='')
    parser.add_argument('--sandbox', choices=('read-only', 'workspace-write', 'danger-full-access'),
                        default='danger-full-access')
    parser.add_argument('--max-runtime', type=int, default=7200)
    args = parser.parse_args()
    if args.max_runtime <= 0:
        parser.error('--max-runtime must be positive')
    stem = args.log.with_suffix('')
    result = stem.with_suffix('.result.json')
    diagnostics = stem.with_suffix('.stderr.log')
    # Never accept a report left behind by an earlier invocation of this helper.
    result.unlink(missing_ok=True)
    command = ['codex', '-a', 'never', 'exec', '--sandbox', args.sandbox, '--json',
               '-c', 'project_doc_max_bytes=65536', '--output-schema',
               str(Path(__file__).with_name('issue_loop_result.schema.json')),
               '--output-last-message', str(result)]
    if args.model:
        command += ['--model', args.model]
    command.append('-')
    prompt = sys.stdin.read()
    prompt += (f'\nReturn the required JSON report for issue #{args.issue}. '
               'Use resolved only after verifying the issue is fixed; otherwise use milestone, '
               'blocked, or incomplete and list remaining work.\n')
    def interrupted(signum: int, frame: object) -> None:
        raise KeyboardInterrupt
    signal.signal(signal.SIGTERM, interrupted)
    with args.log.open('w', encoding='utf-8') as events, diagnostics.open('w', encoding='utf-8') as errors:
        try:
            process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=events,
                                       stderr=errors, start_new_session=True, text=True)
        except OSError as error:
            print(str(error), file=errors)
            return 127
        try:
            process.communicate(input=prompt, timeout=args.max_runtime)
        except subprocess.TimeoutExpired:
            stop_process_group(process)
            print(f'timeout: attempt exceeded {args.max_runtime}s', file=errors)
            return 124
        except KeyboardInterrupt:
            stop_process_group(process)
            print('interrupted: stopped attempt and child tools', file=errors)
            return 130
        if process.returncode != 0:
            events.flush()
            errors.flush()
            if configuration_failure(args.log, diagnostics):
                print('Codex configuration/authentication failure: check model and login', file=errors)
                return 78
            return process.returncode if process.returncode > 0 else 128 - process.returncode
        if not valid_report(result, args.issue):
            print('invalid or missing structured final report', file=errors)
            return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
