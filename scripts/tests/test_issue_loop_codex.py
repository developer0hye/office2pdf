"""Exercise unattended runners without changing GitHub or spending model tokens."""
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import tempfile
import unittest

SCRIPTS = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(SCRIPTS))


class CodexLoopTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.bin = self.root / "bin"
        self.bin.mkdir()
        self.state = self.root / "state.json"
        self.calls = self.root / "calls.jsonl"
        self.state.write_text(json.dumps({"issues": [2101, 2107], "closed": []}))
        self.env = {**os.environ, "PATH": str(self.bin) + os.pathsep + os.environ["PATH"],
                    "FAKE_STATE": str(self.state), "FAKE_CALLS": str(self.calls)}
        self.tool("sleep", "pass")
        self.tool("gh", '''
args = sys.argv[1:]
with open(os.environ['FAKE_CALLS'], 'a') as f:
    f.write(json.dumps({'tool': 'gh', 'args': args}) + '\\n')
state = json.loads(Path(os.environ['FAKE_STATE']).read_text())
if args[:2] == ['auth', 'status']: pass
elif args[:2] == ['api', 'user']: print(state.get('login', os.environ.get('FAKE_LOGIN', 'developer0hye')))
elif args[:2] == ['auth', 'switch'] and os.environ.get('ALLOW_SWITCH'):
    state['login'] = args[-1]
    Path(os.environ['FAKE_STATE']).write_text(json.dumps(state))
elif args[:2] == ['repo', 'view']: print('developer0hye')
elif args[:2] == ['label', 'create']: pass
elif args[:2] == ['issue', 'list']:
    if os.environ.get('QUEUE_ERROR'): sys.exit(1)
    pending = [n for n in state['issues'] if n not in state['closed']]
    print(pending[0] if pending else 'null')
elif args[:2] == ['pr', 'list']: print('0' if 'merged' in args else '')
elif args[:2] == ['issue', 'view']:
    if 'state' in args: print('CLOSED' if int(args[2]) in state['closed'] else 'OPEN')
    else: print('false')
else: sys.exit(1)
''')
        agent = '''
args = sys.argv[1:]
prompt = sys.stdin.read() if sys.argv[0].endswith('codex') else args[args.index('-p') + 1]
issue = int(re.search(r'GitHub issue #(\\d+)', prompt).group(1))
with open(os.environ['FAKE_CALLS'], 'a') as f:
    f.write(json.dumps({'tool': Path(sys.argv[0]).name, 'args': args, 'prompt': prompt, 'issue': issue}) + '\\n')
if sys.argv[0].endswith('codex'):
    print(json.dumps({'type': 'thread.started', 'thread_id': str(issue)}), flush=True)
    if os.environ.get('FAKE_AUTH_ERROR'):
        print('Transport channel closed, when Auth(AuthorizationRequired)', file=sys.stderr)
        sys.exit(1)
    if os.environ.get('FAKE_CONFIG_ERROR'):
        print(json.dumps({'type': 'turn.failed', 'error': {'message': "The 'unavailable' model is not supported when using Codex with a ChatGPT account."}}), flush=True)
        sys.exit(1)
    if os.environ.get('FAKE_QUOTA'):
        print(json.dumps({'type': 'turn.failed', 'error': {'message': "You've hit your usage limit."}}), flush=True)
        sys.exit(1)
    print(json.dumps({'type': 'turn.completed', 'usage': {}}), flush=True)
    result_path = Path(args[args.index('--output-last-message') + 1])
    result_path.write_text(json.dumps({
        'issue': int(os.environ.get('REPORT_ISSUE', issue)), 'status': 'resolved', 'pr_number': None, 'summary': 'Verified fix.', 'remaining': []}))
    if os.environ.get('MALFORMED_REPORT'): result_path.write_text('{')
if not os.environ.get('NO_PROGRESS'):
    path = Path(os.environ['FAKE_STATE'])
    state = json.loads(path.read_text())
    state['closed'].append(issue)
    path.write_text(json.dumps(state))
sys.exit(int(os.environ.get('AGENT_EXIT', '0')))
'''
        self.tool("codex", agent)
        self.tool("claude", agent)

    def tool(self, name, body):
        path = self.bin / name
        path.write_text(f"#!{sys.executable}\nimport json, os, re, sys\nfrom pathlib import Path\n" + body)
        path.chmod(0o755)

    def run_loop(self, *args, **env):
        return subprocess.run(["bash", str(SCRIPTS / "issue_loop.sh"), "--once", "--min-free-gb", "0",
                               "--log-dir", str(self.root / "logs"), *args],
                              env={**self.env, **env}, text=True, capture_output=True, timeout=15)

    def records(self, tool):
        return [r for r in map(json.loads, self.calls.read_text().splitlines()) if r['tool'] == tool]

    def test_two_issues_get_separate_codex_exec_and_reports(self):
        done = self.run_loop("--agent", "codex", "--model", "test-model")
        self.assertEqual(done.returncode, 0, done.stderr)
        calls = self.records("codex")
        self.assertEqual([r['issue'] for r in calls], [2101, 2107])
        for record in calls:
            args = record['args']
            self.assertEqual(args[:3], ['-a', 'never', 'exec'])
            self.assertIn('--json', args)
            self.assertIn('--output-schema', args)
            self.assertIn('project_doc_max_bytes=65536', args)
            self.assertIn('test-model', args)
            self.assertNotIn('resume', args)
            self.assertNotIn('--max-turns', args)
        self.assertEqual(len(list((self.root / 'logs').glob('*.jsonl'))), 2)
        self.assertEqual(len(list((self.root / 'logs').glob('*.result.json'))), 2)
        self.assertIn('queue empty', done.stdout)

    def test_default_still_runs_claude(self):
        done = self.run_loop('--max-issues', '1')
        self.assertEqual(done.returncode, 0, done.stderr)
        call = self.records('claude')[0]
        self.assertIn('--dangerously-skip-permissions', call['args'])
        self.assertIn('--max-turns', call['args'])

    def test_resolved_summary_without_github_progress_is_not_success(self):
        done = self.run_loop('--agent', 'codex', '--max-issues', '1', NO_PROGRESS='1', AGENT_EXIT='1')
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn('no progress (exit 1)', done.stdout)
        self.assertNotIn('closed — success', done.stdout)
        self.assertEqual(len(self.records('codex')), 1)

    def test_queue_failure_is_not_reported_as_empty(self):
        done = self.run_loop('--agent', 'codex', QUEUE_ERROR='1')
        self.assertNotEqual(done.returncode, 0)
        self.assertNotIn('queue empty', done.stdout)
        self.assertEqual(self.records('codex'), [])

    def test_wrong_account_cannot_mutate_repo(self):
        done = self.run_loop('--agent', 'codex', FAKE_LOGIN='someone-else')
        self.assertNotEqual(done.returncode, 0)
        writes = [r for r in self.records('gh') if r['args'][:2] == ['label', 'create']]
        self.assertEqual(writes, [])

    def test_account_switch_is_verified_before_write(self):
        done = self.run_loop('--agent', 'codex', '--max-issues', '1',
                             FAKE_LOGIN='someone-else', ALLOW_SWITCH='1')
        self.assertEqual(done.returncode, 0, done.stderr)
        args = [r['args'] for r in self.records('gh')]
        switch = args.index(['auth', 'switch', '--user', 'developer0hye'])
        write = next(i for i, a in enumerate(args) if a[:2] == ['label', 'create'])
        self.assertLess(switch, write)
        self.assertIn(['api', 'user', '--jq', '.login'], args[switch + 1:write])

    def test_authentication_configuration_failure_does_not_block_issue(self):
        done = self.run_loop('--agent', 'codex', FAKE_AUTH_ERROR='1')
        self.assertEqual(done.returncode, 78, done.stderr)
        self.assertEqual(len(self.records('codex')), 1)
        self.assertFalse(any(r['args'][:2] == ['issue', 'edit'] for r in self.records('gh')))

    def test_model_configuration_error_stops_without_blocking_issue(self):
        done = self.run_loop('--agent', 'codex', FAKE_CONFIG_ERROR='1')
        self.assertEqual(done.returncode, 78, done.stderr)
        self.assertEqual(len(self.records('codex')), 1)
        self.assertFalse(any(r['args'][:2] == ['issue', 'edit'] for r in self.records('gh')))

    def test_usage_limit_obeys_trial_attempt_cap(self):
        done = self.run_loop('--agent', 'codex', '--max-issues', '1', FAKE_QUOTA='1')
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn('usage limit reached', done.stdout)
        self.assertEqual(len(self.records('codex')), 1)
        self.assertNotIn('sleeping', done.stdout)

    def test_report_for_wrong_issue_is_rejected(self):
        done = self.run_loop('--agent', 'codex', '--max-issues', '1', NO_PROGRESS='1', REPORT_ISSUE='999')
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn('no progress (exit 1)', done.stdout)
        diagnostic = next((self.root / 'logs').glob('*.stderr.log')).read_text()
        self.assertIn('invalid or missing', diagnostic)

    def test_malformed_final_report_is_not_a_successful_attempt(self):
        done = self.run_loop('--agent', 'codex', '--max-issues', '1', NO_PROGRESS='1', MALFORMED_REPORT='1')
        self.assertEqual(done.returncode, 0, done.stderr)
        self.assertIn('no progress (exit 1)', done.stdout)

    def test_interrupted_attempt_stops_supervisor(self):
        done = self.run_loop('--agent', 'codex', '--max-issues', '1', NO_PROGRESS='1', AGENT_EXIT='130')
        self.assertEqual(done.returncode, 130)

    def test_invalid_agent_rejected(self):
        self.assertEqual(self.run_loop('--agent', 'unknown').returncode, 2)

    def test_nonpositive_runtime_rejected(self):
        self.assertEqual(self.run_loop('--agent', 'codex', '--max-runtime', '0').returncode, 2)


class CodexRunnerTest(unittest.TestCase):
    def test_timeout_stops_child_tool_even_if_it_ignores_termination(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            marker = root / 'orphan-write'
            ready = root / 'child-ready'
            child = root / 'child.py'
            child.write_text('import signal, time\nfrom pathlib import Path\n'
                             'signal.signal(signal.SIGTERM, signal.SIG_IGN)\n'
                             f'Path({str(ready)!r}).touch()\n'
                             f'time.sleep(3)\nPath({str(marker)!r}).touch()\n')
            codex = root / 'codex'
            codex.write_text(f'#!{sys.executable}\nimport subprocess, sys, time\n'
                             f'subprocess.Popen([sys.executable, {str(child)!r}])\ntime.sleep(60)\n')
            codex.chmod(0o755)
            done = subprocess.run([sys.executable, str(SCRIPTS / 'issue_loop_codex.py'),
                                   '--log', str(root / 'run.jsonl'), '--max-runtime', '1', '--issue', '2101'],
                                  input='GitHub issue #2101', text=True, capture_output=True,
                                  env={**os.environ, 'PATH': tmp + os.pathsep + os.environ['PATH']}, timeout=10)
            self.assertEqual(done.returncode, 124, done.stderr)
            self.assertTrue(ready.exists())
            time.sleep(3)
            self.assertFalse(marker.exists(), 'timed-out child continued writing')

    def test_timeout_stops_attempt_and_keeps_diagnostics(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            codex = root / 'codex'
            codex.write_text(f'#!{sys.executable}\nimport time\nprint(\'{{"type":"thread.started"}}\', flush=True)\ntime.sleep(60)\n')
            codex.chmod(0o755)
            done = subprocess.run([sys.executable, str(SCRIPTS / 'issue_loop_codex.py'),
                                   '--log', str(root / 'run.jsonl'), '--max-runtime', '1', '--issue', '2101'],
                                  input='GitHub issue #2101', text=True, capture_output=True,
                                  env={**os.environ, 'PATH': tmp + os.pathsep + os.environ['PATH']}, timeout=10)
            self.assertEqual(done.returncode, 124, done.stderr)
            self.assertIn('thread.started', (root / 'run.jsonl').read_text())
            self.assertIn('timeout', (root / 'run.stderr.log').read_text())


if __name__ == '__main__':
    unittest.main()
