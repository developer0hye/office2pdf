# OpenAI Codex

- Project: https://github.com/openai/codex
- Headless CLI: https://learn.chatgpt.com/docs/non-interactive-mode
- Repair-loop example: https://developers.openai.com/cookbook/examples/codex/build_iterative_repair_loops_with_codex

`scripts/issue_loop_codex.py` applies the documented subprocess pattern:
stdin prompts, `codex exec --json`, an output schema, and a separate final-message
file. Each attempt starts a new process without session resume. The supervisor
checks GitHub issue/PR state independently of the report; this is a progress
signal, not independent verification of merge CI or visual correctness.

The local runner additionally bounds wall time and terminates the attempt's
process group on timeout or interruption. It preserves stdout events and stderr
separately so live monitoring does not depend on private Codex session storage.
Attempt history also copies these events and their exact thread ID for retries;
reflection uses the selected agent with the same timeout and report transport.
