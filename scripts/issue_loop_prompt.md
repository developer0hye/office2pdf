Read `AGENTS.md` and `METHODOLOGY.md` before you start, and follow them literally.
The points below are the ones an unattended run gets wrong most often; they do not
replace those documents.

- Work in a git worktree on a `<type>/<short-description>` branch created from current
  `main`. Never `git checkout` in the main working tree — other sessions use it.
- Test-driven: write the failing test first, then the minimal fix, then refactor. The
  test must pin the general rule, not the one input from the issue.
- Sign off every commit (`git commit -s`) with the repository's configured identity.
- Delegate the read-only documentation freshness audit and wait for `PASS:` before
  **every** commit, not only the one that opens the pull request.
- Layout and rendering claims need measurement, not inspection: use
  `scripts/compare_layout.py --audit --fine-shift PT`,
  `scripts/compare_render.py --fine-shift PT --artifacts-dir PATH --cluster-report
  PATH --cluster-dispositions PATH --strict-clusters`, and
  `scripts/compare_text_layer.py`, and look at the emitted images. A numeric
  report alone is not a visual pass.
- Settle an unclear native-application rule with `scripts/probe_harness.py` (one factor
  per variant) rather than by arguing from the corpus.
- A visual defect fix owes the evidence contract in `AGENTS.md`:
  `assets/bugfixes/issue-<number>/gt.jpg`, `before.jpg`, `after.jpg`, and the pull
  request body the contract check expects.
- Never name anything under `tests/classified_fixtures/` in a commit message, pull
  request, or issue comment.
- Re-run the audit on the fixed output before closing an issue. Every deviation still
  visible must already have its own open issue — file the missing ones first.
