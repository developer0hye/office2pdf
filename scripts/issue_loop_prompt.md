Read `AGENTS.md` and `METHODOLOGY.md` before you start, and follow them literally.
The points below are the ones an unattended run gets wrong most often; they do not
replace those documents.

- Work in a git worktree on a `<type>/<short-description>` branch created from current
  `main`. Never `git checkout` in the main working tree — other sessions use it.
- A shared `CARGO_TARGET_DIR` can let a build from another worktree replace the
  binary you intended to measure. Build baseline and candidate binaries in separate,
  run-owned scratch targets and preserve the binaries used for evidence. `cmp`
  detects identical binaries; it does not establish baseline provenance or fidelity.
  Before ending a run, including one leaving a draft PR, remove only private build
  directories created by that run and no longer needed for active work, after
  preserving required evidence. Never delete the runner's shared target, shared
  caches, or another session's build output. Keep uncertain ownership or required
  directories and record why they remain.
- Test-driven: write the failing test first, then the minimal fix, then refactor. The
  test must pin the general rule, not the one input from the issue.
- Sign off every commit (`git commit -s`) with the repository's configured identity.
- Merge only when every check on the pull request passes, including one that also
  fails on `main`: an inherited failure is a red main to fix first (the `main-red`
  issue), never a reason to merge. Read the full `gh pr checks` output, unfiltered.
- Delegate the read-only documentation freshness audit and wait for `PASS:` before
  **every** commit, not only the one that opens the pull request.
- Layout and rendering claims need measurement, not inspection: use
  `scripts/compare_layout.py --audit --fine-shift PT`,
  `scripts/compare_render.py --fine-shift PT --artifacts-dir PATH --cluster-report
  PATH --cluster-dispositions PATH --strict-clusters`, and
  `scripts/compare_text_layer.py`, and look at the emitted images. A numeric
  report alone is not a visual pass.
- Re-derive cluster dispositions from current crops on every audit; never carry
  an old explanation forward by stable ID or bounding box.
- Settle an unclear native-application rule with `scripts/probe_harness.py` (one factor
  per variant) rather than by arguing from the corpus.
- A visual defect fix owes the evidence contract in `AGENTS.md`:
  `assets/bugfixes/issue-<number>/gt.jpg`, `before.jpg`, `after.jpg`,
  `layout-audit.json`, one `render-clusters-page-<page>.json` per compared page, and
  the pull request body the contract check expects. `AGENTS.md` is authoritative if
  this list and it disagree.
- Before implementing a visual fix, check whether the native export its evidence
  needs is committed, is supplied by an external reference corpus the issue names
  (check its producer with `pdfinfo`), or can be made on this host. If none: when
  the issue itself asks for new native exports or probes, block it without code
  (rule 8); otherwise implement and test the fix, open it as a **draft** pull
  request, comment on the issue what a Mac must export to finish it, add the
  blocked label, and merge nothing. This is the settled route; do not re-derive
  it from earlier drafts.
- Never name anything under `tests/classified_fixtures/` in a commit message, pull
  request, or issue comment.
- Re-run the audit on the fixed output before closing an issue. Every material deviation
  still visible must already have its own open issue — file the missing ones first.
- Apply the Materiality rule in `AGENTS.md`: spend effort on what a reader would
  notice, not on sub-point drift; send sub-material findings to #1874.
- After visual acceptance, complete the final memory and speed pass in `AGENTS.md`
  before final CI and merge. Record elapsed-time/peak-memory measurements and the
  optimization or measured reason no simple improvement is worthwhile in the PR's
  `Final performance pass`. Revalidate tests and the full affected-file visual
  comparison after optimization changes; a visual pass alone is not completion.
  The Visual PR Contract rejects missing or unfinished performance declarations.
  Verify measurements and revalidation before marking Complete, and verify no
  conversion runtime impact before using an exemption; CI does not establish these.
