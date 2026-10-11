# Agent Rules

Canonical instructions for all coding agents.

- Communicate and work in English.
- Read and follow `METHODOLOGY.md` before any task and `PRD.md`, if present, before development.
- Prefer native Rust and the standard library; justify third-party dependencies.
- Keep instructions and documentation concise. After code changes or merges, update `README.md` when usage, setup, architecture, or APIs change.
- Complete, validate, and commit each planned task before starting the next.
- **Before every commit, delegate a read-only documentation freshness audit and wait for `PASS:`.** Codex uses `documentation_freshness_reviewer`; Claude Code uses `documentation-freshness-reviewer`. Fix stale or unverified documentation, examples, and comments in the same commit.

## Testing (TDD)

- For every feature or fix, write a failing behavioral test, implement the general rule, then refactor. Use realistic inputs and triangulation; never hardcode test cases or weaken tests to pass.
- Keep tests isolated and parallel; run fast unit tests first and investigate slow suites.
- Use the tracked `Cargo.lock`: pass `--locked` to build, test, and evidence commands. Update dependencies explicitly and review the lockfile diff.
- Run the full Rust suite with `cargo test --locked --workspace --profile ci`.
- Skip runtime tests for changes without runtime impact; runtime configuration changes require tests. Run applicable documentation/CI contract checks.

## Test Fixture Storage

- Do not introduce Git LFS without running `lfs-cost-advisor` and following its recommendation, including before adding `lfs: true` or `git lfs pull` to CI.
- Track every fixture referenced outside `tests/bulk_conversion_baseline.json`, including fixtures used by default tests, golden mocks, and visual audits. Bulk-only fixtures come from the release archive via `.github/actions/bulk-fixtures`.
- Before force-adding an ignored fixture, confirm a non-ignored test needs it.
- To replace the bulk corpus, use `scripts/download-third-party-fixtures.sh`, publish a new `fixtures-vN` release, and update `tag`, `sha256`, and `min-files` in `.github/actions/bulk-fixtures/action.yml`. On macOS, build the archive with `COPYFILE_DISABLE=1 tar czf` to exclude AppleDouble members.
- Declare every `mc:Ignorable` namespace prefix in scope; `fixture_package_validity.rs` checks authored packages.

## Code Style

- Comments explain constraints and reasons; use `TODO(reason)` or `FIXME(reason)` for limitations.
- Use `tracing` for key decisions, transitions, and external calls, recording context, elapsed time, and outcome. Avoid hot-path noise and sensitive data.

## Reference Projects

- Before design decisions, check `references/INDEX.md`. If no relevant reference exists, research maintained open-source implementations across languages.
- Add useful discoveries to the index and a detail file under 50 lines; cite references that informed the implementation.
- If a dependency bug or limitation breaks conversion, fix and test it upstream and open a PR following that project's conventions.

## Confidentiality

- Never publish names, paths, companies, people, titles, or other content from classified fixtures in commits, PRs, issues, or public text. Use generic descriptions such as "internal test document" or "ground truth PDF".
- Obtain user confirmation before publishing images from classified fixtures.

## Git and Pull Requests

- Use `developer0hye` for this repository's GitHub operations. Before writes, verify `gh auth status` and `gh api user --jq .login`; switch accounts and reverify if necessary. Do not use connectors with an unverified acting identity.
- Before committing, verify local `user.name=Yonghye Kwon` and `user.email=developer.0hye@gmail.com`. Use `git commit -s`; author and sign-off must match.
- All changes go through PRs. Use one focused `<type>/<short-description>` branch in a **git worktree**, based on current `main`. Work and push there; never switch branches or commit directly in the main worktree.

### PR Merge Procedure

1. Give the PR a clear description of the problem, change, and validation. Cross-reference related issues with `Related: #N`; avoid auto-close keywords unless instructed.
2. Resolve conflicts with current `main`, then wait for passing CI using `gh pr checks <number> --watch`. Do not merge failing checks.
3. Review `gh pr diff <number>` for correctness, debug code, hardcoded paths, credentials, and unused imports.
4. Merge with `gh pr merge <number> --merge`; never use `--delete-branch` while its worktree exists.
5. Sync the main worktree and verify CI on the exact merge commit before cleanup.
6. Check worktree status, merge ancestry, and open PRs. Preserve unrelated changes, active work, unmerged commits, required evidence, and shared build data.
7. In the same turn, remove the completed worktree with `git worktree remove`, delete its local branch with `git branch -d` and its remote branch with `git push origin --delete`, then prune stale records. Do not force-delete unfinished work; no additional confirmation is needed for verified cleanup.

## MSRV Policy

- `rust-version` must be a stable release at least six months old and meet the edition floor (Rust 1.85 for edition 2024). Check release dates at https://releases.rs/.
- Bump only when required by language features or dependencies. Verify code and dependency compatibility with the target, and update the `Check MSRV` toolchain in `.github/workflows/ci.yml`.

## Visual Comparison Workflow

For every visual fix or defect report, **read and follow
[the evidence contract](assets/bugfixes/README.md)** and
[the PR template](.github/pull_request_template.md). They define required images,
reports, JPEG settings, text-layer-only exceptions, and exact-ID exporter dispositions.

### Scope and provenance

- Audit **every page of every reported file**. Before changing code, record deviations using the PR template's complete deviation checklist.
- Verify GT provenance with `pdfinfo`; an attached PDF may be converter output. For user-supplied files on macOS, export from the matching Microsoft app when available. Otherwise generate and identify a LibreOffice reference; never assume it reproduces native Office behavior.
- Run `python3 scripts/check_gt_integrity.py GT.pdf [--source FILE.xlsx]` before deriving measurements. Read source XML and identify the element/attribute responsible before attributing a defect; reference-side extraction or font substitutions are not automatically converter bugs.
- Use `mutool` for geometry (`brew install mupdf-tools` on macOS). Do not quote APPROXIMATE fallback measurements in issues or PRs.

### Required comparisons

Run geometry first, then text, color/pixels, and direct image inspection:

```sh
python3 scripts/compare_layout.py GT.pdf OUTPUT.pdf --json --audit --fine-shift PT
python3 scripts/compare_text_layer.py GT.pdf OUTPUT.pdf
python3 scripts/compare_render.py GT.pdf OUTPUT.pdf --page N --dpi 150 \
  --fine-shift PT --artifacts-dir ARTIFACT_DIR --cluster-report REPORT.json
```

- Run the render comparison per page in separate artifact directories. Read each tool's interpretation and use the acting agent's image vision to inspect every full GT/output page, pixel diff, and matched crop.
- Inspect full-resolution crops at **150 DPI or higher**, including hairlines (at most 1pt), dashes, and bold/italic/underlined runs, including CJK. Record concrete `Model vision findings`; numbers alone are insufficient.
- Choose `--fine-shift PT` by materiality, at least the trace noise floor; see the layout differ's docstring for format-specific `--noise-floor` settings.
- Fix or track every page-count, missing/extra/reflow/changed-wrap, painted-text visibility, visible-fill occlusion, rectangle geometry, and fine/large text-instance shift finding. Safe split/join topology is informational; recovered fragments still require position and visibility checks.
- Geometry shifts are not antialiasing. Assess histogram **color shift**, not bin intersection, and never infer correctness from pixel counts alone. For text changes, inspect both codepoint classes and normalized content; identical pixels do not prove searchable text is correct.

### Cluster dispositions and materiality

- Follow the contract's per-ID cluster review and strict-report procedure. Re-derive dispositions from current crops; never inherit approvals by ID/bbox or rewrite historical reports.
- **One issue and PR per root cause; fix sequentially.** File separate issues for independent material deviations, with the identifying measurements.
- **Materiality:** a deviation is material if a reader notices it without an overlay. Changed wrapping/pagination, missing/overlapping/clipped text, or a wrong face/color are always material. Calibrate reference placement around a couple of points, internal alignment around a device pixel, and stroke/size differences proportionally (roughly a quarter); these are guides, not fixed gates. Record sub-material findings and their measurements under #1874 instead of creating separate issues.

### Evidence and completion

- Capture before evidence before the fix; generate after images and reports from the same GT/output PDFs. Follow the contract for publishing fix/defect evidence and classifying all remaining findings.
- Embed a commit-pinned raw URL for `compare.jpg` in every visual defect issue.
- Close an issue only after a fresh comparison proves its defect gone and every remaining material deviation has an open issue or verified exact-ID exporter disposition. Re-audit every posted after image and reference all remaining deviations in the PR.
- After each merge, repeat the full reported-file audit, file newly exposed defects, and fix them in turn. Report a file done only when a full pass finds nothing new.

### Native probes and artifacts

- Settle unclear layout rules with one-factor XML probes using `scripts/probe_harness.py`; read its docstring for the spec and backend options. Require the no-patch re-zip control to match the base; a failed control or differ invalidates the probe.
- Stage native Office input/output in the **driven app's own container on the internal disk**, never `/Volumes` or another app's container; copy evidence out afterward. See the probe docstring for paths.
- Generate internal comparison artifacts with `cargo test --locked -p office2pdf --test artifact_generator -- --ignored --nocapture`. Use the generated report to locate pages, then apply the same full-page audit and confidentiality rules.

## Unattended Issue Loop

- `scripts/issue_loop.sh` starts fresh Claude attempts (`--agent codex` for Codex); instructions live in `scripts/issue_loop_prompt.md`. Use `--expect-user developer0hye`, `autofix-skip` for trackers, and `scripts/issue_loop_watch.py` for activity. See the runner for options and stop/retry behavior.
- In headless sessions, wait for CI in the foreground and complete the PR lifecycle before ending the turn. Runner progress does not prove completion.

## Release Procedure

Follow `RELEASING.md` through publication verification and cleanup; a version bump
or GitHub Release alone is not completion.
