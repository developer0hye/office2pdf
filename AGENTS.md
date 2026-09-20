# Agent Rules

This file is the canonical repository instruction source for all coding agents.

- Always communicate and work in English.
- Before starting development, check if `PRD.md` exists in the project root. If it does, read and follow the requirements defined in it throughout the development process.
- **IMPORTANT: Always prefer Rust native implementations.** Avoid unnecessary external dependencies and leverage the Rust standard library as much as possible. Only use third-party crates when there is a clear, justified need.
- **IMPORTANT: Follow Test-Driven Development (TDD).** See the **Testing (TDD)** section below for detailed rules.
- **IMPORTANT: Read and follow `METHODOLOGY.md`** before starting any task.
- When editing `AGENTS.md` or `CLAUDE.md`, use the minimum words and sentences needed to convey 100% of the meaning.
- After completing each planned task, run tests and commit before moving to the next task. **Skip tests if the change has no impact on runtime behavior** (e.g., docs, comments, CI config). Changes to runtime config files (YAML, JSON, etc. read by code) must still trigger tests.
- **After any code change (feature addition, bug fix, refactoring, PR merge), check if `README.md` needs updating.** If project description, usage, setup, architecture, or API changed, update `README.md` with clear, concise language. Keep it minimal — only document what users need to know.
- **Before every commit, delegate a read-only freshness audit and wait for `PASS:`.** Codex must use `documentation_freshness_reviewer`; Claude Code must use `documentation-freshness-reviewer`. Compare existing documentation, examples, and code comments with the current code and configuration; update or remove stale versions, commands, APIs, behavior, defaults, paths, architecture, limitations, and unverified claims in the same commit.

## Testing (TDD)

- Write tests first. Follow Red-Green-Refactor: (1) failing test, (2) minimal code to pass, (3) refactor.
- Use real-world scenarios and realistic data in tests. Prefer actual use cases over trivial/contrived examples.
- **Never overfit to tests.** Implementation must solve the general problem, not just the specific test cases. No hardcoded returns, no input-matching conditionals, no logic that only handles test values. Use triangulation — when a fake/hardcoded implementation passes, add tests with different inputs to force generalization.
- Test behavior, not implementation. Assert on observable outcomes, not internal details — tests must survive refactoring.
- Every new feature or bug fix must have corresponding tests.
- **Optimize test execution speed.** Run independent tests in parallel. Use `cargo test` default parallelism. Keep each test isolated — no shared mutable state — so parallel execution is safe.
- **Use the reviewed dependency graph.** The workspace `Cargo.lock` is tracked. Pass `--locked` to build, test, and evidence commands; update dependencies explicitly and review the lockfile diff.
- For I/O-bound tests (network, file), prefer async or use mocks to avoid blocking. For CPU-bound tests, use multi-thread parallelism.
- If full test suite exceeds 30 seconds, investigate: split slow integration tests from fast unit tests, run unit tests first for quick feedback.
- **Skip tests when no runtime impact.** In CI/CD, use path filters to trigger tests only when source code, test files, or runtime config files are modified. Non-runtime changes (docs, README, `.md`, CI pipeline config) should not trigger test runs.

## Test Fixture Storage

**This repository uses no Git LFS.** LFS bandwidth is billed to the repo owner for every clone and fork — including forks we do not control and cannot rate-limit — so the corpus was moved off it. Release asset downloads are not metered.

- **IMPORTANT: Run the `lfs-cost-advisor` agent and follow its recommendation before adding any file to Git LFS, or adding `lfs: true` / `git lfs pull` to a workflow job.** Reintroducing LFS restores the fork-billing exposure this split removed.
- Fixtures the default `cargo test` suite reads are **tracked in git normally**. Ordinary git objects are not metered, so a new fixture under a few MB just gets committed.
- The bulk corpus (2,695 files) lives in the `fixtures-v1` release asset and is fetched by `.github/actions/bulk-fixtures`, which verifies the archive checksum and that every archived file survived extraction. Only the `#[ignore]`d gate in `bulk_conversion.rs` needs it.
- `.gitignore` covers `tests/fixtures/*/libreoffice/*` and `tests/fixtures/*/poi/*` so extracted bulk files stay untracked. Committing a fixture there needs `git add -f`, which is the intended friction — confirm the default suite actually reads it first.
- Regenerate the corpus with `scripts/download-third-party-fixtures.sh`, then publish a new `fixtures-vN` release and bump the `tag`, `sha256`, and `min-files` inputs in `.github/actions/bulk-fixtures/action.yml`. Build the archive with `COPYFILE_DISABLE=1 tar czf` on macOS, or it gains `._` xattr members that macOS `tar` hides from listings but Linux extracts as bogus `.docx` fixtures.
- A fixture referenced anywhere outside `tests/bulk_conversion_baseline.json` — including `tests/golden_mocks/**` and `tests/visual_audits/**`, which non-`#[ignore]`d tests read — must be tracked, not left to the release asset.

## Logging

- Add structured logs at key decision points, state transitions, and external calls — not every line. Logs alone should reveal the execution flow and root cause.
- Include context: request/correlation IDs, input parameters, elapsed time, and outcome (success/failure with reason).
- Use appropriate log levels: `error!` for failures requiring action, `warn!` for recoverable issues, `info!` for business events, `debug!`/`trace!` for development diagnostics.
- Use the `tracing` crate for structured, async-safe logging. Prefer `tracing::instrument` for automatic span creation.
- Never log sensitive data (credentials, tokens, PII). Mask or omit them.
- Avoid excessive logging in hot paths — logging must not degrade performance or increase latency noticeably.

## Naming

- Names must be self-descriptive — understandable without reading surrounding code. Avoid cryptic abbreviations (`proc`, `mgr`, `tmp`).
- Prefer clarity over brevity, but don't over-pad. `user_email` > `e`, `calculate_shipping_cost` > `calc`.
- Booleans should read as yes/no questions: `is_valid`, `has_permission`, `should_retry`.
- Functions/methods should describe the action and target: `parse_config`, `send_notification`, `validate_input`.

## Types

- Prefer explicit type annotations over type inference. Annotate function signatures (parameters and return types) always.
- Annotate variables when the type isn't obvious from the assigned value.
- Use newtypes to enforce domain semantics (e.g., `struct Emu(f64)` instead of bare `f64`).

## Comments

- Explain **why**, not what. Code already shows what it does — comments should capture intent, constraints, and non-obvious decisions.
- Comment business rules, workarounds, and "why this approach over the obvious one" — context that can't be inferred from code alone.
- Mark known limitations with `TODO(reason)` or `FIXME(reason)` — always include why, not just what.
- Delete comments when the code changes — outdated comments are worse than no comments.

## Reference Projects

- When facing design decisions or implementation challenges, first check if `references/INDEX.md` exists and find relevant reference projects.
- If no relevant project exists in `references/`, search the web for well-maintained open-source projects that solve similar problems. Search across all languages — architectural patterns transfer regardless of language.
- When a new useful project is discovered and `references/` exists, add it to `references/INDEX.md` and create a corresponding detail file. Keep detail files under 50 lines.
- Cite which reference project informed your approach when applying patterns from it.
- If a dependency limitation or bug breaks PDF conversion, clone that library, fix and test it upstream, and open a PR. Follow its repository conventions and match the tone and scope of its recently merged PRs.

## Confidentiality

- **NEVER mention `tests/classified_fixtures/` content** (file names, paths, company names, personal names, document titles) in commit messages, PR titles/descriptions, issue comments, or any public-facing text.
- Use generic references instead: "classified fixture", "internal test document", "ground truth PDF", etc.

## Git Configuration

- All commits must use the local git config `user.name` and `user.email`. Verify with `git config user.name` and `git config user.email` before committing.
- All commits must include `Signed-off-by` line (always use `git commit -s`). The `Signed-off-by` name must match the commit author.

## Branching & PR Workflow

- All changes go through pull requests. No direct commits to `main`.
- Branch naming: `<type>/<short-description>` (e.g., `feat/add-parser`, `fix/table-bug`).
- One branch = one focused unit of work.
- **Use git worktrees** for all branch work. Do not use `git checkout`/`git switch` in the main repo.
  - Create: `git worktree add ../<repo-name>-<branch-name> -b <type>/<short-description>`
  - Work and push from inside the worktree.
  - Once the task is verified complete, its PR is merged, and exact-merge CI passes on `main`, remove its worktree and local/remote topic branches in the same turn. Follow the preservation checks below; no additional confirmation is needed.

## Worktree and Branch Cleanup

- After a task is verified complete, its PR is merged, and CI passes on the exact merge commit on `main`, remove its worktree and delete its local and remote topic branches in the same turn. No additional confirmation is needed.
- Before cleanup, check `git worktree list`, worktree status, merge ancestry, and open PRs. Preserve active work, uncommitted changes, and commits not merged into `main`.
- Preserve required evidence and shared build data outside the worktree before deleting it. Use `git worktree remove` and `git branch -d`; do not force-delete unfinished work. Prune stale worktree and remote-tracking records afterward.

## PR Merge Procedure

Follow all steps in order:

1. Rewrite PR description if empty/unclear via `gh pr edit`. Include: what changed, why, key changes, and relevant context.
2. Cross-reference related issues (`gh issue list`). Use "Related: #N" — avoid auto-close keywords unless instructed.
3. Check for conflicts. If `main` has advanced, rebase/merge as needed.
4. Wait for CI to pass: `gh pr checks <number> --watch`. Abort if tests fail.
5. Final code review via `gh pr diff <number>` — check for debug statements, hardcoded paths, credentials, unused imports.
6. Merge: `gh pr merge <number> --merge`. **Never use `--delete-branch`** (worktree depends on the branch).
7. Return to main repo, `git pull` to sync; verify completion and CI on the exact merge commit, then preserve required evidence and shared build data outside the worktree.
8. Remove worktree: `git worktree remove ../<repo-name>-<branch-name>`
9. Delete local branch: `git branch -d <branch-name>`
10. Delete remote branch: `git push origin --delete <branch-name>`

## MSRV Policy — 6-Month Rolling Minimum

This project follows a **6-month rolling MSRV policy** (aligned with [tokio](https://crates.io/crates/tokio) and other major crates):

- The `rust-version` in `Cargo.toml` MUST target a Rust stable release that was published **at least 6 months ago**
- Rust stable releases ship every 6 weeks — consult [releases.rs](https://releases.rs/) for exact dates
- When a newer Rust version crosses the 6-month threshold, updating the MSRV is **allowed but not required** — only bump when a newer language feature or dependency demands it
- **Floor:** the MSRV can never go below the minimum required by `edition` in `Cargo.toml` (edition 2024 = Rust 1.85)

**Before any MSRV change:**
1. Verify no language features or APIs exclusive to versions above the target are used
2. Confirm all dependencies compile on the target version (`cargo check` with the target toolchain, or review dependency MSRV metadata)
3. Update CI matrix to include the new MSRV version

## Visual Comparison Workflow

- For visual bug fixes tied to an issue, keep `assets/bugfixes/issue-<number>/gt.jpg`, `before.jpg`, `after.jpg`, `layout-audit.json`, and one `render-clusters-page-<page>.json` per compared page from the same fixture, page set, renderer, and resolution; `before.jpg` captures the pre-fix output, while `after.jpg` and the reports use the current output. Every fix PR with a raster change must change `after.jpg`, `layout-audit.json`, and every page's strict cluster report. Generate the layout report from the same GT and current-output PDFs used for `gt.jpg` and `after.jpg` with `compare_layout.py --json --audit --fine-shift PT`; record the same fine-detail threshold in the PR and classify page-count, missing/extra/reflow or changed-wrap, painted-text visibility, visible-fill occlusion, rectangle geometry, and fine/large-shift failures with open issues in the `Visual audit` fields and matching `Remaining:` rows. A reported safe split/join `topology` group is informational rather than a text-flow failure, but each recovered fragment's position and visibility findings remain auditable. Generate each render report with `compare_render.py --cluster-report PATH --cluster-dispositions PATH --strict-clusters`; every cluster ID must map to an accepted renderer class or an open issue, and all report paths must be listed in `Visual audit`. A verified GT-exporter difference may instead use the exact-ID contract documented in `assets/bugfixes/README.md`; it requires `reference-exporter-differences.json`, a hashed full-page `native.jpg`, structured source/reference/native provenance, and a stable verification comment. Use progressive JPEG quality 86 with metadata stripped, preserve the source pixel dimensions, and verify text and images remain legible for direct GitHub links. Strip first, then re-set the density (`-strip -density 150 -units PixelsPerInch`): the contract check rejects evidence recording under 150 DPI, which stripping alone removes.
- A text-layer-only fix may correctly produce byte- and pixel-identical `before.jpg` and `after.jpg`. Regenerate and verify the current `after.jpg` without adding annotations; it need not appear in the diff when the render is unchanged. Set `Text-layer-only: Yes` and `Pixel delta: 0` in the PR's `Visual audit`, change `layout-audit.json`, and require that report to show the missing/extra searchable-text finding resolved. The contract requires both evidence files and uses ImageMagick's exact decoded-pixel comparison to verify that the declared zero delta is true.
- **When filing a visual defect issue, attach a side-by-side image (GT left, office2pdf output at filing time right)** rendered from the same page and resolution, committed as `assets/bugfixes/issue-<number>/compare.jpg` (same JPEG rules as above) and embedded in the issue body via a commit-pinned raw URL. For classified fixtures, confirm with the user before publishing the image; the surrounding issue text must still follow the Confidentiality rules.

### Visual check discipline (harness rules)

- **Enumerate before fixing.** For every compared page, walk this checklist and record each deviation before touching code: page count/order; element presence; position; size; rotation/flip; fill; stroke/border (incl. dash style); shape outline geometry (corner rounding, curved edges — a panel drawn as its bounding box, #1029); text content; font family/weight/style; text color; alignment; line/paragraph spacing; clipping/overflow.
- **One issue per root cause.** When one image reveals multiple independent defects, file a separate issue for each — never bundle them into one issue or one PR. Fix them sequentially.
- **Closing condition.** An issue may be closed only when a fresh GT comparison shows its specific defect gone. Every remaining visible deviation on that comparison must already have its own open issue or a verified exact-ID reference-exporter disposition — file the missing issue for every other unresolved deviation before closing.
- **After images are re-audited.** When posting an after image, re-run the checklist on it; each still-visible deviation gets an issue reference in the PR body ("remaining, tracked in #N") or the verified reference-exporter identifier (`ref:<id>`).

### Fine-level difference analysis (reported files)

When a user reports files, audit **every page of every file** to this depth, not
just the pages a screenshot shows.

1. **Check the reference's provenance first.** `pdfinfo x.pdf | grep -Ei
   'producer|creator'`. A PDF attached beside a source file is often *our own*
   output — office2pdf reports `Creator: Typst 0.15.x`, LibreOffice reports
   `Producer: LibreOffice`. Two PDFs sharing embedded font subset tags came from
   the same producer. Generate a real reference with `soffice --headless
   --convert-to pdf` when in doubt.
2. **Run all four axes**, not one: `compare_layout.py` (geometry, per page),
   `compare_text_layer.py` (what a reader can select), `compare_render.py`
   (colour and pixels), and a page-by-page visual at >=150 DPI.
   Run the geometry axis with `--audit --fine-shift PT`; every fine/large
   text-instance shift, rectangle geometry deviation, painted-text visibility
   mismatch, or visible-fill occlusion it names must be fixed or recorded as a
   remaining deviation with an open issue.
   Do not classify the pixel diff as antialiasing while this audit is failing.
   Source/XML inspection and numeric reports only route attention; they never
   constitute a visual pass. The acting Codex or Claude agent must use its image
   vision to open and inspect the full GT/output pages, the pixel diff, and all
   matched region crops before completing the audit. Record the concrete visual
   observations in `Model vision findings`; numeric output or `None` does not
   satisfy the PR contract.
3. **Read the source XML before attributing a deviation.** Name the element and
   attribute that produced it; a measurement without one is a guess.
4. **A token the reference has and we lack is not automatically our defect.**
   LibreOffice fragments words itself (`Gullfi sk`, `eff ektivt`,
   `C O N T O S O`). Check which side is wrong before filing.
5. File **one issue per root cause**, each with the numbers that identify it.

**Iterate until a pass finds nothing new.** Every fix changes what is visible
underneath it, so after each merge re-run the whole audit, file what the fix
uncovered, and fix that in turn. Report a file "done" only after a pass over it
produces no new finding.

### Three-axis comparison (`scripts/compare_render.py`)

Run `python3 scripts/compare_render.py <GT.pdf> <output.pdf> [--page N]` before
judging a rendered difference. It reports geometry, colour histogram, and pixel
difference together, then states what the combination means.
For a visual audit, also pass `--fine-shift PT --artifacts-dir <directory>`: for
the page chosen by `--page`, it preserves the full GT/output pages, their side-by-side image,
the 5% pixel diff, and a matched GT/output crop for every text-instance shift
past the active fine or coarse gate on that page. Run it once for every compared page, using a distinct
directory per page. It fails if ImageMagick is unavailable rather than silently
omitting evidence. Open every emitted path with Codex/Claude image vision;
producing the files is not itself inspection.
Also write one machine report per page with `--cluster-report PATH`. First use
its stable IDs to build explicit disposition groups, then rerun with
`--cluster-dispositions PATH --strict-clusters`. Accepted renderer-only classes
are glyph-edge rasterization, photo resampling, gradient rasterization, and
shape-edge antialiasing; otherwise use an open issue. Page-, region-, or
bbox-wide approvals are invalid, so every new cluster requires review. A
bounded `renderer_observations` entry may record inspected fragments below the
material-cluster floor, but never dispositions a cluster inside that bbox.
For a proven non-native GT difference, pass the issue-local manifest with
`--reference-differences`; its `reference-exporter-difference` group must name
the manifest's exact cluster set. The PR body names each ID as
`Reference difference: ref:<id>`. Layout `ref:<id>` dispositions are limited to
an exact painted-text visibility occurrence; every other finding needs an open
issue. See `assets/bugfixes/README.md` for the required hashes and fields.

For layout defects, run `python3 scripts/compare_layout.py <GT.pdf> <output.pdf> --audit --fine-shift PT`
first: it matches text lines from `mutool` traces and reports missing/extra/
re-wrapped lines, informational safe split/join topology groups, spatial-anchor
dy, pitch and width drift, painted-text visibility,
visible-fill occlusions, and geometry-aware axis-aligned rectangle/line
x/y/size/edge deltas, with GT noise floors built in (`--noise-floor 0.12` Word,
`0.5` Excel). Painted visibility uses trace order and flat-background contrast:
text covered by a later opaque image, single closed axis-aligned rectangle, or
fully extended shading under such a rectangular clip is `hidden`; same-colour
text on an opaque flat fill is `low_contrast`, and text that remains visible is
`painted`. The harness retains the page media box and rectangular clip stack,
intersects each glyph's conservative ink box with those visible bounds, and
keeps any partially visible run auditable. A trace line proven fully hidden by
the visibility analysis — including bounds, transparency, absent path ink, or
later opaque paint — is still reported under
`visibility.unmatched_hidden_*`, but when it has no counterpart it does not
become a visual missing/extra failure. An exact counterpart remains matched so
painted-versus-hidden differences still fail. An `ignore_text` record is
normally hidden. The harness promotes it
to painted geometry only when a compact, intersecting `fill_path` appeared
since the preceding text operation and lies within half an em of the glyph's
conservative bounding box. This trace-order heuristic recovers path-painted
text seen in issue #1407 without claiming to identify a Type3 program;
pathless invisible/OCR text stays hidden, and ambiguous cases still require
the pixel-difference and visual-inspection passes. Visible-fill analysis compares
paint order and final overlap colour where a later opaque rectangle cuts into a
thin earlier rule. Point and area floors reject trace slivers, and same-colour
operation splitting is ignored.
Touching same-paint rectangle/line operations are merged before geometry
matching; non-rectangular path bounds and raw draw counts remain informational.
Matched position, size, and edge deltas use the active fine or coarse shift
threshold and fail the audit.
Horizontal text uses its true baseline; a rotated or skewed `fill_text` stays
one visual run and uses the minimum fully transformed glyph x/y as its
comparable anchor. Its numbers are assertable; pixel counts are only a
tripwire. Repeated strings are matched as separate spatial instances: a chart
title and legend both named `Sales` appear as `Sales [1/2]` and `Sales [2/2]`,
with their own `dx`/`dy`. The coarse audit exits nonzero when any instance moves
more than 5pt (override with `--large-shift PT` for a justified fixture-specific
threshold). Add `--fine-shift PT` for high-DPI evidence; it keeps that coarse
report and also fails on matched text beyond the recorded fine tolerance. The
fine threshold must be at least the trace noise floor. Raster-only edge or
colour changes may be font antialiasing, but trace-derived x/y anchor movement
is geometry and must be fixed or tracked. Painted versus hidden/low-contrast
changes also fail; inspect and track every named instance before marking
alignment or text flow as matching.

**Its width column trims leading and trailing whitespace glyphs.** Their
advances position no visible ink and previously invented large width deltas
when only one exporter retained a terminal space (#1482). Internal spaces still
contribute through the following glyph's origin. Compare per-glyph advances
before attributing a remaining width delta to the font.

**Install `mupdf-tools` first** (`brew install mupdf-tools`). Without `mutool`
the geometry axis falls back to `pdftotext -bbox`, whose `yMin` is each glyph's
font-descriptor box rather than its baseline. The two PDFs always embed
different subsets, so that fallback carries an error proportional to font size —
on the newsletter mock it reported +2.90pt for a 22pt heading whose baseline is
really 1.07pt the other way, and issue #501 was filed against a defect that did
not exist. The report labels itself APPROXIMATE when it falls back; do not quote
those numbers in an issue or PR.

No single measure is reliable, and each one's blind spot is another's strength:

| Axis | Catches | Blind to |
| --- | --- | --- |
| Geometry (`mutool draw -F trace`) | position, size, row pitch, pagination | colour, missing elements |
| Colour histogram | fill colour, recolouring, missing elements, ink coverage | position, size, font |

Judge the colour axis on the reported **colour shift**, not on the bin-wise
intersection printed beside it. Intersection punishes a one-level shift as hard
as a recolour, and renderers dither smooth gradients by a channel step or two:
three PPTX decks score 0.9745-0.9860 there while being pixel-identical to
within +-2 per channel. The shift compares cumulative distributions, so that
noise reaches 0.0003 while the half-width borders of #487 reached 0.0016.
| Pixel difference (`AE`, `RMSE`) | whatever the other two were not watching | see below |



**Never conclude from the pixel count alone.** Measured on this corpus: two
shadow variants that look plainly different both scored an identical 173,524
at `AE -fuzz 5%`; a correct anchor-height fix made the count *rise* because the
shape was still displaced by an unrelated defect; and a column-width fix that
cut the error from 86pt to 4.4pt barely moved it, because re-proportioning
columns changes where pixels are, not how many differ. Use `RMSE` or
`AE -fuzz 1%` when magnitude matters, and treat all of them as a regression
tripwire rather than evidence.

The script's **Diff clusters** section localises that count: contiguous mask
regions with bounding boxes in points, largest first. Disposition every listed
cluster — an accepted rendering difference, a verified exact-ID reference-exporter
difference, or an issue reference. The two
~5,100pt² corner clusters of #1029 sailed through a bare 84,400-pixel count;
the census names them and their page region, and text-line geometry cannot see
them at all.

**Text layer (`scripts/compare_text_layer.py`)** is a fourth axis none of the
three above can see. Some defects change what a reader can select and search
for while leaving every pixel identical: injected U+2060/U+00A0 (#664), or a
ligature swallowing letter-spacing so the run extracts as `o ffi c e 2 p d f`
(#684 — measured 24 occurrences in the pre-fix output, against 0 of the
`ofce2pdf` that issue's body reports). It reports a codepoint-class census and a normalized-content check
separately — a class delta with matching content is an encoding difference, a
content residual needs review for extraction order or missing/extra text.
Costs milliseconds; run it on any fix that
touches text.

The tool's `## Reading` section is the part to act on — it routes attention to
the defect class (pagination, line advance, indent, fill, size) instead of
leaving a bare number to be over-interpreted.

### One-factor probe harness (`scripts/probe_harness.py`)

To settle a layout rule, don't compare corpus files — patch one factor:
`python3 scripts/probe_harness.py SPEC.json [--backend office2pdf|soffice|office]`
takes a base fixture plus one-factor XML patches (JSON spec; see the script
docstring), builds a variant package per patch and a no-patch re-zip control,
exports the batch, and runs `compare_layout.py` into one table — one row per
variant, keyed by the factor value. The control is a hard gate: if its export
deviates from the base's, the run aborts, because repackaging (not the factor)
changed the output. Differ failures abort likewise, never an empty row.
Backend `office2pdf` answers "what does our code do" anywhere in seconds
(`--converter` names the binary); `office` drives the native AppleScripts and
answers "what does Office do".

The `office` stage is constrained by the app sandbox on both sides. It defaults
into the driven app's own container — `~/Library/Containers/com.microsoft.Word`,
`.Powerpoint` or `.Excel` under `Data/probes/`, matched to the fixture's format
— and the report and PDFs are copied back to `target/probes/` afterwards.
Anywhere else, including another Office app's container, costs a per-file
"Grant Access" dialog that stalls an unattended run. An explicit `--stage-root`
overrides that and must still sit on the internal disk: the sandbox cannot
write to `/Volumes` at all (AppleEvent timeout -1712), which the harness
enforces.

The same container rule binds every other caller of those AppleScripts, and
every one of them now follows it. The `GENERATE_MICROSOFT_GT=1` exports in
`public_visual_audit.rs` and `scripts/measure_powerpoint_chart_axis.py` stage
the fixtures and the PDFs in the driven app's container and copy the results
out afterwards. `scripts/macos/export_business_golden_pdfs.sh` drives three
apps, so it uses three stages — `Data/business-golden-export/` in each app's
own container, removed when the run ends — and copies the PDFs back to its
stage-root argument, which only the unsandboxed `pdfunite`/`pdfinfo` steps read
(#1128).

### Fine-detail analysis (thin and small elements)

Whole-page thumbnails at 80 DPI hide hairlines, dash patterns, font weight, and sub-pixel offsets. For every compared page:

1. **High-DPI pass.** Render both sides at ≥150 DPI (`pdftoppm -r 150`) before judging any checklist item involving stroke width, dash style, font weight/italic, or small glyphs. Never mark those items "OK" from an 80 DPI image.
2. **Region crops.** For each region containing text, lines, or decorations, cut matched crops from GT and output (`magick input.png -crop WxH+X+Y crop.png`) and view them side by side at full scale.
3. **Pixel-difference sweep.** Run `magick compare -metric AE -fuzz 5% gt.png out.png diff.png` on size-normalized pages; view `diff.png` and inspect every highlighted cluster. A checklist pass is complete only when each cluster is either explained by an accepted rendering difference (fonts/antialiasing), dispositioned by a verified exact-ID reference-exporter difference, or captured as an issue.
4. **Hairline inventory.** Explicitly enumerate elements ≤1pt (rules, underlines, dashed/dotted lines, borders, tick marks) found in GT and confirm each exists in the output at matching position, width, and dash pattern.
5. **Weight/emphasis inventory.** Enumerate bold/italic/underlined runs visible in GT (including CJK) and confirm the same emphasis in the output — weight differences must be checked on the high-DPI crops, not thumbnails.

`magick` is ImageMagick 7 only. On ImageMagick 6 drop it: `convert input.png -crop ...`, `compare -metric AE ...`. `scripts/compare_render.py` resolves this itself.

**A GT primitive is not always the quantity you want.** Native exports draw
things we draw differently, and comparing the wrong quantity has produced fixes
that were exactly backwards:

- **Sprite box vs ink.** Excel prints iconSet icons as `fill_image` sprites. The
  placement box is 11x11pt; extracting the bitmap and measuring its non-white
  ink gives 10.08x11.00pt. Sizing our vector glyph to the box oversizes it by a
  whole sprite's padding (#651).
- **Rasterised shapes.** PowerPoint flattens shapes carrying shadows to
  bitmaps, each ~18pt larger than its shape and offset by the blur margin. An
  image-count comparison then reports phantom missing pictures — three of them
  on one slide whose XML declares a single `p:pic` (#666).
- **Filled rects vs strokes.** Word draws table borders as filled rectangles
  with the outer edge on the cell boundary, emitting no `stroke_path` at all;
  we stroke a line centred on it. Border geometry has to be compared as fills
  (#724).

**Charts have no committed ground truth.** No workbook or deck under
`tests/golden_mocks/business/sources/` contains a `c:lineChart`, `barChart`,
`pieChart`, `doughnutChart`, `areaChart` or `scatterChart`. Every chart issue
therefore rests on figures nobody can re-derive from the repo, and at least one
(#636) points the wrong way when measured: our plot line is 2.0pt against a
reference's 2.235pt, not "roughly twice" it. Treat a chart GT number as
unverified until a native export is attached.

The auto-scaled value axis is the one exception:
`python3 scripts/measure_chart_axis.py <data maxima>` rescales `WithChart.xlsx`
to each maximum and reads the tick labels back off a render, so those figures
are reproducible without a native export (#634).

**Validate the GT before deriving anything from it** —
`python3 scripts/check_gt_integrity.py GT.pdf [--source FILE.xlsx]`. A corrupt
export poisons every downstream number silently: the workbook-per-sheet
duplication of #616 survived because nothing checked the GT itself. The gate
detects page-sequence periodicity (the duplication signature), implausible page
counts against the worksheets the source would actually print (hidden sheets
excluded, #1211), and GT-side font substitutions —
the last so an exporter's Arial-to-Liberation swap is never filed as a converter
defect. Exits non-zero only when the GT cannot be trusted at all.

When comparing PDF output against ground truth (classified fixtures):

1. Run `cargo test --locked -p office2pdf --test artifact_generator -- --ignored --nocapture` to generate artifacts.
2. Read `tests/classified_fixtures/_work/report.json` — contains per-file page counts, text lengths, and PNG paths.
3. Identify worst files: page count mismatches, large text length differences, conversion errors.
4. For worst files, use the **Read tool to view PNG images** in `tests/classified_fixtures/_work/<work_dir>/`:
   - `output-*.png` — rendered pages from office2pdf output
   - `gt-*.png` — rendered pages from ground truth PDF
   - `output.txt` / `gt.txt` — extracted text
5. Compare output and GT PNGs **visually** to identify specific rendering differences (layout, font, table, image, margin, page break, etc.).
6. For user-provided DOCX/XLSX/PPTX files on macOS, if Word/Excel/PowerPoint is available for that file type, export a PDF from the native Microsoft app first and compare that GT before guessing.
7. Fix root causes in parser/codegen via TDD. Prioritize high-leverage fixes that improve multiple files.

## Unattended Issue Loop

`scripts/issue_loop.sh` works through the open issues without supervision: it picks the
oldest open issue, starts a **fresh** `claude -p` for it, and repeats. One agent per
issue is the point — a single long session fills its context, auto-compaction drops the
early instructions, and quality decays; a new process starts empty every time, with
GitHub holding the state (the open-issue list is the queue, a closed issue or a merged
PR is the completion mark).

```sh
scripts/issue_loop.sh --once --max-issues 1   # trial one issue
scripts/issue_loop.sh --model claude-opus-5   # then run it unattended
```

`scripts/issue_loop_prompt.md` carries the repository-specific instructions appended to
every agent's prompt; edit it rather than the script when the rules change. The agent
inside a run is headless, so it must never background the CI wait — the process dies
when its turn ends, taking the watch with it.

Guardrails, each of which has already caught a real failure: two attempts without a
merged PR label an issue `autofix-blocked` and take it out of the queue; three
consecutive no-progress runs trip a circuit breaker; a usage limit pauses for an hour
instead of counting as failure; and a disk-space floor stops the loop before the
filesystem fills. Label an umbrella issue `autofix-skip` so no agent tries to "fix" a
tracker that closes with its children.

Because a run reports only when its agent exits, use `scripts/issue_loop_watch.py` to
see what the current agent is doing.

## Release Procedure

- Follow `RELEASING.md` end to end in one turn; a version bump or GitHub Release alone is not completion.
- Perform every GitHub operation for this repository as `developer0hye`. Ignore stale shell tokens with `GH_TOKEN=''` and verify `gh api user --jq .login` before the first write.
- After the version PR merges, start releases only by dispatching `release.yml` once with the tag; do not manually create a normal release. Monitor that exact run. Completion requires a green release run, tag/version alignment, both crates.io packages, and all six binary assets.
