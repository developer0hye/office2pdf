# Project Rules

- Always communicate and work in English.
- Before starting development, check if `PRD.md` exists in the project root. If it does, read and follow the requirements defined in it throughout the development process.
- **IMPORTANT: Always prefer Rust native implementations.** Avoid unnecessary external dependencies and leverage the Rust standard library as much as possible. Only use third-party crates when there is a clear, justified need.
- **IMPORTANT: Follow Test-Driven Development (TDD).** See the **Testing (TDD)** section below for detailed rules.
- **IMPORTANT: Read and follow `METHODOLOGY.md`** before starting any task.
- When editing `CLAUDE.md`, use the minimum words and sentences needed to convey 100% of the meaning.
- After completing each planned task, run tests and commit before moving to the next task. **Skip tests if the change has no impact on runtime behavior** (e.g., docs, comments, CI config). Changes to runtime config files (YAML, JSON, etc. read by code) must still trigger tests.
- **After any code change (feature addition, bug fix, refactoring, PR merge), check if `README.md` needs updating.** If project description, usage, setup, architecture, or API changed, update `README.md` with clear, concise language. Keep it minimal — only document what users need to know.

## Testing (TDD)

- Write tests first. Follow Red-Green-Refactor: (1) failing test, (2) minimal code to pass, (3) refactor.
- Use real-world scenarios and realistic data in tests. Prefer actual use cases over trivial/contrived examples.
- **Never overfit to tests.** Implementation must solve the general problem, not just the specific test cases. No hardcoded returns, no input-matching conditionals, no logic that only handles test values. Use triangulation — when a fake/hardcoded implementation passes, add tests with different inputs to force generalization.
- Test behavior, not implementation. Assert on observable outcomes, not internal details — tests must survive refactoring.
- Every new feature or bug fix must have corresponding tests.
- **Optimize test execution speed.** Run independent tests in parallel. Use `cargo test` default parallelism. Keep each test isolated — no shared mutable state — so parallel execution is safe.
- For I/O-bound tests (network, file), prefer async or use mocks to avoid blocking. For CPU-bound tests, use multi-thread parallelism.
- If full test suite exceeds 30 seconds, investigate: split slow integration tests from fast unit tests, run unit tests first for quick feedback.
- **Skip tests when no runtime impact.** In CI/CD, use path filters to trigger tests only when source code, test files, or runtime config files are modified. Non-runtime changes (docs, README, `.md`, CI pipeline config) should not trigger test runs.

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
  - Do not delete worktrees immediately after task completion — remove only when starting new work or upon user confirmation.

## PR Merge Procedure

Follow all steps in order:

1. Rewrite PR description if empty/unclear via `gh pr edit`. Include: what changed, why, key changes, and relevant context.
2. Cross-reference related issues (`gh issue list`). Use "Related: #N" — avoid auto-close keywords unless instructed.
3. Check for conflicts. If `main` has advanced, rebase/merge as needed.
4. Wait for CI to pass: `gh pr checks <number> --watch`. Abort if tests fail.
5. Final code review via `gh pr diff <number>` — check for debug statements, hardcoded paths, credentials, unused imports.
6. Merge: `gh pr merge <number> --merge`. **Never use `--delete-branch`** (worktree depends on the branch).
7. Return to main repo, `git pull` to sync.
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

When comparing PDF output against ground truth (classified fixtures):

1. Run `cargo test -p office2pdf --test artifact_generator -- --ignored --nocapture` to generate artifacts.
2. Read `tests/classified_fixtures/_work/report.json` — contains per-file page counts, text lengths, and PNG paths.
3. Identify worst files: page count mismatches, large text length differences, conversion errors.
4. For worst files, use the **Read tool to view PNG images** in `tests/classified_fixtures/_work/<work_dir>/`:
   - `output-*.png` — rendered pages from office2pdf output
   - `gt-*.png` — rendered pages from ground truth PDF
   - `output.txt` / `gt.txt` — extracted text
5. Compare output and GT PNGs **visually** to identify specific rendering differences (layout, font, table, image, margin, page break, etc.).
6. For user-provided DOCX/XLSX/PPTX files on macOS, if Word/Excel/PowerPoint is available for that file type, export a PDF from the native Microsoft app first and compare that GT before guessing.
7. Fix root causes in parser/codegen via TDD. Prioritize high-leverage fixes that improve multiple files.

## Release Procedure

When asked to "release", always perform **both** GitHub Release and crates.io publish:

1. **Version bump** — Create a PR (`chore/publish-<version>`) that bumps `version` in both `crates/office2pdf/Cargo.toml` and `crates/office2pdf-cli/Cargo.toml`, and updates the CLI's `office2pdf` dependency version. Merge via standard PR workflow.
2. **GitHub Release** — `gh release create v<version>` with changelog and contributors section.
   - Use `git log <prev-tag>..HEAD --format='%an' | sort -u` to find contributors. List each with their GitHub profile link.
3. **crates.io publish** — Publish lib first, then CLI:
   - `cargo publish -p office2pdf`
   - `cargo publish -p office2pdf-cli`
4. **Tag alignment** — Ensure the GitHub release tag (`v<version>`) and Cargo.toml versions match.
