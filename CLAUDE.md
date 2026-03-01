# Project Rules

- Always communicate and work in English.
- Before starting development, check if `PRD.md` exists in the project root. If it does, read and follow the requirements defined in it throughout the development process.
- **IMPORTANT: Always prefer Rust native implementations.** Avoid unnecessary external dependencies and leverage the Rust standard library as much as possible. Only use third-party crates when there is a clear, justified need.
- **IMPORTANT: Follow Test-Driven Development (TDD).** Always write tests first before implementing functionality. Follow the Red-Green-Refactor cycle: (1) Write a failing test, (2) Write the minimal code to make it pass, (3) Refactor while keeping tests green. Every new feature or bug fix must have corresponding tests.
- **IMPORTANT: Read and follow `METHODOLOGY.md`** before starting any task.
- When editing `CLAUDE.md`, use the minimum words and sentences needed to convey 100% of the meaning.
- After completing each planned task, run tests and commit before moving to the next task.

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

## Release Procedure

When asked to "release", always perform **both** GitHub Release and crates.io publish:

1. **Version bump** — Create a PR (`chore/publish-<version>`) that bumps `version` in both `crates/office2pdf/Cargo.toml` and `crates/office2pdf-cli/Cargo.toml`, and updates the CLI's `office2pdf` dependency version. Merge via standard PR workflow.
2. **GitHub Release** — `gh release create v<version>` with changelog and contributors section.
   - Use `git log <prev-tag>..HEAD --format='%an' | sort -u` to find contributors. List each with their GitHub profile link.
3. **crates.io publish** — Publish lib first, then CLI:
   - `cargo publish -p office2pdf`
   - `cargo publish -p office2pdf-cli`
4. **Tag alignment** — Ensure the GitHub release tag (`v<version>`) and Cargo.toml versions match.
