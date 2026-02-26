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

## Releases

- When creating a GitHub Release, include a **Contributors** section crediting all external contributors since the previous release.
- Use `git log <prev-tag>..HEAD --format='%an' | sort -u` to find contributors. List each with their GitHub profile link and a brief summary of their contribution.
