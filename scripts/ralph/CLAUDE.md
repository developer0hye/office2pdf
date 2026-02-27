# Ralph Agent Instructions for office2pdf

You are an autonomous coding agent working on the **office2pdf** project — a pure Rust library and CLI tool that converts DOCX, XLSX, and PPTX files to PDF using the Typst engine.

## Project Context

- **Project root**: The current working directory
- **Language**: Rust (edition 2024, MSRV 1.89)
- **Architecture**: Input file → Parser → IR (Intermediate Representation) → Typst Codegen → Typst Compile → PDF
- **Workspace**: Two crates — `crates/office2pdf` (library) and `crates/office2pdf-cli` (CLI wrapper)

Read these project files for full context:
- `PRD.md` — Full product requirements
- `CLAUDE.md` — Development guidelines and conventions
- `METHODOLOGY.md` — Software methodology principles

## Your Task

1. Read the PRD at `scripts/ralph/prd.json`
2. Read the progress log at `scripts/ralph/progress.txt` (check Codebase Patterns section first)
3. Check you're on the correct branch from PRD `branchName`. If not, create it from `main`.
4. Pick the **highest priority** user story where `passes: false`
5. **Write tests first** (TDD: Red-Green-Refactor cycle)
6. Implement that single user story
7. Run quality checks (see below)
8. If checks pass, commit ALL changes with message: `feat: [Story ID] - [Story Title]`
9. Update the PRD to set `passes: true` for the completed story
10. Append your progress to `scripts/ralph/progress.txt`

## Quality Check Commands

Run ALL of these before committing. ALL must pass:

```bash
cargo fmt --all -- --check        # Format check
cargo clippy --workspace -- -D warnings  # Lint check
cargo test --workspace            # All tests
cargo check --workspace           # Compile check
```

If any check fails, fix the issues before committing.

## Rust Development Rules

- **Prefer Rust native implementations.** Avoid unnecessary external dependencies. Use the standard library as much as possible.
- **Follow TDD.** Write failing tests first, then minimal implementation, then refactor.
- **Only add third-party crates when clearly justified** (e.g., `typst` for PDF rendering, `docx-rs` for DOCX parsing).
- When adding a new dependency, add it to the appropriate `Cargo.toml` (workspace root or crate-level).
- Follow existing code patterns — check how existing modules are structured before adding new ones.
- Keep all public APIs documented with doc comments.

## Git Rules

- Verify git config before first commit: `git config user.name` and `git config user.email`
- All commits must include `Signed-off-by` line: use `git commit -s`
- Commit message format: `feat: [Story ID] - [Story Title]`
- Commit only code that passes ALL quality checks

## Key Dependencies (approved for use)

| Crate | Purpose |
|---|---|
| `typst`, `typst-pdf`, `typst-kit` | Layout engine and PDF output |
| `docx-rs` | DOCX parsing |
| `umya-spreadsheet` | XLSX parsing |
| `zip`, `quick-xml` | PPTX parsing (direct ZIP+XML, ppt-rs is write-only) |
| `serde`, `serde_json` | docx-rs private field extraction |
| `comemo` | Required by Typst World trait |
| `thiserror` | Library error types |
| `anyhow` | CLI error handling |
| `clap` | CLI argument parsing |

Do NOT add dependencies beyond these unless absolutely necessary. If you must add one, document why in the progress log.

## Existing Code Overview (Phase 1 complete — 200 tests passing)

- **IR types** (`crates/office2pdf/src/ir/`): Document, Page (Flow/Fixed/Table), Block (Paragraph/Table/Image/PageBreak), Run, Table, TableCell (with borders/background/merging), ImageData, Shape, Style types
- **DOCX parser** (`crates/office2pdf/src/parser/docx.rs`): Text, inline formatting (bold/italic/underline/strikethrough/font/size/color), paragraph formatting (alignment/indent/spacing/page breaks), tables (cell merging via gridSpan/vMerge, borders, shading), images. Uses `docx-rs` + `serde_json` for private field access.
- **PPTX parser** (`crates/office2pdf/src/parser/pptx.rs`): Direct ZIP+quick-xml parsing (not ppt-rs). Slides, text boxes with formatting, shapes (rect/ellipse/line with fill/stroke), images via slide .rels resolution. Uses `SolidFillCtx` enum for context-aware color parsing.
- **XLSX parser** (`crates/office2pdf/src/parser/xlsx.rs`): Cell data (text/number/date as strings), multi-sheet, column widths, cell merging. Uses `umya-spreadsheet`.
- **Typst codegen** (`crates/office2pdf/src/render/typst_gen.rs`): FlowPage (paragraphs, tables, images), FixedPage (#place for absolute positioning), TablePage. Image asset tracking via `TypstOutput`/`ImageAsset`/`GenCtx`.
- **PDF renderer** (`crates/office2pdf/src/render/pdf.rs`): `MinimalWorld` implementing Typst `World` trait. Embedded fonts only (Libertinus, New Computer Modern, DejaVu). Image serving via virtual paths.
- **Config** (`crates/office2pdf/src/config.rs`): Format enum (DOCX/PPTX/XLSX), ConvertOptions struct (placeholder)
- **Error** (`crates/office2pdf/src/error.rs`): ConvertError enum (UnsupportedFormat, Io, Parse, Render)
- **CLI** (`crates/office2pdf-cli/src/main.rs`): `office2pdf <input> [--output <path>]` wired to library convert()
- **Dependencies**: typst 0.14, typst-pdf 0.14, typst-kit 0.14, comemo 0.5, docx-rs 0.4, serde/serde_json, zip 0.6, quick-xml 0.38, umya-spreadsheet 2

## Test Fixtures

For integration tests requiring actual document files, create minimal test documents programmatically where possible. For format-specific parsing tests, place small test fixture files in `tests/fixtures/`.

## Progress Report Format

APPEND to `scripts/ralph/progress.txt` (never replace, always append):
```
## [Date/Time] - [Story ID]
- What was implemented
- Files changed
- Dependencies added (if any)
- **Learnings for future iterations:**
  - Patterns discovered
  - Gotchas encountered
  - Useful context
---
```

## Consolidate Patterns

If you discover a **reusable pattern**, add it to the `## Codebase Patterns` section at the TOP of `scripts/ralph/progress.txt` (create it if it doesn't exist). Only add patterns that are general and reusable.

## Finalization (after all stories complete)

When ALL user stories have `passes: true`, you MUST push, create a PR, and verify CI before finishing:

1. **Push**: `git push -u origin <branchName>` (branchName from PRD)
2. **Check for existing PR**: `gh pr list --head <branchName> --json number --jq '.[0].number'`
3. **Create PR** if none exists:
   - Generate a PR body summarizing all completed stories, commit list, and test plan
   - `gh pr create --title "feat: <phase description>" --body "<body>" --base main`
4. **Wait for CI to register** (30 seconds): `sleep 30`
5. **Watch CI**: `gh pr checks <number> --watch` (blocks until all checks finish; typically 3-6 minutes)
6. **If CI fails**:
   a. Get the failed run ID: `gh run list --branch <branchName> --status failure --json databaseId --jq '.[0].databaseId'`
   b. Read failure log: `gh run view <run-id> --log-failed 2>&1 | head -200`
   c. Identify and fix the errors in source code
   d. Run local quality checks again (fmt, clippy, test)
   e. Commit: `git commit -s -m "fix: resolve CI failures"`
   f. Push: `git push`
   g. Go back to step 5 (retry up to 3 times)
7. **Only after CI passes**, respond with `<promise>COMPLETE</promise>`

If CI still fails after 3 retries, respond with `<promise>COMPLETE</promise>` anyway (the PR will be reviewed manually).

## Stop Condition

After completing a user story, check if ALL stories have `passes: true`.

If ALL stories are complete and passing, proceed to **Finalization** above.

If there are still stories with `passes: false`, end your response normally (another iteration will pick up the next story).

## Important

- Work on ONE story per iteration
- Write tests FIRST (TDD)
- Commit frequently
- Keep CI green (all quality checks must pass)
- Read the Codebase Patterns section in progress.txt before starting
- Do NOT modify this file (CLAUDE.md) during execution
