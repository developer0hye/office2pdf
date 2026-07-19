# XLSX Six-Digit Style Colors Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve six-digit OOXML `RRGGBB` colors in XLSX fills and fonts so issue #328 matches the native Excel export.

**Architecture:** Normalize the shared OOXML color parser to accept exactly six-digit RGB and eight-digit ARGB inputs while continuing to reject malformed lengths. Verify the parser directly, verify the contributor fixture through the public XLSX parser, then regenerate the same page at 150 DPI for the visual contract.

**Tech Stack:** Rust, `umya-spreadsheet`, office2pdf IR, Typst, Cargo tests, Poppler, ImageMagick

---

## Chunk 1: Parser and fixture behavior

### Task 1: Accept RGB and ARGB color encodings

**Files:**
- Modify: `crates/office2pdf/src/parser/xml_util.rs`
- Test: `crates/office2pdf/src/parser/xml_util_tests.rs`

- [ ] **Step 1: Write failing parser tests**

Add tests proving `D9EAF7` becomes `(0xD9, 0xEA, 0xF7)`, `16324F` becomes `(0x16, 0x32, 0x4F)`, and lengths other than six or eight remain rejected.

- [ ] **Step 2: Run the focused parser tests and verify RED**

Run: `cargo test --offline -p office2pdf parser::xml_util::tests::parse_argb_color -- --nocapture`

Expected: six-digit cases fail because `parse_argb_color` currently rejects values shorter than eight characters.

- [ ] **Step 3: Implement exact-length RGB/ARGB parsing**

For six characters, parse all three byte pairs. For eight characters, ignore the alpha byte and parse the remaining byte pairs. Return `None` for every other length or invalid hexadecimal input.

- [ ] **Step 4: Re-run the focused parser tests and verify GREEN**

Run: `cargo test --offline -p office2pdf parser::xml_util::tests::parse_argb_color -- --nocapture`

Expected: all focused color tests pass.

### Task 2: Lock the real fixture behavior

**Files:**
- Test: `crates/office2pdf/tests/xlsx_fixtures.rs`

- [ ] **Step 1: Add a contributor-fixture acceptance test**

Parse `pr_186_contributor_acceptance.xlsx` and assert the header cell fill is `D9EAF7` and its header run text color is `16324F`.

- [ ] **Step 2: Run the acceptance test**

Run: `cargo test --offline -p office2pdf --test xlsx_fixtures acceptance_pr_186_contributor_acceptance_six_digit_colors -- --nocapture`

Expected: PASS after the shared parser fix and fail on the parent `main` baseline.

- [ ] **Step 3: Run the XLSX and full workspace regressions**

Run: `cargo test --offline -p office2pdf --test xlsx_fixtures`

Run: `cargo test --offline --workspace`

Expected: all tests pass.

## Chunk 2: Visual contract and delivery

### Task 3: Generate reproducible issue evidence

**Files:**
- Create: `assets/bugfixes/issue-328/gt.jpg`
- Create: `assets/bugfixes/issue-328/before.jpg`
- Create: `assets/bugfixes/issue-328/after.jpg`

- [ ] **Step 1: Render the same fixture/page at 150 DPI**

Use the existing native Excel PDF for GT, the captured parent-`main` output for before, and this branch output for after. Render page 1 with `pdftoppm -r 150`.

- [ ] **Step 2: Re-audit the after image**

Check page count/order, presence, position, size, rotation, fill, stroke, text, font, color, alignment, spacing, clipping, matched crops, 5% pixel diff clusters, hairlines, and emphasis. Record every unrelated remaining deviation as an existing open issue reference.

- [ ] **Step 3: Encode and validate images**

Store all three as progressive JPEG quality 86, stripped metadata, 150 DPI, original rendered dimensions. Run `python3 -m unittest scripts.tests.test_check_visual_pr` and the visual PR contract validator.

### Task 4: Commit, publish, and merge

**Files:**
- Review all files changed above

- [ ] **Step 1: Format, lint, and review**

Run: `cargo fmt --all -- --check`

Run: `cargo clippy --offline --workspace --all-targets -- -D warnings`

Expected: both pass and `git diff --check` is clean.

- [ ] **Step 2: Commit with DCO**

Commit with `git commit -s` using the configured `Yonghye Kwon <developer.0hye@gmail.com>` identity.

- [ ] **Step 3: Push and open the PR as developer0hye**

Immediately verify `gh auth status` and `gh api user --jq .login`, push the branch, and create an English PR with `Related: #328`, GT/before/after links, and remaining-issue references.

- [ ] **Step 4: Wait for CI, review the final diff, and merge**

Require all checks green, review `gh pr diff`, merge with `gh pr merge --merge`, sync `main`, and verify post-merge CI before cleanup.
