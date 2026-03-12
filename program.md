# office2pdf autoresearch

This repository can be run as an autoresearch loop.

The mission is simple: make the PDFs generated from `project/原文档/` look as close as possible to the reference PDFs in `project/standard/`.

This is not a semantics-only conversion project. The success criterion is visual fidelity against the standard PDFs, with strong preference for generic fixes that improve multiple files at once.

## Goal

- Minimize dataset loss reported by `python3 project/compare_pdfs.py`.
- Reduce worst-case file loss and worst-page pixel loss.
- Match standard page counts whenever possible.
- Improve word matching, word positions, and text similarity without sacrificing layout.
- Prefer fixes that generalize across the corpus over document-specific hacks.

## Non-goals

- Do not optimize for a single document by hardcoding filename checks, literal text checks, or one-off geometry exceptions.
- Do not edit `project/standard/` to make the metrics look better.
- Do not accept a small metric gain if it introduces obvious overfitting or fragile complexity.

## Read First

Before starting an experiment loop, read these files:

1. `README.md`
2. `PRD.md`
3. `METHODOLOGY.md`
4. `project/README.md`
5. `project/compare_pdfs.py`
6. The latest `project/report/report.md` if it exists

Then read only the code relevant to the current hypothesis. In most cases that means one or more of:

- `crates/office2pdf/src/parser/docx*.rs`
- `crates/office2pdf/src/parser/pptx*.rs`
- `crates/office2pdf/src/render/typst_gen*.rs`
- `crates/office2pdf/src/render/font_*.rs`
- the corresponding tests

Use `PRD.md` section `5.5 High-Fidelity OOXML -> Typst Conversion Strategy` as the rendering playbook.

## Setup

1. Create a dedicated worktree for the run.
2. Create a dedicated branch in that worktree. Follow the repo branch rules. A good name is `feat/autoresearch-<tag>`.
3. Verify required tools exist:
   - `cargo`
   - `python3`
   - `pdfinfo`
   - `pdftotext`
   - `pdftoppm`
4. Initialize `project/results.tsv` if it does not exist.
5. Run a baseline conversion and comparison before changing code.

Use this header for `project/results.tsv`:

```tsv
commit	dataset_score	dataset_loss	worst_case_loss	page_mismatch_cases	status	description
```

## Baseline Run

Regenerate `project/target/` from the current code before the first comparison:

```bash
python3 - <<'PY'
from pathlib import Path
import subprocess

root = Path("project/原文档")
target_root = Path("project/target")

for format_name in ("docx", "pptx"):
    for src in sorted((root / format_name).rglob(f"*.{format_name}")):
        rel = src.relative_to(root / format_name).with_suffix(".pdf")
        out = target_root / format_name / rel
        out.parent.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            ["cargo", "run", "-p", "office2pdf-cli", "--", str(src), "-o", str(out)],
            check=True,
        )
PY
```

Then compare:

```bash
python3 project/compare_pdfs.py
```

The main outputs are:

- `project/report/report.json`
- `project/report/report.md`
- `project/report/artifacts/<format>/<relative-stem>/`

Record the baseline result in `project/results.tsv`.

## Metrics That Matter

Primary metric:

- `dataset_loss` from `project/report/report.md` or `project/report/report.json`

Hard gates:

- No new conversion failures
- No intentional degradation of page-count parity
- No document-specific hacks

Secondary metrics:

- worst-case file loss
- pixel mean loss
- word F1
- word bbox score
- text similarity
- page count match rate

When the metrics disagree, prefer the change that produces the better visual result in the rendered artifacts.

## Where to Look

When a comparison run finishes, inspect the worst cases first.

For each bad case, check:

- page count mismatch
- `standard.txt` vs `target.txt`
- `renders/standard-*.ppm`
- `renders/target-*.ppm`
- `renders/diff-*.ppm`

Typical root-cause buckets:

- font resolution or wrong fallback font
- paragraph spacing or line-height mismatch
- list indentation or numbering layout
- page-break or section-break handling
- table width, row height, border, or cell padding mismatch
- DOCX style inheritance not fully flattened
- PPTX master/layout/theme inheritance not fully resolved
- PPTX absolute positioning or sizing drift
- shape fallback quality problems

Fix root causes, not symptoms.

## Experiment Rules

Every experiment should be a single clear hypothesis.

Good hypotheses:

- flattening DOCX paragraph spacing before Typst emission will reduce page drift in long contracts
- fixing Office font discovery will improve both page counts and word positions
- resolving PPTX placeholder inheritance earlier will improve multiple decks at once

Bad hypotheses:

- special-case one file
- tweak output only for one literal heading
- add a hack that only exists to move one page break in one document

## Required Workflow

The loop is:

1. Start from the current best commit in the dedicated worktree.
2. Read the latest comparison report and choose one high-leverage hypothesis.
3. Add or update a regression test before changing runtime code.
4. Implement the smallest generic fix that can prove or disprove the hypothesis.
5. Run targeted tests first.
6. If the change still looks promising, regenerate `project/target/`.
7. Run `python3 project/compare_pdfs.py`.
8. Inspect the worst cases and confirm whether the change helped visually.
9. Log the result in `project/results.tsv`.
10. Keep the commit only if the overall result is better.
11. Continue immediately with the next hypothesis.

The loop does not stop on its own. Keep going until the human interrupts.

## Keep or Discard

Keep a change when most of the following are true:

- `dataset_loss` decreases
- the worst case improves or at least does not regress badly
- page-count mismatches do not get worse
- there is no new conversion failure
- the implementation remains general and defensible

Discard a change when any of the following are true:

- `dataset_loss` is worse
- the improvement is limited to one document because of a special-case hack
- one metric improves but the rendered artifacts clearly look worse
- the complexity cost is too high for the gain

If a run crashes, log it as `crash`, record the attempted idea, fix obvious mistakes if they are trivial, and move on.

## Testing Strategy

Follow TDD.

- Add a failing regression test for the target behavior.
- Prefer fixture-driven tests that reflect real document structures.
- Validate the smallest affected surface first.
- Run broader tests before accepting a keep decision.

Useful checks include:

- targeted `cargo test` for the parser or renderer area you changed
- full `cargo test` when the change affects shared behavior
- `python3 project/compare_pdfs.py` for corpus-level validation

If a code change affects runtime behavior, do not skip tests.

## What “Better” Looks Like

A strong improvement usually has this shape:

- the same fix improves more than one file
- text similarity stays stable or rises
- word placement improves
- page counts move closer to the standard PDFs
- visible layout drift decreases in the artifact renders

The ideal end state is that a human can flip between `project/standard/` and `project/target/` PDFs and see only minor differences.

## Result Logging

Each experiment must append one row to `project/results.tsv`.

Example:

```tsv
commit	dataset_score	dataset_loss	worst_case_loss	page_mismatch_cases	status	description
1a2b3c4	0.620369	0.379631	0.547229	4	keep	baseline
2b3c4d5	0.641000	0.359000	0.498000	3	keep	resolve paragraph spacing before Typst generation
3c4d5e6	0.612000	0.388000	0.560000	5	discard	try tighter fallback font substitution
4d5e6f7	0.000000	1.000000	1.000000	0	crash	break list layout while refactoring numbering
```

Be brief and factual in the description.

## Final Principle

The project wins by repeatedly closing the gap between generated PDFs and the standard PDFs.

That means:

- measure
- inspect
- hypothesize
- test
- fix
- compare
- keep only real gains
- repeat

Never confuse “Typst code looks cleaner” with “the PDF is better”. The PDF is the product.
