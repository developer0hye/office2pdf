# Issue 2041 visual evidence

Source: `tests/fixtures/pptx/oxp_PB001-Input1.pptx`, both slides. Microsoft PowerPoint 16.113.4 exported the native PDF locally with Best for printing on macOS 26.6.2 (25G83), 2026-10-11.

SHA-256:

- Source: `14b54006862844a1e7c8c4e9e8170df17b7fe69ebd8cfbbfcb6dcd4e16388ffc`
- Native GT PDF: `6aedd3bd31b2c8a34d2746279e9ad75a391afb10c71b964b8e47a3393c1a342d`
- Current PDF: `6b9a2f78a0aeba4f6c790e0ac203bb2554800d42a8158266565a47bfd4b7a2ff`

Before uses main `71e103115ee19b1c72cd765ff04a27b5b8dede6f`; after uses `d06dc8da`. Both were built with `cargo build --locked --profile ci -p office2pdf-cli`. PDFs, binaries, logs and inspected crops are retained outside the worktree.

The JPEGs vertically stack complete pages 1 and 2 without rescaling. Each page retains its 1625 x 1125 pixels from pdftoppm at 150 DPI. Images are progressive quality 86, stripped before resetting 150 DPI. Both pages have fresh layout, text-layer and strict render-cluster audits at a 1pt fine threshold. There are no missing/extra/reflowed lines, visibility failures, shifts over the threshold, or material clusters.

Reproduce with `cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/pptx/oxp_PB001-Input1.pptx -o after.pdf`; export the same file with native PowerPoint. Run `compare_layout.py --json --audit --fine-shift 1`, `compare_text_layer.py`, and `compare_render.py --page N --dpi 150 --fine-shift 1 --strict-clusters` for N=1,2 with the report's renderer observations.

Full pages, diffs and matched title/header/photo/logo/table/footer crops were inspected. Page 2's four affected entries now align with the native table. Their a:pPr/@indent=-228600 is an -18pt first-line offset, separate from marL=771525. The cover, all table labels and page numbers remain intact; bold headings, blue/gray text and dotted leaders agree. The 0.5pt solid rules/branches and 0.5/0.5pt dotted leaders retain their trace widths and positions. Native strokes have negligible endpoint slope, giving different fractional pixel coverage. Photo resampling and glyph-edge differences stay below the material-cluster floor. No new material deviation was observed.
