# Issue 1879 visual evidence

Source: `tests/fixtures/pptx/1-slide.pptx`. Native reference exported through Microsoft PowerPoint 16.113.4 on macOS 26.6.2 (25G83), using the local Best for printing PDF option on 2026-10-11.

SHA-256:

- Source: `181622a0431e7be842b56f1783643fe373ea8936629e033d3879f492e5a5349a`
- Native GT PDF: `a5ba138cb826237928f823ffe381e17e7bb497c2527207b1d41333292edf1ef1`
- Current output PDF: `d58a35a2b133257723048ab4e1db650558ab03c21ef454bb784e048aeac0eadf`

Before uses main `71e103115ee19b1c72cd765ff04a27b5b8dede6f`; after uses this PR's `3338dce9922bfc376c01c11c3f1ea284ca2a4cff` code. Both converters were built with `cargo build --locked --profile ci -p office2pdf-cli`. Original PDFs, binaries, logs and inspected PNGs are retained outside the worktree.

Render the source with the native app's local PDF export and with `cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/pptx/1-slide.pptx -o after.pdf`. Compare all pages with `compare_layout.py --json --audit --fine-shift 1`, `compare_text_layer.py`, and `compare_render.py --page 1 --dpi 150 --fine-shift 1 --strict-clusters` using the exact IDs in the committed report.

One page was inspected at 150 DPI. JPEGs preserve pixel dimensions, use progressive quality 86, and strip metadata before resetting 150 DPI density. The 1pt fine threshold is below the couple-of-points external-reference placement materiality scale; smaller observed differences remain documented. Layout and normalized searchable content pass.

The previously purple panel now takes the white slide background, as p:sp/@useBgFill="1" specifies. PAGE 1, the right-side picture, both black rules, their flat endpoints, and the picture crop agree with native PowerPoint. The rules are solid 3.5pt and 1pt strokes (a:ln/@w=44450 and 12700); no dash, emphasis, rotation, clipping, or missing element was observed. Full pages, the pixel diff, and matched title/rule/picture crops show only boundary coverage and photo resampling fragments below the material-cluster floor. There are zero current material clusters.
