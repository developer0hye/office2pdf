# Issue 1884 visual evidence

Four synthetic one-page probes are tracked as `tests/fixtures/docx/issue-1884-{normal-absent,normal-partial,custom-absent,custom-partial}.docx`. Each was exported locally with Microsoft Word 16.113.4, Best for printing, on macOS 26.6.2 (25G83), 2026-10-11. The order below is used in the combined reports and page images.

Before uses main `71e103115ee19b1c72cd765ff04a27b5b8dede6f`; after uses this PR integrated with that main. Both binaries were built with `cargo build --locked --profile ci -p office2pdf-cli`. PDFs, binaries, logs and inspected full pages/diffs/crops are retained outside the worktree. Reproduce each output with `cargo run --locked --profile ci -p office2pdf-cli -- FIXTURE -o after.pdf` and export the same fixture with Word.

The JPEGs vertically stack all four complete 1275 x 1650px pages in the recorded order without resizing. They use progressive quality 86, stripped metadata, and reset 150 DPI density. Combined PDFs preserve the same order with `pdfunite`. `compare_layout.py --json --audit --fine-shift 1` passes all four pages; normalized text matches separately for every probe. `compare_render.py --page N --dpi 150 --fine-shift 1 --strict-clusters` passes for N=1..4 with zero material clusters.

Full pages, diffs and matched content crops show both cells and the following paragraph in the native positions. The custom partial style's explicit top inset remains visible; absent sides no longer introduce Typst's unrelated 5pt inset. No missing text, wrap, clipping, overlap, rotation, fill, rule, or emphasis discrepancy was observed. These probes contain no painted borders, hairlines, pictures, bold, italic, or underlined runs. Sub-material cell-text anchor offsets are at most 0.22pt horizontally and 0.167pt vertically, tracked in #1874; they are not claimed to be rasterization.

| Page / probe | Source SHA-256 | Native PDF SHA-256 | Current PDF SHA-256 |
| --- | --- | --- | --- |
| 1: normal-absent | `71772c46317aaf46177aadd4c188f2c7a297aabb2ddb97b25170fcf661c1fce9` | `75123458dfcba99245f45ac75b7c30b730075f2fba3f54dc229b1f32fda14b35` | `7560d71b66f1c3d277ecefd8392a9b6325fb9f2726d10a3af5c4859ccfb06d21` |
| 2: normal-partial | `67e92d9d1b55fefb9da30e3e0fb9a8c9c57326fd90d963e0a95ec5d047f3805a` | `3d837255a9a4ce4c679c4c7c4827bd81cd2420285b671011f210ba036f789475` | `7560d71b66f1c3d277ecefd8392a9b6325fb9f2726d10a3af5c4859ccfb06d21` |
| 3: custom-absent | `96e953f461e7f3ce2b17ace4c0b4a9af8b46c1507ecb3a1fea50a15d607cc89b` | `109654b22410641d6ddeef7839b26c31e66e67785eab5e71520aa0e88c669b1e` | `4b2fa1ceff77fce8be459d5ba7bfee33fadbecf4f91a311bd01ed454ec300279` |
| 4: custom-partial | `13fdc14be21f69e640d02010e4f93629cccdf2765961a400ce54efaf859985f2` | `4dea12c3b187ebe55930757b6c0ec0baa35137ab813241e203369f4da04e9d72` | `0e08e134d8d1ca865ca0d9e84acf2c2a7d3f4faf6962b7d36423728dac82c286` |

Combined GT PDF: `0affa108a7ea3eed1d004f85b350e369f20a2f92166674dcea56547ad66f745d`. Combined current PDF: `a0d99a9b0bf81c02fb38dac7776cc14c19d209ea5d98b20f12076cc49c83c079`.
