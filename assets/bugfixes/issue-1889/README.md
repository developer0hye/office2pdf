# Issue 1889 visual evidence

Source: `tests/fixtures/docx/issue-1889-inline-table.docx`. Native reference exported through Microsoft Word 16.113.4 on macOS 26.6.2 (25G83), using the local Best for printing PDF option on 2026-10-11.

SHA-256:

- Source: `42114ae5d9a08156f57b5e6b828f5a630e4d98203a8454dae806630cc0751050`
- Native GT PDF: `2ba48b423170cbbf04c183201e65b75951479d6d9c8cbc1c17bf8f9f2725e2a8`
- Current output PDF: `776178c223db0a2f519675af853a7c6d8666662f9a4dce93bfb13592609be6d5`

Before uses main `71e103115ee19b1c72cd765ff04a27b5b8dede6f`; after uses this PR's `e623b9de1acce599b99efd18efbb508cef2c38c9` code. Both converters were built with `cargo build --locked --profile ci -p office2pdf-cli`. Original PDFs, binaries, logs and inspected PNGs are retained outside the worktree.

Render the source with the native app's local PDF export and with `cargo run --locked --profile ci -p office2pdf-cli -- tests/fixtures/docx/issue-1889-inline-table.docx -o after.pdf`. Compare all pages with `compare_layout.py --json --audit --fine-shift 1`, `compare_text_layer.py`, and `compare_render.py --page 1 --dpi 150 --fine-shift 1 --strict-clusters` using the exact IDs in the committed report.

One page was inspected at 150 DPI. JPEGs preserve pixel dimensions, use progressive quality 86, and strip metadata before resetting 150 DPI density. The 1pt fine threshold is below the couple-of-points external-reference placement materiality scale; smaller observed differences remain documented. Layout and normalized searchable content pass.

The 216 x 72pt white box with a solid 0.75pt black outline now contains its two-column table. The anchor remains alongside the box, the following paragraph stays below it, and all six text fragments are visible and searchable. No rotation, clipping, overlap, missing row, bold/italic/underline, or color mismatch was observed. The single current cluster is assigned to #1874: table-text offsets reach 0.31325pt vertically and 0.26925pt horizontally. Those sub-material offsets are geometry, not a rasterization exemption.
