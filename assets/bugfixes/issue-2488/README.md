# Alternate-content audit

Public synthetic fixture: `tests/fixtures/pptx/office_math_alternate_content.pptx`, from [the original report](https://github.com/developer0hye/office2pdf/issues/2488#issuecomment-7058982614). Creator/last modifier are Synthetic reproduction. All package parts were inspected: no notes, comments, custom XML, embedded objects or external relationships.

Fresh GT: Microsoft PowerPoint 16.115.1 on macOS 27.2.1 (26B131), exported from the unmodified input in PowerPoint's own internal-disk container. `pdfinfo` and `check_gt_integrity.py` passed: two 960 x 540 pt pages. Before is e02eee43 (0.11.0); after changes branch selection. Evidence stacks both full pages vertically at 150 DPI; `compare.jpg` stacks native/before comparisons, and `math-control-compare.jpg` isolates native/after on page 2. JPEGs are progressive quality 86 with metadata stripped.

- Source SHA-256: ae9baa8b22a432fb859830c394cdfbbbe89e50d19c9cb4e6305a43f37d3436e2
- Native PDF SHA-256: e85c669f9a55a0802730741c1edcebc82d28b452d5b16cbaf4e8c76bd6562db7
- Before PDF SHA-256: f424c77a4fd9002f989f9b80e9398df533aa5b10021935dea98e3ec3eea205f8
- After PDF SHA-256: eb244925aa229f74adcefd089e628fd4faed7dd7b613501f946ae5014bac5af5

Commands: `compare_layout.py GT AFTER --json --audit --fine-shift 0.5`, `compare_text_layer.py GT AFTER`, then `compare_render.py GT AFTER --page N --dpi 150 --fine-shift 0.5 --strict-clusters` for each page with current exact-ID dispositions. Layout/text reports retain the native-equation findings under #2493. All other layout categories pass. The final grouped-shape correction produces a byte-identical reported-file PDF and was re-audited using the same commands.

Model vision: inspected every full GT/before/after page, diff, and matched title/caption/equation crop at 150 DPI. The overlapping bitmap disappears on page 1; both equation crops are now pixel-identical. The full page-2 output remains pixel-identical to before (`magick compare -metric AE`: 0). `pdfimages -list` reports zero images. Black color is preserved (histogram color shift 0.0000). Native math still has the known 6.88pt baseline shift, compact operator spacing, and slanted plus/digit on both pages (#2493); these are material, not antialiasing. Current title/caption edge differences and ~0.12pt title baseline differences remain sub-material under #1874. No CJK, bold, underline, hairlines or dashes exist; italic/upright equation runs were inspected directly.

## Final memory and speed pass

After final visual acceptance: Apple M2, 16 GiB RAM, macOS 27.2.1, system fonts including ArialMT and Cambria Math, `ci` profile; warm filesystem/font cache and fresh process per sample. One untimed warm-up plus five measurements per build/input, without concurrent local builds/tests. Command: `/usr/bin/time -l BINARY INPUT -o OUTPUT.pdf`; Python perf_counter measures elapsed process time and time's maximum resident set size supplies peak memory.

| Input | Before median seconds / MiB | Accepted pre-optimization median seconds / MiB |
| --- | --- | --- |
| Reported two-slide presentation | 0.6838 / 84.88 | 0.6818 / 84.73 |
| 1-slide.pptx control | 0.9252 / 107.91 | 0.9277 / 107.97 |

No worthwhile simple optimization was supported: timing and memory differences are small. The selector borrows XML when no branches are removed and skips scanning parts without AlternateContent. No optimization code was added after acceptance; the accepted pre-optimization build remains unchanged. Full workspace tests follow this pass. Scoped measurements do not establish product SLAs.
