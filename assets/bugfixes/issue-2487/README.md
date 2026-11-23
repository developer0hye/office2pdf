# Brace-pair geometry audit

Public fixture: `tests/fixtures/pptx/brace_pair.pptx`, from [the original report](https://github.com/developer0hye/office2pdf/issues/2487#issuecomment-7058970574). Synthetic metadata and all package parts were checked; no notes, comments, external relationships, custom XML or embedded objects are present.

Fresh GT: Microsoft PowerPoint 16.115.1 on macOS 27.2.1 (26B131), exported from the unmodified submitted PPTX in PowerPoint's internal-disk container. `pdfinfo` and `check_gt_integrity.py` passed: one 960 x 540 pt page. Before is the validated fill-only tree 7c0fabe4; after adds the brace geometry.

- Source SHA-256: 61122cf5dc0548c8e73fb0fff264d7e4b214b90cfe9f4c3461c2e73817f84f62
- Native PDF SHA-256: 0744fa51e0137076bd060f86c1b54eff4d7235580d8fa848e8f52bcaf32d8c94
- Before PDF SHA-256: 3b027b651e4c808ea4d12083beaa79041f1b643b430925e7a9e8b6c2c41c121e
- After PDF SHA-256: 720015a699e623f152dfffae895b08b5ac7a8dd778e78f318fd1211c60a3754d

Commands: `compare_layout.py --json --audit --fine-shift 0.5`, `compare_text_layer.py`, `compare_render.py --page 1 --dpi 150 --fine-shift 0.5 --strict-clusters` with current exact-ID dispositions. All layout categories pass; normalized text and codepoint census match. Evidence JPEGs are progressive quality 86 at 150 DPI with metadata removed.

Model vision: reviewed the full GT/before/after pages, current diff, both full-height brace crops, title and rectangle-control crops. The two magenta curves match native tips, central cusps, straight segments and open endpoints; no top or bottom joining stroke remains. Interior stays white; filled control and all text are preserved. Sparse arc-edge differences are below the material cluster floor; current title glyph-edge and approximately 0.12pt baseline differences remain under #1874. No new material defect was found. This fixture has 2pt solid strokes; no hairlines, dashes, bold, italic, underline or CJK.

## Final memory and speed pass

After visual acceptance: Apple M2, 16 GiB, macOS 27.2.1; ci profile; system fonts with ArialMT embedded; warm filesystem/font cache, fresh process per sample; one untimed warm-up and five timed samples per input/build. `/usr/bin/time -l BINARY INPUT -o OUTPUT.pdf` provides peak resident memory; Python perf_counter measures elapsed process time. No local compilation/tests ran during sampling.

| Input | Before geometry median seconds / MiB | Accepted pre-optimization median seconds / MiB |
| --- | --- | --- |
| brace_pair.pptx | 0.4331 / 69.31 | 0.4414 / 69.62 |
| 1-slide.pptx control | 1.0099 / 108.52 | 0.9973 / 108.34 |

No worthwhile simple optimization was supported: the new work is a bounded set of arc samples and normalized points; the measured differences are small and control timing moves in the opposite direction. The accepted build is the pre-optimization build and remains unchanged, with no optimization code added. Required full workspace tests follow this pass. These scoped measurements are not product SLAs.
