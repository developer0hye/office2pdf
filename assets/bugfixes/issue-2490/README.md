# Fill-reference audit

The public synthetic input is from [#2487](https://github.com/developer0hye/office2pdf/issues/2487#issuecomment-7058970574). Metadata contains only Synthetic reproduction; no notes, comments, external relationships, custom XML or embedded objects were found.

Fresh reference: Microsoft PowerPoint 16.115.1 on macOS 27.2.1 (26B131), exported from the original input in PowerPoint's internal-disk container. `pdfinfo` reports one 960 x 540 pt page; `check_gt_integrity.py` passed. Baseline: e02eee43. Candidate adds only the no-fill index fix and the tested URL-advisory prerequisite.

- Input SHA-256: 61122cf5dc0548c8e73fb0fff264d7e4b214b90cfe9f4c3461c2e73817f84f62
- Native PDF SHA-256: 0744fa51e0137076bd060f86c1b54eff4d7235580d8fa848e8f52bcaf32d8c94
- Before PDF SHA-256: 4c3616c5f501d8ed9c45a5f49674ad951d854defb032393c777cfe7baa21b970
- After PDF SHA-256: 3b027b651e4c808ea4d12083beaa79041f1b643b430925e7a9e8b6c2c41c121e

Compared the complete page with `compare_layout.py --json --audit --fine-shift 0.5`, `compare_text_layer.py`, and `compare_render.py --page 1 --dpi 150 --fine-shift 0.5 --strict-clusters`. Layout categories pass and normalized text is identical. PDF font: embedded ArialMT.

Model vision findings: inspected full GT/before/after pages, current pixel diff, the title, a full-height brace-side crop, and the rectangle-control corner at 150 DPI. The unexpected teal interior disappears. Magenta control, labels, positions, page count, content and wrapping are preserved. The independent missing brace outline remains in #2487 (current exact clusters named in the strict report). Title edge/baseline differences around 0.12pt are sub-material and tracked in #1874. No hairlines, dashes, bold, italic, underline or CJK appear in this fixture. Solid 2pt strokes were inspected.

## Final memory and speed pass

Measured after visual acceptance on Apple M2, 16 GiB RAM, macOS 27.2.1, system fonts (ArialMT embedded), `ci` profile, warm filesystem/font cache, fresh process per sample; no concurrent local build/tests during sampling. One untimed warm-up and five repetitions per input/build. Command: `/usr/bin/time -l BINARY INPUT -o OUTPUT.pdf`; elapsed time includes process launch and was measured with Python perf_counter; peak memory is the process maximum resident set size.

| Input | Before fix median seconds / MiB | Accepted pre-optimization median seconds / MiB |
| --- | --- | --- |
| brace_pair.pptx | 0.3981 / 68.67 | 0.4017 / 68.88 |
| 1-slide.pptx control | 0.9159 / 107.88 | 0.9307 / 108.00 |

The accepted build is the pre-optimization build. No worthwhile low-complexity optimization is supported by these measurements: both changes are small and within run-to-run variability, and this rule adds a constant-time index check with no allocation. No optimization code was introduced; the visually accepted build is unchanged. These scoped measurements do not establish product SLAs. Final full workspace tests run after this pass.
