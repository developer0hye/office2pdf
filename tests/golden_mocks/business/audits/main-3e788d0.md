# Visual baseline: office2pdf main at `3e788d0`

- Date: 2026-07-22
- Converter: `office2pdf 0.6.3` dev build from `3e788d03f7d9408b2f6da5433df271a4701c0166`
- Reference: native Microsoft Office 16.111.1 PDFs from `expected/`
- Render: Poppler at 150 DPI; converted pages size-normalized to the native render before ImageMagick `AE` with 5% fuzz
- Result: all 30 files converted; all 54 pages preserved in source order
- Workbook check: artifact-tool imported all 10 XLSX files, found no displayed formula-error tokens, and rendered every one of the 11 worksheets for a separate visual pass

## Human review summary

All 54 expected/converted page pairs were reviewed in format contact sheets. The largest outlier in each format was also inspected at full rendered resolution. Content and page order are preserved, and no unexplained clipping, overflow, rotation, flipping, or missing page was observed.

- **DOCX:** text, images, fills, borders, bullets, explicit page breaks, and emphasis remain present. Native and converted A4 renders differ by one horizontal pixel at 150 DPI. Table row heights, column widths, paragraph spacing, and text metrics remain visibly different. `docx-research-report-ko` is the material pagination-within-page outlier: native page 1 ends at 2025-11 while the converter fits through 2026-05; all rows remain present across two pages.
- **PPTX:** all 26 pages are an exact 2000×1125 size match. Shapes, images, fills, shadows, tables, bullets, and slide order are present. Remaining deviations are dominated by text rasterization/spacing and effect rendering; `pptx-conference-talk-en` page 2 is the largest outlier because monospace code glyph metrics differ, without clipping.
- **XLSX:** formulas resolve to the expected displayed values; merged cells, print order, conditional fills, data bars, color scales, icon sets, and table borders are present. Native pages render as 1240×1755 and converted pages as 1241×1754. `xlsx-inventory-en` has the largest table-scale, row-height, and data-bar geometry differences, while preserving all four pages and row order.
- **Hairlines and emphasis:** table/card hairlines visible in the native references remain present in the converter output, though DOCX/XLSX stroke weight and antialiasing differ. The source-level emphasis inventory below was checked against the paired renders; no missing bold, italic, or underline treatment was observed at contact-sheet scale.

These are fidelity baselines, not reasons to weaken a fixture. Historical issue context remains attached to each case in `manifest.json`; every page row below links to the first related regression issue or to the audit PR when no case-specific issue exists.

## Case-level metrics

| Case | Pages | Exact rendered size | Median 5% fuzz diff | Max 5% fuzz diff |
| --- | ---: | ---: | ---: | ---: |
| `docx-invoice-en` | 1 | 0/1 | 0.0659 | 0.0659 |
| `docx-contract-ko` | 1 | 0/1 | 0.0914 | 0.0914 |
| `docx-meeting-minutes-ko` | 1 | 0/1 | 0.0868 | 0.0868 |
| `docx-resume-en` | 1 | 0/1 | 0.0517 | 0.0517 |
| `docx-technical-manual-en` | 2 | 0/2 | 0.0465 | 0.0480 |
| `docx-official-letter-ko` | 1 | 0/1 | 0.0411 | 0.0411 |
| `docx-product-spec-en` | 1 | 0/1 | 0.0926 | 0.0926 |
| `docx-newsletter-en` | 1 | 0/1 | 0.0709 | 0.0709 |
| `docx-offer-letter-en` | 1 | 0/1 | 0.0387 | 0.0387 |
| `docx-research-report-ko` | 2 | 0/2 | 0.0716 | 0.1185 |
| `pptx-startup-pitch-en` | 3 | 3/3 | 0.0082 | 0.0302 |
| `pptx-quarterly-review-ko` | 3 | 3/3 | 0.0040 | 0.0043 |
| `pptx-product-launch-en` | 3 | 3/3 | 0.0129 | 0.0162 |
| `pptx-training-deck-ko` | 3 | 3/3 | 0.0132 | 0.0194 |
| `pptx-company-intro-en` | 3 | 3/3 | 0.0060 | 0.0066 |
| `pptx-project-status-ko` | 2 | 2/2 | 0.0053 | 0.0068 |
| `pptx-conference-talk-en` | 2 | 2/2 | 0.0182 | 0.0331 |
| `pptx-marketing-report-en` | 3 | 3/3 | 0.0078 | 0.0278 |
| `pptx-lecture-ko` | 2 | 2/2 | 0.0142 | 0.0252 |
| `pptx-sales-proposal-en` | 2 | 2/2 | 0.0122 | 0.0204 |
| `xlsx-quotation-ko` | 2 | 0/2 | 0.0222 | 0.0361 |
| `xlsx-financial-model-en` | 2 | 0/2 | 0.0298 | 0.0459 |
| `xlsx-inventory-en` | 4 | 0/4 | 0.0718 | 0.1676 |
| `xlsx-payroll-ko` | 2 | 0/2 | 0.0212 | 0.0364 |
| `xlsx-project-schedule-en` | 1 | 0/1 | 0.0368 | 0.0368 |
| `xlsx-sales-dashboard-en` | 1 | 0/1 | 0.0306 | 0.0306 |
| `xlsx-attendance-ko` | 1 | 0/1 | 0.0252 | 0.0252 |
| `xlsx-budget-ko` | 1 | 0/1 | 0.0392 | 0.0392 |
| `xlsx-expense-report-en` | 1 | 0/1 | 0.0346 | 0.0346 |
| `xlsx-kpi-tracker-en` | 1 | 0/1 | 0.0278 | 0.0278 |

## Per-page ledger

| Case | Page | Native px | Converted px | Size match | 5% fuzz diff | Mean channel error | Tracking |
| --- | ---: | ---: | ---: | :---: | ---: | ---: | --- |
| `docx-invoice-en` | 1 | 1240×1754 | 1241×1754 | no | 0.065905 | 11.7216 | [#352](https://github.com/developer0hye/office2pdf/issues/352) |
| `docx-contract-ko` | 1 | 1240×1754 | 1241×1754 | no | 0.091367 | 15.4719 | [#357](https://github.com/developer0hye/office2pdf/issues/357) |
| `docx-meeting-minutes-ko` | 1 | 1240×1754 | 1241×1754 | no | 0.086808 | 15.5720 | [#356](https://github.com/developer0hye/office2pdf/issues/356) |
| `docx-resume-en` | 1 | 1240×1754 | 1241×1754 | no | 0.051652 | 8.2058 | [#352](https://github.com/developer0hye/office2pdf/issues/352) |
| `docx-technical-manual-en` | 1 | 1240×1754 | 1241×1754 | no | 0.048014 | 7.9045 | [#351](https://github.com/developer0hye/office2pdf/issues/351) |
| `docx-technical-manual-en` | 2 | 1240×1754 | 1241×1754 | no | 0.045023 | 8.5381 | [#351](https://github.com/developer0hye/office2pdf/issues/351) |
| `docx-official-letter-ko` | 1 | 1240×1754 | 1241×1754 | no | 0.041144 | 6.7577 | [#352](https://github.com/developer0hye/office2pdf/issues/352) |
| `docx-product-spec-en` | 1 | 1754×1240 | 1754×1241 | no | 0.092553 | 15.9294 | [#355](https://github.com/developer0hye/office2pdf/issues/355) |
| `docx-newsletter-en` | 1 | 1240×1754 | 1241×1754 | no | 0.070917 | 11.4162 | [#368](https://github.com/developer0hye/office2pdf/issues/368) |
| `docx-offer-letter-en` | 1 | 1240×1754 | 1241×1754 | no | 0.038694 | 5.5694 | [#354](https://github.com/developer0hye/office2pdf/issues/354) |
| `docx-research-report-ko` | 1 | 1240×1754 | 1241×1754 | no | 0.118456 | 21.5406 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `docx-research-report-ko` | 2 | 1240×1754 | 1241×1754 | no | 0.024735 | 3.8877 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-startup-pitch-en` | 1 | 2000×1125 | 2000×1125 | yes | 0.002198 | 0.5849 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-startup-pitch-en` | 2 | 2000×1125 | 2000×1125 | yes | 0.030198 | 4.7344 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-startup-pitch-en` | 3 | 2000×1125 | 2000×1125 | yes | 0.008229 | 0.9353 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-quarterly-review-ko` | 1 | 2000×1125 | 2000×1125 | yes | 0.003956 | 0.3932 | [#360](https://github.com/developer0hye/office2pdf/issues/360) |
| `pptx-quarterly-review-ko` | 2 | 2000×1125 | 2000×1125 | yes | 0.004296 | 0.4437 | [#360](https://github.com/developer0hye/office2pdf/issues/360) |
| `pptx-quarterly-review-ko` | 3 | 2000×1125 | 2000×1125 | yes | 0.003826 | 0.3270 | [#360](https://github.com/developer0hye/office2pdf/issues/360) |
| `pptx-product-launch-en` | 1 | 2000×1125 | 2000×1125 | yes | 0.002311 | 0.5894 | [#358](https://github.com/developer0hye/office2pdf/issues/358) |
| `pptx-product-launch-en` | 2 | 2000×1125 | 2000×1125 | yes | 0.016174 | 2.1800 | [#358](https://github.com/developer0hye/office2pdf/issues/358) |
| `pptx-product-launch-en` | 3 | 2000×1125 | 2000×1125 | yes | 0.012936 | 1.9716 | [#358](https://github.com/developer0hye/office2pdf/issues/358) |
| `pptx-training-deck-ko` | 1 | 2000×1125 | 2000×1125 | yes | 0.003746 | 0.3336 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-training-deck-ko` | 2 | 2000×1125 | 2000×1125 | yes | 0.019384 | 2.8647 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-training-deck-ko` | 3 | 2000×1125 | 2000×1125 | yes | 0.013230 | 1.4382 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-company-intro-en` | 1 | 2000×1125 | 2000×1125 | yes | 0.001764 | 0.1789 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-company-intro-en` | 2 | 2000×1125 | 2000×1125 | yes | 0.005952 | 0.5059 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-company-intro-en` | 3 | 2000×1125 | 2000×1125 | yes | 0.006554 | 0.6243 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-project-status-ko` | 1 | 2000×1125 | 2000×1125 | yes | 0.003822 | 0.3388 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-project-status-ko` | 2 | 2000×1125 | 2000×1125 | yes | 0.006754 | 0.5004 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-conference-talk-en` | 1 | 2000×1125 | 2000×1125 | yes | 0.003361 | 0.3272 | [#352](https://github.com/developer0hye/office2pdf/issues/352) |
| `pptx-conference-talk-en` | 2 | 2000×1125 | 2000×1125 | yes | 0.033076 | 4.4657 | [#352](https://github.com/developer0hye/office2pdf/issues/352) |
| `pptx-marketing-report-en` | 1 | 2000×1125 | 2000×1125 | yes | 0.003845 | 0.3512 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-marketing-report-en` | 2 | 2000×1125 | 2000×1125 | yes | 0.007798 | 0.8077 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-marketing-report-en` | 3 | 2000×1125 | 2000×1125 | yes | 0.027758 | 4.3043 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-lecture-ko` | 1 | 2000×1125 | 2000×1125 | yes | 0.003236 | 0.2034 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-lecture-ko` | 2 | 2000×1125 | 2000×1125 | yes | 0.025168 | 3.3270 | [#353](https://github.com/developer0hye/office2pdf/issues/353) |
| `pptx-sales-proposal-en` | 1 | 2000×1125 | 2000×1125 | yes | 0.003966 | 0.7484 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `pptx-sales-proposal-en` | 2 | 2000×1125 | 2000×1125 | yes | 0.020412 | 2.6827 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `xlsx-quotation-ko` | 1 | 1240×1755 | 1241×1754 | no | 0.036112 | 4.2748 | [#366](https://github.com/developer0hye/office2pdf/issues/366) |
| `xlsx-quotation-ko` | 2 | 1240×1755 | 1241×1754 | no | 0.008265 | 0.9968 | [#366](https://github.com/developer0hye/office2pdf/issues/366) |
| `xlsx-financial-model-en` | 1 | 1240×1755 | 1241×1754 | no | 0.013642 | 1.6723 | [#363](https://github.com/developer0hye/office2pdf/issues/363) |
| `xlsx-financial-model-en` | 2 | 1240×1755 | 1241×1754 | no | 0.045923 | 5.4502 | [#363](https://github.com/developer0hye/office2pdf/issues/363) |
| `xlsx-inventory-en` | 1 | 1240×1755 | 1241×1754 | no | 0.167610 | 18.6122 | [#362](https://github.com/developer0hye/office2pdf/issues/362) |
| `xlsx-inventory-en` | 2 | 1240×1755 | 1241×1754 | no | 0.113507 | 15.0435 | [#362](https://github.com/developer0hye/office2pdf/issues/362) |
| `xlsx-inventory-en` | 3 | 1240×1755 | 1241×1754 | no | 0.030108 | 2.2269 | [#362](https://github.com/developer0hye/office2pdf/issues/362) |
| `xlsx-inventory-en` | 4 | 1240×1755 | 1241×1754 | no | 0.016060 | 1.9953 | [#362](https://github.com/developer0hye/office2pdf/issues/362) |
| `xlsx-payroll-ko` | 1 | 1240×1755 | 1241×1754 | no | 0.036424 | 4.1397 | [#366](https://github.com/developer0hye/office2pdf/issues/366) |
| `xlsx-payroll-ko` | 2 | 1240×1755 | 1241×1754 | no | 0.005909 | 0.7581 | [#366](https://github.com/developer0hye/office2pdf/issues/366) |
| `xlsx-project-schedule-en` | 1 | 1240×1755 | 1241×1754 | no | 0.036781 | 4.4868 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `xlsx-sales-dashboard-en` | 1 | 1240×1755 | 1241×1754 | no | 0.030556 | 3.2295 | [#362](https://github.com/developer0hye/office2pdf/issues/362) |
| `xlsx-attendance-ko` | 1 | 1240×1755 | 1241×1754 | no | 0.025225 | 2.5506 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `xlsx-budget-ko` | 1 | 1240×1755 | 1241×1754 | no | 0.039192 | 5.1171 | [PR #369](https://github.com/developer0hye/office2pdf/pull/369) |
| `xlsx-expense-report-en` | 1 | 1240×1755 | 1241×1754 | no | 0.034637 | 4.2642 | [#364](https://github.com/developer0hye/office2pdf/issues/364) |
| `xlsx-kpi-tracker-en` | 1 | 1240×1755 | 1241×1754 | no | 0.027795 | 3.1525 | [#363](https://github.com/developer0hye/office2pdf/issues/363) |

## Source emphasis inventory

Counts are explicit emphasized runs for DOCX/PPTX and emphasized styled cells for XLSX. Theme/default emphasis is not double-counted.

| Case | Bold | Italic | Underline |
| --- | ---: | ---: | ---: |
| `docx-invoice-en` | 29 | 2 | 1 |
| `docx-contract-ko` | 16 | 1 | 1 |
| `docx-meeting-minutes-ko` | 27 | 1 | 1 |
| `docx-resume-en` | 8 | 4 | 1 |
| `docx-technical-manual-en` | 18 | 1 | 1 |
| `docx-official-letter-ko` | 5 | 1 | 1 |
| `docx-product-spec-en` | 51 | 2 | 1 |
| `docx-newsletter-en` | 14 | 1 | 1 |
| `docx-offer-letter-en` | 6 | 2 | 2 |
| `docx-research-report-ko` | 157 | 2 | 1 |
| `pptx-startup-pitch-en` | 6 | 0 | 0 |
| `pptx-quarterly-review-ko` | 10 | 0 | 0 |
| `pptx-product-launch-en` | 11 | 0 | 0 |
| `pptx-training-deck-ko` | 6 | 1 | 0 |
| `pptx-company-intro-en` | 12 | 0 | 0 |
| `pptx-project-status-ko` | 11 | 0 | 0 |
| `pptx-conference-talk-en` | 2 | 0 | 0 |
| `pptx-marketing-report-en` | 6 | 0 | 0 |
| `pptx-lecture-ko` | 5 | 0 | 0 |
| `pptx-sales-proposal-en` | 8 | 1 | 0 |
| `xlsx-quotation-ko` | 25 | 0 | 0 |
| `xlsx-financial-model-en` | 17 | 0 | 0 |
| `xlsx-inventory-en` | 14 | 0 | 0 |
| `xlsx-payroll-ko` | 29 | 0 | 0 |
| `xlsx-project-schedule-en` | 12 | 0 | 0 |
| `xlsx-sales-dashboard-en` | 23 | 0 | 0 |
| `xlsx-attendance-ko` | 24 | 0 | 0 |
| `xlsx-budget-ko` | 25 | 0 | 0 |
| `xlsx-expense-report-en` | 14 | 0 | 0 |
| `xlsx-kpi-tracker-en` | 12 | 0 | 0 |
