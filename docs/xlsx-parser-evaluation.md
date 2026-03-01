# XLSX Parser Evaluation: calamine vs umya-spreadsheet

**Date**: 2026-03-01
**Context**: Issue #83 — 22 XLSX files fail due to upstream `umya-spreadsheet` bugs
**Purpose**: Evaluate `calamine` as a potential alternative for XLSX parsing

## 1. Library Overview

| Property | umya-spreadsheet | calamine |
|---|---|---|
| Version evaluated | 2.3.3 | 0.33.0 |
| License | MIT | MIT |
| MSRV | — | 1.75.0 |
| Read support | Yes | Yes |
| Write support | Yes | No (read-only) |
| Formats | XLSX | XLSX, XLS, ODS, XLSB |
| Monthly downloads | ~30k | ~619k |

## 2. API Coverage Comparison

| Feature | umya-spreadsheet | calamine | Notes |
|---|---|---|---|
| Cell values (text, number, date) | Yes | Yes | Both extract cell data |
| Cell formatting (font, color, border) | Yes | **No** | calamine is value-only |
| Column widths / row heights | Yes | **No** | Not exposed in API |
| Merged cells | Yes | Yes | calamine added in v0.26.0 |
| Number formats | Yes | Partial | calamine uses internally for datetime detection only |
| Charts | Partial | **No** | umya stores chart data; calamine has no chart support |
| Images | Yes | Yes | calamine via `picture` feature flag |
| Formulas | Yes | Yes | Both can read formula strings |
| Sheet names / selection | Yes | Yes | Both support |
| Conditional formatting | Yes | **No** | Not supported |
| Print area | Yes | **No** | Not supported |
| Headers / footers | Yes | **No** | Not supported |
| Page breaks | Yes | **No** | Not supported |

## 3. Robustness: Failing Fixture Test Results

Tested the 22 XLSX files that fail with `umya-spreadsheet` against `calamine`:

| File | umya-spreadsheet | calamine | Notes |
|---|---|---|---|
| libreoffice/chart_hyperlink.xlsx | PANIC (FileNotFound) | **PASS** | 2 sheets, 36 cells |
| libreoffice/hyperlink.xlsx | PANIC (FileNotFound) | **PASS** | 2 sheets, 0 cells |
| libreoffice/tdf130959.xlsx | PANIC (FileNotFound) | **PASS** | 1 sheet, 0 cells |
| libreoffice/test_115192.xlsx | PANIC (FileNotFound) | **PASS** | 2 sheets, 0 cells |
| poi/47504.xlsx | PANIC (FileNotFound) | **PASS** | 1 sheet, 0 cells |
| poi/bug63189.xlsx | PANIC (FileNotFound) | **PASS** | 4 sheets, 1 cell |
| poi/ConditionalFormattingSamples.xlsx | PANIC (FileNotFound) | **PASS** | 18 sheets, 2469 cells |
| libreoffice/check-boolean.xlsx | PANIC (ParseFloatError) | **PASS** | 1 sheet, 2 cells |
| libreoffice/functions-excel-2010.xlsx | PANIC (ParseIntError) | **PASS** | 2 sheets, 1635 cells |
| poi/FormulaEvalTestData_Copy.xlsx | PANIC (ParseIntError) | **PASS** | 4 sheets, 58057 cells |
| libreoffice/tdf100709.xlsx | PANIC (unwrap on None) | **PASS** | 1 sheet, 156 cells |
| poi/64450.xlsx | PANIC (unwrap on None) | **PASS** | 2 sheets, 16 cells |
| poi/sample-beta.xlsx | PANIC (unwrap on None) | FAIL | shared string index error |
| libreoffice/tdf162948.xlsx | PANIC (dataBar) | **PASS** | 1 sheet, 8 cells |
| poi/NewStyleConditionalFormattings.xlsx | PANIC (dataBar) | **PASS** | 1 sheet, 357 cells |
| libreoffice/forcepoint107.xlsx | ERROR (invalid checksum) | FAIL | range parsing error |
| libreoffice/tdf121887.xlsx | ERROR (ZipError) | **PASS** | 1 sheet, 1 cell |
| libreoffice/tdf131575.xlsx | ERROR (ZipError) | FAIL | unrecognized sheet type |
| libreoffice/tdf76115.xlsx | ERROR (ZipError) | FAIL | unrecognized sheet type |
| poi/49609.xlsx | ERROR (ZipError) | FAIL | unrecognized sheet type |
| poi/56278.xlsx | ERROR (ZipError) | **PASS** | 10 sheets, 889 cells |
| poi/59021.xlsx | ERROR (ZipError) | **PASS** | 1 sheet, 16 cells |

**Summary**: calamine passes **17/22** files (77%) vs **0/22** for umya-spreadsheet.

calamine still fails on 5 files:
- 3 files with non-standard sheet types (calamine doesn't recognize them)
- 1 file with corrupt zip entry (both libraries fail)
- 1 file with invalid shared string reference

## 4. Performance

Not benchmarked in this evaluation. Both libraries are pure Rust and expected to have comparable performance for read operations. calamine is widely used (~619k monthly downloads) and likely well-optimized for read-only workloads.

## 5. Maintenance

| Metric | umya-spreadsheet | calamine |
|---|---|---|
| Active development | Yes | Yes |
| Open issues | ~50 | ~39 |
| Panic reduction planned | Yes (v3.0.0 milestone) | N/A (fewer panics) |
| Maintainer responsiveness | Moderate | Active (jmcnamara) |
| Last release | Recent | Feb 2026 |

## 6. MSRV Compatibility

calamine requires MSRV 1.75.0. Our project's MSRV is 1.85+ (edition 2024). calamine is compatible.

## 7. Recommendation

**Stay with umya-spreadsheet** for now, but monitor calamine for formatting support.

### Rationale

- **calamine cannot extract cell formatting** (fonts, colors, borders, fills, column widths). This is a fundamental requirement for PDF conversion with visual fidelity (PRD §3.1 XLSX P1 features).
- Switching to calamine would **regress formatting quality** — all cell styling would be lost.
- calamine's formatting support is tracked in [calamine#404](https://github.com/tafia/calamine/issues/404) but has no timeline.
- umya-spreadsheet's panic issues are tracked in [umya-spreadsheet#271](https://github.com/MathNya/umya-spreadsheet/issues/271) (v3.0.0 milestone) and our filed issue [umya-spreadsheet#310](https://github.com/MathNya/umya-spreadsheet/issues/310).

### Possible future approaches

1. **Partial hybrid**: Use calamine as a fallback when umya-spreadsheet panics/fails — extract cell values without formatting. This would increase success rate from 0% to 77% on the failing files, with degraded visual quality.
2. **Custom parser**: Build a focused XLSX parser using `quick-xml` + `zip` that extracts exactly the data we need (cell values + formatting). Higher effort but full control.
3. **Wait for upstream fixes**: Monitor umya-spreadsheet v3.0.0 for panic reduction and calamine for formatting support.

For now, our `catch_unwind` wrapper and improved error messages (phase 17) provide adequate mitigation for the 22 failing files (0.8% failure rate).
