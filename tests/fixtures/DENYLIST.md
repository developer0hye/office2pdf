# Denylisted Test Fixtures

Fixtures excluded from bulk testing because they are adversarial, trigger
upstream OOM/hangs, or contain intentional security payloads. They are listed
in `crates/office2pdf/tests/bulk_conversion.rs::DENYLIST`.

Related: [#77](https://github.com/developer0hye/office2pdf/issues/77)

## DOCX

| File | Source | Size | Reason |
|------|--------|------|--------|
| `poi/clusterfuzz-testcase-minimized-POIXWPFFuzzer-6733884933668864.docx` | Apache POI (ClusterFuzz) | 20 KB | Fuzzer-generated corrupted file with invalid checksum. Triggers upstream panic in `docx-rs`. |

## XLSX

| File | Source | Size | Reason |
|------|--------|------|--------|
| `poi/clusterfuzz-testcase-minimized-POIXSSFFuzzer-5265527465181184.xlsx` | Apache POI (ClusterFuzz) | 10 KB | Fuzzer-generated corrupted file. Triggers upstream panic in `umya-spreadsheet`. |
| `poi/clusterfuzz-testcase-minimized-POIXSSFFuzzer-5937385319563264.xlsx` | Apache POI (ClusterFuzz) | 53 KB | Fuzzer-generated corrupted file. Triggers upstream panic in `umya-spreadsheet`. |
| `poi/poc-xmlbomb.xlsx` | Apache POI | 5.1 KB | XML billion-laughs attack PoC. Entity expansion (`<!ENTITY>`) causes exponential memory growth during XML parsing. |
| `poi/poc-xmlbomb-empty.xlsx` | Apache POI | 5.1 KB | Variant of XML billion-laughs attack with empty content. Same entity expansion issue. |
| `poi/54764.xlsx` | Apache POI | 7.8 KB | XML bomb variant using lol9 entity expansion (Apache POI bug 54764). |
| `poi/54764-2.xlsx` | Apache POI | 7.9 KB | Second variant of Apache POI bug 54764 XML bomb. |
| `poi/poc-shared-strings.xlsx` | Apache POI | 107 KB | Shared string table bomb. `xl/sharedStrings.xml` contains a massive number of entries that cause OOM when `umya-spreadsheet::read_reader()` loads them. |
| `libreoffice/too-many-cols-rows.xlsx` | LibreOffice | 5.4 KB | Extreme dimension metadata (1024 cols x 1,048,576 rows). The `umya-spreadsheet::read_reader()` call hangs/OOMs before our guards can run. |
| `poi/bug62181.xlsx` | Apache POI | 854 KB | Complex workbook that triggers OOM/hang in `umya-spreadsheet` during parsing. |
