# Hidden header picture is drawn

`compare.jpg` shows page 1 of the public upstream budget letter: supplied
Microsoft Word reference on the left, main `da90c1c1` output on the right.
Both panels are full-page renders at 150 DPI. No source document or PDF is copied.

Source: `sverrejb/docxide-pdf` at `fe7b8beafa7612cc89ad62a7d27b2c6141c4d6ed`,
`tests/fixtures/scraped/dutch_government_budget_letter/input.docx`.
Input SHA256: `20567e12c37c5f45455fe790d508a9476e82ff0e1cb900a247870ee4201d7260`.
Native PDF SHA256: `f962ee0e6ce4f74c51b480150f57d773f44b5988d6e0da7688112ee5fd2e665c`.
Output PDF SHA256: `cebb26b47df46c114b162afd0e698b786151e072fa3b27dffe2c199a6524b9aa`.

The first-page header anchors a 36.75 × 111pt picture whose `wp:docPr` states
`hidden="1"`. Word draws no such image on the page. Main draws it at the margin
origin (79.4, 119.9), where its blue band covers the start of the address lines.
This records a defect; conversion code is unchanged. Other deviations on the
page are listed, unattributed, in the issue.
