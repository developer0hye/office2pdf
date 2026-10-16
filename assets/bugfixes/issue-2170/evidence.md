# Geometry-only theme-paint defect evidence

Related: #2170. `compare.jpg` shows native Word (left) and filing-time main
(right), full page 8, rendered at 150 DPI without resizing or annotations.
It is a metadata-stripped progressive JPEG, quality 86, 2482 × 1755 px.
SHA-256: `4423ac38d7fc3f151c81e3530124ece16fab0d81a25831fc01b15722c0c13814`.

Source: [pinned docxide-pdf fixture](https://github.com/sverrejb/docxide-pdf/tree/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/tests/fixtures/scraped/strategi_pengembangan_information_center_resort).
Input SHA-256: `ba11097d3b97feecc8ca108a227e624ed7cdb7ad843d537ffd65d3ddce9ebb0b`.
Word reference: `b11096d871edd3df036eaffc301596a649ec2e6fe9bda02f79f184d7a356b183`.
Main output: `15bbb9aa0baf6d87528d4bbc3d9a18cac671760ebaf30d0e32f067ca158f8634`.
Filing main: `2542b9f8309051daa03d18db6485c221964c8c99`. Native PDF provenance
and GT integrity were checked; no Office version beyond its metadata is asserted.

The native matrix has a black star in its first cell; main paints no star.
The source shape has `star5` geometry and theme fill/line references without
direct paint. A one-factor converter probe adding only a direct black fill
restores an object without changing matched page-8 text layout; that diagnostic
object is a rectangle because preset geometry falls back independently.
This is source isolation, not a fix or native acceptance. Vertical-rule gaps
(#2145), table-style line spacing (#2147), repeated headers (#2163), list labels
(#2168), the diagonal (#2167), body clearance (#2056), and remaining flow
findings stay separate. No strict-cluster pass or whole-file acceptance is claimed.

Attribution: aldo tambalo, Junaedin Wadu and Diecky Arif Rachman (2025),
*Strategi Pengembangan Information Center Resort Kambata Wundut Taman Nasional
Manupeu Tanah Daru*, Prosiding Seminar Nasional SATI 4(1), 20–31.
[Publisher and license](https://ojs.unkriswina.ac.id/index.php/semnas-FST/article/view/1310):
[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/).
The right panel is an altered conversion; no author/publisher endorsement is implied.
Only this public page comparison is submitted. It has no personal contact details;
no classified fixture, source DOCX/PDF, credentials, or fonts are included.
