# Diagonal cell-border defect evidence

Related: #2167. `compare.jpg` is native Word (left) versus filing-time main
(right), page 9, rendered with pdftoppm at 150 DPI, without resizing or annotations.
The final image is metadata-stripped progressive JPEG, quality 86, 2482 × 1755 px.

Source: [docxide-pdf snapshot](https://github.com/sverrejb/docxide-pdf/tree/adb8c2b82baad8deb2973d1d059bfb7513b8e1fe/tests/fixtures/scraped/strategi_pengembangan_information_center_resort).
Input SHA-256: `ba11097d3b97feecc8ca108a227e624ed7cdb7ad843d537ffd65d3ddce9ebb0b`.
Word reference: `b11096d871edd3df036eaffc301596a649ec2e6fe9bda02f79f184d7a356b183`.
Main output: `15bbb9aa0baf6d87528d4bbc3d9a18cac671760ebaf30d0e32f067ca158f8634`.
Filing main: `2542b9f8309051daa03d18db6485c221964c8c99`; its runtime source is
identical to the captured baseline build. Native provenance and GT integrity
were checked; no Office version beyond the upstream PDF metadata is asserted.

The native first SWOT cell has one black 0.5pt diagonal; main has none.
The source declares `w:tl2br`, `single`, `sz=4`, `color=auto`.
Repeated header content (#2163), automatic-row border gaps (#2145), omitted
list labels, body/table redistribution and 12-versus-13 pagination remain
separate findings. This image documents the defect; it is not an after image,
a full-file visual acceptance, a strict-cluster pass or a performance result.

Attribution: aldo tambalo, Junaedin Wadu and Diecky Arif Rachman (2025),
*Strategi Pengembangan Information Center Resort Kambata Wundut Taman Nasional
Manupeu Tanah Daru*, Prosiding Seminar Nasional SATI 4(1), 20–31.
[Publisher and license](https://ojs.unkriswina.ac.id/index.php/semnas-FST/article/view/1310):
[CC BY 4.0](https://creativecommons.org/licenses/by/4.0/).
The right panel is an altered conversion; no author/publisher endorsement is implied.
Only this public page comparison is submitted; no source DOCX/PDF or fonts are included.
