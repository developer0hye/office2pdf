# Issue 1647 trace excerpt

The two 70%-opacity groups are copied from page 10 of the public
[LibreOffice reference](https://github.com/user-attachments/files/31185150/GENERAL.SERVICES.libreoffice.pdf)
attached to [issue 1220](https://github.com/developer0hye/office2pdf/issues/1220).
They preserve the background tint and slide-number glyphs responsible for the
false low-contrast result; unrelated page content is omitted.

Reference PDF SHA-256: `8c8f471d8baaf0abd79e752ef49951c240752d36e4512651f6cb9a7cfffb650f`.
Trace command: `mutool draw -F trace -o page10.xml reference.pdf 10`.
