# Real files saved by WPS Office (third-party, Apache License 2.0)

These six files were saved by WPS Office and are copied unchanged from the Apache Tika test
corpus, where they were added for TIKA-3526 ("can't extract content from attachments in
Office docs created by WPS"):

- Source: https://github.com/apache/tika/tree/main/tika-parsers/tika-parsers-standard/tika-parsers-standard-modules/tika-parser-microsoft-module/src/test/resources/test-documents
  (`testWPSAttachment.doc`, `.xls`, `.ppt`, `.docx`, `.xlsx`, `.pptx`), downloaded 2026-10-13.
- Issue: https://issues.apache.org/jira/browse/TIKA-3526
- License: Apache License 2.0 (Apache Tika); see https://www.apache.org/licenses/LICENSE-2.0.
  Not covered by this repository's MIT license. Copies: `LICENSE-APACHE-2.0.txt` and `NOTICE-tika.txt`
  (from the Tika repository root).

Their metadata names WPS as the producer (`WPS Office_11.1.0.10700`, `WPS 演示`, `WPS 表格`),
and each carries `KSOProductBuildVer`. Each embeds other Office files, PDF, XML and text as
OLE objects, so the inspector rates them high (`ole-package`). They are not malicious.
`SHA256SUMS` lists their hashes. They are kept out of the npm package.

No WPS-saved native `.wps`, `.et` or `.dps` file, and no WPS file with DISPIMG pictures or JS
macros, was found under a public license (searched: Apache POI test-data, LibreOffice core QA
data via GitHub code search, Apache Tika). Apache POI issue 1041 attaches WPS-saved `.doc`
files, but they were never merged into POI's test-data, so their license is unclear and they are
not included.

## Independent DISPIMG sample (third-party, MIT, not WPS-saved)

`third-party-dispimg-ooxml-excel-editor.xlsx` is copied unchanged from
https://github.com/ojaDeveloper/ooxml-excel-editor (`public/wps-dispimg-sample.xlsx`, commit
c672235, 2026-06-05), MIT license, copy in `LICENSE-MIT-ooxml-excel-editor.txt`. Its author
generated it with ExcelJS and injected `xl/cellimages.xml` (script `scripts/gen-wps-sample.mjs`
in that repository); WPS did not save it. It is useful because someone else wrote it from
their own reading of the WPS layout: it uses a non-hex picture name (`ID_demo_0001`) and an
XML-escaped formula, which exposed a bug in `cellimages.py` 0.1.2 (fixed in 0.1.3).
