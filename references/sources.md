# Sources and what was tested

Read 2026-10-13. "Tested" means run on the box; everything else is documented, not tested here.

## Tested on the box
- Python 3.13.5, LibreOffice 25.2.3.2, bubblewrap, olefile 0.47 (optional test), Linux x86_64.
- LibreOffice reads `.doc`/`.xls`/`.ppt` renamed to `.wps`/`.et`/`.dps` by content and converts them.
- Unsandboxed LibreOffice fetched an external linked picture from a `.docx` (OPTIONS and GET to
  a local server); the bwrap sandbox in `convert_wps.py` made no request.
- LibreOffice 25.2 profile option names `DisableMacrosExecution`, `MacroSecurityLevel` and
  `BlockUntrustedRefererLinks` exist in `share/registry/main.xcd`.
- Python's expat expands internal entities in a DOCTYPE (why DOCTYPE is refused).
- `npm view wps-office`: 404, the name was free on 2026-10-13. `npm view wpsjs`: 2.2.3.
- Not tested: WPS Office itself (not installed), files saved by WPS, Windows and macOS (CI
  workflow written, not run), the AirScript API, wpsjs.

## Formats and fingerprints
- WPS WebOffice supported formats: https://solution.wps.cn/docs/product/format.html
- Look Scanned, "WPS to PDF" (2005+ layout, Works collision, WPS 2000-2003 and `.wpsx`): https://lookscanned.io/en/blog/wps-to-scanned-pdf
- file-extension.info `.et` page (ZIP and OLE2 variants, Excel cannot open `.et`): https://www.file-extension.info/format/et
- wps2office (LibreOffice-based converter; fidelity and Works notes): https://github.com/AndersonBY/wps2office
- KSOProductBuildVer survives Office edits: https://jishuzhan.net/article/1829343151328858113
- apache/poi issue 1041, WPS-saved `.doc` with FAT past end of file: https://github.com/apache/poi/issues/1041
- Microsoft compound file format: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/ and OpenOffice's description https://www.openoffice.org/sc/compdocfileformat.pdf
- [MS-DOC], [MS-XLS], [MS-PPT], [MS-OFFCRYPTO] on Microsoft Learn (encryption markers)
- Microsoft Works detection: libwps https://sourceforge.net/p/libwps/wiki/Home/ and file(1) magic `Magdir/wordprocessors`

## Corroboration for single-source claims (added 0.1.1)
- 2005+ WPS/ET/DPS use the doc/xls/ppt binary layout: Look Scanned format notes, plus
  https://www.wps2office.com/ notes and https://file-extension.info/ entries for .et and .dps
  (three independent descriptions; the converter test here confirms LibreOffice reads a .doc
  renamed .wps as Writer).
- VBA by edition and mobile macros: see `platforms.md` (several wps.cn answers plus one official article).

## DISPIMG cell pictures
- https://www.hangge.com/blog/cache/detail_3956.html (Excel `#NAME?`, convert to floating picture)
- https://www.shulou.com/a544869 (DISPIMG arguments)
- https://blog.eo2suite.cn/?p=144 (`xl/cellimages.xml` structure, etCustomData namespace)
- https://github.com/exceljs/exceljs/issues/2786 and https://www.yingdao.com/community/detaildiscuss?id=bac9aeac-fb3a-4a1f-a49a-a722c1438fc0 (mapping cells to pictures)

## Macros and automation
- Application object: https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Application/obj
- Document object: https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/obj
- CreateObject: https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/macro-editor-api/function/createobject
- JS runtime (V8, ES6, no browser objects): https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/macro-editor-api/javascript-runtime-instructions
- JSA since WPS 2021: https://developer.aliyun.com/article/1549238
- Macros need macro-enabled formats: https://www.wps.cn/learning/course/detail/id/330744.html
- `xl/JDEData.bin`: https://bbs.wps.cn/topic/36292 ; `Application.Run` from outside: https://bbs.wps.cn/topic/36290
- JSA macro files across platforms: https://forum.wps.cn/topic/93172
- Add-ins with wpsjs: https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/generate-the-first-wps-addin and https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/wps-integration-mode/wps-addin-development/wps-addin-development-instructions
- AirScript script token API: https://airsheet.wps.cn/docs/apitoken/api.html

## Linux and cloud
- ArchWiki WPS Office: https://wiki.archlinux.org/title/WPS_Office
- nixpkgs package: https://github.com/NixOS/nixpkgs/blob/release-25.05/pkgs/by-name/wp/wpsoffice/package.nix
- wps.com install guides: https://www.wps.com/blog/install-wps-office-via-terminal/ and https://www.wps.com/blog/how-to-install-wps-office-on-linux/
- Linux update log (12.1.2.29182): https://www.wps.com/whatsnew/linux/
- Usage figure (676M monthly active devices, June 2026): Kingsoft Office interim report on HKEXnews, linked from TOP_100_SOFTWARE.md
