# File formats and how to tell them apart

Load this when a file's type is unclear, or before explaining why another app cannot open it.

## The WPS native extensions
| App (WPS name) | Document | Template | Linux launcher | Laid out like (2005 and later) |
| :--- | :--- | :--- | :--- | :--- |
| Writer (WPS 文字) | `.wps` | `.wpt` | `wps` | Word 97-2003 `.doc` |
| Spreadsheets (WPS 表格) | `.et` | `.ett` | `et` | Excel 97-2003 `.xls` (BIFF) |
| Presentation (WPS 演示) | `.dps` | `.dpt` | `wpp` | PowerPoint 97-2003 `.ppt` |
| PDF | `.pdf` | | `wpspdf` | PDF |

Sources: WPS WebOffice supported-formats page (lists `.et`/`.ett`, `.dps`/`.dpt` and the MS
formats; says `.et` is compatible with `.xls`/`.xlsx`); Look Scanned's write-up (2005+ files are
"built almost exactly like" `.doc`/`.xls`/`.ppt`; WPS 2000 to 2003 files used a different format;
`.wpsx`, `.etx`, `.dpsx` exist); ArchWiki (launchers). See `sources.md`.

WPS also saves and opens `.docx`, `.xlsx`, `.pptx`, their macro-enabled forms (`.docm`,
`.xlsm`, `.pptm`), `.doc`, `.xls`, `.ppt`, templates, `.rtf` and `.pdf`. Many WPS users save in the
Microsoft formats by default, so "a WPS file" is often an ordinary `.docx` with WPS fingerprints.

## Identify by bytes, not the name
| First bytes | Container | What `inspect_wps.py` then looks for |
| :--- | :--- | :--- |
| `D0 CF 11 E0 A1 B1 1A E1` | OLE2 compound file (MS-CFB) | Stream names: `WordDocument` (Writer/.doc), `Workbook` or `Book` (Spreadsheets/.xls), `PowerPoint Document` (Presentation/.ppt), `MN0` (Microsoft Works), `EncryptionInfo` + `EncryptedPackage` (password-protected OOXML) |
| `50 4B 03 04` (`PK`) | ZIP | `[Content_Types].xml` and `word/`, `xl/` or `ppt/` main parts; `mimetype` means OpenDocument |
| `%PDF-` | PDF | Action names (`/JavaScript`, `/Launch`, `/OpenAction`, `/EmbeddedFile`) by byte scan |
| `{\rtf` | RTF | `\objdata` (embedded OLE object) |
| bytes `01 FE` at offset 112 | Microsoft Works 1 to 3 (non-OLE) | nothing more; it is not a WPS file (file(1) magic; not tested on a real Works file) |

Tested here: LibreOffice 25.2.3.2 opened a `.doc`, `.xls` and `.ppt` renamed to `.wps`, `.et`
and `.dps` by content (Writer, Calc and Impress respectively) and converted them. That shows
LibreOffice sniffs content; it does not prove every WPS-made file converts cleanly.

## The `.wps` collision with Microsoft Works
Microsoft Works used `.wps` for its own unrelated word processor format. Works 4 files are OLE2
with an `MN0` stream (libwps); Works 1 to 3 files are not OLE2. WPS Office does not claim to
open them; LibreOffice reads them through libwps. `inspect_wps.py` reports `family: works`.

## WPS fingerprints
| Where | What | Notes |
| :--- | :--- | :--- |
| `docProps/custom.xml` (OOXML) | custom property `KSOProductBuildVer`, value like `2052-11.1.0.9662` (2052 is the zh-CN locale id) | Added when WPS creates or saves the file, and it stays when Microsoft Office later edits it (jishuzhan.net test) |
| `docProps/custom.xml` | `ICV` | Seen beside `KSOProductBuildVer` in WPS files; meaning not documented |
| OLE2 property streams | strings `WPS Office_11.1.0.10009`, `KSOTemplate`, `KSOProductBuildVer` | Found in a WPS-made `.doc` (apache/poi issue 1041) |
| `docProps/app.xml` `Application` | often `WPS Office` or `WPS Writer` | Not reliable: WPS may write the Microsoft name |
| XML namespace `http://www.wps.cn/officeDocument/2017/etCustomData` | WPS-only parts such as `xl/cellimages.xml` | |

A fingerprint says WPS touched the file, not who wrote it last. It is personal-ish metadata
(product build and locale), so `--share-check` lists custom properties as blockers.

## WPS-only parts inside OOXML
| Part | Purpose | Effect in Microsoft Office |
| :--- | :--- | :--- |
| `xl/cellimages.xml` + `xl/_rels/cellimages.xml.rels` | Pictures placed inside cells; cells hold `_xlfn.DISPIMG("ID_<32 hex>",1)` | Excel has no DISPIMG: `#NAME?`, pictures not shown. Second argument: 0 crop, 1 fit (shulou write-up) |
| `xl/JDEData.bin` | WPS JS (JSA) macro code for a workbook, plain text, tied in by a relationship and a content-type override | Ignored by Excel; still code. Only created once the macro editor saved code (bbs.wps.cn 36292) |

Where WPS stores JS macros inside Writer and Presentation files, and inside binary
`.wps`/`.et`/`.dps`, is not documented in the sources found. The inspector flags any part named
`JDEData*` or `jsaProject*` anywhere, plus all VBA storages; other carriers may be missed.

## Damaged OLE2 structure in WPS-saved binaries
apache/poi issue 1041 (April 2026): a `.doc` saved by WPS 11.1.0.10009 had a valid header but a
FAT chain pointing past the end of the file; Apache POI threw `Block 52 not found` and
Outlook.com called it corrupt, while WPS opened it, and "Save As .doc" in WPS fixed it.
`inspect_wps.py` reads such files, lists `sector-past-end-of-file` under `incomplete`, and
exits 1 (or 0 with `--allow-incomplete`). Tested on a synthetic fixture only.

## Encryption markers (detected, never decrypted)
| Format | Marker | Source |
| :--- | :--- | :--- |
| Writer / `.doc` | FIB flag `fEncrypted` (bit 0x0100 of the 16-bit word at offset 10 of `WordDocument`) | [MS-DOC] FibBase |
| Spreadsheets / `.xls` | `FILEPASS` record (0x002F) in the first substream | [MS-XLS] |
| Presentation / `.ppt` | `EncryptedSummary` stream | [MS-PPT]; not tested on a real file |
| OOXML | OLE2 file with `EncryptionInfo` and `EncryptedPackage` | [MS-OFFCRYPTO] |

## Limits the inspector enforces
| Limit | Default | Option |
| :--- | :--- | :--- |
| Input size | 100 MB | `--max-bytes` |
| Wall time | 60 s | `--timeout` |
| ZIP entries | 5000 | fixed |
| ZIP total uncompressed | 512 MB | fixed |
| ZIP member | 256 MB, ratio 200 for members over 10 MB | fixed |
| OLE2 directory entries | 4096 | fixed |
| Directory tree depth | 32 | fixed |
