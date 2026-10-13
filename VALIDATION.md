# Validation: what was tested against what

Version 0.1.3, 2026-10-13. "Tested" below means a test in `tests/smoke_test.py` or a recorded
manual run on this date. Anything not listed was not tested.

## Environments
| What | Version | Result |
| :--- | :--- | :--- |
| Linux (Debian, x86_64) | Python 3.13.5 and 3.9.25 | Full suite passes, also under `ulimit -v 4000000` |
| LibreOffice | 25.2.3.2 (build 520) | Conversion tests pass inside bwrap |
| bubblewrap | 0.12.0 | Sandbox tests pass, including the network canary |
| GitHub Actions | ubuntu-latest, macos-15-intel, windows-latest; Python 3.9 and 3.13 | All pass. Linux runs all 63 tests (bwrap allowed via sysctl); macOS and Windows skip the conversion and sandbox tests (no bwrap) |
| WPS Office (any OS) | none | Not installed or run. No claim here is checked against WPS itself |
| Other LibreOffice versions | none | Not tested; the converter does not refuse them |

## Files
| Set | Where | What it proves | What it does not |
| :--- | :--- | :--- | :--- |
| 41 synthetic fixtures | built by `tests/make_fixtures.py` at test time | Each detector, limit and refusal on a known input (macros, DDE, links, DOCTYPE, ZIP bombs, traversal, damaged OLE, Works) | That real WPS files look like this |
| 6 WPS-saved Office files | `tests/samples/testWPSAttachment.*` (Apache Tika, Apache-2.0) | Family, markers (`KSOProductBuildVer`), embedded-object detection and conversion on files WPS Office 11.1 really wrote; found two bugs (0.1.2) | Native `.wps/.et/.dps`, DISPIMG, JS macros |
| 1 third-party DISPIMG workbook | `tests/samples/third-party-dispimg-ooxml-excel-editor.xlsx` (MIT) | `cellimages.py` against someone else's reading of the layout; found a bug (0.1.3: non-hex picture names) | That WPS writes it this way (it was generated with ExcelJS) |
| Seeded fuzz | `fuzz_never_crashes` | No traceback, JSON on every path, for mutated ZIP and OLE inputs | Coverage of every parser branch |

## Not validated, and why
- **Native WPS files (.wps, .et, .dps saved by WPS).** Searched on 2026-10-13: Apache POI
  test-data, Apache Tika test documents, LibreOffice core QA data (GitHub code search), GitHub
  code search for the WPS cell-image namespace (`etCustomData`, 60 results, all code or files
  made by other tools), and the web. Found none under a license that allows copying. Apache POI
  issue 1041 attaches WPS-saved `.doc` files that were never merged, so their license is unclear.
  The OLE2 layout claim for these extensions comes from documentation (`references/sources.md`).
- **WPS-saved DISPIMG workbook.** None found under a public license; the only public sample is the
  generated MIT one above.
- **WPS JS macro file (`xl/JDEData.bin` or a JSA project).** None found; the detector is tested on
  a synthetic `jsa.xlsm` only.
- **Rendering fidelity.** Conversion success means LibreOffice produced a file that re-inspects
  cleanly, not that it looks like WPS.

## Reproduce
```bash
ulimit -v 4000000; python3 tests/smoke_test.py
sha256sum -c tests/samples/SHA256SUMS            # run inside tests/samples
```
If you have a WPS-saved native file you may share under an open license, a sample plus its
expected `inspect_wps.py --json` output is the most useful contribution.
