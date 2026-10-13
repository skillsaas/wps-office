# Converting WPS files

Load this before converting, or when a conversion looks wrong.

## Which route
| Situation | Route | Fidelity |
| :--- | :--- | :--- |
| WPS is installed and the user can click | WPS: File, Save As, choose `.docx`/`.xlsx`/`.pptx` (Writer, Spreadsheets, Presentation) | Best: WPS's own engine |
| Agent, no GUI, Linux with bwrap | `convert_wps.py IN OUT --to docx` | LibreOffice's import of the binary layout |
| macOS or Windows agent | `convert_wps.py` refuses (no bwrap). Ask the user to Save As in WPS, or to run the converter in a Linux VM or container | |
| Online converters | Only with the user's consent for that file and service | Unknown |

WPS Office has no documented command-line converter; the Linux launchers (`wps`, `et`, `wpp`,
`wpspdf`) open files in the GUI (ArchWiki, wps.com install guide).

## Decision tree
1. `inspect_wps.py FILE --json`. `error` or `refused`: stop and report. `works`: only docx, odt, pdf.
2. Encrypted: ask the user to save an unprotected copy in WPS. Stop.
3. Medium or high `detected_risks`: tell the user; convert only if they agree
   (`--allow-risky-input`, exit 1 expected). Prefer asking for a WPS Save As instead.
4. WPS available to the user and fidelity matters: WPS Save As. Else Linux with bwrap:
   `convert_wps.py`. Else macOS/Windows: the converter in a Linux VM or container.
5. DISPIMG pictures in a ZIP spreadsheet: `cellimages.py --extract` before converting.
6. Check the result (fidelity checklist below).

## convert_wps.py
```bash
python3 <skill-dir>/scripts/convert_wps.py report.wps /tmp/out/report.docx --to docx --json
python3 <skill-dir>/scripts/convert_wps.py budget.et /tmp/out/budget.xlsx --to xlsx --json
python3 <skill-dir>/scripts/convert_wps.py deck.dps /tmp/out/deck.pdf --to pdf --json
```
Targets: Writer to `docx`, `odt` or `pdf`; Spreadsheets to `xlsx`, `ods` or `pdf`; Presentation
to `pptx`, `odp` or `pdf`; Microsoft Works `.wps` to `docx`, `odt` or `pdf`. A family mismatch
(for example `.et` to `.docx`) is exit 2. Password-protected input is exit 2: ask the user to
save an unprotected copy in WPS.

The output name must end in the target extension and must not exist (exit 3) unless `--force`
(replaces it, no backup). LibreOffice runs with `--max-memory` (default 4096 MB) and a file-size
cap; only one regular output file is accepted (symlinks refused). The result is written to a temporary name and renamed at the end; the
output is then inspected and its risks reported (`output_check`), so macros or links that
survived are visible.

## What can change in conversion
Not tested against WPS output; these are the documented classes of loss:
- WPS-only features: DISPIMG cell pictures (save them first with `cellimages.py` if the source is
  ZIP-based; binary `.et` must be re-saved in WPS), JS macros, WPS-specific fonts.
- Layout: fonts missing on the converting machine are substituted (wps2office notes; ArchWiki
  for Linux fonts), so line and page breaks move.
- Complex layouts, embedded objects and macros may render differently (wps2office notes).

After converting, spot-check text, page or slide count, sheet names and a few cell values
against the source, and say the engine was LibreOffice.

## Tested here
LibreOffice 25.2.3.2 in the bwrap sandbox: a `.doc` renamed `.wps` became a `.docx` with the
text intact; an `.xls` renamed `.et` became an `.xlsx` with the cell values intact; a `.ppt`
renamed `.dps` became a PDF. These are LibreOffice-made binaries, not WPS-made files.

## Expected loss by content type (LibreOffice import, documented classes, not measured here)
| Content | Expect | Check |
| :--- | :--- | :--- |
| Body text, tables, simple styles | Usually kept | Text spot-check |
| Fonts | Substituted if missing on the converting machine; line and page breaks move | Page count |
| Formulas | Kept as formulas; functions Excel or LibreOffice lack (DISPIMG) become errors | Three key cells |
| Charts | Kept but may restyle | Look at each chart |
| Comments, tracked changes | Usually kept; review them before sending | `inspect_wps.py --share-check` on the output |
| OLE objects | Kept as objects or a picture of them | `embedded_objects` in `output_check` |
| Macros (VBA, JSA) | VBA may survive into docm/xlsm; JSA is not kept | `macros_left` |
| WPS-only layout and cell pictures | Changed or lost | Compare with the original in WPS |

## Fidelity checklist (after every conversion)
- [ ] Page count (Writer, PDF) or slide count matches the original.
- [ ] Sheet names and order match; three or more key cells or totals match.
- [ ] Headers, footers, footnotes and tables are present.
- [ ] Chinese text renders (no boxes): fonts on the converting machine (`linux.md`).
- [ ] DISPIMG pictures were saved separately, or the user knows they are lost.
- [ ] `output_check` shows no macros left; links that survived are listed to the user.
- [ ] Say: "converted with LibreOffice, not WPS".
