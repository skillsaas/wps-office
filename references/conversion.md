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
(replaces it, no backup). The result is written to a temporary name and renamed at the end; the
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
