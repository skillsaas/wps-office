# Troubleshooting

| Symptom | Check | Fix |
| :--- | :--- | :--- |
| `error: not a WPS, Microsoft Office or PDF file` (exit 2) | First bytes | It may be text, an image, a Works 1-3 file renamed, or a WPS 2000-2003 file: ask where it came from |
| `refused: duplicate ZIP entry name` / `collide by case` (exit 3) | ZIP listing | Crafted or broken package; do not "repair" it, ask the sender for a fresh copy |
| `refused: XML part contains DOCTYPE` (exit 3) | Which part | Office and WPS do not write DOCTYPE in OOXML parts; treat as hostile |
| `refused: compression ratio` (exit 3) | Member sizes | Possible ZIP bomb; stop |
| `refused: time limit reached` (exit 3) | Size, structure loops | Raise `--timeout` only if the file is large and trusted by the user |
| `findings`, `incomplete: compound file structure: sector-past-end-of-file` | WPS-saved binary | Re-save in WPS or convert; other readers may reject it |
| `findings`, `incomplete: encrypted` | Password | Ask the user to save an unprotected copy |
| `extension-mismatch` | `format` vs name | Rename only in a copy and only to the real type; tell the user |
| `convert_wps.py` `refused: no sandbox` | OS, `command -v bwrap` | Linux: install bubblewrap. Others: see `fallbacks.md`. There is no bypass flag |
| bwrap fails with "No permissions to create new namespace" | Unprivileged user namespaces disabled (some distros, AppArmor) | Ask the admin, or use a VM or container |
| `LibreOffice produced no output` (exit 2) | stderr in `error` | Damaged or unsupported file; try `--to pdf`; ask for a re-save in WPS |
| `LibreOffice did not finish within --timeout` (exit 3) | Size | Retry once with a longer `--timeout` for big trusted files; the process group was killed |
| Converted file differs from WPS view | Fonts, WPS-only features | `references/compatibility.md`; say the engine was LibreOffice |
| `cellimages.py` `MISSING` cells (exit 1) | `missing` list | Picture id not in `xl/cellimages.xml`: the picture is not in the file |
| `cellimages.py` skipped "bytes are not a known picture type" | The part's bytes | Not saved on purpose; report it as suspicious |
| `cellimages.py` `binary (OLE2) spreadsheet` (exit 2) | `.et` container | Save as `.xlsx` in WPS first |
| JS macro error "ActiveDocument" | No document open | Check `Documents.Count` before using `ActiveDocument` |
| `et`/`wpp` not found on Linux | `dpkg -L wps-office` | Launcher names vary by build (`references/linux.md`) |
| `status: error`, `internal error while reading the file` | The JSON `error` names the exception | A parser bug on a malformed file; the file is likely damaged. Re-save in WPS; report the message |
| `refused: input has detected risks` (convert) | `input_risks` | Ask the user; if they accept, rerun with `--allow-risky-input --consent "<their words>"` (exit 1 is expected) |
| `refused: pictures add up to more than --max-total` | `cells` count and part sizes | Raise `--max-total` only after the user agrees, or extract on a machine with space |
| Converter `refused` or LibreOffice crashes under the memory cap | `--max-memory` | Large decks may need `--max-memory 6144`; do not remove the cap |
| Converter `notes` says --max-memory was clamped | The process runs under `ulimit -v` | Expected: LibreOffice gets the lower limit. If it then fails, raise the ulimit, not the flag |
| Many `dangerous-link-scheme` for `ms-word:` links | SharePoint or Office links | Expected: every `ms-*` handler is high by policy (`security.md`); explain, do not call it malicious |
