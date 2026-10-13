# wps-office

An agent skill (`SKILL.md`) for **WPS Office** by Kingsoft: Writer (`.wps`, `.wpt`),
Spreadsheets (`.et`, `.ett`), Presentation (`.dps`, `.dpt`), PDF, and the Microsoft formats WPS
saves. Unofficial; not affiliated with or endorsed by Kingsoft Office or Zhuhai Kingsoft Office
Software. WPS is a trademark of its owner; Microsoft Office, Word, Excel and PowerPoint are
trademarks of Microsoft.

## Use it when
- A `.wps`, `.et` or `.dps` will not open in Word, Excel or PowerPoint, or a `.wps` may be a
  Microsoft Works file.
- Excel shows `#NAME?` or `=DISPIMG("ID_...",1)` where WPS showed pictures in cells.
- You need a WPS file as `.docx`, `.xlsx`, `.pptx` or PDF, converted in a sandbox.
- You must check a WPS or Office file for VBA or WPS JS (JSA) macros, remote templates, DDE,
  embedded objects, hidden content or metadata (`KSOProductBuildVer`) before opening or sharing.
- You write WPS JS macros, wpsjs add-ins or KDocs AirScript, or run WPS for Linux.

## Install (planned: not published yet)
```bash
npx @skillsaas/cli add wps-office --claude     # planned
npx skills add skillsaas/wps-office            # planned
```
Today: copy this folder into your agent's skills folder (for example `.claude/skills/wps-office`).

## Preflight (copy and run)
```bash
cd <skill-dir>                       # the folder that holds SKILL.md
python3 --version                    # 3.9+ (Windows: py -3 --version)
python3 tests/smoke_test.py          # self-check; skips what is not installed
command -v soffice bwrap             # both needed for conversion (Linux only)
```
Verify before running: read `scripts/*.py` (standard library only, no network code) and
`SECURITY.md`. Without LibreOffice or bwrap, inspection and picture extraction still work;
conversion refuses (exit 3) and never falls back to unsandboxed LibreOffice. On macOS or
Windows, run the converter inside a Linux VM or container.

## Contents
| Path | What |
| :--- | :--- |
| `SKILL.md` | Router: task, first command, reference; safety rules; exit codes |
| `references/` | formats, security, conversion, compatibility, extraction, automation (JSA, wpsjs, AirScript), linux, cloud, fallbacks, troubleshooting, json-output, recipes, sources |
| `scripts/inspect_wps.py` | Read-only inventory and share check |
| `scripts/cellimages.py` | List or save WPS in-cell (DISPIMG) pictures |
| `scripts/convert_wps.py` | LibreOffice conversion inside bubblewrap, no network |
| `tests/smoke_test.py` | Smoke and security tests (`python tests/smoke_test.py`), including seeded fuzz |
| `tests/samples/` | Six real WPS-saved files from the Apache Tika corpus (Apache-2.0; git repo only, not in npm) |

Scripts need Python 3.9+ only; conversion also needs LibreOffice and, on Linux, bubblewrap.
Tested only on Linux; the CI workflow targets Linux, macOS and Windows but has not run yet. WPS Office itself
was not run. Catalog of all skills: https://github.com/skillsaas/index

License: MIT.
