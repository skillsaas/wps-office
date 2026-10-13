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

## Install
```bash
npx skills add skillsaas/wps-office         # skills CLI, reads SKILL.md at the repo root
npm install wps-office                      # files land in node_modules/wps-office/
python3 node_modules/wps-office/scripts/install_skill.py --claude   # copies the skill into ~/.claude/skills/wps-office
# Planned, not yet published: npx @skillsaas/cli add wps-office
```
From a checkout or the npm folder, `scripts/install_skill.py` copies the skill in one step
(standard library only, no network):
```bash
python3 scripts/install_skill.py --claude              # ~/.claude/skills/wps-office (Claude Code, personal)
python3 scripts/install_skill.py --claude-project      # ./.claude/skills/wps-office (this project)
python3 scripts/install_skill.py --dest ~/my-agent/skills   # any agent that reads <dir>/<name>/SKILL.md
```
It copies only what the skill needs, checks every file's SHA-256, and refuses to overwrite an
existing copy (`--replace` moves the old one next to the skills folder; nothing is deleted).

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
| `tests/samples/` | Six real WPS-saved files from the Apache Tika corpus (Apache-2.0) and one third-party DISPIMG workbook (MIT); git repo only, not in npm |
| `scripts/install_skill.py` | One-step install into an agent's skills folder |
| `VALIDATION.md` | What was tested against what, and what was not |

Scripts need Python 3.9+ only; conversion also needs LibreOffice and, on Linux, bubblewrap.
Smoke tests pass in CI on Linux, macOS and Windows (Python 3.9 and 3.13); the sandboxed conversion tests run on Linux only. WPS Office itself
was not run. Catalog of all skills: https://github.com/skillsaas/index

License: MIT.
