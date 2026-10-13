# wps-office

**TL;DR:** `wps-office` is a free, MIT-licensed agent skill that teaches AI agents (Claude Code, Cursor, Codex, Gemini CLI and other agents that read SKILL.md skills) to work with WPS Office files: identify `.wps`, `.et` and `.dps` by their bytes, list macros, remote templates, DDE and risky links, save the in-cell pictures Excel shows as `#NAME?` (`DISPIMG`), and convert to `.docx`, `.xlsx`, `.pptx` or PDF with LibreOffice inside a sandbox. Install: `npx skills add skillsaas/wps-office` or `npm install wps-office`.

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

## What it does
- `inspect_wps.py`: read-only report on any WPS or Office file (family, format, WPS markers such as `KSOProductBuildVer`, macros, DDE, remote templates, embedded objects, links, metadata), JSON output, never runs or fetches anything.
- Tells a Kingsoft `.wps` from a Microsoft Works `.wps`.
- `cellimages.py`: lists and saves WPS in-cell pictures (`=DISPIMG("ID_...",1)`) from `.xlsx` and ZIP `.et` files.
- `convert_wps.py`: LibreOffice conversion inside bubblewrap with no network, memory and time limits, and a consent record for risky input (Linux only).
- References for WPS JS macros (JSA), wpsjs add-ins, KDocs AirScript and WPS for Linux, labelled as documented, not run.
- Python 3.9+ standard library only; 63 smoke tests run in CI on Linux, macOS and Windows.

## FAQ
**What is the wps-office skill?**
An Agent Skills package (a `SKILL.md` plus references and Python scripts) that tells an AI agent how to handle WPS Office files and tasks safely. It is unofficial and not affiliated with Kingsoft.

**How do I open a .wps, .et or .dps file without WPS Office?**
Run `python3 scripts/inspect_wps.py FILE --json` to see what the file really is. To read it in Microsoft Office, convert it with `convert_wps.py` (Linux with LibreOffice and bubblewrap) or ask the owner to use Save As `.docx`, `.xlsx` or `.pptx` in WPS.

**Why does Excel show #NAME? or =DISPIMG in a WPS spreadsheet?**
WPS stores in-cell pictures in `xl/cellimages.xml` and refers to them with a `DISPIMG` formula that Excel does not know. `cellimages.py FILE.xlsx --extract NEW_DIR` saves those pictures.

**How do I install it in Claude Code?**
`python3 scripts/install_skill.py --claude` copies it to `~/.claude/skills/wps-office`; `--dest DIR` works for any agent that reads `<dir>/<name>/SKILL.md`.

**Does it run macros or open links?**
No. The inspector and picture tool never run, render or fetch anything, and the converter refuses to run without its sandbox. A clean report means the known checks found nothing; it is not a malware scan.

**Does it work on macOS and Windows?**
Inspection and picture extraction do. Conversion needs Linux with bubblewrap, so on macOS or Windows run it in a Linux VM or container.

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
