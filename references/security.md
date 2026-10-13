# Security: what is risky in a WPS or Office file, and what to do

Load this for "is it safe", "can I send this", or any high or medium finding.

## Risk levels used by inspect_wps.py
| Risk | Kind | Why |
| :--- | :--- | :--- |
| high | `macros` | VBA project (`vbaProject.bin`, `Macros`, `_VBA_PROJECT_CUR`, `VBA`) or WPS JS macro part (`JDEData*`). Code that runs once macros are enabled |
| high | `macros-under-plain-extension` | Macro code inside a `.docx`/`.xlsx`/`.pptx` name, a common disguise |
| high | `ole-package` | OLE Package (`\x01Ole10Native`) can carry any file, including a program, behind an icon |
| high | `activex` | ActiveX controls |
| high | `dde-field` | `DDE`/`DDEAUTO` field can start a program when fields update |
| high | `external-attachedTemplate` | Remote template fetched on open (template injection, credential leaks) |
| high | `dangerous-link-scheme`, any `file:`/UNC link | `file:`, `smb:`, `javascript:`, `data:` targets; UNC paths can leak Windows credentials (NTLM) |
| high | `rtf-object`, `pdf-active-content` with JavaScript/Launch | Classic exploit carriers |
| medium | `external-image`, `external-oleObject`, `external-frame`, `external-externalLinkPath`, `external-field` | Fetched or updated from elsewhere; tested: LibreOffice 25.2 fetched a linked picture when converting unsandboxed |
| medium | `embedded-objects` | Other files inside the document |
| medium | `encrypted` | Contents could not be inspected |
| medium | `extension-mismatch` | Name and content disagree |
| medium | `very-hidden-sheet` | Sheet hidden from the Unhide list |
| low | `wps-cell-images`, `damaged-structure`, `pdf-active-content` with only OpenAction | Compatibility or structure issues, not attacks by themselves |

`overall_risk` is the highest level found by the checks that ran. It is triage, not a verdict.
`none` means "nothing found by these checks"; say that, never "safe".

## What the scripts never do
They never run macros, render, print, convert (except `convert_wps.py`, sandboxed), open links,
fetch remote content, contact a server, or write anywhere except a new output path.

## Macros in WPS
- WPS Office has two macro languages: VBA, and JS macros (JSA, "JS for Application"), the
  default macro language since WPS 2021 (aliyun article), run by a V8 runtime with ES6 support
  and no browser objects (open.wps.cn).
- Macro code can only be saved in macro-enabled formats (`.docm`, `.xlsm`, `.pptm`; wps.cn help).
  A file named `.docx` that contains macro parts is disguised; the inspector flags it.
- JS macros can call `CreateObject` to reach other WPS apps, and through `Application.Run`
  external programs can call them (bbs.wps.cn 36290). Treat JS macros exactly like VBA.
- Never enable macros on a received file, and never run a macro to "see what it does".

## Before sharing: --share-check
`inspect_wps.py FILE --share-check` lists blockers and exits 1 if there are any:
active content, embedded objects, external links (hyperlinks only when the scheme is risky),
author / last saved by / company / manager / template / comments / keywords, every custom
property (including `KSOProductBuildVer`), tracked changes, comments, hidden text, hidden or
very hidden sheets, hidden slides, an incomplete report, and for binary formats a reminder that
comments and hidden content inside binary streams were not checked.

This skill does not remove anything. Remove blockers in WPS or Microsoft Office (document
inspector / "remove personal information" style features; exact menu names vary by version and
were not verified here), save a copy, and run `--share-check` on the copy until it exits 0. For
binary `.wps`/`.et`/`.dps`, convert to OOXML first so hidden content can be checked.

## Running a converter on an untrusted file
LibreOffice parses and renders the whole document, so it is attack surface. `convert_wps.py`:
- caps the input size before LibreOffice sees it and inspects it first;
- runs LibreOffice in bubblewrap with `--unshare-all` (no network), only `/usr` and `/etc`
  read-only, a private `/tmp`, a fresh HOME and profile, `--clearenv`, and a private copy of the
  input, so the document cannot read the user's files or reach the network;
- sets `DisableMacrosExecution=true`, `MacroSecurityLevel=3` and `BlockUntrustedRefererLinks=true`
  in the throwaway profile (names checked in LibreOffice 25.2's `main.xcd`);
- kills the whole process group at `--timeout`;
- refuses without bwrap (exit 3). `--sandboxed` is for an operator who already isolated the run.

Tested here: a `.docx` with an external linked picture made unsandboxed LibreOffice send
`OPTIONS` and `GET` requests to a local server; the same file through `convert_wps.py` sent none.

## What these tools do not protect against
- Exploits in LibreOffice itself beyond what the sandbox contains.
- Macro carriers in unknown places (for example JS macros inside binary files).
- Malicious content in pictures, fonts or embedded files (listed, never scanned).
- Social engineering in the visible text.
