# Security: what is risky in a WPS or Office file, and what to do

Load this for "is it safe", "can I send this", or any high or medium finding.

**Read first:** detection is by known storage places and patterns. Unknown or untested macro
carriers (for example JS macros inside binary `.wps/.et/.dps`, objects inside compressed PDF
streams, new link handlers) can be missed. An empty `detected_risks` means "nothing found by
these checks", never "safe". There is no malware scan. Each risk has a `category`: `execution`
(code can run), `fetch` (something is loaded from elsewhere), `deception` (the file misleads),
`privacy`, `coverage` (a check could not see inside), `compatibility`.

## Link scheme policy (version 0.1.2, `scripts/wpslib.py` `link_target`)
1. Percent-decode until stable (at most 8 passes), NFKC-normalize (fullwidth letters fold),
   map hyphen look-alikes (soft hyphen, U+2010 to U+2015, U+2212) to `-`, delete format
   characters (Unicode Cf) and any whitespace or control character inside the scheme token.
2. Allowlist, risk none: `http`, `https`, `mailto`, and relative or `#` targets.
3. High: `HIGH_SCHEMES` (`file`, drive letters, UNC, `smb`, `javascript`, `vbscript`, `data`,
   `mhtml`, `its`, `ms-its`, `shell`, `search`, `search-ms`, `ms-msdt`, `ms-search`, `jar`, `res`,
   `cdl`), every `ms-*` scheme (`ms-officecmd`, `ms-cxh`, `ms-excel:ofe|u|...`, `ms-word`, ...),
   and a scheme that matches one of those with hyphens removed.
4. High: a relative target whose last segment ends in an executable or script extension
   (`EXEC_EXT`: exe, scr, bat, cmd, vbs, js, ps1, hta, lnk, msi, docm, xlsm and more).
5. Medium: any other scheme (`unusual-link-scheme`).
Why every `ms-*`: these launch local Office or Windows handlers, and several (`ms-msdt`, the
`ms-office` URI forms) were exploited. Legitimate uses exist (`ms-word:ofe|u|https://...` in
SharePoint links); expect them flagged, and tell the user why instead of calling them malicious.
A new handler needs a code change here; the list is versioned with the scripts.

## Risk levels used by inspect_wps.py
| Risk | Kind | Why |
| :--- | :--- | :--- |
| high | `macros` | VBA project (`vbaProject.bin`, `Macros`, `_VBA_PROJECT_CUR`, `VBA`) or WPS JS macro part (`JDEData*`). Code that runs once macros are enabled |
| high | `macros-under-plain-extension` | Macro code inside a `.docx`/`.xlsx`/`.pptx` name, a common disguise |
| high | `ole-package` | OLE Package (`\x01Ole10Native`) can carry any file, including a program, behind an icon |
| high | `activex` | ActiveX controls |
| high | `dde-field` | `DDE`/`DDEAUTO` field can start a program when fields update. Field code is joined across runs before matching; RTF `\fldinst` is checked too |
| high | `external-attachedTemplate` | Remote template fetched on open (template injection, credential leaks); OOXML settings relationship or RTF `{\*\template ...}` |
| high | `dangerous-link-scheme` | Any scheme outside http, https, mailto and relative paths that is known dangerous: `file:` (also a drive letter), UNC, `smb:`, `javascript:`, `vbscript:`, `data:`, `ms-msdt:` (Follina), `search-ms:`, `mhtml:`, every `ms-*` handler such as `ms-excel:ofe`. The target is percent-decoded (up to 3 times) and stripped of whitespace and invisible characters first, so ` file:` and `%66ile:` are caught |
| medium | `unusual-link-scheme` | Any other non-allowlisted scheme |
| high | `renamed-macro-part` | VBA found by relationship type or content type under another part name |
| high | `dde-formula`, `call-formula`, `dde-link`, `ole-link` | Spreadsheet `app|topic!item` DDE formulas, `CALL`/`REGISTER`/`EXEC`, and `ddeLink`/`oleLink` in `xl/externalLinks/` |
| high | `odf-event-script` | OpenDocument event listener or script link (`Basic/` and `Scripts/` libraries count as `macros`) |
| medium | `web-formula`, `unknown-binary-or-script-part`, `equation-editor-object`, `text-shows-other-host`, `hidden-characters-in-metadata` (bidi) | Fetch on recalc; an unrecognized container or script part; legacy Equation object; a hyperlink whose visible text names another host; reordering characters in metadata |
| low | `hyperlink-formula`, `formula-like-text` | `HYPERLINK()`; text starting with `= + - @` (CSV injection if exported) |
| high | `rtf-object`, `pdf-active-content` with JavaScript/Launch | Classic exploit carriers |
| medium | `external-image`, `external-oleObject`, `external-frame`, `external-externalLinkPath`, `external-field` | Fetched or updated from elsewhere; tested: LibreOffice 25.2 fetched a linked picture when converting unsandboxed |
| medium | `embedded-objects` | Other files inside the document |
| medium | `encrypted` | Contents could not be inspected |
| medium | `extension-mismatch` | Name and content disagree, including RTF or OLE2 bytes under `.docx`/`.xlsx`/`.pptx` (encrypted OOXML is the one legitimate OLE2-under-`.docx` case) |
| medium | `very-hidden-sheet` | Sheet hidden from the Unhide list |
| low | `wps-cell-images`, `damaged-structure`, `pdf-active-content` with only OpenAction | Compatibility or structure issues, not attacks by themselves |

`overall_risk` is the highest level in `detected_risks`. Both are triage, not a verdict: an
empty list means "nothing found by these checks"; say that, never "safe". There is no malware,
exploit or signature scan.

## Remediation per blocker
| Blocker | What the user does (in WPS or Office, on a copy) |
| :--- | :--- |
| macros, JS macros, renamed macro part | Save as `.docx`/`.xlsx`/`.pptx` (macro-free); check again |
| DDE field or formula, CALL, ddeLink | Delete the field or formula; for links use Data, Edit Links, Break Link (Excel naming; WPS menus vary) |
| external template, image, frame, link | Remove the link or embed the content; for a template, detach it in the template settings |
| embedded object, OLE package | Delete the object or replace it with a picture |
| metadata, custom properties (`KSOProductBuildVer`, `ICV`) | Remove personal information / document properties; WPS writes KSO properties again on every save, so expect them back |
| tracked changes, comments | Accept or reject all changes; delete all comments |
| hidden sheets, rows, columns, text, slides | Unhide and review, then delete what should not be sent |
| incomplete or coverage incomplete | Convert to OOXML with `convert_wps.py` and check the result |
Exact menu names were not verified in WPS here.

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
comments and hidden content inside binary streams were not checked. It is a privacy and
content review, not a malware scan.

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
- sets an address-space cap (`--max-memory`, default 4096 MB) and a file-size cap with setrlimit;
- copies out only a single regular file (lstat, then `O_NOFOLLOW`), never a symlink;
- refuses inputs with medium or high detected risks unless `--allow-risky-input` (still exit 1);
- refuses without bwrap (exit 3). There is no bypass flag (0.1.0 had `--sandboxed`; removed).

Tested here: a `.docx` with an external linked picture made unsandboxed LibreOffice send
`OPTIONS` and `GET` requests to a local server; the same file through `convert_wps.py` sent none.

## Sandbox boundary, precisely
What bwrap gives LibreOffice: no network namespace (`--unshare-all`), an empty environment,
read-only `/usr` and `/etc` (system files readable, not writable), a private `/tmp`, a private
work folder holding only the input copy, a fresh profile and output folder, and nothing else
of the user's home. The whole process group is killed at `--timeout`; RLIMIT_AS and
RLIMIT_FSIZE cap memory and file size. Only one regular output file is copied out (no symlinks).
What remains exposed: the Linux kernel (a kernel or bwrap escape bug defeats it), LibreOffice
bugs inside the sandbox (they can read `/usr` and `/etc`, use CPU until the timeout, and write
a hostile output file, which is then inspected and still untrusted), and anything the input
itself contains. If bwrap exists but user namespaces are disabled, bwrap fails to start and the
converter reports an error rather than running unsandboxed.
`--allow-risky-input` changes none of this: it only lets the conversion be attempted.

## Reproduce the canary
`tests/smoke_test.py` tests `canary_threat_is_real_without_sandbox` and
`canary_no_network_in_sandbox`: a local HTTP server on 127.0.0.1 and a `.docx` whose picture is
linked to it (`tests/make_fixtures.py` `canary_docx`). Expected: plain `soffice --convert-to pdf`
makes requests (OPTIONS and GET seen with LibreOffice 25.2.3.2); `convert_wps.py` makes none.

## Extracted and converted files are untrusted
Pictures saved by `cellimages.py` and files written by `convert_wps.py` come from the same
untrusted document. Do not open, render, run or upload them automatically; inspect converted
output (the converter does) and treat pictures as data.

## What these tools do not protect against
- Exploits in LibreOffice itself beyond what the sandbox contains.
- Macro carriers in unknown places (for example JS macros inside binary files).
- Malicious content in pictures, fonts or embedded files (listed, never scanned).
- Social engineering in the visible text.
