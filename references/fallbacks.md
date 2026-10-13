# Fallbacks

Load this when a tool is missing or the file is in an unusual form.

| Situation | Do | Do not |
| :--- | :--- | :--- |
| No Python 3.9+ | Ask the user to install Python 3, or to Save As `.docx`/`.xlsx`/`.pptx` in WPS and send that | Improvise a parser in shell; open the file in an app |
| No LibreOffice | Inspect only; ask the user to Save As in WPS, or to install LibreOffice | Upload to an online converter without consent |
| No bwrap on Linux | Ask the user to install bubblewrap (`apt install bubblewrap`); if user namespaces are disabled, a VM or container is needed | Run soffice unsandboxed |
| macOS or Windows | Inspect and extract cell pictures normally; for conversion ask the user to Save As in WPS, or run the converter inside a Linux VM or container (WSL2, Docker with bubblewrap, UTM); the converter has no bypass flag | Run soffice directly on a received file |
| WPS 2000 to 2003 era files | Not the 2005+ layout (Look Scanned); `inspect_wps.py` may say `unknown`. Ask the user to open and re-save in current WPS | Guess the format |
| `.wpsx`, `.etx`, `.dpsx` | Unknown layout in the sources found; inspect, then ask | Claim support |
| Microsoft Works `.wps` | `convert_wps.py --to docx` (LibreOffice reads Works through libwps) | Rename to `.doc` |
| Password-protected | Ask the user to remove the password in WPS on a copy | Try to crack it |
| Cloud-only document | Ask the user to download a copy (`references/cloud.md`) | Sign in for them |
| Binary `.et` with in-cell pictures | Ask the user to save as `.xlsx` in WPS, then `cellimages.py` | Expect LibreOffice to keep them (not verified) |
