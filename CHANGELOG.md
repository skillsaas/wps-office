# Changelog

## 0.1.0 (2026-10-13)
- First version: SKILL.md router; references for formats, security, conversion, compatibility,
  extraction, automation (JS macros, wpsjs, AirScript), Linux, cloud, fallbacks,
  troubleshooting, JSON output, recipes and sources.
- `inspect_wps.py`: byte-based identification of WPS native, Microsoft binary, OOXML, Works,
  PDF and RTF files; WPS fingerprints; macros (VBA, WPS JS `JDEData`), OLE packages, ActiveX,
  DDE, remote templates, external links; DISPIMG cell pictures; hidden and review content;
  encryption markers; damaged OLE2 chains; `--share-check`.
- `cellimages.py`: maps DISPIMG cells to pictures and saves real pictures to a new folder.
- `convert_wps.py`: LibreOffice conversion in bubblewrap with no network; refuses without it.
- Smoke tests with synthetic fixtures, a network canary and a test proving the threat is real.
