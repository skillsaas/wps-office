# Worked examples (real output)

Each output below was produced on the box on 2026-10-13 with version 0.1.1 of the scripts
(Python 3.13.5, LibreOffice 25.2.3.2). Inputs 1 to 3 are the synthetic fixtures from
`tests/make_fixtures.py` under friendlier names; input 4 is an `.xls` made by LibreOffice and
renamed `.et`. None of them was saved by WPS itself. Long arrays are trimmed where marked.

## 1. "What is this .wps? Word will not open it."
```bash
python3 <skill-dir>/scripts/inspect_wps.py report.wps --json
```
```json
{"ok": true, "status": "ok", "container": "cfb", "family": "writer",
 "format": "Word 97-2003 binary layout (.doc; WPS Writer .wps/.wpt use the same layout)",
 "extension": ".wps", "extension_matches": true,
 "wps_markers": ["KSOProductBuildVer", "WPS Office"], "producer": "WPS Office",
 "metadata": {"author": "Zhang San", "last_saved_by": "Li Si", "application": "WPS Office", "company": "Example Co"},
 "active_content": [], "embedded_objects": [], "external_links": [], "detected_risks": [],
 "overall_risk": "none", "complete": true, "coverage_complete": false, "incomplete": [], "exit_reasons": [],
 "not_checked": ["...", "binary formats: where WPS stores JS (JSA) macros inside .wps/.et/.dps is not documented; only VBA storages are detected"],
 "exit_code": 0}
```
Tell the user: it is a WPS Writer file in the old Word binary layout, made with WPS; no macros
or links were found by these checks (`coverage_complete` is false: binary streams are not fully
parsed, so this is not a safety verdict). Offer to convert it
to `.docx` (recipe 4) or ask them to Save As `.docx` in WPS.

## 2. "Excel shows #NAME? where the pictures were."
```bash
python3 <skill-dir>/scripts/cellimages.py prices.xlsx --extract /tmp/prices-pictures --json
```
```json
{"ok": true, "status": "findings",
 "cells": [
  {"sheet": "Sheet1", "sheet_index": 1, "cell": "A1", "image_id": "ID_A16C8CE6E39B4F25934D584A1C41E1E1", "media": "xl/media/image1.png"},
  {"sheet": "Sheet1", "sheet_index": 1, "cell": "B1", "image_id": "ID_0088DCB6A4964EEE86E72379D016275C", "media": "xl/media/image2.png"},
  {"sheet": "Sheet1", "sheet_index": 1, "cell": "B2", "image_id": "ID_FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF", "media": null}],
 "missing": ["Sheet1!B2"],
 "saved": [{"file": "sheet1_A1.png", "cells": ["Sheet1!A1"], "shared_by": 1, "bytes": 67, "type": "png", "width": 1, "height": 1}],
 "skipped": [{"cell": "B1", "reason": "bytes are not a valid picture header"}], "notes": [],
 "output_dir": "/tmp/prices-pictures", "pictures_in_part": 2, "exit_code": 1}
```
Tell the user: these are WPS in-cell pictures (DISPIMG), which Excel cannot show. One picture
was saved; B1's "picture" is not an image (not saved, worth a look); B2 points to a picture that
is not in the file. For Excel, the lasting fix is in WPS: convert the pictures to floating
pictures and save again.

## 3. "Can I send this offer letter to the client?"
```bash
python3 <skill-dir>/scripts/inspect_wps.py offer.docx --share-check --json
```
```json
{"overall_risk": "high",
 "detected_risks": [
  {"risk": "medium", "kind": "hidden-characters-in-metadata", "detail": "author contains bidi and invisible characters (removed when printed)"},
  {"risk": "high", "kind": "dangerous-link-scheme", "detail": "hyperlink with file: target"},
  {"risk": "high", "kind": "external-attachedTemplate", "detail": "remote template is fetched when the document opens (template injection); host attacker.example, scheme https"},
  {"risk": "high", "kind": "dde-field", "detail": "DDE or DDEAUTO field: can start a program when fields update"}],
 "external_links": [
  {"type": "hyperlink", "scheme": "file", "host": "", "part": "word/_rels/document.xml.rels"},
  {"type": "attachedTemplate", "scheme": "https", "host": "attacker.example", "part": "word/_rels/settings.xml.rels"}],
 "share_check": {"passed": false, "blockers": [
  "active content: dde-field (Writer field)", "external link: hyperlink to file",
  "external link: attachedTemplate to attacker.example", "metadata: author",
  "tracked changes", "comments", "hidden content: hidden_text_runs"],
  "note": "privacy and content review, not a malware scan; nothing was removed"},
 "exit_reasons": ["share-blockers"], "exit_code": 1}
```
Tell the user: do not send it, and do not open it with editing enabled. It has a DDE field and
a remote template (both can act when the file opens), a link to a local program, tracked
changes, comments, hidden text and an author name containing hidden bidi characters. If they wrote it themselves, those
must be removed in WPS on a copy and the check run again; if it came from someone else, treat it
as hostile.

## 4. "Turn this budget.et into an Excel file."
```bash
python3 <skill-dir>/scripts/convert_wps.py budget.et /tmp/out/budget.xlsx --to xlsx --json
```
```json
{"ok": true, "status": "ok", "exit_reasons": [], "input_family": "spreadsheet",
 "engine": "LibreOffice at /usr/lib/libreoffice/program/soffice",
 "sandbox": "bwrap (no network, private HOME and /tmp, /usr and /etc read-only)",
 "output": "/tmp/out/budget.xlsx",
 "output_check": {"family": "spreadsheet", "overall_risk": "none", "detected_risks": [], "complete": true, "macros_left": []},
 "input_risks": [],
 "not_checked": ["layout, fonts, formulas and charts are LibreOffice's rendering, not WPS's: compare with the original",
                 "WPS-only features (DISPIMG cell pictures, WPS fonts, JS macros) may be dropped or changed"],
 "exit_code": 0}
```
Tell the user: converted with LibreOffice in a sandbox, not with WPS; open both and compare a
few totals. In the smoke test the same route kept the cell values of a test sheet.

## 5. Bad case: "Convert this WPS spreadsheet" and it carries embedded objects
Input: the real WPS-saved `tests/samples/testWPSAttachment.xls` (Apache Tika corpus).
```bash
python3 <skill-dir>/scripts/convert_wps.py budget.xls /tmp/out/budget.xlsx --to xlsx --json
```
```json
{"ok": false, "status": "refused", "exit_reasons": ["input-risks"], "exit_code": 3,
 "error": "input has detected risks (ole-package); convert only if the user accepts that, with --allow-risky-input --consent \"...\""}
```
Tell the user: the file embeds other files as OLE packages (they can carry programs); nothing
was converted. Ask: "Convert anyway in the sandbox? The embedded files are not opened." Only
after a clear yes:
```bash
python3 <skill-dir>/scripts/convert_wps.py budget.xls /tmp/out/budget.xlsx --to xlsx --allow-risky-input \
  --consent "User said: convert it, I made this file" --json
```
```json
{"status": "findings", "exit_reasons": ["input-risks"], "exit_code": 1,
 "conversion_attempted": true, "output_written": true,
 "consent": {"text": "User said: convert it, I made this file", "risks_accepted": ["ole-package"],
   "time_utc": "2026-10-13T15:01:17Z", "input_sha256": "d96a5644...78507"},
 "engine_version": "LibreOffice 25 build 520(Build:2) The Document Foundation/Debian", "memory_limit_mb": 4096,
 "output_check": {"family": "spreadsheet", "overall_risk": "none", "detected_risks": [], "complete": true, "macros_left": []}}
```
Tell the user: converted with LibreOffice in the sandbox. The output check found no embedded
objects, so they were most likely dropped in conversion: if the user needs them, keep the
original. Exit 1 here is expected (the input had risks), not a failure.

## 6. Bad case: a damaged or hostile package
```bash
python3 <skill-dir>/scripts/inspect_wps.py big.xlsx --json
```
```json
{"ok": false, "status": "refused", "error": "ZIP member compression ratio over 200: xl/workbook.xml", "exit_code": 3}
```
Tell the user: the file looks like a ZIP bomb (a small file that expands enormously); it was not
unpacked. Do not raise limits to "try again"; ask where the file came from.

## 7. Bad case: "Check this file" and it is password-protected
```bash
python3 <skill-dir>/scripts/inspect_wps.py offer.docx --json
```
Real 0.1.3 output (fixture `encrypted_ooxml.docx`, trimmed):
```json
{"status": "findings", "family": "encrypted-ooxml",
 "format": "password-protected OOXML package (EncryptionInfo + EncryptedPackage)",
 "complete": false, "incomplete": ["encrypted"], "coverage_complete": false,
 "detected_risks": [{"risk": "medium", "kind": "encrypted", "category": "coverage",
   "detail": "password-protected: contents cannot be inspected without the password"}],
 "exit_reasons": ["incomplete"], "exit_code": 1}
```
Say: "This file is password-protected, so I could not look inside it at all: no result about
macros or links is possible. If you trust the sender, open it in WPS with the password, save an
unprotected copy under a new name, and I will check that copy." Never call it clean, never try
to guess or crack the password, and do not convert it (the converter refuses encrypted input).

## 8. Partial coverage: a clean list that is not the whole story
```bash
python3 <skill-dir>/scripts/inspect_wps.py report.wps --json
```
Real 0.1.3 output (fixture `kingsoft.wps`, trimmed):
```json
{"status": "ok", "family": "writer", "wps_markers": ["KSOProductBuildVer", "WPS Office"],
 "complete": true, "coverage_complete": false,
 "coverage_gaps": ["binary-streams-not-parsed", "jsa-in-binary-unknown"],
 "detected_risks": [], "exit_code": 0}
```
`complete: true` means the file was read without damage; `coverage_complete: false` means some
checks do not exist for this layout. Here the Writer text stream is not parsed (so DDE fields or
links inside it are not seen), and where WPS keeps JS macros in a binary file is not known. Say:
"This is a WPS Writer file. No macro storage, embedded object or remote template was found, but
the document text itself was not scanned for fields and links, and WPS JS macros in this binary
format cannot be checked. Converting it to .docx (sandboxed) lets me check those too."
