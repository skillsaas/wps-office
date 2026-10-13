# WPS and Microsoft Office: compatibility pitfalls

Load this when a file moves between WPS and Microsoft Office (or Apache POI, openpyxl and
similar libraries) and something is missing or refuses to open.

| Pitfall | What the user sees | Check | Fix | Source |
| :--- | :--- | :--- | :--- | :--- |
| Native extension | Word, Excel or PowerPoint will not open `.wps`, `.et`, `.dps` | `inspect_wps.py` family | Save As `.docx`/`.xlsx`/`.pptx` in WPS, or `convert_wps.py` | Look Scanned; file-extension.info |
| `.wps` is a Works file | Neither WPS nor Word opens it as expected | `family: works` | LibreOffice reads Works via libwps (`convert_wps.py --to docx`) | Look Scanned; wps2office notes |
| In-cell pictures | Excel shows `#NAME?` or `=_xlfn.DISPIMG("ID_...",1)`; pictures missing | `wps_features.cell_images`, `cellimages.py` | In WPS: right-click the picture, convert to a floating picture, save. Or save the pictures with `cellimages.py --extract` | hangge.com; shulou; exceljs issue 2786 |
| What `cellimages.py` recovers | Picture bytes and the cell-to-picture map; not the placement, size, crop, alt text, z-order or a working Excel picture | n/a | To get real in-cell or floating pictures in Excel, re-save in WPS after converting them to floating pictures | Tested on synthetic files only |
| Libraries ignore cell pictures | openpyxl, exceljs, pandas see a formula string, no picture | `cellimages.py FILE --json` | Map cell to picture through `xl/cellimages.xml` and its rels (what `cellimages.py` does) | eo2suite blog; yingdao post |
| WPS-saved `.doc` damaged | Apache POI `Block NN not found`; Outlook.com says corrupt; WPS opens it | `incomplete` has `sector-past-end-of-file` | Re-save in WPS (Save As `.doc` fixed it in the report) or convert | apache/poi issue 1041 |
| JS macros | Macros do nothing in Microsoft Office | `active_content` kind `wps-js-macro` | JS macros are WPS-only; rewrite as VBA or Office Scripts if needed (ask) | WPS forum topic 93172 |
| Fingerprint metadata | `KSOProductBuildVer` custom property stays after Office edits | `custom_properties` | Remove in a copy before sharing if it matters | jishuzhan.net |
| Fonts | Text reflows; boxes for Chinese text on Linux | fonts installed | Install the needed fonts (Linux: `references/linux.md`) | ArchWiki |
| Formula results | Values differ after conversion | recalculated by which engine? | Say which engine produced the value; compare a few cells | project LEARNINGS (engine differences) |

Rules:
- Always say which engine produced a converted file or a calculated value (WPS, Microsoft
  Office or LibreOffice).
- "Opens in WPS" is not "valid": other readers may reject the same file (POI case above).
- When a user must keep working in both suites, suggest saving in the Microsoft OOXML formats
  rather than the WPS native ones, and avoiding in-cell pictures and JS macros.
