# JSON output

Every script prints one JSON object with `--json`, always with `ok`, `status`, `error`,
`tool`, `version`, `schema_version` (now "1") and `exit_code`. Any internal parser
error still produces JSON with `status` `error` and exit 2. `status` is `ok`, `findings`, `error` or `refused`
(exit 0, 1, 2, 3). Strings taken from the file are cleaned (no control, bidi or zero-width
characters) and capped.

## inspect_wps.py
| Key | Meaning |
| :--- | :--- |
| `file` | `name`, `size`, `sha256` |
| `container` | `cfb`, `zip`, `pdf`, `rtf`, `works-legacy` or `other` |
| `family` | `writer`, `spreadsheet`, `presentation`, `works`, `encrypted-ooxml`, `opendocument`, `pdf`, `unknown` |
| `format` | Human description of the layout |
| `extension`, `extension_matches` | Lower-case extension and whether it fits the content |
| `wps_markers` | `KSOProductBuildVer`, `KSOTemplate`, `WPS Office`, `Kingsoft`, `wps.cn XML namespace` when found |
| `producer` | Application name from the metadata (not proof of origin) |
| `metadata` | `author`, `last_saved_by`, `title`, `subject`, `keywords`, `description`/`comments`, `company`, `manager`, `template`, `application`, `app_version`, dates |
| `custom_properties` | Name to value (OOXML) |
| `streams` | OLE2 only: first 200 entries with `path`, `type`, `size` (`streams_cut` if more) |
| `active_content` | `{kind, where}`: `vba-project`, `wps-js-macro`, `activex`, `dde-field`, `ribbon-customization`, `pdf-actions` |
| `embedded_objects` | `{kind, path}`: `ole-object`, `ole-package`, `embedded-file` |
| `external_links` | `{type, scheme, host, part}` and `target` only with `--show-targets` |
| `hyperlinks` | Count of external hyperlinks (risky schemes also appear in `external_links`) |
| `wps_features` | `cell_images: {dispimg_cells, cellimages_part}`, `external_workbook_parts` |
| `hidden` | `sheets`, `rows`, `columns`, `hidden_slides`, `hidden_text_runs` (vanish, specVanish, webHidden; `w:val="0"` ignored) |
| `review` | `tracked_changes: {insertions, deletions}`, `comment_parts` |
| `assessment` | Fixed reminder: detected risks are not a safety verdict |
| `detected_risks` | `{risk, kind, category, detail}`, see `security.md`. `category`: `execution`, `fetch`, `deception`, `privacy`, `coverage`, `compatibility`. Renamed from `risks` in 0.1.1: old parsers must read the new key |
| `metadata_flags` | Metadata fields that held bidi or invisible characters |
| `complete` | The parser read every structure it tried (false: damaged or encrypted, see `incomplete`) |
| `coverage_complete`, `coverage_gaps` | False plus slugs (`binary-streams-not-parsed`, `jsa-in-binary-unknown`, `rtf-partial`, `pdf-object-streams`, `odf-content-not-parsed`, `works-not-parsed`) when some checks do not apply to this format; text in `not_checked` |
| `exit_reasons` | Why exit 1: `risk-threshold`, `share-blockers`, `incomplete` |
| `overall_risk` | `none`, `low`, `medium`, `high` |
| `complete`, `incomplete` | False plus reasons when structure was damaged or encrypted |
| `not_checked` | What this run could not check: say these to the user |
| `share_check` | With `--share-check`: `{passed, blockers, note}`; a privacy review, not a malware scan |

## cellimages.py
`cells` (`sheet`, `sheet_index`, `cell`, `image_id`, `media`), `pictures_in_part`, `missing`,
`saved` (`file`, `cells`, `shared_by`, `bytes`, `type`, `width`, `height`, `gps`): one file per
distinct picture, listing every cell that shows it; `skipped` (`cell`, `reason`); `notes`
(EXIF GPS, huge declared dimensions); `would_save` and `would_write_bytes` (dry run);
`output_dir`. `--max-total` (default 512 MB) caps all bytes written; pictures are streamed.

## convert_wps.py
`plan` (with `--dry-run`: `target`, `would_run`, `blockers`, `sandbox_available`,
`expected_losses`), `memory_limit_mb` (after clamping to `ulimit -v`), `engine_version`, `notes`,
`exit_reasons` (`input-risks`, `output-risks`), `input_family`, `input_risks`, `engine`,
`sandbox`, `output`, `output_check` (`family`, `overall_risk`, `detected_risks`, `complete`,
`macros_left`), `not_checked`.

## Exit codes
| Code | inspect_wps.py | cellimages.py | convert_wps.py |
| :--- | :--- | :--- | :--- |
| 0 | Report done, under thresholds | Listed or all saved | Converted, output clean |
| 1 | `--fail-on-risk` met, share blockers, incomplete | Missing pictures or skipped files | Input or output has medium+ risks (`exit_reasons`) |
| 2 | Not an office file, bad option | Not a ZIP spreadsheet, nothing to extract | Bad option, unsupported, LibreOffice missing or failed |
| 3 | Caps, ZIP rules, DOCTYPE (UTF-8/16/32), time, memory | Caps, `--max-total`, traversal, output exists | No sandbox, risky input without `--allow-risky-input`, caps, time, output exists, output not one regular file |

## Flags
| Script | Flag | Default | Effect |
| :--- | :--- | :--- | :--- |
| all | `--json` | off | One JSON object on stdout |
| all | `--max-bytes N` | 100 MB inspect, 50 MB convert | Refuse larger input (exit 3) |
| all | `--timeout S` | 60 inspect, 120 convert | Stop (exit 3); convert kills LibreOffice |
| inspect | `--fail-on-risk low/medium/high` | off | Exit 1 when a detected risk reaches the level |
| inspect | `--share-check` | off | Privacy blockers; exit 1 if any |
| inspect | `--allow-incomplete` | off | Do not exit 1 only because the report is incomplete |
| inspect | `--show-targets` | off | Print link targets as text (never fetched) |
| cellimages | `--extract DIR` / `--dry-run` | off | Save to a new folder / only list |
| cellimages | `--max-file`, `--max-total`, `--max-count` | 25 MB, 512 MB, 5000 | Per picture, all bytes, cells |
| convert | `--to FMT` | required | docx, xlsx, pptx, pdf, odt, ods, odp |
| convert | `--max-memory MB` | 4096 | RLIMIT_AS for LibreOffice |
| convert | `--force` | off | Replace OUT only |
| convert | `--allow-risky-input` | off | Convert medium/high-risk input (exit 1) |

Resource-limit hits (ZIP bomb ratio, member or total size, `--max-bytes`, `--timeout`, memory)
are always exit 3 `refused`, never exit 2. Exit 3 is not only "no sandbox".

## Canonical examples (real 0.1.2 output, trimmed)
Clean OOXML, exit 0:
```json
{"status": "ok", "family": "writer", "detected_risks": [], "overall_risk": "none",
 "complete": true, "coverage_complete": true, "coverage_gaps": [], "exit_reasons": [], "exit_code": 0}
```
Real WPS-saved `.ppt` (Tika sample), exit 0 with medium risk and a coverage gap:
```json
{"status": "ok", "family": "presentation", "wps_markers": ["KSOProductBuildVer"],
 "detected_risks": [{"risk": "medium", "kind": "embedded-objects", "category": "coverage",
   "detail": "PowerPoint stream has 9 ExOleObjStg record header(s): embedded OLE objects (byte scan; contents not parsed, may include OLE packages)"}],
 "coverage_complete": false, "coverage_gaps": ["binary-streams-not-parsed", "jsa-in-binary-unknown"], "exit_code": 0}
```
ZIP bomb, exit 3:
```json
{"ok": false, "status": "refused", "error": "ZIP member compression ratio over 200: xl/workbook.xml", "exit_code": 3}
```
Risky input to the converter without consent, exit 3:
```json
{"ok": false, "status": "refused", "exit_reasons": ["input-risks"],
 "error": "input has detected risks (ole-package); convert only if the user accepts that, with --allow-risky-input", "exit_code": 3}
```
