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
| `detected_risks` | `{risk, kind, detail}`, see `security.md` (was `risks` in 0.1.0) |
| `metadata_flags` | Metadata fields that held bidi or invisible characters |
| `coverage_complete` | False when some checks do not apply to this format (binary, RTF, PDF); see `not_checked` |
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
