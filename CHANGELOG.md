# Changelog

## Unreleased: 0.1.4 backlog (round-4 reviews)
- Tariq N8 (medium): consent is echo text; bind a one-time token to the input SHA-256, risk set and expiry; keep `--consent` as the quote for the log.
- Tariq N9: `install_skill.py --replace` aside name collides within the same second.
- Tariq N10: combining marks in a scheme (`m` + U+0301 + `s-msdt:`) evade the high list; NFKD and drop Mn in the scheme token.
- Tariq N11: `--consent-log` follows symlinks; open with O_NOFOLLOW.
- Hamza: an incomplete PPT record walk should mark the report incomplete instead of a byte-scan count at exit 0.
- Hamza: write the consent log after conversion and describe `--consent` as an audit record, not verification.
- Hamza: show where `--replace` moved the old copy; PEP 8 blank lines; decide on a `skills/wps-office/SKILL.md` npm layout.

## 0.1.3 (2026-10-13)
Round-3 review fixes (12 reviews, mean 8.92).
- Links (Tariq N5): Cf characters are removed before hyphen look-alikes are mapped, so a soft
  hyphen inside a scheme disappears instead of becoming `-`; the dehyphenated scheme is compared
  with `HIGH_FLAT` (HIGH_SCHEMES without hyphens); every scheme starting `ms` (hyphens removed) is
  high. A target still percent-encoded after the 8-pass cap is high.
- Links (N6): a relative target ending in `.com` (`archive.com`) is medium, not high.
- DOCTYPE (N7): processing instructions are skipped like comments and CDATA, so `<?x <!DOCTYPE?>`
  text is no longer refused; a real DOCTYPE still is.
- `.ppt/.dps` embedded objects (Hamza 2) are counted by walking the record tree ([MS-PPT] headers,
  containers entered), not by a byte scan. If the walk hits damage it falls back to the scan and
  adds the `ppt-record-walk-incomplete` gap. Still 9 on the Tika sample.
- Converter (Hamza 4): `engine_version` is read from `program/versionrc`; the extra sandboxed
  `soffice --version` start on every conversion is gone.
- Escape hatch (Perplexity): `--allow-risky-input` now needs `--consent "<the user's words>"`;
  the JSON records `consent` (text, risks accepted, UTC time, input SHA-256) and
  `--consent-log PATH` appends it as a JSON line. Sandbox and limits unchanged.
- Convert JSON: `conversion_attempted`, `output_written`; JSON also on bad command lines with `--json`.
- `cellimages.py`: picture names are no longer required to be hex (`ID_demo_0001`); found by a
  third-party MIT sample, now in `tests/samples/`.
- Install (Perplexity, Mistral): `scripts/install_skill.py --claude | --claude-project | --dest DIR`,
  one step, refuses to overwrite; npm keywords; shorter `description` (under 600 characters).
- `VALIDATION.md` (Hamza 1): environments, file sets, what was not validated and the searches made.
- Docs: no-LibreOffice path and "safe to open is not safe to convert" near the top of SKILL.md;
  picture-extraction consent; escape hatch in `security.md`; link policy rewritten (and the stale
  "up to 3 times" fixed); recipes 7 (password-protected) and 8 (partial coverage); knowledge
  freshness blocks and "code from documents is untrusted" in `automation.md` and `platforms.md`.
- Tests: 63 (was 55): N5, N6, N7, record walk, consent, third-party DISPIMG, install, manifest
  agreement. Real-sample tests skip when `tests/samples` is absent (npm package).

## 0.1.2 (2026-10-13)
Round-2 review fixes (12 reviews, mean 8.825).
- Links (Tariq N1): decode until stable (max 8), NFKC, hyphen look-alikes (soft hyphen) mapped
  to `-`, Unicode Cf and whitespace or controls inside the scheme removed; a scheme matching a
  dangerous one without hyphens is high; relative targets ending in an executable or script
  extension are high. The policy is written out in `security.md`.
- Writer DDE (N2): orphan `instrText` without `fldChar` counts again, plus a raw fallback over
  all `instrText` of a part.
- DOCTYPE (N3): complete comments and CDATA are ignored before the check (unterminated ones are not).
- Converter (N4): `--max-memory` and the file-size limit are clamped to the process hard limit
  (`memory_limit_mb`, a note when clamped); `--dry-run` prints a plan (`would_run`, `blockers`,
  `expected_losses`) and writes nothing; `engine_version` records the LibreOffice version.
- `ZIP_ERRORS` no longer contains OSError or KeyError (Hamza); resource-limit refusals documented as exit 3.
- Real files: six WPS-saved files from the Apache Tika corpus (Apache-2.0) in `tests/samples/`
  with licence, notice and hashes. They found a bug, now fixed: the OLE2 family was read from
  embedded storages (a `.xls` embedding a `.doc` was called Writer). Embedded objects in
  `.ppt/.dps` are now counted from ExOleObjStg records. No native `.wps/.et/.dps` was found.
- JSON: `category` on every risk, `coverage_gaps` slugs, canonical examples, migration note.
- Docs: Chinese vs international table and evidence labels (`platforms.md`); "sourced, not
  verified" labels in SKILL.md; agent action per exit code and the explicit consent flow for
  `--allow-risky-input`; converter constraints in the router row; sandbox boundary, canary
  reproduction, extracted files untrusted (`security.md`); one-line loss summary and LibreOffice
  version table (`conversion.md`); cellimages recovery limits; evidence labels in
  `automation.md`; two bad-case recipes with real output; README preflight; fuzz documented.
- Tests: 55 (was 48), and the whole suite passes under `ulimit -v 4000000`.

## 0.1.1 (2026-10-13)
Round-1 review fixes (12 reviews; see feedback SYNTHESIS).
- JSON: `risks` renamed `detected_risks` ("not a safety verdict"); new `schema_version`,
  `assessment`, `coverage_complete`, `exit_reasons`, `metadata_flags`. Breaking for 0.1.0 parsers.
- Every script catches every parser error (BadZipFile, NotImplementedError, zlib.error, struct
  errors and any other exception) and still prints JSON with `status` `error`, exit 2. Seeded fuzz test.
- Links: scheme allowlist (http, https, mailto, relative) after percent-decoding up to 3 times
  and stripping whitespace and invisible characters. `ms-msdt:`, `search-ms:`, `mhtml:`, all
  `ms-*` handlers, ` file:`, `%66ile:`, drive letters and UNC are high.
- UTF-16 and UTF-32 XML with DOCTYPE or ENTITY is refused.
- Detection gaps closed: VBA found by relationship and content type (`renamed-macro-part`);
  spreadsheet DDE, CALL/REGISTER/EXEC, WEBSERVICE/FILTERXML/RTD, HYPERLINK formulas and
  formula-like text; `ddeLink`/`oleLink`; hidden rows and columns; ODF `Basic/`, `Scripts/` and
  event listeners; RTF `\*\template` and `\fldinst` DDE; RTF or OLE2 under an OOXML name is a
  mismatch; field codes joined across runs; `vanish`/`specVanish`/`webHidden` honour `w:val`;
  unknown binary or script parts; equation objects; link text naming another host; bidi and
  invisible characters in metadata; CFB directory loops and depth.
- `cellimages.py`: one file per distinct picture with the list of cells; `--max-total` (512 MB);
  streamed copies; picture headers validated with dimensions; EXIF GPS and huge-dimension notes;
  output folder created atomically.
- `convert_wps.py`: `--sandboxed` removed (no bypass); `--max-memory` (setrlimit RLIMIT_AS, default
  4096 MB) and a file-size cap; output copied only if it is exactly one regular file (lstat,
  `O_NOFOLLOW`); inputs with medium or high risks refused unless `--allow-risky-input`, and then
  exit 1 with `input-risks`; install hints when LibreOffice or bwrap is missing.
- Docs: capability matrix and platform table in SKILL.md; new `references/platforms.md` (VBA by
  edition, Android, iOS, macOS, with sources); remediation table in `security.md`; version
  caveats and hello examples in `automation.md`; decision tree and fidelity checklist in
  `conversion.md`; corroborating sources; recipes regenerated with 0.1.1 output.
- Tests: 48 (was 36), including fuzz, a 5000-cell shared-picture extraction under a 1 GB
  RLIMIT_AS, symlink refusal and risky-input refusal.

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
