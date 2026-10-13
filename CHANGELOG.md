# Changelog

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
