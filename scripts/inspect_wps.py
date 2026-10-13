#!/usr/bin/env python3
"""inspect_wps.py: read-only inventory of a WPS Office or Microsoft Office file.

Usage:
  python <skill-dir>/scripts/inspect_wps.py FILE [--json] [--share-check]
         [--fail-on-risk low|medium|high] [--allow-incomplete] [--show-targets]
         [--max-bytes N] [--timeout S] [--version]

Reads .wps .wpt .et .ett .dps .dpt (Kingsoft WPS Office), .doc .xls .ppt and
their templates, .docx .xlsx .pptx and macro-enabled variants, OpenDocument,
RTF and PDF, and tells a Kingsoft .wps from a Microsoft Works .wps. The type
comes from the bytes, never the extension.

It reports detected_risks: static indicators found by the checks that ran.
This is not a safety verdict and not a malware scan. Checks: WPS markers
(KSOProductBuildVer and friends), metadata (and hidden bidi or invisible
characters in it), macros (VBA found by name, relationship type or content
type; WPS JS macro parts; OpenDocument Basic and scripts), unknown binary or
script-like parts, OLE packages, ActiveX, DDE (Writer fields, also split across
runs; spreadsheet formulas and ddeLink/oleLink), WEBSERVICE/CALL/REGISTER
formulas, remote templates (OOXML and RTF), external links with one scheme
normalizer (allowlist http, https, mailto; ms-*, search-ms, mhtml, file, UNC
and others flagged after whitespace stripping and percent-decoding), hyperlink
text that shows another host, WPS cell images (DISPIMG), hidden sheets, rows,
columns, slides and text, comments, tracked changes, encryption, and damaged
compound-file structure.

It never runs, renders, converts or fetches anything, and never writes a file.
Standard library only, Python 3.9+.

--share-check: a privacy and content review before sending, not a malware scan:
lists active content, external links, embedded objects, metadata, custom
properties, comments, tracked changes and hidden content. Exit 1 if anything is
listed. This tool removes nothing.

Exit codes (also "status"; "exit_reasons" says which rule fired):
  0 ok         report done; nothing at or above --fail-on-risk; share-check clean
  1 findings   exit_reasons: risk-threshold, share-blockers, incomplete (damaged
               structure or encrypted, unless --allow-incomplete)
  2 error      not an office file, damaged beyond reading, bad option, or an
               internal error (still printed as JSON)
  3 refused    a safety cap: --max-bytes, --timeout, ZIP bomb caps, duplicate or
               case-colliding ZIP names, unsafe ZIP paths, DOCTYPE in XML
"""
import argparse
import hashlib
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wpslib  # noqa: E402
from wpslib import Refusal, FormatError, clean  # noqa: E402

FAMILY_EXT = {
    "writer": {".wps", ".wpt", ".doc", ".dot", ".docx", ".docm", ".dotx", ".dotm"},
    "spreadsheet": {".et", ".ett", ".xls", ".xlt", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam"},
    "presentation": {".dps", ".dpt", ".ppt", ".pot", ".pps", ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm"},
    "works": {".wps"},
    "pdf": {".pdf"},
}
OOXML_EXT = {".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam",
             ".pptx", ".pptm", ".ppsx", ".ppsm", ".potx", ".potm"}
BINARY_EXT = {".doc", ".dot", ".xls", ".xlt", ".ppt", ".pot", ".pps"}
MACRO_EXT = {".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".pptm", ".ppsm", ".potm"}
WPS_NATIVE = {".wps", ".wpt", ".et", ".ett", ".dps", ".dpt"}
RANK = {"none": 0, "low": 1, "medium": 2, "high": 3}
WPS_MARKERS = [b"KSOProductBuildVer", b"KSOTemplate", b"KSORubyTemplateID", b"WPS Office", b"Kingsoft"]
ASSESSMENT = "detected_risks are static indicators from the checks that ran; this is not a safety verdict or a malware scan"

NOT_CHECKED_ALWAYS = [
    "nothing was rendered or run: layout, fonts and formula results are not checked",
    "links are listed as text only; no URL or path is fetched",
    "macro code is detected by its storage location and type, not read or judged",
    "no malware scan: embedded files and pictures are listed, never scanned",
]


CATEGORY = {
    "execution": ("macros", "renamed-macro-part", "macros-under-plain-extension", "ole-package", "activex",
                  "dde-field", "dde-formula", "call-formula", "dde-link", "ole-link", "odf-event-script",
                  "rtf-object", "pdf-active-content", "equation-editor-object", "unknown-binary-or-script-part"),
    "fetch": ("external-", "dangerous-link-scheme", "unusual-link-scheme", "web-formula", "hyperlink-formula"),
    "deception": ("extension-mismatch", "text-shows-other-host", "hidden-characters-in-metadata", "very-hidden-sheet"),
    "privacy": ("formula-like-text",),
    "coverage": ("encrypted", "damaged-structure", "embedded-objects"),
    "compatibility": ("wps-cell-images",),
}


def category(kind):
    """What kind of concern a risk is: execution, fetch, deception, privacy, coverage, compatibility."""
    for cat, kinds in CATEGORY.items():
        if any(kind == k or (k.endswith("-") and kind.startswith(k)) for k in kinds):
            return cat
    return "other"


class Report:
    def __init__(self, path, show_targets):
        self.show = show_targets
        self.r = {"ok": True, "status": "ok", "error": None, "tool": "inspect_wps.py",
                  "version": wpslib.VERSION, "schema_version": wpslib.SCHEMA_VERSION,
                  "assessment": ASSESSMENT, "file": {}, "container": None,
                  "family": "unknown", "format": None, "extension": None,
                  "extension_matches": None, "wps_markers": [], "producer": None,
                  "metadata": {}, "metadata_flags": {}, "custom_properties": {}, "active_content": [],
                  "embedded_objects": [], "external_links": [], "hyperlinks": 0,
                  "wps_features": {}, "hidden": {}, "review": {},
                  "detected_risks": [], "overall_risk": "none", "complete": True,
                  "coverage_complete": True, "coverage_gaps": [], "incomplete": [], "not_checked": list(NOT_CHECKED_ALWAYS),
                  "exit_reasons": []}

    def risk(self, level, kind, detail):
        self.r["detected_risks"].append({"risk": level, "kind": kind, "category": category(kind),
                                         "detail": clean(detail, 300)})

    def has(self, kind):
        return any(x["kind"] == kind for x in self.r["detected_risks"])

    def incomplete(self, why):
        if why not in self.r["incomplete"]:
            self.r["incomplete"].append(why)
        self.r["complete"] = False

    def gap(self, what, slug):
        self.r["not_checked"].append(what)
        if slug not in self.r["coverage_gaps"]:
            self.r["coverage_gaps"].append(slug)
        self.r["coverage_complete"] = False

    def meta(self, key, value):
        if isinstance(value, str) and value.strip():
            hc = wpslib.hidden_chars(value)
            if hc:
                self.r["metadata_flags"][key] = hc
                level = "medium" if "bidi" in hc else "low"
                self.risk(level, "hidden-characters-in-metadata",
                          "%s contains %s characters (removed when printed)" % (key, " and ".join(hc)))
            self.r["metadata"][key] = clean(value)

    def link(self, rtype, target, part):
        norm, scheme, host, risk = wpslib.link_target(target)
        item = {"type": rtype, "scheme": scheme, "host": clean(host, 120), "part": clean(part, 120)}
        if self.show:
            item["target"] = clean(norm, 500)
        self.r["external_links"].append(item)
        return item, risk


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def scan_markers(rep, blobs):
    found = set(rep.r["wps_markers"])
    for b in blobs:
        for m in WPS_MARKERS:
            if m in b or m.decode().encode("utf-16-le") in b:
                found.add(m.decode())
    rep.r["wps_markers"] = sorted(found)


# ------------------------------------------------------------------ CFB files
SUMMARY_PIDS = {2: "title", 3: "subject", 4: "author", 5: "keywords", 6: "comments",
                7: "template", 8: "last_saved_by", 18: "application"}
DOCSUM_PIDS = {14: "manager", 15: "company"}


def inspect_cfb(rep, data, dl, cap):
    cfb = wpslib.Cfb(data, dl)
    names = {e["path"].lower() for e in cfb.entries if "/" not in e["path"]}   # root level decides the family
    paths = [e["path"] for e in cfb.entries]
    r = rep.r
    if "encryptioninfo" in names and "encryptedpackage" in names:
        r["family"], r["format"] = "encrypted-ooxml", "password-protected OOXML package (EncryptionInfo + EncryptedPackage)"
        rep.risk("medium", "encrypted", "password-protected: contents cannot be inspected without the password")
        rep.incomplete("encrypted")
    elif "worddocument" in names:
        r["family"], r["format"] = "writer", "Word 97-2003 binary layout (.doc; WPS Writer .wps/.wpt use the same layout)"
        fib = cfb.read([e for e in cfb.entries if e["path"].lower() == "worddocument"][0], 64)
        if len(fib) >= 12:
            ident, flags = int.from_bytes(fib[0:2], "little"), int.from_bytes(fib[10:12], "little")
            if ident != 0xA5EC:
                rep.incomplete("WordDocument stream has no Word FIB signature")
            if flags & 0x0100:
                rep.risk("medium", "encrypted", "document is password-encrypted (FIB fEncrypted)")
                rep.incomplete("encrypted")
    elif "workbook" in names or "book" in names:
        r["family"], r["format"] = "spreadsheet", "Excel 97-2003 BIFF layout (.xls; WPS Spreadsheets .et/.ett use the same layout)"
        ent = [e for e in cfb.entries if e["path"].lower() in ("workbook", "book")][0]
        wb = cfb.read(ent, 256 * 1024)
        pos = 0
        while pos + 4 <= len(wb):
            rt, ln = int.from_bytes(wb[pos:pos + 2], "little"), int.from_bytes(wb[pos + 2:pos + 4], "little")
            if rt == 0x002F:
                rep.risk("medium", "encrypted", "workbook is password-encrypted (BIFF FILEPASS)")
                rep.incomplete("encrypted")
                break
            if rt == 0x000A:
                break
            pos += 4 + ln
    elif "powerpoint document" in names:
        r["family"], r["format"] = "presentation", "PowerPoint 97-2003 binary layout (.ppt; WPS Presentation .dps/.dpt use the same layout)"
        if "encryptedsummary" in names:
            rep.risk("medium", "encrypted", "presentation is password-encrypted (EncryptedSummary stream)")
            rep.incomplete("encrypted")
        # Embedded objects in .ppt/.dps live inside the PowerPoint Document stream as
        # ExOleObjStg records (type 0x1011), not as storages: count record headers.
        pd = cfb.read([e for e in cfb.entries if e["path"].lower() == "powerpoint document"][0], min(cap, 64 * 1024 * 1024))
        n_obj = len(re.findall(rb"[\x00\x10]\x00\x11\x10", pd))
        if n_obj:
            r["embedded_objects"].append({"kind": "ppt-exoleobjstg", "path": "PowerPoint Document (%d record header(s))" % n_obj})
            rep.risk("medium", "embedded-objects", "PowerPoint stream has %d ExOleObjStg record header(s): embedded OLE objects (byte scan; contents not parsed, may include OLE packages)" % n_obj)
    elif "mn0" in names:
        r["family"], r["format"] = "works", "Microsoft Works word processor (.wps, MN0 stream); not a Kingsoft WPS file"
    else:
        r["family"], r["format"] = "unknown", "OLE2 compound file of unknown type"
    r["streams"] = [{"path": clean(e["path"], 120), "type": e["type"], "size": e["size"]} for e in cfb.entries[:200]]
    if len(cfb.entries) > 200:
        r["streams_cut"] = len(cfb.entries) - 200

    blobs = []
    for nm, pids, key in (("\x05SummaryInformation", SUMMARY_PIDS, "metadata"),
                          ("\x05DocumentSummaryInformation", DOCSUM_PIDS, "metadata")):
        for ent in cfb.find(nm):
            blob = cfb.read(ent, 1024 * 1024)
            blobs.append(blob)
            props = wpslib.read_summary(blob)
            for pid, field in pids.items():
                v = props.get(pid)
                if isinstance(v, str):
                    rep.meta(field, v)
    for ent in cfb.entries:
        if ent["type"] == "stream" and ent["path"].split("/")[-1] in ("WordDocument", "Workbook", "Book", "PowerPoint Document", "1Table", "0Table"):
            blobs.append(cfb.read(ent, min(cap, 4 * 1024 * 1024)))
    scan_markers(rep, blobs)
    r["producer"] = r["metadata"].get("application")

    low = [p.lower() for p in paths]
    if any(p.split("/")[-1] in ("_vba_project_cur", "vba", "_vba_project", "macros") for p in low) or \
            any(p.endswith("/project") or p == "project" for p in low):
        r["active_content"].append({"kind": "vba-project", "where": "compound file storage (Macros, _VBA_PROJECT_CUR or VBA)"})
        rep.risk("high", "macros", "VBA project present: code that runs when macros are enabled")
    for p in paths:
        last = p.split("/")[-1]
        if last.lower() == "\x01ole10native":
            r["embedded_objects"].append({"kind": "ole-package", "path": clean(p, 120)})
            rep.risk("high", "ole-package", "OLE Package object (Ole10Native) can carry any file, including programs: %s" % p)
        elif p.lower().startswith("objectpool/") and p.count("/") == 1 or re.match(r"^MBD[0-9A-F]+$", last):
            r["embedded_objects"].append({"kind": "ole-object", "path": clean(p, 120)})
    if r["embedded_objects"] and not rep.has("embedded-objects") and not any(x["risk"] == "high" and x["kind"] == "ole-package" for x in r["detected_risks"]):
        rep.risk("medium", "embedded-objects", "%d embedded OLE object(s)" % len(r["embedded_objects"]))
    if cfb.issues:
        for i in cfb.issues:
            rep.incomplete("compound file structure: %s" % i)
        if "sector-past-end-of-file" in cfb.issues:
            rep.risk("low", "damaged-structure", "sector chain points past the end of the file (seen in some WPS-saved .doc files; other readers may refuse it)")
    rep.gap("binary formats: text, comments, tracked changes, hidden content, links and DDE inside the binary streams are not parsed", "binary-streams-not-parsed")
    rep.gap("binary formats: where WPS stores JS (JSA) macros inside .wps/.et/.dps is not documented; only VBA storages are detected", "jsa-in-binary-unknown")


# ------------------------------------------------------------------ OOXML files
REL_RISK = {"attachedTemplate": "high", "oleObject": "medium", "frame": "medium",
            "externalLinkPath": "medium", "image": "medium", "subDocument": "medium",
            "package": "medium", "video": "medium", "audio": "medium", "media": "medium"}
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
OFF = ("0", "false", "off")
KNOWN_BIN = re.compile(r"(^|/)(printersettings\d*\.bin|.*/media/.*|.*/embeddings/.*|vbaproject\.bin|vbadata\.xml|"
                       r"activex\d*\.bin|jdedata.*|jsaproject.*|.*fonts/.*\.(odttf|bin))$")
SCRIPTISH = re.compile(rb"(\bfunction\s+\w+\s*\(|\bSub\s+(Auto_?Open|Document_Open|Workbook_Open)\b|\beval\s*\(|ActiveXObject|WScript\.Shell)", re.I)
DDE_FORMULA = re.compile(r"(^|[=+\-*/(,&\s])'?[A-Za-z0-9_.\-]+'?\|'?[^!|]*'?!")


def on(el):
    v = el.get(W + "val")
    return v is None or v.lower() not in OFF


def inspect_zip(rep, data, dl, cap):
    zf = wpslib.open_zip(data)
    names = zf.namelist()
    lower = {n.lower(): n for n in names}
    r = rep.r

    def get(n, limit=32 * 1024 * 1024):
        dl.check()
        return wpslib.read_member(zf, n, min(cap, limit))

    def xml(n):
        return wpslib.parse_xml(get(n))

    if "[content_types].xml" not in lower:
        if "mimetype" in lower:
            return inspect_odf(rep, zf, names, lower, get)
        raise FormatError("ZIP file is not an Office Open XML or OpenDocument package")
    ctroot = xml(lower["[content_types].xml"])
    ct_by_part, ct_by_ext = {}, {}
    for el in ctroot:
        if wpslib.local(el.tag) == "Override":
            ct_by_part[(el.get("PartName") or "").lstrip("/").lower()] = (el.get("ContentType") or "").lower()
        elif wpslib.local(el.tag) == "Default":
            ct_by_ext[(el.get("Extension") or "").lower()] = (el.get("ContentType") or "").lower()
    ct_all = " ".join(list(ct_by_part.values()))

    def ctype(n):
        ln = n.lower()
        return ct_by_part.get(ln) or ct_by_ext.get(ln.rsplit(".", 1)[-1], "")

    if "word/document.xml" in lower or "wordprocessingml.document" in ct_all:
        r["family"] = "writer"
    elif "xl/workbook.xml" in lower or "spreadsheetml.sheet" in ct_all:
        r["family"] = "spreadsheet"
    elif "ppt/presentation.xml" in lower or "presentationml" in ct_all:
        r["family"] = "presentation"
    r["format"] = "Office Open XML package" + (" (macro-enabled content type)" if "macroenabled" in ct_all else "")

    # metadata
    if "docprops/core.xml" in lower:
        for el in xml(lower["docprops/core.xml"]):
            k, v = wpslib.local(el.tag), (el.text or "")
            key = {"creator": "author", "lastModifiedBy": "last_saved_by"}.get(k, k)
            if key in ("author", "last_saved_by", "title", "subject", "keywords", "description", "category", "created", "modified"):
                rep.meta(key, v)
    if "docprops/app.xml" in lower:
        for el in xml(lower["docprops/app.xml"]):
            k, v = wpslib.local(el.tag), (el.text or "")
            if k in ("Application", "AppVersion", "Company", "Manager", "Template"):
                rep.meta(k.lower() if k != "AppVersion" else "app_version", v)
        r["producer"] = r["metadata"].get("application")
    if "docprops/custom.xml" in lower:
        for el in xml(lower["docprops/custom.xml"]):
            nm = el.get("name")
            val = "".join(c.text or "" for c in el)
            if nm:
                r["custom_properties"][clean(nm, 80)] = clean(val, 120)
    blobs = [k.encode() + b"=" + v.encode() for k, v in r["custom_properties"].items()]
    blobs.append((r["metadata"].get("application") or "").encode())
    scan_markers(rep, blobs)
    for n in names:
        ln = n.lower()
        if ln.endswith(".xml") and ln.startswith(("xl/workbook", "word/document", "ppt/presentation", "xl/cellimages")):
            if b"wps.cn" in get(n):
                r["wps_markers"] = sorted(set(r["wps_markers"]) | {"wps.cn XML namespace"})

    # relationships: collect once, used for macros and links
    rels = []
    for n in names:
        if n.lower().endswith(".rels"):
            for rel in xml(n):
                rels.append((n, rel.get("Id"), (rel.get("Type") or "").rsplit("/", 1)[-1],
                             rel.get("Target") or "", rel.get("TargetMode", "")))
    vba_targets = set()
    for part, rid, rtype, target, mode in rels:
        if rtype in ("vbaProject", "wordVbaData") and mode != "External":
            base = part.rsplit("_rels/", 1)[0]
            vba_targets.add(os.path.normpath(os.path.join(base, target)).replace("\\", "/").lstrip("/").lower())

    # active content, by name, relationship type and content type
    for n in names:
        ln = n.lower()
        base = ln.rsplit("/", 1)[-1]
        ct = ctype(n)
        if base == "vbaproject.bin" or ln in vba_targets or "vbaproject" in ct:
            r["active_content"].append({"kind": "vba-project", "where": clean(n, 120)})
            how = "name" if base == "vbaproject.bin" else ("relationship type" if ln in vba_targets else "content type")
            rep.risk("high", "macros", "VBA project (%s, found by %s): code that runs when macros are enabled" % (n, how))
            if base != "vbaproject.bin":
                rep.risk("high", "renamed-macro-part", "macro part hidden under the name %s" % n)
        elif base.startswith("jdedata") or "jsaproject" in base or "jsa" in ct:
            r["active_content"].append({"kind": "wps-js-macro", "where": clean(n, 120)})
            rep.risk("high", "macros", "WPS JS (JSA) macro part (%s): code that runs in WPS when macros are enabled" % n)
        elif ("/activex/" in ln and base.endswith((".xml", ".bin"))) or "activex" in ct:
            if not any(a["kind"] == "activex" for a in r["active_content"]):
                r["active_content"].append({"kind": "activex", "where": clean(n.rsplit("/", 1)[0], 120)})
                rep.risk("high", "activex", "ActiveX controls present")
        elif ln.startswith(("customui/", "customui14")) or base in ("customui.xml", "customui14.xml"):
            r["active_content"].append({"kind": "ribbon-customization", "where": clean(n, 120)})
        elif "/embeddings/" in ln and not ln.endswith("/"):
            kind = "ole-object" if base.endswith(".bin") else "embedded-file"
            r["embedded_objects"].append({"kind": kind, "path": clean(n, 120)})
            if base.endswith(".bin"):
                head = get(n, 8 * 1024 * 1024)
                if head.startswith(wpslib.CFB_MAGIC):
                    try:
                        sub = wpslib.Cfb(head, dl)
                        if sub.find("\x01Ole10Native"):
                            r["embedded_objects"][-1]["kind"] = "ole-package"
                            rep.risk("high", "ole-package", "OLE Package object (%s) can carry any file, including programs" % n)
                        if any("equation" in e["path"].lower() for e in sub.entries):
                            rep.risk("medium", "equation-editor-object", "legacy Equation object (%s): old Equation Editor exploits used this carrier" % n)
                    except (FormatError, Refusal):
                        rep.incomplete("embedded object %s could not be read" % clean(n, 80))
        elif not ln.endswith((".xml", ".rels", "/")) and not KNOWN_BIN.search(ln):
            head = get(n, 1024 * 1024)
            if wpslib.sniff(head[:16]) in ("cfb", "zip") or SCRIPTISH.search(head):
                r["active_content"].append({"kind": "unknown-part", "where": clean(n, 120)})
                rep.risk("medium", "unknown-binary-or-script-part",
                         "part %s is not a known Office part and holds %s" % (
                             n, "script-like text" if SCRIPTISH.search(head) else "a container"))
    if r["embedded_objects"] and not rep.has("ole-package"):
        rep.risk("medium", "embedded-objects", "%d embedded object(s) in */embeddings/" % len(r["embedded_objects"]))
    ext = r.get("extension") or ""
    if any(a["kind"] in ("vba-project", "wps-js-macro") for a in r["active_content"]) \
            and ext and ext not in MACRO_EXT and ext not in WPS_NATIVE:
        rep.risk("high", "macros-under-plain-extension", "macro code inside a file named %s (expected .docm/.xlsm/.pptm)" % ext)

    # external relationships
    rel_by_id = {}
    for part, rid, rtype, target, mode in rels:
        rel_by_id[(part, rid)] = (rtype, target, mode)
        if mode != "External":
            continue
        if rtype == "hyperlink":
            r["hyperlinks"] += 1
            _norm, scheme, _h, risk = wpslib.link_target(target)
            if risk != "none":
                item, _ = rep.link(rtype, target, part)
                rep.risk(risk, "dangerous-link-scheme" if risk == "high" else "unusual-link-scheme",
                         "hyperlink with %s: target" % item["scheme"])
            continue
        item, srisk = rep.link(rtype, target, part)
        level = REL_RISK.get(rtype, "medium")
        if srisk == "high":
            level = "high"
        why = {"attachedTemplate": "remote template is fetched when the document opens (template injection)",
               "image": "linked picture is fetched when the document is opened or converted"}.get(rtype, "external %s link" % rtype)
        rep.risk(level, "external-" + rtype, "%s; host %s, scheme %s" % (why, item["host"] or "-", item["scheme"]))

    if r["family"] == "writer":
        writer_checks(rep, names, lower, xml, rel_by_id)
    com = [n for n in names if re.search(r"(^|/)(comments\d*\.xml|threadedcomment\w*\.xml|commentsextended\.xml)$", n.lower())
           or n.lower().startswith("ppt/comments/")]
    if com:
        r["review"]["comment_parts"] = [clean(c, 80) for c in com[:50]]
    if r["family"] == "spreadsheet" and "xl/workbook.xml" in lower:
        sheet_checks(rep, names, lower, xml, get)
    if r["family"] == "presentation":
        hs = 0
        for n in names:
            if re.match(r"^ppt/slides/slide\d+\.xml$", n.lower()):
                if re.search(rb"<p:sld\b[^>]*\sshow=\"(0|false)\"", get(n)):
                    hs += 1
        if hs:
            r["hidden"]["hidden_slides"] = hs


def writer_checks(rep, names, lower, xml, rel_by_id):
    r = rep.r
    parts = [n for n in names if re.match(r"^word/(document|header\d*|footer\d*|footnotes|endnotes|comments)\.xml$", n.lower())]
    ins = dels = van = 0
    fields = []
    shown_other = 0
    for n in parts:
        root = xml(n)
        stack = []
        for el in root.iter():
            t = el.tag
            if t == W + "fldChar":
                kind = el.get(W + "fldCharType")
                if kind == "begin":
                    stack.append([[], True])
                elif kind == "separate" and stack:
                    stack[-1][1] = False
                elif kind == "end" and stack:
                    fields.append("".join(stack.pop()[0]))
            elif t == W + "instrText" and stack and stack[-1][1]:
                stack[-1][0].append(el.text or "")
            elif t == W + "instrText" and not stack:
                fields.append(el.text or "")       # orphan instrText (no fldChar) still counts
            elif t == W + "fldSimple":
                fields.append(el.get(W + "instr") or "")
            elif t in (W + "ins", W + "moveTo"):
                ins += 1
            elif t in (W + "del", W + "moveFrom"):
                dels += 1
            elif t in (W + "vanish", W + "specVanish", W + "webHidden") and on(el):
                van += 1
            elif t == W + "hyperlink" and el.get(R_NS + "id"):
                rel = rel_by_id.get(("word/_rels/%s.rels" % n.split("/", 1)[1], el.get(R_NS + "id")))
                text = "".join(x.text or "" for x in el.iter(W + "t")).strip()
                if rel and text:
                    m = re.search(r"(?:https?://)?((?:[a-z0-9-]+\.)+[a-z]{2,})", text.lower())
                    _n, _s, host, _r = wpslib.link_target(rel[1])
                    if m and host and m.group(1) != host.lower().split(":")[0] and \
                            not host.lower().endswith("." + m.group(1)):
                        shown_other += 1
        fields.extend("".join(st[0]) for st in stack)  # unterminated fields still count
        # raw fallback: every instrText of the part joined, whatever the field structure
        allinstr = "".join((x.text or "") for x in root.iter(W + "instrText"))
        if re.search(r"(^|\s)DDE(AUTO)?\s", allinstr, re.I):
            fields.append("DDE " + allinstr[:80])
    for f in fields:
        fs = " ".join(f.split())
        if re.match(r"^\s*DDE(AUTO)?\b", fs, re.I):
            if not any(a["kind"] == "dde-field" for a in r["active_content"]):
                rep.risk("high", "dde-field", "DDE or DDEAUTO field: can start a program when fields update")
                r["active_content"].append({"kind": "dde-field", "where": "Writer field"})
        elif re.match(r"^\s*(INCLUDETEXT|INCLUDEPICTURE|LINK)\b", fs, re.I):
            _n, scheme, host, risk = wpslib.link_target((re.findall(r'"([^"]+)"', fs) or fs.split()[1:2] or [""])[0])
            if scheme not in ("relative",) or "\\\\" in fs:
                if not rep.has("external-field"):
                    rep.risk("medium" if risk != "high" else "high", "external-field",
                             "INCLUDETEXT/INCLUDEPICTURE/LINK field with a %s path" % scheme)
    if shown_other:
        rep.risk("medium", "text-shows-other-host", "%d hyperlink(s) whose visible text names a different host" % shown_other)
    if ins or dels:
        r["review"]["tracked_changes"] = {"insertions": ins, "deletions": dels}
    if van:
        r["hidden"]["hidden_text_runs"] = van


def sheet_checks(rep, names, lower, xml, get):
    r = rep.r
    wb = xml(lower["xl/workbook.xml"])
    hid = []
    for el in wb.iter():
        if wpslib.local(el.tag) == "sheet" and el.get("state") in ("hidden", "veryHidden"):
            hid.append({"name": clean(el.get("name"), 80), "state": el.get("state")})
    if hid:
        r["hidden"]["sheets"] = hid
        if any(h["state"] == "veryHidden" for h in hid):
            rep.risk("medium", "very-hidden-sheet", "sheet hidden so it does not appear in the Unhide list")
    xl = [n for n in names if n.lower().startswith("xl/externallinks/") and n.lower().endswith(".xml")]
    if xl:
        r["wps_features"]["external_workbook_parts"] = len(xl)
        for n in xl:
            for el in xml(n).iter():
                k = wpslib.local(el.tag)
                if k == "ddeLink":
                    rep.risk("high", "dde-link", "DDE link to %s (external link part %s)" % (clean(el.get("ddeService"), 40), n))
                elif k == "oleLink":
                    rep.risk("high", "ole-link", "OLE link in external link part %s" % n)
    counts = {"dde": 0, "web": 0, "call": 0, "hyperlink": 0, "formula_like_text": 0}
    hidden_rows = hidden_cols = dispimg = 0
    for n in names:
        ln = n.lower()
        if ln.startswith("xl/worksheets/") and ln.endswith(".xml"):
            root = xml(n)
            for el in root.iter():
                k = wpslib.local(el.tag)
                if k == "row" and el.get("hidden") in ("1", "true"):
                    hidden_rows += 1
                elif k == "col" and el.get("hidden") in ("1", "true"):
                    hidden_cols += 1
                elif k == "c":
                    f = "".join((x.text or "") for x in el if wpslib.local(x.tag) == "f")
                    if "DISPIMG(" in f or any("DISPIMG(" in (x.text or "") for x in el if wpslib.local(x.tag) == "v"):
                        dispimg += 1
                    if f:
                        fu = f.upper()
                        if DDE_FORMULA.search(f):
                            counts["dde"] += 1
                        if re.search(r"\b(WEBSERVICE|FILTERXML|RTD)\s*\(", fu):
                            counts["web"] += 1
                        if re.search(r"\b(CALL|REGISTER|REGISTER\.ID|EXEC)\s*\(", fu):
                            counts["call"] += 1
                        if re.search(r"\bHYPERLINK\s*\(", fu):
                            counts["hyperlink"] += 1
                    if el.get("t") == "inlineStr":
                        txt = "".join(x.text or "" for x in el.iter() if wpslib.local(x.tag) == "t")
                        if txt[:1] in ("=", "+", "-", "@"):
                            counts["formula_like_text"] += 1
    if "xl/sharedstrings.xml" in lower:
        for si in xml(lower["xl/sharedstrings.xml"]):
            txt = "".join(x.text or "" for x in si.iter() if wpslib.local(x.tag) == "t")
            if txt[:1] in ("=", "+", "-", "@") and len(txt) > 1:
                counts["formula_like_text"] += 1
    if counts["dde"]:
        rep.risk("high", "dde-formula", "%d formula(s) with a DDE reference (app|topic!item, for example cmd|' /c ...'!A1)" % counts["dde"])
    if counts["call"]:
        rep.risk("high", "call-formula", "%d formula(s) using CALL/REGISTER/EXEC" % counts["call"])
    if counts["web"]:
        rep.risk("medium", "web-formula", "%d formula(s) using WEBSERVICE/FILTERXML/RTD (fetch or link out on recalculation)" % counts["web"])
    if counts["hyperlink"]:
        rep.risk("low", "hyperlink-formula", "%d HYPERLINK() formula(s); targets not resolved" % counts["hyperlink"])
    if counts["formula_like_text"]:
        rep.risk("low", "formula-like-text", "%d text cell(s) starting with = + - @ (formula injection if exported to CSV)" % counts["formula_like_text"])
    if hidden_rows or hidden_cols:
        r["hidden"]["rows"] = hidden_rows
        r["hidden"]["columns"] = hidden_cols
    has_ci = "xl/cellimages.xml" in lower
    if dispimg or has_ci:
        r["wps_features"]["cell_images"] = {"dispimg_cells": dispimg, "cellimages_part": has_ci}
        rep.risk("low", "wps-cell-images",
                 "%d cell(s) with a DISPIMG picture formula: WPS-only; Microsoft Excel shows #NAME? and no picture. "
                 "List or save them with cellimages.py" % dispimg)


# ------------------------------------------------------------------ OpenDocument
def inspect_odf(rep, zf, names, lower, get):
    r = rep.r
    mt = get(lower["mimetype"], 256).decode("ascii", "replace").strip()
    r["family"], r["format"] = "opendocument", "OpenDocument (%s)" % clean(mt, 80)
    for n in names:
        ln = n.lower()
        if ln.startswith(("basic/", "scripts/")) and not ln.endswith("/"):
            if not any(a["kind"] == "odf-macro" for a in r["active_content"]):
                r["active_content"].append({"kind": "odf-macro", "where": clean(n.split("/")[0] + "/", 40)})
                rep.risk("high", "macros", "OpenDocument macro library (%s)" % n.split("/")[0])
    for n in ("content.xml", "styles.xml", "settings.xml"):
        if n in lower:
            b = get(lower[n])
            if re.search(rb"(script:event-listener|office:event-listeners|xlink:href=\"(vnd\.sun\.star\.script|macro):)", b):
                r["active_content"].append({"kind": "odf-event-script", "where": n})
                rep.risk("high", "odf-event-script", "event listener or script link in %s (runs on open or click)" % n)
    if "meta-inf/manifest.xml" in lower and b"manifest:encryption-data" in get(lower["meta-inf/manifest.xml"]):
        rep.risk("medium", "encrypted", "OpenDocument parts are encrypted")
        rep.incomplete("encrypted")
    rep.gap("OpenDocument: metadata, links, comments and hidden content are not parsed; only macros and event scripts are checked", "odf-content-not-parsed")


def share_blockers(r):
    b = []
    for a in r["active_content"]:
        b.append("active content: %s (%s)" % (a["kind"], a["where"]))
    for e in r["embedded_objects"]:
        b.append("embedded object: %s" % e["path"])
    for x in r["external_links"]:
        b.append("external link: %s to %s" % (x["type"], x["host"] or x["scheme"]))
    for x in r["detected_risks"]:
        if x["kind"] in ("dde-formula", "call-formula", "web-formula", "dde-link", "ole-link", "external-field", "text-shows-other-host"):
            b.append("risk: %s" % x["kind"])
    for k in ("author", "last_saved_by", "company", "manager", "template", "comments", "keywords"):
        if r["metadata"].get(k):
            b.append("metadata: %s" % k)
    for k in r["custom_properties"]:
        b.append("custom property: %s" % k)
    if r["review"].get("tracked_changes"):
        b.append("tracked changes")
    if r["review"].get("comment_parts"):
        b.append("comments")
    for k, v in r["hidden"].items():
        if v:
            b.append("hidden content: %s" % k)
    if not r["complete"]:
        b.append("report incomplete: %s" % "; ".join(r["incomplete"]))
    if not r["coverage_complete"]:
        b.append("coverage incomplete: some checks do not apply to this format (see not_checked); convert to OOXML and check again")
    return b


def ext_matches(r, ext):
    fam, cont = r["family"], r["container"]
    if fam == "encrypted-ooxml":
        return ext in OOXML_EXT or ext in WPS_NATIVE
    if fam == "opendocument":
        return ext.startswith((".od", ".ot"))
    if cont == "rtf":
        return ext in (".rtf",)
    if cont == "cfb" and ext in OOXML_EXT:
        return False
    if cont == "zip" and ext in BINARY_EXT:
        return False
    return ext in FAMILY_EXT.get(fam, set())


def inspect_rtf(rep, data):
    r = rep.r
    r["family"], r["format"] = "writer", "RTF"
    if re.search(rb"\\objdata|\\object\b", data):
        rep.risk("high", "rtf-object", "RTF embeds an OLE object (\\objdata): a common exploit carrier")
        r["active_content"].append({"kind": "rtf-object", "where": "\\objdata"})
    m = re.search(rb"\{\\\*\\template\s+([^}]*)\}", data)
    if m:
        _n, scheme, host, risk = wpslib.link_target(m.group(1).decode("latin-1").strip())
        item, _ = rep.link("attachedTemplate", m.group(1).decode("latin-1").strip(), "\\*\\template")
        rep.risk("high", "external-attachedTemplate", "RTF \\*\\template points to %s (%s): fetched when opened" % (host or "-", scheme))
    for fm in re.finditer(rb"\\fldinst\s*(\{[^}]*\}|[^}\\]+)", data[:8 * 1024 * 1024]):
        txt = re.sub(rb"[{}]|\\[a-z]+-?\d* ?", b"", fm.group(1)).decode("latin-1")
        if re.match(r"^\s*DDE(AUTO)?\b", txt, re.I):
            rep.risk("high", "dde-field", "RTF DDE/DDEAUTO field")
            r["active_content"].append({"kind": "dde-field", "where": "RTF \\fldinst"})
            break
    rep.gap("RTF: only \\objdata, \\*\\template and DDE fields are checked", "rtf-partial")


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="Read-only inventory of WPS Office and Microsoft Office files.")
    ap.add_argument("file")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--share-check", action="store_true")
    ap.add_argument("--fail-on-risk", choices=["low", "medium", "high"])
    ap.add_argument("--allow-incomplete", action="store_true")
    ap.add_argument("--show-targets", action="store_true", help="print full link targets as text (never fetched; may be sensitive)")
    ap.add_argument("--max-bytes", type=int, default=100 * 1024 * 1024)
    ap.add_argument("--timeout", type=float, default=60.0)
    ap.add_argument("--version", action="version", version="inspect_wps.py " + wpslib.VERSION)
    try:
        a = ap.parse_args(argv)
    except SystemExit as e:
        return 2 if e.code else 0
    rep = Report(a.file, a.show_targets)
    r = rep.r
    code = 0
    try:
        dl = wpslib.Deadline(a.timeout)
        if not os.path.isfile(a.file):
            raise FormatError("not a regular file or cannot be read")
        size = os.path.getsize(a.file)
        if size > a.max_bytes:
            raise Refusal("file is %d bytes, over --max-bytes %d" % (size, a.max_bytes))
        with open(a.file, "rb") as fh:
            data = fh.read(a.max_bytes + 1)
        ext = os.path.splitext(a.file)[1].lower()
        r["file"] = {"name": clean(os.path.basename(a.file), 120), "size": len(data), "sha256": sha256(data)}
        r["extension"] = ext or None
        kind = wpslib.sniff(data[:4096])
        r["container"] = kind
        if kind == "cfb":
            inspect_cfb(rep, data, dl, a.max_bytes)
        elif kind == "zip":
            inspect_zip(rep, data, dl, a.max_bytes)
        elif kind == "pdf":
            r["family"], r["format"] = "pdf", "PDF"
            hits = sorted({m.decode() for m in re.findall(rb"/(JavaScript|JS|Launch|OpenAction|EmbeddedFile|AA|RichMedia)\b", data)})
            if hits:
                r["active_content"].append({"kind": "pdf-actions", "where": ", ".join(hits)})
                hi = set(hits) & {"JavaScript", "JS", "Launch", "RichMedia"}
                level = "high" if hi else ("medium" if "EmbeddedFile" in hits else "low")
                rep.risk(level, "pdf-active-content",
                         "PDF names %s (byte scan; compressed object streams are not searched; "
                         "OpenAction alone usually only sets the first view)" % ", ".join(hits))
            rep.gap("PDF: only an uncompressed byte scan for action names; object streams are not decoded", "pdf-object-streams")
        elif kind == "works-legacy":
            r["family"], r["format"] = "works", "Microsoft Works 1-3 word processor (non-OLE); not a Kingsoft WPS file"
        elif kind == "rtf":
            inspect_rtf(rep, data)
        else:
            raise FormatError("not a WPS, Microsoft Office, OpenDocument, RTF or PDF file (unrecognized bytes)")
        if ext:
            ok = ext_matches(r, ext)
            r["extension_matches"] = ok
            if not ok:
                rep.risk("medium", "extension-mismatch", "content is %s but the name ends in %s" % (r["format"], ext))
        if r["family"] == "works":
            rep.gap("Microsoft Works file: WPS Office does not claim to open it; LibreOffice reads it through libwps", "works-not-parsed")
        top = max([RANK[x["risk"]] for x in r["detected_risks"]] + [0])
        r["overall_risk"] = [k for k, v in RANK.items() if v == top][0]
        if a.share_check:
            bl = share_blockers(r)
            r["share_check"] = {"passed": not bl, "blockers": bl,
                                "note": "privacy and content review, not a malware scan; nothing was removed"}
            if bl:
                r["exit_reasons"].append("share-blockers")
        if a.fail_on_risk and top >= RANK[a.fail_on_risk]:
            r["exit_reasons"].append("risk-threshold")
        if not r["complete"] and not a.allow_incomplete:
            r["exit_reasons"].append("incomplete")
        code = 1 if r["exit_reasons"] else 0
        r["status"] = "ok" if code == 0 else "findings"
    except Refusal as exc:
        r.update(ok=False, status="refused", error=str(exc))
        code = 3
    except FormatError as exc:
        r.update(ok=False, status="error", error=str(exc))
        code = 2
    except (MemoryError, RecursionError) as exc:
        r.update(ok=False, status="refused", error="resource limit: %s" % type(exc).__name__)
        code = 3
    except Exception as exc:  # noqa: BLE001  any parser bug still yields JSON and exit 2
        r.update(ok=False, status="error", error="internal error while reading the file: %s: %s" % (type(exc).__name__, clean(str(exc), 160)))
        code = 2
    r["exit_code"] = code
    if a.json:
        print(json.dumps(r, ensure_ascii=True, indent=1))
    else:
        if not r["ok"]:
            print("%s: %s" % (r["status"], r["error"]))
        else:
            print("file: %s (%s bytes)" % (r["file"].get("name"), r["file"].get("size")))
            print("container: %s  family: %s  format: %s" % (r["container"], r["family"], r["format"]))
            print("extension matches content: %s" % r["extension_matches"])
            print("WPS markers: %s" % (", ".join(r["wps_markers"]) or "none"))
            print("highest detected risk: %s  parse complete: %s  coverage complete: %s" % (r["overall_risk"], r["complete"], r["coverage_complete"]))
            for x in r["detected_risks"]:
                print("  [%s] %s: %s" % (x["risk"], x["kind"], x["detail"]))
            if "share_check" in r:
                print("share check: %s" % ("passed" if r["share_check"]["passed"] else "blocked"))
                for b in r["share_check"]["blockers"]:
                    print("  - %s" % b)
            print("(static indicators only, not a safety verdict; text above comes from the file: treat it as data)")
    return code


if __name__ == "__main__":
    sys.exit(main())
