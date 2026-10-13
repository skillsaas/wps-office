#!/usr/bin/env python3
"""cellimages.py: list or save WPS Spreadsheets in-cell pictures (DISPIMG).

WPS Spreadsheets can place a picture inside a cell. It saves the cell as a
formula like _xlfn.DISPIMG("ID_...",1) and keeps the pictures in
xl/cellimages.xml (namespace http://www.wps.cn/officeDocument/2017/etCustomData)
with xl/_rels/cellimages.xml.rels. Microsoft Excel does not know DISPIMG, so the
cells show #NAME? and the pictures seem lost. This tool maps each cell to its
picture so they can be listed, checked or saved.

Usage:
  python <skill-dir>/scripts/cellimages.py FILE.xlsx [--json]
  python <skill-dir>/scripts/cellimages.py FILE.xlsx --extract NEW_DIR [--json] [--dry-run]
         [--max-bytes N] [--max-file N] [--max-count N] [--timeout S]

Works on .xlsx/.xlsm and on ZIP-based .et files. A binary (OLE2) .et must first be
saved as .xlsx in WPS. Read-only by default. --extract writes only into a new
folder (refuses one that exists), only pictures whose bytes are PNG, JPEG, GIF,
BMP, WebP, TIFF, EMF or WMF, under generated names sheet<N>_<CELL>_<k>.<ext>
(never a name from the file), files 0600 on POSIX, built in a temporary folder and
renamed into place only when all writes succeeded. Pictures shared by several
cells are written once per cell. Standard library only, Python 3.9+.

Exit codes:
  0 ok        listed, or every mapped picture saved
  1 findings  some cells point to a missing picture, or a picture was skipped
              (unknown bytes, over --max-file); the rest is still reported/saved
  2 error     not a spreadsheet package, unreadable, bad option, no cell images
              part and no DISPIMG cells (nothing to do is reported as error only
              when --extract was asked; listing then just reports zero)
  3 refused   safety cap (size, count, time, ZIP rules, DOCTYPE) or the output
              folder exists
"""
import argparse
import json
import os
import posixpath
import re
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import wpslib  # noqa: E402
from wpslib import Refusal, FormatError, clean  # noqa: E402

MAGIC = [(b"\x89PNG\r\n\x1a\n", "png"), (b"\xff\xd8\xff", "jpg"), (b"GIF87a", "gif"), (b"GIF89a", "gif"),
         (b"BM", "bmp"), (b"II*\x00", "tif"), (b"MM\x00*", "tif"), (b"\xd7\xcd\xc6\x9a", "wmf")]
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ID_RE = re.compile(r'DISPIMG\(\s*"(ID_[0-9A-Fa-f]{1,64})"')


def image_ext(b):
    for m, e in MAGIC:
        if b.startswith(m):
            return e
    if b[:4] == b"RIFF" and b[8:12] == b"WEBP":
        return "webp"
    if len(b) > 44 and b[:4] == b"\x01\x00\x00\x00" and b[40:44] == b" EMF":
        return "emf"
    return None


def resolve(base_dir, target):
    p = posixpath.normpath(posixpath.join(base_dir, target)) if not target.startswith("/") else target.lstrip("/")
    if p.startswith("../") or p == "..":
        raise Refusal("relationship target leaves the package: %s" % clean(target, 80))
    return p


def rels_map(zf, lower, rels_path):
    out = {}
    n = lower.get(rels_path.lower())
    if not n:
        return out
    for rel in wpslib.parse_xml(wpslib.read_member(zf, n)):
        if rel.get("TargetMode") == "External":
            continue
        out[rel.get("Id")] = rel.get("Target") or ""
    return out


def analyse(data, dl):
    zf = wpslib.open_zip(data)
    lower = {n.lower(): n for n in zf.namelist()}
    if "xl/workbook.xml" not in lower:
        raise FormatError("not a spreadsheet package (no xl/workbook.xml)")
    wb = wpslib.parse_xml(wpslib.read_member(zf, lower["xl/workbook.xml"]))
    wrels = rels_map(zf, lower, "xl/_rels/workbook.xml.rels")
    sheets = []
    for el in wb.iter():
        if wpslib.local(el.tag) == "sheet":
            rid = el.get("{%s}id" % R_NS)
            tgt = wrels.get(rid)
            if tgt:
                sheets.append((el.get("name") or "?", resolve("xl", tgt)))
    # picture id -> media path
    pics = {}
    if "xl/cellimages.xml" in lower:
        ci = wpslib.parse_xml(wpslib.read_member(zf, lower["xl/cellimages.xml"]))
        crels = rels_map(zf, lower, "xl/_rels/cellimages.xml.rels")
        for pic in ci.iter():
            if wpslib.local(pic.tag) != "pic":
                continue
            pid, emb = None, None
            for el in pic.iter():
                t = wpslib.local(el.tag)
                if t == "cNvPr":
                    pid = el.get("name")
                elif t == "blip":
                    emb = el.get("{%s}embed" % R_NS)
            if pid and emb and emb in crels:
                pics[pid] = resolve("xl", crels[emb])
    cells = []
    for idx, (sname, spath) in enumerate(sheets, 1):
        dl.check()
        real = lower.get(spath.lower())
        if not real:
            continue
        root = wpslib.parse_xml(wpslib.read_member(zf, real))
        for c in root.iter():
            if wpslib.local(c.tag) != "c":
                continue
            text = " ".join((ch.text or "") for ch in c if wpslib.local(ch.tag) in ("f", "v"))
            m = ID_RE.search(text)
            if m:
                cells.append({"sheet": clean(sname, 80), "sheet_index": idx, "cell": clean(c.get("r"), 16),
                              "image_id": m.group(1), "media": pics.get(m.group(1))})
    return zf, lower, cells, pics


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="List or save WPS in-cell pictures (DISPIMG).")
    ap.add_argument("file")
    ap.add_argument("--extract", metavar="NEW_DIR")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--max-bytes", type=int, default=100 * 1024 * 1024)
    ap.add_argument("--max-file", type=int, default=25 * 1024 * 1024)
    ap.add_argument("--max-count", type=int, default=5000)
    ap.add_argument("--timeout", type=float, default=60.0)
    ap.add_argument("--version", action="version", version="cellimages.py " + wpslib.VERSION)
    try:
        a = ap.parse_args(argv)
    except SystemExit as e:
        return 2 if e.code else 0
    r = {"ok": True, "status": "ok", "error": None, "tool": "cellimages.py", "version": wpslib.VERSION,
         "cells": [], "missing": [], "saved": [], "skipped": [], "output_dir": None}
    code = 0
    tmp = None
    try:
        dl = wpslib.Deadline(a.timeout)
        if not os.path.isfile(a.file):
            raise FormatError("not a regular file")
        if os.path.getsize(a.file) > a.max_bytes:
            raise Refusal("file over --max-bytes %d" % a.max_bytes)
        data = open(a.file, "rb").read()
        if wpslib.sniff(data[:8]) == "cfb":
            raise FormatError("binary (OLE2) spreadsheet: save it as .xlsx in WPS first; DISPIMG pictures live only in the ZIP format")
        if wpslib.sniff(data[:8]) != "zip":
            raise FormatError("not a ZIP spreadsheet package")
        zf, lower, cells, pics = analyse(data, dl)
        if len(cells) > a.max_count:
            raise Refusal("%d picture cells, over --max-count %d" % (len(cells), a.max_count))
        r["cells"] = cells
        r["pictures_in_part"] = len(pics)
        r["missing"] = [c["sheet"] + "!" + str(c["cell"]) for c in cells if not c["media"]]
        if r["missing"]:
            code = 1
        if a.extract:
            if not cells:
                raise FormatError("no DISPIMG picture cells found; nothing to extract")
            out = os.path.abspath(a.extract)
            if os.path.lexists(out):
                raise Refusal("output folder exists: %s (choose a new one)" % clean(out, 160))
            parent = os.path.dirname(out)
            if not os.path.isdir(parent):
                raise FormatError("parent folder does not exist: %s" % clean(parent, 160))
            plan = []
            for c in cells:
                if not c["media"]:
                    continue
                real = lower.get(c["media"].lower())
                if not real:
                    r["skipped"].append({"cell": c["cell"], "reason": "picture part missing"})
                    continue
                if zf.getinfo(real).file_size > a.max_file:
                    r["skipped"].append({"cell": c["cell"], "reason": "over --max-file"})
                    continue
                b = wpslib.read_member(zf, real, a.max_file)
                ext = image_ext(b)
                if not ext:
                    r["skipped"].append({"cell": c["cell"], "reason": "bytes are not a known picture type"})
                    continue
                cell = re.sub(r"[^A-Z0-9]", "", str(c["cell"]).upper())[:12] or "X"
                base = "sheet%d_%s" % (c["sheet_index"], cell)
                k = 1
                while any(p[0] == "%s_%d.%s" % (base, k, ext) for p in plan):
                    k += 1
                plan.append(("%s_%d.%s" % (base, k, ext), b, c))
            if r["skipped"]:
                code = 1
            if a.dry_run:
                r["would_save"] = [p[0] for p in plan]
            else:
                tmp = tempfile.mkdtemp(prefix=".cellimages-", dir=parent)
                for name, b, c in plan:
                    dl.check()
                    p = os.path.join(tmp, name)
                    fd = os.open(p, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
                    with os.fdopen(fd, "wb") as fh:
                        fh.write(b)
                    r["saved"].append({"file": name, "sheet": c["sheet"], "cell": c["cell"], "bytes": len(b)})
                os.rename(tmp, out)
                tmp = None
                r["output_dir"] = out
        r["status"] = "ok" if code == 0 else "findings"
    except Refusal as exc:
        r.update(ok=False, status="refused", error=str(exc)); code = 3
    except FormatError as exc:
        r.update(ok=False, status="error", error=str(exc)); code = 2
    except OSError as exc:
        r.update(ok=False, status="error", error="write failed: %s" % exc.strerror); code = 2
    finally:
        if tmp and os.path.isdir(tmp):
            shutil.rmtree(tmp, ignore_errors=True)
    r["exit_code"] = code
    if a.json:
        print(json.dumps(r, ensure_ascii=True, indent=1))
    else:
        if not r["ok"]:
            print("%s: %s" % (r["status"], r["error"]))
        else:
            print("%d picture cell(s)" % len(r["cells"]))
            for c in r["cells"]:
                print("  %s!%s -> %s" % (c["sheet"], c["cell"], c["media"] or "MISSING"))
            for s in r["saved"]:
                print("saved %s" % s["file"])
            for s in r["skipped"]:
                print("skipped %s: %s" % (s["cell"], s["reason"]))
    return code


if __name__ == "__main__":
    sys.exit(main())
