#!/usr/bin/env python3
"""Smoke tests for the wps-office skill. Run: python tests/smoke_test.py

Builds synthetic fixtures in a temp folder, runs every script, and prints
PASS/FAIL/SKIP with a summary. Exit 0 when nothing failed. Optional parts skip
cleanly: olefile cross-read (pip install olefile), LibreOffice conversion and
the network canary (need soffice; the sandbox tests need Linux with bwrap).
"""
import http.server
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
S = os.path.join(ROOT, "scripts")
sys.path.insert(0, os.path.join(ROOT, "tests"))
sys.path.insert(0, S)
import make_fixtures  # noqa: E402

RESULTS = []
TMP = tempfile.mkdtemp(prefix="wps-skill-test-")
FX = make_fixtures.build(os.path.join(TMP, "fx"))


def run(script, *args, timeout=300, env=None):
    p = subprocess.run([sys.executable, os.path.join(S, script)] + list(args), capture_output=True,
                       text=True, timeout=timeout, env=env)
    return p.returncode, p.stdout, p.stderr


def js(script, *args, **kw):
    rc, out, err = run(script, *(list(args) + ["--json"]), **kw)
    try:
        return rc, json.loads(out)
    except ValueError:
        raise AssertionError("no JSON (rc %s): %s %s" % (rc, out[:300], err[:300]))


def f(name):
    return os.path.join(FX, name)


def test(fn):
    name = fn.__name__
    try:
        res = fn()
        if res == "SKIP" or (isinstance(res, tuple) and res[0] == "SKIP"):
            RESULTS.append(("SKIP", name, res[1] if isinstance(res, tuple) else ""))
        else:
            RESULTS.append(("PASS", name, ""))
    except Exception as exc:  # noqa: BLE001
        RESULTS.append(("FAIL", name, "%s: %s" % (type(exc).__name__, exc)))
    return fn


def kinds(d):
    return [x["kind"] for x in d["detected_risks"]]


# ---------------------------------------------------------------- inspector
@test
def kingsoft_writer_detected_with_markers():
    rc, d = js("inspect_wps.py", f("kingsoft.wps"))
    assert rc == 0 and d["ok"] and d["family"] == "writer" and d["container"] == "cfb", d
    assert "KSOProductBuildVer" in d["wps_markers"] and d["metadata"]["author"] == "Zhang San"
    assert d["metadata"]["company"] == "Example Co" and d["producer"] == "WPS Office"


@test
def spreadsheet_and_presentation_families():
    assert js("inspect_wps.py", f("sheet.et"))[1]["family"] == "spreadsheet"
    assert js("inspect_wps.py", f("slides.dps"))[1]["family"] == "presentation"


@test
def works_wps_is_not_kingsoft():
    for n in ("works.wps", "works_legacy.wps"):
        rc, d = js("inspect_wps.py", f(n))
        assert rc == 0 and d["family"] == "works" and "Works" in d["format"], d


@test
def vba_and_ole_package_are_high():
    rc, d = js("inspect_wps.py", f("macro.wps"), "--fail-on-risk", "high")
    assert rc == 1 and "macros" in kinds(d) and "ole-package" in kinds(d) and d["overall_risk"] == "high"


@test
def wps_js_macro_part_detected():
    rc, d = js("inspect_wps.py", f("jsa.xlsm"))
    assert any(a["kind"] == "wps-js-macro" for a in d["active_content"]) and d["overall_risk"] == "high"


@test
def macro_under_plain_extension():
    assert "macros-under-plain-extension" in kinds(js("inspect_wps.py", f("macro_under_docx.docx"))[1])


@test
def encrypted_variants_incomplete_exit_1():
    for n in ("encrypted.wps", "encrypted.et", "encrypted_ooxml.docx"):
        rc, d = js("inspect_wps.py", f(n))
        assert rc == 1 and "encrypted" in kinds(d) and d["complete"] is False, (n, d)
        rc2, _ = js("inspect_wps.py", f(n), "--allow-incomplete")
        assert rc2 == 0


@test
def damaged_fat_past_eof_reported_not_crash():
    rc, d = js("inspect_wps.py", f("damaged.wps"))
    assert rc == 1 and d["ok"] and any("past-end" in i for i in d["incomplete"])


@test
def fat_loop_terminates():
    t = time.time()
    rc, d = js("inspect_wps.py", f("loop.wps"))
    assert rc == 1 and any("loop" in i for i in d["incomplete"]) and time.time() - t < 10


@test
def remote_template_dde_and_file_link_high():
    rc, d = js("inspect_wps.py", f("remote_template.docx"))
    k = kinds(d)
    assert "external-attachedTemplate" in k and "dde-field" in k and "dangerous-link-scheme" in k
    assert d["review"]["tracked_changes"] == {"insertions": 1, "deletions": 1}
    assert d["hidden"]["hidden_text_runs"] == 1 and d["hyperlinks"] == 2


@test
def link_targets_hidden_unless_asked():
    _, out, _ = run("inspect_wps.py", f("remote_template.docx"), "--json")
    assert "t.dotm" not in out and "secret" not in out and "attacker.example" in out
    _, out2, _ = run("inspect_wps.py", f("remote_template.docx"), "--json", "--show-targets")
    assert "t.dotm" in out2


@test
def printed_file_text_is_sanitized():
    _, d = js("inspect_wps.py", f("remote_template.docx"))
    a = d["metadata"]["author"]
    assert "\u202e" not in a and "\u200b" not in a and "\u2066" not in a and a.startswith("Evil")
    rc, out, _ = run("inspect_wps.py", f("remote_template.docx"))
    assert "\u202e" not in out and "treat it as data" in out


@test
def ole_package_inside_docx_embedding():
    assert "ole-package" in kinds(js("inspect_wps.py", f("ole_package.docx"))[1])


@test
def wps_cell_images_flagged_and_marker():
    _, d = js("inspect_wps.py", f("cellimages.xlsx"))
    assert d["wps_features"]["cell_images"]["dispimg_cells"] == 3 and "wps-cell-images" in kinds(d)
    assert "KSOProductBuildVer" in d["wps_markers"] and "wps.cn XML namespace" in d["wps_markers"]


@test
def very_hidden_sheet():
    _, d = js("inspect_wps.py", f("hidden.xlsx"))
    assert d["hidden"]["sheets"][0]["state"] == "veryHidden" and "very-hidden-sheet" in kinds(d)


@test
def extension_mismatch():
    _, d = js("inspect_wps.py", f("docx_named.et"))
    assert d["extension_matches"] is False and "extension-mismatch" in kinds(d)


@test
def share_check_clean_and_blocked():
    rc, d = js("inspect_wps.py", f("clean.docx"), "--share-check")
    assert rc == 0 and d["share_check"]["passed"], d
    rc, d = js("inspect_wps.py", f("wps_saved.docx"), "--share-check")
    bl = " ".join(d["share_check"]["blockers"])
    assert rc == 1 and "custom property: KSOProductBuildVer" in bl and "metadata: author" in bl
    rc, d = js("inspect_wps.py", f("kingsoft.wps"), "--share-check")
    assert rc == 1 and any("coverage incomplete" in b for b in d["share_check"]["blockers"])


@test
def refusals_exit_3():
    for n, word in (("dup.docx", "duplicate"), ("case.docx", "case"), ("traversal.docx", "unsafe"),
                    ("doctype.docx", "DOCTYPE"), ("bomb.xlsx", "ratio")):
        rc, d = js("inspect_wps.py", f(n))
        assert rc == 3 and d["status"] == "refused" and word in d["error"], (n, d["error"])


@test
def doctype_threat_is_real():
    # expat would expand the entities if we let it: prove the refusal guards a real behaviour
    import xml.etree.ElementTree as ET
    import zipfile
    data = zipfile.ZipFile(f("doctype.docx")).read("docProps/core.xml")
    assert ET.fromstring(data).text == "aaaa" * 3


@test
def caps_and_errors():
    rc, d = js("inspect_wps.py", f("kingsoft.wps"), "--max-bytes", "100")
    assert rc == 3 and "max-bytes" in d["error"]
    rc, d = js("inspect_wps.py", f("macro.wps"), "--timeout", "0")
    assert rc == 3 and "time" in d["error"]
    rc, d = js("inspect_wps.py", f("notes.txt"))
    assert rc == 2 and d["status"] == "error"
    rc, d = js("inspect_wps.py", os.path.join(TMP, "missing.wps"))
    assert rc == 2
    rc, _, _ = run("inspect_wps.py", "--fail-on-risk", "bogus", f("clean.docx"))
    assert rc == 2


@test
def worst_case_cfb_timing():
    import cfbwriter
    big = cfbwriter.build([("WordDocument", make_fixtures.FIB + b"A" * (4 * 1024 * 1024))])
    p = os.path.join(TMP, "big.wps")
    open(p, "wb").write(big)
    t = time.time()
    rc, d = js("inspect_wps.py", p)
    assert rc == 0 and time.time() - t < 20, time.time() - t


@test
def olefile_cross_read():
    try:
        import olefile
    except ImportError:
        return ("SKIP", "olefile not installed")
    o = olefile.OleFileIO(f("kingsoft.wps"))
    assert o.get_metadata().author == b"Zhang San"
    names = sorted("/".join(x) for x in o.listdir(streams=True, storages=False))
    _, d = js("inspect_wps.py", f("kingsoft.wps"))
    mine = sorted(s["path"] for s in d["streams"] if s["type"] == "stream")
    assert sorted(n.lstrip("\x05") for n in names) == mine, (names, mine)


# ---------------------------------------------------------------- cellimages
@test
def cellimages_list_maps_cells():
    rc, d = js("cellimages.py", f("cellimages.xlsx"))
    m = {c["cell"]: c["media"] for c in d["cells"]}
    assert rc == 1 and m == {"A1": "xl/media/image1.png", "B1": "xl/media/image2.png", "B2": None}
    assert d["missing"] == ["Sheet1!B2"]


@test
def cellimages_extract_only_real_images():
    out = os.path.join(TMP, "ci1")
    rc, d = js("cellimages.py", f("cellimages.xlsx"), "--extract", out)
    assert rc == 1 and [s["file"] for s in d["saved"]] == ["sheet1_A1.png"]
    assert d["skipped"][0]["reason"].startswith("bytes are not")
    assert os.listdir(out) == ["sheet1_A1.png"]
    if os.name == "posix":
        assert oct(os.stat(os.path.join(out, "sheet1_A1.png")).st_mode & 0o777) == "0o600"


@test
def cellimages_ok_et_and_dry_run():
    out = os.path.join(TMP, "ci2")
    rc, d = js("cellimages.py", f("cellimages_ok.et"), "--extract", out, "--dry-run")
    assert rc == 0 and d["would_save"] == ["sheet1_C3.png"] and not os.path.exists(out)
    rc, d = js("cellimages.py", f("cellimages_ok.et"), "--extract", out)
    assert rc == 0 and os.listdir(out) == ["sheet1_C3.png"]
    rc, d = js("cellimages.py", f("cellimages_ok.et"), "--extract", out)
    assert rc == 3 and "exists" in d["error"]


@test
def cellimages_refuses_traversal_and_binary():
    rc, d = js("cellimages.py", f("traversal_cellimages.xlsx"))
    assert rc == 3 and "leaves the package" in d["error"]
    rc, d = js("cellimages.py", f("sheet.et"))
    assert rc == 2 and "save it as .xlsx" in d["error"]
    rc, d = js("cellimages.py", f("clean.docx"))
    assert rc == 2


@test
def cellimages_no_leftover_temp_dirs():
    left = [n for n in os.listdir(TMP) if n.startswith(".cellimages-")]
    assert not left, left


# ---------------------------------------------------------------- converter
def soffice():
    import convert_wps
    return convert_wps.find_soffice(None)


def sandbox_ok():
    if not sys.platform.startswith("linux") or not shutil.which("bwrap"):
        return False
    p = subprocess.run(["bwrap", "--ro-bind", "/usr", "/usr", "--symlink", "usr/bin", "/bin", "--symlink", "usr/lib", "/lib",
                        "--symlink", "usr/lib64", "/lib64", "--unshare-all", "/usr/bin/true"], capture_output=True)
    return p.returncode == 0


@test
def convert_refuses_without_sandbox():
    if not sys.platform.startswith("linux"):
        rc, d = js("convert_wps.py", f("clean.docx"), os.path.join(TMP, "x.pdf"), "--to", "pdf")
        assert rc in (2, 3), d  # 2 when LibreOffice is missing, 3 = no sandbox
        if soffice():
            assert rc == 3 and "sandbox" in d["error"]
        return
    if not soffice():
        return ("SKIP", "LibreOffice not installed")
    fake = os.path.join(TMP, "nobwrap")
    os.makedirs(fake, exist_ok=True)
    env = dict(os.environ, PATH=fake)
    rc, d = js("convert_wps.py", f("clean.docx"), os.path.join(TMP, "x.pdf"), "--to", "pdf", "--soffice", soffice(), env=env)
    assert rc == 3 and "sandbox" in d["error"] and "--sandboxed" not in d["error"], d


@test
def convert_guards_before_running():
    out = os.path.join(TMP, "exists.docx")
    open(out, "w").close()
    rc, d = js("convert_wps.py", f("clean.docx"), out, "--to", "docx")
    assert rc == 3 and "exists" in d["error"]
    rc, d = js("convert_wps.py", f("sheet.et"), os.path.join(TMP, "s.docx"), "--to", "docx")
    assert rc == 2 and "cannot become" in d["error"]
    rc, d = js("convert_wps.py", f("encrypted.wps"), os.path.join(TMP, "e.docx"), "--to", "docx")
    assert rc == 2 and "password" in d["error"]
    rc, d = js("convert_wps.py", f("kingsoft.wps"), os.path.join(TMP, "k.docx"), "--to", "docx", "--max-bytes", "10")
    assert rc == 3
    rc, d = js("convert_wps.py", f("dup.docx"), os.path.join(TMP, "d.pdf"), "--to", "pdf")
    assert rc == 3
    rc, d = js("convert_wps.py", f("clean.docx"), os.path.join(TMP, "wrong.pdf"), "--to", "docx")
    assert rc == 2


def real_binary(kind):
    """A real .doc/.xls made by LibreOffice from our own text (trusted input)."""
    src = os.path.join(TMP, "src_%s" % kind)
    os.makedirs(src, exist_ok=True)
    if kind == "doc":
        open(os.path.join(src, "w.txt"), "w").write("hello wps writer\n")
        inp, filt = "w.txt", "doc"
    else:
        open(os.path.join(src, "s.csv"), "w").write("item,qty\nwidget,42\n")
        inp, filt = "s.csv", "xls"
    home = os.path.join(src, "home")
    os.makedirs(home, exist_ok=True)
    subprocess.run([soffice(), "-env:UserInstallation=file://" + os.path.join(src, "prof"), "--headless",
                    "--convert-to", filt, "--outdir", src, os.path.join(src, inp)], capture_output=True,
                   timeout=180, env=dict(os.environ, HOME=home))
    out = os.path.join(src, os.path.splitext(inp)[0] + "." + filt)
    return out if os.path.exists(out) else None


@test
def convert_wps_and_et_round_trip():
    if not soffice() or not sandbox_ok():
        return ("SKIP", "needs LibreOffice and a working bwrap")
    import zipfile
    doc = real_binary("doc")
    wps = os.path.join(TMP, "real.wps")
    shutil.copyfile(doc, wps)
    rc, d = js("inspect_wps.py", wps)
    assert d["family"] == "writer"
    out = os.path.join(TMP, "real.docx")
    rc, d = js("convert_wps.py", wps, out, "--to", "docx")
    assert rc == 0 and d["sandbox"].startswith("bwrap"), d
    assert "hello wps writer" in zipfile.ZipFile(out).read("word/document.xml").decode()
    xls = real_binary("xls")
    et = os.path.join(TMP, "real.et")
    shutil.copyfile(xls, et)
    out = os.path.join(TMP, "real.xlsx")
    rc, d = js("convert_wps.py", et, out, "--to", "xlsx")
    z = zipfile.ZipFile(out)
    blob = b"".join(z.read(n) for n in z.namelist() if n.endswith(".xml"))
    assert rc == 0 and b"widget" in blob
    assert not [n for n in os.listdir(TMP) if ".partial-" in n]


class _H(http.server.BaseHTTPRequestHandler):
    hits = []

    def do_GET(self):
        _H.hits.append(self.path)
        self.send_response(404)
        self.end_headers()

    do_OPTIONS = do_HEAD = do_PROPFIND = do_GET

    def log_message(self, *a):
        pass


def server():
    srv = http.server.HTTPServer(("127.0.0.1", 0), _H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


@test
def canary_threat_is_real_without_sandbox():
    if not soffice() or os.name != "posix":
        return ("SKIP", "needs LibreOffice on POSIX")
    srv = server()
    _H.hits = []
    p = os.path.join(TMP, "canary_raw.docx")
    make_fixtures.canary_docx(p, srv.server_address[1])
    w = os.path.join(TMP, "raw")
    os.makedirs(os.path.join(w, "home"), exist_ok=True)
    subprocess.run([soffice(), "-env:UserInstallation=file://" + os.path.join(w, "prof"), "--headless",
                    "--convert-to", "pdf", "--outdir", w, p], capture_output=True, timeout=180,
                   env=dict(os.environ, HOME=os.path.join(w, "home")))
    time.sleep(0.5)
    srv.shutdown()
    if not _H.hits:
        return ("SKIP", "this LibreOffice build did not fetch the linked image even unsandboxed")
    assert any("canary" in h for h in _H.hits)


@test
def canary_no_network_in_sandbox():
    if not soffice() or not sandbox_ok():
        return ("SKIP", "needs LibreOffice and a working bwrap")
    srv = server()
    _H.hits = []
    p = os.path.join(TMP, "canary.docx")
    make_fixtures.canary_docx(p, srv.server_address[1])
    rc, d = js("convert_wps.py", p, os.path.join(TMP, "canary.pdf"), "--to", "pdf", "--allow-risky-input")
    time.sleep(0.5)
    srv.shutdown()
    assert rc in (0, 1) and d["output"], d
    assert _H.hits == [], _H.hits
    assert any(x["kind"] == "external-image" for x in d["input_risks"])


@test
def convert_timeout_kills():
    if not soffice() or not sandbox_ok():
        return ("SKIP", "needs LibreOffice and a working bwrap")
    rc, d = js("convert_wps.py", f("clean.docx"), os.path.join(TMP, "slow.pdf"), "--to", "pdf", "--timeout", "0.01")
    assert rc == 3 and "timeout" in d["error"]
    assert not os.path.exists(os.path.join(TMP, "slow.pdf"))


# ---------------------------------------------------------------- release consistency
SHIPPED_TEXT = None


def shipped():
    out = []
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "__pycache__", "node_modules")]
        for n in files:
            if n.endswith((".md", ".py", ".json", ".txt", ".yml", ".gitignore", ".gitattributes")) or n == "LICENSE":
                out.append(os.path.join(base, n))
    return out


@test
def no_em_dashes():
    bad = [os.path.relpath(p, ROOT) for p in shipped() if "\u2014" in open(p, encoding="utf-8").read()]
    assert not bad, bad


@test
def versions_and_names_in_sync():
    pkg = json.load(open(os.path.join(ROOT, "package.json")))
    sk = json.load(open(os.path.join(ROOT, "skill.json")))
    md = open(os.path.join(ROOT, "SKILL.md"), encoding="utf-8").read()
    fm = md.split("---")[1]
    name = re.search(r"^name:\s*(\S+)", fm, re.M).group(1)
    ver = re.search(r'^\s+version:\s*"([^"]+)"', fm, re.M).group(1)
    import wpslib
    assert pkg["name"] == sk["name"] == name == "wps-office" == os.path.basename(ROOT) or os.path.basename(ROOT) != "wps-office"
    assert pkg["version"] == sk["version"] == ver == wpslib.VERSION, (pkg["version"], sk["version"], ver, wpslib.VERSION)
    desc = re.search(r'^description:\s*"(.*)"\s*$', fm, re.M).group(1)
    assert desc.startswith("Use when") and len(desc) <= 1024
    assert len(md.splitlines()) < 500
    assert not pkg.get("scripts") and "bin" not in pkg and pkg["publishConfig"]["access"] == "public"
    for s in ("inspect_wps.py", "cellimages.py", "convert_wps.py"):
        rc, out, _ = run(s, "--version")
        assert rc == 0 and wpslib.VERSION in out


@test
def referenced_files_exist():
    md = open(os.path.join(ROOT, "SKILL.md"), encoding="utf-8").read()
    for ref in set(re.findall(r"`?(references/[a-z0-9-]+\.md)", md)) | set("references/" + x for x in re.findall(r"`([a-z0-9-]+\.md)`", md)):
        assert os.path.exists(os.path.join(ROOT, ref)), ref
    pkg = json.load(open(os.path.join(ROOT, "package.json")))
    for entry in pkg["files"]:
        assert os.path.exists(os.path.join(ROOT, entry.rstrip("/").replace("*.py", ""))), entry


@test
def link_schemes_normalized():
    rc, d = js("inspect_wps.py", f("links.docx"))
    high = [x for x in d["detected_risks"] if x["kind"] == "dangerous-link-scheme"]
    assert len(high) == 7, d["detected_risks"]
    sys.path.insert(0, os.path.join(ROOT, "scripts"))
    import wpslib
    for raw, scheme in (("ms-msdt:/id x", "ms-msdt"), (" file:///c:/x", "file"), ("%66ile://s/x", "file"),
                        ("%2566ile://s/x", "file"), ("search-ms:q", "search-ms"), ("mhtml:http://a/b!c", "mhtml"),
                        ("ms-excel:ofe|u|http://a/b", "ms-excel"), ("\\\\srv\\s", "unc"), ("C:\\x.exe", "file"),
                        ("fi\u200ble:///x", "file")):
        _n, sc, _h, risk = wpslib.link_target(raw)
        assert sc == scheme and risk == "high", (raw, sc, risk)
    assert wpslib.link_target("https://ok.example/")[3] == "none"


@test
def utf16_doctype_refused():
    rc, d = js("inspect_wps.py", f("doctype16.docx"))
    assert rc == 3 and "DOCTYPE" in d["error"], d


@test
def writer_field_and_hidden_text_rules():
    rc, d = js("inspect_wps.py", f("split_dde.docx"))
    assert "dde-field" in kinds(d) and d["hidden"].get("hidden_text_runs") == 1, d
    rc, d = js("inspect_wps.py", f("vanish_off.docx"))
    assert not d["hidden"].get("hidden_text_runs"), d


@test
def renamed_vba_and_unknown_binary_part():
    _, d = js("inspect_wps.py", f("renamed_vba.docx"))
    assert {"macros", "renamed-macro-part"} <= set(kinds(d)), kinds(d)
    _, d = js("inspect_wps.py", f("unknown_bin.docx"))
    assert "unknown-binary-or-script-part" in kinds(d)


@test
def spreadsheet_formula_risks_and_hidden_rows():
    rc, d = js("inspect_wps.py", f("formulas.xlsx"), "--share-check")
    k = set(kinds(d))
    assert {"dde-link", "dde-formula", "call-formula", "web-formula", "hyperlink-formula", "formula-like-text"} <= k, k
    assert d["hidden"]["rows"] == 1 and d["hidden"]["columns"] == 1 and rc == 1


@test
def odf_macros_and_events():
    _, d = js("inspect_wps.py", f("macro.odt"))
    assert {"macros", "odf-event-script"} <= set(kinds(d)), kinds(d)


@test
def rtf_template_dde_and_rtf_under_docx():
    _, d = js("inspect_wps.py", f("template.rtf"))
    assert {"external-attachedTemplate", "dde-field"} <= set(kinds(d)), kinds(d)
    _, d = js("inspect_wps.py", f("rtf_named.docx"))
    assert "extension-mismatch" in kinds(d) and d["extension_matches"] is False


@test
def detected_risks_framing_and_exit_reasons():
    rc, d = js("inspect_wps.py", f("macro.wps"), "--fail-on-risk", "high")
    assert rc == 1 and "risk-threshold" in d["exit_reasons"] and d["assessment"] and "risks" not in d
    assert any("malware" in x for x in d["not_checked"])


@test
def fuzz_never_crashes():
    """Seeded mutation fuzz: every input gives JSON and an exit code in 0..3."""
    import random
    import io
    import contextlib
    sys.path.insert(0, os.path.join(ROOT, "scripts"))
    import inspect_wps
    import cellimages
    rnd = random.Random(20261013)
    seeds = [open(f(n), "rb").read() for n in ("clean.docx", "kingsoft.wps", "cellimages.xlsx", "formulas.xlsx",
                                                "macro.odt", "template.rtf", "encrypted.et", "sheet.et")]
    path = os.path.join(TMP, "fuzz.bin")
    n = int(os.environ.get("WPS_FUZZ_N", "300"))
    for i in range(n):
        b = bytearray(rnd.choice(seeds))
        for _ in range(rnd.randint(1, 12)):
            op = rnd.random()
            if not b:
                break
            pos = rnd.randrange(len(b))
            if op < 0.6:
                b[pos] = rnd.randrange(256)
            elif op < 0.8:
                del b[pos:pos + rnd.randint(1, 64)]
            else:
                b[pos:pos] = bytes(rnd.randrange(256) for _ in range(rnd.randint(1, 16)))
        open(path, "wb").write(bytes(b))
        for mod, args in ((inspect_wps, [path, "--json", "--timeout", "10"]), (cellimages, [path, "--json"])):
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                rc = mod.main(args)
            d = json.loads(buf.getvalue())
            assert rc in (0, 1, 2, 3) and d["exit_code"] == rc, (i, rc)
            assert "Traceback" not in buf.getvalue()


@test
def cellimages_shared_picture_written_once_under_ulimit():
    if not sys.platform.startswith("linux"):
        return ("SKIP", "RLIMIT_AS is enforced reliably only on Linux")
    import random
    big = make_fixtures.PNG[:-12] + bytes(random.Random(1).randrange(256) for _ in range(9 * 1024 * 1024)) + make_fixtures.PNG[-12:]
    cells = "".join('<row r="%d">%s</row>' % (i, make_fixtures.dispimg_cell("A%d" % i, make_fixtures.ID1)) for i in range(1, 5001))
    p = os.path.join(TMP, "shared.xlsx")
    make_fixtures.xlsx(p, cells, extra=[("xl/cellimages.xml", make_fixtures.cellimages([make_fixtures.ID1])),
                                        ("xl/_rels/cellimages.xml.rels", make_fixtures.RELS % make_fixtures.rel("rId1", "image", "media/image1.png")),
                                        ("xl/media/image1.png", big)])
    out = os.path.join(TMP, "shared_out")
    code = ("import resource,sys,runpy;resource.setrlimit(resource.RLIMIT_AS,(1024**3,1024**3));"
            "sys.argv=['cellimages.py',%r,'--extract',%r,'--json'];runpy.run_path(%r,run_name='__main__')"
            % (p, out, os.path.join(ROOT, "scripts", "cellimages.py")))
    pr = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, timeout=300)
    d = json.loads(pr.stdout)
    assert pr.returncode == 0, (pr.returncode, d.get("error"))
    assert len(d["saved"]) == 1 and d["saved"][0]["shared_by"] == 5000 and len(os.listdir(out)) == 1
    rc, d = js("cellimages.py", p, "--extract", os.path.join(TMP, "shared_cap"), "--max-total", "1000")
    assert rc == 3 and "max-total" in d["error"] and not os.path.exists(os.path.join(TMP, "shared_cap"))


@test
def convert_output_copy_refuses_symlinks_and_has_no_bypass():
    sys.path.insert(0, os.path.join(ROOT, "scripts"))
    import convert_wps
    from wpslib import Refusal
    src = open(os.path.join(ROOT, "scripts", "convert_wps.py")).read()
    assert "--sandboxed" not in src and "--max-memory" in src
    d = os.path.join(TMP, "outdir")
    os.makedirs(d)
    if hasattr(os, "symlink"):
        try:
            os.symlink(os.path.join(ROOT, "SKILL.md"), os.path.join(d, "x.pdf"))
        except OSError:
            return ("SKIP", "cannot create symlinks here")
        try:
            convert_wps.single_output(d, 10 ** 9)
            raise AssertionError("symlink accepted")
        except Refusal:
            pass
        os.remove(os.path.join(d, "x.pdf"))
    open(os.path.join(d, "a.pdf"), "w").write("a")
    open(os.path.join(d, "b.pdf"), "w").write("b")
    try:
        convert_wps.single_output(d, 10 ** 9)
        raise AssertionError("two outputs accepted")
    except Refusal:
        pass


@test
def convert_refuses_risky_input_by_default():
    if not soffice() or not sandbox_ok():
        return ("SKIP", "needs LibreOffice and a working bwrap")
    rc, d = js("convert_wps.py", f("macro_under_docx.docx"), os.path.join(TMP, "m.pdf"), "--to", "pdf")
    assert rc == 3 and "input-risks" in d["exit_reasons"], d


def main():
    for status, name, why in RESULTS:
        print("%s %s%s" % (status, name, (" (" + why + ")") if why else ""))
    n = {k: sum(1 for r in RESULTS if r[0] == k) for k in ("PASS", "FAIL", "SKIP")}
    print("\n%d passed, %d failed, %d skipped (of %d)" % (n["PASS"], n["FAIL"], n["SKIP"], len(RESULTS)))
    shutil.rmtree(TMP, ignore_errors=True)
    return 1 if n["FAIL"] else 0


if __name__ == "__main__":
    sys.exit(main())
