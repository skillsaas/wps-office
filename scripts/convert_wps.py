#!/usr/bin/env python3
"""convert_wps.py: convert a WPS or Microsoft Office file with LibreOffice, sandboxed.

Usage:
  python <skill-dir>/scripts/convert_wps.py IN OUT --to docx|xlsx|pptx|pdf|odt|ods|odp
         [--json] [--max-bytes N] [--timeout S] [--soffice PATH] [--force]
         [--sandboxed]   (operator only, see below)

Typical use: turn a Kingsoft .wps/.et/.dps (or a .doc/.xls/.ppt) into .docx,
.xlsx or .pptx so Office libraries and Microsoft Office can read it, or make a
PDF. WPS Office itself has no documented headless converter; LibreOffice reads
these files by their content (a 2005+ .wps is read like a .doc, .et like .xls,
.dps like .ppt). Fidelity is LibreOffice's, not WPS's: check the result.

Safety, always on:
- The input is inspected first (inspect_wps.py rules); ZIP and size caps apply.
- LibreOffice runs inside bubblewrap (bwrap) with no network (--unshare-all),
  a fresh empty HOME and profile, only /usr and /etc read-only, a private /tmp,
  and a private copy of the input. Document macros are disabled in the
  throwaway profile. The whole process group is killed at --timeout.
- Without bwrap (macOS, Windows, or Linux without user namespaces) it refuses
  (exit 3). --sandboxed tells it the operator already runs it in a VM or
  container with no network and no user files; an agent must never add it to
  get past a refusal. It only skips the bwrap requirement.
- OUT must not exist (exit 3) unless --force (replaces it, keeps no backup). The
  result is written to a temporary file next to OUT and renamed at the end.
- The result is inspected with inspect_wps.py and every risk left in it (macros,
  links, embedded objects) is reported; nothing is assumed to be dropped.

Exit codes:
  0 ok        converted and the output passed the post-check
  1 findings  converted, but the output still has risks (listed) or is incomplete
  2 error     bad option, unsupported input (for example a Works .wps to .xlsx),
              LibreOffice missing or the conversion produced nothing
  3 refused   no sandbox, input over a cap, time limit, OUT exists, ZIP rules
"""
import argparse
import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import wpslib  # noqa: E402
import inspect_wps  # noqa: E402
from wpslib import Refusal, FormatError, clean  # noqa: E402

TARGETS = {
    "docx": ("writer", "docx:MS Word 2007 XML"), "odt": ("writer", "odt"),
    "xlsx": ("spreadsheet", "xlsx:Calc MS Excel 2007 XML"), "ods": ("spreadsheet", "ods"),
    "pptx": ("presentation", "pptx:Impress MS PowerPoint 2007 XML"), "odp": ("presentation", "odp"),
    "pdf": (None, "pdf"),
}
IN_EXT = {"writer": ".doc", "spreadsheet": ".xls", "presentation": ".ppt"}
PROFILE_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="DisableMacrosExecution" oor:op="fuse"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop></item>
</oor:items>
"""


def find_soffice(given):
    if given:
        return given if os.path.isfile(given) else None
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return os.path.realpath(p)
    for p in ("/Applications/LibreOffice.app/Contents/MacOS/soffice",
              r"C:\Program Files\LibreOffice\program\soffice.exe"):
        if os.path.isfile(p):
            return p
    return None


def bwrap_cmd(bwrap, work, soffice_args):
    cmd = [bwrap, "--ro-bind", "/usr", "/usr", "--ro-bind", "/etc", "/etc"]
    for d in ("bin", "lib", "lib64", "sbin"):
        p = "/" + d
        if os.path.islink(p):
            cmd += ["--symlink", os.readlink(p), p]
        elif os.path.isdir(p):
            cmd += ["--ro-bind", p, p]
    if os.path.isdir("/opt") and any(os.path.isdir(os.path.join("/opt", x)) and "libreoffice" in x.lower() for x in os.listdir("/opt")):
        cmd += ["--ro-bind", "/opt", "/opt"]
    cmd += ["--proc", "/proc", "--dev", "/dev", "--tmpfs", "/tmp", "--bind", work, "/work",
            "--unshare-all", "--die-with-parent", "--new-session", "--clearenv",
            "--setenv", "HOME", "/work/home", "--setenv", "PATH", "/usr/bin:/bin",
            "--setenv", "TMPDIR", "/tmp", "--setenv", "LANG", "C.UTF-8", "--chdir", "/work"]
    return cmd + soffice_args


def run(cmd, timeout, env=None):
    kw = {}
    if os.name == "posix":
        kw["start_new_session"] = True
    else:
        kw["creationflags"] = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env, **kw)
    try:
        out, err = p.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        if os.name == "posix":
            try:
                os.killpg(p.pid, signal.SIGKILL)
            except OSError:
                pass
        else:
            p.kill()
        p.communicate()
        raise Refusal("LibreOffice did not finish within --timeout %ss (killed)" % timeout)
    return p.returncode, out.decode("utf-8", "replace"), err.decode("utf-8", "replace")


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="Sandboxed LibreOffice conversion of WPS and Office files.")
    ap.add_argument("input")
    ap.add_argument("output")
    ap.add_argument("--to", required=True, choices=sorted(TARGETS))
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--max-bytes", type=int, default=50 * 1024 * 1024)
    ap.add_argument("--timeout", type=float, default=120.0)
    ap.add_argument("--soffice")
    ap.add_argument("--force", action="store_true", help="replace OUT; keeps no backup")
    ap.add_argument("--sandboxed", action="store_true", help="OPERATOR ONLY: already isolated, skip bwrap")
    ap.add_argument("--version", action="version", version="convert_wps.py " + wpslib.VERSION)
    try:
        a = ap.parse_args(argv)
    except SystemExit as e:
        return 2 if e.code else 0
    r = {"ok": True, "status": "ok", "error": None, "tool": "convert_wps.py", "version": wpslib.VERSION,
         "input_family": None, "engine": None, "sandbox": None, "output": None, "output_check": None}
    code, work, tmp_out = 0, None, None
    try:
        out = os.path.abspath(a.output)
        if not os.path.isfile(a.input):
            raise FormatError("input is not a regular file")
        if os.path.getsize(a.input) > a.max_bytes:
            raise Refusal("input over --max-bytes %d (cap applies before LibreOffice sees it)" % a.max_bytes)
        if os.path.lexists(out) and not a.force:
            raise Refusal("output exists: %s (use a new name, or --force to replace it with no backup)" % clean(out, 160))
        if not os.path.isdir(os.path.dirname(out)):
            raise FormatError("output folder does not exist")
        if os.path.splitext(out)[1].lower() != "." + a.to:
            raise FormatError("output name must end in .%s" % a.to)
        # inspect first
        import io
        import contextlib
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            icode = inspect_wps.main([a.input, "--json", "--allow-incomplete", "--max-bytes", str(a.max_bytes)])
        pre = json.loads(buf.getvalue())
        if icode in (2, 3):
            if icode == 3:
                raise Refusal("input refused by the inspector: %s" % pre.get("error"))
            raise FormatError("input not convertible: %s" % pre.get("error"))
        fam = pre["family"]
        r["input_family"] = fam
        r["input_risks"] = pre["risks"]
        if fam in ("encrypted-ooxml",) or any(x["kind"] == "encrypted" for x in pre["risks"]):
            raise FormatError("input is password-protected; open it in WPS with the password and save an unprotected copy")
        want, filt = TARGETS[a.to]
        if fam == "works" and a.to not in ("docx", "odt", "pdf"):
            raise FormatError("Microsoft Works word processor file can only become docx, odt or pdf")
        if want and fam not in (want, "works", "opendocument") and not (fam == "works" and want == "writer"):
            raise FormatError("a %s file cannot become .%s" % (fam, a.to))
        soffice = find_soffice(a.soffice)
        if not soffice:
            raise FormatError("LibreOffice (soffice) not found; install it or see references/fallbacks.md")
        r["engine"] = "LibreOffice at %s" % soffice
        bwrap = shutil.which("bwrap") if sys.platform.startswith("linux") else None
        if not bwrap and not a.sandboxed:
            raise Refusal("no sandbox: bubblewrap (bwrap) is needed to run LibreOffice on an untrusted file on this system. "
                          "Ask the user; do not add --sandboxed yourself (operator only)")
        work = tempfile.mkdtemp(prefix="wps-convert-")
        for d in ("in", "out", "home", "profile/user"):
            os.makedirs(os.path.join(work, d))
        with open(os.path.join(work, "profile", "user", "registrymodifications.xcu"), "w") as fh:
            fh.write(PROFILE_XCU)
        src_ext = os.path.splitext(a.input)[1].lower()
        name = "input" + (src_ext if src_ext and src_ext.isascii() and len(src_ext) <= 6 and src_ext[1:].isalnum() else IN_EXT.get(fam, ".bin"))
        shutil.copyfile(a.input, os.path.join(work, "in", name))
        if bwrap and not a.sandboxed:
            r["sandbox"] = "bwrap (no network, private HOME and /tmp, /usr and /etc read-only)"
            args = [soffice, "-env:UserInstallation=file:///work/profile", "--headless", "--norestore",
                    "--nologo", "--nolockcheck", "--convert-to", filt, "--outdir", "/work/out", "/work/in/" + name]
            # soffice path must exist inside the sandbox: only /usr, /opt (LibreOffice) are bound
            if not soffice.startswith(("/usr/", "/opt/")):
                raise Refusal("LibreOffice is outside /usr and /opt and cannot run in the sandbox")
            cmd = bwrap_cmd(bwrap, work, args)
            env = {"PATH": "/usr/bin:/bin"}
        else:
            r["sandbox"] = "operator asserted (--sandboxed); no bwrap"
            prof = "file:///" + os.path.join(work, "profile").replace("\\", "/").lstrip("/")
            cmd = [soffice, "-env:UserInstallation=" + prof, "--headless", "--norestore", "--nologo",
                   "--nolockcheck", "--convert-to", filt, "--outdir", os.path.join(work, "out"),
                   os.path.join(work, "in", name)]
            env = {"PATH": os.environ.get("PATH", ""), "HOME": os.path.join(work, "home"),
                   "TMPDIR": os.path.join(work, "home"), "SYSTEMROOT": os.environ.get("SYSTEMROOT", "")}
        rc, so, se = run(cmd, a.timeout, env)
        produced = [f for f in os.listdir(os.path.join(work, "out"))]
        if not produced:
            raise FormatError("LibreOffice produced no output (exit %s): %s" % (rc, clean(se or so, 300)))
        res = os.path.join(work, "out", produced[0])
        if os.path.getsize(res) == 0:
            raise FormatError("LibreOffice wrote an empty file")
        tmp_out = out + ".partial-%d" % os.getpid()
        shutil.copyfile(res, tmp_out)
        os.replace(tmp_out, out)
        tmp_out = None
        r["output"] = out
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            pcode = inspect_wps.main([out, "--json"])
        post = json.loads(buf.getvalue())
        r["output_check"] = {"family": post.get("family"), "overall_risk": post.get("overall_risk"),
                             "risks": post.get("risks"), "complete": post.get("complete"),
                             "macros_left": [x["kind"] for x in post.get("active_content", [])]}
        if pcode != 0 or post.get("overall_risk") in ("medium", "high"):
            code = 1
        r["not_checked"] = ["layout, fonts, formulas and charts are LibreOffice's rendering, not WPS's: compare with the original",
                            "WPS-only features (DISPIMG cell pictures, WPS fonts, JS macros) may be dropped or changed"]
        r["status"] = "ok" if code == 0 else "findings"
    except Refusal as exc:
        r.update(ok=False, status="refused", error=str(exc)); code = 3
    except FormatError as exc:
        r.update(ok=False, status="error", error=str(exc)); code = 2
    except OSError as exc:
        r.update(ok=False, status="error", error="OS error: %s" % exc); code = 2
    finally:
        if work:
            shutil.rmtree(work, ignore_errors=True)
        if tmp_out and os.path.exists(tmp_out):
            os.remove(tmp_out)
    r["exit_code"] = code
    if a.json:
        print(json.dumps(r, ensure_ascii=True, indent=1))
    else:
        if not r["ok"]:
            print("%s: %s" % (r["status"], r["error"]))
        else:
            print("converted (%s) -> %s" % (r["input_family"], r["output"]))
            print("sandbox: %s" % r["sandbox"])
            oc = r["output_check"]
            print("output risk: %s; complete: %s" % (oc["overall_risk"], oc["complete"]))
            for x in oc["risks"] or []:
                print("  [%s] %s: %s" % (x["risk"], x["kind"], x["detail"]))
    return code


if __name__ == "__main__":
    sys.exit(main())
