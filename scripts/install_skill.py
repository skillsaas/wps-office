#!/usr/bin/env python3
"""install_skill.py: copy this skill into an agent's skills folder in one step.

Usage:
  python3 scripts/install_skill.py --claude              # ~/.claude/skills/wps-office
  python3 scripts/install_skill.py --claude-project      # ./.claude/skills/wps-office
  python3 scripts/install_skill.py --dest DIR            # DIR/wps-office (any agent)
  add --dry-run to list what would be copied, --replace to move an existing copy
  aside (to wps-office.old-<time> next to the skills folder,
  so the agent does not load two copies; never deleted), --json for a JSON report.

Agents that read skills look for <skills dir>/<name>/SKILL.md; the folder name
must match the "name" in SKILL.md (wps-office). Claude Code reads
~/.claude/skills (personal) and .claude/skills (project). For other agents,
pass their skills folder with --dest.

Copies only the files the skill needs (SKILL.md, references, scripts, the
self-test); no network, standard library only. Checks every copied file's
SHA-256 against the source. Exit codes: 0 installed (or dry run), 2 error,
3 refused (destination exists without --replace, or is a symlink).
"""
import argparse
import hashlib
import json
import os
import shutil
import sys
import time

NAME = "wps-office"
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = ["SKILL.md", "skill.json", "AGENTS.md", "README.md", "LICENSE", "CHANGELOG.md", "SECURITY.md",
         "VALIDATION.md", "scripts/requirements.txt", "tests/smoke_test.py", "tests/make_fixtures.py",
         "tests/cfbwriter.py"]
DIRS = [("references", (".md",)), ("scripts", (".py",))]


def plan():
    out = [f for f in FILES if os.path.isfile(os.path.join(HERE, f))]
    for d, exts in DIRS:
        base = os.path.join(HERE, d)
        for n in sorted(os.listdir(base)):
            p = os.path.join(base, n)
            if n.endswith(exts) and os.path.isfile(p) and not os.path.islink(p):
                out.append(d + "/" + n)
    return sorted(set(out))


def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as fh:
        for b in iter(lambda: fh.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def main(argv=None):
    ap = argparse.ArgumentParser(description="Install the wps-office skill into an agent's skills folder.")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--dest", help="skills folder; the skill goes into DEST/wps-office")
    g.add_argument("--claude", action="store_true", help="Claude Code personal skills: ~/.claude/skills")
    g.add_argument("--claude-project", action="store_true", help="Claude Code project skills: ./.claude/skills")
    ap.add_argument("--replace", action="store_true", help="move an existing wps-office folder aside first")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    root = a.dest if a.dest else (os.path.join(os.path.expanduser("~"), ".claude", "skills") if a.claude
                                  else os.path.join(os.getcwd(), ".claude", "skills"))
    target = os.path.join(os.path.abspath(root), NAME)
    r = {"ok": True, "status": "ok", "error": None, "tool": "install_skill.py", "target": target,
         "files": plan(), "moved_aside": None, "dry_run": a.dry_run}
    code = 0
    try:
        if os.path.realpath(target) == os.path.realpath(HERE):
            raise ValueError("refuse:destination is this source folder")
        if os.path.islink(target):
            raise ValueError("refuse:destination is a symlink; remove it yourself")
        if os.path.lexists(target):
            if not a.replace:
                raise ValueError("refuse:%s exists; rerun with --replace to move it aside (it is never deleted)" % target)
            # Outside the skills folder, so the agent does not load two copies.
            r["moved_aside"] = os.path.join(os.path.dirname(os.path.dirname(target)),
                                            NAME + ".old-" + time.strftime("%Y%m%d-%H%M%S"))
            if os.path.lexists(r["moved_aside"]):
                raise ValueError("refuse:%s already exists" % r["moved_aside"])
        if not a.dry_run:
            os.makedirs(os.path.dirname(target), exist_ok=True)
            if r["moved_aside"]:
                os.rename(target, r["moved_aside"])
            tmp = target + ".partial-%d" % os.getpid()
            try:
                for f in r["files"]:
                    dst = os.path.join(tmp, *f.split("/"))
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copyfile(os.path.join(HERE, f), dst)
                    if sha(dst) != sha(os.path.join(HERE, f)):
                        raise OSError("copy check failed for %s" % f)
                os.rename(tmp, target)
            finally:
                if os.path.isdir(tmp):
                    shutil.rmtree(tmp, ignore_errors=True)
    except ValueError as exc:
        msg = str(exc)
        r.update(ok=False, status="refused", error=msg.split(":", 1)[1]); code = 3
    except OSError as exc:
        r.update(ok=False, status="error", error="OS error: %s" % exc); code = 2
    r["exit_code"] = code
    if a.json:
        print(json.dumps(r, indent=1))
    elif r["ok"]:
        print("%s %d files to %s" % ("would copy" if a.dry_run else "installed", len(r["files"]), target))
        if r["moved_aside"]:
            print("previous copy %s %s" % ("would move to" if a.dry_run else "moved to", r["moved_aside"]))
        if not a.dry_run:
            print("check it: python3 %s" % os.path.join(target, "tests", "smoke_test.py"))
    else:
        print("%s: %s" % (r["status"], r["error"]), file=sys.stderr)
    return code


if __name__ == "__main__":
    sys.exit(main())
