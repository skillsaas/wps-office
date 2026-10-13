"""Shared, hardened readers for the wps-office skill (standard library only).

Used by inspect_wps.py, cellimages.py and convert_wps.py. Not a CLI.

What is here:
- sniff(): identify a file by its bytes (OLE2/CFB, ZIP/OOXML, PDF, RTF, other),
  never by its extension.
- Cfb: a small read-only reader for OLE2 compound files (.wps .et .dps .doc .xls
  .ppt and encrypted OOXML). Every chain is bounded and loop-checked; sectors
  past the end of the file are reported, not fatal (WPS-saved .doc files with
  such FAT chains exist: apache/poi issue 1041).
- open_zip(): ZIP with entry, size and ratio caps; refuses duplicate names,
  case-colliding names, absolute or parent paths and encrypted members.
- parse_xml(): refuses DOCTYPE and ENTITY before parsing, caps size.
- clean(): strips control, bidi and zero-width characters and caps length, for
  any text taken from a file before it is printed.
"""
import io
import re
import struct
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET

VERSION = "0.1.0"

CFB_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ENDOFCHAIN, FREESECT, FATSECT, DIFSECT = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD, 0xFFFFFFFC
NOSTREAM = 0xFFFFFFFF


class Refusal(Exception):
    """A safety cap or rule was hit (exit code 3)."""


class FormatError(Exception):
    """The file is not what it claims, or is too damaged to read (exit code 2)."""


class Deadline:
    def __init__(self, seconds):
        self.end = time.monotonic() + float(seconds)

    def check(self):
        if time.monotonic() > self.end:
            raise Refusal("time limit reached (--timeout)")


_STRIP = re.compile("[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f"
                    "\u200b-\u200f\u202a-\u202e\u2060-\u2064\u2066-\u2069\ufeff]")


def clean(text, limit=200):
    """Text from a file, made safe to print: no controls, bidi or zero-width, capped."""
    if text is None:
        return None
    if isinstance(text, bytes):
        text = text.decode("utf-8", "replace")
    text = unicodedata.normalize("NFC", str(text))
    text = _STRIP.sub("", text).replace("\r", " ").replace("\n", " ").replace("\t", " ")
    if len(text) > limit:
        text = text[:limit] + "...[cut]"
    return text


def sniff(head):
    """Kind of file from its first bytes."""
    if head.startswith(CFB_MAGIC):
        return "cfb"
    if head.startswith(b"PK\x03\x04") or head.startswith(b"PK\x05\x06"):
        return "zip"
    if head.startswith(b"%PDF-"):
        return "pdf"
    if head.lstrip(b"\xef\xbb\xbf \r\n\t").startswith(b"{\\rtf"):
        return "rtf"
    if len(head) > 113 and head[112:114] == b"\x01\xfe":
        return "works-legacy"  # non-OLE Microsoft Works 1 to 3 (file(1) magic)
    return "other"


# ---------------------------------------------------------------- CFB reader
class Cfb:
    """Read-only OLE2 compound file reader with hard bounds."""

    def __init__(self, data, deadline, max_streams=4096):
        self.data = data
        self.deadline = deadline
        self.issues = []
        if len(data) < 512 or not data.startswith(CFB_MAGIC):
            raise FormatError("not an OLE2 compound file")
        (minor, major, bom, sshift, mshift) = struct.unpack_from("<HHHHH", data, 24)
        if bom != 0xFFFE or sshift not in (9, 12) or mshift != 6:
            raise FormatError("bad compound file header")
        self.ssize = 1 << sshift
        self.msize = 64
        (self.nfat, self.dirstart, _sig, self.mcutoff, self.mfatstart,
         self.nmfat, self.difstart, self.ndif) = struct.unpack_from("<IIIIIIII", data, 44)
        self.major = major
        self.nsect = max(0, (len(data) - self.ssize) // self.ssize) + 1
        self.past_eof = 0
        self._fat()
        self._minifat_loaded = False
        self.entries = []
        self._dir(max_streams)

    def _issue(self, what):
        if what not in self.issues:
            self.issues.append(what)

    def _sector(self, sid):
        off = self.ssize + sid * self.ssize
        if off >= len(self.data):
            self.past_eof += 1
            self._issue("sector-past-end-of-file")
            return b""
        return self.data[off:off + self.ssize]

    def _fat(self):
        difat = list(struct.unpack_from("<109I", self.data, 76))
        sid, seen = self.difstart, set()
        while sid not in (ENDOFCHAIN, FREESECT) and len(seen) < self.nsect:
            self.deadline.check()
            if sid in seen or sid >= self.nsect:
                self._issue("bad-difat-chain")
                break
            seen.add(sid)
            sec = self._sector(sid)
            n = self.ssize // 4
            vals = struct.unpack("<%dI" % n, sec.ljust(self.ssize, b"\xff"))
            difat.extend(vals[:-1])
            sid = vals[-1]
        fatsids = [s for s in difat if s < self.nsect][: max(self.nfat, 0) or len(difat)]
        raw = b"".join(self._sector(s) for s in fatsids)
        n = len(raw) // 4
        self.fat = list(struct.unpack("<%dI" % n, raw[: n * 4]))

    def chain(self, start, table, limit):
        out, sid, seen = [], start, set()
        while sid not in (ENDOFCHAIN, FREESECT, NOSTREAM):
            self.deadline.check()
            if sid in seen:
                self._issue("loop-in-sector-chain")
                break
            if sid >= len(table) or len(out) >= limit:
                if sid >= len(table):
                    self._issue("chain-points-outside-table")
                break
            seen.add(sid)
            out.append(sid)
            sid = table[sid]
        return out

    def _read_big(self, start, size):
        need = (size + self.ssize - 1) // self.ssize if size else self.nsect
        sids = self.chain(start, self.fat, min(need, self.nsect + 1))
        buf = b"".join(self._sector(s) for s in sids)
        if size and len(buf) < size:
            self._issue("stream-shorter-than-declared")
        return buf[:size] if size else buf

    def _dir(self, max_streams):
        raw = self._read_big(self.dirstart, 0)
        count = len(raw) // 128
        if count == 0:
            raise FormatError("compound file has no directory")
        ents = []
        for i in range(count):
            e = raw[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(0, min(nlen, 64) - 2)].decode("utf-16-le", "replace")
            etype = e[66]
            left, right, child = struct.unpack_from("<III", e, 68)
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0]
            if self.major == 3:
                size &= 0xFFFFFFFF
            ents.append({"name": name, "type": etype, "left": left, "right": right,
                         "child": child, "start": start, "size": size})
        self.raw_entries = ents
        if ents[0]["type"] != 5:
            raise FormatError("compound file root entry missing")
        root = ents[0]
        self.root = root
        # walk the red-black tree with a visited set and depth cap
        seen = set()
        stack = [(root["child"], "", 0)]
        while stack:
            self.deadline.check()
            idx, prefix, depth = stack.pop()
            if idx == NOSTREAM or idx >= len(ents):
                continue
            if idx in seen or depth > 32:
                self._issue("loop-in-directory")
                continue
            seen.add(idx)
            if len(self.entries) >= max_streams:
                raise Refusal("more than %d directory entries" % max_streams)
            e = ents[idx]
            path = prefix + e["name"]
            if e["type"] in (1, 2):
                self.entries.append({"path": path, "type": "storage" if e["type"] == 1 else "stream",
                                     "size": e["size"], "_idx": idx})
            stack.append((e["left"], prefix, depth + 1))
            stack.append((e["right"], prefix, depth + 1))
            if e["type"] == 1:
                stack.append((e["child"], path + "/", depth + 1))

    def _ministream(self):
        if not self._minifat_loaded:
            raw = self._read_big(self.mfatstart, 0) if self.nmfat else b""
            n = len(raw) // 4
            self.minifat = list(struct.unpack("<%dI" % n, raw[: n * 4]))
            self.mini = self._read_big(self.root["start"], self.root["size"])
            self._minifat_loaded = True

    def find(self, name):
        """Entries whose last path part equals name (case-insensitive)."""
        low = name.lower()
        return [e for e in self.entries if e["path"].split("/")[-1].lower() == low]

    def read(self, entry, cap):
        e = self.raw_entries[entry["_idx"]]
        size = min(e["size"], cap)
        if e["size"] < self.mcutoff:
            self._ministream()
            sids = self.chain(e["start"], self.minifat, (size + 63) // 64 + 1)
            buf = b"".join(self.mini[s * 64:(s + 1) * 64] for s in sids)
            if len(buf) < size:
                self._issue("stream-shorter-than-declared")
            return buf[:size]
        need = (size + self.ssize - 1) // self.ssize
        sids = self.chain(e["start"], self.fat, need)
        buf = b"".join(self._sector(s) for s in sids)
        if len(buf) < size:
            self._issue("stream-shorter-than-declared")
        return buf[:size]


# ---------------------------------------------------------------- property sets
def read_summary(blob):
    """Strings from a \\x05SummaryInformation or DocumentSummaryInformation stream.

    Returns {pid: text} for the first section. Only string, int and FILETIME
    values; everything else is skipped. Bounded by the blob length.
    """
    out = {}
    try:
        if len(blob) < 48:
            return out
        nsec = struct.unpack_from("<I", blob, 24)[0]
        if nsec < 1:
            return out
        off = struct.unpack_from("<I", blob, 44)[0]
        if off + 8 > len(blob):
            return out
        count = struct.unpack_from("<I", blob, off + 4)[0]
        codepage = 1252
        pairs = []
        for i in range(min(count, 256)):
            p = off + 8 + i * 8
            if p + 8 > len(blob):
                break
            pairs.append(struct.unpack_from("<II", blob, p))
        for pid, poff in pairs:
            q = off + poff
            if q + 8 > len(blob):
                continue
            vt = struct.unpack_from("<I", blob, q)[0] & 0xFFFF
            if pid == 1 and vt == 2:
                codepage = struct.unpack_from("<h", blob, q + 4)[0] & 0xFFFF
        for pid, poff in pairs:
            q = off + poff
            if q + 8 > len(blob):
                continue
            vt = struct.unpack_from("<I", blob, q)[0] & 0xFFFF
            if vt == 0x1E:
                n = struct.unpack_from("<I", blob, q + 4)[0]
                raw = blob[q + 8:q + 8 + min(n, 4096)].split(b"\x00")[0]
                enc = "utf-8" if codepage == 65001 else ("cp%d" % codepage)
                try:
                    out[pid] = raw.decode(enc, "replace")
                except LookupError:
                    out[pid] = raw.decode("latin-1")
            elif vt == 0x1F:
                n = struct.unpack_from("<I", blob, q + 4)[0]
                raw = blob[q + 8:q + 8 + min(n * 2, 8192)]
                out[pid] = raw.decode("utf-16-le", "replace").split("\x00")[0]
            elif vt == 3:
                out[pid] = struct.unpack_from("<i", blob, q + 4)[0]
    except struct.error:
        pass
    return out


# ---------------------------------------------------------------- ZIP and XML
def open_zip(data, max_entries=5000, max_total=512 * 1024 * 1024,
             max_member=256 * 1024 * 1024, max_ratio=200):
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as exc:
        raise FormatError("bad ZIP: %s" % exc)
    infos = zf.infolist()
    if len(infos) > max_entries:
        raise Refusal("ZIP has %d entries (cap %d)" % (len(infos), max_entries))
    seen, folded, total = set(), set(), 0
    for i in infos:
        n = i.filename
        if n in seen:
            raise Refusal("duplicate ZIP entry name: %s" % clean(n, 80))
        f = unicodedata.normalize("NFC", n).casefold()
        if f in folded:
            raise Refusal("ZIP entry names collide by case: %s" % clean(n, 80))
        seen.add(n)
        folded.add(f)
        if n.startswith(("/", "\\")) or re.match(r"^[A-Za-z]:", n) or ".." in re.split(r"[\\/]", n):
            raise Refusal("unsafe ZIP entry path: %s" % clean(n, 80))
        if i.flag_bits & 0x1:
            raise Refusal("encrypted ZIP member: %s" % clean(n, 80))
        if i.file_size > max_member:
            raise Refusal("ZIP member over %d bytes: %s" % (max_member, clean(n, 80)))
        if i.compress_size and i.file_size > 10 * 1024 * 1024 and i.file_size / i.compress_size > max_ratio:
            raise Refusal("ZIP member compression ratio over %d: %s" % (max_ratio, clean(n, 80)))
        total += i.file_size
        if total > max_total:
            raise Refusal("ZIP expands to more than %d bytes" % max_total)
    return zf


def read_member(zf, name, cap=64 * 1024 * 1024):
    with zf.open(name) as fh:
        data = fh.read(cap + 1)
    if len(data) > cap:
        raise Refusal("ZIP member larger than %d bytes when read: %s" % (cap, clean(name, 80)))
    return data


_DTD = re.compile(rb"<!\s*(DOCTYPE|ENTITY)", re.I)


def parse_xml(data):
    if _DTD.search(data[:65536]) or _DTD.search(data):
        raise Refusal("XML part contains DOCTYPE or ENTITY (refused, entity expansion risk)")
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise FormatError("XML parse error: %s" % exc)


def local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""
