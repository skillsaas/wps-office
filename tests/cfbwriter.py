"""Tiny OLE2 compound file writer used only to build test fixtures.

Writes version 3 files (512-byte sectors) with a mini stream for streams under
4096 bytes. Sibling entries are chained through the right pointer (a valid,
if unbalanced, tree). Cross-read by olefile in the optional smoke test.
"""
import struct

ENDC, FREE, FATS = 0xFFFFFFFE, 0xFFFFFFFF, 0xFFFFFFFD
NOS = 0xFFFFFFFF


def _pad(b, n):
    return b + b"\x00" * ((-len(b)) % n)


def build(items, mutate=None):
    """items: list of (path, bytes) for streams or (path, None) for storages."""
    nodes = [{"name": "Root Entry", "type": 5, "kids": [], "data": None}]
    index = {"": 0}
    for path, data in items:
        parts = path.split("/")
        for i in range(1, len(parts) + 1):
            p = "/".join(parts[:i])
            if p in index:
                continue
            last = i == len(parts)
            node = {"name": parts[i - 1], "type": 2 if (last and data is not None) else 1,
                    "kids": [], "data": data if last else None}
            index[p] = len(nodes)
            nodes.append(node)
            nodes[index["/".join(parts[:i - 1])]]["kids"].append(index[p])
    mini, minifat = b"", []
    big = []  # (node, bytes)
    for n in nodes[1:]:
        if n["type"] != 2:
            continue
        d = n["data"]
        if len(d) < 4096:
            n["start"] = len(mini) // 64 if d else ENDC
            cnt = (len(d) + 63) // 64
            for k in range(cnt):
                minifat.append(n["start"] + k + 1 if k < cnt - 1 else ENDC)
            mini += _pad(d, 64)
        else:
            big.append((n, d))
    # sector plan: [big streams][ministream][minifat][directory]
    plan, fat_chain = [], []
    def alloc(data):
        start = len(plan)
        chunks = [data[i:i + 512] for i in range(0, len(data), 512)] or []
        for i, c in enumerate(chunks):
            plan.append(_pad(c, 512))
            fat_chain.append(start + i + 1 if i < len(chunks) - 1 else ENDC)
        return start if chunks else ENDC
    for n, d in big:
        n["start"] = alloc(d)
    nodes[0]["start"] = alloc(mini) if mini else ENDC
    nodes[0]["size"] = len(mini)
    mf_raw = b"".join(struct.pack("<I", x) for x in minifat)
    mf_raw = mf_raw + b"\xff" * ((-len(mf_raw)) % 512)
    mfstart = alloc(mf_raw) if minifat else ENDC
    nmf = len(mf_raw) // 512 if minifat else 0
    # directory entries
    for idx, n in enumerate(nodes):
        kids = n["kids"]
        n["child"] = kids[0] if kids else NOS
        for a, b in zip(kids, kids[1:]):
            nodes[a]["right"] = b
    raw = b""
    for n in nodes:
        name = n["name"].encode("utf-16-le") + b"\x00\x00"
        e = name.ljust(64, b"\x00") + struct.pack("<HBB", len(name), n["type"], 1)
        e += struct.pack("<III", NOS, n.get("right", NOS), n["child"])
        e += b"\x00" * 16 + b"\x00" * 4 + b"\x00" * 16
        size = n.get("size", len(n["data"]) if n["data"] is not None else 0)
        start = n.get("start", ENDC if n["type"] != 2 else ENDC)
        if n["type"] == 1:
            start, size = 0, 0
        e += struct.pack("<IQ", start, size)
        raw += e
    raw = _pad(raw, 512)
    dirstart = alloc(raw)
    # FAT sectors go at the end
    ncontent = len(plan)
    nfat = 1
    while (ncontent + nfat) > nfat * 128:
        nfat += 1
    fat = list(fat_chain) + [FATS] * nfat
    fat += [FREE] * (nfat * 128 - len(fat))
    if mutate:
        mutate(fat, locals())
    fatbytes = b"".join(struct.pack("<I", x) for x in fat)
    difat = [ncontent + i for i in range(nfat)] + [FREE] * (109 - nfat)
    hdr = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 16
    hdr += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\x00" * 6
    hdr += struct.pack("<IIIIIIIII", 0, nfat, dirstart, 0, 4096, mfstart, nmf, ENDC, 0)
    hdr += b"".join(struct.pack("<I", x) for x in difat)
    assert len(hdr) == 512, len(hdr)
    return hdr + b"".join(plan) + fatbytes


def summary_stream(props, fmtid=b"\xe0\x85\x9f\xf2\xf9\x4f\x68\x10\xab\x91\x08\x00\x2b\x27\xb3\xd9"):
    """A property set stream with codepage 65001 and string properties {pid: str}."""
    vals = [(1, struct.pack("<Ih", 2, -535) + b"\x00\x00")]  # codepage 65001 as signed int16
    for pid, s in props.items():
        b = s.encode("utf-8") + b"\x00"
        b = _pad(b, 4)
        vals.append((pid, struct.pack("<II", 0x1E, len(b)) + b))
    off = 8 + 8 * len(vals)
    table, body = b"", b""
    for pid, v in vals:
        table += struct.pack("<II", pid, off + len(body))
        body += v
    sec = struct.pack("<II", 8 + len(table) + len(body), len(vals)) + table + body
    head = struct.pack("<HHI", 0xFFFE, 0, 0x00020006) + b"\x00" * 16 + struct.pack("<I", 1) + fmtid + struct.pack("<I", 48)
    return head + sec
