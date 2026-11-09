#!/usr/bin/env python3
"""Export VBA source from an .xlsm (no Excel needed).

Usage: python export_vba.py <workbook.xlsm> <output_dir>

Reads xl/vbaProject.bin (an OLE compound file), decompresses each module
stream (MS-OVBA) and writes .bas / .cls / .frm text files in the same form
the VBA editor's File > Export produces (CRLF line endings, ANSI code page).
Standard library only.
"""
import io, os, re, struct, sys, zipfile

ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF


class CFB:
    def __init__(self, data):
        self.d = data
        if data[:8] != b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1":
            raise ValueError("not an OLE compound file")
        self.ss = 1 << struct.unpack_from("<H", data, 30)[0]
        self.mss = 1 << struct.unpack_from("<H", data, 32)[0]
        (nfat, dirstart, _, self.cutoff, minifat_start, nminifat,
         difat_start, ndifat) = struct.unpack_from("<IIIIIIII", data, 44)
        difat = [x for x in struct.unpack_from("<109I", data, 76) if x != FREESECT]
        s = difat_start
        seen = set()
        while s not in (ENDOFCHAIN, FREESECT) and s not in seen:
            seen.add(s)
            ents = struct.unpack_from("<%dI" % (self.ss // 4), self._sec(s))
            difat += [x for x in ents[:-1] if x != FREESECT]
            s = ents[-1]
        self.fat = []
        for s in difat[:nfat] if nfat else difat:
            self.fat += struct.unpack_from("<%dI" % (self.ss // 4), self._sec(s))
        dirdata = self._chain(dirstart)
        self.entries = []
        for i in range(0, len(dirdata) - 127, 128):
            e = dirdata[i:i + 128]
            nlen = struct.unpack_from("<H", e, 64)[0]
            name = e[:max(nlen - 2, 0)].decode("utf-16-le", "replace")
            typ = e[66]
            left, right, child = struct.unpack_from("<III", e, 68)
            start = struct.unpack_from("<I", e, 116)[0]
            size = struct.unpack_from("<Q", e, 120)[0]
            if self.ss == 512:
                size &= 0xFFFFFFFF
            self.entries.append(dict(name=name, type=typ, left=left, right=right,
                                     child=child, start=start, size=size))
        root = self.entries[0]
        self.ministream = self._chain(root["start"])[:root["size"]]
        mf = self._chain(minifat_start) if nminifat else b""
        self.minifat = struct.unpack_from("<%dI" % (len(mf) // 4), mf) if mf else ()
        self.paths = {}
        self._walk(root["child"], "")

    def _sec(self, n):
        o = (n + 1) * self.ss
        return self.d[o:o + self.ss]

    def _chain(self, s):
        out, seen = [], set()
        while s not in (ENDOFCHAIN, FREESECT) and s not in seen:
            seen.add(s)
            out.append(self._sec(s))
            s = self.fat[s]
        return b"".join(out)

    def _minichain(self, s):
        out, seen = [], set()
        while s not in (ENDOFCHAIN, FREESECT) and s not in seen:
            seen.add(s)
            out.append(self.ministream[s * self.mss:(s + 1) * self.mss])
            s = self.minifat[s]
        return b"".join(out)

    def _walk(self, idx, prefix):
        if idx == FREESECT or idx >= len(self.entries):
            return
        e = self.entries[idx]
        self._walk(e["left"], prefix)
        p = prefix + e["name"]
        self.paths[p.lower()] = e
        if e["type"] == 1:
            self._walk(e["child"], p + "/")
        self._walk(e["right"], prefix)

    def has(self, path):
        return path.lower() in self.paths

    def read(self, path):
        e = self.paths[path.lower()]
        if e["size"] < self.cutoff:
            return self._minichain(e["start"])[:e["size"]]
        return self._chain(e["start"])[:e["size"]]


def decompress(buf):
    """MS-OVBA 2.4.1 decompression."""
    if not buf or buf[0] != 1:
        raise ValueError("bad compressed container signature")
    out = bytearray()
    pos = 1
    n = len(buf)
    while pos + 2 <= n:
        hdr = struct.unpack_from("<H", buf, pos)[0]
        size = (hdr & 0x0FFF) + 3
        compressed = hdr >> 15
        end = min(pos + size, n)
        pos += 2
        if not compressed:
            out += buf[pos:pos + 4096]
            pos += 4096
            continue
        start = len(out)
        while pos < end:
            flags = buf[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                if not (flags >> bit) & 1:
                    out.append(buf[pos])
                    pos += 1
                else:
                    tok = struct.unpack_from("<H", buf, pos)[0]
                    pos += 2
                    diff = len(out) - start
                    bc = 4
                    while (1 << bc) < diff:
                        bc += 1
                    bc = min(bc, 12)
                    lmask = 0xFFFF >> bc
                    length = (tok & lmask) + 3
                    offset = (tok >> (16 - bc)) + 1
                    src = len(out) - offset
                    for i in range(length):
                        out.append(out[src + i])
    return bytes(out)


def parse_dir(d):
    codepage = 1252
    mods, cur = [], None
    pos = 0
    while pos + 6 <= len(d):
        rid, size = struct.unpack_from("<HI", d, pos)
        pos += 6
        if rid == 0x0009:          # PROJECTVERSION: size field is 4, data is 6
            pos += 6
            continue
        data = d[pos:pos + size]
        pos += size
        if rid == 0x0003:
            codepage = struct.unpack_from("<H", data)[0]
        elif rid == 0x0019:
            cur = dict(name=data, stream=None, offset=0)
            mods.append(cur)
        elif rid == 0x0047 and cur is not None:
            cur["uname"] = data.decode("utf-16-le", "replace")
        elif rid == 0x001A and cur is not None:
            cur["stream"] = data
        elif rid == 0x0032 and cur is not None:
            cur["ustream"] = data.decode("utf-16-le", "replace")
        elif rid == 0x0031 and cur is not None:
            cur["offset"] = struct.unpack_from("<I", data)[0]
        elif rid == 0x002B:
            cur = None
        elif rid == 0x0010:
            break
    return codepage, mods


def py_codec(cp):
    return {65001: "utf-8", 10000: "mac_roman"}.get(cp, "cp%d" % cp)


CLASS_HEADER = ('VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n')


def export(xlsm, outdir):
    with zipfile.ZipFile(xlsm) as z:
        name = next(n for n in z.namelist() if n.lower().endswith("vbaproject.bin"))
        cfb = CFB(z.read(name))
    codepage, mods = parse_dir(decompress(cfb.read("VBA/dir")))
    codec = py_codec(codepage)
    kinds = {}
    if cfb.has("PROJECT"):
        for line in cfb.read("PROJECT").decode(codec, "replace").splitlines():
            m = re.match(r"(Module|Class|Document|BaseClass)=([^/\r\n]+)", line)
            if m:
                kinds[m.group(2).strip().lower()] = m.group(1)
    os.makedirs(outdir, exist_ok=True)
    rows = []
    for m in mods:
        mname = m.get("uname") or m["name"].decode(codec, "replace")
        sname = m.get("ustream") or (m["stream"] or m["name"]).decode(codec, "replace")
        raw = cfb.read("VBA/" + sname)
        src = decompress(raw[m["offset"]:]) if len(raw) > m["offset"] else b""
        kind = kinds.get(mname.lower(), "Module")
        text = src.decode(codec, "replace")
        text = re.sub(r"\r\n|\r|\n", "\r\n", text)
        if text and not text.endswith("\r\n"):
            text += "\r\n"
        if kind == "Module":
            ext = ".bas"
        elif kind == "BaseClass":
            ext = ".frm"
        else:
            ext = ".cls"
            text = CLASS_HEADER + text
        code_lines = sum(1 for l in text.split("\r\n")
                         if l.strip() and not l.startswith("Attribute ")
                         and not l.startswith(("VERSION ", "BEGIN", "END", "  MultiUse")))
        fn = mname + ext
        with open(os.path.join(outdir, fn), "wb") as f:
            f.write(text.encode(codec, "replace"))
        rows.append((fn, kind, code_lines))
    return codepage, rows


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    cp, rows = export(sys.argv[1], sys.argv[2])
    print("code page", cp)
    for fn, kind, n in rows:
        print("%-40s %-10s %5d lines" % (fn, kind, n))
