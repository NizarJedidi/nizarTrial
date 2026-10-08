"""Génère un xl/vbaProject.bin (MS-OVBA) à partir de modules VBA (.bas), sans Excel.

Usage : python3 make_vbaproject.py out.bin Module1.bas [Module2.bas ...]
Le nom de module est pris dans la ligne 'Attribute VB_Name = "..."' (sinon le nom du fichier).
"""
import re, struct, sys, uuid

SECTOR = 512
MINI = 64
CUTOFF = 4096
ENDOFCHAIN = 0xFFFFFFFE
FREESECT = 0xFFFFFFFF
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF
CODEPAGE = 1252


# ----------------------------------------------------------------- compression MS-OVBA 2.4.1
def ovba_compress(data: bytes) -> bytes:
    out = bytearray([0x01])
    pos = 0
    n = len(data)
    while pos < n:
        chunk_start = pos
        chunk_end = min(pos + 4096, n)
        body = bytearray()
        cur = chunk_start
        while cur < chunk_end:
            flag = 0
            tokens = bytearray()
            for bit in range(8):
                if cur >= chunk_end:
                    break
                diff = cur - chunk_start
                bitcount = 4
                while (1 << bitcount) < diff:
                    bitcount += 1
                length_mask = 0xFFFF >> bitcount
                max_len = length_mask + 3
                max_off = 1 << bitcount
                best_len, best_off = 0, 0
                if diff > 0:
                    lo = max(chunk_start, cur - max_off)
                    limit = min(max_len, chunk_end - cur)
                    for cand in range(cur - 1, lo - 1, -1):
                        k = 0
                        while k < limit and data[cand + k] == data[cur + k]:
                            k += 1
                        if k > best_len:
                            best_len, best_off = k, cur - cand
                            if k == limit:
                                break
                if best_len >= 3:
                    token = ((best_off - 1) << (16 - bitcount)) | (best_len - 3)
                    tokens += struct.pack("<H", token)
                    flag |= 1 << bit
                    cur += best_len
                else:
                    tokens.append(data[cur])
                    cur += 1
            body.append(flag)
            body += tokens
        if len(body) > 4096:
            raise ValueError("chunk incompressible (non gere)")
        header = ((len(body) - 1) & 0x0FFF) | 0x3000 | 0x8000
        out += struct.pack("<H", header) + body
        pos = chunk_end
    return bytes(out)


def ovba_decompress(data: bytes) -> bytes:
    """Décompresseur de contrôle (MS-OVBA 2.4.1.3.1)."""
    assert data[0] == 0x01
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        size = (header & 0x0FFF) + 3
        compressed = header & 0x8000
        chunk_end = pos + size
        pos += 2
        dstart = len(out)
        if not compressed:
            out += data[pos:pos + 4096]
            pos = chunk_end
            continue
        while pos < chunk_end:
            flag = data[pos]; pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if flag & (1 << bit):
                    token = struct.unpack_from("<H", data, pos)[0]; pos += 2
                    diff = len(out) - dstart
                    bitcount = 4
                    while (1 << bitcount) < diff:
                        bitcount += 1
                    length = (token & (0xFFFF >> bitcount)) + 3
                    offset = (token >> (16 - bitcount)) + 1
                    for _ in range(length):
                        out.append(out[len(out) - offset])
                else:
                    out.append(data[pos]); pos += 1
    return bytes(out)


# ----------------------------------------------------------------- flux dir / PROJECT / PROJECTwm
def rec(rid, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def build_dir_stream(modules, project_name="VBAProject"):
    u16 = lambda s: s.encode("utf-16-le")
    ansi = lambda s: s.encode("cp1252")
    d = bytearray()
    d += rec(0x0001, struct.pack("<I", 3))                 # PROJECTSYSKIND (Win64)
    d += rec(0x0002, struct.pack("<I", 0x409))             # PROJECTLCID
    d += rec(0x0014, struct.pack("<I", 0x409))             # PROJECTLCIDINVOKE
    d += rec(0x0003, struct.pack("<H", CODEPAGE))          # PROJECTCODEPAGE
    d += rec(0x0004, ansi(project_name))                   # PROJECTNAME
    d += rec(0x0005, b"") + rec(0x0040, b"")               # PROJECTDOCSTRING
    d += rec(0x0006, b"") + rec(0x003D, b"")               # PROJECTHELPFILEPATH
    d += rec(0x0007, struct.pack("<I", 0))                 # PROJECTHELPCONTEXT
    d += rec(0x0008, struct.pack("<I", 0))                 # PROJECTLIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1627084032, 1)    # PROJECTVERSION (major, minor)
    d += rec(0x000C, b"") + rec(0x003C, b"")               # PROJECTCONSTANTS
    # référence stdole (OLE Automation)
    libid = ansi("*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation")
    d += rec(0x0016, ansi("stdole")) + rec(0x003E, u16("stdole"))
    d += rec(0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))
    d += rec(0x000F, struct.pack("<H", len(modules)))      # PROJECTMODULES
    d += rec(0x0013, struct.pack("<H", 0xFFFF))            # PROJECTCOOKIE
    for name, _src in modules:
        d += rec(0x0019, ansi(name))                       # MODULENAME
        d += rec(0x0047, u16(name))                        # MODULENAMEUNICODE
        d += rec(0x001A, ansi(name)) + rec(0x0032, u16(name))  # MODULESTREAMNAME
        d += rec(0x001C, b"") + rec(0x0048, b"")           # MODULEDOCSTRING
        d += rec(0x0031, struct.pack("<I", 0))             # MODULEOFFSET
        d += rec(0x001E, struct.pack("<I", 0))             # MODULEHELPCONTEXT
        d += rec(0x002C, struct.pack("<H", 0xFFFF))        # MODULECOOKIE
        d += struct.pack("<HI", 0x0021, 0)                 # MODULETYPE (procedural)
        d += struct.pack("<HI", 0x002B, 0)                 # Terminator
    d += struct.pack("<HI", 0x0010, 0)                     # dir Terminator
    return bytes(d)


def build_project_stream(modules, project_id):
    lines = ['ID="{%s}"' % project_id]
    lines += ["Module=%s" % name for name, _ in modules]
    lines += ['Name="VBAProject"', 'HelpContextID="0"', 'VersionCompatible32="393222000"', "",
              "[Host Extender Info]", "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000", "",
              "[Workspace]"]
    lines += ["%s=0, 0, 0, 0, C" % name for name, _ in modules]
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def build_projectwm_stream(modules):
    b = bytearray()
    for name, _ in modules:
        b += name.encode("cp1252") + b"\x00" + name.encode("utf-16-le") + b"\x00\x00"
    b += b"\x00\x00"
    return bytes(b)


# ----------------------------------------------------------------- conteneur OLE (CFB v3)
class _Entry:
    def __init__(self, name, typ, data=b"", children=None):
        self.name, self.typ, self.data = name, typ, data
        self.children = children or []
        self.sid = None
        self.left = self.right = self.child = NOSTREAM
        self.start = 0
        self.size = 0


def _sort_key(e):
    return (len(e.name), e.name.upper())


def _build_tree(entries):
    """BST équilibré ; retourne le sid racine."""
    if not entries:
        return NOSTREAM
    entries = sorted(entries, key=_sort_key)
    mid = len(entries) // 2
    root = entries[mid]
    root.left = _build_tree(entries[:mid])
    root.right = _build_tree(entries[mid + 1:])
    return root.sid


def build_ole(root_children):
    root = _Entry("Root Entry", 5, children=root_children)
    # numérotation des entrées (parcours en largeur)
    flat = []
    def walk(e):
        e.sid = len(flat); flat.append(e)
        for c in e.children:
            walk(c)
    walk(root)
    # mini stream et flux normaux
    mini = bytearray(); minifat = []
    big_sectors = []   # (entry, data)
    for e in flat:
        if e.typ != 2:
            continue
        e.size = len(e.data)
        if e.size == 0:
            e.start = ENDOFCHAIN
        elif e.size < CUTOFF:
            n = -(-e.size // MINI)
            e.start = len(minifat)
            for i in range(n):
                minifat.append(len(minifat) + 1 if i < n - 1 else ENDOFCHAIN)
            mini += e.data + b"\x00" * (n * MINI - e.size)
        else:
            big_sectors.append(e)
    for e in flat:
        if e.typ in (1, 5):
            e.child = _build_tree(e.children)
    root.size = len(mini)

    n_dir = max(1, -(-len(flat) * 128 // SECTOR))
    n_minifat = -(-len(minifat) * 4 // SECTOR) if minifat else 0
    n_mini = -(-len(mini) // SECTOR)
    n_big = sum(-(-e.size // SECTOR) for e in big_sectors)
    n_fat = 1
    while True:
        total = n_fat + n_dir + n_minifat + n_mini + n_big
        if n_fat * (SECTOR // 4) >= total:
            break
        n_fat += 1
    assert n_fat <= 109
    fat = []
    sectors = []
    def chain(count):
        start = len(fat) + 0
        for i in range(count):
            fat.append(len(fat) + 1 if i < count - 1 else ENDOFCHAIN)
        return start
    for _ in range(n_fat):
        fat.append(FATSECT)
    dir_start = chain(n_dir)
    minifat_start = chain(n_minifat) if n_minifat else ENDOFCHAIN
    mini_start = chain(n_mini) if n_mini else ENDOFCHAIN
    root.start = mini_start
    for e in big_sectors:
        e.start = chain(-(-e.size // SECTOR))
    while len(fat) < n_fat * (SECTOR // 4):
        fat.append(FREESECT)

    # en-tête
    difat = [i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header = struct.pack("<8s16sHHHHHHIIIIIIIIII", b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", b"\x00" * 16,
                         0x003E, 0x0003, 0xFFFE, 9, 6, 0, 0, 0, n_fat, dir_start, 0, CUTOFF,
                         minifat_start, n_minifat, ENDOFCHAIN, 0)
    header += struct.pack("<109I", *difat)
    assert len(header) == 512

    def dir_entry(e):
        if e is None:
            return b"\x00" * 128
        name = e.name.encode("utf-16-le") + b"\x00\x00"
        return (name + b"\x00" * (64 - len(name)) + struct.pack("<HBBIII", len(name), e.typ, 1, e.left, e.right, e.child)
                + b"\x00" * 16 + struct.pack("<IQQIQ", 0, 0, 0, e.start, e.size))
    dir_data = b"".join(dir_entry(e) for e in flat)
    pad_entries = n_dir * 4 - len(flat)
    dir_data += b"".join(
        (b"\x00" * 64 + struct.pack("<HBBIII", 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM) + b"\x00" * 16 + struct.pack("<IQQIQ", 0, 0, 0, 0, 0))
        for _ in range(pad_entries))

    out = bytearray(header)
    out += b"".join(struct.pack("<I", v) for v in fat)
    out += dir_data
    if n_minifat:
        mf = minifat + [FREESECT] * (n_minifat * (SECTOR // 4) - len(minifat))
        out += b"".join(struct.pack("<I", v) for v in mf)
    out += mini + b"\x00" * (n_mini * SECTOR - len(mini))
    for e in big_sectors:
        out += e.data + b"\x00" * (-(-e.size // SECTOR) * SECTOR - e.size)
    assert len(out) == (1 + total) * SECTOR, (len(out), total)
    return bytes(out)


def build_vbaproject(modules, project_id=None):
    """modules : liste de (nom, source_texte). Retourne les octets de vbaProject.bin."""
    project_id = project_id or str(uuid.uuid4()).upper()
    mods = []
    for name, src in modules:
        src = src.replace("\r\n", "\n").replace("\n", "\r\n")
        if not re.search(r'^Attribute VB_Name\s*=', src, re.M):
            src = 'Attribute VB_Name = "%s"\r\n' % name + src
        mods.append((name, src))
    dir_stream = ovba_compress(build_dir_stream(mods))
    vba_children = [_Entry("dir", 2, dir_stream), _Entry("_VBA_PROJECT", 2, b"\xCC\x61\xFF\xFF\x00\x00\x00")]
    for name, src in mods:
        comp = ovba_compress(src.encode("cp1252"))
        assert ovba_decompress(comp) == src.encode("cp1252")
        vba_children.append(_Entry(name, 2, comp))
    root_children = [_Entry("VBA", 1, children=vba_children),
                     _Entry("PROJECT", 2, build_project_stream(mods, project_id)),
                     _Entry("PROJECTwm", 2, build_projectwm_stream(mods))]
    return build_ole(root_children)


if __name__ == "__main__":
    out = sys.argv[1]
    modules = []
    for path in sys.argv[2:]:
        src = open(path, encoding="utf-8").read()
        m = re.search(r'^Attribute VB_Name\s*=\s*"([^"]+)"', src, re.M)
        name = m.group(1) if m else re.sub(r"\.bas$", "", path.split("/")[-1])
        modules.append((name, src))
    data = build_vbaproject(modules)
    open(out, "wb").write(data)
    print("ecrit", out, len(data), "octets ;", [m[0] for m in modules])
