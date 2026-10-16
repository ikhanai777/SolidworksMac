#!/usr/bin/env python3
"""
Build the SOLIDWORKS VBA macro files in dist/ from the sources in src/
(currently dist/SafeRename.swp; AddPartProperties is shipped as a plain .bas).

A .swp file is an OLE compound file (MS-CFB) holding a VBA project (MS-OVBA)
inside an "APC" host storage written by SOLIDWORKS:

    apc/
      The VBA Project/
        Host Project Item Names/...      SOLIDWORKS host data (copied verbatim)
        Host Project Items               "
        VBA Project Data                 "
        VBA Project Signature            "
        _VBA_Project/
          PROJECT, PROJECTwm             project text streams
          <form>/                        UserForm designer (CompObj, VBFrame, f, o)
          VBA/
            _VBA_PROJECT                 7-byte header, no p-code (see below)
            dir                          compressed module/reference directory
            ThisLibrary, <module>,       compressed module source
            <form>

The project is written source-only: _VBA_PROJECT has Version 0xFFFF and no
performance cache, as MS-OVBA 2.3.4.1 prescribes for writers. VBA therefore
compiles the modules from source the first time SOLIDWORKS opens the macro.

The host streams and the project references (stdole, SldWorks, SwConst,
MSForms) were taken from a macro saved by SOLIDWORKS: codestack.net
"FreezeRebuild.swp" (https://github.com/xarial/codestack, MIT licence).

Usage:  python3 tools/build_swp.py                 build every macro into dist/
        python3 tools/build_swp.py NAME [out.swp]  build one macro (SafeRename)
"""
import os
import struct
import sys
import uuid

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")

# One entry per macro: module = project name, form, form caption, design size of
# the form in points (the code sets the real size at run time), and a fixed GUID
# namespace so that every build of the same sources is byte-identical.
MACROS = {
    "SafeRename": dict(form="frmSafeRename", caption="Safe Rename", size=(618, 512),
                       ns="6b7f3a52-0c1e-4b8e-9a51-5afe7e4a3e11"),
}

PROJECT_NAME = MODULE_NAME = FORM_NAME = FORM_CAPTION = None
PROJECT_ID = FORM_CLSID = FORM_TYPELIB = None
FORM_W_PT = FORM_H_PT = None


def configure(name):
    """Select the macro that the functions below build."""
    global PROJECT_NAME, MODULE_NAME, FORM_NAME, FORM_CAPTION
    global PROJECT_ID, FORM_CLSID, FORM_TYPELIB, FORM_W_PT, FORM_H_PT
    m = MACROS[name]
    PROJECT_NAME = MODULE_NAME = name
    FORM_NAME, FORM_CAPTION = m["form"], m["caption"]
    FORM_W_PT, FORM_H_PT = m["size"]
    ns = uuid.UUID(m["ns"])
    PROJECT_ID = "{%s}" % str(uuid.uuid5(ns, "project")).upper()
    FORM_CLSID = "{%s}" % str(uuid.uuid5(ns, "form-clsid")).upper()
    FORM_TYPELIB = "{%s}" % str(uuid.uuid5(ns, "form-typelib")).upper()

# ---------------------------------------------------------------------------
# Data copied from a SOLIDWORKS-saved macro
# ---------------------------------------------------------------------------
HOST_STREAMS = {
    "Host Project Item Names/Host Project Item List Data":
        "3601040000000000000040010400000002000000",
    "Host Project Item Names/Project_Data_CurVer":
        "2c010400000001000000",
    "Host Project Item Names/ThisLibrary/Control List Data/Control List Data":
        "d2000400000000000000dc000400000000000000e6000400000000000000f0000400000001000000",
    "Host Project Item Names/ThisLibrary/Control List Data/Control_Data_CurVer":
        "c8000400000001000000",
    "Host Project Item Names/ThisLibrary/Project Item Data":
        "4a010400000001000000",
    "Host Project Item Names/ThisLibrary/ReqControl List Data/ReqControl List Data":
        "62020400000000000000",
    "Host Project Items":
        "f4010400000001000000167b01000000c003701908021600000054006800690073004c00690062007200"
        "6100720079002a7b01000000000470191c020400000001000000357b01000000400470191e024c000000"
        "7b00380030003100410032003000300034002d0041003600360043002d0031003100440033002d004100"
        "3800420044002d003000300030003000380036003100450042004200440036007d00377b010000008004"
        "701920024c0000007b00380030003100410030003000300030002d0041003600360043002d0031003100"
        "440033002d0041003800420044002d003000300030003000380036003100450042004200440036007d00"
        "397b01000000c004701922020200000019003b7b010000000005701924020200000000003d7b01000000"
        "4005701926020400000000000000",
    "VBA Project Data":
        "900104000000010000009a010400000001000000a4010400000026000000ae014c0000007b0038003000"
        "3100410030003000300030002d0041003600360043002d0031003100440033002d004100380042004400"
        "2d003000300030003000380036003100450042004200440036007d00b8010400000038000000c2017000"
        "000043003a005c00550073006500720073005c0061007200740065006d002e0074006100740075007200"
        "650076006900630068002e00430049005c0044006f0077006e006c006f006100640073005c0046007200"
        "650065007a006500520065006200750069006c0064002e00730077007000cc0104000000572d8707",
    "VBA Project Signature":
        "cd010400000000000000",
}

# MS Forms 2.0 Form CompObj stream
FORM_COMPOBJ = (
    "0100feff030a0000ffffffff00000000000000000000000000000000190000004d6963726f736f667420"
    "466f726d7320322e3020466f726d0010000000456d626564646564204f626a6563740000000000f439b2"
    "71000000000000000000000000")

# dir stream REFERENCE records: stdole, SldWorks, SwConst, MSForms
REFERENCES = (
    "1600060000007374646f6c653e000c0000007300740064006f006c0065000d00680000005e0000002a5c"
    "477b30303032303433302d303030302d303030302d433030302d3030303030303030303034367d23322e"
    "30233023433a5c57696e646f77735c73797374656d33325c7374646f6c65322e746c62234f4c45204175"
    "746f6d6174696f6e000000000000160008000000536c64576f726b733e001000000053006c0064005700"
    "6f0072006b0073000d0092000000880000002a5c477b38334133334433312d323743352d313143452d42"
    "4644342d3030343030353133424235377d2331392e30233023433a5c50726f6772616d2046696c65735c"
    "534f4c4944574f524b5320436f72705c534f4c4944574f524b53202832295c736c64776f726b732e746c"
    "6223536c64576f726b7320323031372054797065204c6962726172790000000000001600070000005377"
    "436f6e73743e000e0000005300770043006f006e00730074000d009c000000920000002a5c477b343638"
    "37463335392d353544302d344344332d423643462d3245423432433131463938397d2331392e30233023"
    "433a5c50726f6772616d2046696c65735c534f4c4944574f524b5320436f72705c534f4c4944574f524b"
    "53202832295c7377636f6e73742e746c6223534f4c4944574f524b53203230313720436f6e7374616e74"
    "2074797065206c6962726172790000000000001600070000004d53466f726d733e000e0000004d005300"
    "46006f0072006d00730033006f0000002a5c477b30443435324545312d453038462d313031412d383532"
    "452d3032363038433444304242347d23322e30233023433a5c57696e646f77735c73797374656d33325c"
    "464d32302e444c4c234d6963726f736f667420466f726d7320322e30204f626a656374204c6962726172"
    "792f003b000000310000002a5c477b30303030303030302d303030302d303030302d303030302d303030"
    "3030303030303030307d23302e30233023230000000000001600070000004d53466f726d733e000e0000"
    "004d00530046006f0072006d0073003000a80000008a0000002a5c477b44344337313231452d30373238"
    "2d343342342d393341332d4437413835314635303932437d23322e30233023433a5c55736572735c4152"
    "54454d547e312e43495c417070446174615c4c6f63616c5c54656d705c5642455c4d53466f726d732e65"
    "7864234d6963726f736f667420466f726d7320322e30204f626a656374204c6962726172790000000000"
    "00e12e450d8fe01a10852e02608c4d0bb401000000")

THISLIBRARY_SOURCE = (
    'Attribute VB_Name = "ThisLibrary"\r\n'
    'Attribute VB_Base = "0{801A2004-A66C-11D3-A8BD-0000861EBBD6}"\r\n'
    'Attribute VB_GlobalNameSpace = False\r\n'
    'Attribute VB_Creatable = False\r\n'
    'Attribute VB_PredeclaredId = True\r\n'
    'Attribute VB_Exposed = True\r\n'
    'Attribute VB_TemplateDerived = False\r\n'
    'Attribute VB_Customizable = True\r\n')

CODEPAGE = 1252


# ---------------------------------------------------------------------------
# MS-OVBA 2.4.1 compression
# ---------------------------------------------------------------------------
def _copy_token_help(pos):
    bit_count = max((pos - 1).bit_length(), 4)
    length_mask = 0xFFFF >> bit_count
    return bit_count, length_mask, length_mask + 3


def _compress_chunk(data):
    out = bytearray()
    pos = 0
    n = len(data)
    index = {}                      # 3-byte prefix -> list of positions
    while pos < n:
        flag_pos = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            best_len, best_off = 0, 0
            if pos >= 1 and pos + 3 <= n:
                _, _, max_len = _copy_token_help(pos)
                for cand in reversed(index.get(bytes(data[pos:pos + 3]), [])):
                    length = 0
                    limit = min(max_len, n - pos)
                    while length < limit and data[cand + length] == data[pos + length]:
                        length += 1
                    if length > best_len:
                        best_len, best_off = length, pos - cand
                        if length == limit:
                            break
            step = best_len if best_len >= 3 else 1
            if best_len >= 3:
                bit_count, _, _ = _copy_token_help(pos)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
            else:
                out.append(data[pos])
            for p in range(pos, pos + step):
                if p + 3 <= n:
                    index.setdefault(bytes(data[p:p + 3]), []).append(p)
            pos += step
        out[flag_pos] = flags
    return bytes(out)


def compress(data):
    out = bytearray(b"\x01")
    for start in range(0, len(data), 4096):
        chunk = data[start:start + 4096]
        comp = _compress_chunk(chunk)
        if len(comp) > 4096 and len(chunk) == 4096:
            out += struct.pack("<H", 0x3000 | (4096 + 2 - 3)) + chunk     # raw chunk
        else:
            out += struct.pack("<H", 0xB000 | (len(comp) + 2 - 3)) + comp
    return bytes(out)


# ---------------------------------------------------------------------------
# MS-OVBA 2.4.3 data encryption (CMG / DPB / GC)
# ---------------------------------------------------------------------------
def encrypt(project_id, data, seed):
    version_enc = 2 ^ seed
    proj_key = sum(ord(c) for c in project_id) & 0xFF
    proj_key_enc = proj_key ^ seed
    unenc1, enc1, enc2 = proj_key, proj_key_enc, version_enc
    out = bytearray([seed, version_enc, proj_key_enc])
    for _ in range((seed & 6) // 2):            # ignored bytes
        temp = 0
        b = temp ^ ((unenc1 + enc2) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, temp
    for byte in struct.pack("<I", len(data)) + bytes(data):
        b = byte ^ ((unenc1 + enc2) & 0xFF)
        out.append(b)
        enc2, enc1, unenc1 = enc1, b, byte
    return out.hex().upper()


# ---------------------------------------------------------------------------
# VBA project streams
# ---------------------------------------------------------------------------
def rec(rid, payload=b""):
    return struct.pack("<HI", rid, len(payload)) + payload


def u16(s):
    return s.encode("utf-16-le")


def mbcs(s):
    return s.encode("cp%d" % CODEPAGE)


def build_dir(modules):
    """modules: list of (name, is_procedural, is_private)"""
    d = bytearray()
    d += rec(0x01, struct.pack("<I", 3))            # SYSKIND: 64-bit Windows
    d += rec(0x02, struct.pack("<I", 0x409))        # LCID
    d += rec(0x14, struct.pack("<I", 0x409))        # LCIDINVOKE
    d += rec(0x03, struct.pack("<H", CODEPAGE))     # CODEPAGE
    d += rec(0x04, mbcs(PROJECT_NAME))              # NAME
    d += rec(0x05) + rec(0x40)                      # DOCSTRING (+unicode)
    d += rec(0x06) + rec(0x3D)                      # HELPFILEPATH 1/2
    d += rec(0x07, struct.pack("<I", 0))            # HELPCONTEXT
    d += rec(0x08, struct.pack("<I", 0))            # LIBFLAGS
    d += struct.pack("<HIIH", 0x09, 4, 0x5D10A12F, 0x1F)   # VERSION
    d += rec(0x0C) + rec(0x3C)                      # CONSTANTS (+unicode)
    d += bytes.fromhex(REFERENCES)
    d += rec(0x0F, struct.pack("<H", len(modules)))  # PROJECTMODULES
    d += rec(0x13, struct.pack("<H", 0xFFFF))        # PROJECTCOOKIE
    for name, procedural, private in modules:
        d += rec(0x19, mbcs(name)) + rec(0x47, u16(name))   # MODULENAME
        d += rec(0x1A, mbcs(name)) + rec(0x32, u16(name))   # MODULESTREAMNAME
        d += rec(0x1C) + rec(0x48)                          # MODULEDOCSTRING
        d += rec(0x31, struct.pack("<I", 0))                # MODULEOFFSET: source at 0
        d += rec(0x1E, struct.pack("<I", 0))                # MODULEHELPCONTEXT
        d += rec(0x2C, struct.pack("<H", 0xFFFF))           # MODULECOOKIE
        d += rec(0x21 if procedural else 0x22)              # MODULETYPE
        if private:
            d += rec(0x28)                                  # MODULEPRIVATE
        d += rec(0x2B)                                      # terminator
    d += rec(0x10)                                          # dir terminator
    return bytes(d)


def build_project_text():
    lines = [
        'ID="%s"' % PROJECT_ID,
        "Document=ThisLibrary/&H00000000",
        "Module=%s" % MODULE_NAME,
        "Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}",
        "BaseClass=%s" % FORM_NAME,
        'Name="%s"' % PROJECT_NAME,
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG="%s"' % encrypt(PROJECT_ID, b"\x00\x00\x00\x00", 0x51),  # not protected
        'DPB="%s"' % encrypt(PROJECT_ID, b"\x00", 0x2C),              # no password
        'GC="%s"' % encrypt(PROJECT_ID, b"\xff", 0x9E),               # visible
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
        "ThisLibrary=0, 0, 0, 0, C",
        "%s=30, 30, 1200, 700, " % MODULE_NAME,
        "%s=60, 60, 1200, 700, Z, 30, 30, 1100, 650, " % FORM_NAME,
        "",
    ]
    return mbcs("\r\n".join(lines))


def build_projectwm(names):
    out = bytearray()
    for n in names:
        out += mbcs(n) + b"\x00" + u16(n) + b"\x00\x00"
    return bytes(out + b"\x00\x00")


def form_streams():
    tw = lambda pt: int(round(pt * 20))
    him = lambda pt: int(round(pt * 2540 / 72))
    vbframe = (
        "VERSION 5.00\r\n"
        "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} %s \r\n"
        '   Caption         =   "%s"\r\n'
        "   ClientHeight    =   %d\r\n"
        "   ClientLeft      =   108\r\n"
        "   ClientTop       =   432\r\n"
        "   ClientWidth     =   %d\r\n"
        "   StartUpPosition =   1  'CenterOwner\r\n"
        "   TypeInfoVer     =   1\r\n"
        "End\r\n") % (FORM_NAME, FORM_CAPTION, tw(FORM_H_PT), tw(FORM_W_PT))
    # MS-OFORMS 2.2.10 FormControl without child controls.
    prop_mask = (1 << 3) | (1 << 10) | (1 << 11) | (1 << 26) | (1 << 27)
    data_block = struct.pack("<IIi", 0, 0, 32000)            # NextAvailableID, ShapeCookie, DrawBuffer
    extra_block = struct.pack("<iiii", him(FORM_W_PT), him(FORM_H_PT), 0, 0)  # DisplayedSize, LogicalSize
    body = struct.pack("<I", prop_mask) + data_block + extra_block
    f = struct.pack("<BBH", 0, 4, len(body)) + body
    f += struct.pack("<HII", 0, 0, 0)                        # site class info, sites, bytes
    return {
        "\x01CompObj": bytes.fromhex(FORM_COMPOBJ),
        "\x03VBFrame": mbcs(vbframe),
        "f": f,
        "o": b"",
    }


def read_source(path):
    with open(path, "r", encoding="ascii") as fh:      # VBA source must be ANSI; keep it ASCII
        text = fh.read()
    return text.replace("\r\n", "\n").replace("\n", "\r\n")


def build_tree():
    module_src = read_source(os.path.join(SRC, MODULE_NAME + ".bas"))
    if not module_src.startswith('Attribute VB_Name = "%s"' % MODULE_NAME):
        raise SystemExit("%s.bas must start with Attribute VB_Name" % MODULE_NAME)
    form_header = (
        'Attribute VB_Name = "%s"\r\n'
        'Attribute VB_Base = "0%s%s"\r\n'
        "Attribute VB_GlobalNameSpace = False\r\n"
        "Attribute VB_Creatable = False\r\n"
        "Attribute VB_PredeclaredId = True\r\n"
        "Attribute VB_Exposed = False\r\n"
        "Attribute VB_TemplateDerived = False\r\n"
        "Attribute VB_Customizable = False\r\n") % (FORM_NAME, FORM_CLSID, FORM_TYPELIB)
    form_src = form_header + read_source(os.path.join(SRC, FORM_NAME + ".vba"))

    modules = [("ThisLibrary", False, False), (MODULE_NAME, True, False), (FORM_NAME, False, True)]
    vba = {
        "_VBA_PROJECT": bytes.fromhex("cc61ffff000300"),
        "dir": compress(build_dir(modules)),
        "ThisLibrary": compress(mbcs(THISLIBRARY_SOURCE)),
        MODULE_NAME: compress(mbcs(module_src)),
        FORM_NAME: compress(mbcs(form_src)),
    }
    vba_project = {
        "PROJECT": build_project_text(),
        "PROJECTwm": build_projectwm([m[0] for m in modules]),
        FORM_NAME: form_streams(),
        "VBA": vba,
    }
    the_vba_project = {"_VBA_Project": vba_project}
    for path, hexdata in HOST_STREAMS.items():
        node = the_vba_project
        parts = path.split("/")
        for p in parts[:-1]:
            node = node.setdefault(p, {})
        node[parts[-1]] = bytes.fromhex(hexdata)
    return {"apc": {"The VBA Project": the_vba_project}}


# ---------------------------------------------------------------------------
# MS-CFB writer (version 3, 512-byte sectors)
# ---------------------------------------------------------------------------
SECTOR, MINI, CUTOFF = 512, 64, 4096
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
FILETIME_2026 = 134100000000000000          # fixed timestamp for reproducible builds


def _cfb_key(name):
    return (len(name), name.upper())


def write_cfb(tree, path):
    entries = []        # dicts: name, type, data, children(list of idx)

    def add(name, node):
        idx = len(entries)
        e = {"name": name, "children": []}
        entries.append(e)
        if isinstance(node, dict):
            e["type"], e["data"] = (5 if name == "Root Entry" else 1), b""
            for child_name in sorted(node, key=_cfb_key):
                e["children"].append(add(child_name, node[child_name]))
        else:
            e["type"], e["data"] = 2, node
        return idx

    add("Root Entry", tree)
    for e in entries:
        e.update(left=NOSTREAM, right=NOSTREAM, child=NOSTREAM, start=ENDOFCHAIN, size=0)

    def balance(ids):                       # balanced binary tree, all nodes black
        if not ids:
            return NOSTREAM
        mid = len(ids) // 2
        entries[ids[mid]]["left"] = balance(ids[:mid])
        entries[ids[mid]]["right"] = balance(ids[mid + 1:])
        return ids[mid]

    for e in entries:
        if e["type"] in (1, 5):
            e["child"] = balance(e["children"])

    # mini stream
    ministream, minifat = bytearray(), []
    big = []
    for e in entries:
        if e["type"] != 2:
            continue
        data = e["data"]
        e["size"] = len(data)
        if len(data) == 0:
            continue
        if len(data) < CUTOFF:
            first = len(ministream) // MINI
            count = (len(data) + MINI - 1) // MINI
            e["start"] = first
            minifat += [first + i + 1 for i in range(count - 1)] + [ENDOFCHAIN]
            ministream += data + b"\x00" * (count * MINI - len(data))
        else:
            big.append(e)

    def nsec(nbytes):
        return (nbytes + SECTOR - 1) // SECTOR

    dir_bytes = 128 * len(entries)
    n_dir = nsec(dir_bytes)
    n_minifat = nsec(4 * len(minifat))
    n_mini = nsec(len(ministream))
    n_big = sum(nsec(len(e["data"])) for e in big)
    n_other = n_dir + n_minifat + n_mini + n_big
    n_fat = 1
    while n_fat * (SECTOR // 4) < n_fat + n_other:
        n_fat += 1
    if n_fat > 109:
        raise SystemExit("file too large for header DIFAT")

    fat = [FATSECT] * n_fat
    sectors = []

    def place(data):
        first = n_fat + len(sectors)
        count = nsec(len(data))
        for i in range(count):
            sectors.append(data[i * SECTOR:(i + 1) * SECTOR].ljust(SECTOR, b"\x00"))
        fat.extend([first + i + 1 for i in range(count - 1)] + [ENDOFCHAIN])
        return first

    # directory placeholder first (filled after start sectors are known)
    dir_first = n_fat
    sectors.extend([b""] * n_dir)
    fat.extend([dir_first + i + 1 for i in range(n_dir - 1)] + [ENDOFCHAIN])
    minifat_first = place(b"".join(struct.pack("<I", x) for x in minifat)) if minifat else ENDOFCHAIN
    root = entries[0]
    if ministream:
        root["start"], root["size"] = place(bytes(ministream)), len(ministream)
    for e in big:
        e["start"] = place(e["data"])

    fat += [FREESECT] * (n_fat * (SECTOR // 4) - len(fat))

    d = bytearray()
    for e in entries:
        name = e["name"].encode("utf-16-le")
        if len(name) > 62:
            raise SystemExit("name too long: " + e["name"])
        ts = FILETIME_2026 if e["type"] == 1 else 0
        d += name.ljust(64, b"\x00")
        d += struct.pack("<HBB", len(name) + 2, e["type"], 1)
        d += struct.pack("<III", e["left"], e["right"], e["child"])
        d += b"\x00" * 16 + struct.pack("<I", 0)
        d += struct.pack("<QQ", ts, ts)
        d += struct.pack("<IQ", e["start"] if e["type"] != 1 else 0, e["size"])
    while len(d) % SECTOR:
        d += b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) \
             + b"\x00" * 36 + struct.pack("<IQ", 0, 0)
    for i in range(n_dir):
        sectors[i] = bytes(d[i * SECTOR:(i + 1) * SECTOR])

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1") + b"\x00" * 16
    header += struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\x00" * 6
    header += struct.pack("<IIIIIIIII", 0, n_fat, dir_first, 0, CUTOFF, minifat_first,
                          n_minifat, ENDOFCHAIN, 0)
    difat = list(range(n_fat)) + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", x) for x in difat)
    assert len(header) == 512

    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    with open(path, "wb") as fh:
        fh.write(header)
        fh.write(fat_bytes)
        for s in sectors:
            fh.write(s)


def build(name, out=None):
    configure(name)
    out = out or os.path.join(ROOT, "dist", name + ".swp")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    write_cfb(build_tree(), out)
    print("wrote", out, os.path.getsize(out), "bytes")


def main():
    args = sys.argv[1:]
    if not args:
        for name in MACROS:
            build(name)
    elif args[0] in MACROS:
        build(args[0], args[1] if len(args) > 1 else None)
    else:
        raise SystemExit("usage: build_swp.py [%s [out.swp]]" % "|".join(MACROS))


if __name__ == "__main__":
    main()
