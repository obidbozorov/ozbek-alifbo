#!/usr/bin/env python3
"""macOS klaviatura joylashuvi «Özbekça (yangi alifbo)» (.keylayout + .bundle).
Asos: AQSh (QWERTY). Option + O/G/S/C -> ö ğ ş ç (Shift bilan katta), Option + ' -> ʼ, Shift+Option+' -> ʻ.
Option qatlamidagi boshqa tugmalar oddiy harfni beradi (ortiqcha belgilar yo'q).
Ishga tushirish: python3 tools/make_keylayout.py"""
import os, plistlib, zipfile, io
from xml.sax.saxutils import quoteattr

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "6-macOS")
NAME = "Özbekça (yangi alifbo)"

# ANSI klaviatura kodlari -> (oddiy, Shift)
KEYS = {
    0: ("a", "A"), 1: ("s", "S"), 2: ("d", "D"), 3: ("f", "F"), 4: ("h", "H"), 5: ("g", "G"), 6: ("z", "Z"),
    7: ("x", "X"), 8: ("c", "C"), 9: ("v", "V"), 10: ("§", "±"), 11: ("b", "B"), 12: ("q", "Q"), 13: ("w", "W"),
    14: ("e", "E"), 15: ("r", "R"), 16: ("y", "Y"), 17: ("t", "T"), 18: ("1", "!"), 19: ("2", "@"), 20: ("3", "#"),
    21: ("4", "$"), 22: ("6", "^"), 23: ("5", "%"), 24: ("=", "+"), 25: ("9", "("), 26: ("7", "&"), 27: ("-", "_"),
    28: ("8", "*"), 29: ("0", ")"), 30: ("]", "}"), 31: ("o", "O"), 32: ("u", "U"), 33: ("[", "{"), 34: ("i", "I"),
    35: ("p", "P"), 37: ("l", "L"), 38: ("j", "J"), 39: ("'", '"'), 40: ("k", "K"), 41: (";", ":"), 42: ("\\", "|"),
    43: (",", "<"), 44: ("/", "?"), 45: ("n", "N"), 46: ("m", "M"), 47: (".", ">"), 50: ("`", "~"),
}
# barcha qatlamlarda bir xil: boshqaruv tugmalari va raqamli blok
FIXED = {
    36: "\r", 48: "\t", 49: " ", 51: "\x08", 52: "\x03", 53: "\x1b", 64: "\x10", 65: ".", 66: "\x1d", 67: "*",
    69: "+", 70: "\x1c", 71: "\x1b", 72: "\x1f", 75: "/", 76: "\x03", 77: "\x1e", 78: "-", 79: "\x10", 80: "\x10",
    81: "=", 82: "0", 83: "1", 84: "2", 85: "3", 86: "4", 87: "5", 88: "6", 89: "7", 91: "8", 92: "9",
    96: "\x10", 97: "\x10", 98: "\x10", 99: "\x10", 100: "\x10", 101: "\x10", 103: "\x10", 105: "\x10",
    106: "\x10", 107: "\x10", 109: "\x10", 111: "\x10", 113: "\x10", 114: "\x05", 115: "\x01", 116: "\x0b",
    117: "\x7f", 118: "\x10", 119: "\x04", 120: "\x10", 121: "\x0c", 122: "\x10", 123: "\x1c", 124: "\x1d",
    125: "\x1f", 126: "\x1e",
}
NEW = {31: ("ö", "Ö"), 5: ("ğ", "Ğ"), 1: ("ş", "Ş"), 8: ("ç", "Ç"), 39: ("ʼ", "ʻ")}


def esc(ch):
    if ord(ch) < 0x20 or ord(ch) == 0x7f or ch in "<>&\"'":
        return "&#x%04X;" % ord(ch)
    return ch


def keymap(index, layer):
    rows = []
    for code in sorted(set(KEYS) | set(FIXED)):
        if code in FIXED:
            out = FIXED[code]
        else:
            lo, up = KEYS[code]
            if layer == "base": out = lo
            elif layer == "shift": out = up
            elif layer == "caps": out = up if lo.isalpha() else lo
            elif layer == "opt": out = NEW[code][0] if code in NEW else lo
            elif layer == "shiftopt": out = NEW[code][1] if code in NEW else up
            elif layer == "capsopt":
                out = NEW[code][1] if code in NEW and code != 39 else (NEW[code][0] if code in NEW else (up if lo.isalpha() else lo))
            elif layer == "cmd": out = lo
            elif layer == "ctrl":
                out = chr(ord(lo) - 96) if len(lo) == 1 and "a" <= lo <= "z" else lo
        rows.append('\t\t\t<key code="%d" output="%s"/>' % (code, esc(out)))
    return '\t\t<keyMap index="%d">\n%s\n\t\t</keyMap>' % (index, "\n".join(rows))


layers = ["base", "shift", "caps", "opt", "shiftopt", "capsopt", "cmd", "ctrl"]
xml = """<?xml version="1.1" encoding="UTF-8"?>
<!DOCTYPE keyboard SYSTEM "file://localhost/System/Library/DTDs/KeyboardLayout.dtd">
<!-- %s — Option + O/G/S/C -> ö ğ ş ç. Muallif: Bozorov Obidjon Norqobilovich (MIT). tools/make_keylayout.py yasaydi. -->
<keyboard group="0" id="-23417" name=%s maxout="1">
	<layouts>
		<layout first="0" last="17" mapSet="ozb" modifiers="mods"/>
	</layouts>
	<modifierMap id="mods" defaultIndex="0">
		<keyMapSelect mapIndex="0"><modifier keys=""/></keyMapSelect>
		<keyMapSelect mapIndex="1"><modifier keys="anyShift caps?"/></keyMapSelect>
		<keyMapSelect mapIndex="2"><modifier keys="caps"/></keyMapSelect>
		<keyMapSelect mapIndex="3"><modifier keys="anyOption"/></keyMapSelect>
		<keyMapSelect mapIndex="4"><modifier keys="anyShift anyOption caps?"/></keyMapSelect>
		<keyMapSelect mapIndex="5"><modifier keys="caps anyOption"/></keyMapSelect>
		<keyMapSelect mapIndex="6"><modifier keys="command anyShift? caps? anyOption?"/></keyMapSelect>
		<keyMapSelect mapIndex="7"><modifier keys="anyControl anyShift? caps? anyOption? command?"/></keyMapSelect>
	</modifierMap>
	<keyMapSet id="ozb">
%s
	</keyMapSet>
</keyboard>
""" % (NAME, quoteattr(NAME), "\n".join(keymap(i, l) for i, l in enumerate(layers)))

kl_path = os.path.join(OUT, "Ozbekca.keylayout")
open(kl_path, "w", encoding="utf-8").write(xml)

# .bundle (Klaviatura joylashuvlari papkasiga butunligicha ko'chiriladi)
b = "Ozbekca.bundle/Contents/"
info = plistlib.dumps({
    "CFBundleIdentifier": "uz.nuu.ozbekca.keyboardlayout",
    "CFBundleName": "Özbekça",
    "CFBundleVersion": "1.0",
    "CFBundleShortVersionString": "1.0",
    "KLInfo_Özbekça (yangi alifbo)": {
        "TISInputSourceID": "uz.nuu.ozbekca.keyboardlayout.ozbekca",
        "TISIntendedLanguage": "uz-Latn",
    },
})
strings = '"Özbekça (yangi alifbo)" = "Özbekça (yangi alifbo)";\n'.encode("utf-16")
zbuf = io.BytesIO()
with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr(b + "Info.plist", info)
    z.writestr(b + "Resources/Özbekça (yangi alifbo).keylayout", xml.encode("utf-8"))
    z.writestr(b + "Resources/English.lproj/InfoPlist.strings", strings)
    z.write(os.path.join(OUT, "README.md"), "README.md")
open(os.path.join(ROOT, "dist", "Ozbekca-macOS.zip"), "wb").write(zbuf.getvalue())
print(kl_path, "va dist/Ozbekca-macOS.zip tayyor")
