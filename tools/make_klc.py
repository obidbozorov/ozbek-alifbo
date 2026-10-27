#!/usr/bin/env python3
"""«Özbekça (yangi alifbo)» klaviatura joylashuvi (.klc) — Microsoft Keyboard Layout Creator uchun.
Asos: AQSh (QWERTY). AltGr (o'ng Alt) + O/G/S/C -> ö ğ ş ç (Shift bilan katta), AltGr + ' -> ʼ (tutuq), Shift+AltGr+' -> ʻ.
Ishga tushirish: python3 tools/make_klc.py"""
import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "5-Klaviatura", "OzbekYangi.klc")

def h(c):                       # belgi -> KLC ko'rinishi
    if c is None: return "-1"
    if c.isalnum() and ord(c) < 128: return c
    return "%04x" % ord(c)

# (scancode, VK, base, shift, altgr, shift+altgr)
rows = [
    ("29", "OEM_3", "`", "~"), ("02", "1", "1", "!"), ("03", "2", "2", "@"), ("04", "3", "3", "#"),
    ("05", "4", "4", "$"), ("06", "5", "5", "%"), ("07", "6", "6", "^"), ("08", "7", "7", "&"),
    ("09", "8", "8", "*"), ("0a", "9", "9", "("), ("0b", "0", "0", ")"), ("0c", "OEM_MINUS", "-", "_"),
    ("0d", "OEM_PLUS", "=", "+"),
]
letters = [("10","Q"),("11","W"),("12","E"),("13","R"),("14","T"),("15","Y"),("16","U"),("17","I"),("18","O"),("19","P"),
           ("1e","A"),("1f","S"),("20","D"),("21","F"),("22","G"),("23","H"),("24","J"),("25","K"),("26","L"),
           ("2c","Z"),("2d","X"),("2e","C"),("2f","V"),("30","B"),("31","N"),("32","M")]
punct = [("1a","OEM_4","[","{","\x1b"),("1b","OEM_6","]","}","\x1d"),("27","OEM_1",";",":",None),
         ("28","OEM_7","'",'"',None),("2b","OEM_5","\\","|","\x1c"),("33","OEM_COMMA",",","<",None),
         ("34","OEM_PERIOD",".",">",None),("35","OEM_2","/","?",None),("56","OEM_102","\\","|","\x1c")]
NEW = {"O": ("ö", "Ö"), "G": ("ğ", "Ğ"), "S": ("ş", "Ş"), "C": ("ç", "Ç")}

L = []
for sc, vk, b, s in rows:
    L.append(f"{sc}\t{vk}\t\t0\t{h(b)}\t{h(s)}\t-1\t-1\t-1")
for sc, vk in letters:
    lo, up = vk.lower(), vk
    if vk in NEW:
        a, A = NEW[vk]
        # 5 = CapsLock AltGr qatlamiga ham ta'sir qiladi (CapsLock + AltGr+O -> Ö)
        L.append(f"{sc}\t{vk}\t\t5\t{lo}\t{up}\t-1\t{h(a)}\t{h(A)}\t\t// {a} {A}")
    else:
        L.append(f"{sc}\t{vk}\t\t1\t{lo}\t{up}\t-1\t-1\t-1")
for sc, vk, b, s, ctl in punct:
    ag, sag = (("\u02bc", "\u02bb") if vk == "OEM_7" else (None, None))
    note = "\t\t// AltGr+' -> tutuq (ʼ), Shift+AltGr+' -> ʻ" if ag else ""
    L.append(f"{sc}\t{vk}\t\t0\t{h(b)}\t{h(s)}\t{h(ctl)}\t{h(ag)}\t{h(sag)}{note}")
L.append("39\tSPACE\t\t0\t0020\t0020\t0020\t-1\t-1")
L.append("53\tDECIMAL\t\t0\t002e\t002e\t-1\t-1\t-1")

keyname = """01\tEsc
0e\tBackspace
0f\tTab
1c\tEnter
1d\tCtrl
2a\tShift
36\t"Right Shift"
37\t"Num *"
38\tAlt
39\tSpace
3a\t"Caps Lock"
3b\tF1
3c\tF2
3d\tF3
3e\tF4
3f\tF5
40\tF6
41\tF7
42\tF8
43\tF9
44\tF10
45\tPause
46\t"Scroll Lock"
47\t"Num 7"
48\t"Num 8"
49\t"Num 9"
4a\t"Num -"
4b\t"Num 4"
4c\t"Num 5"
4d\t"Num 6"
4e\t"Num +"
4f\t"Num 1"
50\t"Num 2"
51\t"Num 3"
52\t"Num 0"
53\t"Num Del"
54\t"Sys Req"
57\tF11
58\tF12
7c\tF13
7d\tF14
7e\tF15
7f\tF16
80\tF17
81\tF18
82\tF19
83\tF20
84\tF21
85\tF22
86\tF23
87\tF24"""
keyname_ext = """1c\t"Num Enter"
1d\t"Right Ctrl"
35\t"Num /"
37\t"Prnt Scrn"
38\t"Right Alt"
45\t"Num Lock"
46\tBreak
47\tHome
48\tUp
49\t"Page Up"
4b\tLeft
4d\tRight
4f\tEnd
50\tDown
51\t"Page Down"
52\tInsert
53\tDelete
54\t<00>
56\tHelp
5b\t"Left Windows"
5c\t"Right Windows"
5d\tApplication"""

text = f"""KBD\tOzbYangi\t"Özbekça (yangi alifbo)"

COPYRIGHT\t"(c) 2026 Bozorov Obidjon Norqobilovich"

COMPANY\t"Mirzo Ulugbek nomidagi Ozbekiston Milliy universiteti"

LOCALENAME\t"uz-Latn-UZ"

LOCALEID\t"00000443"

VERSION\t1.0

SHIFTSTATE

0\t//Column 4
1\t//Column 5 : Shft
2\t//Column 6 :       Ctrl
6\t//Column 7 :       Ctrl Alt
7\t//Column 8 : Shft  Ctrl Alt

LAYOUT\t\t;an extra '@' at the end is a dead key

//SC\tVK_\t\tCap\t0\t1\t2\t6\t7
//--\t----\t\t----\t----\t----\t----\t----\t----

""" + "\n".join(L) + f"""


KEYNAME

{keyname}

KEYNAME_EXT

{keyname_ext}

DESCRIPTIONS

0409\tÖzbekça (yangi alifbo)
0443\tÖzbekça (yangi alifbo)

LANGUAGENAMES

0409\tUzbek (Latin, Uzbekistan)
0443\tOzbek (lotin)

ENDKBD
"""
with open(OUT, "w", encoding="utf-16", newline="\r\n") as f:   # MSKLC: UTF-16 LE + BOM
    f.write(text)
print(OUT, len(L), "ta tugma")
