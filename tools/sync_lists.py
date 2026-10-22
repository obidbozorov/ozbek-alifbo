#!/usr/bin/env python3
"""istisnolar/standart.json dagi standart ro'yxatlarni JS, C# va VBA kodiga yozadi.
Ishga tushirish (loyiha ildizida): python3 tools/sync_lists.py"""
import json, os, re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
data = json.load(open(os.path.join(ROOT, "istisnolar", "standart.json"), encoding="utf-8"))
LAT, CYR, LC = data["lotin"], data["kirill"], data.get("lotin_kirill", [])


def wrap(items, indent, width=104, quote=json.dumps):
    out, line = [], indent
    for x in items:
        t = quote(x) + ", "
        if len(line) + len(t) > width:
            out.append(line.rstrip())
            line = indent
        line += t
    out.append(line.rstrip().rstrip(","))
    return "\n".join(out)


def q(x):
    return json.dumps(x, ensure_ascii=False)


def sub(path, pattern, repl, enc="utf-8", newline=None):
    p = os.path.join(ROOT, path)
    s = open(p, encoding=enc).read().replace("\r\n", "\n")
    s2, n = re.subn(pattern, lambda m: repl, s, flags=re.S)
    if n != 1:
        raise SystemExit(f"{path}: naqsh topilmadi ({pattern[:40]}...)")
    open(p, "w", encoding=enc, newline=newline).write(s2)


# --- JS
sub("2-Office-Add-in/converter.js", r"  var DEFAULT_EXCEPTIONS = \[.*?\];",
    "  var DEFAULT_EXCEPTIONS = [\n" + wrap(LAT, "    ", quote=q) + "\n  ];")
sub("2-Office-Add-in/converter.js", r"  var DEFAULT_CYR_EXCEPTIONS = \[.*?\];",
    "  var DEFAULT_CYR_EXCEPTIONS = [\n" + wrap(CYR, "    ", quote=q) + "\n  ];")

sub("2-Office-Add-in/converter.js", r"  var DEFAULT_LAT_CYR = \[.*?\];",
    "  var DEFAULT_LAT_CYR = [\n" + wrap(LC, "    ", quote=q) + "\n  ];")

# --- C#
sub("3-VSTO/OzbekAlifboAddIn/Converter.cs", r"(public static readonly string\[\] DefaultExceptions =\n        \{\n).*?(\n        \};)",
    "public static readonly string[] DefaultExceptions =\n        {\n" + wrap(LAT, "            ", quote=q) + "\n        };")
sub("3-VSTO/OzbekAlifboAddIn/Converter.cs", r"(public static readonly string\[\] DefaultCyrExceptions =\n        \{\n).*?(\n        \};)",
    "public static readonly string[] DefaultCyrExceptions =\n        {\n" + wrap(CYR, "            ", quote=q) + "\n        };")


sub("3-VSTO/OzbekAlifboAddIn/Converter.cs", r"(public static readonly string\[\] DefaultLatCyr =\n        \{\n).*?(\n?        \};)",
    "public static readonly string[] DefaultLatCyr =\n        {\n" + wrap(LC, "            ", quote=q) + "\n        };")

# --- VBA (faqat ASCII: kirill harflar ChrW orqali)
def vba_str(txt):
    parts, buf = [], ""
    for ch in txt:
        if ord(ch) < 128:
            buf += ch
        else:
            if buf:
                parts.append('"' + buf + '"'); buf = ""
            parts.append("ChrW(&H%X)" % ord(ch))
    if buf:
        parts.append('"' + buf + '"')
    return " & ".join(parts)


chunks, line = [], ""
for w in LAT:
    if len(line) + len(w) + 1 > 80:
        chunks.append(line); line = ""
    line += w + ","
chunks.append(line.rstrip(","))
sub("1-VBA/OzbekAlifbo.bas", r'Private Const STANDART_ISTISNOLAR As String = _\n.*?"\n(?=Private Const MONO)',
    "Private Const STANDART_ISTISNOLAR As String = _\n" + " & _\n".join('    "%s"' % c for c in chunks) + "\n",
    enc="ascii", newline="\r\n")
sub("1-VBA/OzbekAlifbo.bas", r"(Private Sub InitCyrDefaults\(\)\n    mCyrDefCount = 0\n).*?(\nEnd Sub)",
    "Private Sub InitCyrDefaults()\n    mCyrDefCount = 0\n" + "\n".join("    AddCyrDefault " + vba_str(e) for e in CYR) + "\nEnd Sub",
    enc="ascii", newline="\r\n")
third = (len(LC) + 2) // 3
for n_, part in enumerate([LC[:third], LC[third:2 * third], LC[2 * third:]], 1):
    sub("1-VBA/OzbekAlifbo.bas", r"(Private Sub InitLatCyr%d\(\)\n).*?(End Sub)" % n_,
        "Private Sub InitLatCyr%d()\n" % n_ + "".join("    AddLatCyr " + vba_str(e) + "\n" for e in part) + "End Sub",
        enc="ascii", newline="\r\n")
print(f"Yangilandi: lotin {len(LAT)} ta, kirill {len(CYR)} ta, lotin->kirill lug'ati {len(LC)} ta")
