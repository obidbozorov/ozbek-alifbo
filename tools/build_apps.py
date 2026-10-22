#!/usr/bin/env python3
"""Excel va PowerPoint uchun VBA modullarini yig'adi:
  OzbekAlifbo.bas dagi sof (Wordga bog'liq bo'lmagan) qism + 1-VBA/ilovalar/Umumiy.bas + Excel.bas / PowerPoint.bas
Natija: 1-VBA/OzbekAlifbo_Excel.bas, 1-VBA/OzbekAlifbo_PowerPoint.bas
Ishga tushirish (loyiha ildizida): python3 tools/build_apps.py"""
import os, re, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vba_split

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APPS = os.path.join(ROOT, "1-VBA", "ilovalar")
# Word tezkor tugmalari va yozishda o'girish — Excel/PowerPoint da kerak emas
SKIP = {"KeyDefault", "KeyFor", "KeyIds", "KeyWhat", "ParseKey", "AddDisabled", "RemoveDisabled", "LiveIsOn", "KeyToChar"}
# Wordga bog'liq deb belgilangan, lekin kodi umumiy bo'lgan menyular
COPY = ["YangiAlifbo_Istisnolar", "YangiAlifbo_KirillIstisnolar"]
LISTS = ["ShowExceptionList", "ShowCyrList"]


def read(p):
    return open(p, encoding="latin-1", newline="").read().replace("\r\n", "\n").replace("\n", "\r\n")


def to_list_view(lines):
    """ShowExceptionList/ShowCyrList: Word hujjati o'rniga Bloknot."""
    out = []
    skipping = False
    for ln in lines:
        s = ln.strip()
        if re.match(r"Set d = Documents\.Add", s):
            out.append("    ShowText t"); skipping = True; continue
        if skipping:
            if s.startswith("End Sub"):
                skipping = False; out.append(ln)
            continue
        ln = re.sub(r",\s*d As Document", "", ln)
        out.append(ln)
    return out


def build(app):
    head, procs = vba_split.analyse()
    host_names = set()
    extra = read(os.path.join(APPS, "Umumiy.bas")) + "\r\n" + read(os.path.join(APPS, app + ".bas"))
    for m in re.finditer(r"^(?:Public |Private )?(?:Sub|Function) (\w+)", extra, re.M):
        host_names.add(m.group(1))
    out = []
    for ln in head:
        if ln.startswith("'  YANGI O'ZBEK ALIFBOSI - Word uchun makros"):
            ln = "'  YANGI O'ZBEK ALIFBOSI - " + app + " uchun makros (avtomatik yig'ilgan: tools/build_apps.py)"
        out.append(ln)
    used = 0
    for p in procs:
        name = p["name"]
        if name in host_names or name in SKIP:
            continue
        if not p["word"] or name in COPY:
            body = p["lines"]
        elif name in LISTS:
            body = to_list_view(p["lines"])
        else:
            continue
        out.extend(body)
        out.extend(p.get("after", [""]) or [""])
        used += 1
    # ilova qismidagi modul darajasidagi o'zgaruvchilar (Private mX As ...) protseduralardan OLDIN turishi shart
    decl_re = re.compile(r"^(?:Private|Public|Dim) (?!Sub|Function|Declare|Const)\w+ As [\w.]+\s*$", re.M)
    decls = decl_re.findall(extra)
    extra = decl_re.sub("", extra)
    if decls:
        first = next(i for i, l in enumerate(out) if re.match(r"(Public |Private )?(Sub|Function) ", l))
        k = first
        while k > 0 and (out[k - 1].startswith("'") or not out[k - 1].strip()):
            k -= 1
        out[k:k] = ["' " + app + " uchun"] + [d.strip() for d in decls]
    text = "\r\n".join(out).rstrip() + "\r\n\r\n" + extra.strip() + "\r\n"
    dst = os.path.join(ROOT, "1-VBA", "OzbekAlifbo_" + app + ".bas")
    open(dst, "w", encoding="latin-1", newline="").write(text)
    long_lines = [i for i, l in enumerate(text.split("\r\n"), 1) if len(l) > 1023]
    bad = [c for c in text if ord(c) > 127]
    print(f"{os.path.basename(dst)}: {used} ta umumiy protsedura + {len(host_names)} ta {app} protsedurasi"
          + (f"  [UZUN QATOR: {long_lines}]" if long_lines else "") + (f"  [ASCII EMAS: {len(bad)}]" if bad else ""))
    return dst


if __name__ == "__main__":
    for a in ("Excel", "PowerPoint"):
        build(a)


def build_ribbons():
    """Word lentasidan Excel/PowerPoint lentalari: yozishda o'girish, tezkor tugma sozlamasi va
    o'ng tugma menyusi olib tashlanadi (bu ilovalarda yo'q)."""
    src = open(os.path.join(ROOT, "4-Ornatuvchi-EXE", "src", "customUI14.xml"), encoding="utf-8").read()
    s = re.sub(r"\s*<toggleButton id=\"tglYozishda\".*?/>", "", src, flags=re.S)
    s = re.sub(r"\s*<button id=\"btnTugmalar\".*?/>", "", s, flags=re.S)
    s = re.sub(r"\s*<contextMenus>.*?</contextMenus>", "", s, flags=re.S)
    s = s.replace("Ozbekca.dotm lentasi", "Excel/PowerPoint lentasi (tools/build_apps.py yasaydi)")
    for app in ("Excel", "PowerPoint"):
        dst = os.path.join(ROOT, "4-Ornatuvchi-EXE", "src", "customUI14_" + app.lower() + ".xml")
        open(dst, "w", encoding="utf-8").write(s)
        print(os.path.basename(dst), "tayyor")


if __name__ == "__main__":
    build_ribbons()
