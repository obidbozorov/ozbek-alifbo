#!/usr/bin/env python3
"""VBA modulini (OzbekAlifbo.bas) Wordsiz, LibreOffice Basic da ishga tushirib sinaydi.
Natijalar JS konvertori bilan solishtiriladi. Kerak: node, soffice (LibreOffice).
Ishga tushirish (loyiha ildizida): python3 tools/test_vba_lo.py"""
import os, subprocess, sys, tempfile, shutil, json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
T = os.path.join(ROOT, "tools", "vba_lo")
work = tempfile.mkdtemp(prefix="vba_lo_")
prof = os.path.join(work, "profile")

# 1. Kirish ma'lumotlari va JS natijalari
js = r"""
const C=require(process.argv[2]); const fs=require("fs"); const W=process.argv[3];
const load=f=>{const s=fs.readFileSync(f,"utf8"); const m=s.match(/const cases = (\[[\s\S]*?\n\]);/); return eval(m[1]);};
const t=process.argv[4];
const put=(n,rows,fn)=>fs.writeFileSync(W+"/in_"+n+".tsv", rows.map(r=>r+"\t"+fn(r)).join("\n"));
put("lat", load(t+"/converter.test.js").filter(c=>!(c[2]&&c[2].tutuq)).map(c=>c[0]), x=>C.convertText(x).text);
put("cyr", load(t+"/cyrillic.test.js").map(c=>c[0]), x=>C.cyrToLatin(x).text);
put("l2c", load(t+"/latin2cyr.test.js").map(c=>c[0]), x=>C.latToCyr(x).text);
put("cyr2", ["Майкрософтнинг маҳсулотлари","ГУГЛГА","21 октябрь, Сентябрда, АБЗАЦ, абзацлар","Айфон"], x=>C.cyrToLatin(x).text);
put("lat2", ["Ishoq va Ishoqov mushafni o’qidi","Schneier va Shannon shifrlari","Shor tuproq"], x=>C.convertText(x).text);
const words=[]; for (const e of C.DEFAULT_LAT_CYR) words.push(e.split("=")[1].replace("*",""));
put("l2cw", words.map(w=>C.cyrToLatin(w,{skipRussian:false}).text), x=>C.latToCyr(x).text);
put("n2o", load(t+"/new2old.test.js").map(c=>c[0]).concat(load(t+"/converter.test.js").filter(c=>!(c[2]&&c[2].tutuq)).map(c=>C.convertText(c[0]).text)), x=>C.newToOld(x).text);
put("l2ct", ["Ўзбекистон Республикаси Олий Мажлиси Сенати қонунни маъқуллади.".replace(/./g,c=>c)].map(x=>C.cyrToLatin(x).text), x=>C.latToCyr(x).text);
"""
open(os.path.join(work, "gen.js"), "w").write(js)
subprocess.run(["node", os.path.join(work, "gen.js"), os.path.join(ROOT, "2-Office-Add-in", "converter.js"), work,
                os.path.join(ROOT, "tests")], check=True)

# 2. LibreOffice profili va modul
subprocess.run(["soffice", "--headless", f"-env:UserInstallation=file://{prof}", "--terminate_after_init"],
               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=180)
subprocess.run([sys.executable, os.path.join(T, "build_module.py"), os.path.join(T, "harness.bas"),
                os.path.join(prof, "user", "basic", "Standard", "Module1.xba"), work], check=True, stdout=subprocess.DEVNULL)
subprocess.run(["soffice", "--headless", "--invisible", f"-env:UserInstallation=file://{prof}",
                "macro:///Standard.Module1.Main"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600)

# 3. Solishtirish
st = os.path.join(work, "status.txt")
if not os.path.exists(st):
    sys.exit("VBA modul LibreOffice da kompilyatsiya bo'lmadi (status.txt yo'q)")
status = open(st, encoding="utf-8").read()
if status != "OK":
    sys.exit("VBA xatosi: " + status)
total = bad = 0
for n in ["lat", "cyr", "l2c", "cyr2", "lat2", "l2cw", "l2ct", "n2o"]:
    rows = [l.split("\t") for l in open(os.path.join(work, f"in_{n}.tsv"), encoding="utf-8").read().split("\n") if l]
    outs = open(os.path.join(work, f"out_in_{n}.tsv"), encoding="utf-8").read().split("\n")
    for r, o in zip(rows, outs):
        total += 1
        if o != r[1]:
            bad += 1
            print(f"[{n}] {r[0]!r}\n   VBA: {o!r}\n   JS : {r[1]!r}")
shutil.rmtree(work, ignore_errors=True)
print(f"VBA (LibreOffice) natijalari JS bilan: {total - bad}/{total} mos")
sys.exit(1 if bad else 0)
