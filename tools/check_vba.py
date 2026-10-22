#!/usr/bin/env python3
"""OzbekAlifbo.bas ni Wordsiz tekshiradi: qator uzunligi, ASCII, e'lonlar joyi,
takroriy nomlar, blok juftliklari (If/For/Do/Select/With) va lenta callbacklari."""
import os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
bas = open(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "1-VBA", "OzbekAlifbo.bas"), encoding="ascii").read()
lines = bas.replace("\r\n", "\n").split("\n")
err = []


def code(ln):
    # izoh va satrlarni olib tashlash
    ln = re.sub(r'"[^"]*"', '""', ln)
    i = ln.find("'")
    return (ln[:i] if i >= 0 else ln).strip()


for n, ln in enumerate(lines, 1):
    if len(ln) > 1023:
        err.append(f"{n}: qator 1023 belgidan uzun")

first_proc = next(i for i, l in enumerate(lines) if re.match(r"(Public |Private )?(Sub|Function) ", l))
for i, l in enumerate(lines[first_proc:], first_proc + 1):
    if re.match(r"(Private|Public|Dim) m\w+ As ", l) and not l.startswith(" "):
        err.append(f"{i}: modul o'zgaruvchisi protseduralardan keyin e'lon qilingan")

names = re.findall(r"^(?:Public |Private )?(?:Sub|Function) (\w+)", bas, re.M)
for nm in set(names):
    if names.count(nm) > 1:
        err.append(f"takroriy nom: {nm}")

# blok juftliklari (har protsedura ichida)
stack, proc = [], None
joined, buf = [], ""
for l in lines:                       # davom qatorlarini birlashtirish
    c = code(l)
    if c.endswith(" _"):
        buf += c[:-2] + " "
        continue
    joined.append(buf + c); buf = ""
for n, c in enumerate(joined, 1):
    if not c:
        continue
    for part in re.split(r":(?=\s)", c) if not re.match(r"^\w+:$", c) else [""]:
        p = part.strip()
        if not p:
            continue
        m = re.match(r"(?:Public |Private )?(Sub|Function) (\w+)", p)
        if m:
            if stack: err.append(f"{proc}: yopilmagan bloklar {stack}")
            stack, proc = [], m.group(2); continue
        if re.match(r"End (Sub|Function)$", p):
            if stack: err.append(f"{proc}: yopilmagan bloklar {stack}")
            stack = []; continue
        if re.match(r"If .* Then$", p): stack.append("If")
        elif re.match(r"(ElseIf .* Then|Else)$", p): pass
        elif p == "End If":
            if not stack or stack.pop() != "If": err.append(f"{proc}: ortiqcha End If")
        elif re.match(r"For (Each )?\w+", p): stack.append("For")
        elif re.match(r"Next\b", p):
            if not stack or stack.pop() != "For": err.append(f"{proc}: ortiqcha Next")
        elif re.match(r"Do\b", p): stack.append("Do")
        elif re.match(r"Loop\b", p):
            if not stack or stack.pop() != "Do": err.append(f"{proc}: ortiqcha Loop")
        elif re.match(r"Select Case\b", p): stack.append("Select")
        elif p == "End Select":
            if not stack or stack.pop() != "Select": err.append(f"{proc}: ortiqcha End Select")
        elif re.match(r"With\b", p): stack.append("With")
        elif p == "End With":
            if not stack or stack.pop() != "With": err.append(f"{proc}: ortiqcha End With")
        if re.match(r"If .* Then .+", p):
            break          # bir qatorli If — qolgan qismi shu If ga tegishli

ui = os.path.join(ROOT, "4-Ornatuvchi-EXE", "src", "customUI14.xml")
for cb in set(re.findall(r'(?:onAction|getPressed)="(\w+)"', open(ui, encoding="utf-8").read())):
    if cb not in names:
        err.append(f"lenta callbacki yo'q: {cb}")

print("\n".join(err) if err else f"VBA tekshiruvi o'tdi: {len(names)} ta protsedura, {len(lines)} qator")
sys.exit(1 if err else 0)
