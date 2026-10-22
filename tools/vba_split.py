#!/usr/bin/env python3
"""OzbekAlifbo.bas ni tahlil qiladi: qaysi protseduralar Wordga bog'liq, qaysilari sof (Excel/PowerPoint uchun umumiy)."""
import re, sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "1-VBA", "OzbekAlifbo.bas")
WORD = re.compile(r"\b(Selection|ActiveDocument|Documents|StoryRanges|KeyBindings|FindKey|CustomizationContext|NormalTemplate|"
                  r"MacroContainer|UndoRecord|BuildKeyCode|wd[A-Z]\w*|IRibbonControl|IRibbonUI|Options\.|ScreenUpdating|StatusBar|"
                  r"\.Range\b|Range\b|Paragraphs|ActiveWindow|WordBasic|Dialogs|Application\.)")

def parse(text):
    lines = text.split("\r\n")
    head, procs, cur = [], [], None
    for ln in lines:
        m = re.match(r"^(Public |Private )?(Sub|Function|Property Get|Property Let) (\w+)", ln)
        if cur is None and m:
            cur = {"name": m.group(3), "lines": [ln], "public": (m.group(1) or "Public ").strip() == "Public"}
            # oldingi izoh qatorlari
            k = len(head)
            while procs == [] and False: pass
        elif cur is not None:
            cur["lines"].append(ln)
            if re.match(r"^End (Sub|Function|Property)\b", ln):
                procs.append(cur); cur = None
        else:
            if not procs: head.append(ln)
            else: procs[-1].setdefault("after", []).append(ln)
    return head, procs

def analyse():
    text = open(SRC, encoding="latin-1", newline="").read()
    head, procs = parse(text)
    names = {p["name"] for p in procs}
    for p in procs:
        body = "\r\n".join(l for l in p["lines"] if not l.strip().startswith("'"))
        body = re.sub(r'"[^"\r\n]*"', '""', body)
        p["word"] = bool(WORD.search(body)) or p["name"].startswith("RB_")
        p["calls"] = {n for n in names if n != p["name"] and re.search(r"\b%s\b" % n, body)}
    by = {p["name"]: p for p in procs}
    changed = True
    while changed:
        changed = False
        for p in procs:
            if not p["word"] and any(by[c]["word"] for c in p["calls"]):
                p["word"] = True; changed = True
    return head, procs

if __name__ == "__main__":
    head, procs = analyse()
    pure = [p["name"] for p in procs if not p["word"]]
    word = [p["name"] for p in procs if p["word"]]
    print("sof:", len(pure)); print(" ".join(pure)); print("\nword:", len(word)); print(" ".join(word))
