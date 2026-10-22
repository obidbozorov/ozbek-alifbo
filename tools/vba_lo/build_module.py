"""OzbekAlifbo.bas dan Word ga bog'liq bo'lmagan qismini ajratib, LibreOffice Basic moduli yasaydi.
Foydalanish: build_module.py harness.bas chiqish.xba [ishchi_papka]"""
import re, sys, html
import os
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
src = open(os.path.join(ROOT, '1-VBA', 'OzbekAlifbo.bas'), encoding='ascii').read().replace('\r\n', '\n')
lines = src.split('\n')
# header (before first procedure)
first = next(i for i, l in enumerate(lines) if re.match(r'(Public |Private )?(Sub|Function) ', l))
header = lines[:first]
# drop Attribute line and #If..#End If block
out, skip = [], False
for l in header:
    if l.startswith('Attribute '): continue
    if l.startswith('#If'): skip = True
    if skip:
        if l.startswith('#End If'): skip = False
        continue
    out.append(l)
# split procedures
procs, cur = [], []
for l in lines[first:]:
    cur.append(l)
    if re.match(r'End (Sub|Function)\s*$', l):
        procs.append(cur); cur = []
BAD = ['Selection', 'ActiveDocument', 'Documents', 'KeyBindings', 'CustomizationContext', 'IRibbonControl',
       'Application.', 'CommandBar', 'As Range', 'Shell ', 'MacroContainer', 'NormalTemplate', 'GetKeyboardLayout',
       'ToUnicodeEx', 'BuildKeyCode', 'FindKey', 'CreateObject', 'KeyBinding', 'As Paragraph', 'As Document',
       'wdSelection', 'Environ$', 'Dir$']
kept, dropped = [], []
for p in procs:
    body = '\n'.join(p)
    name = next(re.match(r'(?:Public |Private )?(?:Sub|Function) (\w+)', l).group(1) for l in p if re.match(r'(?:Public |Private )?(?:Sub|Function) ', l))
    if name == 'ReadUtf8File':
        kept.append('Private Function ReadUtf8File(ByVal path As String) As String\n    ReadUtf8File = ""\nEnd Function'); continue
    if name in ('ExceptionsPath', 'CyrExceptionsPath'):
        kept.append(f'Private Function {name}() As String\n    {name} = "/nonexistent"\nEnd Function'); continue
    if any(b in body for b in BAD): dropped.append(name); continue
    kept.append(body)
harness = open(sys.argv[1], encoding='utf-8').read().replace('@DIR@', sys.argv[3] if len(sys.argv) > 3 else '')
code = 'Option VBASupport 1\n' + '\n'.join(out) + '\n' + '\n\n'.join(kept) + '\n\n' + harness
xml = ('<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
       '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">'
       + html.escape(code, quote=False) + '</script:module>')
open(sys.argv[2], 'w', encoding='utf-8').write(xml)
print('kept', len(kept), 'dropped', len(dropped), ':', ' '.join(dropped))
