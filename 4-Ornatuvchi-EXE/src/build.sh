#!/bin/sh
# Özbekça-1.2.exe ni yig'ish (Mono mcs yoki Windows dagi csc bilan). Ishga tushirish: sh build.sh
cd "$(dirname "$0")"
RES=""
for f in icons/*.png; do n=$(basename "$f"); RES="$RES -resource:$f,icon_$n"; done
mcs -nologo -target:winexe -platform:anycpu -optimize+ -out:../Özbekça-1.2.exe -win32icon:icon.ico \
  -resource:../../1-VBA/OzbekAlifbo.bas,OzbekAlifbo.bas -resource:customUI14.xml,customUI14.xml -resource:klaviatura/OzbYangi_amd64.msi,OzbYangi_amd64.msi -resource:klaviatura/OzbYangi_i386.msi,OzbYangi_i386.msi -resource:klaviatura/amd64/OzbYangi.dll,kbd_amd64_OzbYangi.dll -resource:klaviatura/wow64/OzbYangi.dll,kbd_wow64_OzbYangi.dll -resource:klaviatura/i386/OzbYangi.dll,kbd_i386_OzbYangi.dll -resource:customUI14_excel.xml,customUI14_excel.xml -resource:customUI14_powerpoint.xml,customUI14_powerpoint.xml -resource:../../1-VBA/OzbekAlifbo_Excel.bas,OzbekAlifbo_Excel.bas -resource:../../1-VBA/OzbekAlifbo_PowerPoint.bas,OzbekAlifbo_PowerPoint.bas $RES \
  -r:System.Windows.Forms.dll -r:System.Drawing.dll -r:System.IO.Compression.dll \
  -r:System.IO.Compression.FileSystem.dll -r:System.Xml.Linq.dll -r:System.Xml.dll Ornatuvchi.cs
