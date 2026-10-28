#!/bin/sh
# Özbekça-Yozuv-1.0.exe ni yig'ish (Mono mcs). Ishga tushirish: sh build.sh
cd "$(dirname "$0")"
mcs -nologo -target:winexe -platform:anycpu -optimize+ -out:"../Özbekça-Yozuv-1.0.exe" -win32icon:tray.ico \
  -resource:tray.ico,tray.ico -r:System.Windows.Forms.dll -r:System.Drawing.dll -r:Accessibility.dll \
  Yozuv.cs Engine.cs
