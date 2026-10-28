#!/usr/bin/env python3
"""Onlayn konvertor sayti (docs/) uchun konvertorni yangilaydi: 2-Office-Add-in/converter.js -> docs/converter.js"""
import os, shutil
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
shutil.copy(os.path.join(ROOT, "2-Office-Add-in", "converter.js"), os.path.join(ROOT, "docs", "converter.js"))
print("docs/converter.js yangilandi")
