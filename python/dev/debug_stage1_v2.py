"""Debug Stage 1 v2: tambahkan patching untuk trace DAFTAR PUSTAKA."""
import sys, subprocess, json, zipfile, os
sys.stdout.reconfigure(encoding='utf-8')

# Jalankan daftar_isi.py tapi dengan env variable debug
import importlib.util, os, re

# Import langsung untuk bisa inspect
sys.path.insert(0, r"D:\Freelaces\Server\Lib\site-packages")
sys.path.insert(0, r"D:\Freelaces\Server\htdocs\adk\python")

import daftar_isi as d

# Test _FRONT_MATTER_RE vs DAFTAR PUSTAKA
print("=== Test _FRONT_MATTER_RE ===")
tests = ["DAFTAR PUSTAKA", "DAFTAR ISI", "KATA PENGANTAR", "ABSTRAK", "LAMPIRAN"]
for t in tests:
    m = d._FRONT_MATTER_RE.match(t)
    print(f"  {t!r:35} -> {bool(m)}")

# Test get_para_level dengan mock para
class MockStyle:
    def __init__(self, name=None):
        self.name = name
        self.base_style = None

class MockPara:
    def __init__(self, text, style_name=None):
        self._text = text
        self._style = MockStyle(style_name) if style_name else MockStyle("Normal")
    @property
    def text(self):
        return self._text
    @property
    def style(self):
        return self._style

print("\n=== Test get_para_level ===")
test_paras = [
    ("DAFTAR PUSTAKA", None),
    ("DAFTAR ISI", None),
    ("BAB I PENDAHULUAN", None),
    ("1.1 Latar Belakang", None),
]
for txt, sn in test_paras:
    mp = MockPara(txt, sn)
    lvl = d.get_para_level(mp)
    print(f"  {txt!r:35} -> {lvl}")
