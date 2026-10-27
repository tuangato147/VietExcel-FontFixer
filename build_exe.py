# -*- coding: utf-8 -*-
"""
Build Script for PyInstaller
Compiles VietExcel FontFixer into a standalone .exe
"""

import os
import subprocess
import sys

# Ensure stdout handles UTF-8 on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def build():
    print("=== Bat dau dong goi VietExcel FontFixer thanh .exe ===")

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--windowed",
        "--name", "VietExcel_FontFixer",
        "--collect-all", "customtkinter",
        "--hidden-import", "xlrd",
        "--hidden-import", "openpyxl",
        "app_gui.py"
    ]

    print(f"Chạy lệnh: {' '.join(cmd)}")
    result = subprocess.run(cmd)

    if result.returncode == 0:
        print("\n=======================================================")
        print("[SUCCESS] DONG GOI THANH CONG!")
        exe_path = os.path.abspath(os.path.join("dist", "VietExcel_FontFixer.exe"))
        print(f"File thuc thi tai: {exe_path}")
        print("=======================================================")
    else:
        print("\n[ERROR] Dong goi that bai voi ma loi:", result.returncode)

if __name__ == '__main__':
    build()
