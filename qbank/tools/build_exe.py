#!/usr/bin/env python3
"""Đóng gói QBank thành một file chạy: dist/QBank.exe trên Windows, dist/QBank trên Linux/macOS.

Chạy từ thư mục qbank/:  python3 tools/build_exe.py
Cần: pip install pyinstaller python-docx
Kết quả: dist/QBank(.exe). Dữ liệu (data/, templates/) được tạo cạnh file khi chạy lần đầu.
"""
import os
import subprocess
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SEP = ";" if os.name == "nt" else ":"  # dấu phân tách --add-data theo hệ điều hành

cmd = [
    sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onefile",
    "--name", "QBank",
    "--add-data", f"static{SEP}static",
    "--add-data", f"templates{SEP}templates",
    "--collect-data", "pykakasi",  # từ điển chữ Hán (.db) để tự thêm furigana
    "app.py",
]
subprocess.run(cmd, cwd=BASE, check=True)
print("Xong:", os.path.join(BASE, "dist", "QBank.exe" if os.name == "nt" else "QBank"))
