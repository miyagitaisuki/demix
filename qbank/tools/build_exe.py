#!/usr/bin/env python3
"""Đóng gói QBank thành một file chạy: dist/QBank.exe trên Windows, dist/QBank trên Linux/macOS.

Chạy từ thư mục qbank/:  python3 tools/build_exe.py
Cần: pip install pyinstaller python-docx
Kết quả: dist/QBank(.exe) và dist/QBankMCP(.exe) (MCP server để Claude điều khiển app; luôn đi kèm). Dữ liệu (data/, templates/) được tạo cạnh file khi chạy lần đầu.
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
    "--add-data", f"VERSION{SEP}.",  # số phiên bản để nút Cập nhật so sánh với GitHub Releases
    "--collect-data", "pykakasi",  # từ điển chữ Hán (.db) để tự thêm furigana
    "app.py",
]
# MCP phải khớp app: lệch thì dừng, không đóng gói bản MCP cũ
subprocess.run([sys.executable, os.path.join("tools", "check_mcp_sync.py")], cwd=BASE, check=True)
subprocess.run(cmd, cwd=BASE, check=True)
subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--onefile", "--name", "QBankMCP", "mcp_server.py"],
               cwd=BASE, check=True)
ext = ".exe" if os.name == "nt" else ""
print("Xong:", os.path.join(BASE, "dist", "QBank" + ext), "và", os.path.join(BASE, "dist", "QBankMCP" + ext))
