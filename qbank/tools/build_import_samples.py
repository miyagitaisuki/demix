#!/usr/bin/env python3
"""Dựng lại file mẫu nhập câu hỏi trong mau_nhap/ (docx + csv) từ app.py, đúng cột CSDL hiện tại.

Chạy từ thư mục qbank/:  python3 tools/build_import_samples.py
Hook CSDL (tools/check_db_sync.py) báo lỗi khi hai file này lệch với cột của app; chạy lệnh này để cập nhật."""
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
import app  # noqa: E402

OUT = os.path.join(BASE, "mau_nhap")


def main():
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, "mau_nhap_cau_hoi.docx"), "wb") as f:
        f.write(app.make_import_template())
    with open(os.path.join(OUT, "mau_nhap_cau_hoi.csv"), "wb") as f:
        f.write(app.make_import_template_csv())
    print("Đã ghi mau_nhap/mau_nhap_cau_hoi.docx và mau_nhap_cau_hoi.csv")


if __name__ == "__main__":
    main()
