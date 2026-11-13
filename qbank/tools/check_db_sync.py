#!/usr/bin/env python3
"""Hook CSDL: cột bảng questions, bộ nhập/xuất và file mẫu nhập (docx, csv) phải luôn khớp nhau.

Chạy từ thư mục qbank/:  python3 tools/check_db_sync.py   (mã thoát 1 nếu lệch)

Kiểm tra:
1. Mọi cột của bảng questions (CREATE TABLE + ensure_column trong app.py) phải được khai báo: hoặc ứng với một
   cột trong file nhập (COLUMN_TO_HEADER), hoặc là cột app tự quản lý (NOT_IMPORTED). Thêm cột mới vào CSDL mà
   quên bộ nhập/mẫu thì bị chặn tại đây.
2. Mọi tiêu đề trong IMPORT_HEADERS phải được map_import_header nhận ra; EXPORT_HEADERS = Mã câu + IMPORT_HEADERS + Trạng thái.
3. Mẫu do app sinh ra (make_import_template, make_import_template_csv) có đúng tiêu đề và mỗi dòng ví dụ đủ cột.
4. File mẫu trong mau_nhap/ (docx, csv) có đúng tiêu đề hiện tại; lệch thì chạy tools/build_import_samples.py.
"""
import csv
import io
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE)
SAMPLE_DIR = os.path.join(BASE, "mau_nhap")

# cột CSDL của câu hỏi -> tiêu đề cột trong file nhập/xuất
COLUMN_TO_HEADER = {
    "code": "Mã câu", "chapter_id": "Mã chương", "content": "Nội dung", "options": "Nội dung", "answer": "Đáp án",
    "qtype": "Loại", "bloom": "Mức", "passage": "Đoạn văn", "solution": "Lời giải", "transcript": "Transcript",
    "translation": "Dịch", "clo": "CLO", "tags": "Thẻ", "word_limit": "Số từ", "status": "Trạng thái",
}
# cột do app tự quản lý, không nằm trong file nhập (hình/âm thanh, điểm, rubric, khóa nội bộ)
NOT_IMPORTED = {"id", "passage_id", "media", "points", "rubric"}


def question_columns(src):
    m = re.search(r"CREATE TABLE IF NOT EXISTS questions\((.*?)\)\"\"\"\)", src, re.S)
    cols = set()
    for ln in (m.group(1).splitlines() if m else []):
        w = ln.strip().split(" ")[0].strip(",")
        if w and w.isidentifier() and w.upper() not in ("PRIMARY", "FOREIGN", "UNIQUE"):
            cols.add(w)
    cols |= set(re.findall(r'ensure_column\("questions",\s*"(\w+)"', src))
    return cols


def sample_headers():
    """Tiêu đề của file mẫu trong mau_nhap/: {tên file: danh sách tiêu đề} (file thiếu thì không có trong kết quả)."""
    out = {}
    p = os.path.join(SAMPLE_DIR, "mau_nhap_cau_hoi.csv")
    if os.path.isfile(p):
        with open(p, encoding="utf-8-sig", newline="") as f:
            out["mau_nhap_cau_hoi.csv"] = next(csv.reader(f), [])
    p = os.path.join(SAMPLE_DIR, "mau_nhap_cau_hoi.docx")
    if os.path.isfile(p):
        from docx import Document
        d = Document(p)
        out["mau_nhap_cau_hoi.docx"] = [c.text for c in d.tables[0].rows[0].cells] if d.tables else []
    return out


def check():
    import app
    problems = []
    with open(os.path.join(BASE, "app.py"), encoding="utf-8") as f:
        src = f.read()
    for c in sorted(question_columns(src) - set(COLUMN_TO_HEADER) - NOT_IMPORTED):
        problems.append(f"Cột CSDL mới questions.{c} chưa có trong bộ nhập/mẫu: thêm vào IMPORT_HEADERS, IMPORT_COLUMNS, "
                        f"parse_rows, export_row, mẫu nhập, rồi khai báo trong COLUMN_TO_HEADER (hoặc NOT_IMPORTED nếu app tự quản lý)")
    for c, h in COLUMN_TO_HEADER.items():
        if c in question_columns(src) and h not in app.IMPORT_HEADERS + ["Mã câu", "Trạng thái"]:
            problems.append(f"Cột questions.{c} ứng với tiêu đề '{h}' nhưng IMPORT_HEADERS không có")
    cols, _ = app.map_import_header(app.IMPORT_HEADERS)
    mapped = set(cols.values())
    for i, h in enumerate(app.IMPORT_HEADERS):
        if i not in mapped:
            problems.append(f"Tiêu đề '{h}' trong IMPORT_HEADERS không được map_import_header nhận ra (thiếu trong IMPORT_COLUMNS)")
    if app.EXPORT_HEADERS != ["Mã câu"] + app.IMPORT_HEADERS + ["Trạng thái"]:
        problems.append("EXPORT_HEADERS không còn = Mã câu + IMPORT_HEADERS + Trạng thái")
    for i, ex in enumerate(app.IMPORT_EXAMPLES, 1):
        if len(ex) != len(app.IMPORT_HEADERS):
            problems.append(f"Dòng ví dụ {i} trong IMPORT_EXAMPLES có {len(ex)} ô, cần {len(app.IMPORT_HEADERS)}")
    from docx import Document
    gen_docx = [c.text for c in Document(io.BytesIO(app.make_import_template())).tables[0].rows[0].cells]
    gen_csv = next(csv.reader(io.StringIO(app.make_import_template_csv().decode("utf-8-sig"))))
    for name, hdr in (("make_import_template (docx)", gen_docx), ("make_import_template_csv", gen_csv)):
        if hdr != app.IMPORT_HEADERS:
            problems.append(f"{name} có tiêu đề khác IMPORT_HEADERS")
    samples = sample_headers()
    for name in ("mau_nhap_cau_hoi.docx", "mau_nhap_cau_hoi.csv"):
        if name not in samples:
            problems.append(f"Thiếu file mẫu mau_nhap/{name}: chạy python3 tools/build_import_samples.py")
        elif samples[name] != app.IMPORT_HEADERS:
            problems.append(f"File mẫu mau_nhap/{name} đã cũ (tiêu đề khác cột hiện tại): chạy python3 tools/build_import_samples.py")
    return problems


def main():
    problems = check()
    if problems:
        sys.stderr.write("CSDL CHƯA ĐỒNG BỘ VỚI BỘ NHẬP / FILE MẪU:\n" + "".join(f" - {p}\n" for p in problems))
        return 1
    print("CSDL, bộ nhập/xuất và file mẫu nhập đồng bộ.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
