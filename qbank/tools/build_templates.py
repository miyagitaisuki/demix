#!/usr/bin/env python3
"""Dựng mẫu chính thức từ file Word gốc của trường.

- mau_goc/de_goc_mau.docx      -> templates/de_thi_chuan.docx
- mau_goc/dap_an_goc_mau.docx  -> templates/dap_an_chuan.docx
Khung, tên trường, khoa, chân trang giữ nguyên. Chỉ thay các giá trị thay đổi bằng placeholder
và thay phần câu hỏi/lưới đáp án mẫu bằng điểm chèn.
Chạy: python3 tools/build_templates.py
"""
import copy
import os
import re
from docx import Document
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "mau_goc")
DST = os.path.join(BASE, "templates")


def _marker_para(body_parent_el, text):
    """Tạo đoạn mới (chỉ có chữ) dùng làm điểm chèn."""
    from docx.oxml import OxmlElement
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    p.append(r)
    return p


def build_exam():
    doc = Document(os.path.join(SRC, "de_goc_mau.docx"))
    hdr = doc.tables[0]
    left, right = hdr.rows[0].cells
    lp = left.paragraphs
    rp = right.paragraphs
    # Trái
    lp[4].runs[0].text = "ĐỀ SỐ: {{MA_DE}}"
    lp[5].runs[1].text = " {{SO_TRANG}} "
    # Phải
    rp[2].runs[1].text = "{{HOC_KY}}"
    rp[2].runs[3].text = " NĂM HỌC {{NAM_HOC}}"
    rp[2].runs[4].text = ""
    rp[3].runs[1].text = " {{TEN_HP}}"
    rp[4].runs[2].text = "{{MA_HP}}"
    rp[4].runs[5].text = " {{SO_TC}}"
    rp[5].runs[1].text = " {{KHOA_HOC}}"
    rp[6].runs[1].text = "{{THOI_GIAN}}"
    # Thân: giữ các phần tử đến đoạn có ngắt mục (chân trang), bỏ câu mẫu, chèn {{CAU_HOI}}
    body = doc.element.body
    kids = list(body)
    # Đoạn trống đầu tiên sau bảng tiêu đề mang ngắt mục (chân trang). Giữ đoạn đó, bỏ các phần còn lại.
    keep_until = next(i for i, el in enumerate(kids)
                      if i >= 1 and el.tag == qn("w:p") and el.find(".//" + qn("w:sectPr")) is not None)
    keep = {0, keep_until}
    for i, el in enumerate(kids):
        if i in keep or el.tag == qn("w:sectPr"):  # sectPr cuối thân tài liệu: giữ
            continue
        body.remove(el)
    # Xóa chữ còn sót trong đoạn mang ngắt mục (nếu có)
    for r in list(kids[keep_until].iter(qn("w:r"))):
        r.getparent().remove(r)
    kids[keep_until].addnext(_marker_para(body, "{{CAU_HOI}}"))
    doc.save(os.path.join(DST, "de_thi_chuan.docx"))


def build_key():
    doc = Document(os.path.join(SRC, "dap_an_goc_mau.docx"))
    # Thay giá trị trong bảng tiêu đề (bỏ qua các bảng lưới đáp án)
    grid_tables = [t for t in doc.element.body.iter(qn("w:tbl")) if _is_grid(t)]
    grid_ids = {id(t) for t in grid_tables}
    labels = {  # nhãn -> placeholder cho giá trị đứng sau nhãn
        "Đáp án đề số:": "{{MA_DE}}", "KỲ": "{{HOC_KY}}", "NĂM": "{{NAM_HOC}}",
        "Học phần:": "{{TEN_HP}}", "Mã HP": "{{MA_HP}}", "Số tín chỉ:": "{{SO_TC}}",
        "Khóa": "{{KHOA_HOC}}", "Thời gian làm bài:": "{{THOI_GIAN}}",
    }
    tokens = []
    for t in doc.element.body.iter(qn("w:t")):
        if any(id(a) in grid_ids for a in _ancestors(t)) or _in_alt(t):
            continue
        tokens.append(t)
    pending, skip_dash = None, False
    for t in tokens:
        txt = (t.text or "").strip()
        if not txt:
            continue
        if skip_dash and re.fullmatch(r"\s*-?\d{4}(-\d{4})?\s*", t.text or ""):
            t.text = ""  # phần năm học còn sót ("2025" "-2026"): đã được thay bằng {{NAM_HOC}}
            continue
        skip_dash = False
        if pending:
            raw = t.text or ""
            lead = raw[:len(raw) - len(raw.lstrip())]
            trail = raw[len(raw.rstrip()):]
            t.text = lead + pending + trail  # giữ khoảng trắng quanh giá trị
            skip_dash = pending == "{{NAM_HOC}}"
            pending = None
            continue
        if txt in labels:
            pending = labels[txt]
            continue
    # Điểm chèn lưới đáp án: đặt trước bảng lưới đầu tiên, rồi xóa các bảng lưới mẫu
    if grid_tables:
        first = grid_tables[0]
        marker = _marker_para(None, "{{BANG_DAP_AN}}")
        first.addprevious(marker)
        for gt in grid_tables:
            gt.getparent().remove(gt)
    doc.save(os.path.join(DST, "dap_an_chuan.docx"))


def _ancestors(el):
    out, p = [], el.getparent()
    while p is not None:
        out.append(p)
        p = p.getparent()
    return out


def _in_alt(el):
    return any(a.tag == "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent"
               for a in _ancestors(el))


def _is_grid(tbl):
    """Bảng lưới đáp án: ô đầu tiên của dòng đầu là chữ 'Câu'."""
    first_tc = tbl.find(".//" + qn("w:tc"))
    if first_tc is None:
        return False
    texts = "".join(x.text or "" for x in first_tc.iter(qn("w:t"))).strip()
    return texts == "Câu"


# Biến thể của đề thi chuẩn: cùng khung, tiêu đề và thông tin trường. Đổi tiêu đề loại đề,
# và (nếu có) thay dòng ghi chú về tài liệu ở cuối khối tiêu đề.
VARIANTS = {
    "kiem_tra_giua_ky.docx": ("ĐỀ KIỂM TRA GIỮA KỲ", None),
    "kiem_tra_nhanh.docx": ("ĐỀ KIỂM TRA NHANH", None),
    "bai_tap_ve_nha.docx": ("BÀI TẬP VỀ NHÀ", "(Sinh viên được sử dụng tài liệu khi làm bài)."),
}
# bài tập về nhà không tính thời gian phát đề: bỏ cụm "(Không kể thời gian phát đề)" ở dòng thời gian
NO_HANDOUT_TIME = {"bai_tap_ve_nha.docx"}


def build_variants():
    src = os.path.join(DST, "de_thi_chuan.docx")
    for name, (title, note) in VARIANTS.items():
        doc = Document(src)
        right = doc.tables[0].rows[0].cells[1]
        right.paragraphs[0].runs[0].text = title + " "
        if note:
            runs = right.paragraphs[7].runs
            runs[0].text = note
            for r in runs[1:]:
                r.text = ""
        if name in NO_HANDOUT_TIME:
            time_runs = right.paragraphs[6].runs
            for r in time_runs:
                if "Không kể" in r.text:
                    r.text = ""
        doc.save(os.path.join(DST, name))


def _set_para(p, text):
    """Đặt lại chữ của một đoạn: giữ run đầu tiên, bỏ các run sau."""
    runs = p.runs
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def build_answer_sheet():
    """Phiếu đáp án trắc nghiệm trên giấy theo Mẫu M2 của trường: khung tiêu đề và các lưới đáp án.
    Giữ bảng tiêu đề lồng (khoa, học phần, học kỳ...), thay giá trị bằng placeholder, và thay khối
    'Mã đề' + lưới mẫu bằng điểm chèn {{BANG_DAP_AN}}. Lưới 20 câu mỗi dòng do render_key_doc dựng."""
    from docx.table import Table
    doc = Document(os.path.join(SRC, "phieu_dap_an_goc_mau.docx"))
    body = doc.element.body
    kids = list(body)
    outer_el = [el for el in kids if el.tag == qn("w:tbl")][-1]  # bảng phiếu đáp án là bảng cuối cùng
    for el in kids:  # chỉ giữ bảng khung và sectPr; bỏ tiêu đề "Mẫu M2" và các mục hướng dẫn phía sau
        if el is not outer_el and el.tag != qn("w:sectPr"):
            body.remove(el)
    nested = Table(outer_el, doc).rows[0].cells[0].tables[0]
    left, right = nested.rows[0].cells
    for p in left.paragraphs:
        if p.text.startswith("Đáp án đề số"):
            _set_para(p, "Đáp án đề số: {{MA_DE}}")
    values = [
        ("ĐÀO TẠO", "ĐÀO TẠO TRÌNH ĐỘ ĐẠI HỌC HỆ CHÍNH QUY"),
        ("HỌC KỲ", "HỌC KỲ {{HOC_KY}} NĂM HỌC {{NAM_HOC}}"),
        ("Học phần", "Học phần: {{TEN_HP}}"),
        ("Mã HP", "Mã HP {{MA_HP}}          Số tín chỉ: {{SO_TC}}"),
        ("Khóa", "Khóa {{KHOA_HOC}}"),
        ("Thời gian", "Thời gian làm bài: {{THOI_GIAN}} phút (Không kể thời gian phát đề)"),
    ]
    for p in right.paragraphs:
        for prefix, new in values:
            if p.text.startswith(prefix):
                _set_para(p, new)
                break
    # ô chứa khối mã đề và lưới mẫu (hàng 3 của bảng lồng): thay bằng điểm chèn
    cell = nested.rows[2].cells[0]
    for child in list(cell._tc):
        if child.tag != qn("w:tcPr"):
            cell._tc.remove(child)
    cell.add_paragraph("Mã đề: {{MA_DE}}")
    cell.add_paragraph("{{BANG_DAP_AN}}")
    cell.add_paragraph("Ghi chú:  - Giảng viên ghi mã đề và điền đáp án A, B, C, D vào ô tương ứng.")
    cell.add_paragraph("- (1): Ghi cụ thể hình thức đào tạo (chính quy, thường xuyên)")
    cell.add_paragraph("- (2): Ghi địa điểm liên kết đào tạo đối với đào tạo thường xuyên.")
    doc.save(os.path.join(DST, "phieu_dap_an.docx"))



    build_exam()
    build_key()
    build_variants()
    build_answer_sheet()
    print("Đã dựng: templates/de_thi_chuan.docx, dap_an_chuan.docx, phieu_dap_an.docx và các biến thể đề")
