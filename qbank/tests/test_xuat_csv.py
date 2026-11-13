"""Xuất câu hỏi ra CSV cùng cột với file nhập; nhập lại file đã sửa cập nhật đúng câu (theo Mã câu)."""
import os
import sqlite3
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


@pytest.fixture
def bank(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "data" / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    os.makedirs(tmp_path / "data")
    app.init_storage()
    c1 = app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B01", "Bài 1")
    c2 = app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B02", "Bài 2")
    tid = app.q("SELECT id FROM textbooks WHERE name='Trainer 3'", fetch=True)[0]["id"]
    specs = [
        {"chapter_id": c1, "bloom": 1, "qtype": "TN", "content": "Từ nào đúng?\nDòng hai",
         "options": ["cat", "dog", "#cố định"], "answer": "B", "solution": "Giải thích", "clo": "CLO1", "tags": "ngữ pháp"},
        {"chapter_id": c1, "bloom": 2, "qtype": "TN", "content": "Thư viện mới.", "options": [],
         "answer": "A", "tags": "tfng", "passage": "Đoạn văn về thư viện.", "clo": "CLO2"},
        {"chapter_id": c1, "bloom": 2, "qtype": "TN", "content": "Đoạn 1", "options": ["Lợi ích", "Lịch sử", "Cách mượn"],
         "answer": "A", "tags": "noitieude", "passage": "Đoạn 1: thư viện yên tĩnh."},
        {"chapter_id": c1, "bloom": 2, "qtype": "DS", "content": "Câu đúng/sai:",
         "options": ["Tôi là sinh viên.", "Sách đọc."], "answer": "Đ S"},
        {"chapter_id": c2, "bloom": 1, "qtype": "DT", "content": "Thư viện có ___ tầng.", "options": [],
         "answer": "ba|3", "word_limit": 2},
        {"chapter_id": c2, "bloom": 3, "qtype": "TL", "content": "Viết về gia đình.", "options": [],
         "answer": "", "solution": "Tiêu chí: nội dung, ngữ pháp", "clo": "CLO3", "status": "nhap"},
    ]
    for spec in specs:
        app.insert_q(app.clean_question(dict(spec, status=spec.get("status", "da_duyet"))))
    return {"tid": tid, "c1": c1, "c2": c2}


def snapshot():
    with sqlite3.connect(app.DB_PATH) as conn:
        return conn.execute("""SELECT qs.code, ch.code, qs.bloom, qs.qtype, qs.content, qs.options, qs.answer, qs.solution,
                               qs.clo, qs.tags, qs.word_limit, qs.status,
                               COALESCE((SELECT p.content FROM passages p WHERE p.id = qs.passage_id), qs.passage, '')
                               FROM questions qs JOIN chapters ch ON ch.id = qs.chapter_id ORDER BY qs.code""").fetchall()


def count():
    with sqlite3.connect(app.DB_PATH) as conn:
        return conn.execute("SELECT COUNT(*) FROM questions").fetchone()[0]


def read(raw):
    return app.read_table_text(raw)


def test_export_has_import_columns_and_question_code(bank):
    rows = read(app.export_questions_csv(bank["tid"]))
    assert rows[0] == app.EXPORT_HEADERS
    assert rows[0][0] == "Mã câu" and rows[0][1] == "Mã chương"
    assert len(rows) - 1 == 6


def test_round_trip_updates_same_questions_without_changes(bank):
    before = snapshot()
    rows = read(app.export_questions_csv(bank["tid"]))
    res = app.parse_rows([rows], bank["tid"], dry_run=False, docx=False)
    assert res["errors"] == [], res["errors"]
    assert res["updated"] == 6 and res["added"] == 0
    assert count() == 6
    assert snapshot() == before  # nhập lại file chưa sửa không làm đổi gì


def test_edit_in_csv_updates_that_question_only(bank):
    rows = read(app.export_questions_csv(bank["tid"]))
    # sửa nội dung và đáp án câu TN đầu tiên (mã nằm ở dòng 2)
    rows[1][2] = "Từ nào đúng hơn?\nDòng hai\nA. cat\nB. dog\nC. #cố định"  # giữ nguyên các phương án
    rows[1][3] = "A"
    res = app.parse_rows([rows], bank["tid"], dry_run=False, docx=False)
    assert res["errors"] == [] and res["updated"] == 6
    code = rows[1][0]
    with sqlite3.connect(app.DB_PATH) as conn:
        content, answer = conn.execute("SELECT content, answer FROM questions WHERE code=?", (code,)).fetchone()
    assert content == "Từ nào đúng hơn?\nDòng hai" and answer == "A"
    assert app.q("SELECT options FROM questions WHERE code=?", (code,), True)[0]["options"] == '["cat", "dog", "#cố định"]'
    assert count() == 6


def test_new_row_without_code_is_added_and_unknown_code_warns(bank):
    rows = read(app.export_questions_csv(bank["tid"]))
    rows.append(["", "B02", "Câu mới\nA. một\nB. hai", "B", "TN", "1", "", "", "", "", "", "Nháp"])
    rows.append(["MÃ-KHÔNG-CÓ", "B01", "Câu lạ", "A", "TN", "1", "", "", "", "", "", "Nháp"])
    rows[-1][2] = "Câu lạ\nA. x\nB. y"
    res = app.parse_rows([rows], bank["tid"], dry_run=False, docx=False)
    assert res["errors"] == []
    assert res["added"] == 2 and res["updated"] == 6
    assert count() == 8
    assert any("MÃ-KHÔNG-CÓ" in w for w in res["warnings"])


def test_question_from_other_textbook_is_not_overwritten(bank):
    app.ensure_chapter(["Tiếng Nhật 4"], "Minna", "M01", "Minna 1")
    other = app.ensure_chapter(["Tiếng Nhật 4"], "Minna", "M01", "Minna 1")
    app.insert_q(app.clean_question({"chapter_id": other, "bloom": 1, "qtype": "DT", "content": "Khác ___.",
                                     "answer": "x", "options": []}))
    code = app.q("SELECT code FROM questions WHERE chapter_id=?", (other,), True)[0]["code"]
    rows = read(app.export_questions_csv(bank["tid"]))
    rows.append([code, "B01", "Ghi đè thử", "x", "DT", "1", "", "", "", "", "", "Nháp"])
    res = app.parse_rows([rows], bank["tid"], dry_run=False, docx=False)
    assert any("giáo trình khác" in e for e in res["errors"])
    assert app.q("SELECT content FROM questions WHERE code=?", (code,), True)[0]["content"] == "Khác ___."


def test_export_single_chapter_and_dry_run_writes_nothing(bank):
    rows = read(app.export_questions_csv(bank["tid"], bank["c2"]))
    assert len(rows) - 1 == 2  # chương B02 có 2 câu
    before = snapshot()
    res = app.parse_rows([rows], bank["tid"], dry_run=True, docx=False)
    assert res["errors"] == []
    assert [r["status"] for r in res["rows"]] == ["update", "update"]
    assert snapshot() == before
