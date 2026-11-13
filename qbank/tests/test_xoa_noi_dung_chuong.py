"""Xóa nội dung chương: xóa hẳn câu, giữ tên chương, sao lưu trước và khôi phục được từ sao lưu."""
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
    monkeypatch.setattr(app, "BACKUP_DIR", str(tmp_path / "data" / "backup"))
    os.makedirs(tmp_path / "data")
    app.init_storage()
    c1 = app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B01", "Bài 1")
    c2 = app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B02", "Bài 2")
    for chap, text in [(c1, "câu một"), (c1, "câu hai"), (c2, "câu khác chương")]:
        app.insert_q(app.clean_question({"chapter_id": chap, "bloom": 1, "qtype": "TN", "content": text,
                                         "options": ["x", "y"], "answer": "A", "passage": "đoạn của " + text}))
    app.q("INSERT INTO exam_usage(question_id, exam_code, used_at) SELECT id, '001', '2026-10-01' FROM questions WHERE chapter_id=?",
          (c1,))
    return {"c1": c1, "c2": c2}


def count(sql, args=()):
    with sqlite3.connect(app.DB_PATH) as conn:
        return conn.execute(sql, args).fetchone()[0]


def test_clear_removes_questions_and_keeps_chapter_name(bank):
    res = app.clear_chapter_questions(bank["c1"], confirm=True)
    assert res["deleted"] == 2 and res["backup"]
    assert count("SELECT COUNT(*) FROM questions WHERE chapter_id=?", (bank["c1"],)) == 0
    assert count("SELECT COUNT(*) FROM chapters WHERE id=?", (bank["c1"],)) == 1  # tên chương còn
    assert count("SELECT COUNT(*) FROM questions") == 1  # chương khác không bị ảnh hưởng
    assert count("SELECT COUNT(*) FROM exam_usage") == 0  # lịch sử dùng của câu đã xóa
    assert count("SELECT COUNT(*) FROM passages WHERE chapter_id=?", (bank["c1"],)) == 0


def test_backup_is_written_before_and_restorable(bank):
    before = app.build_items(app.q(app.QUESTION_SQL + " WHERE qs.chapter_id=?", (bank["c1"],), fetch=True))
    res = app.clear_chapter_questions(bank["c1"], confirm=True)
    path = os.path.join(app.BACKUP_DIR, res["backup"])
    assert os.path.exists(path)
    import json
    data = json.load(open(path, encoding="utf-8"))
    assert [i["content"] for i in data["questions"]] == [i["content"] for i in before]
    for item in data["questions"]:  # khôi phục từ sao lưu: đủ câu, đúng chương
        app.restore_item(item, app.dup_index())
    assert count("SELECT COUNT(*) FROM questions WHERE chapter_id=?", (bank["c1"],)) == 2
    assert count("SELECT COUNT(*) FROM passages WHERE chapter_id=?", (bank["c1"],)) == 2


def test_refuses_without_confirmation(bank):
    with pytest.raises(ValueError):
        app.clear_chapter_questions(bank["c1"], confirm=False)
    assert count("SELECT COUNT(*) FROM questions WHERE chapter_id=?", (bank["c1"],)) == 2


def test_backup_failure_deletes_nothing(bank, monkeypatch):
    def broken(_):
        raise OSError("đĩa đầy")
    monkeypatch.setattr(app, "write_chapter_backup", broken)
    with pytest.raises(OSError):
        app.clear_chapter_questions(bank["c1"], confirm=True)
    assert count("SELECT COUNT(*) FROM questions WHERE chapter_id=?", (bank["c1"],)) == 2


def test_empty_chapter_is_a_no_op(bank):
    empty = app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B09", "Trống")
    before = set(os.listdir(app.BACKUP_DIR)) if os.path.isdir(app.BACKUP_DIR) else set()
    res = app.clear_chapter_questions(empty, confirm=True)
    assert res == {"deleted": 0, "backup": None}
    assert count("SELECT COUNT(*) FROM chapters WHERE id=?", (empty,)) == 1
    after = set(os.listdir(app.BACKUP_DIR)) if os.path.isdir(app.BACKUP_DIR) else set()
    assert after == before  # không ghi file sao lưu rỗng


def test_two_backups_in_the_same_second_do_not_overwrite(bank):
    p1, n1 = app.write_chapter_backup(bank["c1"])
    p2, n2 = app.write_chapter_backup(bank["c1"])
    assert p1 != p2 and os.path.exists(p1) and os.path.exists(p2)
    assert n1 == n2 == 2
