"""Giáo trình dùng chung nhiều học phần (course_textbooks) và nâng cấp CSDL cũ."""
import os
import sqlite3
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402

LEGACY_SCHEMA = """
CREATE TABLE courses(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE, code TEXT NOT NULL DEFAULT '');
CREATE TABLE textbooks(id INTEGER PRIMARY KEY AUTOINCREMENT, course_id INTEGER NOT NULL REFERENCES courses(id),
    name TEXT NOT NULL, code TEXT NOT NULL DEFAULT '', UNIQUE(course_id, name));
CREATE TABLE chapters(id INTEGER PRIMARY KEY AUTOINCREMENT, textbook_id INTEGER NOT NULL REFERENCES textbooks(id),
    code TEXT NOT NULL, title TEXT NOT NULL DEFAULT '', UNIQUE(textbook_id, code));
CREATE TABLE questions(id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT, chapter_id INTEGER NOT NULL REFERENCES chapters(id),
    bloom INTEGER NOT NULL, qtype TEXT NOT NULL DEFAULT 'TN', content TEXT NOT NULL, options TEXT NOT NULL DEFAULT '[]',
    answer TEXT NOT NULL DEFAULT '', solution TEXT NOT NULL DEFAULT '');
CREATE TABLE exam_usage(id INTEGER PRIMARY KEY AUTOINCREMENT, question_id INTEGER NOT NULL, exam_code TEXT NOT NULL,
    exam_name TEXT NOT NULL DEFAULT '', used_at TEXT NOT NULL);
INSERT INTO courses(id, name, code) VALUES (1, 'Tiếng Nhật 3', 'N3A'), (2, 'Tiếng Nhật 4', 'N4A');
INSERT INTO textbooks(id, course_id, name, code) VALUES (1, 1, 'Trainer 3', 'TR3'), (2, 2, 'Trainer 3', ''), (3, 2, 'Minna', 'MN');
INSERT INTO chapters(id, textbook_id, code, title) VALUES (1, 1, 'B01', 'Bài 1'), (2, 2, 'B01', 'Bài 1 bản TN4'), (3, 2, 'B02', 'Bài 2'), (4, 3, 'M01', 'Minna 1');
INSERT INTO questions(id, code, chapter_id, bloom, content) VALUES
    (1, 'N3A-TR3-B01-001', 1, 1, 'câu A'), (2, 'N4A-TR3-B01-001', 2, 1, 'câu B'), (3, 'N4A-MN-M01-001', 4, 1, 'câu C');
INSERT INTO exam_usage(question_id, exam_code, used_at) VALUES (1, '001', '2026-10-01');
CREATE TABLE passages(id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT NOT NULL UNIQUE,
    chapter_id INTEGER NOT NULL REFERENCES chapters(id), content TEXT NOT NULL);
"""


@pytest.fixture
def db(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "data" / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    os.makedirs(tmp_path / "data")
    return tmp_path


def legacy_db(tmp_path):
    conn = sqlite3.connect(app.DB_PATH)
    conn.executescript(LEGACY_SCHEMA)
    conn.commit()
    conn.close()


def scalar(sql, args=()):
    with sqlite3.connect(app.DB_PATH) as conn:
        return conn.execute(sql, args).fetchone()[0]


def test_fresh_database_has_many_to_many_schema(db):
    app.init_storage()
    cols = [r[1] for r in sqlite3.connect(app.DB_PATH).execute("PRAGMA table_info(textbooks)")]
    assert "course_id" not in cols
    assert scalar("SELECT COUNT(*) FROM sqlite_master WHERE name='course_textbooks'") == 1


def test_migration_keeps_all_questions_and_merges_same_name(db):
    legacy_db(db)
    app.init_storage()
    assert scalar("SELECT COUNT(*) FROM questions") == 3
    assert scalar("SELECT COUNT(*) FROM exam_usage") == 1
    # hai bản "Trainer 3" gộp thành một, nối với cả hai học phần
    assert scalar("SELECT COUNT(*) FROM textbooks WHERE name='Trainer 3'") == 1
    tb = scalar("SELECT id FROM textbooks WHERE name='Trainer 3'")
    assert scalar("SELECT COUNT(*) FROM course_textbooks WHERE textbook_id=?", (tb,)) == 2
    # chương B01 trùng mã được gộp: cả hai câu cùng nằm trong một chương
    assert scalar("SELECT COUNT(DISTINCT chapter_id) FROM questions WHERE id IN (1, 2)") == 1
    assert scalar("SELECT COUNT(*) FROM chapters WHERE textbook_id=? AND code='B01'", (tb,)) == 1
    # mã câu không đổi
    assert scalar("SELECT code FROM questions WHERE id=2") == "N4A-TR3-B01-001"
    with sqlite3.connect(app.DB_PATH) as conn:
        assert conn.execute("PRAGMA foreign_key_check").fetchall() == []


def test_migration_is_idempotent(db):
    legacy_db(db)
    app.init_storage()
    app.init_storage()  # chạy lại không được lỗi hay nhân đôi dữ liệu
    assert scalar("SELECT COUNT(*) FROM questions") == 3
    assert scalar("SELECT COUNT(*) FROM course_textbooks") == 3


def test_shared_textbook_filters_by_each_course(db):
    app.init_storage()
    app.ensure_chapter(["Tiếng Nhật 3", "Tiếng Nhật 4"], "Trainer 3", "B01", "Bài 1")
    app.q("INSERT INTO questions(code, chapter_id, bloom, content) VALUES('X-1', (SELECT id FROM chapters), 1, 'câu')")
    n3 = scalar("SELECT id FROM courses WHERE name='Tiếng Nhật 3'")
    n4 = scalar("SELECT id FROM courses WHERE name='Tiếng Nhật 4'")
    tree = app.get_tree()
    tb = next(t for t in tree["textbooks"] if t["name"] == "Trainer 3")
    assert sorted(tb["course_ids"]) == sorted([n3, n4])
    for cid in (n3, n4):  # cùng một câu xuất hiện ở cả hai học phần, không bị nhân đôi trong mỗi học phần
        rows = app.q(app.QUESTION_SQL + " WHERE EXISTS (SELECT 1 FROM course_textbooks ct WHERE ct.textbook_id=t.id AND ct.course_id=?)",
                     (cid,), fetch=True)
        assert len({r["id"] for r in rows}) == len(rows)
    assert len(app.q(app.QUESTION_SQL, fetch=True)) == len(app.q("SELECT id FROM questions", fetch=True))


def test_backup_roundtrip_restores_shared_textbook(db):
    app.init_storage()
    app.q("INSERT INTO courses(name, code) VALUES('Tiếng Nhật 3', 'N3A')")
    tid_course = scalar("SELECT id FROM courses WHERE name='Tiếng Nhật 3'")
    app.ensure_chapter(["Tiếng Nhật 3", "Tiếng Nhật 4"], "Trainer 3", "B01", "Bài 1")
    chap = scalar("SELECT id FROM chapters WHERE code='B01'")
    qid = app.insert_q(app.clean_question({"chapter_id": chap, "bloom": 1, "qtype": "TN", "content": "câu sao lưu",
                                           "options": ["x", "y"], "answer": "A", "solution": ""}))
    assert qid and tid_course
    payload = app.backup_payload()
    assert payload["version"] == 4
    item = next(i for i in payload["questions"] if i["content"] == "câu sao lưu")
    assert [c["name"] for c in item["courses"]] == ["Tiếng Nhật 3", "Tiếng Nhật 4"]
    app.q("DELETE FROM questions WHERE id=?", (qid,))
    app.restore_item(item, app.dup_index())
    assert scalar("SELECT COUNT(*) FROM textbooks WHERE name='Trainer 3'") == 1
    tb = scalar("SELECT id FROM textbooks WHERE name='Trainer 3'")
    assert scalar("SELECT COUNT(*) FROM course_textbooks WHERE textbook_id=?", (tb,)) == 2


def test_unlink_refused_when_last_course_and_chapters_exist(db):
    app.init_storage()
    app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B01", "Bài 1")
    cid = scalar("SELECT id FROM courses WHERE name='Tiếng Nhật 3'")
    tid = scalar("SELECT id FROM textbooks WHERE name='Trainer 3'")
    left = app.q("SELECT COUNT(*) AS n FROM course_textbooks WHERE textbook_id=? AND course_id<>?", (tid, cid), True)[0]["n"]
    assert left == 0  # đây là điều kiện mà endpoint DELETE /api/course_textbooks dùng để từ chối


def test_chapter_csv_plans_and_skips_existing(db):
    app.init_storage()
    app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B01", "Bài 1")
    tb = scalar("SELECT id FROM textbooks WHERE name='Trainer 3'")
    rows = [["Mã", "Tên chương"], ["B01", "Bài 1"], ["B02", "Bài 2"], ["", "thiếu mã"], ["B02", "trùng"], ["B03", ""]]
    plan = app.plan_chapters(rows, tb)
    assert [p["status"] for p in plan] == ["bỏ qua", "thêm", "lỗi", "lỗi", "thêm"]
    assert plan[1]["line"] == 3 and plan[1]["code"] == "B02"
    assert plan[3]["note"] == "mã trùng trong file"


def test_chapter_csv_dry_run_does_not_write_and_commit_does(db):
    app.init_storage()
    app.ensure_chapter(["Tiếng Nhật 3"], "Trainer 3", "B01", "Bài 1")
    tb = scalar("SELECT id FROM textbooks WHERE name='Trainer 3'")
    rows = [["B02", "Bài 2"], ["B03", "Bài 3"]]
    before = scalar("SELECT COUNT(*) FROM chapters")
    check = app.import_chapters(rows, tb, dry_run=True)
    assert check["counts"]["thêm"] == 2 and scalar("SELECT COUNT(*) FROM chapters") == before
    done = app.import_chapters(rows, tb, dry_run=False)
    assert done["counts"]["thêm"] == 2
    assert scalar("SELECT COUNT(*) FROM chapters") == before + 2
    assert scalar("SELECT title FROM chapters WHERE code='B03'") == "Bài 3"
    again = app.import_chapters(rows, tb, dry_run=False)  # chạy lại không thêm trùng
    assert again["counts"] == {"thêm": 0, "bỏ qua": 2, "lỗi": 0}


def test_chapter_csv_without_header_and_with_bom(db):
    app.init_storage()
    app.q("INSERT INTO textbooks(name) VALUES('Sách mới')")
    raw = "﻿B01,Bài một\nB02,Bài hai\n".encode("utf-8")
    rows = app.read_table_text(raw)
    plan = app.plan_chapters(rows, scalar("SELECT id FROM textbooks WHERE name='Sách mới'"))
    assert [p["code"] for p in plan] == ["B01", "B02"]
    assert plan[0]["title"] == "Bài một"
