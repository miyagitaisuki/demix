"""Hook CSDL: cột questions, bộ nhập/xuất và file mẫu nhập phải khớp; phát hiện cột mới chưa được khai báo."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import check_db_sync  # noqa: E402


def test_repo_is_in_sync():
    assert check_db_sync.check() == []


def test_transcript_and_translation_are_questions_columns():
    with open(os.path.join(check_db_sync.BASE, "app.py"), encoding="utf-8") as f:
        cols = check_db_sync.question_columns(f.read())
    assert {"transcript", "translation", "solution", "word_limit"} <= cols


def test_new_column_without_import_support_is_detected():
    src = 'ensure_column("questions", "hint", "TEXT NOT NULL DEFAULT \'\'")\n'
    extra = check_db_sync.question_columns(src) - set(check_db_sync.COLUMN_TO_HEADER) - check_db_sync.NOT_IMPORTED
    assert extra == {"hint"}


def test_stale_sample_is_detected(tmp_path, monkeypatch):
    (tmp_path / "mau_nhap_cau_hoi.csv").write_text("﻿Mã chương,Nội dung\n", encoding="utf-8")
    monkeypatch.setattr(check_db_sync, "SAMPLE_DIR", str(tmp_path))
    problems = check_db_sync.check()
    assert any("đã cũ" in p for p in problems) and any("Thiếu file mẫu" in p and ".docx" in p for p in problems)
