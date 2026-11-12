"""Điền từ có giới hạn số từ: lưu word_limit, xuất đề ghi "(không quá N từ)", chỉ áp dụng cho DT."""
import os
import sys

import pytest
from docx import Document

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


@pytest.fixture(autouse=True)
def chapter_exists(monkeypatch):
    monkeypatch.setattr(app, "q", lambda sql, args=(), fetch=False: [{"id": 1}] if fetch else None)


def dt(**kw):
    d = {"chapter_id": 1, "content": "図書館は___です。", "qtype": "DT", "bloom": 1, "answer": "静か",
         "word_limit": 2}
    d.update(kw)
    return d


def test_word_limit_saved_for_dt_only():
    assert app.clean_question(dt())["word_limit"] == 2
    assert app.clean_question(dt(word_limit=0))["word_limit"] == 0
    assert app.clean_question(dt(qtype="TL", answer="", word_limit=5))["word_limit"] == 0


def test_word_limit_range_checked():
    with pytest.raises(ValueError, match="từ 1 đến 20"):
        app.clean_question(dt(word_limit=25))
    with pytest.raises(ValueError, match="số nguyên"):
        app.clean_question(dt(word_limit="nhiều"))


def test_export_shows_limit(tmp_path):
    tpl = os.path.join(os.path.dirname(app.__file__), "templates", "de_thi_chuan.docx")
    item = {"id": 1, "code": "D01", "qtype": "DT", "content": "図書館は___です。", "options": [], "answer": "静か",
            "passage": "", "passage_id": None, "passage_media": [], "media": [], "tags": [], "points": None,
            "word_limit": 2}
    data = app.render_docx(tpl, {}, [("Phần III. Điền từ", [item])], "T002", False)
    path = tmp_path / "de.docx"
    path.write_bytes(data)
    text = "\n".join(p.text for p in Document(str(path)).paragraphs)
    assert "Đáp án (không quá 2 từ):" in text


def test_label_item_has_hint(tmp_path):
    tpl = os.path.join(os.path.dirname(app.__file__), "templates", "de_thi_chuan.docx")
    item = {"id": 2, "code": "L01", "qtype": "DT", "content": "図の 3 の部分は何ですか。", "options": [],
            "answer": "玄関", "passage": "", "passage_id": None, "passage_media": [],
            "media": [{"kind": "image", "file": "missing.png", "name": "so_do.png"}],
            "tags": ["ghinhan"], "points": None, "word_limit": 0}
    data = app.render_docx(tpl, {}, [("Phần IV. Ghi nhãn", [item])], "T003", False)
    path = tmp_path / "de.docx"
    path.write_bytes(data)
    text = "\n".join(p.text for p in Document(str(path)).paragraphs)
    assert app.LABEL_HINT in text
    assert "so_do.png" not in text  # hình không có file trong thư mục media thì bỏ qua, không làm hỏng việc xuất


def test_import_reads_word_limit_column():
    cols, missing = app.map_import_header(["Mã chương", "Nội dung", "Đáp án", "Loại", "Mức", "Số từ"])
    assert missing == [] and cols["so_tu"] == 5
    out = app.clean_question(dict(dt(), word_limit="2"))
    assert out["word_limit"] == 2
