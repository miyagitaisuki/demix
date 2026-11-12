"""Mẫu nhập câu hỏi phải có đủ cột và các ví dụ của dạng câu mới (TFNG, nối tiêu đề, ghi nhãn, số từ)."""
import io
import os
import sys

import pytest
from docx import Document

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


@pytest.fixture
def captured(monkeypatch):
    """Chương 1 = B01, chương 2 = B02; không đụng CSDL. Ghi lại payload mà clean_question trả về."""
    def fake_q(sql, args=(), fetch=False):
        if "FROM chapters" in sql:
            return [{"id": 1, "code": "B01"}, {"id": 2, "code": "B02"}]
        return []
    monkeypatch.setattr(app, "q", fake_q)
    payloads = []
    real_clean = app.clean_question

    def spy(data):
        out = real_clean(data)
        payloads.append(out)
        return out
    monkeypatch.setattr(app, "clean_question", spy)
    return payloads


def template_header():
    doc = Document(io.BytesIO(app.make_import_template()))
    return [c.text for c in doc.tables[0].rows[0].cells]


def test_template_has_tag_and_word_limit_columns():
    assert template_header() == ["Mã chương", "Nội dung", "Đáp án", "Loại", "Mức",
                                 "Đoạn văn", "Lời giải", "CLO", "Thẻ", "Số từ"]


def test_template_guidance_mentions_new_types():
    doc = Document(io.BytesIO(app.make_import_template()))
    text = "\n".join(p.text for p in doc.paragraphs)
    for keyword in ("tfng", "noitieude", "ghinhan", "Số từ"):
        assert keyword in text


def test_template_examples_import_without_errors(captured):
    result = app.parse_import(app.make_import_template(), 1, dry_run=True)
    assert result["errors"] == []
    assert len(result["rows"]) == 9


def test_new_types_reach_the_question_payload(captured):
    app.parse_import(app.make_import_template(), 1, dry_run=True)
    tags = [p["tags"] for p in captured]
    assert any("tfng" in t for t in tags)
    assert any("noitieude" in t for t in tags)
    assert any("ghinhan" in t for t in tags)
    limits = [p["word_limit"] for p in captured if p["qtype"] == "DT"]
    assert 2 in limits  # ví dụ điền từ có Số từ = 2
    tfng = next(p for p in captured if "tfng" in p["tags"])
    assert tfng["qtype"] == "TN" and tfng["answer"] == "A"
