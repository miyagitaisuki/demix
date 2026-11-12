"""Nối tiêu đề (matching headings): không xáo phương án, danh sách tiêu đề in một lần, tối đa 12 tiêu đề."""
import os
import sys

import pytest
from docx import Document

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402

HEADINGS = ["生活の変化", "天気の予報", "図書館の利用", "新しい店", "駅の案内", "学校の行事"]


@pytest.fixture(autouse=True)
def chapter_exists(monkeypatch):
    monkeypatch.setattr(app, "q", lambda sql, args=(), fetch=False: [{"id": 1}] if fetch else None)


def item(i, pid=7, tags=("noitieude",)):
    return {"id": i, "code": f"M{i:02d}", "qtype": "TN", "content": f"段落{i}", "options": list(HEADINGS),
            "answer": "B", "passage": "本文の段落", "passage_id": pid, "passage_media": [], "media": [],
            "tags": list(tags), "points": None}


def test_clean_question_allows_up_to_12_headings():
    many = [f"見出し{k}" for k in range(12)]
    out = app.clean_question({"chapter_id": 1, "content": "段落1", "qtype": "TN", "bloom": 2,
                              "options": many, "answer": "L", "tags": "noitieude"})
    assert len(app.json.loads(out["options"])) == 12
    with pytest.raises(ValueError, match="tối đa 8"):
        app.clean_question({"chapter_id": 1, "content": "x", "qtype": "TN", "bloom": 2,
                            "options": many, "answer": "A", "tags": "khac"})


def test_matching_items_are_not_shuffled():
    import random
    base = [item(i) for i in range(1, 4)]
    out = app.make_variant(base, True, random.Random(3))
    assert all(it["options"] == HEADINGS and it["answer"] == "B" for it in out)


def test_heading_list_printed_once_per_group(tmp_path):
    tpl = os.path.join(os.path.dirname(app.__file__), "templates", "de_thi_chuan.docx")
    exam = [("Phần I. Nối tiêu đề", [item(1), item(2), item(3)])]
    data = app.render_docx(tpl, {}, exam, "T001", False)
    path = tmp_path / "de.docx"
    path.write_bytes(data)
    text = "\n".join(p.text for p in Document(str(path)).paragraphs)
    assert text.count("Danh sách tiêu đề") == 1
    assert text.count("Đáp án (chữ cái tiêu đề)") == 3
    assert "A. 生活の変化" not in text.split("Danh sách tiêu đề")[0]
