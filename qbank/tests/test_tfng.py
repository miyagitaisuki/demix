"""Câu TN có thẻ TFNG (Đúng / Sai / Không có thông tin): phương án cố định, đáp án A/B/C."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


@pytest.fixture(autouse=True)
def chapter_exists(monkeypatch):
    monkeypatch.setattr(app, "q", lambda sql, args=(), fetch=False: [{"id": 1}] if fetch else None)


def base(**kw):
    d = {"chapter_id": 1, "content": "本文によると、図書館は新しい。", "qtype": "TN", "bloom": 2,
         "options": ["x", "y"], "answer": "A", "tags": "tfng"}
    d.update(kw)
    return d


def test_tfng_forces_fixed_options():
    out = app.clean_question(base(options=["bỏ", "qua"]))
    assert out["options"] == '["正（真）", "誤（偽）", "情報なし（書かれていない）"]'
    assert out["answer"] == "A"


def test_tfng_answer_must_be_a_b_or_c():
    with pytest.raises(ValueError, match="A \\(正\\)"):
        app.clean_question(base(answer="D"))


def test_tfng_tag_case_insensitive_and_normal_tn_unchanged():
    out = app.clean_question(base(tags="TFNG, ngữ pháp", answer="c"))
    assert out["answer"] == "C"
    plain = app.clean_question(base(tags="ngữ pháp", options=["a", "b", "c"], answer="B"))
    assert plain["options"] == '["a", "b", "c"]'
