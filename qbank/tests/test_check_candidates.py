"""Kiểm trùng hàng loạt trước khi Claude tạo câu: trùng cùng chương, chương khác, gần giống, trùng trong lô."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402

LONG = "ふくやで　おんなの　ひとと　てんいんが　はなして　います。おんなの　ひとは　どの　シャツを　かいますか。"


@pytest.fixture
def bank(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "data" / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    monkeypatch.setattr(app, "BACKUP_DIR", str(tmp_path / "data" / "backup"))
    os.makedirs(tmp_path / "data")
    app.init_storage()
    c1 = app.ensure_chapter(["Nghe N5"], "N5 Nghe", "MD51", "Mondai 1")
    c2 = app.ensure_chapter(["Nghe N5"], "N5 Nghe", "MD54", "Mondai 4")
    app.insert_q(app.clean_question({"chapter_id": c1, "bloom": 2, "qtype": "TN", "content": LONG,
                                     "options": ["しろい", "くろい", "あかい", "あおい"], "answer": "D"}))
    return {"c1": c1, "c2": c2}


def item(chapter, content, options=("しろい", "くろい", "あかい", "あおい"), qtype="TN"):
    return {"chapter_id": chapter, "qtype": qtype, "content": content, "options": list(options)}


def test_exact_duplicate_same_and_other_chapter(bank):
    res = app.check_candidates([item(bank["c1"], LONG), item(bank["c2"], LONG + "　"), item(bank["c1"], "まったく　ちがう　もんだい")])
    assert res[0]["status"] == "trùng" and res[0]["exact"][0]["same_chapter"] is True
    assert res[1]["status"] == "trùng chương khác" and res[1]["exact"][0]["same_chapter"] is False
    assert res[2]["status"] == "mới" and not res[2]["exact"] and not res[2]["similar"]


def test_near_duplicate_is_flagged_with_ratio(bank):
    near = LONG.replace("シャツ", "ぼうし")
    res = app.check_candidates([item(bank["c1"], near, options=("しろい", "くろい", "あかい", "あおい"))])
    assert res[0]["status"] == "gần giống" and res[0]["similar"][0]["ratio"] >= 0.9


def test_duplicates_inside_the_batch(bank):
    a = "みせで　りんごを　みっつ　かいます。ぜんぶで　いくらですか。"
    res = app.check_candidates([item(bank["c1"], a), item(bank["c1"], a + "　"), item(bank["c1"], "べつの　もんだい　です")])
    assert res[0]["in_batch"] and res[0]["in_batch"][0]["index"] == 1
    assert res[1]["in_batch"][0]["index"] == 0 and res[0]["status"] == "gần giống"
    assert not res[2]["in_batch"]


def test_qtype_is_respected_for_similarity(bank):
    res = app.check_candidates([item(bank["c1"], LONG.replace("シャツ", "ぼうし"), qtype="DS")])
    assert res[0]["status"] == "mới"
