"""Trang luyện nghe: danh sách bài có audio, nội dung một bài (audio, transcript, đáp án), ẩn bài thiếu audio."""
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


@pytest.fixture()
def db(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    media = tmp_path / "media"
    media.mkdir()
    monkeypatch.setattr(app, "MEDIA_DIR", str(media))
    app.init_storage()
    (media / "a1.mp3").write_bytes(b"ID3")
    app.q("INSERT INTO courses(name, code) VALUES('Tiếng Nhật 1','N5')")
    app.q("INSERT INTO textbooks(name, code) VALUES('Nghe N5','NG5')")
    co = app.q("SELECT id FROM courses", fetch=True)[0]["id"]
    tb = app.q("SELECT id FROM textbooks", fetch=True)[0]["id"]
    app.q("INSERT INTO course_textbooks(course_id, textbook_id) VALUES(?,?)", (co, tb))
    app.q("INSERT INTO chapters(textbook_id, code, title) VALUES(?,?,?)", (tb, "B01", "Mua sắm"))
    app.q("INSERT INTO chapters(textbook_id, code, title) VALUES(?,?,?)", (tb, "B02", "Không audio"))
    ids = {r["code"]: r["id"] for r in app.q("SELECT id, code FROM chapters", fetch=True)}
    ch_audio, ch_text = ids["B01"], ids["B02"]
    return {"audio": ch_audio, "text": ch_text, "media": media}


def make(chapter, media_list, code):
    n = app.clean_question({"chapter_id": chapter, "content": "どこへ いきますか。", "qtype": "TN", "bloom": 1,
                            "options": ["1かい", "2かい"], "answer": "B", "solution": "2階 = tầng 2"})
    n.update(code=code, media=json.dumps(media_list), passage="", passage_id=None, points=1, clo="", tags="",
             status="da_duyet")
    app.q(app.INSERT_Q, n)


def test_only_chapters_with_audio_are_listed(db):
    make(db["audio"], [{"kind": "audio", "file": "a1.mp3", "name": "ban1.mp3",
                        "tts": {"script": "[1] すみません // Xin lỗi", "engine": "aivis", "level": "N5"}}], "Q1")
    make(db["text"], [], "Q2")
    sets = app.listening_sets()
    assert [s["chapter_id"] for s in sets] == [db["audio"]]
    assert sets[0]["count"] == 1 and sets[0]["course"] == "Tiếng Nhật 1"


def test_set_returns_audio_transcript_and_answer(db):
    make(db["audio"], [{"kind": "audio", "file": "a1.mp3", "name": "ban1.mp3",
                        "tts": {"script": "[1] すみません // Xin lỗi", "engine": "aivis", "level": "N5"}}], "Q1")
    s = app.listening_set(db["audio"])
    it = s["items"][0]
    assert it["audio"][0]["url"] == "/media/a1.mp3"
    assert "Xin lỗi" in it["transcript"] and it["answer"] == "B" and it["solution"].startswith("2階")


def test_missing_audio_file_or_chapter_is_rejected(db):
    make(db["audio"], [{"kind": "audio", "file": "khong-co.mp3", "name": "x.mp3"}], "Q1")
    assert app.listening_sets() == []
    with pytest.raises(ValueError):
        app.listening_set(db["audio"])


def test_question_transcript_and_translation_override_tts_script(db):
    make(db["audio"], [{"kind": "audio", "file": "a1.mp3", "name": "b.mp3", "tts": {"script": "[1] cũ", "engine": "aivis", "level": "N5"}}], "Q1")
    app.q("UPDATE questions SET transcript=?, translation=?", ("[1] はい // vâng", "Vâng"))
    it = app.listening_set(db["audio"])["items"][0]
    assert it["transcript"].startswith("[1] はい") and it["translation"] == "Vâng"
