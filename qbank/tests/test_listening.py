"""Phần nghe: chỉ các câu của phần được đánh dấu mới vào file đề nghe; số thứ tự toàn đề giữ nguyên."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


def item(code, passage=None, media=None):
    d = {"code": code, "id": code, "passage": passage, "media": media or [], "passage_media": []}
    return d


def test_only_marked_part_is_kept_with_global_numbers():
    part1 = [item("A1"), item("A2")]
    part2 = [item("B1", media=[{"kind": "audio", "file": "b1.mp3"}]), item("B2")]
    part3 = [item("C1")]
    exam = [("I. Từ vựng", part1), ("II. Nghe", part2), ("III. Đọc", part3)]
    kept = app.listening_only(exam, [False, True, False])
    assert [(n, it["code"]) for n, _, it in kept] == [(3, "B1"), (4, "B2")]


def test_no_listening_part_gives_empty_list():
    exam = [("I", [item("A1")]), ("II", [item("A2")])]
    assert app.listening_only(exam, [False, False]) == []


def test_passage_group_stays_together_in_listening_part():
    p = [item("L1", passage="P"), item("L2", passage="P")]
    kept = app.listening_only([("Nghe", p), ("Đọc", [item("R1")])], [True, False])
    assert [it["code"] for _, _, it in kept] == ["L1", "L2"]
