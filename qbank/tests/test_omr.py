"""Phiếu trả lời trắc nghiệm chấm máy: bố cục, tờ đáp án đục lỗ, đáp án theo mã đề, chấm điểm, lưu tỷ lệ đúng, mẫu đề TN."""
import json
import os
import sys

import pytest
from docx import Document

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402
import omr  # noqa: E402


def test_layout_capacity_and_positions():
    L = omr.layout({"tn": 100})
    assert len(L["tn"]) == 100 and [r["n"] for r in L["tn"]] == list(range(1, 101))
    xs = [b["x"] for r in L["tn"] for b in r["bubbles"]]
    ys = [r["y"] for r in L["tn"]]
    assert min(xs) > omr.AREA_X0 and max(xs) < omr.AREA_X1 and max(ys) < omr.AREA_Y1
    with pytest.raises(ValueError, match="không đủ chỗ"):
        omr.layout({"tn": 140})
    L = omr.layout({"tn": 30, "ds": [31, 32, 33, 34], "sbd": 6, "made": 3})
    assert [b["n"] for b in L["ds"]] == [31, 32, 33, 34] and len(L["ds"][0]["items"]) == 4
    assert len(L["sbd"]["cols"]) == 6 and len(L["made"]["cols"]) == 3 and len(L["made"]["cols"][0]) == 10
    # mọi ô tô nằm trong khung 4 mốc và không chồng lên nhau
    pts = [(b["x"], b["y"]) for r in L["tn"] for b in r["bubbles"]] + \
          [(b["x"], b["y"]) for blk in L["ds"] for it in blk["items"] for b in it["bubbles"]] + \
          [(b["x"], b["y"]) for blk in (L["sbd"], L["made"]) for col in blk["cols"] for b in col]
    assert len(set(pts)) == len(pts)
    (x0, y0), _, _, (x1, y1) = L["fiducials"]
    assert all(x0 < x < x1 and y0 < y < y1 for x, y in pts)


def test_spec_validation():
    with pytest.raises(ValueError):
        omr.layout({"tn": 0, "ds": 0})
    with pytest.raises(ValueError):
        omr.layout({"tn": 10, "opts": 9})
    assert omr.layout({"tn": 5, "ds": 2})["spec"]["ds"] == [6, 7]  # số câu đúng/sai đánh số tiếp sau trắc nghiệm


def test_sheet_and_punched_key_svg():
    blank = omr.sheet_svg({"tn": 10, "ds": [11]}, "Đề thử")
    assert blank.count("<circle") == 10 * 4 + 8 + 6 * 10 + 3 * 10 and "PHIẾU TRẢ LỜI" in blank
    key = omr.sheet_svg({"tn": 10, "ds": [11]}, key={1: "B", 2: "D", 11: "DSSD"}, code="A002")
    assert "TỜ ĐÁP ÁN ĐỤC LỖ" in key and "MÃ ĐỀ A002" in key
    assert key.count('fill="#000" stroke') == 2 + 4 + 3  # 2 câu TN + 4 ý Đ/S + 3 chữ số mã đề (002)
    page = omr.sheet_page([blank, key])
    assert page.count('class="pg"') == 2 and "size: A4" in page


def flat_items():
    return [(1, "I", {"qtype": "TN", "options": ["a", "b", "c", "d"], "answer": "C", "code": "Q1", "points": None}),
            (2, "I", {"qtype": "TN", "options": ["a", "b", "c", "d"], "answer": "A", "code": "Q2", "points": 2}),
            (3, "II", {"qtype": "DS", "options": ["x", "y", "z", "w"], "answer": "DSSD", "code": "Q3", "points": None}),
            (4, "III", {"qtype": "TL", "options": [], "answer": "", "code": "Q4", "points": None}),
            (5, "IV", {"qtype": "TN", "options": [str(i) for i in range(12)], "answer": "K", "code": "Q5", "points": None})]


def test_keys_and_grading():
    key = omr.keys_from_flat(flat_items())
    assert [k["n"] for k in key] == [1, 2, 3]  # tự luận và nối tiêu đề 12 tiêu đề không chấm máy
    spec = omr.spec_from_keys({"001": key, "002": key})
    assert spec["tn"] == 2 and spec["ds"] == [3] and spec["made"] == 3
    g = app.omr_grade_sheet(key, {"1": "C", "2": "B", "3": "DSDD"})
    assert g["correct"] == 1 and g["total"] == 3
    assert g["raw"] == 1 + 0 + 0.75 and g["max"] == 4 and g["score10"] == pytest.approx(4.38, abs=0.01)
    assert app.omr_grade_sheet(key, {"1": "*", "2": "", "3": "????"})["raw"] == 0


def test_match_code_and_manual_keys():
    keys = app.omr_manual_keys("A001: ABCD\n002: DC-BA\n")
    assert [i["n"] for i in keys["002"]] == [1, 2, 4, 5]  # dấu - là câu không chấm
    assert app.omr_match_code(keys, "001") == "A001" and app.omr_match_code(keys, "002") == "002"
    assert app.omr_match_code(keys, "009") is None and app.omr_match_code(keys, "") is None
    with pytest.raises(ValueError, match="không đúng dạng"):
        app.omr_manual_keys("001 = ABCD")
    r = app.omr_grade({"manual_keys": "001: ABCD", "sheets": [{"made": "001", "answers": {"1": "A", "2": "B"}},
                                                              {"made": "555", "answers": {}}]})
    assert r["results"][0]["correct"] == 2 and r["results"][1]["code"] is None


@pytest.fixture
def bank(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "data" / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    monkeypatch.setattr(app, "BACKUP_DIR", str(tmp_path / "data" / "backup"))
    os.makedirs(tmp_path / "data")
    app.init_storage()
    return tmp_path


def test_export_stores_keys_and_results_saved(bank):
    ch = app.ensure_chapter(["TN1"], "Minna", "B01")
    ids = [app.insert_q(app.clean_question({"chapter_id": ch, "bloom": 1, "qtype": "TN", "content": f"問{i}（　）",
                                            "options": ["に", "で", "と", "の"], "answer": "B", "status": "da_duyet"}))
           for i in range(5)]
    codes = [app.code_of(i) for i in ids]
    key = [{"n": i + 1, "qtype": "TN", "answer": "B", "nopt": 4, "qcode": c, "points": None} for i, c in enumerate(codes)]
    app.q("INSERT INTO exam_exports(created_at, name, codes, config, data) VALUES(?,?,?,?,?)",
          ("2026-11-14T10:00:00", "Giữa kỳ", "001", json.dumps({"omr_keys": {"001": key}, "meta": {"ten_hp": "TN1"}}), b""))
    app.q("INSERT INTO exam_exports(created_at, name, codes, config, data) VALUES(?,?,?,?,?)",
          ("2026-11-01T10:00:00", "Bản cũ", "001", "{}", b""))
    exp = app.omr_export(1)
    assert exp["spec"]["tn"] == 5 and exp["codes"] == ["001"]
    with pytest.raises(ValueError, match="chưa có đáp án chấm máy"):
        app.omr_export(2)
    sheets = [{"made": "001", "answers": {"1": "B", "2": "B", "3": "A", "4": "", "5": "B"}},
              {"made": "001", "answers": {"1": "B", "2": "C", "3": "B", "4": "B", "5": "B"}}]
    r = app.omr_save_results({"export_id": 1, "exam_code": "GK-A", "sheets": sheets})
    assert r["questions"] == 5 and r["added"] == 5
    pct = {row["code"]: row["correct_pct"] for row in app.q(
        "SELECT qs.code, r.correct_pct FROM question_results r JOIN questions qs ON qs.id = r.question_id", fetch=True)}
    assert pct[codes[0]] == 100 and pct[codes[1]] == 50 and pct[codes[3]] == 50
    html = app.omr_sheet_html(lambda k: {"export_id": "1", "key": "all"}.get(k, ""))
    assert html.count("<svg") == 1 and "TỜ ĐÁP ÁN ĐỤC LỖ" in html and "TN1" in html


def test_mc_template_has_placeholders(bank):
    path = os.path.join(app.TPL_DIR, app.MC_TEMPLATE)
    assert os.path.isfile(path)
    text = "\n".join(p.text for p in app._every_paragraph(Document(path)))
    assert "{{MA_DE}}" in text and "{{CAU_HOI}}" in text and "PHIẾU TRẢ LỜI TRẮC NGHIỆM" in text
    bundled = os.path.join(os.path.dirname(app.__file__), "templates", app.MC_TEMPLATE)
    assert os.path.isfile(bundled)  # đi kèm bản đóng gói
