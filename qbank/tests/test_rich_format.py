"""Kiểm thử định dạng văn bản giàu: gạch chân <u>, in đậm <b>, in nghiêng <i>, furigana trong đề xuất docx và HTML."""
import io
import os
import sys
import pytest
from docx import Document

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import app  # noqa: E402


def test_parse_rich_segments_underline():
    text = "母は専業<u>しゅふ</u>として家族を支えています。"
    tokens = app.parse_rich_segments(text)
    assert len(tokens) == 3
    assert tokens[0] == ("text", "母は専業", {"u": False, "b": False, "i": False, "s": False})
    assert tokens[1] == ("text", "しゅふ", {"u": True, "b": False, "i": False, "s": False})
    assert tokens[2] == ("text", "として家族を支えています。", {"u": False, "b": False, "i": False, "s": False})


def test_parse_rich_segments_typo_closing_tag():
    # Trường hợp gõ nhầm <u>...<u> thay vì </u>
    text = "大切な書類やお金は<u>金庫<u>に入れています。"
    tokens = app.parse_rich_segments(text)
    assert len(tokens) == 3
    assert tokens[0] == ("text", "大切な書類やお金は", {"u": False, "b": False, "i": False, "s": False})
    assert tokens[1] == ("text", "金庫", {"u": True, "b": False, "i": False, "s": False})
    assert tokens[2] == ("text", "に入れています。", {"u": False, "b": False, "i": False, "s": False})


def test_parse_rich_segments_with_ruby():
    text = "大切な書類やお{金|きん}は<u>{金庫|きんこ}</u>に{入|にゅう}れています。"
    tokens = app.parse_rich_segments(text)
    ruby_kinko = [t for t in tokens if t[0] == "ruby" and t[1][0] == "金庫"]
    assert len(ruby_kinko) == 1
    assert ruby_kinko[0][2]["u"] is True

    ruby_kin = [t for t in tokens if t[0] == "ruby" and t[1][0] == "金"]
    assert len(ruby_kin) == 1
    assert ruby_kin[0][2]["u"] is False


def test_add_rich_docx_runs():
    doc = Document()
    p = doc.add_paragraph()
    app.add_rich(p, "大切な書類やお金は<u>金庫</u>に入れています。")

    runs = p.runs
    assert any(r.text == "金庫" and r.underline is True for r in runs)
    assert all("<u>" not in r.text and "</u>" not in r.text for r in runs)


def test_add_rich_docx_typo_runs():
    doc = Document()
    p = doc.add_paragraph()
    app.add_rich(p, "大切な書類やお金は<u>金庫<u>に入れています。")

    runs = p.runs
    assert any(r.text == "金庫" and r.underline is True for r in runs)
    assert all("<u>" not in r.text for r in runs)


def test_rich_html_rendering():
    text = "母は専業<u>しゅふ</u>として家族を支えています。"
    html = app.rich_html(text)
    assert "<u>しゅふ</u>" in html
    assert "&lt;u&gt;" not in html

    # Typo <u>...<u> được sửa thành <u>...</u>
    text_typo = "大切な書類やお金は<u>金庫<u>に入れています。"
    html_typo = app.rich_html(text_typo)
    assert "<u>金庫</u>" in html_typo
    assert "&lt;u&gt;" not in html_typo


def test_docx_import_underline_preserved():
    doc = Document()
    tbl = doc.add_table(rows=1, cols=1)
    cell = tbl.cell(0, 0)
    p = cell.paragraphs[0]
    p.add_run("母は専業")
    r_u = p.add_run("しゅふ")
    r_u.underline = True
    p.add_run("として家族を支えています。")

    extracted = app._cell_import_text(cell)
    assert "<u>しゅふ</u>" in extracted
