"""Unit tests cho ptc_client.py của QBank."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import mcp_server  # noqa: E402
import ptc_client  # noqa: E402


def test_iter_questions_paginates_and_returns_full_items(monkeypatch):
    calls = []

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        calls.append((method, path, dict(params or {})))
        page = (params or {}).get("page", 0)
        if page == 0:
            return {
                "total": 3,
                "items": [
                    {"id": 1, "code": "C-1", "qtype": "TN", "content": "A" * 400, "options": ["1", "2", "3", "4"]},
                    {"id": 2, "code": "C-2", "qtype": "TN", "content": "B", "options": ["1", "2", "3", "4"]},
                ],
            }
        return {
            "total": 3,
            "items": [
                {"id": 3, "code": "C-3", "qtype": "DT", "content": "これ___ペンです。"},
            ],
        }

    monkeypatch.setattr(mcp_server, "api", fake_api)
    client = ptc_client.QBankClient()
    items = list(client.iter_questions(course_id=1, per=2, full=True))
    assert [x["id"] for x in items] == [1, 2, 3]
    assert len(items[0]["content"]) == 400
    assert len(calls) == 2


def test_create_question_enforces_nhap_and_aivis(monkeypatch):
    captured = {}

    def fake_handler(args):
        captured.update(args)
        return {"ok": True, "id": 99}

    monkeypatch.setitem(mcp_server.TOOL_BY_NAME["create_question"], "handler", fake_handler)
    client = ptc_client.QBankClient()
    res = client.create_question(
        chapter_id=1,
        qtype="TN",
        content="テスト",
        audio_script="[1] こんにちは。",
        audio_engine="voicevox",
    )
    assert res["id"] == 99
    assert captured["status"] == "nhap"
    assert captured["audio_engine"] == "aivis"


def test_audit_question_rules_catches_jlpt_violations():
    q = {
        "id": 10,
        "code": "N5-010",
        "qtype": "DT",
        "content": "日本 語の主語は何ですか。",
        "options": ["は", "は", "を", "に"],
    }
    issues = ptc_client.audit_question_rules(q)
    rules = {it["rule"] for it in issues}
    assert "jp_space" in rules
    assert "linguistic_term" in rules
    assert "dt_blank" in rules

    q_tn = {
        "id": 11,
        "code": "N5-011",
        "qtype": "TN",
        "content": "{漢字|かな}と{壊れ}と山",
        "options": ["A", "A", "B", "C"],
    }
    issues_tn = ptc_client.audit_question_rules(q_tn, allowed_kanji={"漢", "字"})
    rules_tn = {it["rule"] for it in issues_tn}
    assert "ruby_syntax" in rules_tn
    assert "options_invalid" in rules_tn
    assert "scope_kanji" in rules_tn


def test_audit_question_rules_passes_clean_question():
    q = {
        "id": 12,
        "code": "N5-012",
        "qtype": "DT",
        "content": "きょうは駅へ___。",
        "options": ["行きます", "来ます", "見ます", "食べます"],
    }
    assert ptc_client.audit_question_rules(q) == []
