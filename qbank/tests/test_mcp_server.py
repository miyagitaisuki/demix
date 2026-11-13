"""MCP server của QBank: giao thức JSON-RPC và cách gọi API (dùng HTTP server giả)."""
import json
import os
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import mcp_server  # noqa: E402

CALLS = []


class Stub(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _reply(self, code, obj):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        return self.rfile.read(int(self.headers.get("Content-Length") or 0))

    def do_GET(self):
        CALLS.append(("GET", self.path, None))
        if self.path.startswith("/api/tree"):
            return self._reply(200, {"courses": [{"id": 1, "name": "TN3", "code": "N3A", "exam_meta": {}}],
                                     "textbooks": [{"id": 1, "name": "Trainer 3", "code": "TR3", "course_ids": [1], "course_id": 1}],
                                     "chapters": [{"id": 5, "textbook_id": 1, "code": "B01", "title": "Bài 1", "count": 4}]})
        if self.path.startswith("/api/questions"):
            return self._reply(200, {"items": [{"code": "N3A-TR3-B01-001", "qtype": "TN", "bloom": 1, "status": "da_duyet",
                                               "courses_all": "TN3", "textbook": "Trainer 3", "chapter_label": "Trainer 3 / B01",
                                               "content": "x" * 500, "used_count": 0}], "total": 1})
        self._reply(404, {"error": "không tìm thấy"})

    def do_POST(self):
        body = self._body()
        CALLS.append(("POST", self.path, body))
        if self.path.startswith("/api/chapters/import"):
            if "dry_run=1" in self.path:
                return self._reply(200, {"counts": {"thêm": 1, "bỏ qua": 0, "lỗi": 1}, "rows": [
                    {"line": 2, "code": "B02", "title": "Bài 2", "status": "thêm", "note": ""},
                    {"line": 3, "code": "", "title": "", "status": "lỗi", "note": "thiếu mã chương"}]})
            return self._reply(200, {"counts": {"thêm": 1, "bỏ qua": 0, "lỗi": 1}, "rows": [
                {"line": 3, "code": "", "title": "", "status": "lỗi", "note": "thiếu mã chương"}]})
        if self.path == "/api/check-candidates":
            n = len(json.loads(body)["items"])
            return self._reply(200, {"results": [{"index": i, "status": "mới", "exact": [], "similar": [], "in_batch": []} for i in range(n)]})
        if self.path == "/api/chapters":
            return self._reply(400, {"error": "Mã chương đã tồn tại"})
        self._reply(201, {"ok": True, "id": 9})


@pytest.fixture(scope="module")
def stub():
    srv = HTTPServer(("127.0.0.1", 0), Stub)
    t = threading.Thread(target=srv.serve_forever, daemon=True)
    t.start()
    old = mcp_server.BASE_URL
    mcp_server.BASE_URL = f"http://127.0.0.1:{srv.server_address[1]}"
    yield srv
    mcp_server.BASE_URL = old
    srv.shutdown()


def call(name, args):
    resp = mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": name, "arguments": args}})
    res = resp["result"]
    return res["isError"], json.loads(res["content"][0]["text"]) if not res["isError"] else res["content"][0]["text"]


def test_initialize_echoes_protocol_and_lists_tools():
    r = mcp_server.handle({"jsonrpc": "2.0", "id": 0, "method": "initialize", "params": {"protocolVersion": "2025-06-18"}})
    assert r["result"]["protocolVersion"] == "2025-06-18"
    assert "tools" in r["result"]["capabilities"]
    names = [t["name"] for t in mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "tools/list"})["result"]["tools"]]
    assert {"list_tree", "list_questions", "create_chapter", "import_chapters"} <= set(names)
    assert "delete" not in " ".join(names)  # không có công cụ xóa


def test_notifications_get_no_reply():
    assert mcp_server.handle({"jsonrpc": "2.0", "method": "notifications/initialized"}) is None


def test_list_tree_and_questions(stub):
    err, tree = call("list_tree", {})
    assert not err and tree["chapters"][0]["questions"] == 4
    err, res = call("list_questions", {"course_id": 1, "limit": 5})
    assert not err and res["total"] == 1
    assert len(res["items"][0]["content"]) == 300  # nội dung được rút gọn
    assert "course_id=1" in CALLS[-1][1] and "per=5" in CALLS[-1][1]


def test_import_defaults_to_dry_run_and_real_write_needs_false(stub):
    csv = "Mã,Tên chương\nB02,Bài 2\n,thiếu mã\n"
    err, res = call("import_chapters", {"textbook_id": 1, "csv": csv})
    assert not err and res["dry_run"] is True
    assert "dry_run=1" in CALLS[-1][1]
    err, res = call("import_chapters", {"textbook_id": 1, "csv": csv, "dry_run": False})
    assert not err and res["dry_run"] is False
    assert "dry_run" not in CALLS[-1][1]
    assert CALLS[-1][2] == csv.encode("utf-8")  # gửi nguyên nội dung CSV dạng UTF-8


def test_api_error_is_returned_as_tool_error(stub):
    err, msg = call("create_chapter", {"textbook_id": 1, "code": "B01"})
    assert err and "đã tồn tại" in msg


def test_bad_arguments_and_unknown_tool(stub):
    err, msg = call("create_chapter", {"textbook_id": "abc", "code": "B01"})
    assert err and "số nguyên" in msg
    r = mcp_server.handle({"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "xoa_tat_ca"}})
    assert "error" in r


def test_unreachable_app_gives_clear_message():
    old = mcp_server.BASE_URL
    mcp_server.BASE_URL = "http://127.0.0.1:1"
    try:
        err, msg = call("list_tree", {})
    finally:
        mcp_server.BASE_URL = old
    assert err and "Không kết nối được app" in msg


def test_create_question_with_audio_attaches_to_passage(stub):
    err, res = call("create_question", {"chapter_id": 5, "content": "男の人は何をしますか。", "options": ["a", "b", "c", "d"],
                                        "answer": "A", "passage": "[1] こんにちは\n[2] はい", "audio_script": "[1] こんにちは\n[2] はい",
                                        "audio_target": "passage", "audio_level": "N3"})
    assert not err and res["id"] == 9
    paths = [c[1] for c in CALLS[-3:]]
    assert paths == ["/api/check-candidates", "/api/tts", "/api/questions"]  # kiểm trùng chạy trước, rồi mới tạo audio
    sent = json.loads(CALLS[-1][2])
    assert "passage_media" in sent and "media" not in sent and sent["status"] == "nhap"


def test_create_question_transcript_defaults_to_audio_script_and_keeps_translation(stub):
    err, res = call("create_question", {"chapter_id": 5, "content": "x", "options": ["a", "b"], "answer": "A",
                                        "audio_script": "[1] こんにちは", "translation": "Xin chào"})
    assert not err
    sent = json.loads(CALLS[-1][2])
    assert sent["transcript"] == "[1] こんにちは" and sent["translation"] == "Xin chào"


def test_create_question_audio_needs_passage_when_target_passage(stub):
    err, msg = call("create_question", {"chapter_id": 5, "content": "x", "audio_script": "[1] a", "audio_target": "passage"})
    assert err and "passage" in msg


def test_new_tools_are_listed():
    names = [t["name"] for t in mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "tools/list"})["result"]["tools"]]
    assert {"import_questions", "attach_audio"} <= set(names)


def test_import_questions_dry_run_default_and_extension_check(tmp_path, monkeypatch):
    f = tmp_path / "de.docx"
    f.write_bytes(b"PK-fake")
    seen = []

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        seen.append((method, path, params, body))
        return {"added": 1, "errors": [], "rows": [{"where": "B1", "code": "B01", "status": "ok", "note": ""}]}

    monkeypatch.setattr(mcp_server, "api", fake_api)
    err, res = call("import_questions", {"textbook_id": 3, "file_path": str(f)})
    assert not err and res["dry_run"] is True and seen[-1][2]["dry_run"] == "1" and seen[-1][3] == b"PK-fake"
    err, res = call("import_questions", {"textbook_id": 3, "file_path": str(f), "dry_run": False})
    assert not err and res["dry_run"] is False and seen[-1][2]["dry_run"] is None
    txt = tmp_path / "a.txt"
    txt.write_text("x")
    err, msg = call("import_questions", {"textbook_id": 3, "file_path": str(txt)})
    assert err and ".docx" in msg


def test_attach_audio_keeps_fields_and_old_media(tmp_path, monkeypatch):
    mp3 = tmp_path / "ban1.mp3"
    mp3.write_bytes(b"ID3")
    calls = []
    question = {"code": "X-001", "chapter_id": 5, "bloom": 2, "qtype": "TN", "content": "c", "options": ["a", "b"],
                "answer": "B", "solution": "s", "transcript": "t", "translation": "d", "passage": "p", "rubric": [], "points": 1, "clo": "CLO2", "tags": [],
                "status": "nhap", "word_limit": 0, "media": [{"kind": "image", "file": "old.png", "name": "old.png"}],
                "passage_id": None, "extra_ignored": 1}

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        calls.append((method, path, params, body))
        if method == "GET":
            return question
        if path == "/api/media":
            return {"kind": "audio", "file": "abc.mp3", "name": "ban1.mp3"}
        return {"ok": True}

    monkeypatch.setattr(mcp_server, "api", fake_api)
    err, res = call("attach_audio", {"question_id": 7, "file_path": str(mp3)})
    assert not err and res["code"] == "X-001"
    method, path, _, body = calls[-1]
    assert (method, path) == ("PUT", "/api/questions/7")
    assert body["status"] == "nhap" and body["answer"] == "B" and "extra_ignored" not in body
    assert body["transcript"] == "t" and body["translation"] == "d"  # PUT không làm mất transcript/dịch
    assert [m["file"] for m in body["media"]] == ["old.png", "abc.mp3"]
    err, msg = call("attach_audio", {"question_id": 7, "file_path": str(mp3), "target": "passage"})
    assert err and "đoạn văn" in msg  # câu không có đoạn văn
    bad = tmp_path / "a.png"
    bad.write_bytes(b"x")
    err, msg = call("attach_audio", {"question_id": 7, "file_path": str(bad)})
    assert err and ".mp3" in msg


def _fake_bank(monkeypatch, question):
    puts = []

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        if method == "GET":
            return dict(question)
        if path == "/api/media":
            return {"kind": "audio", "file": "new.mp3", "name": "new.mp3"}
        puts.append((path, body))
        return {"ok": True}

    monkeypatch.setattr(mcp_server, "api", fake_api)
    return puts


def _q(**kw):
    base = {"code": "X-1", "chapter_id": 5, "bloom": 1, "qtype": "TN", "content": "c", "options": ["a", "b"], "answer": "A",
            "status": "nhap", "passage": "p", "passage_id": 9,
            "media": [{"kind": "audio", "file": "old.mp3", "name": "old.mp3"}, {"kind": "image", "file": "i.png", "name": "i.png"}],
            "passage_media": [{"kind": "audio", "file": "pold.mp3", "name": "pold.mp3"}]}
    base.update(kw)
    return base


def test_attach_audio_replaces_old_audio_by_default(tmp_path, monkeypatch):
    mp3 = tmp_path / "new.mp3"
    mp3.write_bytes(b"ID3")
    puts = _fake_bank(monkeypatch, _q())
    err, res = call("attach_audio", {"question_id": 7, "file_path": str(mp3)})
    assert not err and res["replaced"] == ["old.mp3"]
    assert [m["file"] for m in puts[-1][1]["media"]] == ["i.png", "new.mp3"]  # ảnh giữ, audio cũ bị thay
    assert "đoạn văn" in res["warning"]  # đoạn văn vẫn còn file nghe -> báo 2 file
    err, res = call("attach_audio", {"question_id": 7, "file_path": str(mp3), "replace": False})
    assert [m["file"] for m in puts[-1][1]["media"]] == ["old.mp3", "i.png", "new.mp3"]


def test_remove_audio_selects_files_and_dry_run(monkeypatch):
    puts = _fake_bank(monkeypatch, _q())
    err, res = call("remove_audio", {"question_id": 7, "dry_run": True})
    assert not err and res["removed"] == {"question": ["old.mp3"]} and not puts  # xem trước, chưa ghi
    err, res = call("remove_audio", {"question_id": 7, "target": "both"})
    assert res["removed"] == {"question": ["old.mp3"], "passage": ["pold.mp3"]}
    assert [m["file"] for m in puts[0][1]["media"]] == ["i.png"]  # chỉ gỡ audio, ảnh còn
    assert puts[1] == ("/api/passages/9", {"media": []})
    err, msg = call("remove_audio", {"question_id": 7, "file": "khong-co.mp3"})
    assert err and "Không thấy" in msg


def test_get_question_audio_lists_both_levels(monkeypatch):
    _fake_bank(monkeypatch, _q())
    err, res = call("get_question_audio", {"question_id": 7})
    assert not err and res["question_audio"][0]["file"] == "old.mp3" and res["passage_audio"][0]["file"] == "pold.mp3"

def test_check_duplicates_tool_summarises(monkeypatch):
    seen = []

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        seen.append((method, path, body))
        return {"results": [{"index": 0, "status": "mới"}, {"index": 1, "status": "trùng"}, {"index": 2, "status": "mới"}]}

    monkeypatch.setattr(mcp_server, "api", fake_api)
    err, res = call("check_duplicates", {"items": [{"chapter_id": 1, "content": "a"}] * 3})
    assert not err and res["summary"] == {"mới": 2, "gần giống": 0, "trùng chương khác": 0, "trùng": 1}
    assert seen[-1][:2] == ("POST", "/api/check-candidates")
    err, msg = call("check_duplicates", {"items": []})
    assert err and "items" in msg


def test_create_question_passes_confirm_passage_only_when_true(monkeypatch):
    seen = []

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        if path == "/api/check-candidates":
            return {"results": [{"status": "mới", "exact": [], "similar": []}]}
        seen.append(body)
        return {"id": 1, "code": "X-1"}

    monkeypatch.setattr(mcp_server, "api", fake_api)
    base = {"chapter_id": 1, "content": "c", "options": ["a", "b"], "answer": "A", "passage": "ばん3"}
    call("create_question", base)
    assert "confirm_passage" not in seen[-1]
    call("create_question", {**base, "confirm_passage": True})
    assert seen[-1]["confirm_passage"] is True
    tool = mcp_server.TOOL_BY_NAME["create_question"]
    assert "confirm_passage" in tool["inputSchema"]["properties"]


def _create_fake(monkeypatch, status, exact=None, similar=None, check_error=None):
    calls = []

    def fake_api(method, path, params=None, body=None, content_type="application/json"):
        calls.append(path)
        if path == "/api/check-candidates":
            if check_error:
                raise mcp_server.ApiError(check_error)
            return {"results": [{"index": 0, "status": status, "exact": exact or [], "similar": similar or [], "in_batch": []}]}
        if path == "/api/tts":
            return {"kind": "audio", "file": "a.mp3", "name": "a.mp3"}
        return {"id": 3, "code": "X-3", "warnings": []}

    monkeypatch.setattr(mcp_server, "api", fake_api)
    return calls


BASE_Q = {"chapter_id": 1, "content": "c", "options": ["a", "b"], "answer": "A"}


def test_create_question_auto_checks_and_blocks_similar(monkeypatch):
    calls = _create_fake(monkeypatch, "gần giống", similar=[{"code": "N-1", "ratio": 0.94}])
    err, msg = call("create_question", {**BASE_Q, "audio_script": "[1] あ"})
    assert err and "N-1" in msg and "94%" in msg
    assert "/api/tts" not in calls and "/api/questions" not in calls  # chặn trước khi tạo audio và ghi
    err, res = call("create_question", {**BASE_Q, "allow_similar": True})
    assert not err and res["code"] == "X-3"


def test_create_question_blocks_exact_and_warns_other_chapter(monkeypatch):
    _create_fake(monkeypatch, "trùng", exact=[{"code": "N-2", "same_chapter": True}])
    err, msg = call("create_question", {**BASE_Q, "allow_similar": True})
    assert err and "N-2" in msg  # trùng y hệt: allow_similar cũng không cho qua
    _create_fake(monkeypatch, "trùng chương khác", exact=[{"code": "N-5", "same_chapter": False}])
    err, res = call("create_question", BASE_Q)
    assert not err and any("N-5" in w for w in res["warnings"])


def test_create_question_skips_check_on_old_app(monkeypatch):
    _create_fake(monkeypatch, "mới", check_error="không tìm thấy")
    err, res = call("create_question", BASE_Q)
    assert not err and any("build lại" in w for w in res["warnings"])
