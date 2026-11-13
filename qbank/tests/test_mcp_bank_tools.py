"""MCP: công cụ rà soát, sửa/duyệt, ra đề/xuất đề, sao lưu/khôi phục. Mặc định dry_run; không có công cụ xóa ngoài merge_duplicates."""
import inspect
import json
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import mcp_server  # noqa: E402

NEW_TOOLS = {"bank_stats", "find_duplicates", "find_similar", "difficulty_report", "update_question", "set_status",
             "preview_exam", "export_exam", "backup_bank", "list_trash", "restore_from_trash", "preview_restore",
             "merge_duplicates"}
WRITE_TOOLS = {"update_question", "set_status", "export_exam", "backup_bank", "restore_from_trash", "merge_duplicates"}
READ_POSTS = ("/api/exam/preview", "/api/check-candidates")  # POST chỉ để đọc


def call(name, args):
    resp = mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": name, "arguments": args}})
    res = resp["result"]
    text = res["content"][0]["text"]
    return res["isError"], (text if res["isError"] else json.loads(text))


class Fake:
    """Giả lập API của app; ghi lại mọi lệnh gọi."""

    def __init__(self, questions=None):
        self.calls = []
        self.questions = questions or {}

    def writes(self):
        return [c for c in self.calls if c[0] in ("PUT", "POST", "DELETE") and c[1] not in READ_POSTS]

    def api(self, method, path, params=None, body=None, content_type="application/json"):
        self.calls.append((method, path, body))
        if method == "GET" and path.startswith("/api/questions/"):
            return dict(self.questions[int(path.rsplit("/", 1)[1])])
        if path == "/api/trash":
            return [{"id": 1, "code": "T-1", "content": "x" * 200}]
        if path == "/api/exam/preview":
            return {"sections": [{"items": [{"id": 1, "code": "A-1", "qtype": "TN", "bloom": 1, "content": "c"}]}]}
        if path == "/api/stats":
            return {"total": 2, "by_type": {"TN": 2}, "chapters": [{"id": 1, "label": "L", "total": 2, "counts": {"1": 2}}]}
        if path == "/api/duplicates/merge":
            return {"removed": 1, "kept": 1}
        if path == "/api/questions/status":
            return {"ok": True, "updated": len(body["ids"])}
        if path == "/api/check-candidates":
            return {"results": [{"status": "mới"} for _ in body["items"]]}
        return {"ok": True}

    def api_raw(self, method, path, body=None):
        self.calls.append((method, path, body))
        if path == "/api/backup":
            return json.dumps({"questions": [{"content": "a"}, {"content": "b"}]}).encode(), {}
        return b"PK-zip", {"Content-Disposition": 'attachment; filename="De_001.zip"', "X-Variants": "1", "X-Seed": "5", "X-Reused": "0"}


def q(i, content="かさ", **kw):
    d = {"id": i, "code": f"X-{i}", "chapter_id": 5, "bloom": 1, "qtype": "TN", "content": content, "options": ["a", "b"],
         "answer": "A", "solution": "s", "status": "nhap", "passage": "p", "passage_id": None, "rubric": [], "points": 1,
         "clo": "CLO1", "tags": [], "word_limit": 0, "transcript": "t", "translation": "tr",
         "media": [{"kind": "audio", "file": "a.mp3", "name": "a.mp3"}], "passage_media": []}
    d.update(kw)
    return d


@pytest.fixture
def fake(monkeypatch):
    f = Fake({1: q(1), 2: q(2), 3: q(3, content="ぜんぜん　ちがう")})
    monkeypatch.setattr(mcp_server, "api", f.api)
    monkeypatch.setattr(mcp_server, "api_raw", f.api_raw)
    return f


# R1
def test_new_tools_are_listed_with_permission_wording():
    tools = {t["name"]: t for t in mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "tools/list"})["result"]["tools"]}
    assert NEW_TOOLS <= set(tools)
    for name in WRITE_TOOLS:
        assert "dry_run" in tools[name]["inputSchema"]["properties"]
        assert "dry_run" in tools[name]["description"] or name == "merge_duplicates"
    assert "da_duyet" in tools["set_status"]["description"] and "CHỈ khi" in tools["set_status"]["description"]
    assert "confirm" in tools["export_exam"]["inputSchema"]["properties"] and "confirm" in tools["merge_duplicates"]["inputSchema"]["properties"]


# R2
@pytest.mark.parametrize("name,args", [
    ("update_question", {"question_id": 1, "content": "mới"}),
    ("set_status", {"ids": [1, 2], "status": "da_duyet"}),
    ("export_exam", {"sections": [{"matrix": [{"chapter_id": 5, "bloom": 1, "n": 1}]}], "out_dir": "x", "confirm": True}),
    ("backup_bank", {"out_dir": "x"}),
    ("restore_from_trash", {"ids": [1]}),
    ("merge_duplicates", {"keep_id": 1, "remove_ids": [2], "out_dir": "x", "confirm": True}),
])
def test_write_tools_default_to_dry_run_and_do_not_write(fake, tmp_path, name, args):
    if "out_dir" in args:
        args = {**args, "out_dir": str(tmp_path)}
    err, res = call(name, args)
    assert not err and res["dry_run"] is True
    assert fake.writes() == []
    assert os.listdir(tmp_path) == []  # không tạo file


# R3
def test_update_question_changes_only_named_fields(fake):
    err, res = call("update_question", {"question_id": 1, "content": "あたらしい", "dry_run": False})
    assert not err and list(res["changed"]) == ["content"]
    method, path, body = fake.writes()[-1]
    assert (method, path) == ("PUT", "/api/questions/1")
    assert body["content"] == "あたらしい" and body["status"] == "nhap" and body["answer"] == "A"
    assert body["media"] == q(1)["media"]  # file nghe giữ nguyên
    err, msg = call("update_question", {"question_id": 1})
    assert err and "Chưa nêu trường" in msg
    err, res = call("update_question", {"question_id": 1, "status": "da_duyet", "content": "かさ", "dry_run": False})
    assert res["changed"] == {}  # status không phải trường sửa được; nội dung không đổi


# R4
def test_set_status_reports_count_and_validates(fake):
    err, res = call("set_status", {"ids": [1, 2], "status": "da_duyet", "dry_run": False})
    assert not err and res["changed"] == 2 and fake.writes()[-1][1] == "/api/questions/status"
    err, msg = call("set_status", {"ids": [1], "status": "xoa"})
    assert err and "status" in msg
    err, msg = call("set_status", {"ids": list(range(201)), "status": "nhap"})
    assert err and "Tối đa" in msg


# R5
def test_merge_duplicates_backs_up_first_and_needs_confirm(fake, tmp_path):
    args = {"keep_id": 1, "remove_ids": [2], "out_dir": str(tmp_path), "dry_run": False}
    err, msg = call("merge_duplicates", args)
    assert err and "confirm" in msg and fake.writes() == [] and os.listdir(tmp_path) == []
    err, res = call("merge_duplicates", {**args, "confirm": True})
    assert not err and os.path.isfile(res["backup"])
    saved = json.load(open(res["backup"], encoding="utf-8"))
    assert [x["code"] for x in saved["questions"]] == ["X-2"]
    assert fake.writes()[-1][:2] == ("POST", "/api/duplicates/merge")


def test_merge_duplicates_refuses_non_duplicates_and_keep_in_remove(fake, tmp_path):
    err, msg = call("merge_duplicates", {"keep_id": 1, "remove_ids": [3], "out_dir": str(tmp_path), "dry_run": False, "confirm": True})
    assert err and "không trùng" in msg and fake.writes() == [] and os.listdir(tmp_path) == []
    err, msg = call("merge_duplicates", {"keep_id": 1, "remove_ids": [1, 2]})
    assert err and "keep_id" in msg


# R6
def test_export_exam_needs_confirm_outdir_and_writes_zip(fake, tmp_path):
    sections = [{"matrix": [{"chapter_id": 5, "bloom": 1, "n": 1}]}]
    err, res = call("export_exam", {"sections": sections})
    assert not err and res["dry_run"] is True and res["preview"]["sections"][0]["count"] == 1
    err, msg = call("export_exam", {"sections": sections, "dry_run": False, "out_dir": str(tmp_path)})
    assert err and "confirm" in msg and not [c for c in fake.calls if c[1] == "/api/export"]
    err, msg = call("export_exam", {"sections": sections, "dry_run": False, "confirm": True})
    assert err and "out_dir" in msg and not [c for c in fake.calls if c[1] == "/api/export"]  # chưa có out_dir: không gọi export
    err, res = call("export_exam", {"sections": sections, "dry_run": False, "confirm": True, "out_dir": str(tmp_path)})
    assert not err and os.path.basename(res["file"]) == "De_001.zip" and open(res["file"], "rb").read() == b"PK-zip"


def test_export_exam_does_not_call_app_when_outdir_is_unwritable(fake, tmp_path):
    blocker = tmp_path / "file"
    blocker.write_text("x")
    err, msg = call("export_exam", {"sections": [{"matrix": []}], "dry_run": False, "confirm": True, "out_dir": str(blocker / "sub")})
    assert err and "Không ghi được" in msg
    assert not [c for c in fake.calls if c[1] == "/api/export"]  # app chưa ghi lịch sử dùng câu


def test_backup_bank_writes_json(fake, tmp_path):
    err, res = call("backup_bank", {"out_dir": str(tmp_path), "dry_run": False})
    assert not err and res["questions"] == 2 and json.load(open(res["file"], encoding="utf-8"))["questions"]


def test_restore_from_trash_and_preview_restore(fake, tmp_path):
    err, res = call("restore_from_trash", {"ids": [1, 9], "dry_run": False})
    assert not err and res["not_in_trash"] == [9] and fake.writes()[-1][1] == "/api/trash/restore"
    f = tmp_path / "b.json"
    f.write_text(json.dumps({"questions": [{"content": "a"}, {"content": "b"}]}), encoding="utf-8")
    err, res = call("preview_restore", {"file_path": str(f)})
    assert not err and res["questions_in_file"] == 2 and res["new"] == 2
    err, msg = call("preview_restore", {"file_path": str(tmp_path / "x.txt")})
    assert err and ".json" in msg


# R7
def test_no_delete_capability_besides_merge():
    src = inspect.getsource(mcp_server)
    assert 'api("DELETE"' not in src and "'DELETE'" not in src
    for route in ("bulk-delete", "/api/trash/empty", "/api/usage/reset", "/api/restore\""):
        assert route not in src
    names = [t["name"] for t in mcp_server.TOOLS]
    assert not [n for n in names if "delete" in n or n.startswith("remove_") and n != "remove_audio"]
    users = [t["name"] for t in mcp_server.TOOLS if "tool_merge_duplicates" in str(t["handler"])]
    assert users == ["merge_duplicates"]
    assert src.count('"/api/duplicates/merge"') == 1  # chỉ một chỗ xóa câu
