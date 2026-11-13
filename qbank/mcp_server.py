"""MCP server (stdio) cho QBank: để Claude đọc và thao tác ngân hàng câu hỏi qua API của app.

Không đọc/ghi SQLite trực tiếp: mọi thao tác đi qua API HTTP của app đang chạy (mặc định
http://127.0.0.1:8765), nên mã câu, chống trùng, sao lưu và kiểm tra nghiệp vụ vẫn được giữ.
Chỉ dùng thư viện chuẩn của Python, không cần cài thêm gói.

Chạy: mở app (python3 app.py) trước, rồi khai báo server này trong .mcp.json; xem qbank/README.md.
Biến môi trường: QBANK_URL (mặc định http://127.0.0.1:8765).
"""
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

BASE_URL = os.environ.get("QBANK_URL", "http://127.0.0.1:8765").rstrip("/")
PROTOCOL = "2024-11-05"
SERVER_INFO = {"name": "qbank", "version": "1.0.0"}


class ApiError(Exception):
    pass


def api(method, path, params=None, body=None, content_type="application/json"):
    """Gọi API của app. Trả về JSON; lỗi nghiệp vụ (HTTP 4xx/5xx có trường error) thành ApiError."""
    url = BASE_URL + path
    if params:
        url += "?" + urllib.parse.urlencode({k: v for k, v in params.items() if v not in (None, "")})
    data = None
    if body is not None:
        data = body if isinstance(body, bytes) else json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(url, data=data, method=method)
    if data is not None:
        req.add_header("Content-Type", content_type)
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read()
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf-8")).get("error", str(e))
        except Exception:
            msg = str(e)
        raise ApiError(msg)
    except urllib.error.URLError:
        raise ApiError(f"Không kết nối được app tại {BASE_URL}. Hãy mở app (python3 app.py) trước.")
    return json.loads(raw.decode("utf-8")) if raw else {}


def _int(v, name):
    try:
        return int(v)
    except (TypeError, ValueError):
        raise ApiError(f"{name} phải là số nguyên")


# ---------------------------------------------------------------- tools
def tool_list_tree(args):
    """Cây học phần → giáo trình → chương, kèm số câu của từng chương."""
    t = api("GET", "/api/tree")
    return {
        "courses": [{"id": c["id"], "name": c["name"], "code": c["code"]} for c in t["courses"]],
        "textbooks": [{"id": b["id"], "name": b["name"], "code": b["code"], "course_ids": b["course_ids"]}
                      for b in t["textbooks"]],
        "chapters": [{"id": c["id"], "textbook_id": c["textbook_id"], "code": c["code"],
                      "title": c["title"], "questions": c["count"]} for c in t["chapters"]],
    }


def tool_list_questions(args):
    """Tìm câu hỏi. Lọc theo học phần, giáo trình, chương, mức Bloom, dạng (TN/DS/DT/TL), trạng thái, từ khóa."""
    per = min(max(int(args.get("limit", 20)), 1), 50)
    page = max(int(args.get("page", 0)), 0)
    params = {
        "course_id": args.get("course_id"), "textbook_id": args.get("textbook_id"),
        "chapter_id": args.get("chapter_id"), "bloom": args.get("bloom"), "qtype": args.get("qtype"),
        "status": args.get("status"), "q": args.get("query"), "per": per, "page": page,
    }
    r = api("GET", "/api/questions", params)
    items = [{"code": i["code"], "qtype": i["qtype"], "bloom": i["bloom"], "status": i["status"],
              "course": i.get("courses_all") or i.get("course"), "textbook": i.get("textbook"),
              "chapter": i.get("chapter_label"), "content": (i.get("content") or "")[:300],
              "used_count": i.get("used_count", 0)} for i in r["items"]]
    return {"total": r["total"], "page": page, "per": per, "items": items}


def tool_create_course(args):
    name = str(args.get("name", "")).strip()
    if not name:
        raise ApiError("Thiếu tên học phần")
    api("POST", "/api/courses", body={"name": name, "code": str(args.get("code", "")).strip()})
    return {"ok": True, "name": name}


def tool_create_textbook(args):
    """Tạo giáo trình trong học phần. Nếu giáo trình cùng tên đã có thì gắn thêm vào học phần đó (dùng chung)."""
    cid = _int(args.get("course_id"), "course_id")
    name = str(args.get("name", "")).strip()
    if not name:
        raise ApiError("Thiếu tên giáo trình")
    r = api("POST", "/api/textbooks", body={"course_id": cid, "name": name, "code": str(args.get("code", "")).strip()})
    return {"ok": True, "textbook_id": r.get("id")}


def tool_link_textbook(args):
    """Gắn giáo trình có sẵn vào thêm một học phần (giáo trình dùng chung)."""
    cid = _int(args.get("course_id"), "course_id")
    tid = _int(args.get("textbook_id"), "textbook_id")
    api("POST", "/api/course_textbooks", body={"course_id": cid, "textbook_id": tid})
    return {"ok": True}


def tool_create_chapter(args):
    tid = _int(args.get("textbook_id"), "textbook_id")
    code = str(args.get("code", "")).strip()
    if not code:
        raise ApiError("Thiếu mã chương")
    api("POST", "/api/chapters", body={"textbook_id": tid, "code": code, "title": str(args.get("title", "")).strip()})
    return {"ok": True, "code": code}


def tool_import_chapters(args):
    """Nhập hàng loạt chương từ nội dung CSV (cột Mã, Tên chương; dòng đầu có thể là tiêu đề).
    Mặc định dry_run=true: chỉ kiểm tra và báo thêm/bỏ qua/lỗi. Đặt dry_run=false để ghi thật."""
    tid = _int(args.get("textbook_id"), "textbook_id")
    text = args.get("csv")
    if not isinstance(text, str) or not text.strip():
        raise ApiError("Thiếu nội dung csv")
    dry = args.get("dry_run", True) is not False
    params = {"textbook_id": tid, "dry_run": "1" if dry else None}
    r = api("POST", "/api/chapters/import", params, body=text.encode("utf-8"), content_type="text/csv; charset=utf-8")
    return {"dry_run": dry, "counts": r["counts"],
            "rows": [{"line": x["line"], "code": x["code"], "title": x["title"], "status": x["status"],
                      "note": x["note"]} for x in r["rows"] if x["status"] != "thêm" or dry]}


def _obj(props, required=()):
    return {"type": "object", "properties": props, "required": list(required), "additionalProperties": False}


TOOLS = [
    {"name": "list_tree", "handler": tool_list_tree,
     "description": "Liệt kê học phần, giáo trình và chương (kèm số câu). Dùng để lấy id trước khi thao tác.",
     "inputSchema": _obj({})},
    {"name": "list_questions", "handler": tool_list_questions,
     "description": "Tìm câu hỏi trong ngân hàng. Trả về tối đa 50 câu mỗi trang.",
     "inputSchema": _obj({
         "course_id": {"type": "integer"}, "textbook_id": {"type": "integer"}, "chapter_id": {"type": "integer"},
         "bloom": {"type": "integer", "minimum": 1, "maximum": 5},
         "qtype": {"type": "string", "enum": ["TN", "DS", "DT", "TL"]},
         "status": {"type": "string", "enum": ["nhap", "da_duyet", "ngung"]},
         "query": {"type": "string", "description": "từ khóa trong nội dung, phương án, lời giải, đáp án, mã"},
         "limit": {"type": "integer", "minimum": 1, "maximum": 50},
         "page": {"type": "integer", "minimum": 0}})},
    {"name": "create_course", "handler": tool_create_course,
     "description": "Tạo học phần mới.",
     "inputSchema": _obj({"name": {"type": "string"}, "code": {"type": "string"}}, ["name"])},
    {"name": "create_textbook", "handler": tool_create_textbook,
     "description": "Tạo giáo trình trong học phần. Tên trùng giáo trình đã có sẽ gắn thêm vào học phần (không tạo bản mới).",
     "inputSchema": _obj({"course_id": {"type": "integer"}, "name": {"type": "string"}, "code": {"type": "string"}},
                         ["course_id", "name"])},
    {"name": "link_textbook", "handler": tool_link_textbook,
     "description": "Gắn một giáo trình có sẵn vào thêm một học phần (dùng chung).",
     "inputSchema": _obj({"course_id": {"type": "integer"}, "textbook_id": {"type": "integer"}},
                         ["course_id", "textbook_id"])},
    {"name": "create_chapter", "handler": tool_create_chapter,
     "description": "Thêm một chương vào giáo trình. Mã chương phải duy nhất trong giáo trình.",
     "inputSchema": _obj({"textbook_id": {"type": "integer"}, "code": {"type": "string"}, "title": {"type": "string"}},
                         ["textbook_id", "code"])},
    {"name": "import_chapters", "handler": tool_import_chapters,
     "description": "Nhập hàng loạt chương từ nội dung CSV (cột Mã, Tên chương). Mặc định chỉ kiểm tra (dry_run=true); "
                    "chỉ đặt dry_run=false sau khi người dùng đã xác nhận kết quả kiểm tra.",
     "inputSchema": _obj({"textbook_id": {"type": "integer"}, "csv": {"type": "string"},
                          "dry_run": {"type": "boolean", "default": True}}, ["textbook_id", "csv"])},
]
TOOL_BY_NAME = {t["name"]: t for t in TOOLS}


# ---------------------------------------------------------------- JSON-RPC (MCP stdio)
def ok(req_id, result):
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def err(req_id, code, message):
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


def handle(msg):
    """Xử lý một thông điệp JSON-RPC. Trả về phản hồi, hoặc None với thông báo (không có id)."""
    method = msg.get("method")
    req_id = msg.get("id")
    if req_id is None:  # thông báo (ví dụ notifications/initialized): không trả lời
        return None
    if method == "initialize":
        client_version = (msg.get("params") or {}).get("protocolVersion", PROTOCOL)
        return ok(req_id, {"protocolVersion": client_version, "capabilities": {"tools": {}}, "serverInfo": SERVER_INFO})
    if method == "ping":
        return ok(req_id, {})
    if method == "tools/list":
        return ok(req_id, {"tools": [{k: t[k] for k in ("name", "description", "inputSchema")} for t in TOOLS]})
    if method == "tools/call":
        params = msg.get("params") or {}
        tool = TOOL_BY_NAME.get(params.get("name"))
        if not tool:
            return err(req_id, -32602, f"Không có công cụ {params.get('name')!r}")
        try:
            result = tool["handler"](params.get("arguments") or {})
            return ok(req_id, {"content": [{"type": "text", "text": json.dumps(result, ensure_ascii=False, indent=1)}],
                               "isError": False})
        except ApiError as e:
            return ok(req_id, {"content": [{"type": "text", "text": str(e)}], "isError": True})
        except Exception as e:  # lỗi bất ngờ vẫn trả về cho Claude, không làm sập server
            return ok(req_id, {"content": [{"type": "text", "text": f"{type(e).__name__}: {e}"}], "isError": True})
    return err(req_id, -32601, f"Không hỗ trợ phương thức {method!r}")


def main():
    # Windows dùng cp1252 cho stdin/stdout mặc định, sẽ lỗi với tiếng Việt: luôn đọc/ghi UTF-8
    sys.stdin.reconfigure(encoding="utf-8")
    sys.stdout.reconfigure(encoding="utf-8")
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except json.JSONDecodeError:
            resp = err(None, -32700, "JSON không hợp lệ")
        else:
            resp = handle(msg)
        if resp is not None:
            sys.stdout.write(json.dumps(resp) + "\n")  # ensure_ascii: đầu ra toàn ASCII, an toàn trên mọi bảng mã
            sys.stdout.flush()


if __name__ == "__main__":
    main()
