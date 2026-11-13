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
SERVER_INFO = {"name": "qbank", "version": "1.2.0"}


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


def tool_create_question(args):
    """Tạo một câu hỏi (mặc định trạng thái Nháp). Có audio_script thì nhờ textts đọc kịch bản thành mp3 rồi đính kèm:
    audio_target="passage" gắn vào đoạn văn (dùng chung cho mọi câu của đoạn, cần có passage), "question" gắn riêng vào câu."""
    body = {"chapter_id": _int(args.get("chapter_id"), "chapter_id"), "bloom": args.get("bloom", 1),
            "qtype": args.get("qtype", "TN"), "content": args.get("content", ""), "options": args.get("options") or [],
            "answer": args.get("answer", ""), "solution": args.get("solution", ""), "passage": args.get("passage", ""),
            "clo": args.get("clo", ""), "tags": args.get("tags", ""), "status": args.get("status", "nhap"),
            "transcript": args.get("transcript", ""), "translation": args.get("translation", "")}
    if args.get("points") is not None:
        body["points"] = args["points"]
    script = str(args.get("audio_script") or "").strip()
    audio = None
    if script and not str(body["transcript"]).strip():
        body["transcript"] = script  # kịch bản dùng tạo audio cũng là transcript của câu (trang luyện nghe hiển thị)
    if script:
        target = args.get("audio_target", "question")
        if target not in ("question", "passage"):
            raise ApiError('audio_target phải là "question" hoặc "passage"')
        if target == "passage" and not str(body["passage"]).strip():
            raise ApiError("Gắn audio vào đoạn văn thì phải có passage")
        audio = api("POST", "/api/tts", body={"script": script, "engine": args.get("audio_engine") or "voicevox",
                                              "level": args.get("audio_level")})
        body["passage_media" if target == "passage" else "media"] = [audio]
    r = api("POST", "/api/questions", body=body)
    out = {"ok": True, "id": r.get("id"), "code": r.get("code"), "warnings": r.get("warnings", [])}
    if audio:
        out["audio"] = {"file": audio.get("name"), "duration_s": audio.get("duration_s")}
    return out


AUDIO_EXT = (".mp3", ".wav", ".ogg", ".m4a")
QUESTION_FIELDS = ("chapter_id", "bloom", "qtype", "content", "options", "answer", "solution", "transcript", "translation", "passage",
                   "rubric", "points", "clo", "tags", "status", "word_limit")


def _read_file(path, exts, what):
    path = str(path or "").strip()
    if not path:
        raise ApiError(f"Thiếu đường dẫn {what}")
    if not path.lower().endswith(exts):
        raise ApiError(f"{what} phải có đuôi {', '.join(exts)}")
    if not os.path.isfile(path):
        raise ApiError(f"Không thấy file: {path}")
    with open(path, "rb") as f:
        return f.read()


def tool_import_questions(args):
    """Nhập câu hỏi từ file .docx theo mẫu 8 cột. Mặc định dry_run=true: chỉ kiểm tra.
    Câu nhập vào ở trạng thái Nháp. Đặt dry_run=false để ghi thật."""
    tid = _int(args.get("textbook_id"), "textbook_id")
    data = _read_file(args.get("file_path"), (".docx",), "File nhập")
    dry = args.get("dry_run", True) is not False
    r = api("POST", "/api/import", {"textbook_id": tid, "dry_run": "1" if dry else None}, body=data,
            content_type="application/octet-stream")
    return {"dry_run": dry, "added": r.get("added"), "errors": r.get("errors", []),
            "skipped": r.get("skipped", []), "warnings": r.get("warnings", []),
            "rows": [{"where": x.get("where"), "code": x.get("code"), "status": x.get("status"),
                      "note": x.get("note")} for x in r.get("rows", [])]}


def _is_audio(m):
    return (m or {}).get("kind") == "audio"


def _put_media(q, target, media):
    """Ghi lại danh sách đính kèm của câu (target=question) hoặc của đoạn văn (target=passage), giữ nguyên các trường khác."""
    if target == "passage":
        if not q.get("passage_id"):
            raise ApiError("Câu này không có đoạn văn")
        api("PUT", f"/api/passages/{q['passage_id']}", body={"media": media})
    else:
        body = {k: q[k] for k in QUESTION_FIELDS if k in q}
        body["media"] = media
        body["confirm_passage"] = True
        api("PUT", f"/api/questions/{q['id']}", body=body)


def tool_attach_audio(args):
    """Gắn file âm thanh có sẵn (mp3, wav, ogg, m4a) vào câu (target=question) hoặc đoạn văn của câu (target=passage).
    Mặc định THAY file nghe cũ cùng cấp (replace=true) để mỗi câu chỉ có một file nghe; replace=false thì gắn thêm.
    Giữ nguyên mọi trường khác, ảnh đính kèm và trạng thái câu."""
    qid = _int(args.get("question_id"), "question_id")
    target = args.get("target", "question")
    if target not in ("question", "passage"):
        raise ApiError('target phải là "question" hoặc "passage"')
    replace = args.get("replace", True) is not False
    raw = _read_file(args.get("file_path"), AUDIO_EXT, "File âm thanh")
    q = api("GET", f"/api/questions/{qid}")
    q.setdefault("id", qid)
    if target == "passage" and not q.get("passage_id"):
        raise ApiError("Câu này không có đoạn văn để gắn audio")
    name = os.path.basename(str(args["file_path"]))
    up = api("POST", "/api/media", {"name": name}, body=raw, content_type="application/octet-stream")
    item = {"kind": up["kind"], "file": up["file"], "name": up["name"]}
    old = (q.get("passage_media") if target == "passage" else q.get("media")) or []
    removed = [m.get("name") or m.get("file") for m in old if replace and _is_audio(m)]
    _put_media(q, target, [m for m in old if not (replace and _is_audio(m))] + [item])
    other = (q.get("media") if target == "passage" else q.get("passage_media")) or []
    out = {"ok": True, "question_id": qid, "code": q.get("code"), "target": target, "file": name, "replaced": removed}
    if any(_is_audio(m) for m in other):
        out["warning"] = ("Câu còn file nghe ở " + ("câu" if target == "passage" else "đoạn văn")
                          + ", nên câu có 2 file nghe. Dùng remove_audio để bỏ file không cần.")
    return out


def tool_remove_audio(args):
    """Xóa file nghe khỏi câu (target=question), khỏi đoạn văn (target=passage; đoạn dùng chung nên MỌI câu của đoạn mất file) hoặc cả hai.
    Chọn file bằng `file` (tên hiển thị hoặc tên lưu trong list_questions); bỏ trống thì xóa mọi file nghe ở cấp đó.
    Chỉ gỡ khỏi dữ liệu câu/đoạn; ảnh và các trường khác giữ nguyên. Đặt dry_run=true để chỉ xem sẽ xóa gì."""
    qid = _int(args.get("question_id"), "question_id")
    target = args.get("target", "question")
    if target not in ("question", "passage", "both"):
        raise ApiError('target phải là "question", "passage" hoặc "both"')
    wanted = str(args.get("file") or "").strip()
    dry = args.get("dry_run") is True
    q = api("GET", f"/api/questions/{qid}")
    q.setdefault("id", qid)
    removed = {}
    for level in (("question", "passage") if target == "both" else (target,)):
        old = (q.get("passage_media") if level == "passage" else q.get("media")) or []
        drop = [m for m in old if _is_audio(m) and (not wanted or wanted in (m.get("name"), m.get("file")))]
        if not drop:
            continue
        removed[level] = [m.get("name") or m.get("file") for m in drop]
        if not dry:
            _put_media(q, level, [m for m in old if m not in drop])
    if wanted and not removed:
        raise ApiError(f"Không thấy file nghe {wanted!r} ở {target}. Xem các file hiện có bằng get_question_audio.")
    return {"ok": True, "dry_run": dry, "question_id": qid, "code": q.get("code"), "removed": removed}


def tool_get_question_audio(args):
    """Liệt kê file nghe hiện có của một câu: ở câu và ở đoạn văn (để chọn file cần xóa/thay)."""
    qid = _int(args.get("question_id"), "question_id")
    q = api("GET", f"/api/questions/{qid}")
    f = lambda media: [{"name": m.get("name"), "file": m.get("file"), "has_script": bool(m.get("tts"))}
                       for m in (media or []) if _is_audio(m)]
    return {"question_id": qid, "code": q.get("code"), "question_audio": f(q.get("media")),
            "passage_audio": f(q.get("passage_media"))}


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
    {"name": "create_question", "handler": tool_create_question,
     "description": "Tạo một câu hỏi (mặc định Nháp, chờ người dùng duyệt). Câu nghe (聴解): đặt kịch bản hội thoại vào audio_script "
                    "(mỗi dòng một lượt nói, [1] ..., [2] ...; furigana {漢字|かな}; nghỉ [pause=800]); QBank gọi textts tạo mp3 và đính kèm. "
                    "Cần bật textts (cổng 8000). Hội thoại dùng chung nhiều câu: điền passage và audio_target=passage.",
     "inputSchema": _obj({
         "chapter_id": {"type": "integer"}, "bloom": {"type": "integer", "minimum": 1, "maximum": 5},
         "qtype": {"type": "string", "enum": ["TN", "DS", "DT", "TL"]},
         "content": {"type": "string"}, "options": {"type": "array", "items": {"type": "string"}},
         "answer": {"type": "string", "description": "TN: chữ cái A/B/C/D; DT: đáp án (nhiều đáp án cách nhau bằng |)"},
         "solution": {"type": "string"},
         "transcript": {"type": "string", "description": "Câu nghe: kịch bản/lời thoại (mặc định lấy từ audio_script)"},
         "translation": {"type": "string", "description": "Bản dịch tiếng Việt của transcript, cùng số dòng"},
         "passage": {"type": "string", "description": "Đoạn văn / kịch bản hiển thị kèm câu"},
         "clo": {"type": "string"}, "tags": {"type": "string"}, "points": {"type": "number"},
         "status": {"type": "string", "enum": ["nhap", "da_duyet", "ngung"]},
         "audio_script": {"type": "string"},
         "audio_target": {"type": "string", "enum": ["question", "passage"], "default": "question"},
         "audio_engine": {"type": "string", "description": "voicevox (mặc định), aivis, edge… theo textts"},
         "audio_level": {"type": "string", "enum": ["N5", "N4", "N3", "N2", "N1"]}},
         ["chapter_id", "content"])},
    {"name": "import_questions", "handler": tool_import_questions,
     "description": "Nhập câu hỏi từ file .docx theo mẫu 8 cột (Mã chương phải có trong giáo trình đích). Mặc định chỉ kiểm tra "
                    "(dry_run=true); chỉ đặt dry_run=false sau khi người dùng đã xác nhận. Câu nhập vào ở trạng thái Nháp.",
     "inputSchema": _obj({"textbook_id": {"type": "integer"},
                          "file_path": {"type": "string", "description": "đường dẫn .docx trên máy chạy app"},
                          "dry_run": {"type": "boolean", "default": True}}, ["textbook_id", "file_path"])},
    {"name": "attach_audio", "handler": tool_attach_audio,
     "description": "Gắn file âm thanh có sẵn (mp3, wav, ogg, m4a; tối đa 15 MB) vào một câu hoặc đoạn văn của câu. "
                    "Mặc định THAY file nghe cũ cùng cấp (replace=true) nên câu chỉ còn một file nghe. Không đổi trạng thái câu. "
                    "Nếu kết quả có warning thì câu còn file nghe ở cấp kia (câu/đoạn văn): dùng remove_audio.",
     "inputSchema": _obj({"question_id": {"type": "integer"},
                          "file_path": {"type": "string", "description": "đường dẫn file âm thanh trên máy chạy app"},
                          "target": {"type": "string", "enum": ["question", "passage"], "default": "question"},
                          "replace": {"type": "boolean", "default": True}}, ["question_id", "file_path"])},
    {"name": "get_question_audio", "handler": tool_get_question_audio,
     "description": "Liệt kê file nghe hiện có của một câu (ở câu và ở đoạn văn) để chọn file cần xóa hoặc thay.",
     "inputSchema": _obj({"question_id": {"type": "integer"}}, ["question_id"])},
    {"name": "remove_audio", "handler": tool_remove_audio,
     "description": "Xóa file nghe khỏi dữ liệu câu hoặc đoạn văn. Chọn file bằng `file` (xem get_question_audio) hoặc bỏ trống để xóa mọi "
                    "file nghe ở cấp đó. target=passage ảnh hưởng MỌI câu dùng chung đoạn văn đó. Nên chạy dry_run=true trước và chỉ xóa "
                    "thật sau khi người dùng đồng ý.",
     "inputSchema": _obj({"question_id": {"type": "integer"},
                          "target": {"type": "string", "enum": ["question", "passage", "both"], "default": "question"},
                          "file": {"type": "string"}, "dry_run": {"type": "boolean", "default": False}}, ["question_id"])},
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
