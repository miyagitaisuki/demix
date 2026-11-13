"""MCP server (stdio) cho QBank: để Claude đọc và thao tác ngân hàng câu hỏi qua API của app.

Không đọc/ghi SQLite trực tiếp: mọi thao tác đi qua API HTTP của app đang chạy (mặc định
http://127.0.0.1:8765), nên mã câu, chống trùng, sao lưu và kiểm tra nghiệp vụ vẫn được giữ.
Chỉ dùng thư viện chuẩn của Python, không cần cài thêm gói.

Chạy: mở app (python3 app.py) trước, rồi khai báo server này trong .mcp.json; xem qbank/README.md.
Biến môi trường: QBANK_URL (mặc định http://127.0.0.1:8765).
"""
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request

BASE_URL = os.environ.get("QBANK_URL", "http://127.0.0.1:8765").rstrip("/")
PROTOCOL = "2024-11-05"
SERVER_INFO = {"name": "qbank", "version": "1.3.0"}


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


def _auto_check(body, allow_similar):
    """Kiểm trùng TỰ ĐỘNG trước khi tạo câu (và trước khi tốn thời gian tạo audio). Trả về cảnh báo (list) hoặc raise ApiError.
    App cũ chưa có route kiểm trùng thì bỏ qua và báo trong cảnh báo."""
    item = {k: body[k] for k in ("chapter_id", "qtype", "content", "options")}
    try:
        res = api("POST", "/api/check-candidates", body={"items": [item]})["results"][0]
    except ApiError as e:
        if "không tìm thấy" in str(e):
            return ["Chưa kiểm câu gần giống: app chưa có route kiểm trùng (cần build lại QBank.exe)."]
        raise
    codes = lambda lst: ", ".join(f"{x['code']}" + (f" ({round(x['ratio'] * 100)}%)" if "ratio" in x else "") for x in lst)
    if res["status"] == "trùng":
        raise ApiError(f"Câu đã có trong cùng chương ({codes(res['exact'])}). Không tạo; hãy viết câu khác.")
    if res["status"] == "gần giống" and not allow_similar:
        raise ApiError(f"Câu gần giống câu đã có: {codes(res['similar'])}. Hãy đổi chủ đề/số liệu/hội thoại, hoặc hỏi người dùng "
                       "rồi gọi lại với allow_similar=true.")
    if res["status"] == "trùng chương khác":
        return [f"Trùng nội dung với câu {codes(res['exact'])} ở chương khác"]
    return []


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
    if args.get("confirm_passage") is True:  # nhãn đoạn văn gần giống đoạn có sẵn: chỉ tạo đoạn mới khi người dùng đã xác nhận
        body["confirm_passage"] = True
    check_warnings = _auto_check(body, args.get("allow_similar") is True)
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
    out = {"ok": True, "id": r.get("id"), "code": r.get("code"), "warnings": r.get("warnings", []) + check_warnings}
    if audio:
        out["audio"] = {"file": audio.get("name"), "duration_s": audio.get("duration_s")}
    return out


def tool_check_duplicates(args):
    """Kiểm trùng nhanh một lô câu DỰ ĐỊNH tạo (chưa ghi gì). Trả về cho từng câu: trùng cùng chương / chương khác /
    gần giống (>= 90%) với câu trong ngân hàng, và trùng với câu khác trong chính lô."""
    items = args.get("items")
    if not isinstance(items, list) or not items:
        raise ApiError("Thiếu danh sách items")
    r = api("POST", "/api/check-candidates", body={"items": items})
    res = r["results"]
    return {"summary": {k: sum(1 for x in res if x["status"] == k) for k in ("mới", "gần giống", "trùng chương khác", "trùng")},
            "results": res}


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


# ---------------------------------------------------------------- công cụ rà soát, sửa, ra đề, sao lưu
STATUS_VALUES = ("nhap", "da_duyet", "ngung")
EDITABLE_FIELDS = ("content", "options", "answer", "solution", "passage", "bloom", "qtype", "clo", "tags", "points",
                   "word_limit", "transcript", "translation", "rubric")


def api_raw(method, path, body=None):
    """Như api() nhưng trả về (bytes, headers): dùng cho file zip đề và file sao lưu."""
    data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
    req = urllib.request.Request(BASE_URL + path, data=data, method=method)
    if data is not None:
        req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=300) as resp:
            return resp.read(), resp.headers
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf-8")).get("error", str(e))
        except Exception:
            msg = str(e)
        raise ApiError(msg)
    except urllib.error.URLError:
        raise ApiError(f"Không kết nối được app tại {BASE_URL}. Hãy mở app (python3 app.py) trước.")


def _out_dir(args):
    """Thư mục ghi file do người dùng chỉ định (bắt buộc). Kiểm tra ghi được TRƯỚC khi làm việc có tác dụng phụ."""
    d = str(args.get("out_dir") or "").strip()
    if not d:
        raise ApiError("Thiếu out_dir: hãy hỏi người dùng muốn lưu file ở thư mục nào")
    d = os.path.abspath(d)
    try:
        os.makedirs(d, exist_ok=True)
        probe = os.path.join(d, ".qbank_write_test")
        with open(probe, "w") as f:
            f.write("ok")
        os.remove(probe)
    except OSError as e:
        raise ApiError(f"Không ghi được vào thư mục {d}: {e}")
    return d


def _stamp():
    import datetime
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def _ids(args, key="ids", limit=200):
    ids = args.get(key)
    if not isinstance(ids, list) or not ids:
        raise ApiError(f"Thiếu danh sách {key}")
    if len(ids) > limit:
        raise ApiError(f"Tối đa {limit} mục mỗi lần")
    return [_int(i, key) for i in ids]


def _writes(args):
    """Công cụ ghi: mặc định dry_run (chỉ xem trước). dry_run=false mới ghi thật."""
    return args.get("dry_run", True) is False


# --- rà soát (chỉ đọc)
def tool_bank_stats(args):
    """Thống kê ngân hàng: tổng số câu, theo dạng, và theo từng chương (số câu theo mức Bloom)."""
    s = api("GET", "/api/stats")
    return {"total": s["total"], "by_type": s["by_type"],
            "chapters": [{"id": c["id"], "label": c["label"], "total": c["total"], "by_bloom": c["counts"]}
                         for c in s["chapters"]]}


def tool_find_duplicates(args):
    """Các nhóm câu trùng nội dung y hệt trong ngân hàng (chỉ báo cáo, không gộp)."""
    limit = min(max(int(args.get("limit", 20)), 1), 100)
    groups = api("GET", "/api/duplicates")
    groups = groups if isinstance(groups, list) else groups.get("groups", [])
    return {"groups": len(groups), "shown": groups[:limit]}


def tool_find_similar(args):
    """Các cặp câu gần giống (>= 90%, cùng dạng) để rà soát (chỉ báo cáo)."""
    limit = min(max(int(args.get("limit", 20)), 1), 100)
    pairs = api("GET", "/api/similar")
    pairs = pairs if isinstance(pairs, list) else pairs.get("pairs", [])
    return {"pairs": len(pairs), "shown": pairs[:limit]}


def tool_difficulty_report(args):
    """Độ khó thực tế từ kết quả thi đã nhập: trung bình theo mức Bloom và các câu dễ/khó bất thường."""
    return api("GET", "/api/difficulty")


# --- sửa và duyệt
def tool_update_question(args):
    """Sửa một câu: chỉ các trường được nêu thay đổi; trường khác, file đính kèm và trạng thái giữ nguyên.
    Mặc định dry_run: chỉ cho xem thay đổi. Không đổi trạng thái (dùng set_status)."""
    qid = _int(args.get("question_id"), "question_id")
    changes = {k: args[k] for k in EDITABLE_FIELDS if k in args}
    if not changes:
        raise ApiError("Chưa nêu trường nào cần sửa: " + ", ".join(EDITABLE_FIELDS))
    q = api("GET", f"/api/questions/{qid}")
    diff = {k: {"cũ": q.get(k), "mới": v} for k, v in changes.items() if q.get(k) != v}
    if diff and _writes(args):
        body = {k: q[k] for k in QUESTION_FIELDS if k in q}
        body.update(changes)
        body["media"] = q.get("media") or []
        body["confirm_passage"] = args.get("confirm_passage") is True
        api("PUT", f"/api/questions/{qid}", body=body)
    return {"ok": True, "dry_run": not _writes(args), "question_id": qid, "code": q.get("code"), "changed": diff}


def tool_set_status(args):
    """Đổi trạng thái hàng loạt (nhap, da_duyet, ngung). Chuyển sang da_duyet CHỈ khi người dùng đã bảo duyệt các câu này.
    Mặc định dry_run: liệt kê các câu sẽ đổi."""
    ids = _ids(args)
    status = args.get("status")
    if status not in STATUS_VALUES:
        raise ApiError("status phải là nhap, da_duyet hoặc ngung")
    rows = []
    for i in ids[:100]:
        q = api("GET", f"/api/questions/{i}")
        rows.append({"id": i, "code": q.get("code"), "từ": q.get("status")})
    if _writes(args):
        r = api("POST", "/api/questions/status", body={"status": status, "ids": ids})
        return {"ok": True, "dry_run": False, "status": status, "changed": r.get("updated", r.get("changed", len(ids))),
                "questions": rows}
    return {"ok": True, "dry_run": True, "status": status, "would_change": len(ids), "questions": rows}


# --- ra đề và xuất đề
def _sections(args):
    s = args.get("sections")
    if not isinstance(s, list) or not s:
        raise ApiError("Thiếu sections (mỗi phần có matrix [{chapter_id, bloom, n}] và/hoặc passages)")
    return s


def tool_preview_exam(args):
    """Xem trước các câu sẽ được chọn cho đề (không ghi lịch sử dùng, không tạo file)."""
    r = api("POST", "/api/exam/preview", body={"sections": _sections(args)})
    parts = []
    for i, s in enumerate(r["sections"]):
        parts.append({"section": i + 1, "count": len(s["items"]),
                      "items": [{"id": it.get("id"), "code": it.get("code"), "qtype": it.get("qtype"), "bloom": it.get("bloom"),
                                 "last_used": it.get("last_used"), "content": (it.get("content") or "")[:80]}
                                for it in s["items"]]})
    return {"sections": parts}


def _filename(headers, fallback):
    cd = headers.get("Content-Disposition", "") or ""
    m = re.search(r"filename\*=UTF-8''([^;]+)", cd) or re.search(r'filename="?([^";]+)"?', cd)
    name = urllib.parse.unquote(m.group(1)) if m else fallback
    return os.path.basename(name) or fallback


def tool_export_exam(args):
    """Xuất đề (zip gồm .docx, đáp án, âm thanh) vào out_dir. Xuất THẬT ghi file và GHI LỊCH SỬ DÙNG CÂU (câu bị đánh dấu đã ra đề),
    nên cần dry_run=false VÀ confirm=true sau khi người dùng đồng ý. Mặc định dry_run: chỉ xem trước các câu sẽ chọn."""
    sections = _sections(args)
    body = {k: args[k] for k in ("template", "variants", "start", "prefix", "meta", "answer_key", "key_separate", "furigana",
                                 "shuffle", "seed", "clo_mode", "listening", "picked", "key_template") if k in args}
    body["sections"] = sections
    if not _writes(args):
        return {"dry_run": True, "note": "Chưa ghi gì. Để xuất thật: dry_run=false và confirm=true, kèm out_dir.",
                "preview": tool_preview_exam({"sections": sections})}
    if args.get("confirm") is not True:
        raise ApiError("Xuất thật cần confirm=true (sẽ ghi file và đánh dấu các câu đã ra đề). Hãy hỏi người dùng trước.")
    d = _out_dir(args)  # kiểm tra ghi được TRƯỚC khi app ghi lịch sử dùng câu
    if "record_usage" in args:
        body["record_usage"] = bool(args["record_usage"])
    data, headers = api_raw("POST", "/api/export", body)
    path = os.path.join(d, _filename(headers, f"de_{_stamp()}.zip"))
    with open(path, "wb") as f:
        f.write(data)
    return {"ok": True, "dry_run": False, "file": path, "bytes": len(data), "variants": headers.get("X-Variants"),
            "seed": headers.get("X-Seed"), "reused_questions": headers.get("X-Reused")}


# --- sao lưu và khôi phục
def tool_backup_bank(args):
    """Sao lưu toàn bộ ngân hàng ra file JSON trong out_dir (không đổi gì trong app). Mặc định dry_run: chỉ báo số câu sẽ sao lưu."""
    data, _ = api_raw("GET", "/api/backup")
    info = json.loads(data.decode("utf-8"))
    n = len(info.get("questions", []))
    if not _writes(args):
        return {"dry_run": True, "questions": n, "bytes": len(data)}
    d = _out_dir(args)
    path = os.path.join(d, f"qbank_backup_{_stamp()}.json")
    with open(path, "wb") as f:
        f.write(data)
    return {"ok": True, "dry_run": False, "file": path, "questions": n, "bytes": len(data)}


def tool_list_trash(args):
    """Câu trong thùng rác (đã xóa trong app) có thể khôi phục."""
    return {"items": api("GET", "/api/trash")}


def tool_restore_from_trash(args):
    """Khôi phục câu từ thùng rác (câu trùng nội dung sẽ không khôi phục được và ở lại thùng rác). Mặc định dry_run."""
    ids = _ids(args)
    trash = {t["id"]: t for t in api("GET", "/api/trash")}
    found = [{"id": i, "code": trash[i]["code"], "content": trash[i]["content"][:80]} for i in ids if i in trash]
    missing = [i for i in ids if i not in trash]
    if not _writes(args):
        return {"dry_run": True, "would_restore": found, "not_in_trash": missing}
    r = api("POST", "/api/trash/restore", body={"ids": [t["id"] for t in found]})
    return {"ok": True, "dry_run": False, "result": r, "not_in_trash": missing}


def tool_preview_restore(args):
    """Xem trước việc nhập lại một file sao lưu JSON: bao nhiêu câu đã có sẵn trong ngân hàng, bao nhiêu câu mới. KHÔNG ghi gì;
    khôi phục thật từ file sao lưu làm trong app."""
    path = str(args.get("file_path") or "").strip()
    if not path.lower().endswith(".json") or not os.path.isfile(path):
        raise ApiError("file_path phải là file sao lưu .json có thật")
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    items = data.get("questions") if isinstance(data, dict) else data
    if not isinstance(items, list):
        raise ApiError("File sao lưu không đúng định dạng")
    present, new, near = 0, 0, 0
    for i in range(0, len(items), 100):
        batch = [{"chapter_id": 0, "qtype": it.get("qtype", "TN"), "content": it.get("content", ""),
                  "options": it.get("options") or []} for it in items[i:i + 100]]
        for r in api("POST", "/api/check-candidates", body={"items": batch})["results"]:
            if r["status"] in ("trùng", "trùng chương khác"):
                present += 1
            elif r["status"] == "gần giống":
                near += 1
            else:
                new += 1
    return {"file": path, "questions_in_file": len(items), "already_in_bank": present, "near_duplicates": near, "new": new}


DUP_IGNORE = re.compile(r"[\s。、，．,.!?！？「」『』“”\"'()（）]+")  # cùng quy tắc với dup_key của app (app vẫn kiểm lại)


def _dup_key(q):
    return DUP_IGNORE.sub("", (str(q.get("content") or "") + "".join(q.get("options") or [])).lower())


def tool_merge_duplicates(args):
    """Gộp câu trùng: giữ keep_id, XÓA HẲN các câu remove_ids (cùng nội dung; không vào thùng rác; lịch sử dùng chuyển sang câu giữ).
    Trước khi xóa tự lưu bản sao các câu sẽ mất vào out_dir. Mặc định dry_run; xóa thật cần dry_run=false VÀ confirm=true
    sau khi người dùng đồng ý từng lần."""
    keep = _int(args.get("keep_id"), "keep_id")
    removes = _ids(args, "remove_ids", limit=20)
    if keep in removes:
        raise ApiError("keep_id không được nằm trong remove_ids")
    kept = api("GET", f"/api/questions/{keep}")
    gone = [api("GET", f"/api/questions/{i}") for i in removes]
    for i, g in zip(removes, gone):
        if _dup_key(g) != _dup_key(kept):
            raise ApiError(f"Câu {g.get('code')} không trùng nội dung với câu giữ {kept.get('code')}: không gộp.")
    plan = {"keep": {"id": keep, "code": kept.get("code")}, "remove": [{"id": i, "code": g.get("code")} for i, g in zip(removes, gone)]}
    if not _writes(args):
        return {"dry_run": True, "plan": plan, "note": "Chưa xóa gì. Xóa thật: dry_run=false, confirm=true, kèm out_dir để lưu bản sao."}
    if args.get("confirm") is not True:
        raise ApiError("Gộp thật cần confirm=true (câu bị gộp bị xóa hẳn). Hãy hỏi người dùng trước.")
    d = _out_dir(args)
    backup = os.path.join(d, f"truoc_gop_{_stamp()}.json")
    with open(backup, "w", encoding="utf-8") as f:
        json.dump({"app": "QBank", "version": 4, "note": "Câu bị xóa khi gộp", "keep": keep, "questions": gone}, f, ensure_ascii=False, indent=1)
    if os.path.getsize(backup) == 0:
        raise ApiError("Không lưu được bản sao, hủy gộp")
    r = api("POST", "/api/duplicates/merge", body={"keep": keep, "remove": removes})
    return {"ok": True, "dry_run": False, "result": r, "backup": backup}


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
     "description": "Tạo một câu hỏi (mặc định Nháp, chờ người dùng duyệt). TỰ KIỂM TRÙNG trước khi ghi: câu trùng cùng chương bị từ chối, câu gần giống (>= 90%) bị chặn trừ khi allow_similar=true sau khi người dùng đồng ý. Câu nghe (聴解): đặt kịch bản hội thoại vào audio_script "
                    "(mỗi dòng một lượt nói, [1] ..., [2] ...; furigana {漢字|かな}; nghỉ [pause=800]); QBank gọi textts tạo mp3 và đính kèm. "
                    "Cần bật textts (cổng 8000). Nếu báo đoạn văn gần giống đoạn có sẵn: dùng đúng chữ của đoạn đó để dùng chung, hoặc hỏi người dùng rồi gọi lại với confirm_passage=true. Hội thoại dùng chung nhiều câu: điền passage và audio_target=passage.",
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
         "audio_level": {"type": "string", "enum": ["N5", "N4", "N3", "N2", "N1"]},
         "allow_similar": {"type": "boolean", "default": False, "description": "Chỉ đặt true khi người dùng đồng ý tạo câu dù gần giống câu đã có"},
         "confirm_passage": {"type": "boolean", "default": False, "description": "Đặt true khi app báo nhãn/đoạn văn gần giống đoạn có sẵn trong chương và người dùng đồng ý tạo đoạn mới (hoặc dùng đúng chữ của đoạn có sẵn để dùng chung)"}},
         ["chapter_id", "content"])},
    {"name": "check_duplicates", "handler": tool_check_duplicates,
     "description": "GỌI TRƯỚC create_question: kiểm trùng một lô câu dự định tạo (tối đa 100) với toàn ngân hàng, chưa ghi gì. "
                    "status: mới (tạo được), gần giống (>= 90%, nên đổi chủ đề/số liệu hoặc hỏi người dùng), "
                    "trùng chương khác (chỉ cảnh báo), trùng (cùng chương: app sẽ từ chối, hãy bỏ hoặc viết câu khác). "
                    "Mỗi kết quả có exact/similar (mã câu đã có) và in_batch (trùng với câu khác trong lô, theo index).",
     "inputSchema": _obj({"items": {"type": "array", "maxItems": 100, "items": _obj({
         "chapter_id": {"type": "integer"}, "qtype": {"type": "string", "enum": ["TN", "DS", "DT", "TL"]},
         "content": {"type": "string"}, "options": {"type": "array", "items": {"type": "string"}}}, ["chapter_id", "content"])}},
         ["items"])},
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
    # ---- rà soát (chỉ đọc)
    {"name": "bank_stats", "handler": tool_bank_stats,
     "description": "CHỈ ĐỌC. Thống kê ngân hàng: tổng số câu, theo dạng, và theo chương (số câu theo mức Bloom).",
     "inputSchema": _obj({})},
    {"name": "find_duplicates", "handler": tool_find_duplicates,
     "description": "CHỈ ĐỌC. Các nhóm câu trùng nội dung y hệt trong ngân hàng. Chỉ báo cáo; gộp bằng merge_duplicates khi người dùng đồng ý.",
     "inputSchema": _obj({"limit": {"type": "integer", "minimum": 1, "maximum": 100}})},
    {"name": "find_similar", "handler": tool_find_similar,
     "description": "CHỈ ĐỌC. Các cặp câu gần giống (>= 90%, cùng dạng) để rà soát. Chỉ báo cáo.",
     "inputSchema": _obj({"limit": {"type": "integer", "minimum": 1, "maximum": 100}})},
    {"name": "difficulty_report", "handler": tool_difficulty_report,
     "description": "CHỈ ĐỌC. Độ khó thực tế từ kết quả thi đã nhập: trung bình theo mức Bloom và các câu dễ/khó bất thường.",
     "inputSchema": _obj({})},
    # ---- sửa và duyệt
    {"name": "update_question", "handler": tool_update_question,
     "description": "Sửa một câu: chỉ các trường được nêu thay đổi, trường khác/file đính kèm/trạng thái giữ nguyên. Mặc định dry_run=true "
                    "(chỉ cho xem thay đổi cũ/mới); dry_run=false mới ghi. KHÔNG đổi trạng thái (dùng set_status).",
     "inputSchema": _obj({"question_id": {"type": "integer"}, "content": {"type": "string"},
                          "options": {"type": "array", "items": {"type": "string"}}, "answer": {"type": "string"},
                          "solution": {"type": "string"}, "passage": {"type": "string"},
                          "bloom": {"type": "integer", "minimum": 1, "maximum": 5},
                          "qtype": {"type": "string", "enum": ["TN", "DS", "DT", "TL"]}, "clo": {"type": "string"},
                          "tags": {"type": "string"}, "points": {"type": "number"}, "word_limit": {"type": "integer"},
                          "transcript": {"type": "string"}, "translation": {"type": "string"},
                          "rubric": {"type": "array", "items": {}}, "confirm_passage": {"type": "boolean"},
                          "dry_run": {"type": "boolean", "default": True}}, ["question_id"])},
    {"name": "set_status", "handler": tool_set_status,
     "description": "Đổi trạng thái hàng loạt (nhap, da_duyet, ngung; tối đa 200 câu). Chuyển sang da_duyet CHỈ khi người dùng đã bảo duyệt "
                    "đúng các câu này; không tự duyệt câu Claude vừa tạo. Mặc định dry_run=true: liệt kê các câu sẽ đổi.",
     "inputSchema": _obj({"ids": {"type": "array", "items": {"type": "integer"}, "maxItems": 200},
                          "status": {"type": "string", "enum": ["nhap", "da_duyet", "ngung"]},
                          "dry_run": {"type": "boolean", "default": True}}, ["ids", "status"])},
    # ---- ra đề và xuất đề
    {"name": "preview_exam", "handler": tool_preview_exam,
     "description": "CHỈ ĐỌC. Xem trước các câu sẽ được chọn cho đề theo cấu trúc sections (matrix chương x mức Bloom x số câu, passages). "
                    "Không ghi lịch sử dùng câu, không tạo file.",
     "inputSchema": _obj({"sections": {"type": "array", "items": {"type": "object"}}}, ["sections"])},
    {"name": "export_exam", "handler": tool_export_exam,
     "description": "Xuất đề thành file zip vào out_dir. Xuất THẬT ghi file và GHI LỊCH SỬ DÙNG CÂU (câu bị đánh dấu đã ra đề): cần dry_run=false "
                    "VÀ confirm=true sau khi người dùng đồng ý. Mặc định dry_run=true: chỉ xem trước các câu sẽ chọn. out_dir bắt buộc khi xuất thật.",
     "inputSchema": _obj({"sections": {"type": "array", "items": {"type": "object"}}, "template": {"type": "string"},
                          "variants": {"type": "integer", "minimum": 1}, "start": {"type": "integer"}, "prefix": {"type": "string"},
                          "meta": {"type": "object"}, "answer_key": {"type": "boolean"}, "key_separate": {"type": "boolean"},
                          "furigana": {"type": "string"}, "shuffle": {"type": "boolean"}, "seed": {"type": "integer"},
                          "clo_mode": {"type": "boolean"}, "listening": {"type": "object"}, "picked": {"type": "array"},
                          "key_template": {"type": "string"}, "record_usage": {"type": "boolean"},
                          "out_dir": {"type": "string", "description": "thư mục lưu file zip (do người dùng chỉ định)"},
                          "dry_run": {"type": "boolean", "default": True}, "confirm": {"type": "boolean", "default": False}},
                         ["sections"])},
    # ---- sao lưu và khôi phục
    {"name": "backup_bank", "handler": tool_backup_bank,
     "description": "Sao lưu toàn bộ ngân hàng ra file JSON trong out_dir (không đổi gì trong app). Mặc định dry_run=true: chỉ báo số câu.",
     "inputSchema": _obj({"out_dir": {"type": "string"}, "dry_run": {"type": "boolean", "default": True}})},
    {"name": "list_trash", "handler": tool_list_trash,
     "description": "CHỈ ĐỌC. Câu trong thùng rác (đã xóa trong app) có thể khôi phục.",
     "inputSchema": _obj({})},
    {"name": "restore_from_trash", "handler": tool_restore_from_trash,
     "description": "Khôi phục câu từ thùng rác (id lấy từ list_trash). Câu trùng nội dung không khôi phục được. Mặc định dry_run=true.",
     "inputSchema": _obj({"ids": {"type": "array", "items": {"type": "integer"}, "maxItems": 200},
                          "dry_run": {"type": "boolean", "default": True}}, ["ids"])},
    {"name": "preview_restore", "handler": tool_preview_restore,
     "description": "CHỈ ĐỌC. Xem trước việc nhập lại file sao lưu .json: bao nhiêu câu đã có trong ngân hàng, gần giống, mới. Không ghi gì; "
                    "khôi phục thật từ file sao lưu do người dùng làm trong app.",
     "inputSchema": _obj({"file_path": {"type": "string"}}, ["file_path"])},
    {"name": "merge_duplicates", "handler": tool_merge_duplicates,
     "description": "NGUY HIỂM, không hoàn tác trong app. Gộp câu trùng: giữ keep_id, XÓA HẲN remove_ids (cùng nội dung, tối đa 20; không vào "
                    "thùng rác). Tự lưu bản sao các câu sẽ mất vào out_dir trước khi xóa. Mặc định dry_run=true; xóa thật cần dry_run=false "
                    "VÀ confirm=true sau khi người dùng đồng ý từng lần.",
     "inputSchema": _obj({"keep_id": {"type": "integer"},
                          "remove_ids": {"type": "array", "items": {"type": "integer"}, "maxItems": 20},
                          "out_dir": {"type": "string"}, "dry_run": {"type": "boolean", "default": True},
                          "confirm": {"type": "boolean", "default": False}}, ["keep_id", "remove_ids"])},
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
