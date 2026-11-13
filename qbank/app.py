#!/usr/bin/env python3
"""QBank - ngân hàng câu hỏi offline, giao diện web nhẹ chạy trên máy cục bộ.

Cấu trúc phân loại: Học phần -> Giáo trình -> Chương (mã duy nhất trong giáo trình) -> Câu hỏi.

Chạy:  python3 app.py   rồi mở http://127.0.0.1:8765
Phụ thuộc: python-docx (đọc bảng câu hỏi và xuất đề .docx).
Dữ liệu: data/qbank.db (SQLite). Template Word: templates/*.docx
"""
import base64
import csv
import hashlib
import io
import statistics
import json
import os
import random
import re
import sqlite3
import shutil
import subprocess
import sys
import threading
import time
import unicodedata
import webbrowser
import zipfile
from contextlib import closing
from datetime import datetime
from difflib import SequenceMatcher
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote
from urllib import request as urlreq
from urllib.error import HTTPError, URLError

from docx import Document
from html import escape as hesc
from xml.sax.saxutils import escape
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, "frozen", False):  # chạy từ file đóng gói PyInstaller
    RES_DIR = sys._MEIPASS  # tài nguyên đi kèm (static, templates mẫu) nằm trong thư mục tạm của PyInstaller
    DATA_DIR = os.path.dirname(os.path.abspath(sys.executable))  # dữ liệu nằm cạnh file exe, không mất khi tắt app
else:
    RES_DIR = DATA_DIR = BASE
DB_PATH = os.path.join(DATA_DIR, "data", "qbank.db")
TPL_DIR = os.path.join(DATA_DIR, "templates")
STATIC = os.path.join(RES_DIR, "static")
STATUSES = {"nhap": "Nháp", "da_duyet": "Đã duyệt", "ngung": "Ngưng dùng"}  # chỉ câu "da_duyet" được ra đề
MEDIA_DIR = os.path.join(DATA_DIR, "data", "media")  # hình ảnh và âm thanh đính kèm câu hỏi
MEDIA_EXT = {"png": "image", "jpg": "image", "jpeg": "image", "gif": "image", "webp": "image",
             "mp3": "audio", "wav": "audio", "ogg": "audio", "m4a": "audio"}
MEDIA_MAX = {"image": 5 * 1024 * 1024, "audio": 15 * 1024 * 1024}
MEDIA_NAME = re.compile(r"[a-f0-9]{32}\.(png|jpe?g|gif|webp|mp3|wav|ogg|m4a)")
BLOOM = {1: "Biết", 2: "Hiểu", 3: "Vận dụng", 4: "Phân tích", 5: "Đánh giá"}
LETTERS = "ABCDEFGHIJKL"  # tối đa 12 phương án (nối tiêu đề); TN thường vẫn tối đa 8
QTYPES = ("TN", "DS", "DT", "TL")  # trắc nghiệm, đúng/sai, điền từ, tự luận
# Câu TN có thẻ "TFNG" là dạng Đúng / Sai / Không có thông tin (IELTS Reading), dùng phương án cố định
TFNG_TAG = ",tfng,"
TFNG_OPTIONS = ("正（真）", "誤（偽）", "情報なし（書かれていない）")
# Câu TN có thẻ "noitieude" là nối tiêu đề (matching headings): nhiều đoạn dùng chung một danh sách tiêu đề,
# danh sách in một lần trước nhóm câu và không xáo thứ tự phương án
MATCH_TAG = "noitieude"
MATCH_MAX_OPTIONS = 12


def has_tag(it, name):
    return name in [str(t).lower() for t in (it.get("tags") or [])]


def is_match(it):
    return it.get("qtype") == "TN" and has_tag(it, MATCH_TAG)


# Ghi nhãn sơ đồ: câu DT có thẻ "ghinhan" kèm hình (ảnh có các vị trí đánh số); mỗi vị trí là một câu
LABEL_TAG = "ghinhan"
LABEL_HINT = "Nhìn hình và ghi tên đúng cho vị trí được đánh số."


def is_label(it):
    return it.get("qtype") == "DT" and has_tag(it, LABEL_TAG)
DEFAULT_COURSE = "Chưa phân loại"
PORT = int(os.environ.get("QBANK_PORT", "8765"))
TEXTTS_URL = os.environ.get("TEXTTS_API_URL", "http://127.0.0.1:8000").rstrip("/")  # API của textts (python server.py)
TEXTTS_ENGINES = ("voicevox", "aivis", "edge")


# ---------------------------------------------------------------- database
def q(sql, args=(), fetch=False):
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.execute(sql, args)
        rows = [dict(r) for r in cur.fetchall()] if fetch else None
        conn.commit()
        return rows


def needs_migration():
    """CSDL cũ: questions còn cột chapter dạng chữ, hoặc textbooks còn cột course_id (một giáo trình chỉ thuộc một học phần)."""
    if not os.path.exists(DB_PATH):
        return False
    with closing(sqlite3.connect(DB_PATH)) as conn:
        qcols = [r[1] for r in conn.execute("PRAGMA table_info(questions)")]
        tcols = [r[1] for r in conn.execute("PRAGMA table_info(textbooks)")]
    return "chapter" in qcols or "course_id" in tcols


def backup_before_migration():
    """Chép nguyên file CSDL trước khi nâng cấp cấu trúc, để có thể quay lại nếu có sự cố."""
    backup_dir = os.path.join(os.path.dirname(DB_PATH), "backup")
    os.makedirs(backup_dir, exist_ok=True)
    dest = os.path.join(backup_dir, f"qbank_truoc_nang_cap_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db")
    shutil.copy2(DB_PATH, dest)
    print(f"Đã sao lưu CSDL trước khi nâng cấp: {dest}", flush=True)


def migrate_textbook_links():
    """Giáo trình dùng chung được nhiều học phần: textbooks không còn course_id, học phần nối với giáo trình
    qua course_textbooks. Giáo trình trùng tên (tạo riêng ở nhiều học phần) được gộp vào bản có id nhỏ nhất:
    chương trùng mã được gộp, câu và đoạn văn chuyển sang chương giữ lại. Không xóa câu hỏi nào."""
    if not os.path.exists(DB_PATH):
        return
    with closing(sqlite3.connect(DB_PATH, isolation_level=None)) as conn:
        conn.row_factory = sqlite3.Row
        cols = [r[1] for r in conn.execute("PRAGMA table_info(textbooks)")]
        if "course_id" not in cols:
            return
        conn.execute("PRAGMA foreign_keys = OFF")  # bật lại sau khi xong; cần tắt để đổi bảng cha
        conn.execute("BEGIN IMMEDIATE")
        try:
            conn.execute("""CREATE TABLE IF NOT EXISTS course_textbooks(
                    course_id INTEGER NOT NULL REFERENCES courses(id),
                    textbook_id INTEGER NOT NULL REFERENCES textbooks(id),
                    PRIMARY KEY(course_id, textbook_id))""")
            rows = [dict(r) for r in conn.execute("SELECT id, course_id, name, code FROM textbooks ORDER BY id")]
            keep = {}  # tên giáo trình -> id giữ lại (nhỏ nhất)
            for r in rows:
                keep.setdefault(r["name"], r["id"])
            for r in rows:
                tgt = keep[r["name"]]
                if tgt == r["id"]:
                    continue
                print(f"Gộp giáo trình trùng tên '{r['name']}' (id {r['id']} -> {tgt})", flush=True)
                if r["code"]:
                    conn.execute("UPDATE textbooks SET code=? WHERE id=? AND code=''", (r["code"], tgt))
                for ch in [dict(c) for c in conn.execute("SELECT id, code FROM chapters WHERE textbook_id=?", (r["id"],))]:
                    same = conn.execute("SELECT id FROM chapters WHERE textbook_id=? AND code=?", (tgt, ch["code"])).fetchone()
                    if same:  # chương trùng mã: chuyển câu và đoạn sang chương giữ lại
                        conn.execute("UPDATE questions SET chapter_id=? WHERE chapter_id=?", (same[0], ch["id"]))
                        conn.execute("UPDATE passages SET chapter_id=? WHERE chapter_id=?", (same[0], ch["id"]))
                        conn.execute("DELETE FROM chapters WHERE id=?", (ch["id"],))
                    else:
                        conn.execute("UPDATE chapters SET textbook_id=? WHERE id=?", (tgt, ch["id"]))
            for r in rows:  # mọi học phần từng chứa giáo trình (kể cả bản bị gộp) nay nối với giáo trình giữ lại
                conn.execute("INSERT OR IGNORE INTO course_textbooks(course_id, textbook_id) VALUES(?, ?)",
                             (r["course_id"], keep[r["name"]]))
            for r in rows:
                if keep[r["name"]] != r["id"]:
                    conn.execute("DELETE FROM textbooks WHERE id=?", (r["id"],))
            conn.execute("""CREATE TABLE textbooks_new(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    code TEXT NOT NULL DEFAULT '')""")
            conn.execute("INSERT INTO textbooks_new(id, name, code) SELECT id, name, code FROM textbooks")
            conn.execute("DROP TABLE textbooks")
            conn.execute("ALTER TABLE textbooks_new RENAME TO textbooks")
            bad = conn.execute("PRAGMA foreign_key_check").fetchall()
            if bad:
                raise RuntimeError(f"Sai khóa ngoại sau khi nâng cấp: {len(bad)} dòng")
            conn.execute("COMMIT")
        except Exception:
            conn.execute("ROLLBACK")
            raise
        finally:
            conn.execute("PRAGMA foreign_keys = ON")


def init_storage():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(TPL_DIR, exist_ok=True)
    if needs_migration():
        backup_before_migration()
    migrate_if_needed()
    migrate_textbook_links()
    q("""CREATE TABLE IF NOT EXISTS courses(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            code TEXT NOT NULL DEFAULT '')""")
    q("""CREATE TABLE IF NOT EXISTS textbooks(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            code TEXT NOT NULL DEFAULT '')""")
    q("""CREATE TABLE IF NOT EXISTS course_textbooks(
            course_id INTEGER NOT NULL REFERENCES courses(id),
            textbook_id INTEGER NOT NULL REFERENCES textbooks(id),
            PRIMARY KEY(course_id, textbook_id))""")
    q("""CREATE TABLE IF NOT EXISTS chapters(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            textbook_id INTEGER NOT NULL REFERENCES textbooks(id),
            code TEXT NOT NULL,
            title TEXT NOT NULL DEFAULT '',
            UNIQUE(textbook_id, code))""")
    q("""CREATE TABLE IF NOT EXISTS questions(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT,
            chapter_id INTEGER NOT NULL REFERENCES chapters(id),
            bloom INTEGER NOT NULL,
            qtype TEXT NOT NULL DEFAULT 'TN',
            content TEXT NOT NULL,
            options TEXT NOT NULL DEFAULT '[]',
            answer TEXT NOT NULL DEFAULT '',
            solution TEXT NOT NULL DEFAULT '')""")
    q("""CREATE TABLE IF NOT EXISTS exam_usage(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id INTEGER NOT NULL,
            exam_code TEXT NOT NULL,
            exam_name TEXT NOT NULL DEFAULT '',
            used_at TEXT NOT NULL)""")
    q("""CREATE TABLE IF NOT EXISTS exam_configs(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            config TEXT NOT NULL DEFAULT '{}',
            updated_at TEXT NOT NULL DEFAULT '')""")
    q("""CREATE TABLE IF NOT EXISTS exam_exports(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            name TEXT NOT NULL DEFAULT '',
            codes TEXT NOT NULL DEFAULT '',
            config TEXT NOT NULL DEFAULT '{}',
            data BLOB NOT NULL)""")
    q("""CREATE TABLE IF NOT EXISTS question_results(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            question_id INTEGER NOT NULL,
            exam_code TEXT NOT NULL DEFAULT '',
            correct_pct REAL NOT NULL,
            recorded_at TEXT NOT NULL)""")
    q("""CREATE TABLE IF NOT EXISTS passages(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT NOT NULL UNIQUE,
            chapter_id INTEGER NOT NULL REFERENCES chapters(id),
            content TEXT NOT NULL)""")
    ensure_column("courses", "code", "TEXT NOT NULL DEFAULT ''")
    ensure_column("questions", "passage_id", "INTEGER REFERENCES passages(id)")
    ensure_column("questions", "word_limit", "INTEGER NOT NULL DEFAULT 0")  # điền từ: số từ tối đa (0 = không giới hạn)
    ensure_column("courses", "exam_meta", "TEXT NOT NULL DEFAULT '{}'")
    ensure_column("textbooks", "code", "TEXT NOT NULL DEFAULT ''")
    ensure_column("questions", "code", "TEXT")
    ensure_column("questions", "passage", "TEXT NOT NULL DEFAULT ''")
    ensure_column("questions", "points", "REAL")  # điểm câu hỏi, không bắt buộc
    ensure_column("questions", "clo", "TEXT NOT NULL DEFAULT ''")  # chuẩn đầu ra học phần, ví dụ CLO1
    ensure_column("questions", "rubric", "TEXT NOT NULL DEFAULT ''")  # tự luận: tiêu chí × 5 mức (JSON)
    ensure_column("questions", "tags", "TEXT NOT NULL DEFAULT ''")  # thẻ, lưu dạng ",thẻ1,thẻ2,"
    ensure_column("questions", "media", "TEXT NOT NULL DEFAULT '[]'")  # [{kind, file, name, tts?}]
    ensure_column("passages", "media", "TEXT NOT NULL DEFAULT '[]'")  # âm thanh/hình dùng chung cho mọi câu của đoạn
    # câu có sẵn từ trước được coi là đã duyệt, để đề đang dùng không bị thiếu câu
    ensure_column("questions", "status", "TEXT NOT NULL DEFAULT 'da_duyet'")
    q("""CREATE TABLE IF NOT EXISTS trash(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deleted_at TEXT NOT NULL,
            code TEXT NOT NULL DEFAULT '',
            item TEXT NOT NULL)""")
    migrate_passages()
    q("CREATE UNIQUE INDEX IF NOT EXISTS idx_questions_code ON questions(code)")
    for r in q("SELECT id FROM questions WHERE code IS NULL OR code=''", fetch=True):
        q("UPDATE questions SET code=? WHERE id=?", (make_code(r["id"]), r["id"]))
    # mẫu đi kèm bản đóng gói: chép vào thư mục dữ liệu nếu chưa có, để người dùng sửa được
    bundled = os.path.join(RES_DIR, "templates")
    if os.path.abspath(bundled) != os.path.abspath(TPL_DIR) and os.path.isdir(bundled):
        for f in os.listdir(bundled):
            if f.endswith(".docx") and not os.path.exists(os.path.join(TPL_DIR, f)):
                shutil.copy(os.path.join(bundled, f), os.path.join(TPL_DIR, f))
    if not any(f.endswith(".docx") for f in os.listdir(TPL_DIR)):
        make_sample_templates()
    if not os.path.exists(os.path.join(TPL_DIR, "de_thi_chuan.docx")):
        make_official_templates()


# ---------------------------------------------------------------- textts (tạo file nghe)
# textts chạy trên cùng máy nên không đi qua proxy hệ thống.
_textts = urlreq.build_opener(urlreq.ProxyHandler({}))


def textts_call(path, body=None, timeout=10):
    """Gọi API textts; trả về JSON (hoặc bytes nếu không phải JSON). Lỗi kết nối thành ValueError dễ hiểu."""
    url = path if path.startswith("http") else TEXTTS_URL + path
    data = json.dumps(body, ensure_ascii=False).encode("utf-8") if body is not None else None
    req = urlreq.Request(url, data=data, headers={"Content-Type": "application/json"} if data else {})
    try:
        with _textts.open(req, timeout=timeout) as r:
            raw = r.read()
            return json.loads(raw) if "json" in r.headers.get("Content-Type", "") else raw
    except HTTPError as e:
        try:
            detail = json.loads(e.read()).get("detail")
        except Exception:
            detail = None
        raise ValueError(f"textts báo lỗi: {detail or e.reason}")
    except (URLError, TimeoutError, OSError) as e:
        raise ValueError(f"Không kết nối được textts tại {TEXTTS_URL}. Hãy bật textts (python server.py) rồi thử lại. ({e})")


def to_textts_script(text):
    """Furigana của qbank {漢字|かな} thành cú pháp textts 漢字{かな}; dòng chưa có vai thì gán vai [1]."""
    text = re.sub(r"\{([^{}|]+)\|([^{}]+)\}", r"\1{\2}", text or "")
    lines = []
    for ln in text.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        lines.append(ln if re.match(r"^(\[[^\]]+\]|\d+[\s.:：])", ln) else f"[1] {ln}")
    return "\n".join(lines)


def textts_audio(script, engine="voicevox", level=None):
    """Nhờ textts đọc kịch bản, lưu mp3 vào thư mục media; trả về {kind, file, name} để gắn vào câu."""
    raw_script = script
    script = to_textts_script(script)
    if not script:
        raise ValueError("Kịch bản trống")
    if engine not in TEXTTS_ENGINES:
        raise ValueError("Engine không hợp lệ")
    stem = f"qbank_{datetime.now():%Y%m%d_%H%M%S}_{os.urandom(3).hex()}"
    res = textts_call("/api/v1/dialogue", {
        "title": stem, "filename": stem, "script": script, "engine": engine,
        "options": {"format": "mp3", "slow": False, "srt": False, "worksheet": False,
                    "level": level if level in ("N5", "N4", "N3", "N2", "N1") else None},
    }, timeout=600)
    url = (res.get("files") or {}).get("audio_url") if isinstance(res, dict) else None
    if not url:
        raise ValueError("textts không trả về file âm thanh")
    audio = textts_call(url, timeout=120)
    if not isinstance(audio, bytes) or not audio:
        raise ValueError("Không tải được file âm thanh từ textts")
    if len(audio) > MEDIA_MAX["audio"]:
        raise ValueError("File âm thanh quá lớn (tối đa 15 MB)")
    os.makedirs(MEDIA_DIR, exist_ok=True)
    fname = f"{os.urandom(16).hex()}.mp3"
    with open(os.path.join(MEDIA_DIR, fname), "wb") as f:
        f.write(audio)
    return {"kind": "audio", "file": fname, "name": f"{stem}.mp3", "duration_s": res.get("duration_s"),
            "tts": norm_tts({"script": raw_script, "engine": engine, "level": level})}


def _audio_of(media):
    return [m for m in media or [] if m["kind"] == "audio" and os.path.isfile(os.path.join(MEDIA_DIR, m["file"]))]


def listening_segments(flat, repeat=1, answer_s=10, gap_ms=1500):
    """Kịch bản ghép đề nghe của một mã đề, theo đúng thứ tự câu trong đề:
    đọc số câu (n番), phát file của đoạn văn và của câu (lặp `repeat` lần), chừa thời gian làm bài.
    Câu không có file nghe (và đoạn của nó cũng không có) thì bỏ qua. Trả về [] nếu đề không có file nghe nào."""
    def speech(text):
        return {"type": "speech", "text": text}

    def pause(ms):
        return {"type": "pause", "ms": int(ms)}

    def play(files):
        out = []
        for r in range(repeat):
            for m in files:
                with open(os.path.join(MEDIA_DIR, m["file"]), "rb") as f:
                    out.append({"type": "audio", "format": m["file"].rsplit(".", 1)[1],
                                "audio_b64": base64.b64encode(f.read()).decode()})
                out.append(pause(gap_ms))
        return out

    segs, i = [], 0
    while i < len(flat):
        n, _, it = flat[i]
        p_audio = _audio_of(it.get("passage_media")) if it.get("passage") else []
        if p_audio:  # nhóm câu liền nhau cùng đoạn văn: phát đoạn một lần rồi đến từng câu
            group = [flat[i]]
            while i + len(group) < len(flat) and flat[i + len(group)][2].get("passage") == it["passage"]:
                group.append(flat[i + len(group)])
            last = group[-1][0]
            segs += [speech(f"{n}番{{ばん}}から{last}番{{ばん}}。" if last > n else f"{n}番{{ばん}}。"), pause(gap_ms)]
            segs += play(p_audio)
            for k, _, g in group:
                q_audio = _audio_of(g.get("media"))
                if q_audio:
                    segs += [speech(f"{k}番{{ばん}}。"), pause(800)] + play(q_audio)
            segs.append(pause(answer_s * 1000 * len(group)))
            i += len(group)
            continue
        q_audio = _audio_of(it.get("media"))
        if q_audio:
            segs += [speech(f"{n}番{{ばん}}。"), pause(800)] + play(q_audio) + [pause(answer_s * 1000)]
        i += 1
    return segs


def listening_only(exam, flags):
    """Các câu của phần được đánh dấu "Phần nghe", giữ nguyên số thứ tự toàn đề (n) như trong flatten(exam)."""
    keep = {id(it) for (_, items), flag in zip(exam, flags) if flag for it in items}
    return [x for x in flatten(exam) if id(x[2]) in keep]


def listening_track(code, flat, opts):
    """Nhờ textts ghép đề nghe của một mã đề; trả về bytes mp3, hoặc None nếu đề không có file nghe."""
    repeat = 2 if str(opts.get("repeat", 1)) == "2" else 1
    try:
        answer_s = max(0, min(120, int(opts.get("answer_s", 10))))
    except (TypeError, ValueError):
        raise ValueError("Thời gian làm mỗi câu nghe phải là số giây")
    segs = listening_segments(flat, repeat, answer_s)
    if not segs:
        return None
    engine = opts.get("engine") if opts.get("engine") in TEXTTS_ENGINES else "voicevox"
    stem = f"qbank_de_{re.sub(r'[^\w-]', '_', code)}_{os.urandom(3).hex()}"
    res = textts_call("/api/v1/assemble", {"filename": stem, "engine": engine, "segments": segs, "format": "mp3"},
                      timeout=1800)
    url = (res.get("files") or {}).get("audio_url") if isinstance(res, dict) else None
    if not url:
        raise ValueError("textts không trả về file đề nghe")
    data = textts_call(url, timeout=300)
    if not isinstance(data, bytes) or not data:
        raise ValueError("Không tải được file đề nghe từ textts")
    return data


def migrate_passages():
    """Đoạn văn cũ (chữ giống hệt trong cùng chương) thành đối tượng riêng có mã. Chạy lại an toàn."""
    for r in q("SELECT DISTINCT chapter_id, passage FROM questions WHERE passage <> '' AND passage_id IS NULL", fetch=True):
        pid = resolve_passage(r["chapter_id"], r["passage"])
        q("UPDATE questions SET passage_id=? WHERE chapter_id=? AND passage=? AND passage_id IS NULL",
          (pid, r["chapter_id"], r["passage"]))


def passage_code(chapter_id):
    """Mã đoạn văn dạng HọcPhần-GiáoTrình-Chương-Pnnn, ví dụ N4-MN1-B01-P001."""
    r = q("""SELECT ch.code AS chcode, t.id AS tid, t.code AS tcode
             FROM chapters ch JOIN textbooks t ON t.id=ch.textbook_id
             WHERE ch.id=?""", (chapter_id,), True)[0]
    co = textbook_course(r["tid"])
    hp = co["code"] or f"HP{co['id']}"
    gt = r["tcode"] or f"GT{r['tid']}"
    prefix = re.sub(r"\s+", "_", f"{hp}-{gt}-{r['chcode']}")
    n = q("SELECT COUNT(*) AS n FROM passages WHERE chapter_id=?", (chapter_id,), True)[0]["n"] + 1
    while True:
        code = f"{prefix}-P{n:03d}"
        if not q("SELECT 1 FROM passages WHERE code=?", (code,), True):
            return code
        n += 1


def default_exam_name(meta, codes):
    """Tên mặc định của bản xuất khi người dùng không đặt tên: học phần, mã đề, thời gian."""
    base = str(meta.get("ten_hp", "")).strip() or "Đề thi"
    mid = f"mã {codes[0]}" if len(codes) == 1 else f"mã {codes[0]}–{codes[-1]}"
    return " · ".join(x for x in [base, mid, str(meta.get("thoi_gian", "")).strip()] if x)


def export_filename(ma_hp, codes, created_at):
    """Tên file tải về: <mã HP>_<mã đề đầu-cuối>_<yyyymmdd>.zip, ví dụ N4_M001-M012_20261101.zip."""
    hp = (ma_hp or "").strip()
    rng = codes[0] if len(codes) <= 1 else f"{codes[0]}-{codes[-1]}"
    stem = f"De_{hp}_{rng}" if hp else f"De_{rng}"
    name = f"{stem}_{created_at[:10].replace('-', '')}.zip" if codes else "De.zip"
    return re.sub(r"[^\w.-]+", "_", name)


def attachment_header(name):
    """Content-Disposition an toàn với tên Unicode: bản ASCII cho trình duyệt cũ, bản UTF-8 (RFC 5987)."""
    ascii_name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode() or "de_thi.zip"
    return f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(name)}"


def resolve_passage(chapter_id, text):
    """Tìm đoạn văn cùng nội dung trong chương; chưa có thì tạo mới. Không có chữ thì trả None."""
    text = (text or "").strip()
    if not text:
        return None
    hit = q("SELECT id FROM passages WHERE chapter_id=? AND content=?", (int(chapter_id), text), True)
    if hit:
        return hit[0]["id"]
    code = passage_code(int(chapter_id))
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cur = conn.execute("INSERT INTO passages(code, chapter_id, content) VALUES(?,?,?)", (code, int(chapter_id), text))
        conn.commit()
        return cur.lastrowid


def ensure_column(table, col, decl):
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})")]
        if col not in cols:
            conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {decl}")
            conn.commit()


def migrate_if_needed():
    """Chuyển CSDL cũ (questions có cột chapter dạng chữ) sang cấu trúc học phần/giáo trình/chương."""
    if not os.path.exists(DB_PATH):
        return
    with closing(sqlite3.connect(DB_PATH)) as conn:
        cols = [r[1] for r in conn.execute("PRAGMA table_info(questions)")]
        if "chapter" not in cols:
            return
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("ALTER TABLE questions RENAME TO questions_old")
        conn.commit()
    q("""CREATE TABLE courses(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE)""")
    q("""CREATE TABLE textbooks(id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_id INTEGER NOT NULL REFERENCES courses(id), name TEXT NOT NULL, UNIQUE(course_id, name))""")
    q("""CREATE TABLE chapters(id INTEGER PRIMARY KEY AUTOINCREMENT,
            textbook_id INTEGER NOT NULL REFERENCES textbooks(id), code TEXT NOT NULL,
            title TEXT NOT NULL DEFAULT '', UNIQUE(textbook_id, code))""")
    q("""CREATE TABLE questions(id INTEGER PRIMARY KEY AUTOINCREMENT,
            chapter_id INTEGER NOT NULL REFERENCES chapters(id), bloom INTEGER NOT NULL,
            qtype TEXT NOT NULL DEFAULT 'TN', content TEXT NOT NULL, options TEXT NOT NULL DEFAULT '[]',
            answer TEXT NOT NULL DEFAULT '', solution TEXT NOT NULL DEFAULT '')""")
    migrate_textbook_links()  # đưa textbooks về cấu trúc nhiều-nhiều trước khi ghi dữ liệu mới
    # Dữ liệu cũ không có học phần/giáo trình: đưa vào một nhóm "Chưa phân loại" để người dùng gán lại.
    # Mỗi mã chương cũ thành một chương trong giáo trình "Chưa phân loại"
    with closing(sqlite3.connect(DB_PATH)) as conn:
        codes = [r[0] for r in conn.execute("SELECT DISTINCT chapter FROM questions_old")]
    for code in codes:
        ensure_chapter(DEFAULT_COURSE, DEFAULT_COURSE, code)
    q("""INSERT INTO questions(id, chapter_id, bloom, qtype, content, options, answer, solution)
         SELECT o.id, (SELECT c.id FROM chapters c JOIN textbooks t ON t.id=c.textbook_id
                       JOIN course_textbooks ct ON ct.textbook_id=t.id JOIN courses co ON co.id=ct.course_id
                       WHERE co.name=? AND t.name=? AND c.code=o.chapter),
                o.bloom, o.qtype, o.content, o.options, o.answer, o.solution
         FROM questions_old o""", (DEFAULT_COURSE, DEFAULT_COURSE))
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("DROP TABLE questions_old")
        conn.commit()


# ---------------------------------------------------------------- hierarchy helpers
CHAPTER_LABEL = "t.name || ' / ' || ch.code || CASE WHEN ch.title <> '' THEN ' - ' || ch.title ELSE '' END"
# Học phần đại diện của giáo trình: học phần có id nhỏ nhất trong course_textbooks (giáo trình có thể dùng chung nhiều học phần).
PRIMARY_COURSE = "(SELECT {col} FROM course_textbooks ct JOIN courses co ON co.id = ct.course_id WHERE ct.textbook_id = t.id ORDER BY co.id LIMIT 1)"
QUESTION_SQL = f"""SELECT qs.*, ch.code AS chapter_code, ch.title AS chapter_title,
    (SELECT COUNT(DISTINCT u.used_at) FROM exam_usage u WHERE u.question_id = qs.id) AS used_count,
    ch.id AS chapter_id, t.id AS textbook_id, t.name AS textbook, t.code AS textbook_code,
    COALESCE({PRIMARY_COURSE.format(col="co.id")}, 0) AS course_id,
    COALESCE({PRIMARY_COURSE.format(col="co.name")}, '') AS course,
    (SELECT GROUP_CONCAT(name, ', ') FROM (SELECT co.name FROM course_textbooks ct JOIN courses co ON co.id = ct.course_id
        WHERE ct.textbook_id = t.id ORDER BY co.id)) AS courses_all,
    ({CHAPTER_LABEL}) AS chapter_label
    FROM questions qs
    JOIN chapters ch ON ch.id = qs.chapter_id
    JOIN textbooks t ON t.id = ch.textbook_id"""


def textbook_course(textbook_id):
    """Học phần đại diện của giáo trình (dùng cho tiền tố mã câu/đoạn): {id, code} hoặc {id: 0, code: ''}."""
    r = q("""SELECT co.id, co.code FROM course_textbooks ct JOIN courses co ON co.id = ct.course_id
             WHERE ct.textbook_id=? ORDER BY co.id LIMIT 1""", (int(textbook_id),), True)
    return r[0] if r else {"id": 0, "code": ""}


def ensure_chapter(courses, textbook, code, title=""):
    """Tìm hoặc tạo học phần/giáo trình/chương theo tên; giáo trình được nối với mọi học phần trong `courses`.
    Giáo trình trùng tên dùng chung, không tạo bản mới. Trả về id chương."""
    names = [courses] if isinstance(courses, str) else list(courses)
    q("INSERT OR IGNORE INTO textbooks(name) VALUES(?)", (textbook,))
    tid = q("SELECT id FROM textbooks WHERE name=?", (textbook,), True)[0]["id"]
    for course in names:
        q("INSERT OR IGNORE INTO courses(name) VALUES(?)", (course,))
        cid = q("SELECT id FROM courses WHERE name=?", (course,), True)[0]["id"]
        q("INSERT OR IGNORE INTO course_textbooks(course_id, textbook_id) VALUES(?, ?)", (cid, tid))
    q("INSERT OR IGNORE INTO chapters(textbook_id, code, title) VALUES(?, ?, ?)", (tid, code, title))
    return q("SELECT id FROM chapters WHERE textbook_id=? AND code=?", (tid, code), True)[0]["id"]


def chapter_label(chapter_id):
    rows = q(f"""SELECT ({CHAPTER_LABEL}) AS label FROM chapters ch
                 JOIN textbooks t ON t.id=ch.textbook_id
                 WHERE ch.id=?""", (chapter_id,), True)
    return rows[0]["label"] if rows else f"chương #{chapter_id}"


EXAM_META_KEYS = ("ten_hp", "ma_hp", "so_tc", "khoa_hoc", "nam_hoc", "hoc_ky", "thoi_gian")


def get_tree():
    courses = q("SELECT id, name, code, exam_meta FROM courses ORDER BY name", fetch=True)
    for c in courses:
        try:
            c["exam_meta"] = json.loads(c["exam_meta"] or "{}")
        except ValueError:
            c["exam_meta"] = {}
    textbooks = q("SELECT id, name, code FROM textbooks ORDER BY name", fetch=True)
    links = {}  # giáo trình -> danh sách id học phần (theo thứ tự id)
    for r in q("SELECT course_id, textbook_id FROM course_textbooks ORDER BY course_id", fetch=True):
        links.setdefault(r["textbook_id"], []).append(r["course_id"])
    for t in textbooks:
        t["course_ids"] = links.get(t["id"], [])
        t["course_id"] = t["course_ids"][0] if t["course_ids"] else 0  # học phần đại diện (id nhỏ nhất)
    chapters = q("""SELECT ch.id, ch.textbook_id, ch.code, ch.title,
                    (SELECT COUNT(*) FROM questions qs WHERE qs.chapter_id = ch.id) AS count
                    FROM chapters ch ORDER BY ch.code""", fetch=True)
    return {"courses": courses, "textbooks": textbooks, "chapters": chapters}


# ---------------------------------------------------------------- validation
def normalize_ds(ans, n):
    """Đáp án đúng/sai: 'Đ S Đ' / 'DSD' / 'Đ,S,Đ' -> 'DSD' (D = Đúng, S = Sai), đủ n ý."""
    letters = [("D" if c in "ĐD" else "S") for c in ans.upper() if c in "ĐDS"]
    if len(letters) != n:
        raise ValueError(f"Đáp án đúng/sai cần đủ {n} ý (Đ hoặc S)")
    return "".join(letters)


def norm_tags(value):
    """Thẻ từ chuỗi "a, b" hoặc danh sách -> chuỗi lưu ",a,b,". Bỏ trùng (không phân biệt hoa thường), tối đa 15 thẻ."""
    if isinstance(value, str):
        parts = re.split(r"[,;，]", value)
    else:
        parts = [str(x) for x in (value or [])]
    out, seen = [], set()
    for t in parts:
        t = re.sub(r"\s+", " ", t).strip()[:30]
        if not t or t.lower() in seen:
            continue
        if len(out) >= 15:
            raise ValueError("Mỗi câu tối đa 15 thẻ")
        seen.add(t.lower())
        out.append(t)
    return "," + ",".join(out) + "," if out else ""


def norm_media(value):
    """Danh sách file đính kèm câu: chỉ nhận file đã có trong thư mục media."""
    out = []
    for m in value or []:
        f = str((m or {}).get("file", ""))
        if not MEDIA_NAME.fullmatch(f) or not os.path.isfile(os.path.join(MEDIA_DIR, f)):
            raise ValueError("File đính kèm không còn tồn tại, hãy tải lên lại")
        item = {"kind": MEDIA_EXT[f.rsplit(".", 1)[1]], "file": f, "name": str(m.get("name", f))[:120]}
        tts = norm_tts(m.get("tts"))
        if tts and item["kind"] == "audio":
            item["tts"] = tts
        out.append(item)
    if len(out) > 4:
        raise ValueError("Mỗi câu tối đa 4 file đính kèm")
    return out


def norm_tts(value):
    """Kịch bản đã dùng để tạo file nghe bằng textts (lưu kèm file để sửa và tạo lại)."""
    if not isinstance(value, dict) or not str(value.get("script", "")).strip():
        return None
    engine = value.get("engine") if value.get("engine") in TEXTTS_ENGINES else "voicevox"
    level = value.get("level") if value.get("level") in ("N5", "N4", "N3", "N2", "N1") else ""
    return {"script": str(value["script"]).strip()[:5000], "engine": engine, "level": level}


def tags_list(stored):
    return [t for t in (stored or "").split(",") if t]


# Câu đối đáp (từ hai người trở lên, ví dụ "A: ... / B: ..."): mặc định CLO3 (nghe hiểu, phát âm)
SPEAKER_LINE = re.compile(r"(?m)^\s*([^\n:：]{1,12})\s*[:：]")


def is_dialogue(*texts):
    labels = set()
    for t in texts:
        for m in SPEAKER_LINE.finditer(t or ""):
            lab = m.group(1).strip().lower()
            if lab and not re.match(r"^(câu|bài|đáp|phần|\d)", lab):
                labels.add(lab)
    return len(labels) >= 2


def clean_question(data):
    try:
        chapter_id = int(data.get("chapter_id", 0))
    except (TypeError, ValueError):
        raise ValueError("Thiếu chương của câu hỏi")
    if not q("SELECT id FROM chapters WHERE id=?", (chapter_id,), True):
        raise ValueError("Chương không tồn tại")
    content = str(data.get("content", "")).strip()
    qtype = str(data.get("qtype", "TN")).upper()
    try:
        bloom = int(data.get("bloom", 1))
    except (TypeError, ValueError):
        raise ValueError("Mức Bloom phải là số 1-5")
    options = [str(o).strip() for o in data.get("options", []) if str(o).strip()]
    answer = str(data.get("answer", "")).strip()
    if qtype == "TN":
        answer = answer.upper()
    solution = str(data.get("solution", "")).strip()
    if not content:
        raise ValueError("Thiếu nội dung câu hỏi")
    if not 1 <= bloom <= 5:
        raise ValueError("Mức Bloom phải từ 1 đến 5")
    if qtype not in QTYPES:
        raise ValueError("Loại câu phải là TN, DS, DT hoặc TL")
    tags_norm = norm_tags(data.get("tags", ""))
    try:
        word_limit = int(data.get("word_limit") or 0)
    except (TypeError, ValueError):
        raise ValueError("Giới hạn số từ phải là số nguyên")
    if qtype != "DT":
        word_limit = 0  # chỉ điền từ mới có giới hạn số từ
    if qtype == "TN" and TFNG_TAG in tags_norm.lower():
        # Đúng / sai / không có thông tin (kiểu IELTS): 3 phương án cố định, đáp án là A, B hoặc C
        options = list(TFNG_OPTIONS)
        if answer not in LETTERS[:3]:
            raise ValueError("Đúng/Sai/Không có thông tin: đáp án là A (正), B (誤) hoặc C (情報なし)")
    if qtype == "TN":
        if len(options) < 2:
            raise ValueError("Trắc nghiệm cần ít nhất 2 phương án")
        limit = MATCH_MAX_OPTIONS if MATCH_TAG in tags_norm.lower() else 8
        if len(options) > limit:
            raise ValueError(f"Trắc nghiệm tối đa {limit} phương án")
        if answer not in LETTERS[:len(options)]:
            raise ValueError("Đáp án không hợp lệ")
    elif qtype == "DS":
        if not 1 <= len(options) <= 8:
            raise ValueError("Đúng/sai cần từ 1 đến 8 ý")
        answer = normalize_ds(answer, len(options))
    elif qtype == "DT":
        if not answer:
            raise ValueError("Thiếu đáp án điền từ (nhiều đáp án chấp nhận cách nhau bằng |)")
        if word_limit and not 1 <= word_limit <= 20:
            raise ValueError("Giới hạn số từ điền từ phải từ 1 đến 20 (0 = không giới hạn)")
        options = []
    else:  # TL: giữ đáp án/gợi ý chấm nếu có
        options = []
    passage = str(data.get("passage", "")).strip()
    clo = str(data.get("clo", "")).strip()[:20]
    if not clo and is_dialogue(content, passage):  # câu đối đáp mặc định là CLO3
        clo = "CLO3"
    tags = norm_tags(data.get("tags", ""))
    media = norm_media(data.get("media") or [])
    status = str(data.get("status") or "nhap")
    if status not in STATUSES:
        raise ValueError("Trạng thái phải là Nháp, Đã duyệt hoặc Ngưng dùng")
    raw = data.get("points")
    points = None
    if raw not in (None, ""):
        try:
            points = float(str(raw).replace(",", "."))
        except ValueError:
            raise ValueError("Điểm phải là số, ví dụ 0,5 hoặc 2")
        if not 0 < points <= 10:
            raise ValueError("Điểm của câu phải từ 0 đến 10")
    rubric = []  # tự luận: mỗi tiêu chí có điểm và mô tả cho mức 0-4
    if qtype == "TL":
        for crit in data.get("rubric") or []:
            text = str(crit.get("text", "")).strip()
            if not text:
                continue
            raw = crit.get("points")
            pts = None
            if raw not in (None, ""):
                try:
                    pts = float(str(raw).replace(",", "."))
                except ValueError:
                    raise ValueError("Điểm tiêu chí phải là số")
            levels = [str(x).strip() for x in (crit.get("levels") or [])][:5]
            levels += [""] * (5 - len(levels))
            rubric.append({"text": text, "points": pts, "levels": levels})
    return dict(chapter_id=chapter_id, bloom=bloom, qtype=qtype, content=content,
                options=json.dumps(options, ensure_ascii=False), answer=answer, solution=solution,
                passage=passage, points=points, clo=clo, tags=tags, media=json.dumps(media, ensure_ascii=False),
                status=status, rubric=json.dumps(rubric, ensure_ascii=False), word_limit=word_limit)


INSERT_Q = """INSERT INTO questions(code,chapter_id,bloom,qtype,content,options,answer,solution,passage,passage_id,points,clo,tags,media,status,rubric,word_limit)
              VALUES(:code,:chapter_id,:bloom,:qtype,:content,:options,:answer,:solution,:passage,:passage_id,:points,:clo,:tags,:media,:status,:rubric,:word_limit)"""


def row_to_api(r):
    r = dict(r)
    r["options"] = json.loads(r["options"])
    if "rubric" in r:
        r["rubric"] = json.loads(r["rubric"] or "[]")
    if "tags" in r:
        r["tags"] = tags_list(r["tags"])
    if "media" in r:
        r["media"] = json.loads(r["media"] or "[]")
    r["passage_media"] = passage_media(r.get("passage_id"))
    return r


def passage_media(pid):
    """File đính kèm của đoạn văn (dùng chung cho các câu thuộc đoạn)."""
    if not pid:
        return []
    row = q("SELECT media FROM passages WHERE id=?", (int(pid),), True)
    return json.loads(row[0]["media"] or "[]") if row else []


def set_passage_media(pid, media):
    if pid:
        q("UPDATE passages SET media=? WHERE id=?", (json.dumps(norm_media(media), ensure_ascii=False), int(pid)))


def save_passage_media(question_id, body):
    """Ghi file đính kèm của đoạn văn khi lưu câu (chỉ khi form gửi kèm passage_media)."""
    if isinstance(body.get("passage_media"), list):
        row = q("SELECT passage_id FROM questions WHERE id=?", (int(question_id),), True)
        if row and row[0]["passage_id"]:
            set_passage_media(row[0]["passage_id"], body["passage_media"])


def make_code(question_id_or_none, chapter_id=None):
    """Mã câu hỏi ổn định, dạng HọcPhần-GiáoTrình-Chương-Số (ví dụ N5-MN1-B03-001).
    Mã được gán khi tạo câu và không đổi sau đó. Số tăng theo từng chương, bỏ qua mã đã có."""
    if chapter_id is None:  # gọi khi nâng cấp CSDL: lấy chương của câu
        chapter_id = q("SELECT chapter_id FROM questions WHERE id=?", (question_id_or_none,), True)[0]["chapter_id"]
    r = q("""SELECT ch.code AS chcode, t.id AS tid, t.code AS tcode
             FROM chapters ch JOIN textbooks t ON t.id=ch.textbook_id
             WHERE ch.id=?""", (chapter_id,), True)[0]
    co = textbook_course(r["tid"])
    hp = co["code"] or f"HP{co['id']}"
    gt = r["tcode"] or f"GT{r['tid']}"
    prefix = re.sub(r"\s+", "_", f"{hp}-{gt}-{r['chcode']}")
    # số kế tiếp sau số lớn nhất đã dùng (cả câu trong thùng rác), để mã đã cấp không bị cấp lại
    taken = [r["code"] for r in q("SELECT code FROM questions WHERE code LIKE ?", (prefix + "-%",), True)]
    taken += [r["code"] for r in q("SELECT code FROM trash WHERE code LIKE ?", (prefix + "-%",), True)]
    nums = [int(c[len(prefix) + 1:]) for c in taken if c[len(prefix) + 1:].isdigit()]
    n = max(nums, default=0) + 1
    while True:
        code = f"{prefix}-{n:03d}"
        if not q("SELECT 1 FROM questions WHERE code=?", (code,), True):
            return code
        n += 1


def merge_questions(keep, removes):
    """Xóa các câu trùng nội dung với câu giữ lại; chuyển lịch sử dùng sang câu giữ."""
    if not removes:
        raise ValueError("Chưa chọn câu nào để xóa")
    rows = {r["id"]: r for r in q("SELECT id, content, options FROM questions WHERE id IN (%s)"
                                  % ",".join("?" * (len(removes) + 1)), [keep] + removes, fetch=True)}
    if keep not in rows:
        raise ValueError("Câu giữ lại không tồn tại")
    if len(rows) != len(removes) + 1:
        raise ValueError("Có câu không tồn tại trong danh sách xóa")
    key = dup_key(rows[keep]["content"], json.loads(rows[keep]["options"]))
    for rid in removes:
        if dup_key(rows[rid]["content"], json.loads(rows[rid]["options"])) != key:
            raise ValueError(f"Câu {code_of(rid)} không trùng nội dung với câu {code_of(keep)}, không thể gộp")
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        ph = ",".join("?" * len(removes))
        conn.execute(f"UPDATE exam_usage SET question_id=? WHERE question_id IN ({ph})", [keep] + removes)
        conn.execute(f"DELETE FROM questions WHERE id IN ({ph})", removes)
        conn.commit()
    return {"removed": len(removes), "kept": keep}


def auto_merge_same_chapter():
    """Gộp mọi nhóm trùng nằm trong cùng một chương. Trùng khác chương không bị đụng tới."""
    removed = 0
    for v in dup_index().values():
        by_chapter = {}
        for qid, chapter_id, _ in sorted(v):
            by_chapter.setdefault(chapter_id, []).append(qid)
        for ids in by_chapter.values():
            if len(ids) > 1:
                merge_questions(ids[0], ids[1:])
                removed += len(ids) - 1
    return {"removed": removed}


def insert_q(payload, keep_code=None):
    """Thêm câu hỏi (có gán mã), trả về id mới. keep_code: giữ mã cũ nếu chưa có câu nào dùng mã đó (hoàn tác xóa)."""
    code = None
    if keep_code and not q("SELECT id FROM questions WHERE code=?", (keep_code,), fetch=True):
        code = keep_code
    payload = dict(payload, code=code or make_code(None, payload["chapter_id"]),
                   passage_id=resolve_passage(payload["chapter_id"], payload.get("passage")))
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.execute(INSERT_Q, payload)
        conn.commit()
        return cur.lastrowid


# ---------------------------------------------------------------- trùng lặp
def code_of(qid):
    """Mã câu để hiển thị cho người dùng, thay cho số id nội bộ. Giá trị không phải id (nhãn dòng khi xem trước) được giữ nguyên."""
    if not isinstance(qid, int):
        return str(qid)
    rows = q("SELECT code FROM questions WHERE id=?", (qid,), True)
    return rows[0]["code"] if rows else "(đã xóa)"


DUP_IGNORE = re.compile(r"[\s。、，．,.!?！？「」『』“”\"'()（）]+")


def dup_key(content, options):
    """Khóa so trùng: bỏ hoa/thường, khoảng trắng và dấu câu thông thường.
    Giữ ký hiệu toán học và chữ Nhật, chữ Hán (1+1 và 1-1 không còn bị coi là trùng)."""
    return DUP_IGNORE.sub("", (content + "".join(options)).lower())


def dup_index():
    """key -> [(id, chapter_id, nội dung)] cho toàn bộ ngân hàng."""
    index = {}
    for r in q("SELECT id, chapter_id, content, options FROM questions", fetch=True):
        key = dup_key(r["content"], json.loads(r["options"]))
        index.setdefault(key, []).append((r["id"], r["chapter_id"], r["content"]))
    return index


def check_duplicate(index, payload):
    """Trả về (trùng cùng chương, trùng chương khác) dựa trên chỉ mục."""
    key = dup_key(payload["content"], json.loads(payload["options"]))
    hits = index.get(key, [])
    same = [h for h in hits if h[1] == payload["chapter_id"]]
    other = [h for h in hits if h[1] != payload["chapter_id"]]
    return same, other, key


# ---------------------------------------------------------------- lịch sử dùng câu
def usage_map():
    """question_id -> thời điểm dùng gần nhất (ISO), để ưu tiên câu chưa dùng."""
    return {r["question_id"]: r["last"] for r in q(
        "SELECT question_id, MAX(used_at) AS last FROM exam_usage GROUP BY question_id", fetch=True)}


def usage_counts():
    """question_id -> số lần đã ra đề (mỗi lần xuất đếm một lần)."""
    return {r["question_id"]: r["n"] for r in q(
        "SELECT question_id, COUNT(DISTINCT used_at) AS n FROM exam_usage GROUP BY question_id", fetch=True)}


def record_usage(rows):
    """rows: [(question_id, exam_code, exam_name, used_at)]"""
    if not rows:
        return
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.executemany("INSERT INTO exam_usage(question_id, exam_code, exam_name, used_at) VALUES(?,?,?,?)", rows)
        conn.commit()


# ---------------------------------------------------------------- import .docx
OPT_RE = re.compile(r"^([A-Ha-h])[\.\)]\s*(.*)$")



IMPORT_COLUMNS = {
    "ma": ["mã chương", "ma chuong", "mã", "ma"],
    "noi_dung": ["nội dung", "noi dung"],
    "dap_an": ["đáp án", "dap an"],
    "loai": ["loại", "loai"],
    "muc": ["mức", "muc", "mức tn", "muc tn"],
    "doan_van": ["đoạn văn", "doan van"],
    "loi_giai": ["lời giải", "loi giai"],
    "clo": ["clo", "chuẩn đầu ra", "chuan dau ra"],
    "tags": ["thẻ", "the", "tags"],
    "so_tu": ["số từ", "so tu", "số từ tối đa", "so tu toi da", "giới hạn từ", "gioi han tu"],
    "status": ["trạng thái", "trang thai", "status"],
}
IMPORT_REQUIRED = ("ma", "noi_dung", "dap_an", "loai", "muc")


def _norm_header(s):
    return re.sub(r"\s+", " ", s.strip().lower())


def map_import_header(cells):
    """Tìm vị trí cột theo tên tiêu đề (không phụ thuộc thứ tự). Trả về {khóa: chỉ số cột} hoặc None."""
    m = {}
    for idx, c in enumerate(cells):
        n = _norm_header(c)
        for key, names in IMPORT_COLUMNS.items():
            if n in names and key not in m:
                m[key] = idx
    missing = [k for k in IMPORT_REQUIRED if k not in m]
    return (m, missing)


def _xlsx_text(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else ("%g" % v)
    return str(v)


def read_xlsx(raw: bytes):
    """Sheet đầu tiên của file .xlsx thành danh sách dòng (chuỗi). Ô trống thành chuỗi rỗng."""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except Exception:
        raise ValueError("Không đọc được file .xlsx (hãy lưu lại từ Excel rồi thử lại)")
    ws = wb.worksheets[0]
    rows = []
    for row in ws.iter_rows(values_only=True):
        cells = [_xlsx_text(v) for v in row]
        while cells and cells[-1] == "":
            cells.pop()
        rows.append(cells)
    wb.close()
    return rows


def read_table_text(raw: bytes):
    """CSV, dữ liệu dán từ Excel (TSV) hoặc .xlsx. Tự nhận dấu phân cách và bảng mã (UTF-8, UTF-16, Windows-1258)."""
    if raw[:2] == b"PK":  # .xlsx là file zip
        return read_xlsx(raw)
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        text = raw.decode("utf-16")
    else:
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1258", errors="replace")
    first = next((ln for ln in text.splitlines() if ln.strip()), "")
    if "\t" in first:
        delim = "\t"
    elif first.count(";") > first.count(","):
        delim = ";"
    else:
        delim = ","
    return [r for r in csv.reader(io.StringIO(text), delimiter=delim)]


def parse_import(data: bytes, textbook_id, dry_run=False):
    """Nhập câu hỏi từ bảng Word (xem /api/import-template để có mẫu)."""
    try:
        doc = Document(io.BytesIO(data))
    except Exception:
        raise ValueError("File không phải .docx hợp lệ (hãy lưu lại bằng Word dưới định dạng .docx)")
    tables = [[[c.text for c in row.cells] for row in t.rows] for t in doc.tables]
    return parse_rows(tables, textbook_id, dry_run, docx=True)


CHAPTER_CODE_HEADERS = ("mã", "ma", "mã chương", "ma chuong", "code")
CHAPTER_TITLE_HEADERS = ("tên", "ten", "tên chương", "ten chuong", "tiêu đề", "tieu de", "title")


def plan_chapters(rows, textbook_id):
    """Đọc danh sách chương từ bảng: cột Mã và Tên chương, dòng đầu là tiêu đề nếu có.
    Trả về từng dòng kèm trạng thái: thêm / bỏ qua (mã đã có) / lỗi (thiếu hoặc trùng mã trong file)."""
    rows = [list(r) for r in rows]
    if not any(str(c).strip() for r in rows for c in r):
        raise ValueError("File không có dữ liệu")
    code_col, title_col, header = 0, 1, False
    head = [str(c).strip().lower() for c in rows[0]] if rows else []
    if head and head[0] in CHAPTER_CODE_HEADERS:
        header = True
        for i, h in enumerate(head):
            if h in CHAPTER_TITLE_HEADERS:
                title_col = i
    existing = {r["code"] for r in q("SELECT code FROM chapters WHERE textbook_id=?", (textbook_id,), True)}
    seen, out = set(), []
    for n, cells in enumerate(rows, start=1):
        if header and n == 1:
            continue
        cells = [str(c).strip() for c in cells]
        if not any(cells):
            continue
        code = cells[code_col] if len(cells) > code_col else ""
        title = cells[title_col] if len(cells) > title_col else ""
        if not code:
            status, note = "lỗi", "thiếu mã chương"
        elif code in seen:
            status, note = "lỗi", "mã trùng trong file"
        elif code in existing:
            status, note = "bỏ qua", "chương đã có"
        else:
            status, note = "thêm", ""
        if code:
            seen.add(code)
        out.append({"line": n, "code": code, "title": title, "status": status, "note": note})
    return out


def import_chapters(rows, textbook_id, dry_run=False):
    """Nhập hàng loạt chương vào giáo trình đã chọn. dry_run: chỉ kiểm tra, không ghi."""
    plan = plan_chapters(rows, textbook_id)
    counts = {k: sum(1 for r in plan if r["status"] == k) for k in ("thêm", "bỏ qua", "lỗi")}
    if not dry_run and counts["thêm"]:
        with closing(sqlite3.connect(DB_PATH)) as conn:
            conn.executemany("INSERT INTO chapters(textbook_id, code, title) VALUES(?, ?, ?)",
                             [(textbook_id, r["code"], r["title"]) for r in plan if r["status"] == "thêm"])
            conn.commit()
    return {"dry_run": dry_run, "counts": counts, "rows": plan}


STATUS_IMPORT = {**{v.lower(): k for k, v in STATUSES.items()}, **{k: k for k in STATUSES}}


def parse_rows(tables, textbook_id, dry_run=False, docx=True):
    """tables: danh sách bảng, mỗi bảng là danh sách dòng (chuỗi). Dòng đầu là tiêu đề cột.
    Nhận từ bảng Word, CSV hoặc dữ liệu dán (docx=False: không có khái niệm bảng).
    Dòng đầu là tiêu đề cột; cột nhận diện theo tên nên thứ tự cột tùy ý.
    Bắt buộc: Mã chương, Nội dung, Đáp án, Loại, Mức. Tùy chọn: Đoạn văn, Lời giải.
    Mã chương phải đã có trong giáo trình được chọn; không tự tạo chương mới."""
    chapter_by_code = {r["code"]: r["id"] for r in q(
        "SELECT id, code FROM chapters WHERE textbook_id=?", (textbook_id,), True)}
    added, errors, skipped, warnings = 0, [], [], []
    rows_out = []  # xem trước: trạng thái từng dòng
    index = dup_index()
    where = (lambda t, r: f"Bảng {t}, dòng {r}") if docx else (lambda t, r: f"Dòng {r}")
    for t_idx, table in enumerate(tables, 1):
        if not table:
            continue
        cols, missing = map_import_header(table[0])
        if missing:
            names = {"ma": "Mã chương", "noi_dung": "Nội dung", "dap_an": "Đáp án", "loai": "Loại", "muc": "Mức"}
            errors.append(f"{'Bảng ' + str(t_idx) if docx else 'Dữ liệu'}: thiếu cột {', '.join(names[k] for k in missing)} (dòng tiêu đề)")
            continue
        for r_idx, row in enumerate(table[1:], 2):
            cells = [c.strip() for c in row]
            get = lambda k: cells[cols[k]] if k in cols and cols[k] < len(cells) else ""
            if not any(cells):
                continue
            try:
                code = get("ma")
                content = get("noi_dung")
                lines = [ln.strip() for ln in content.splitlines() if ln.strip()]
                solution = get("loi_giai")
                for i, ln in enumerate(lines):
                    if ln.lower().startswith("lời giải:"):
                        solution = "\n".join(lines[i:]).split(":", 1)[1].strip()
                        lines = lines[:i]
                        break
                qtype = (get("loai") or "TN").upper()
                qtype = {"ĐS": "DS", "ĐT": "DT"}.get(qtype, qtype)
                stem, options = [], []
                if qtype in ("TN", "DS"):
                    for ln in lines:
                        m = OPT_RE.match(ln)
                        if m:
                            options.append(m.group(2))
                        elif not options:
                            stem.append(ln)
                        else:
                            options[-1] += " " + ln
                else:
                    stem = lines
                if code not in chapter_by_code:
                    raise ValueError(f"mã chương '{code}' không có trong giáo trình đã chọn")
                muc = get("muc").strip()
                if not muc.isdigit() or not 1 <= int(muc) <= 5:
                    raise ValueError(f"mức '{muc}' phải là số từ 1 đến 5")
                payload = clean_question({
                    "chapter_id": chapter_by_code[code],
                    "bloom": int(muc),
                    "qtype": qtype,
                    "content": "\n".join(stem),  # giữ xuống dòng: câu dẫn nhập và câu hỏi, lời đáp đối thoại
                    "options": options,
                    "answer": get("dap_an"),
                    "solution": solution,
                    "passage": get("doan_van"),
                    "clo": get("clo"),
                    "tags": get("tags"),
                    "word_limit": get("so_tu"),  # điền từ: số từ tối đa, để trống = không giới hạn
                    "status": STATUS_IMPORT.get(get("status").strip().lower(), ""),
                })
                same, other, key = check_duplicate(index, payload)
                label = where(t_idx, r_idx)
                if payload["passage"] and not dry_run and not q(
                        "SELECT 1 FROM passages WHERE chapter_id=? AND content=?",
                        (payload["chapter_id"], payload["passage"]), True):  # đúng từng chữ thì dùng chung, không cảnh báo
                    sim = similar_passage(payload["chapter_id"], payload["passage"])
                    if sim:
                        warnings.append(f"{label}: đoạn văn gần giống đoạn {sim[0]} ({round(sim[1] * 100)}%); kiểm tra lại")
                preview = {"where": label, "code": code, "qtype": qtype, "content": " ".join(stem)[:80]}  # xem trước: một dòng
                if same:
                    ref = same[0][0]
                    ref_txt = f"câu {code_of(ref)}" if isinstance(ref, int) else ref.lower()  # ref là nhãn dòng khi xem trước
                    skipped.append(f"{label}: trùng {ref_txt} trong cùng chương")
                    rows_out.append(dict(preview, status="skip", note=f"trùng {ref_txt}"))
                    continue
                if dry_run:  # chưa ghi; dùng nhãn dòng để so trùng với các dòng sau trong cùng file
                    new_id = f"{label}"
                else:
                    new_id = insert_q(payload)
                index.setdefault(key, []).append((new_id, payload["chapter_id"], payload["content"]))
                added += 1
                note = ""
                if other:
                    note = f"trùng nội dung với câu {code_of(other[0][0])} ở chương khác"
                    warnings.append(f"{label}: {note}")
                rows_out.append(dict(preview, status="warn" if other else "ok", note=note))
            except Exception as e:
                errors.append(f"{where(t_idx, r_idx)}: {e}")
                rows_out.append({"where": where(t_idx, r_idx), "code": "", "qtype": "",
                                 "content": "", "status": "error", "note": str(e)})
    return {"added": added, "skipped": skipped, "warnings": warnings, "errors": errors,
            "dry_run": dry_run, "rows": rows_out}


def pick_base_exam(matrix, qtype=None, usage=None, taken=None, rng=None, counts=None):
    """matrix: [{chapter_id, bloom, n}] -> câu gốc của một phần.
    Ưu tiên câu ra đề ít lần nhất (chưa ra đề = 0), rồi câu có lần dùng gần nhất cũ hơn;
    các câu còn ngang nhau được chọn ngẫu nhiên theo rng (có seed thì tái tạo được).
    Câu thuộc đoạn văn không chọn lẻ ở đây; đoạn văn lấy theo pick_passages.
    qtype: lọc theo dạng câu (TN/DS/DT/TL) của phần; None = mọi dạng."""
    usage = usage or {}
    counts = counts if counts is not None else usage_counts()
    rng = rng or random
    taken = taken if taken is not None else set()  # câu đã chọn cho đề này (mọi phần): không chọn lại
    base, shortage = [], []
    for cell in matrix:
        n = int(cell.get("n", 0))
        if n <= 0:
            continue
        sql = "SELECT * FROM questions WHERE chapter_id=? AND bloom=? AND status='da_duyet' AND passage_id IS NULL"
        args = [int(cell["chapter_id"]), int(cell["bloom"])]
        if qtype:
            sql += " AND qtype=?"
            args.append(qtype)
        pool = [row_to_api(p) for p in q(sql, args, fetch=True) if p["id"] not in taken]
        if len(pool) < n:
            shortage.append(f"{chapter_label(cell['chapter_id'])} / Bloom {cell['bloom']}: cần {n}, có {len(pool)}")
            continue
        rng.shuffle(pool)
        pool.sort(key=lambda it: (counts.get(it["id"], 0), usage.get(it["id"]) or ""))  # ít lần nhất, rồi cũ nhất
        for it in pool[:n]:
            it["last_used"] = usage.get(it["id"])
            taken.add(it["id"])
            base.append(it)
    if shortage:
        raise ValueError("Không đủ câu hỏi: " + "; ".join(shortage))
    return base


def load_items(ids, usage):
    """Lấy đúng các câu theo danh sách id (đã xem trước), giữ thứ tự id."""
    if not ids:
        return []
    rows = {r["id"]: r for r in q(f"SELECT * FROM questions WHERE id IN ({','.join('?' * len(ids))})",
                                  [int(i) for i in ids], fetch=True)}
    out = []
    for i in ids:
        if int(i) not in rows:
            raise ValueError("Có câu đã chọn không còn trong ngân hàng, hãy xem trước lại")
        it = row_to_api(rows[int(i)])
        if it["status"] != "da_duyet":
            raise ValueError(f"Câu {it['code']} chưa được duyệt hoặc đã ngưng dùng; hãy xem trước lại đề")
        it["last_used"] = usage.get(it["id"])
        out.append(it)
    return out


def find_replacement(spec, exclude, usage):
    """Tìm câu thay thế cùng chương, cùng mức Bloom và cùng dạng, chưa có trong đề.
    Ưu tiên câu chưa dùng, rồi câu dùng lâu nhất."""
    if spec.get("passage"):
        raise ValueError("Câu thuộc đoạn văn không đổi riêng lẻ; hãy chọn lại cả phần")
    sql = "SELECT * FROM questions WHERE chapter_id=? AND bloom=? AND qtype=? AND passage_id IS NULL AND status='da_duyet'"
    args = [int(spec["chapter_id"]), int(spec["bloom"]), str(spec["qtype"])]
    pool = [row_to_api(p) for p in q(sql, args, fetch=True) if p["id"] not in exclude]
    if not pool:
        raise ValueError("Không còn câu thay thế cùng chương, mức Bloom và dạng")
    random.shuffle(pool)
    pool.sort(key=lambda it: usage.get(it["id"]) or "")
    it = pool[0]
    it["last_used"] = usage.get(it["id"])
    return it


def pick_passages(entries, qtype, usage, taken, rng=None):
    """Ma trận theo đoạn: entries = [{chapter_id, n}] -> lấy n đoạn của chương (ưu tiên đoạn ít dùng),
    mỗi đoạn lấy mọi câu chưa dùng (đúng dạng của phần nếu có). Đoạn không có câu phù hợp thì bỏ qua."""
    out = []
    for e in entries or []:
        n = int(e.get("n", 0))
        if n <= 0:
            continue
        cands = []
        for p in q("SELECT id FROM passages WHERE chapter_id=?", (int(e["chapter_id"]),), fetch=True):
            sql = "SELECT * FROM questions WHERE passage_id=? AND status='da_duyet'"
            args = [p["id"]]
            if qtype:
                sql += " AND qtype=?"
                args.append(qtype)
            items = [row_to_api(r) for r in q(sql + " ORDER BY id", args, fetch=True) if r["id"] not in taken]
            if items:
                cands.append(items)
        if len(cands) < n:
            raise ValueError(f"Không đủ đoạn văn: {chapter_label(int(e['chapter_id']))} cần {n}, có {len(cands)}")
        (rng or random).shuffle(cands)
        # đoạn có câu được dùng gần nhất xếp sau (chưa dùng xếp trước)
        cands.sort(key=lambda items: max((usage.get(it["id"]) or "") for it in items))
        for items in cands[:n]:
            for it in items:
                it["last_used"] = usage.get(it["id"])
                taken.add(it["id"])
                out.append(it)
    return out


def section_has_content(s):
    return any(int(c.get("n", 0)) > 0 for c in s.get("matrix", [])) or \
           any(int(c.get("n", 0)) > 0 for c in s.get("passages", []) or [])


def build_section(s, usage, taken, ids=None, rng=None, counts=None):
    """Chọn câu cho một phần: theo ma trận chương × Bloom và theo đoạn văn. ids: danh sách đã xem trước."""
    if ids is not None:
        base = load_items(ids, usage)
        taken.update(int(i) for i in ids)
        return base
    qtype = s.get("qtype") or None
    base = []
    if any(int(c.get("n", 0)) > 0 for c in s.get("matrix", [])):
        base += pick_base_exam(s.get("matrix", []), qtype, usage, taken, rng, counts)
    base += pick_passages(s.get("passages", []), qtype, usage, taken, rng)
    return base


def preview_exam(sections):
    """Chọn câu cho mọi phần (không ghi lịch sử). Trả về danh sách song song với sections."""
    usage = usage_map()
    taken = set()
    out = []
    for s in sections:
        if not section_has_content(s):
            out.append({"items": []})
            continue
        out.append({"items": build_section(s, usage, taken)})
    return out


def shuffle_question(item, rng, letter_counts=None, tries=20):
    """Xáo phương án TN; phương án có dấu # giữ nguyên vị trí. Trả về (item, thứ tự mới).
    letter_counts: số lần mỗi chữ cái đã làm đáp án trong mã đề này. Nếu có, thử nhiều hoán vị
    và chọn hoán vị mà đáp án rơi vào chữ đang được dùng ít nhất, để tránh lệch A/B/C/D."""
    if item["qtype"] != "TN":
        return item, None
    opts = item["options"]
    correct = LETTERS.index(item["answer"])
    movable = [i for i, o in enumerate(opts) if not o.startswith("#")]
    best = None
    for _ in range(tries if letter_counts is not None else 1):
        shuffled = movable[:]
        rng.shuffle(shuffled)
        order = list(range(len(opts)))
        for pos, src in zip(movable, shuffled):
            order[pos] = src
        new_answer = LETTERS[order.index(correct)]
        load = letter_counts.get(new_answer, 0) if letter_counts is not None else 0
        if best is None or load < best[0]:
            best = (load, order, new_answer)
        if load == 0 or (letter_counts is not None and load == min(letter_counts.get(c, 0) for c in LETTERS[:len(opts)])):
            break
    _, order, new_answer = best
    if letter_counts is not None:
        letter_counts[new_answer] = letter_counts.get(new_answer, 0) + 1
    new_opts = [opts[i].lstrip("#").strip() for i in order]
    return dict(item, options=new_opts, answer=new_answer), order


def make_variant(base, shuffle, rng):
    """Một phần của mã đề. Các câu cùng đoạn văn luôn đứng liền nhau, theo thứ tự id.
    shuffle=True: xáo thứ tự các nhóm (câu đơn hoặc cả đoạn) và phương án TN;
    shuffle=False: giữ thứ tự ngân hàng."""
    groups = {}
    for it in base:
        key = (f"p{it['passage_id']}" if it.get("passage_id") else
               f"t{it['passage']}" if it.get("passage") else f"#{it['id']}")
        groups.setdefault(key, []).append(it)
    units = [sorted(g, key=lambda it: it["id"]) for g in groups.values()]
    units.sort(key=lambda g: g[0]["id"])
    if shuffle:
        rng.shuffle(units)
        counts = {}  # cân bằng đáp án trong cùng một mã đề
        return [it if is_match(it) else shuffle_question(it, rng, counts)[0] for g in units for it in g]
    return [it for g in units for it in g]


def flatten(exam):
    """exam: [(tiêu đề, items)] -> danh sách (số thứ tự toàn đề, tiêu đề phần, item)."""
    out, n = [], 0
    for title, items in exam:
        for it in items:
            n += 1
            out.append((n, title, it))
    return out


# ---------------------------------------------------------------- docx output
def _set_text(p, text):
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def _all_paragraphs(doc):
    yield from doc.paragraphs
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from cell.paragraphs



RUBY_RE = re.compile(r"\{([^{}|]+)\|([^{}]+)\}")  # cú pháp furigana: {試験|しけん}


def _wrun(text):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _add_ruby(p, base, kana):
    ruby = OxmlElement("w:ruby")
    rpr = OxmlElement("w:rubyPr")
    for tag, val in (("w:rubyAlign", "center"), ("w:hps", "10"), ("w:hpsRaise", "18"),
                     ("w:hpsBaseText", "22"), ("w:lid", "ja-JP")):
        e = OxmlElement(tag)
        e.set(qn("w:val"), val)
        rpr.append(e)
    ruby.append(rpr)
    rt = OxmlElement("w:rt")
    rt.append(_wrun(kana))
    ruby.append(rt)
    rb = OxmlElement("w:rubyBase")
    rb.append(_wrun(base))
    ruby.append(rb)
    holder = OxmlElement("w:r")  # Word đặt w:ruby bên trong một w:r
    holder.append(ruby)
    p._p.append(holder)


KANJI_RUN_RE = re.compile(r"[\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF々〆ヵヶ]+")
_kakasi = None


def _get_kakasi():
    """Bộ đổi chữ Hán sang kana (pykakasi), nạp lần đầu khi cần. Trả về None nếu chưa cài."""
    global _kakasi
    if _kakasi is None:
        try:
            import pykakasi
        except ImportError:
            return None
        _kakasi = pykakasi.kakasi()
    return _kakasi


def auto_ruby(text):
    """Thêm {chữ Hán|cách đọc} cho các cụm chữ Hán chưa có cú pháp furigana.
    Cách đọc lấy từ pykakasi nên có thể sai với từ đọc đặc biệt; người soạn nên kiểm tra
    và viết tay {chữ|cách đọc} để ghi đè. Phần đã có {…|…} được giữ nguyên."""
    kk = _get_kakasi()
    if kk is None:
        raise ValueError("Chưa cài pykakasi nên không tự thêm furigana được (pip install pykakasi)")
    out, pos = [], 0
    for m in RUBY_RE.finditer(text):
        out.append(_auto_plain(kk, text[pos:m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(_auto_plain(kk, text[pos:]))
    return "".join(out)


def _auto_plain(kk, seg):
    def repl(m):
        run = m.group(0)
        reading = "".join(i["hira"] for i in kk.convert(run))
        return "{%s|%s}" % (run, reading) if reading and reading != run else run
    return KANJI_RUN_RE.sub(repl, seg)


def add_rich(p, text, bold=False, mode="ruby"):
    """Thêm văn bản vào đoạn, đổi cú pháp {chữ|cách đọc} thành furigana.
    mode='ruby': chữ đọc phía trên (chuẩn Word); mode='paren': chữ(cách đọc) cùng dòng (dự phòng).
    Thêm '+auto' (vd 'ruby+auto') để tự thêm furigana cho chữ Hán chưa có cú pháp."""
    auto = mode.endswith("+auto")
    mode = mode.split("+")[0]
    if auto:
        text = auto_ruby(text)
    pos = 0
    for m in RUBY_RE.finditer(text):
        if m.start() > pos:
            p.add_run(text[pos:m.start()]).bold = bold
        if mode == "paren":
            p.add_run(m.group(1)).bold = bold
            small = p.add_run(f"({m.group(2)})")
            small.font.size = Pt(8)
            small.bold = bold
        else:
            _add_ruby(p, m.group(1), m.group(2))
        pos = m.end()
    if pos < len(text):
        p.add_run(text[pos:]).bold = bold


def _set_borders(tbl):
    """Kẻ khung đơn cho bảng (không cần kiểu 'Table Grid' có sẵn trong mẫu)."""
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)


def emit_passage(doc, marker, text, mode="ruby"):
    """Đoạn văn dùng chung: khung một ô, in trước câu đầu tiên của nhóm."""
    spacer = marker.insert_paragraph_before("")
    tbl = doc.add_table(rows=1, cols=1)
    _set_borders(tbl)
    cell = tbl.cell(0, 0)
    for i, line in enumerate(text.splitlines() or [""]):
        para = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        add_rich(para, line, mode=mode)
    marker._p.addprevious(tbl._tbl)
    return spacer


def emit_options(doc, marker, options, mode="ruby"):
    """Phương án TN xếp hai cột, đọc theo cột: A C / B D (đúng mẫu đề trường)."""
    half = (len(options) + 1) // 2
    tbl = doc.add_table(rows=half, cols=2)
    for i, opt in enumerate(options):
        row, col = (i, 0) if i < half else (i - half, 1)
        para = tbl.cell(row, col).paragraphs[0]
        para.add_run(f"{LETTERS[i]}. ")
        add_rich(para, opt, mode=mode)
    marker._p.addprevious(tbl._tbl)


def emit_heading_list(doc, marker, options, mode="ruby"):
    """Danh sách tiêu đề dùng chung cho một nhóm nối tiêu đề, in một lần trước câu đầu của nhóm."""
    head = marker.insert_paragraph_before("")
    head.add_run("Danh sách tiêu đề (chọn một tiêu đề cho mỗi đoạn):").bold = True
    for i, opt in enumerate(options):
        line = marker.insert_paragraph_before("")
        line.add_run(f"    {LETTERS[i]}. ")
        add_rich(line, opt, mode=mode)


def emit_question(doc, marker, n, it, mode="ruby"):
    """Nội dung câu: mỗi dòng trong nội dung xuống hàng. Nếu có từ hai dòng trở lên, dòng đầu là
    câu dẫn nhập và in đậm; các dòng sau (câu hỏi, lời đáp đối thoại) in thường."""
    p = marker.insert_paragraph_before("")
    p.add_run(f"Câu {n}. ").bold = True
    lines = [ln.strip() for ln in (it["content"] or "").splitlines() if ln.strip()] or [""]
    bold_lead = len(lines) > 1 and not SPEAKER_LINE.match(lines[0])
    for i, ln in enumerate(lines):
        if i:
            p.add_run().add_break()
        add_rich(p, ln, bold=(bold_lead and i == 0), mode=mode)
    if it.get("points") is not None:
        p.add_run(f" ({fmt_points(it['points'])} điểm)").italic = True
    if is_match(it):
        marker.insert_paragraph_before("    Đáp án (chữ cái tiêu đề): ……")
    elif it["qtype"] == "TN":
        emit_options(doc, marker, it["options"], mode)
    elif it["qtype"] == "DS":
        for k, opt in enumerate(it["options"]):
            line = marker.insert_paragraph_before("")
            line.add_run(f"    {'abcdefgh'[k]}) ")
            add_rich(line, opt, mode=mode)
    elif it["qtype"] == "DT":
        limit = f" (không quá {it['word_limit']} từ)" if it.get("word_limit") else ""
        if is_label(it):
            marker.insert_paragraph_before(f"    {LABEL_HINT}")
        marker.insert_paragraph_before(f"    Đáp án{limit}: ……………………………")
    emit_media(doc, marker, it.get("media"))


def emit_media(doc, marker, media):
    """Hình ảnh nhúng vào đề; âm thanh ghi tên file (file nằm trong thư mục Am_thanh)."""
    for m in media or []:
        if m["kind"] == "image" and os.path.isfile(os.path.join(MEDIA_DIR, m["file"])):
            pic = doc.add_paragraph()  # đoạn mới có đủ đối tượng cha để chèn hình, rồi chuyển lên trước dấu đánh dấu
            marker._p.addprevious(pic._p)
            pic.add_run().add_picture(os.path.join(MEDIA_DIR, m["file"]), width=Cm(7))
        elif m["kind"] == "audio":
            marker.insert_paragraph_before(f"    🔊 Nghe: {m['name']}")


def meta_values(meta, code, n_questions):
    mon = meta.get("ten_hp") or meta.get("mon", "")
    return {
        "{{MON}}": mon, "{{TEN_HP}}": mon,
        "{{MA_HP}}": meta.get("ma_hp", ""), "{{SO_TC}}": str(meta.get("so_tc", "")),
        "{{KHOA_HOC}}": meta.get("khoa_hoc", ""), "{{NAM_HOC}}": meta.get("nam_hoc", ""),
        "{{HOC_KY}}": meta.get("hoc_ky", ""), "{{TEN_DE}}": meta.get("ten_de", ""),
        "{{THOI_GIAN}}": str(meta.get("thoi_gian", "")), "{{MA_DE}}": code,
        "{{SO_CAU}}": str(n_questions), "{{SO_TRANG}}": str(meta.get("so_trang", "")),
    }


def _every_paragraph(doc):
    """Mọi đoạn trong tài liệu, kể cả trong bảng lồng nhau và khung."""
    from docx.text.paragraph import Paragraph
    for p_el in doc.element.body.iter(qn("w:p")):
        yield Paragraph(p_el, None)


def fill_placeholders(doc, values):
    for p in _every_paragraph(doc):
        if "{{" in p.text:
            txt = p.text
            for k, v in values.items():
                txt = txt.replace(k, v)
            _set_text(p, txt)


def render_docx(template_path, meta, exam, code, answer_key, mode="ruby"):
    doc = Document(template_path)
    flat = flatten(exam)
    fill_placeholders(doc, meta_values(meta, code, len(flat)))
    marker = next((p for p in _every_paragraph(doc) if "{{CAU_HOI}}" in p.text), None)
    if marker is None:
        marker = doc.add_paragraph()
    last_title, last_passage, last_match = None, None, None
    for n, title, it in flat:
        if title and title != last_title:  # tiêu đề phần: "I. ...", in một lần trước câu đầu của phần
            head = marker.insert_paragraph_before(title)
            head.runs[0].bold = True
            last_title = title
        if it.get("passage") and it["passage"] != last_passage:
            emit_passage(doc, marker, it["passage"], mode)
            emit_media(doc, marker, it.get("passage_media"))
        last_passage = it.get("passage") or None
        if is_match(it):
            key = (it.get("passage_id"), tuple(it["options"]))
            if key != last_match:
                emit_heading_list(doc, marker, it["options"], mode)
                last_match = key
        else:
            last_match = None
        emit_question(doc, marker, n, it, mode)
    _set_text(marker, "")

    if answer_key:
        doc.add_paragraph("")
        doc.add_paragraph(f"ĐÁP ÁN MÃ {code}").runs[0].bold = True
        doc.add_paragraph("   ".join(f"{n}-{answer_text(it)}" for n, _, it in flat))
        for line in rubric_lines(flat):
            doc.add_paragraph(line)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


KEY_PER_ROW = 10  # số câu mỗi dòng của lưới đáp án


def clo_lines(flat):
    """Đáp án theo CĐR: các câu gom theo CLO, số câu trắc nghiệm và tự luận, điểm của mỗi CLO."""
    by = {}
    for n, _, it in flat:
        by.setdefault(it.get("clo") or "", []).append((n, it))
    lines = ["ĐÁP ÁN THEO CHUẨN ĐẦU RA"]
    for clo in sorted(by):
        items = by[clo]
        tn = sum(1 for _, it in items if it["qtype"] != "TL")
        tl = sum(1 for _, it in items if it["qtype"] == "TL")
        pts = sum(it["points"] for _, it in items if it.get("points") is not None)
        extra = f", {fmt_points(pts)} điểm" if pts else ""
        lines.append(f"{clo}: câu {', '.join(str(n) for n, _ in items)} ({tn} trắc nghiệm, {tl} tự luận{extra})")
    return lines


def rubric_table(doc, marker, n, it):
    """Bảng tiêu chí × mức 0-4 cho một câu tự luận (theo mẫu 'đánh giá theo CĐR')."""
    crits = it.get("rubric") or []
    lines_before = f"Câu {n} (tự luận): hướng dẫn chấm theo tiêu chí"
    marker.insert_paragraph_before(lines_before)
    t = doc.add_table(rows=1 + len(crits), cols=7)
    _set_borders(t)
    head = ["Tiêu chí", "Điểm", "Mức 0 (<40%)", "Mức 1 (40–54%)", "Mức 2 (55–69%)", "Mức 3 (70–84%)", "Mức 4 (85–100%)"]
    for c, h in zip(t.rows[0].cells, head):
        c.text = h
    for r, crit in enumerate(crits, 1):
        cells = t.rows[r].cells
        cells[0].text = crit.get("text", "")
        cells[1].text = fmt_points(crit["points"]) if crit.get("points") is not None else ""
        levels = crit.get("levels") or [""] * 5
        for j in range(5):
            cells[2 + j].text = levels[j] if j < len(levels) else ""
    from docx.shared import Cm  # vừa khổ giấy: tiêu chí rộng, điểm hẹp, năm mức đều nhau
    t.autofit = False
    widths = [Cm(3.1), Cm(1.5)] + [Cm(2.38)] * 5
    for col, w in zip(t.columns, widths):  # lưới cột (tblGrid) để Word và LibreOffice cùng áp dụng
        col.width = w
    for row in t.rows:
        for c, w in zip(row.cells, widths):
            c.width = w
    marker._p.addprevious(t._tbl)


def render_key_doc(template_path, meta, code, exam, clo_mode=False):
    """Đáp án theo mẫu trường: bảng lưới 10 câu mỗi dòng, dòng 'Câu' và dòng 'Đáp án'."""
    doc = Document(template_path)
    flat = flatten(exam)
    fill_placeholders(doc, meta_values(meta, code, len(flat)))
    ans = [(n, answer_text(it)) for n, _, it in flat]
    # lưới đáp án: 10 câu mỗi dòng cho mọi số câu (50, 90...), dòng sau tự xuống; không cố định theo đề
    per_row = KEY_PER_ROW
    groups = [ans[i:i + per_row] for i in range(0, len(ans), per_row)]
    marker = next((p for p in _every_paragraph(doc) if "{{BANG_DAP_AN}}" in p.text), None)
    if marker is None:
        marker = doc.add_paragraph()
    tbl = doc.add_table(rows=2 * len(groups), cols=per_row + 1)
    _set_borders(tbl)
    from docx.shared import Cm  # cột tiêu đề rộng để không ngắt chữ "Câu", "Đáp án"; các cột đáp án đều nhau
    tbl.autofit = False
    for row in tbl.rows:
        for i, c in enumerate(row.cells):
            c.width = Cm(1.8) if i == 0 else Cm((16.5 - 1.8) / per_row)
    for gi, g in enumerate(groups):
        r1, r2 = tbl.rows[2 * gi], tbl.rows[2 * gi + 1]
        r1.cells[0].text, r2.cells[0].text = "Câu", "Đáp án"
        for j, (n, a) in enumerate(g):
            r1.cells[j + 1].text, r2.cells[j + 1].text = str(n), a
    marker._p.addprevious(tbl._tbl)
    if clo_mode:
        for line in clo_lines(flat):
            marker.insert_paragraph_before(line)
    for n, _, it in flat:
        if it["qtype"] == "TL" and it.get("rubric"):
            rubric_table(doc, marker, n, it)
    for line in rubric_lines(flat):
        marker.insert_paragraph_before(line)
    _set_text(marker, "")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def question_lines(i, it):
    """Các dòng hiển thị của một câu theo loại."""
    lines = [f"Câu {i}. {it['content']}"]
    if it["qtype"] == "TN":
        lines += [f"    {LETTERS[k]}. {opt}" for k, opt in enumerate(it["options"])]
    elif it["qtype"] == "DS":
        lines += [f"    {'abcdefgh'[k]}) {opt}" for k, opt in enumerate(it["options"])]
    elif it["qtype"] == "DT":
        lines.append("    Đáp án: ……………………………")
    return lines


def answer_text(it):
    """Đáp án ngắn gọn để in trong bảng đáp án."""
    if it["qtype"] == "TN":
        return it["answer"]
    if it["qtype"] == "DS":
        return "".join("Đ" if c == "D" else "S" for c in it["answer"])
    if it["qtype"] == "DT":
        return it["answer"].replace("|", " / ")
    return "TL"


def fmt_points(x):
    return ("%g" % x).replace(".", ",")


def rubric_lines(flat):
    """Hướng dẫn chấm tự luận và tổng điểm, in dưới bảng đáp án."""
    lines = [f"Câu {n}: {it['answer']}" + (f" ({fmt_points(it['points'])} điểm)" if it.get("points") is not None else "")
             for n, _, it in flat if it["qtype"] == "TL" and it["answer"]]
    pts = [it["points"] for _, _, it in flat if it.get("points") is not None]
    if pts:
        lines.append(f"Tổng điểm: {fmt_points(sum(pts))}")
    if lines:
        lines.insert(0, "HƯỚNG DẪN CHẤM TỰ LUẬN")
    return lines



def render_key(code, exam):
    """Đáp án riêng khi chưa có mẫu đáp án: bảng Câu | Phần | Loại | Đáp án."""
    doc = Document()
    doc.add_paragraph(f"ĐÁP ÁN MÃ {code}").runs[0].bold = True
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for c, h in zip(table.rows[0].cells, ["Câu", "Phần", "Loại", "Đáp án"]):
        c.text = h
    for n, title, it in flatten(exam):
        cells = table.add_row().cells
        for c, v in zip(cells, [str(n), title, it["qtype"], answer_text(it)]):
            c.text = v
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _plain_header(doc, left_lines, right_lines):
    """Khối tiêu đề hai cột không kẻ khung, giống mẫu đề trường."""
    tbl = doc.add_table(rows=1, cols=2)
    for cell, lines in ((tbl.cell(0, 0), left_lines), (tbl.cell(0, 1), right_lines)):
        for i, (text, bold, center) in enumerate(lines):
            para = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            if center:
                para.alignment = 1
            para.add_run(text).bold = bold
    return tbl


def make_official_templates():
    """Mẫu đề và mẫu đáp án theo định dạng trường. Khoa, tên trường và cấu trúc cố định;
    học phần, mã HP, số tín chỉ, khóa, năm học, học kỳ, thời gian và mã đề được điền khi xuất."""
    d = Document()
    _plain_header(d,
        [("TRƯỜNG ĐẠI HỌC QUY NHƠN", True, True), ("KHOA NGOẠI NGỮ", True, True), ("", False, False),
         ("ĐỀ SỐ: {{MA_DE}}", True, False)],
        [("ĐỀ THI KẾT THÚC HỌC PHẦN", True, True),
         ("ĐÀO TẠO TRÌNH ĐỘ ĐẠI HỌC HỆ CHÍNH QUY", True, True),
         ("HỌC KỲ {{HOC_KY}} NĂM HỌC {{NAM_HOC}}", True, True),
         ("Học phần: {{TEN_HP}}", False, False),
         ("Mã HP {{MA_HP}}          Số tín chỉ: {{SO_TC}}", False, False),
         ("Khóa {{KHOA_HOC}}", False, False),
         ("Thời gian làm bài: {{THOI_GIAN}} phút (Không kể thời gian phát đề)", False, False),
         ("(Sinh viên không sử dụng tài liệu tham khảo để làm bài)", False, False)])
    d.add_paragraph("{{CAU_HOI}}")
    d.save(os.path.join(TPL_DIR, "de_thi_chuan.docx"))

    d = Document()
    _plain_header(d,
        [("TRƯỜNG ĐẠI HỌC QUY NHƠN", True, True), ("KHOA NGOẠI NGỮ", True, True), ("", False, False),
         ("Đáp án đề số: {{MA_DE}}", False, False)],
        [("ĐÁP ÁN ĐỀ THI KẾT THÚC HỌC PHẦN", True, True),
         ("ĐÀO TẠO TRÌNH ĐỘ ĐẠI HỌC HỆ CHÍNH QUY", True, True),
         ("HỌC KỲ {{HOC_KY}} NĂM HỌC {{NAM_HOC}}", True, True),
         ("Học phần: {{TEN_HP}}", False, False),
         ("Mã HP {{MA_HP}}          Số tín chỉ: {{SO_TC}}", False, False),
         ("Khóa {{KHOA_HOC}}", False, False),
         ("Thời gian làm bài: {{THOI_GIAN}} phút (Không kể thời gian phát đề)", False, False)])
    d.add_paragraph("{{BANG_DAP_AN}}")
    d.add_paragraph("Ghi chú: - Giảng viên ghi mã đề và điền đáp án A, B, C, D vào ô tương ứng.")
    d.save(os.path.join(TPL_DIR, "dap_an_chuan.docx"))


def make_import_template():
    """File mẫu nhập câu hỏi: đúng tiêu đề cột và các ví dụ cho từng loại câu."""
    d = Document()
    d.add_paragraph("MẪU NHẬP CÂU HỎI (lưu thành .docx, mỗi bảng là một nhóm câu, dòng 1 là tiêu đề cột)")
    d.add_paragraph("Cột: Mã chương (bắt buộc, phải có trong giáo trình đã chọn) | Nội dung (bắt buộc) | "
                    "Đáp án (bắt buộc) | Loại: TN, DS, DT, TL | Mức: 1–5 | Đoạn văn (tùy chọn, các câu cùng đoạn ghi giống hệt) | "
                    "Lời giải (tùy chọn) | CLO (tùy chọn, một CLO cho mỗi câu, ví dụ CLO1) | "
                    "Thẻ (tùy chọn, nhiều thẻ cách nhau bằng dấu phẩy: tfng, noitieude, ghinhan) | "
                    "Số từ (tùy chọn, chỉ cho DT: số từ tối đa 1–20, để trống = không giới hạn). Thứ tự cột tùy ý. "
                    "Furigana viết dạng {試験|しけん}. "
                    "TN: mỗi phương án một dòng bắt đầu bằng A. B. C. D.; # đầu phương án = cố định vị trí. "
                    "DS: mỗi ý một dòng bắt đầu a) b) c); đáp án Đ S Đ S. DT: nhiều đáp án cách nhau bằng |. "
                    "TFNG (thẻ tfng): không ghi phương án; Đáp án A = Đúng, B = Sai, C = Không có thông tin. "
                    "Nối tiêu đề (thẻ noitieude): các câu cùng nhóm ghi cùng Đoạn văn và cùng danh sách tiêu đề A. B. C… "
                    "(tối đa 12); Đáp án là chữ cái tiêu đề. "
                    "Ghi nhãn sơ đồ (thẻ ghinhan, loại DT): Đáp án là tên cần ghi cho vị trí đánh số; "
                    "hình ảnh được thêm trong app sau khi nhập.")
    d.add_paragraph("Hướng dẫn cho AI khi soạn câu hỏi: mỗi câu là một dòng trong bảng; không ghi đáp án hay "
                    "lời giải vào cột Nội dung; giữ đúng tên tiêu đề cột như trên; Mã chương phải lấy từ danh sách "
                    "chương của giáo trình đích; Loại chỉ ghi TN, DS, DT hoặc TL; Mức chỉ ghi số 1 đến 5; "
                    "TN ít nhất 2 phương án và Đáp án là một chữ cái (A, B, C…); DS là các ý a) b) c) và Đáp án "
                    "gồm Đ/S đủ số ý; DT ghi Đáp án là từ cần điền, Số từ chỉ khi giới hạn số từ; TL để trống Đáp án; "
                    "Đoạn văn chỉ điền khi câu thuộc một đoạn, và các câu cùng đoạn phải ghi giống hệt nhau; "
                    "CLO: ghi một CLO cho mỗi câu, ví dụ CLO1; câu đối đáp hoặc tự luận viết hội thoại ghi CLO3; để trống nếu chưa rõ; "
                    "Thẻ: chỉ dùng tfng, noitieude, ghinhan khi đúng dạng câu, để trống với câu thường; "
                    "TFNG và nối tiêu đề chỉ dùng cho Loại TN; ghinhan chỉ dùng cho Loại DT. "
                    "không thêm cột ngoài danh sách trên, không gộp ô, không để dòng trống giữa các câu.")
    headers = ["Mã chương", "Nội dung", "Đáp án", "Loại", "Mức", "Đoạn văn", "Lời giải", "CLO", "Thẻ", "Số từ"]
    t = d.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, headers):
        c.text = h
    examples = [
        ("B01", "{試験|しけん}の{結果|けっか}が悪かった。\nA. おけばよかった\nB. おくべきだ", "A", "TN", "1", "", "", "CLO1", "", ""),
        ("B01", "「インターンシップを通して」と同じ意味はどれですか。\nA. 間に\nB. おかげで", "B", "TN", "2", "私は先月、インターンシップに参加しました。", "", "CLO1", "", ""),
        ("B01", "Câu đúng/sai:\na) 私は学生です。\nb) 本を読むです。", "ĐS", "DS", "2", "", "", "CLO1", "", ""),
        ("B01", "Thư viện được xây dựng năm 1990.", "A", "TN", "2", "Thư viện trường được xây dựng năm 1990 và mở rộng năm 2015.", "", "CLO2", "tfng", ""),
        ("B01", "Đoạn 1\nA. Lợi ích của thư viện\nB. Lịch sử xây dựng\nC. Cách mượn sách", "A", "TN", "2", "Đoạn 1: Nhiều sinh viên chọn học ở thư viện vì yên tĩnh.\nĐoạn 2: Thư viện có nhiều sách mới.", "", "CLO2", "noitieude", ""),
        ("B02", "Vị trí 1 trên sơ đồ tim", "động mạch chủ", "DT", "1", "", "", "CLO1", "ghinhan", ""),
        ("B02", "学校の___", "がっこう", "DT", "1", "", "", "CLO1", "", ""),
        ("B02", "Thư viện có ___ tầng.", "ba|3", "DT", "1", "", "", "CLO1", "", "2"),
        ("B02", "A: どこへ行きますか。\nB: 駅へ行きます。\n\nこの会話の続きを書いてください。", "", "TL", "3", "", "会話の内容と文法の正確さ", "CLO3", "", ""),
    ]
    for ex in examples:
        cells = t.add_row().cells
        for c, v in zip(cells, ex):
            c.text = v
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def make_sample_templates():
    """Tạo 2 template mẫu để người dùng thay bằng mẫu trường mình."""
    d = Document()
    d.add_paragraph("{{MON}} - {{TEN_DE}}").runs[0].bold = True
    d.add_paragraph("Thời gian làm bài: {{THOI_GIAN}} phút        Mã đề: {{MA_DE}}")
    d.add_paragraph("Số câu: {{SO_CAU}}")
    d.add_paragraph("{{CAU_HOI}}")
    d.save(os.path.join(TPL_DIR, "mau_chuan.docx"))

    d = Document()
    d.add_paragraph("ĐỀ {{MA_DE}} - {{MON}}")
    d.add_paragraph("Họ và tên: ..............................   Thời gian: {{THOI_GIAN}}")
    d.add_paragraph("{{CAU_HOI}}")
    d.save(os.path.join(TPL_DIR, "mau_gon.docx"))


# ---------------------------------------------------------------- stats & backup
def record_results(exam_code, rows):
    """rows: [(mã câu, tỷ lệ đúng %)]. Chỉ thêm dòng mới; kết quả cũ của mã đề được giữ nguyên."""
    exam_code = str(exam_code).strip()
    if not exam_code:
        raise ValueError("Thiếu mã đề")

    now = datetime.now().isoformat(timespec="seconds")
    added, errors, to_insert = 0, [], []
    for n, (code, pct) in enumerate(rows, 1):
        code = str(code).strip()
        try:
            pct = float(str(pct).replace(",", ".").replace("%", "").strip())
        except ValueError:
            errors.append(f"Dòng {n}: tỷ lệ '{pct}' không phải số")
            continue
        if not 0 <= pct <= 100:
            errors.append(f"Dòng {n}: tỷ lệ phải từ 0 đến 100")
            continue
        hit = q("SELECT id FROM questions WHERE code=?", (code,), fetch=True)
        if not hit:
            errors.append(f"Dòng {n}: không có câu mã '{code}'")
            continue
        to_insert.append((hit[0]["id"], exam_code, pct, now))
        added += 1
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.executemany("INSERT INTO question_results(question_id, exam_code, correct_pct, recorded_at) VALUES(?,?,?,?)", to_insert)
        conn.commit()
    return {"added": added, "errors": errors}


def difficulty():
    """Tỷ lệ đúng thực tế: trung bình theo mức Bloom và các câu dễ/khó bất thường.
    Mỗi câu lấy trung bình các lần thi đã nhập; mức 'bất thường' là ≤20% hoặc ≥95%."""
    rows = q("""SELECT qs.id, qs.code, qs.bloom, qs.content, AVG(r.correct_pct) AS mean, COUNT(r.id) AS n
                FROM question_results r JOIN questions qs ON qs.id = r.question_id
                GROUP BY qs.id ORDER BY qs.code""", fetch=True)
    by_bloom = {b: [] for b in BLOOM}
    flags = []
    for r in rows:
        by_bloom[r["bloom"]].append(r["mean"])
        if r["mean"] <= 20:
            flags.append({"code": r["code"], "mean": round(r["mean"], 1), "n": r["n"], "flag": "khó bất thường", "content": r["content"][:120]})
        elif r["mean"] >= 95:
            flags.append({"code": r["code"], "mean": round(r["mean"], 1), "n": r["n"], "flag": "dễ bất thường", "content": r["content"][:120]})
    blooms = [{"bloom": b, "label": BLOOM[b], "questions": len(v),
               "mean": round(sum(v) / len(v), 1) if v else None} for b, v in by_bloom.items()]
    return {"blooms": blooms, "flags": flags, "total_questions": len(rows)}


def list_passages():
    rows = q("""SELECT p.id, p.code, p.content, p.media, p.chapter_id, ({CL}) AS chapter,
                (SELECT COUNT(*) FROM questions qs WHERE qs.passage_id = p.id) AS count
                FROM passages p JOIN chapters ch ON ch.id = p.chapter_id
                JOIN textbooks t ON t.id = ch.textbook_id
                ORDER BY p.code""".replace("{CL}", CHAPTER_LABEL), fetch=True)
    for r in rows:
        r["media"] = json.loads(r["media"] or "[]")
    return rows


def update_passage(pid, content):
    content = (content or "").strip()
    if not content:
        raise ValueError("Nội dung đoạn không được trống")
    row = q("SELECT chapter_id FROM passages WHERE id=?", (pid,), True)
    if not row:
        raise ValueError("Đoạn văn không tồn tại")
    if q("SELECT id FROM passages WHERE chapter_id=? AND content=? AND id<>?", (row[0]["chapter_id"], content, pid), True):
        raise ValueError("Trong chương đã có đoạn có nội dung này")
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("UPDATE passages SET content=? WHERE id=?", (content, pid))
        conn.execute("UPDATE questions SET passage=? WHERE passage_id=?", (content, pid))
        conn.commit()


def delete_passage(pid):
    n = q("SELECT COUNT(*) AS n FROM questions WHERE passage_id=?", (pid,), True)[0]["n"]
    if n:
        raise ValueError(f"Còn {n} câu thuộc đoạn này; hãy gỡ hoặc xóa các câu đó trước")
    q("DELETE FROM passages WHERE id=?", (pid,))


def stats():
    tree = get_tree()
    counts = {}  # chương -> mức -> số câu (mọi dạng)
    by_qt = {}   # chương -> dạng -> mức -> số câu (để ô ma trận của từng phần chỉ đếm đúng dạng)
    for r in q("SELECT chapter_id, bloom, qtype, status, (passage_id IS NULL) AS np, COUNT(*) AS n FROM questions GROUP BY chapter_id, bloom, qtype, status, np", fetch=True):
        c = counts.setdefault(r["chapter_id"], {b: 0 for b in BLOOM})
        c[r["bloom"]] += r["n"]
        # ô ma trận chỉ tính câu đã duyệt và không thuộc đoạn (câu đoạn văn chỉ ra đề theo cả đoạn)
        if r["status"] == "da_duyet" and r["np"]:
            by_qt.setdefault(r["chapter_id"], {}).setdefault(r["qtype"], {b: 0 for b in BLOOM})[r["bloom"]] += r["n"]
    # số câu chưa từng ra đề: chương -> dạng -> mức -> số câu
    unused = {}
    for r in q("""SELECT qs.chapter_id, qs.qtype, qs.bloom, COUNT(*) AS n FROM questions qs
                  WHERE qs.status='da_duyet' AND qs.passage_id IS NULL
                    AND NOT EXISTS (SELECT 1 FROM exam_usage u WHERE u.question_id = qs.id)
                  GROUP BY qs.chapter_id, qs.qtype, qs.bloom""", fetch=True):
        unused.setdefault(r["chapter_id"], {}).setdefault(r["qtype"], {b: 0 for b in BLOOM})[r["bloom"]] = r["n"]
    by_type = {t: 0 for t in QTYPES}
    for r in q("SELECT qtype, COUNT(*) AS n FROM questions GROUP BY qtype", fetch=True):
        by_type[r["qtype"]] = r["n"]
    rows = q(f"""SELECT ch.id, ch.code, ch.title, t.id AS textbook_id, t.name AS textbook,
                 COALESCE({PRIMARY_COURSE.format(col="co.id")}, 0) AS course_id,
                 (SELECT GROUP_CONCAT(name, ', ') FROM (SELECT co.name FROM course_textbooks ct
                    JOIN courses co ON co.id = ct.course_id WHERE ct.textbook_id = t.id ORDER BY co.id)) AS course,
                 ({CHAPTER_LABEL}) AS label
                 FROM chapters ch JOIN textbooks t ON t.id=ch.textbook_id
                 ORDER BY t.name, ch.code""", fetch=True)
    chapters = []
    for r in rows:
        r["counts"] = counts.get(r["id"], {b: 0 for b in BLOOM})
        r["counts_by_qtype"] = by_qt.get(r["id"], {})
        r["total"] = sum(r["counts"].values())
        chapters.append(r)
    return {"chapters": chapters, "blooms": BLOOM, "by_type": by_type, "unused": unused,
            "total": sum(by_type.values()), "tree": tree}


def build_items(rows):
    """Các câu (đã có mã, lịch sử dùng, kết quả thi, mã học phần/giáo trình, điểm) theo định dạng sao lưu v3."""
    rows = [row_to_api(r) for r in rows]
    usage, results = {}, {}
    for u in q("""SELECT qs.code AS code, u.exam_code, u.exam_name, u.used_at FROM exam_usage u
                  JOIN questions qs ON qs.id = u.question_id""", fetch=True):
        usage.setdefault(u["code"], []).append({k: u[k] for k in ("exam_code", "exam_name", "used_at")})
    for r in q("""SELECT qs.code AS code, r.exam_code, r.correct_pct AS pct, r.recorded_at FROM question_results r
                  JOIN questions qs ON qs.id = r.question_id""", fetch=True):
        results.setdefault(r["code"], []).append({k: r[k] for k in ("exam_code", "pct", "recorded_at")})
    courses = {c["name"]: c for c in q("SELECT name, code, exam_meta FROM courses", fetch=True)}
    items = []
    for r in rows:
        c = courses.get(r["course"], {})
        all_courses = [{"name": n, "code": courses.get(n, {}).get("code", ""),
                        "exam_meta": json.loads(courses.get(n, {}).get("exam_meta") or "{}")}
                       for n in (r.get("courses_all") or "").split(", ") if n]
        items.append({"code": r["code"], "course": r["course"], "course_code": c.get("code", ""),
                      "course_exam_meta": json.loads(c.get("exam_meta") or "{}"), "courses": all_courses,
                      "textbook": r["textbook"], "textbook_code": r.get("textbook_code", ""),
                      "chapter_code": r["chapter_code"], "chapter_title": r["chapter_title"],
                      "bloom": r["bloom"], "qtype": r["qtype"], "content": r["content"], "options": r["options"],
                      "answer": r["answer"], "solution": r["solution"], "passage": r["passage"],
                      "points": r.get("points"), "clo": r.get("clo", ""), "tags": r.get("tags", []),
                      "media": r.get("media", []), "status": r.get("status", "da_duyet"),
                      "passage_media": r.get("passage_media", []),
                      "media_files": {m["file"]: base64.b64encode(open(os.path.join(MEDIA_DIR, m["file"]), "rb").read()).decode()
                                      for m in r.get("media", []) + r.get("passage_media", [])
                                      if os.path.isfile(os.path.join(MEDIA_DIR, m["file"]))},
                      "used": usage.get(r["code"], []), "results": results.get(r["code"], [])})
    return items


def backup_payload():
    """Bản sao lưu đầy đủ (phiên bản 3): câu hỏi kèm mã, mã học phần/giáo trình, thông tin đề,
    lịch sử dùng, kết quả thi, cấu trúc đề đã lưu. Không gồm file .zip đã xuất và thùng rác."""
    items = build_items(q(QUESTION_SQL + " ORDER BY qs.id", fetch=True))
    configs = [{"name": c["name"], "config": json.loads(c["config"] or "{}"), "updated_at": c["updated_at"]}
               for c in q("SELECT name, config, updated_at FROM exam_configs", fetch=True)]
    return {"app": "QBank", "version": 4, "questions": items, "configs": configs}


BACKUP_DIR = os.path.join(DATA_DIR, "data", "backup")
BACKUP_KEEP = 7


def auto_backup():
    """Ghi bản sao lưu JSON của hôm nay (nếu chưa có) và chỉ giữ BACKUP_KEEP bản gần nhất."""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    today = datetime.now().strftime("%Y-%m-%d")
    path = os.path.join(BACKUP_DIR, f"qbank_{today}.json")
    if not os.path.exists(path):
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(backup_payload(), f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)
    files = sorted(f for f in os.listdir(BACKUP_DIR) if re.fullmatch(r"qbank_\d{4}-\d{2}-\d{2}\.json", f))
    for old in files[:-BACKUP_KEEP]:
        os.remove(os.path.join(BACKUP_DIR, old))
    return path


def backup_loop():
    """Kiểm tra mỗi giờ; sang ngày mới thì sao lưu. Lỗi chỉ in ra, không làm dừng server."""
    import threading
    def run():
        while True:
            try:
                auto_backup()
            except Exception as e:
                print("Sao lưu tự động lỗi:", e)
            threading.Event().wait(3600)
    threading.Thread(target=run, daemon=True).start()


# ---------------------------------------------------------------- câu gần giống
SIMILAR_THRESHOLD = 0.9


def similar_pairs(threshold=SIMILAR_THRESHOLD):
    """Các cặp câu cùng dạng có nội dung gần giống (không tính cặp trùng y hệt đã có mục riêng).
    Ứng viên lấy từ chỉ mục 3-ký-tự (chỉ cặp chung nhiều cụm 3 ký tự mới được so chi tiết),
    thay vì so mọi cặp; mỗi cụm quá phổ biến (> 300 câu) bị bỏ qua để tránh bùng nổ tổ hợp."""
    from difflib import SequenceMatcher
    from collections import defaultdict
    rows = q(QUESTION_SQL + " ORDER BY qs.id", fetch=True)
    keys = {r["id"]: dup_key(r["content"], json.loads(r["options"])) for r in rows}
    by_id = {r["id"]: r for r in rows}
    index = defaultdict(set)
    for r in rows:
        k = keys[r["id"]]
        if len(k) < 3:
            continue
        for g in {k[i:i + 3] for i in range(len(k) - 2)}:
            index[(r["qtype"], g)].add(r["id"])
    shared = defaultdict(int)
    for ids in index.values():
        if 2 <= len(ids) <= 300:
            ordered = sorted(ids)
            for i, a in enumerate(ordered):
                for b in ordered[i + 1:]:
                    shared[(a, b)] += 1
    pairs = []
    for (ia, ib), n in shared.items():
        a, b = by_id[ia], by_id[ib]
        ka, kb = keys[ia], keys[ib]
        if n < 3 or ka == kb or not ka or not kb:
            continue
        sm = SequenceMatcher(None, ka, kb, autojunk=False)
        if sm.real_quick_ratio() < threshold or sm.quick_ratio() < threshold:
            continue
        ratio = sm.ratio()
        if ratio >= threshold:
            pairs.append((ratio, a, b))
    pairs.sort(key=lambda t: -t[0])
    return [{"ratio": round(r, 3),
             "items": [{"id": a["id"], "code": a["code"], "chapter": a["chapter_label"],
                        "content": a["content"][:160]},
                       {"id": b["id"], "code": b["code"], "chapter": b["chapter_label"],
                        "content": b["content"][:160]}]} for r, a, b in pairs]


def bulk_update(ids, chapter_id=None, bloom=None):
    """Đổi chương và/hoặc mức Bloom cho nhiều câu. Kiểm tra trùng trước khi đổi; lỗi thì không đổi câu nào."""
    ids = [int(i) for i in ids]
    if not ids:
        raise ValueError("Chưa chọn câu nào")
    if bloom is not None and int(bloom) not in BLOOM:
        raise ValueError("Mức Bloom không hợp lệ")
    if chapter_id is not None and not q("SELECT id FROM chapters WHERE id=?", (int(chapter_id),), fetch=True):
        raise ValueError("Chương đích không tồn tại")
    rows = q(f"SELECT id, content, options, passage FROM questions WHERE id IN ({','.join('?' * len(ids))})", ids, fetch=True)
    if len(rows) != len(set(ids)):
        raise ValueError("Có câu không còn trong ngân hàng, hãy tải lại trang")
    if chapter_id is not None:
        index, moving = dup_index(), set(ids)
        for r in rows:
            key = dup_key(r["content"], json.loads(r["options"]))
            clash = [h for h in index.get(key, []) if h[1] == int(chapter_id) and h[0] not in moving]
            if clash:
                raise ValueError(f"Câu {code_of(r['id'])} sẽ trùng với câu {code_of(clash[0][0])} trong chương đích; không đổi câu nào")
    new_pids = {}
    if chapter_id is not None:  # đoạn văn thuộc chương: chuyển câu thì lấy đoạn của chương đích
        for r in rows:
            new_pids[r["id"]] = resolve_passage(int(chapter_id), r["passage"])
    with closing(sqlite3.connect(DB_PATH)) as conn:
        for i in ids:
            if chapter_id is not None:
                conn.execute("UPDATE questions SET chapter_id=?, passage_id=? WHERE id=?",
                             (int(chapter_id), new_pids[i], i))
            if bloom is not None:
                conn.execute("UPDATE questions SET bloom=? WHERE id=?", (int(bloom), i))
        conn.commit()
    return {"updated": len(ids)}


def trash_questions(ids):
    """Chuyển câu vào thùng rác (không xóa hẳn): lưu đủ dữ liệu, kể cả lịch sử dùng và kết quả thi.
    Có id không tồn tại thì không chuyển câu nào."""
    ids = sorted({int(i) for i in ids})
    if not ids:
        raise ValueError("Chưa chọn câu nào")
    ph = ",".join("?" * len(ids))
    rows = q(QUESTION_SQL + f" WHERE qs.id IN ({ph}) ORDER BY qs.id", ids, fetch=True)
    if len(rows) != len(ids):
        raise ValueError("Có câu không còn trong ngân hàng, hãy tải lại trang")
    items = build_items(rows)
    now = datetime.now().isoformat(timespec="seconds")
    trash_ids = []
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        for it in items:
            cur = conn.execute("INSERT INTO trash(deleted_at, code, item) VALUES(?,?,?)",
                               (now, it["code"], json.dumps(it, ensure_ascii=False)))
            trash_ids.append(cur.lastrowid)
        conn.execute(f"DELETE FROM exam_usage WHERE question_id IN ({ph})", ids)
        conn.execute(f"DELETE FROM question_results WHERE question_id IN ({ph})", ids)
        conn.execute(f"DELETE FROM questions WHERE id IN ({ph})", ids)
        conn.commit()
    return {"deleted": len(ids), "trash_ids": trash_ids, "items": items}


def list_trash():
    out = []
    for r in q("SELECT id, deleted_at, code, item FROM trash ORDER BY id DESC", fetch=True):
        it = json.loads(r["item"])
        out.append({"id": r["id"], "deleted_at": r["deleted_at"], "code": r["code"],
                    "chapter": f"{it.get('course', '')} / {it.get('textbook', '')} / {it.get('chapter_code', '')}",
                    "qtype": it.get("qtype", ""), "content": (it.get("content") or "")[:160]})
    return out


def restore_from_trash(ids):
    """Khôi phục câu từ thùng rác. Câu nào không khôi phục được (ví dụ trùng nội dung) thì giữ nguyên trong thùng rác."""
    index = dup_index()
    restored, errors = 0, []
    for tid in ids:
        rows = q("SELECT id, item FROM trash WHERE id=?", (int(tid),), fetch=True)
        if not rows:
            continue
        item = json.loads(rows[0]["item"])
        try:
            restore_item(item, index)
        except Exception as e:
            errors.append(f"{item.get('code', '')}: {e}")
            continue
        q("DELETE FROM trash WHERE id=?", (int(tid),))
        restored += 1
    return {"restored": restored, "errors": errors}


def restore_item(item, index, allow_dup=False):
    """Khôi phục một câu từ file sao lưu; tự tạo học phần/giáo trình/chương nếu chưa có.
    Câu trùng trong cùng chương bị bỏ qua (báo lỗi để người dùng biết)."""
    item = dict(item)
    item.setdefault("status", "da_duyet")  # bản sao lưu cũ không có trạng thái: giữ như đang dùng
    course_list = [str(c.get("name", "")).strip() for c in item.get("courses") or [] if str(c.get("name", "")).strip()]
    course_list = course_list or [str(item.get("course", DEFAULT_COURSE)).strip() or DEFAULT_COURSE]
    chapter_id = ensure_chapter(course_list,
                                str(item.get("textbook", DEFAULT_COURSE)).strip() or DEFAULT_COURSE,
                                str(item.get("chapter_code", "Chung")).strip() or "Chung",
                                str(item.get("chapter_title", "")))
    for fname, b64 in (item.get("media_files") or {}).items():  # khôi phục file đính kèm từ bản sao lưu
        if MEDIA_NAME.fullmatch(str(fname)) and not os.path.exists(os.path.join(MEDIA_DIR, fname)):
            os.makedirs(MEDIA_DIR, exist_ok=True)
            with open(os.path.join(MEDIA_DIR, fname), "wb") as f:
                f.write(base64.b64decode(b64))
    payload = clean_question(dict(item, chapter_id=chapter_id))
    same, _, key = check_duplicate(index, payload)
    if same and not allow_dup:
        raise ValueError(f"trùng câu {code_of(same[0][0])} trong cùng chương (bỏ qua)")
    new_id = insert_q(payload, keep_code=str(item.get("code") or "").strip() or None)
    index.setdefault(key, []).append((new_id, chapter_id, payload["content"]))
    pid = q("SELECT passage_id FROM questions WHERE id=?", (new_id,), True)[0]["passage_id"]
    if pid and item.get("passage_media") and not passage_media(pid):  # file của đoạn: chỉ điền khi đoạn chưa có
        try:
            set_passage_media(pid, item["passage_media"])
        except ValueError:
            pass
    # mã học phần/giáo trình, thông tin đề (chỉ điền khi còn trống)
    saved_courses = item.get("courses") or [{"name": course_list[0], "code": item.get("course_code", ""),
                                             "exam_meta": item.get("course_exam_meta") or {}}]
    for c in saved_courses:  # mã và thông tin đề của từng học phần (chỉ điền khi còn trống)
        q("UPDATE courses SET code=? WHERE code='' AND name=?", (str(c.get("code", "")), str(c.get("name", ""))))
        q("UPDATE courses SET exam_meta=? WHERE exam_meta='{}' AND name=?",
          (json.dumps(c.get("exam_meta") or {}, ensure_ascii=False), str(c.get("name", ""))))
    q("UPDATE textbooks SET code=? WHERE code='' AND id=(SELECT textbook_id FROM chapters WHERE id=?)",
      (str(item.get("textbook_code", "")), chapter_id))
    # lịch sử dùng và kết quả thi của câu
    record_usage([(new_id, str(u.get("exam_code", "")), str(u.get("exam_name", "")), str(u.get("used_at", "")))
                  for u in item.get("used") or []])
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.executemany("INSERT INTO question_results(question_id, exam_code, correct_pct, recorded_at) VALUES(?,?,?,?)",
                         [(new_id, str(r.get("exam_code", "")), float(r.get("pct", 0)), str(r.get("recorded_at", "")))
                          for r in item.get("results") or []])
        conn.commit()


# ---------------------------------------------------------------- HTTP
class NeedConfirm(ValueError):
    """Thao tác cần người dùng xác nhận (trả về HTTP 409 kèm confirm=true); gửi lại với confirm_passage=true để tiếp tục."""


class UpdateError(Exception):
    """Lỗi khi kiểm tra hoặc tải bản cập nhật từ GitHub (trả về HTTP 502)."""


def similar_passage(chapter_id, text, threshold=SIMILAR_THRESHOLD):
    """Đoạn văn khác trong chương có nội dung gần giống nhất (>= threshold). Trả về (mã đoạn, tỷ lệ) hoặc None."""
    ka = dup_key(text, [])
    best = None
    for r in q("SELECT code, content FROM passages WHERE chapter_id=?", (int(chapter_id),), True):
        kb = dup_key(r["content"], [])
        if not ka or not kb:
            continue
        sm = SequenceMatcher(None, ka, kb, autojunk=False)
        if sm.real_quick_ratio() < threshold or sm.quick_ratio() < threshold:
            continue
        ratio = sm.ratio()
        if ratio >= threshold and (best is None or ratio > best[1]):
            best = (r["code"], ratio)
    return best


def check_passage(chapter_id, text, confirmed=False):
    """Đoạn văn mới gần giống đoạn có sẵn trong chương: hỏi xác nhận trước khi tạo, tránh đoạn trùng do sửa vài chữ."""
    text = (text or "").strip()
    if not text or confirmed:
        return
    if q("SELECT 1 FROM passages WHERE chapter_id=? AND content=?", (int(chapter_id), text), True):
        return  # trùng đúng từng chữ: dùng chung đoạn đó, không cần hỏi
    sim = similar_passage(chapter_id, text)
    if sim:
        raise NeedConfirm(f"Đoạn văn này gần giống đoạn {sim[0]} ({round(sim[1] * 100)}%) trong cùng chương. "
                          f"Nếu là cùng một đoạn, hãy chọn đoạn đó ở danh sách; nếu vẫn muốn tạo đoạn mới, xác nhận lại.")


class Handler(BaseHTTPRequestHandler):
    server_version = "QBank/2.0"

    def log_message(self, fmt, *args):  # im lặng cho gọn
        pass

    def _send(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _raw(self):
        return self.rfile.read(int(self.headers.get("Content-Length") or 0))

    def _json_body(self):
        return json.loads(self._raw() or b"{}")

    def _guard(self, fn, write=False):
        try:
            if write:
                self._local_only()
            fn()
        except NeedConfirm as e:
            self._json(409, {"error": str(e), "confirm": True})
        except ValueError as e:
            self._json(400, {"error": str(e)})
        except sqlite3.IntegrityError as e:
            msg = "Đã tồn tại (trùng tên/mã trong cùng cấp)" if "UNIQUE" in str(e) else "Dữ liệu bị ràng buộc"
            self._json(400, {"error": msg})
        except UpdateError as e:
            self._json(502, {"error": str(e)})
        except Exception as e:  # lỗi bất ngờ: trả về thông báo, không làm sập server
            self._json(500, {"error": f"{type(e).__name__}: {e}"})

    def _local_only(self):
        """Chặn ghi dữ liệu từ trang web khác (CSRF): chỉ nhận Host và Origin của chính server này."""
        if self.headers.get("Host", "") not in (f"127.0.0.1:{PORT}", f"localhost:{PORT}"):
            raise ValueError("Yêu cầu bị chặn: sai địa chỉ máy chủ")
        origin = self.headers.get("Origin")
        if origin and origin not in (f"http://127.0.0.1:{PORT}", f"http://localhost:{PORT}"):
            raise ValueError("Yêu cầu từ trang khác đã bị chặn")

    def _fix_path(self):
        # máy chủ đọc dòng yêu cầu theo latin-1; URL có chữ UTF-8 thô (không mã hóa) được đọc lại thành UTF-8
        self.path = self.path.encode("latin-1", "replace").decode("utf-8", "replace")

    def do_GET(self):
        self._fix_path()
        mm = re.fullmatch(r"/media/([^/]+)", urlparse(self.path).path)
        if mm:  # file đính kèm: chỉ phục vụ tên đúng định dạng trong thư mục media
            if not MEDIA_NAME.fullmatch(mm.group(1)) or not os.path.isfile(os.path.join(MEDIA_DIR, mm.group(1))):
                return self._json(404, {"error": "không tìm thấy file"})
            ext = mm.group(1).rsplit(".", 1)[1]
            ctype = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "gif": "image/gif", "webp": "image/webp",
                     "mp3": "audio/mpeg", "wav": "audio/wav", "ogg": "audio/ogg", "m4a": "audio/mp4"}[ext]
            with open(os.path.join(MEDIA_DIR, mm.group(1)), "rb") as f:
                return self._send(200, f.read(), ctype)
        self._guard(self._get)

    def do_POST(self):
        self._fix_path()
        self._guard(self._post, write=True)

    def do_PUT(self):
        self._fix_path()
        self._guard(self._put, write=True)

    def do_DELETE(self):
        self._fix_path()
        self._guard(self._delete, write=True)

    # --- GET
    def _get(self):
        u = urlparse(self.path)
        p, qs = u.path, parse_qs(u.query)
        arg = lambda k: qs.get(k, [""])[0].strip()
        if p in ("/", "/index.html"):
            with open(os.path.join(STATIC, "index.html"), "rb") as f:
                return self._send(200, f.read(), "text/html; charset=utf-8")
        if p == "/favicon.ico":
            return self._send(204, b"", "image/x-icon")
        if p == "/api/version":
            return self._json(200, {"version": APP_VERSION})
        if p == "/api/update":  # so sánh với GitHub Releases, không ghi gì
            return self._json(200, check_update())
        if p == "/api/questions":
            sql, args = QUESTION_SQL + " WHERE 1=1", []
            if arg("chapter_id"):
                sql += " AND qs.chapter_id=?"; args.append(int(arg("chapter_id")))
            if arg("textbook_id"):
                sql += " AND t.id=?"; args.append(int(arg("textbook_id")))
            if arg("course_id"):  # giáo trình dùng chung: lọc theo mọi học phần đang nối với giáo trình
                sql += " AND EXISTS (SELECT 1 FROM course_textbooks ct WHERE ct.textbook_id = t.id AND ct.course_id = ?)"
                args.append(int(arg("course_id")))
            if arg("bloom"):
                sql += " AND qs.bloom=?"; args.append(int(arg("bloom")))
            if arg("qtype"):
                sql += " AND qs.qtype=?"; args.append(arg("qtype"))
            if arg("status") in STATUSES:
                sql += " AND qs.status=?"; args.append(arg("status"))
            if arg("tag"):  # thẻ lưu dạng ",a,b," nên khớp ",thẻ," là chính xác
                esc_tag = arg("tag").replace("!", "!!").replace("%", "!%").replace("_", "!_")
                sql += " AND qs.tags LIKE ? ESCAPE '!'"; args.append(f"%,{esc_tag},%")
            kw = arg("q")
            if kw:  # ký tự % _ trong từ khóa được coi là chữ thường (escape bằng !)
                cond = " OR ".join(f"{c} LIKE ? ESCAPE '!'" for c in ("qs.content", "qs.solution", "qs.options", "qs.answer", "qs.code"))
                sql += f" AND ({cond})"
                pat = "%" + kw.replace("!", "!!").replace("%", "!%").replace("_", "!_") + "%"
                args += [pat] * 5
            sql += " ORDER BY t.name, ch.code, qs.bloom, qs.id"
            if "page" in qs:  # bảng ngân hàng: trả từng trang và tổng số câu khớp
                per = min(max(int(arg("per") or 50), 1), 50)  # tối đa 50 câu mỗi trang
                page_no = max(int(arg("page") or 0), 0)
                total = q(f"SELECT COUNT(*) AS n FROM ({sql})", args, fetch=True)[0]["n"]
                rows = q(sql + " LIMIT ? OFFSET ?", args + [per, page_no * per], fetch=True)
                return self._json(200, {"items": [row_to_api(r) for r in rows], "total": total})
            return self._json(200, [row_to_api(r) for r in q(sql, args, fetch=True)])
        one = re.fullmatch(r"/api/questions/(\d+)", p)
        if one:  # một câu, để mở form Sửa không phải tải cả ngân hàng
            rows = q(QUESTION_SQL + " WHERE qs.id=?", (int(one.group(1)),), fetch=True)
            if not rows:
                return self._json(404, {"error": "Câu này không còn trong ngân hàng"})
            return self._json(200, row_to_api(rows[0]))
        if p == "/api/tts/status":  # textts có đang chạy không, engine nào dùng được
            try:
                h = textts_call("/api/v1/health", timeout=4)
                return self._json(200, {"ok": True, "url": TEXTTS_URL, "voicevox": h.get("voicevox_connected", False),
                                        "aivis": h.get("aivis_connected", False)})
            except ValueError as e:
                return self._json(200, {"ok": False, "url": TEXTTS_URL, "error": str(e)})
        if p == "/api/tags":  # thẻ đang dùng và số câu của từng thẻ
            count = {}
            for r in q("SELECT tags FROM questions WHERE tags <> ''", fetch=True):
                for t in tags_list(r["tags"]):
                    count[t] = count.get(t, 0) + 1
            return self._json(200, [{"tag": t, "count": n} for t, n in sorted(count.items(), key=lambda x: x[0].lower())])
        if p == "/api/trash":
            return self._json(200, list_trash())
        if p == "/api/tree":
            return self._json(200, get_tree())
        if p == "/api/similar":  # cặp câu gần giống (>= 90%), để rà soát
            return self._json(200, similar_pairs())
        if p == "/api/duplicates":  # các nhóm câu có nội dung giống nhau (để rà soát)
            groups = []
            for v in dup_index().values():
                if len(v) < 2:
                    continue
                items = q(QUESTION_SQL + " WHERE qs.id IN (%s) ORDER BY qs.id" % ",".join("?" * len(v)),
                          [h[0] for h in v], fetch=True)
                groups.append({"content": items[0]["content"][:120],
                               "items": [{"id": r["id"], "code": r["code"], "chapter": r["chapter_label"],
                                          "used_count": r["used_count"]} for r in items]})
            return self._json(200, groups)
        if p == "/api/passages":
            return self._json(200, list_passages())
        if p == "/api/difficulty":
            return self._json(200, difficulty())
        if p == "/api/stats":
            return self._json(200, stats())
        if p == "/api/backup":
            body = json.dumps(backup_payload(), ensure_ascii=False, indent=1).encode()
            return self._send(200, body, "application/json; charset=utf-8",
                              {"Content-Disposition": 'attachment; filename="qbank_backup.json"'})
        if p == "/api/import-template":
            return self._send(200, make_import_template(),
                              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                              {"Content-Disposition": 'attachment; filename="mau_nhap_cau_hoi.docx"'})
        if p == "/api/templates":
            names = sorted(f for f in os.listdir(TPL_DIR) if f.lower().endswith(".docx"))
            return self._json(200, names)
        if p == "/api/exports":  # lịch sử các lần xuất (không kèm nội dung file)
            return self._json(200, q("""SELECT id, created_at, name, codes, length(data) AS size
                                        FROM exam_exports ORDER BY id DESC LIMIT 200""", fetch=True))
        fm = re.fullmatch(r"/api/exports/(\d+)/file", p)
        if fm:
            rows = q("SELECT codes, data, config, created_at FROM exam_exports WHERE id=?", (int(fm.group(1)),), fetch=True)
            if not rows:
                return self._json(404, {"error": "không tìm thấy bản xuất"})
            row = rows[0]
            meta = (json.loads(row["config"] or "{}").get("meta") or {})
            codes = [c.strip() for c in (row["codes"] or "").split(",") if c.strip()]
            fname = export_filename(meta.get("ma_hp"), codes, row["created_at"])
            return self._send(200, row["data"], "application/zip",
                              {"Content-Disposition": attachment_header(fname)})
        if p == "/api/configs":  # cấu trúc đề đã lưu
            rows = q("SELECT id, name, config, updated_at FROM exam_configs ORDER BY name", fetch=True)
            for r in rows:
                r["config"] = json.loads(r["config"] or "{}")
            return self._json(200, rows)
        self._json(404, {"error": "không tìm thấy"})

    # --- POST
    def _post(self):
        u = urlparse(self.path)
        p = u.path
        if p == "/api/update/install":
            version = install_update()
            self._json(200, {"ok": True, "version": version})
            threading.Timer(1.0, lambda: os._exit(0)).start()  # tắt bản cũ để bản mới nhận cổng
            return
        if p == "/api/questions":
            body = self._json_body()
            payload = clean_question(body)
            check_passage(payload["chapter_id"], payload.get("passage"), bool(body.get("confirm_passage")))
            same, other, _ = check_duplicate(dup_index(), payload)
            if same:  # đã có câu y hệt trong cùng chương: không lưu, báo người dùng
                raise ValueError(f"Câu này đã có trong cùng chương ({code_of(same[0][0])}). Hãy sửa nội dung nếu muốn lưu thêm.")
            new_id = insert_q(payload)
            save_passage_media(new_id, body)
            warnings = [f"Trùng nội dung với câu {code_of(h[0])} ở chương khác" for h in other[:3]]
            return self._json(201, {"ok": True, "id": new_id, "code": code_of(new_id), "warnings": warnings})
        if p == "/api/chapters/import":  # nhập hàng loạt chương từ CSV/xlsx (cột Mã, Tên chương)
            textbook_id = int(parse_qs(u.query).get("textbook_id", [""])[0] or 0)
            if not textbook_id:
                raise ValueError("Chọn giáo trình trước khi nhập chương")
            dry = parse_qs(u.query).get("dry_run", ["0"])[0] == "1"
            return self._json(200, import_chapters(read_table_text(self._raw()), textbook_id, dry))
        if p == "/api/import-text":  # CSV hoặc dữ liệu dán từ Excel
            textbook_id = int(parse_qs(u.query).get("textbook_id", [""])[0] or 0)
            if not textbook_id:
                raise ValueError("Chọn giáo trình trước khi nhập")
            dry = parse_qs(u.query).get("dry_run", ["0"])[0] == "1"
            rows = read_table_text(self._raw())
            rows = [r for r in rows if any(str(c).strip() for c in r)]
            if not rows:
                raise ValueError("Không có dữ liệu để nhập")
            return self._json(200, parse_rows([rows], textbook_id, dry_run=dry, docx=False))
        if p == "/api/tts":  # tạo file nghe bằng textts từ kịch bản; trả về file đính kèm như /api/media
            b = self._json_body()
            return self._json(201, textts_audio(b.get("script", ""), b.get("engine") or "voicevox", b.get("level")))
        if p == "/api/media":  # tải file đính kèm lên; trả về tên file để gắn vào câu
            raw = self._raw()
            name = parse_qs(u.query).get("name", ["file"])[0]
            ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
            if ext not in MEDIA_EXT:
                raise ValueError("Chỉ nhận ảnh (png, jpg, gif, webp) hoặc âm thanh (mp3, wav, ogg, m4a)")
            if len(raw) > MEDIA_MAX[MEDIA_EXT[ext]]:
                raise ValueError("File quá lớn (ảnh tối đa 5 MB, âm thanh tối đa 15 MB)")
            if not raw:
                raise ValueError("File rỗng")
            os.makedirs(MEDIA_DIR, exist_ok=True)
            fname = f"{os.urandom(16).hex()}.{ext}"
            with open(os.path.join(MEDIA_DIR, fname), "wb") as f:
                f.write(raw)
            return self._json(201, {"file": fname, "kind": MEDIA_EXT[ext], "name": os.path.basename(name)[:120]})
        if p == "/api/export/moodle":  # xuất câu hỏi ra Moodle XML: câu đã chọn, hoặc mọi câu đã duyệt
            ids = self._json_body().get("ids") or None
            xml = moodle_xml(ids).encode("utf-8")
            return self._send(200, xml, "application/xml; charset=utf-8",
                              {"Content-Disposition": 'attachment; filename="qbank_moodle.xml"'})
        if p == "/api/questions/status":  # đổi trạng thái hàng loạt: Nháp, Đã duyệt, Ngưng dùng
            b = self._json_body()
            status = str(b.get("status", ""))
            if status not in STATUSES:
                raise ValueError("Trạng thái không hợp lệ")
            ids = sorted({int(i) for i in b.get("ids", [])})
            if not ids:
                raise ValueError("Chưa chọn câu nào")
            with closing(sqlite3.connect(DB_PATH)) as conn:
                cur = conn.execute(f"UPDATE questions SET status=? WHERE id IN ({','.join('?' * len(ids))})", [status] + ids)
                conn.commit()
                return self._json(200, {"updated": cur.rowcount})
        if p == "/api/item-analysis":  # phân tích câu hỏi từ bảng điểm (file .csv, .xlsx hoặc dữ liệu dán)
            raw = self._raw()
            res = item_analysis([read_table_text(raw)])
            if parse_qs(u.query).get("record", ["0"])[0] == "1":  # ghi tỷ lệ đúng vào kết quả của mã đề
                code = parse_qs(u.query).get("exam_code", [""])[0]
                res["recorded"] = record_results(code, [(it["code"], it["pct"]) for it in res["items"]])
            return self._json(200, res)
        if p == "/api/exam/print":  # bản in xem trước (không ghi lịch sử, không ghi lịch sử dùng)
            return self._send(200, print_html(self._json_body()).encode("utf-8"), "text/html; charset=utf-8")
        if p == "/api/import":
            textbook_id = int(parse_qs(u.query).get("textbook_id", [""])[0] or 0)
            if not textbook_id:
                raise ValueError("Chọn giáo trình trước khi nhập")
            dry = parse_qs(u.query).get("dry_run", ["0"])[0] == "1"
            return self._json(200, parse_import(self._raw(), textbook_id, dry_run=dry))
        if p == "/api/restore":
            data = json.loads(self._raw() or b"{}")
            items = data.get("questions") if isinstance(data, dict) else data
            if not isinstance(items, list):
                raise ValueError("File sao lưu không đúng định dạng")
            added, errors = 0, []
            index = dup_index()
            allow_dup = bool(data.get("allow_dup")) if isinstance(data, dict) else False
            for c in (data.get("configs") or []) if isinstance(data, dict) else []:  # cấu trúc đề đã lưu
                if str(c.get("name", "")).strip() and isinstance(c.get("config"), dict):
                    q("""INSERT INTO exam_configs(name, config, updated_at) VALUES(?, ?, ?)
                         ON CONFLICT(name) DO UPDATE SET config=excluded.config, updated_at=excluded.updated_at""",
                      (str(c["name"]).strip(), json.dumps(c["config"], ensure_ascii=False), str(c.get("updated_at", ""))))
            for n, item in enumerate(items, 1):
                try:
                    restore_item(item, index, allow_dup=allow_dup)
                    added += 1
                except Exception as e:
                    errors.append(f"Câu thứ {n}: {e}")
            return self._json(200, {"added": added, "errors": errors})
        if p == "/api/results":  # nhập tỷ lệ đúng của câu sau kỳ thi
            b = self._json_body()
            rows = [(line.get("code", ""), line.get("pct", "")) for line in b.get("rows", [])]
            return self._json(200, record_results(b.get("exam_code", ""), rows))
        if p == "/api/questions/bulk":  # đổi chương/mức Bloom cho nhiều câu
            b = self._json_body()
            return self._json(200, bulk_update(b.get("ids", []), b.get("chapter_id") or None, b.get("bloom") or None))
        if p == "/api/questions/bulk-delete":  # xóa nhiều câu đã chọn (trả về bản sao lưu để hoàn tác)
            return self._json(200, trash_questions(self._json_body().get("ids", [])))
        if p == "/api/trash/restore":
            return self._json(200, restore_from_trash(self._json_body().get("ids", [])))
        if p == "/api/trash/empty":
            q("DELETE FROM trash")
            return self._json(200, {"ok": True})
        if p == "/api/duplicates/merge":  # giữ một câu, xóa các câu trùng đã chọn (lịch sử chuyển sang câu giữ)
            b = self._json_body()
            keep = int(b.get("keep", 0))
            removes = [int(x) for x in b.get("remove", []) if int(x) != keep]
            return self._json(200, merge_questions(keep, removes))
        if p == "/api/duplicates/auto":  # tự gộp các câu trùng trong cùng chương, giữ câu có id nhỏ nhất
            return self._json(200, auto_merge_same_chapter())
        if p == "/api/export":
            return self._export(self._json_body())
        if p == "/api/exam/preview":  # chọn câu để xem trước (chưa ghi lịch sử)
            sections = self._json_body().get("sections") or []
            return self._json(200, {"sections": preview_exam(sections)})
        if p == "/api/usage/reset":  # reset tình trạng ra đề: câu, chương, hoặc toàn bộ (cần gõ RESET)
            return self._json(200, reset_usage(self._json_body()))
        if p == "/api/exam/swap":  # đổi một câu trong bản xem trước
            b = self._json_body()
            usage = usage_map()
            exclude = {int(i) for i in b.get("exclude", [])}
            return self._json(200, find_replacement(b, exclude, usage))
        if p == "/api/configs":  # lưu hoặc ghi đè cấu trúc đề theo tên
            b = self._json_body()
            name = str(b.get("name", "")).strip()
            if not name:
                raise ValueError("Thiếu tên cấu trúc đề")
            config = b.get("config")
            if not isinstance(config, dict) or not config.get("sections"):
                raise ValueError("Cấu trúc đề không hợp lệ")
            now = datetime.now().isoformat(timespec="seconds")
            q("""INSERT INTO exam_configs(name, config, updated_at) VALUES(?, ?, ?)
                 ON CONFLICT(name) DO UPDATE SET config=excluded.config, updated_at=excluded.updated_at""",
              (name, json.dumps(config, ensure_ascii=False), now))
            return self._json(201, {"ok": True})
        if p == "/api/courses":
            body = self._json_body()
            name = str(body.get("name", "")).strip()
            if not name:
                raise ValueError("Thiếu tên học phần")
            q("INSERT INTO courses(name, code) VALUES(?, ?)", (name, str(body.get("code", "")).strip()))
            return self._json(201, {"ok": True})
        if p == "/api/textbooks":  # giáo trình đã có cùng tên thì gắn vào học phần này (dùng chung), không tạo bản mới
            b = self._json_body()
            name = str(b.get("name", "")).strip()
            if not name or not b.get("course_id"):
                raise ValueError("Thiếu học phần hoặc tên giáo trình")
            code = str(b.get("code", "")).strip()
            q("INSERT OR IGNORE INTO textbooks(name, code) VALUES(?, ?)", (name, code))
            if code:
                q("UPDATE textbooks SET code=? WHERE name=? AND code=''", (code, name))
            tid = q("SELECT id FROM textbooks WHERE name=?", (name,), True)[0]["id"]
            q("INSERT OR IGNORE INTO course_textbooks(course_id, textbook_id) VALUES(?, ?)", (int(b["course_id"]), tid))
            return self._json(201, {"ok": True, "id": tid})
        if p == "/api/course_textbooks":  # gắn giáo trình có sẵn vào thêm một học phần
            b = self._json_body()
            if not b.get("course_id") or not b.get("textbook_id"):
                raise ValueError("Thiếu học phần hoặc giáo trình")
            q("INSERT OR IGNORE INTO course_textbooks(course_id, textbook_id) VALUES(?, ?)",
              (int(b["course_id"]), int(b["textbook_id"])))
            return self._json(201, {"ok": True})
        if p == "/api/chapters":
            b = self._json_body()
            code = str(b.get("code", "")).strip()
            if not code or not b.get("textbook_id"):
                raise ValueError("Thiếu giáo trình hoặc mã chương")
            q("INSERT INTO chapters(textbook_id, code, title) VALUES(?, ?, ?)",
              (int(b["textbook_id"]), code, str(b.get("title", "")).strip()))
            return self._json(201, {"ok": True})
        self._json(404, {"error": "không tìm thấy"})

    # --- PUT
    def _put(self):
        path = urlparse(self.path).path
        rn = re.fullmatch(r"/api/(courses|textbooks|chapters)/(\d+)", path)
        if rn:  # sửa tên/mã học phần, giáo trình, chương
            b = self._json_body()
            kind, rid = rn.group(1), int(rn.group(2))
            if kind == "courses":
                name = str(b.get("name", "")).strip()
                if not name:
                    raise ValueError("Thiếu tên học phần")
                q("UPDATE courses SET name=?, code=? WHERE id=?", (name, str(b.get("code", "")).strip(), rid))
            elif kind == "textbooks":
                name = str(b.get("name", "")).strip()
                if not name:
                    raise ValueError("Thiếu tên giáo trình")
                try:
                    q("UPDATE textbooks SET name=?, code=? WHERE id=?", (name, str(b.get("code", "")).strip(), rid))
                except sqlite3.IntegrityError:
                    raise ValueError("Đã có giáo trình cùng tên")
            else:
                code = str(b.get("code", "")).strip()
                if not code:
                    raise ValueError("Thiếu mã chương")
                q("UPDATE chapters SET code=?, title=? WHERE id=?", (code, str(b.get("title", "")).strip(), rid))
            return self._json(200, {"ok": True})
        pm = re.fullmatch(r"/api/passages/(\d+)", path)
        if pm:
            b = self._json_body()
            if "content" in b:
                update_passage(int(pm.group(1)), b.get("content", ""))
            if isinstance(b.get("media"), list):
                if not q("SELECT 1 FROM passages WHERE id=?", (int(pm.group(1)),), True):
                    raise ValueError("Đoạn văn không tồn tại")
                set_passage_media(int(pm.group(1)), b["media"])
            return self._json(200, {"ok": True})
        cm = re.fullmatch(r"/api/courses/(\d+)/exam-meta", urlparse(self.path).path)
        if cm:  # thông tin đề thi gắn với học phần, tự điền khi xuất đề
            body = self._json_body()
            meta = {k: str(body.get(k, "")).strip() for k in EXAM_META_KEYS}
            q("UPDATE courses SET exam_meta=? WHERE id=?",
              (json.dumps(meta, ensure_ascii=False), int(cm.group(1))))
            return self._json(200, {"ok": True})
        m = re.fullmatch(r"/api/questions/(\d+)", urlparse(self.path).path)
        if not m:
            return self._json(404, {"error": "không tìm thấy"})
        body = self._json_body()
        payload = clean_question(body)
        check_passage(payload["chapter_id"], payload["passage"], bool(body.get("confirm_passage")))
        payload["id"] = int(m.group(1))
        same, _, _ = check_duplicate(dup_index(), payload)
        same = [h for h in same if h[0] != payload["id"]]  # không tính chính câu đang sửa
        if same:
            raise ValueError(f"Câu này đã có trong cùng chương ({code_of(same[0][0])}). Hãy sửa nội dung nếu muốn lưu thêm.")
        payload["passage_id"] = resolve_passage(payload["chapter_id"], payload["passage"])
        with closing(sqlite3.connect(DB_PATH)) as conn:
            conn.execute("PRAGMA foreign_keys = ON")
            cur = conn.execute("""UPDATE questions SET chapter_id=:chapter_id, bloom=:bloom, qtype=:qtype, content=:content,
                 options=:options, answer=:answer, solution=:solution, passage=:passage, passage_id=:passage_id, rubric=:rubric,
                 points=:points, clo=:clo, tags=:tags, media=:media, status=:status, word_limit=:word_limit WHERE id=:id""", payload)
            conn.commit()
            updated = cur.rowcount
        if not updated:
            raise ValueError("Câu này không còn trong ngân hàng, hãy tải lại trang")
        save_passage_media(payload["id"], body)
        self._json(200, {"ok": True, "code": code_of(payload["id"])})

    # --- DELETE: chỉ xóa khi không còn dữ liệu bên trong (khóa ngoài RESTRICT)
    def _delete(self):
        path = urlparse(self.path).path
        psg = re.fullmatch(r"/api/passages/(\d+)", path)
        if psg:
            delete_passage(int(psg.group(1)))
            return self._json(200, {"ok": True})
        cfg = re.fullmatch(r"/api/configs/(\d+)", path)
        if cfg:
            q("DELETE FROM exam_configs WHERE id=?", (int(cfg.group(1)),))
            return self._json(200, {"ok": True})
        ln = re.fullmatch(r"/api/course_textbooks/(\d+)/(\d+)", path)
        if ln:  # gỡ giáo trình khỏi học phần; giáo trình còn chương thì phải còn ít nhất một học phần
            cid, tid = int(ln.group(1)), int(ln.group(2))
            left = q("SELECT COUNT(*) AS n FROM course_textbooks WHERE textbook_id=? AND course_id<>?", (tid, cid), True)[0]["n"]
            has_chapters = q("SELECT COUNT(*) AS n FROM chapters WHERE textbook_id=?", (tid,), True)[0]["n"]
            if not left and has_chapters:
                raise ValueError("Giáo trình còn chương, phải còn ít nhất một học phần (hoặc xóa giáo trình trước)")
            q("DELETE FROM course_textbooks WHERE course_id=? AND textbook_id=?", (cid, tid))
            return self._json(200, {"ok": True})
        tr = re.fullmatch(r"/api/trash/(\d+)", path)
        if tr:
            q("DELETE FROM trash WHERE id=?", (int(tr.group(1)),))
            return self._json(200, {"ok": True})
        ex = re.fullmatch(r"/api/exports/(\d+)", path)
        if ex:
            q("DELETE FROM exam_exports WHERE id=?", (int(ex.group(1)),))
            return self._json(200, {"ok": True})
        m = re.fullmatch(r"/api/(questions|courses|textbooks|chapters)/(\d+)", path)
        if not m:
            return self._json(404, {"error": "không tìm thấy"})
        if m.group(1) == "chapters":  # đoạn văn không còn câu nào trong chương thì xóa theo
            q("DELETE FROM passages WHERE chapter_id=? AND id NOT IN (SELECT passage_id FROM questions WHERE passage_id IS NOT NULL)",
              (int(m.group(2)),))
        table = {"questions": "questions", "courses": "courses",
                 "textbooks": "textbooks", "chapters": "chapters"}[m.group(1)]
        if table == "questions":  # câu hỏi vào thùng rác thay vì xóa hẳn
            self._json(200, trash_questions([int(m.group(2))]))
            return
        if table == "textbooks":  # gỡ mọi liên kết học phần trước; giáo trình còn chương thì không xóa
            if q("SELECT COUNT(*) AS n FROM chapters WHERE textbook_id=?", (int(m.group(2)),), True)[0]["n"]:
                raise ValueError("Còn dữ liệu bên trong (câu hỏi, giáo trình hoặc chương), không thể xóa")
            q("DELETE FROM course_textbooks WHERE textbook_id=?", (int(m.group(2)),))
        try:
            q(f"DELETE FROM {table} WHERE id=?", (int(m.group(2)),))
        except sqlite3.IntegrityError:
            raise ValueError("Còn dữ liệu bên trong (câu hỏi, giáo trình hoặc chương), không thể xóa")
        self._json(200, {"ok": True})

    # --- xuất đề
    def _export(self, body):
        tpl = os.path.basename(body.get("template", ""))
        tpl_path = os.path.join(TPL_DIR, tpl)
        if not tpl or not os.path.isfile(tpl_path):
            raise ValueError("Chưa chọn template hợp lệ")
        variants = max(1, int(body.get("variants", 1)))
        start = int(body.get("start", 1))
        prefix = re.sub(r"[^\w-]", "", str(body.get("prefix", "")))  # tiền tố chỉ gồm chữ, số, -, _
        answer_key = bool(body.get("answer_key", True))
        key_separate = bool(body.get("key_separate", False))
        sections = body.get("sections") or [{"title": "", "qtype": None,
                                             "shuffle": body.get("shuffle", True),
                                             "matrix": body.get("matrix", [])}]
        fg = str(body.get("furigana", "ruby"))
        furigana = ("paren" if fg.startswith("paren") else "ruby") + ("+auto" if fg.endswith("_auto") else "")
        listening = body.get("listening") if isinstance(body.get("listening"), dict) else None
        listening_made = 0
        if furigana.endswith("+auto") and _get_kakasi() is None:
            raise ValueError("Chưa cài pykakasi nên không tự thêm furigana được (pip install pykakasi)")
        key_name = os.path.basename(body.get("key_template") or "phieu_dap_an.docx")
        key_tpl_path = os.path.join(TPL_DIR, key_name)
        usage = usage_map()
        counts = usage_counts()
        # seed: cùng seed và cùng lịch sử dùng thì chọn lại đúng các câu cũ; mỗi lần xuất có seed riêng
        seed = int(body["seed"]) if str(body.get("seed", "")).strip() else random.SystemRandom().randrange(1, 10**9)
        clo_mode = bool(body.get("clo_mode"))
        rng = random.Random(seed)
        taken = set()
        picked = body.get("picked")  # câu đã xem trước và đổi tay, aligned theo sections
        parts = []  # (tiêu đề, danh sách câu gốc, có xáo, là phần nghe)
        for idx, s in enumerate(sections):
            if not section_has_content(s):
                continue  # phần không có số câu thì bỏ qua
            ids = picked[idx] if picked and idx < len(picked) else None
            if ids:
                if len(set(ids)) != len(ids):
                    raise ValueError("Bản xem trước có câu trùng; hãy xem trước lại")
                if not any(s.get("passages") or []):  # ma trận chương: số câu phải khớp
                    want = sum(int(c.get("n", 0)) for c in s.get("matrix", []))
                    if len(ids) != want:
                        raise ValueError("Bản xem trước không khớp số câu; hãy xem trước lại")
            base = build_section(s, usage, taken, ids or None, rng, counts)
            parts.append((str(s.get("title", "")).strip(), base, bool(s.get("shuffle", True)),
                          bool(s.get("listening"))))
        if clo_mode:  # đáp án theo CĐR: mọi câu trong đề phải có CLO
            missing = [it["code"] for _, base, _, _ in parts for it in base if not it.get("clo")]
            if missing:
                raise ValueError("Đáp án theo CĐR: còn câu chưa gắn CLO: " + ", ".join(missing[:15]))
        if not parts:
            raise ValueError("Chưa có phần nào có số câu")
        if listening and not any(flag for *_, flag in parts):
            raise ValueError("Chưa đánh dấu phần nào là phần nghe (ô \"Phần nghe\" ở bước Số câu và đoạn)")
        reused = sum(1 for _, base, *_ in parts for it in base if it.get("last_used"))
        meta = body.get("meta", {})
        so_tc = str(meta.get("so_tc", "")).strip()
        st = str(meta.get("so_trang", "")).strip()
        if st and not st.isdigit():
            raise ValueError("Số trang phải là số nguyên")
        if so_tc and not re.fullmatch(r"\d+(\.\d)?", so_tc):
            raise ValueError("Số tín chỉ phải là số, ví dụ 4 hoặc 2.5")
        tg = str(meta.get("thoi_gian", "")).strip()
        if tg and not (tg.isdigit() and int(tg) > 0):
            raise ValueError("Thời gian làm bài phải là số phút dương")
        codes_all = [f"{prefix}{start + i:03d}" for i in range(variants)]
        exam_name = str(meta.get("ten_de", "")).strip() or default_exam_name(meta, codes_all)
        now = datetime.now().isoformat(timespec="seconds")
        used_at = datetime.now().isoformat(timespec="microseconds")  # mỗi lần xuất một mốc riêng để đếm đúng số lần dùng
        out = io.BytesIO()
        usage_rows = []
        codes = []
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for i in range(variants):
                code = f"{prefix}{start + i:03d}"
                codes.append(code)
                exam = [(t, make_variant(b, sh, rng)) for t, b, sh, _ in parts]
                # file nghe chỉ ghép các câu của phần được đánh dấu "Phần nghe", giữ nguyên số thứ tự trong đề
                lis_flags = [flag for *_, flag in parts]
                data = render_docx(tpl_path, body.get("meta", {}), exam, code,
                                   answer_key and not key_separate, furigana)
                z.writestr(f"De_{code}.docx", data)
                if listening:  # file đề nghe ghép sẵn theo thứ tự câu của mã đề này
                    track = listening_track(code, listening_only(exam, lis_flags), listening)
                    if track:
                        z.writestr(f"Am_thanh/De_{code}_nghe.mp3", track)
                        listening_made += 1
                for _, _, it in flatten(exam):  # âm thanh của câu trong đề được đưa vào thư mục Am_thanh
                    for m in (it.get("media") or []) + (it.get("passage_media") or []):
                        path = os.path.join(MEDIA_DIR, m["file"])
                        if m["kind"] == "audio" and os.path.isfile(path) and f"Am_thanh/{m['file']}" not in z.namelist():
                            z.write(path, f"Am_thanh/{m['file']}")
                if key_separate:
                    if os.path.isfile(key_tpl_path):
                        z.writestr(f"DapAn_{code}.docx", render_key_doc(key_tpl_path, body.get("meta", {}), code, exam, clo_mode))
                    else:
                        z.writestr(f"DapAn_{code}.docx", render_key(code, exam))
                if body.get("record_usage", True):
                    usage_rows += [(it["id"], code, exam_name, used_at) for _, _, it in flatten(exam)]
        record_usage(usage_rows)
        data = out.getvalue()
        if body.get("save_history", True):  # lưu bản xuất để tải lại hoặc in lại đáp án
            if q("SELECT COUNT(*) AS n FROM exam_exports", fetch=True)[0]["n"] > 100:
                q("DELETE FROM exam_exports")  # quá 100 bản: reset lịch sử về 0, coi như chưa xuất lần nào
            q("INSERT INTO exam_exports(created_at, name, codes, config, data) VALUES(?,?,?,?,?)",
              (now, exam_name, ", ".join(codes),
               json.dumps({"meta": body.get("meta", {}), "variants": variants, "start": start,
                           "prefix": prefix, "template": tpl, "seed": seed,
                           "picked_ids": [[it["id"] for it in base] for _, base, *_ in parts]}, ensure_ascii=False), data))

        fname = export_filename(meta.get("ma_hp"), codes_all, now)
        self._send(200, data, "application/zip",
                   {"X-Reused": str(reused), "X-Variants": str(variants), "X-Seed": str(seed),
                    "X-Listening": str(listening_made) if listening else "",
                    "Content-Disposition": attachment_header(fname)})


def reset_usage(body):
    """Xóa lịch sử ra đề của câu, của các chương, hoặc toàn bộ. Không xóa câu hỏi hay file đề đã xuất.
    Trước khi xóa luôn ghi bản sao: qbank_<ngày>.json (cả CSDL) và truoc_reset_<thời điểm>.json (lịch sử bị xóa)."""
    scope = body.get("scope")
    if scope in ("questions", "chapters"):
        ids = [int(i) for i in body.get("ids", [])]
        if not ids:
            raise ValueError("Chưa chọn mục nào để reset")
        ph = ",".join("?" * len(ids))
        if scope == "questions":
            where, args = f"question_id IN ({ph})", ids
        else:
            where, args = f"question_id IN (SELECT id FROM questions WHERE chapter_id IN ({ph}))", ids
    elif scope == "all":
        if body.get("confirm") != "RESET":
            raise ValueError("Phải gõ RESET để xác nhận reset toàn bộ")
        where, args = "1=1", []
    else:
        raise ValueError("Phạm vi reset không hợp lệ")
    auto_backup()  # bản sao đầy đủ của hôm nay (đã gồm lịch sử dùng)
    rows = [{k: r[k] for k in r.keys()} for r in q(f"SELECT * FROM exam_usage WHERE {where}", args, fetch=True)]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    snap = os.path.join(BACKUP_DIR, f"truoc_reset_{stamp}.json")
    with open(snap, "w", encoding="utf-8") as f:
        json.dump({"scope": scope, "ids": body.get("ids", []), "usage": rows}, f, ensure_ascii=False, indent=1)
    with closing(sqlite3.connect(DB_PATH)) as conn:
        n = conn.execute(f"DELETE FROM exam_usage WHERE {where}", args).rowcount
        conn.commit()
    return {"deleted": n, "snapshot": os.path.basename(snap)}


def item_analysis(tables):
    """Phân tích câu hỏi từ bảng điểm: mỗi dòng một sinh viên, cột đầu là tên/mã sinh viên,
    các cột sau là mã câu (ví dụ N4-MN1-B01-001). Điểm từng câu là 0–1 (1 = đúng hoàn toàn, 0 = sai, hoặc tỷ lệ điểm).
    Trả về độ khó p (tỷ lệ đúng TB), chỉ số phân biệt D (nhóm 27% cao trừ nhóm 27% thấp),
    hệ số tương quan câu–tổng (không tính câu đó) và độ tin cậy KR-20."""
    rows = [r for r in (tables[0] if tables else []) if any(str(c).strip() for c in r)]
    if len(rows) < 2:
        raise ValueError("Bảng điểm cần dòng tiêu đề và ít nhất một sinh viên")
    head = [h.strip() for h in rows[0]]
    known = {r["code"]: r["id"] for r in q("SELECT id, code FROM questions WHERE code IS NOT NULL", fetch=True)}
    cols, unknown = [], []
    for idx, h in enumerate(head[1:], 1):
        if h in known:
            cols.append((idx, h))
        elif h:
            unknown.append(h)
    if not cols:
        raise ValueError("Không cột nào trùng mã câu trong ngân hàng (tiêu đề cột phải là mã câu, ví dụ N4-MN1-B01-001)")
    students, errors = [], []
    for n, r in enumerate(rows[1:], 2):
        vals = {}
        for idx, code in cols:
            raw = str(r[idx]).strip().replace(",", ".") if idx < len(r) else ""
            if raw == "":
                continue
            try:
                v = float(raw[:-1]) / 100 if raw.endswith("%") else float(raw)
            except ValueError:
                errors.append(f"Dòng {n}, cột {code}: '{raw}' không phải số")
                continue
            if not 0 <= v <= 1:
                errors.append(f"Dòng {n}, cột {code}: {raw} ngoài khoảng 0–1 (dùng 1 đúng, 0 sai, hoặc tỷ lệ điểm)")
                continue
            vals[code] = v
        if vals:
            students.append((str(r[0]).strip() if r else f"dòng {n}", vals))
    if len(students) < 2:
        raise ValueError("Cần ít nhất 2 sinh viên có điểm hợp lệ" + (f". Lỗi: {'; '.join(errors[:3])}" if errors else ""))
    totals = {name: sum(v.values()) for name, v in students}
    N = len(students)
    k_group = max(1, round(0.27 * N))
    items = []
    for _, code in cols:
        scored = [(totals[name], vals[code]) for name, vals in students if code in vals]
        if len(scored) < 2:
            continue
        scores = [x for _, x in scored]
        p = statistics.mean(scores)
        ordered = sorted(scored, key=lambda t: -t[0])
        g = max(1, round(0.27 * len(ordered)))
        D = statistics.mean(x for _, x in ordered[:g]) - statistics.mean(x for _, x in ordered[-g:])
        rest = [totals[name] - vals[code] for name, vals in students if code in vals]  # tổng không tính câu này
        r_pb = None
        try:
            r_pb = statistics.correlation(scores, rest)
        except (statistics.StatisticsError, ValueError):
            pass  # không đủ biến thiên (tất cả đúng hoặc tất cả sai)
        flag = ""
        if p <= 0.2:
            flag = "khó bất thường"
        elif p >= 0.95:
            flag = "dễ bất thường"
        if D < 0:
            flag = (flag + "; " if flag else "") + "phân biệt âm: kiểm tra đáp án"
        elif D < 0.2:
            flag = (flag + "; " if flag else "") + "phân biệt kém"
        items.append({"id": known[code], "code": code, "n": len(scored), "p": round(p, 3),
                      "pct": round(p * 100, 1), "D": round(D, 3),
                      "r": round(r_pb, 3) if r_pb is not None else None, "flag": flag})
    # KR-20 trên các câu đã phân tích, chỉ tính khi mọi sinh viên có đủ các câu đó
    kr20 = None
    codes = [it["code"] for it in items]
    full = [vals for _, vals in students if all(c in vals for c in codes)]
    if len(codes) > 1 and len(full) >= 2:
        pq = sum(statistics.mean(v[c] for v in full) * (1 - statistics.mean(v[c] for v in full)) for c in codes)
        tot = [sum(v[c] for c in codes) for v in full]
        var = statistics.pvariance(tot)
        if var > 0:
            kr20 = round(len(codes) / (len(codes) - 1) * (1 - pq / var), 3)
    return {"students": N, "groups": k_group, "items": items, "kr20": kr20,
            "unknown_columns": unknown, "errors": errors[:30]}


def blueprint(flat):
    """Bảng đặc tả: CLO × mức Bloom (số câu), kèm điểm từng dòng. flat: [(số thứ tự, phần, câu)]."""
    rows = {}
    for _, _, it in flat:
        key = it.get("clo") or "(chưa gắn CLO)"
        r = rows.setdefault(key, {"counts": {b: 0 for b in BLOOM}, "n": 0, "points": 0.0})
        r["counts"][it["bloom"]] += 1
        r["n"] += 1
        r["points"] += it.get("points") or 0
    return rows


def blueprint_html(flat):
    rows = blueprint(flat)
    if not rows:
        return ""
    head = "".join(f"<th>{b}. {v}</th>" for b, v in BLOOM.items())
    body = "".join(
        f"<tr><td>{hesc(k)}</td>" + "".join(f"<td>{r['counts'][b] or ''}</td>" for b in BLOOM)
        + f"<td><b>{r['n']}</b></td><td>{fmt_points(r['points']) if r['points'] else ''}</td></tr>"
        for k, r in sorted(rows.items()))
    total = sum(r["n"] for r in rows.values())
    tot_pts = sum(r["points"] for r in rows.values())
    return ("<h3>Bảng đặc tả đề</h3><table><thead><tr><th>CLO</th>" + head + "<th>Số câu</th><th>Điểm</th></tr></thead>"
            f"<tbody>{body}<tr><td><b>Tổng</b></td>" + "".join(
                f"<td><b>{sum(r['counts'][b] for r in rows.values())}</b></td>" for b in BLOOM)
            + f"<td><b>{total}</b></td><td><b>{fmt_points(tot_pts) if tot_pts else ''}</b></td></tr></tbody></table>")


RUBY_HTML = re.compile(r"\{([^{}|]+)\|([^{}]+)\}")


def rich_html(text):
    """Văn bản câu hỏi sang HTML: escape, furigana {chữ|cách đọc} thành <ruby>, xuống dòng thành <br>."""
    out = RUBY_HTML.sub(lambda m: f"<ruby>{m.group(1)}<rt>{m.group(2)}</rt></ruby>", hesc(text))
    return out.replace("\n", "<br>")


def _print_media(media):
    out = []
    for m in media or []:
        if m["kind"] == "image":
            out.append(f"<div><img src='/media/{m['file']}' style='max-width:100%;max-height:260px'></div>")
        else:
            out.append(f"<div><audio controls preload='none' src='/media/{m['file']}'></audio> <small>{hesc(m['name'])}</small></div>")
    return out


def print_html(body):
    """Bản in xem trước của một mã đề, dựng từ cùng cách chọn câu như khi xuất (không ghi gì vào CSDL)."""
    sections = body.get("sections") or []
    picked = body.get("picked")
    usage = usage_map()
    taken = set()
    parts = []
    for idx, s in enumerate(sections):
        if not section_has_content(s):
            continue
        ids = picked[idx] if picked and idx < len(picked) else None
        base = build_section(s, usage, taken, ids or None)
        parts.append((str(s.get("title", "")).strip(), base, bool(s.get("shuffle", True))))
    if not parts:
        raise ValueError("Chưa có phần nào có số câu")
    rng = random.Random(1)  # cố định để xem lại cho cùng một bản
    exam = [(t, make_variant(b, sh, rng)) for t, b, sh in parts]
    flat = flatten(exam)
    meta = body.get("meta", {})
    code = f"{body.get('prefix', '')}{int(body.get('start', 1) or 1):03d}"
    sec_html = []
    last_title, last_passage, last_match = None, None, None
    for n, title, it in flat:
        if title and title != last_title:
            sec_html.append(f"<h2>{hesc(title)}</h2>")
            last_title = title
        pts = f" <i>({fmt_points(it['points'])} điểm)</i>" if it.get("points") is not None else ""
        if it.get("passage") and it["passage"] != last_passage:  # đoạn văn và file nghe của đoạn, in trước câu đầu
            sec_html.append(f"<div class='q' style='border:1px solid #999;padding:6px'>{rich_html(it['passage'])}"
                            + "".join(_print_media(it.get("passage_media"))) + "</div>")
        last_passage = it.get("passage") or None
        lines = [f"<div class='q'><b>Câu {n}.</b> {rich_html(it['content'])}{pts}"]
        lines += _print_media(it.get("media"))
        if is_match(it):  # danh sách tiêu đề in một lần trước nhóm câu (xem bản xuất đề)
            key = (it.get("passage_id"), tuple(it["options"]))
            if key != last_match:
                sec_html.append("<div class='q'><b>Danh sách tiêu đề:</b> " + " ".join(
                    f"{LETTERS[k]}. {rich_html(o)}" for k, o in enumerate(it["options"])) + "</div>")
                last_match = key
            lines.append("<div class='fill'>Đáp án (chữ cái tiêu đề): ……</div>")
        elif it["qtype"] == "TN":
            lines.append("<div class='opts'>" + "".join(
                f"<span>{LETTERS[k]}. {rich_html(o)}</span>" for k, o in enumerate(it["options"])) + "</div>")
        elif it["qtype"] == "DS":
            lines.append("<div class='opts ds'>" + "".join(
                f"<span>{'abcdefgh'[k]}) {rich_html(o)}</span>" for k, o in enumerate(it["options"])) + "</div>")
        elif it["qtype"] == "DT":
            limit = f" (không quá {it['word_limit']} từ)" if it.get("word_limit") else ""
            lines.append(f"<div class='fill'>Đáp án{limit}: ……………………………</div>")
        else:
            lines.append("<div class='essay'>&nbsp;</div>")
        sec_html.append("".join(lines) + "</div>")
    ans = " · ".join(f"{n}-{answer_text(it)}" for n, _, it in flat)
    rubric = "".join(f"<p>{hesc(line)}</p>" for line in rubric_lines(flat)[1:] if line)
    head = " · ".join(x for x in [hesc(str(meta.get("ten_hp", ""))), f"Mã HP {hesc(str(meta.get('ma_hp', '')))}",
                                  f"Thời gian {hesc(str(meta.get('thoi_gian', '')))} phút", f"Mã đề {code}"] if x and not x.endswith(" "))
    return f"""<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>Bản in {code}</title>
<style>
 body {{ font: 15px/1.5 "Times New Roman", serif; max-width: 780px; margin: 24px auto; padding: 0 16px; color:#000; }}
 h2 {{ font-size: 16px; margin: 18px 0 6px; }} h3 {{ font-size: 15px; margin-top: 24px; }}
 .q {{ margin: 8px 0 10px; }} .opts {{ display: grid; grid-template-columns: 1fr 1fr; gap: 2px 24px; margin-left: 18px; }}
 .ds {{ display: block; }} .ds span {{ display: block; }} .fill {{ margin-left: 18px; }} .essay {{ height: 60px; border-bottom: 1px dotted #888; margin: 6px 18px; }}
 rt {{ font-size: 60%; }} table {{ border-collapse: collapse; font-size: 13px; }} td, th {{ border: 1px solid #999; padding: 3px 8px; text-align: center; }}
 .note {{ background:#fef9c3; padding:8px 10px; font-size:13px; border-radius:6px; }}
 @media print {{ .noprint {{ display: none; }} }}
</style></head><body>
<div class="noprint note">Bản xem trước này để rà soát nội dung. Bố cục thật theo mẫu Word khi xuất đề có thể khác đôi chút.
 <button onclick="window.print()">In</button></div>
<p><b>{head}</b></p>
{''.join(sec_html)}
<h3>Đáp án mã {code}</h3><p>{hesc(ans)}</p>{rubric}
{blueprint_html(flat)}
</body></html>"""


# ---------------------------------------------------------------- xuất Moodle XML
def _mx(text):
    """Escape cho nội dung HTML đặt trong thẻ <text> của Moodle."""
    return escape(text or "", {'"': "&quot;"})


def _mfile(m):
    """Thẻ <file> nhúng file đính kèm (base64) để Moodle lưu cùng câu hỏi."""
    path = os.path.join(MEDIA_DIR, m["file"])
    if not os.path.isfile(path):
        return ""
    b64 = base64.b64encode(open(path, "rb").read()).decode()
    return f'<file name="{_mx(m["file"])}" path="/" encoding="base64">{b64}</file>'


def _mmedia(media):
    """HTML và file đính kèm cho một câu: ảnh nhúng, âm thanh là liên kết đến file."""
    html, files = "", ""
    for m in media or []:
        if m["kind"] == "image":
            html += f'<p><img src="@@PLUGINFILE@@/{_mx(m["file"])}" alt="{_mx(m["name"])}"></p>'
        else:
            html += f'<p><a href="@@PLUGINFILE@@/{_mx(m["file"])}">🔊 {_mx(m["name"])}</a></p>'
        files += _mfile(m)
    return html, files


def _mq(name, qtype, text, media, grade, extra):
    """Một câu hỏi Moodle. extra là các thẻ riêng theo loại câu (đáp án, kiểu trả lời)."""
    mh, mf = _mmedia(media)
    return (f'<question type="{qtype}"><name><text>{_mx(name)}</text></name>'
            f'<questiontext format="html"><text>{_mx(text)}{_mx(mh)}</text>{mf}</questiontext>'
            f'<generalfeedback format="html"><text></text></generalfeedback>'
            f'<defaultgrade>{grade}</defaultgrade><penalty>0.3333333</penalty><hidden>0</hidden>'
            f'{extra}</question>')


def moodle_xml(ids=None):
    """Bộ câu hỏi Moodle XML. Không truyền ids: lấy mọi câu đã duyệt. Mỗi học phần/chương là một danh mục."""
    where, args = "WHERE qs.status='da_duyet'", []  # chỉ câu đã duyệt; câu nháp hoặc ngưng dùng không xuất
    if ids:
        where += f" AND qs.id IN ({','.join('?' * len(ids))})"
        args = [int(i) for i in ids]
    rows = [row_to_api(r) for r in q(QUESTION_SQL + f" {where} ORDER BY qs.id", args, fetch=True)]
    out, cats = [], set()
    for r in rows:
        parts = [str(x).replace("/", "-") for x in (r["course"], r["textbook"], r["chapter_code"])]  # học phần đại diện
        cat = "$course$/" + "/".join(parts)  # danh mục Moodle phân cấp bằng dấu /
        if cat not in cats:
            cats.add(cat)
            out.append(f'<question type="category"><category><text>{_mx(cat)}</text></category></question>')
        grade = fmt_points(r["points"]).replace(",", ".") if r.get("points") is not None else "1"
        r["media"] = (r.get("passage_media") or []) + (r.get("media") or [])
        stem = rich_html(r["content"]).replace("<br>", "<br>")
        if r["qtype"] == "TN":
            letters = LETTERS
            ans = "".join(
                f'<answer fraction="{100 if letters[k] == r["answer"] else 0}" format="html"><text>{_mx(rich_html(o))}</text>'
                f'<feedback format="html"><text></text></feedback></answer>' for k, o in enumerate(r["options"]))
            out.append(_mq(r["code"], "multichoice", stem, r["media"], grade,
                           f'<single>true</single><shuffleanswers>true</shuffleanswers><answernumbering>ABCD</answernumbering>'
                           f'<correctfeedback format="html"><text></text></correctfeedback>'
                           f'<partiallycorrectfeedback format="html"><text></text></partiallycorrectfeedback>'
                           f'<incorrectfeedback format="html"><text></text></incorrectfeedback>{ans}'))
        elif r["qtype"] == "DS":  # mỗi ý đúng/sai thành một câu riêng, có chung đề
            for k, o in enumerate(r["options"]):
                truth = r["answer"][k] == "D"
                text = f"{r['content']}\n{'abcdefgh'[k]}) {o}"
                ans = (f'<answer fraction="{100 if truth else 0}" format="html"><text>true</text><feedback format="html"><text></text></feedback></answer>'
                       f'<answer fraction="{0 if truth else 100}" format="html"><text>false</text><feedback format="html"><text></text></feedback></answer>')
                out.append(_mq(f"{r['code']}-{'abcdefgh'[k]}", "truefalse", rich_html(text), r["media"], grade, ans))
        elif r["qtype"] == "DT":
            ans = "".join(f'<answer fraction="100" format="plain_text"><text>{_mx(a.strip())}</text>'
                          f'<feedback format="html"><text></text></feedback></answer>'
                          for a in r["answer"].split("|") if a.strip())
            out.append(_mq(r["code"], "shortanswer", stem, r["media"], grade,
                           f"<usecase>0</usecase>{ans}"))
        else:  # TL: gợi ý chấm để người chấm đọc
            out.append(_mq(r["code"], "essay", stem, r["media"], grade,
                           f'<responseformat>editor</responseformat><responserequired>1</responserequired>'
                           f'<responsefieldlines>12</responsefieldlines><attachments>0</attachments>'
                           f'<graderinfo format="html"><text>{_mx(r["answer"])}</text></graderinfo>'
                           f'<responsetemplate format="html"><text></text></responsetemplate>'))
    if not rows:
        raise ValueError("Không có câu đã duyệt nào để xuất (câu nháp và câu ngưng dùng không được xuất)")
    return '<?xml version="1.0" encoding="UTF-8"?>\n<quiz>\n' + "\n".join(out) + "\n</quiz>\n"


# ---------------------------------------------------------------- cập nhật từ GitHub Releases
UPDATE_REPO = os.environ.get("QBANK_UPDATE_REPO", "miyagitaisuki/demix")
UPDATE_TAG_PREFIX = "qbank-v"  # release của QBank có tag dạng qbank-v1.2.3
UPDATE_ASSET = "QBank.exe"
UPDATE_HOSTS = ("github.com", "objects.githubusercontent.com", "release-assets.githubusercontent.com")
RELEASES_PAGE = f"https://github.com/{UPDATE_REPO}/releases"
UPDATE_DIR = os.path.join(DATA_DIR, "data", "update")  # file tải về tạm, nằm cạnh dữ liệu nên không mất khi cập nhật


def read_version():
    try:
        with open(os.path.join(RES_DIR, "VERSION"), encoding="utf-8") as f:
            return f.read().strip() or "0.0.0"
    except OSError:
        return "0.0.0"


APP_VERSION = read_version()


def parse_version(v):
    """'1.2.3' hoặc 'qbank-v1.2.3' thành (1, 2, 3); không đọc được thì None."""
    m = re.fullmatch(r"(?:qbank-)?v?(\d+)\.(\d+)\.(\d+)", (v or "").strip())
    return tuple(int(x) for x in m.groups()) if m else None


def safe_download_url(url):
    """Chỉ tải từ HTTPS và các máy chủ file của GitHub."""
    u = urlparse(url or "")
    return u.scheme == "https" and (u.hostname or "") in UPDATE_HOSTS


def pick_update(releases, current):
    """Chọn bản mới nhất (bỏ qua nháp, bản thử, tag không phải qbank-v*, release thiếu QBank.exe).
    Trả về dict nếu mới hơn bản đang chạy, không thì None."""
    cur = parse_version(current)
    best = None
    for r in releases:
        tag = r.get("tag_name") or ""
        ver = parse_version(tag)
        if r.get("draft") or r.get("prerelease") or ver is None or not tag.startswith(UPDATE_TAG_PREFIX):
            continue
        asset = next((a for a in r.get("assets") or [] if a.get("name") == UPDATE_ASSET), None)
        if asset is None or not safe_download_url(asset.get("browser_download_url")):
            continue
        if best is None or ver > best[0]:
            best = (ver, asset, r)
    if best is None or cur is None or best[0] <= cur:
        return None
    ver, asset, r = best
    return {"version": ".".join(map(str, ver)), "url": asset["browser_download_url"],
            "size": asset.get("size") or 0, "digest": asset.get("digest") or "", "notes": r.get("body") or ""}


def fetch_releases():
    req = urlreq.Request(f"https://api.github.com/repos/{UPDATE_REPO}/releases?per_page=30",
                         headers={"Accept": "application/vnd.github+json", "User-Agent": "QBank-updater"})
    with urlreq.urlopen(req, timeout=10) as resp:
        return json.loads(resp.read().decode("utf-8"))


def latest_update():
    try:
        releases = fetch_releases()
    except (URLError, HTTPError, OSError, ValueError) as e:
        raise UpdateError(f"Không kiểm tra được bản cập nhật ({e}). Hãy kiểm tra kết nối mạng.")
    return pick_update(releases, APP_VERSION)


def can_self_update():
    """Chỉ tự thay file khi đang chạy QBank.exe đóng gói trên Windows."""
    return bool(getattr(sys, "frozen", False)) and os.name == "nt"


def check_update():
    upd = latest_update()
    return {"current": APP_VERSION, "available": upd is not None,
            "latest": upd["version"] if upd else APP_VERSION,
            "notes": upd["notes"] if upd else "", "can_install": can_self_update(), "releases_url": RELEASES_PAGE}


def download_update(upd):
    os.makedirs(UPDATE_DIR, exist_ok=True)
    new_path = os.path.join(UPDATE_DIR, "QBank.new.exe")
    h, size = hashlib.sha256(), 0
    try:
        with urlreq.urlopen(urlreq.Request(upd["url"], headers={"User-Agent": "QBank-updater"}), timeout=60) as resp, \
                open(new_path, "wb") as f:
            while chunk := resp.read(1 << 16):
                f.write(chunk)
                h.update(chunk)
                size += len(chunk)
    except (URLError, HTTPError, OSError) as e:
        raise UpdateError(f"Tải bản cập nhật thất bại ({e}).")
    with open(new_path, "rb") as f:
        head = f.read(2)
    digest = upd.get("digest") or ""
    bad = head != b"MZ" or (upd["size"] and size != upd["size"]) or (
        digest.startswith("sha256:") and digest[len("sha256:"):] != h.hexdigest())
    if bad:
        os.remove(new_path)
        raise UpdateError("File tải về không đúng (sai định dạng, kích thước hoặc checksum). Đã hủy cập nhật.")
    return new_path


def swap_exe(new_path, current_path):
    """Windows cho đổi tên file .exe đang chạy nhưng không ghi đè được: đổi bản cũ thành .old rồi đưa bản mới vào chỗ cũ."""
    old_path = current_path + ".old"
    if os.path.exists(old_path):
        os.remove(old_path)  # bản .old của lần cập nhật trước, không còn chạy
    os.replace(current_path, old_path)
    try:
        os.replace(new_path, current_path)
    except OSError:
        os.replace(old_path, current_path)  # lỗi thì trả lại bản cũ
        raise
    return old_path


def install_update():
    if not can_self_update():
        raise ValueError("Chỉ cập nhật tự động khi chạy QBank.exe trên Windows. Hãy tải bản mới từ trang Releases.")
    upd = latest_update()
    if upd is None:
        raise ValueError("Đang là bản mới nhất, không có gì để cài.")
    auto_backup()  # sao lưu dữ liệu của hôm nay trước khi thay file
    new_path = download_update(upd)
    swap_exe(new_path, sys.executable)
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([sys.executable, "--after-update"], close_fds=True, creationflags=flags)
    return upd["version"]


def cleanup_old_exe():
    """Xóa bản .old của lần cập nhật trước (nếu đã tắt được)."""
    if getattr(sys, "frozen", False):
        try:
            os.remove(sys.executable + ".old")
        except OSError:
            pass


def main():
    init_storage()
    backup_loop()
    cleanup_old_exe()
    url = f"http://127.0.0.1:{PORT}"
    # bản vừa cập nhật khởi động khi bản cũ còn giữ cổng: chờ tối đa 30 giây rồi mới báo lỗi
    deadline = time.time() + (30 if "--after-update" in sys.argv else 0)
    srv = None
    while True:
        try:
            srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
            break
        except OSError:
            if time.time() >= deadline:
                break
            time.sleep(0.5)
    if srv is None:  # ra ngoài khối except để không in kèm traceback lồng nhau
        print(f"Cổng {PORT} đang được dùng. Hãy đóng chương trình QBank đang chạy khác rồi thử lại.")
        if getattr(sys, "frozen", False):
            input("Nhấn Enter để thoát...")
        sys.exit(1)
    print(f"QBank đang chạy tại {url}  (đóng cửa sổ này hoặc Ctrl+C để dừng)")
    if not os.environ.get("QBANK_NO_BROWSER"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()  # mở trình duyệt sau khi server sẵn sàng
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        sys.exit(0)


if __name__ == "__main__":
    main()
