#!/usr/bin/env python3
"""QBank - ngân hàng câu hỏi offline, giao diện web nhẹ chạy trên máy cục bộ.

Cấu trúc phân loại: Học phần -> Giáo trình -> Chương (mã duy nhất trong giáo trình) -> Câu hỏi.

Chạy:  python3 app.py   rồi mở http://127.0.0.1:8765
Phụ thuộc: python-docx (đọc bảng câu hỏi và xuất đề .docx).
Dữ liệu: data/qbank.db (SQLite). Template Word: templates/*.docx
"""
import io
import json
import os
import random
import re
import sqlite3
import sys
import zipfile
from contextlib import closing
from datetime import datetime
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE, "data", "qbank.db")
TPL_DIR = os.path.join(BASE, "templates")
STATIC = os.path.join(BASE, "static")
BLOOM = {1: "Biết", 2: "Hiểu", 3: "Vận dụng", 4: "Phân tích", 5: "Đánh giá"}
LETTERS = "ABCDEFGH"
QTYPES = ("TN", "DS", "DT", "TL")  # trắc nghiệm, đúng/sai, điền từ, tự luận
DEFAULT_COURSE = "Chưa phân loại"
PORT = int(os.environ.get("QBANK_PORT", "8765"))


# ---------------------------------------------------------------- database
def q(sql, args=(), fetch=False):
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.execute(sql, args)
        rows = [dict(r) for r in cur.fetchall()] if fetch else None
        conn.commit()
        return rows


def init_storage():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(TPL_DIR, exist_ok=True)
    migrate_if_needed()
    q("""CREATE TABLE IF NOT EXISTS courses(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL UNIQUE,
            code TEXT NOT NULL DEFAULT '')""")
    q("""CREATE TABLE IF NOT EXISTS textbooks(
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_id INTEGER NOT NULL REFERENCES courses(id),
            name TEXT NOT NULL,
            code TEXT NOT NULL DEFAULT '',
            UNIQUE(course_id, name))""")
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
    ensure_column("courses", "code", "TEXT NOT NULL DEFAULT ''")
    ensure_column("courses", "exam_meta", "TEXT NOT NULL DEFAULT '{}'")
    ensure_column("textbooks", "code", "TEXT NOT NULL DEFAULT ''")
    ensure_column("questions", "code", "TEXT")
    ensure_column("questions", "passage", "TEXT NOT NULL DEFAULT ''")
    q("CREATE UNIQUE INDEX IF NOT EXISTS idx_questions_code ON questions(code)")
    for r in q("SELECT id FROM questions WHERE code IS NULL OR code=''", fetch=True):
        q("UPDATE questions SET code=? WHERE id=?", (make_code(r["id"]), r["id"]))
    if not any(f.endswith(".docx") for f in os.listdir(TPL_DIR)):
        make_sample_templates()
    if not os.path.exists(os.path.join(TPL_DIR, "de_thi_chuan.docx")):
        make_official_templates()


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
    # Dữ liệu cũ không có học phần/giáo trình: đưa vào một nhóm "Chưa phân loại" để người dùng gán lại.
    # Mỗi mã chương cũ thành một chương trong giáo trình "Chưa phân loại"
    with closing(sqlite3.connect(DB_PATH)) as conn:
        codes = [r[0] for r in conn.execute("SELECT DISTINCT chapter FROM questions_old")]
    for code in codes:
        ensure_chapter(DEFAULT_COURSE, DEFAULT_COURSE, code)
    q("""INSERT INTO questions(id, chapter_id, bloom, qtype, content, options, answer, solution)
         SELECT o.id, (SELECT c.id FROM chapters c JOIN textbooks t ON t.id=c.textbook_id
                       JOIN courses co ON co.id=t.course_id
                       WHERE co.name=? AND t.name=? AND c.code=o.chapter),
                o.bloom, o.qtype, o.content, o.options, o.answer, o.solution
         FROM questions_old o""", (DEFAULT_COURSE, DEFAULT_COURSE))
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = OFF")
        conn.execute("DROP TABLE questions_old")
        conn.commit()


# ---------------------------------------------------------------- hierarchy helpers
CHAPTER_LABEL = "co.name || ' / ' || t.name || ' / ' || ch.code || CASE WHEN ch.title <> '' THEN ' - ' || ch.title ELSE '' END"
QUESTION_SQL = f"""SELECT qs.*, ch.code AS chapter_code, ch.title AS chapter_title,
    (SELECT COUNT(*) FROM exam_usage u WHERE u.question_id = qs.id) AS used_count,
    ch.id AS chapter_id, t.id AS textbook_id, t.name AS textbook,
    co.id AS course_id, co.name AS course, ({CHAPTER_LABEL}) AS chapter_label
    FROM questions qs
    JOIN chapters ch ON ch.id = qs.chapter_id
    JOIN textbooks t ON t.id = ch.textbook_id
    JOIN courses co ON co.id = t.course_id"""


def ensure_chapter(course, textbook, code, title=""):
    """Tìm hoặc tạo học phần/giáo trình/chương theo tên. Trả về id chương."""
    q("INSERT OR IGNORE INTO courses(name) VALUES(?)", (course,))
    cid = q("SELECT id FROM courses WHERE name=?", (course,), True)[0]["id"]
    q("INSERT OR IGNORE INTO textbooks(course_id, name) VALUES(?, ?)", (cid, textbook))
    tid = q("SELECT id FROM textbooks WHERE course_id=? AND name=?", (cid, textbook), True)[0]["id"]
    q("INSERT OR IGNORE INTO chapters(textbook_id, code, title) VALUES(?, ?, ?)", (tid, code, title))
    return q("SELECT id FROM chapters WHERE textbook_id=? AND code=?", (tid, code), True)[0]["id"]


def chapter_label(chapter_id):
    rows = q(f"""SELECT ({CHAPTER_LABEL}) AS label FROM chapters ch
                 JOIN textbooks t ON t.id=ch.textbook_id JOIN courses co ON co.id=t.course_id
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
    textbooks = q("SELECT id, course_id, name, code FROM textbooks ORDER BY name", fetch=True)
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
    if qtype == "TN":
        if len(options) < 2:
            raise ValueError("Trắc nghiệm cần ít nhất 2 phương án")
        if answer not in LETTERS[:len(options)]:
            raise ValueError("Đáp án không hợp lệ")
    elif qtype == "DS":
        if not 1 <= len(options) <= 8:
            raise ValueError("Đúng/sai cần từ 1 đến 8 ý")
        answer = normalize_ds(answer, len(options))
    elif qtype == "DT":
        if not answer:
            raise ValueError("Thiếu đáp án điền từ (nhiều đáp án chấp nhận cách nhau bằng |)")
        options = []
    else:  # TL
        options, answer = [], ""
    passage = str(data.get("passage", "")).strip()
    return dict(chapter_id=chapter_id, bloom=bloom, qtype=qtype, content=content,
                options=json.dumps(options, ensure_ascii=False), answer=answer, solution=solution,
                passage=passage)


INSERT_Q = """INSERT INTO questions(code,chapter_id,bloom,qtype,content,options,answer,solution,passage)
              VALUES(:code,:chapter_id,:bloom,:qtype,:content,:options,:answer,:solution,:passage)"""


def row_to_api(r):
    r = dict(r)
    r["options"] = json.loads(r["options"])
    return r


def make_code(question_id_or_none, chapter_id=None):
    """Mã câu hỏi ổn định, dạng HọcPhần-GiáoTrình-Chương-Số (ví dụ N5-MN1-B03-001).
    Mã được gán khi tạo câu và không đổi sau đó. Số tăng theo từng chương, bỏ qua mã đã có."""
    if chapter_id is None:  # gọi khi nâng cấp CSDL: lấy chương của câu
        chapter_id = q("SELECT chapter_id FROM questions WHERE id=?", (question_id_or_none,), True)[0]["chapter_id"]
    r = q("""SELECT ch.code AS chcode, t.id AS tid, t.code AS tcode, co.id AS cid, co.code AS ccode
             FROM chapters ch JOIN textbooks t ON t.id=ch.textbook_id JOIN courses co ON co.id=t.course_id
             WHERE ch.id=?""", (chapter_id,), True)[0]
    hp = r["ccode"] or f"HP{r['cid']}"
    gt = r["tcode"] or f"GT{r['tid']}"
    prefix = re.sub(r"\s+", "_", f"{hp}-{gt}-{r['chcode']}")
    n = q("SELECT COUNT(*) AS n FROM questions WHERE chapter_id=? AND code IS NOT NULL AND code <> ''",
          (chapter_id,), True)[0]["n"] + 1
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
            raise ValueError(f"Câu #{rid} không trùng nội dung với câu #{keep}, không thể gộp")
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
    payload = dict(payload, code=code or make_code(None, payload["chapter_id"]))
    with closing(sqlite3.connect(DB_PATH)) as conn:
        conn.execute("PRAGMA foreign_keys = ON")
        cur = conn.execute(INSERT_Q, payload)
        conn.commit()
        return cur.lastrowid


# ---------------------------------------------------------------- trùng lặp
def dup_key(content, options):
    """Khóa so trùng: bỏ hoa/thường, khoảng trắng, dấu câu. Giữ chữ Nhật, chữ Hán."""
    return re.sub(r"[\W_]+", "", (content + "".join(options)).lower())


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


def parse_import(data: bytes, textbook_id, dry_run=False):
    """Nhập câu hỏi từ bảng Word (xem /api/import-template để có mẫu).
    Dòng đầu là tiêu đề cột; cột nhận diện theo tên nên thứ tự cột tùy ý.
    Bắt buộc: Mã chương, Nội dung, Đáp án, Loại, Mức. Tùy chọn: Đoạn văn, Lời giải.
    Mã chương phải đã có trong giáo trình được chọn; không tự tạo chương mới."""
    chapter_by_code = {r["code"]: r["id"] for r in q(
        "SELECT id, code FROM chapters WHERE textbook_id=?", (textbook_id,), True)}
    doc = Document(io.BytesIO(data))
    added, errors, skipped, warnings = 0, [], [], []
    rows_out = []  # xem trước: trạng thái từng dòng
    index = dup_index()
    for t_idx, table in enumerate(doc.tables, 1):
        if not table.rows:
            continue
        cols, missing = map_import_header([c.text for c in table.rows[0].cells])
        if missing:
            errors.append(f"Bảng {t_idx}: thiếu cột {', '.join(missing)} (dòng tiêu đề)")
            continue
        for r_idx, row in enumerate(table.rows[1:], 2):
            cells = [c.text.strip() for c in row.cells]
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
                payload = clean_question({
                    "chapter_id": chapter_by_code[code],
                    "bloom": int(re.sub(r"\D", "", get("muc")) or 1),
                    "qtype": qtype,
                    "content": " ".join(stem),
                    "options": options,
                    "answer": get("dap_an"),
                    "solution": solution,
                    "passage": get("doan_van"),
                })
                same, other, key = check_duplicate(index, payload)
                label = f"Bảng {t_idx}, dòng {r_idx}"
                preview = {"where": label, "code": code, "qtype": qtype, "content": " ".join(stem)[:80]}
                if same:
                    ref = same[0][0]
                    ref_txt = f"câu #{ref}" if isinstance(ref, int) else ref.lower()  # ref là nhãn dòng khi xem trước
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
                    note = f"trùng nội dung với câu #{other[0][0]} ở chương khác"
                    warnings.append(f"{label}: {note}")
                rows_out.append(dict(preview, status="warn" if other else "ok", note=note))
            except Exception as e:
                errors.append(f"Bảng {t_idx}, dòng {r_idx}: {e}")
                rows_out.append({"where": f"Bảng {t_idx}, dòng {r_idx}", "code": "", "qtype": "",
                                 "content": "", "status": "error", "note": str(e)})
    return {"added": added, "skipped": skipped, "warnings": warnings, "errors": errors,
            "dry_run": dry_run, "rows": rows_out}


def pick_base_exam(matrix, qtype=None, usage=None, taken=None):
    """matrix: [{chapter_id, bloom, n}] -> câu gốc của một phần.
    Ưu tiên câu chưa dùng, sau đó câu dùng lâu nhất; các câu cùng mức được chọn ngẫu nhiên.
    qtype: lọc theo dạng câu (TN/DS/DT/TL) của phần; None = mọi dạng."""
    usage = usage or {}
    taken = taken if taken is not None else set()  # câu đã chọn cho đề này (mọi phần): không chọn lại
    base, shortage = [], []
    for cell in matrix:
        n = int(cell.get("n", 0))
        if n <= 0:
            continue
        sql = "SELECT * FROM questions WHERE chapter_id=? AND bloom=?"
        args = [int(cell["chapter_id"]), int(cell["bloom"])]
        if qtype:
            sql += " AND qtype=?"
            args.append(qtype)
        pool = [row_to_api(p) for p in q(sql, args, fetch=True) if p["id"] not in taken]
        if len(pool) < n:
            shortage.append(f"{chapter_label(cell['chapter_id'])} / Bloom {cell['bloom']}: cần {n}, có {len(pool)}")
            continue
        random.shuffle(pool)
        pool.sort(key=lambda it: usage.get(it["id"]) or "")  # chưa dùng ("") lên trước, rồi cũ tới mới
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
            raise ValueError(f"Câu #{i} không còn trong ngân hàng, hãy xem trước lại")
        it = row_to_api(rows[int(i)])
        it["last_used"] = usage.get(it["id"])
        out.append(it)
    return out


def find_replacement(spec, exclude, usage):
    """Tìm câu thay thế cùng chương, cùng mức Bloom và cùng dạng, chưa có trong đề.
    Ưu tiên câu chưa dùng, rồi câu dùng lâu nhất."""
    if spec.get("passage"):
        raise ValueError("Câu thuộc đoạn văn không đổi riêng lẻ; hãy chọn lại cả phần")
    sql = "SELECT * FROM questions WHERE chapter_id=? AND bloom=? AND qtype=? AND passage=''"
    args = [int(spec["chapter_id"]), int(spec["bloom"]), str(spec["qtype"])]
    pool = [row_to_api(p) for p in q(sql, args, fetch=True) if p["id"] not in exclude]
    if not pool:
        raise ValueError("Không còn câu thay thế cùng chương, mức Bloom và dạng")
    random.shuffle(pool)
    pool.sort(key=lambda it: usage.get(it["id"]) or "")
    it = pool[0]
    it["last_used"] = usage.get(it["id"])
    return it


def preview_exam(sections):
    """Chọn câu cho mọi phần (không ghi lịch sử). Trả về danh sách song song với sections."""
    usage = usage_map()
    taken = set()
    out = []
    for s in sections:
        matrix = s.get("matrix", [])
        if not any(int(c.get("n", 0)) > 0 for c in matrix):
            out.append({"items": []})
            continue
        base = pick_base_exam(matrix, s.get("qtype") or None, usage, taken)
        out.append({"items": base})
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
        groups.setdefault(it["passage"] or f"#{it['id']}", []).append(it)
    units = [sorted(g, key=lambda it: it["id"]) for g in groups.values()]
    units.sort(key=lambda g: g[0]["id"])
    if shuffle:
        rng.shuffle(units)
        counts = {}  # cân bằng đáp án trong cùng một mã đề
        return [shuffle_question(it, rng, counts)[0] for g in units for it in g]
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


def add_rich(p, text, bold=False, mode="ruby"):
    """Thêm văn bản vào đoạn, đổi cú pháp {chữ|cách đọc} thành furigana.
    mode='ruby': chữ đọc phía trên (chuẩn Word); mode='paren': chữ(cách đọc) cùng dòng (dự phòng)."""
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


def emit_question(doc, marker, n, it, mode="ruby"):
    p = marker.insert_paragraph_before("")
    p.add_run(f"Câu {n}. ").bold = True
    add_rich(p, it["content"], mode=mode)
    if it["qtype"] == "TN":
        emit_options(doc, marker, it["options"], mode)
    elif it["qtype"] == "DS":
        for k, opt in enumerate(it["options"]):
            line = marker.insert_paragraph_before("")
            line.add_run(f"    {'abcdefgh'[k]}) ")
            add_rich(line, opt, mode=mode)
    elif it["qtype"] == "DT":
        marker.insert_paragraph_before("    Đáp án: ……………………………")


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
    last_title, last_passage = None, None
    for n, title, it in flat:
        if title and title != last_title:  # tiêu đề phần: "I. ...", in một lần trước câu đầu của phần
            head = marker.insert_paragraph_before(title)
            head.runs[0].bold = True
            last_title = title
        if it.get("passage") and it["passage"] != last_passage:
            emit_passage(doc, marker, it["passage"], mode)
        last_passage = it.get("passage") or None
        emit_question(doc, marker, n, it, mode)
    _set_text(marker, "")

    if answer_key:
        doc.add_paragraph("")
        doc.add_paragraph(f"ĐÁP ÁN MÃ {code}").runs[0].bold = True
        doc.add_paragraph("   ".join(f"{n}-{answer_text(it)}" for n, _, it in flat))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def render_key_doc(template_path, meta, code, exam):
    """Đáp án theo mẫu trường: bảng lưới 10 câu mỗi dòng, dòng 'Câu' và dòng 'Đáp án'."""
    doc = Document(template_path)
    flat = flatten(exam)
    fill_placeholders(doc, meta_values(meta, code, len(flat)))
    ans = [(n, answer_text(it)) for n, _, it in flat]
    groups = [ans[i:i + 10] for i in range(0, len(ans), 10)]
    marker = next((p for p in _every_paragraph(doc) if "{{BANG_DAP_AN}}" in p.text), None)
    if marker is None:
        marker = doc.add_paragraph()
    tbl = doc.add_table(rows=2 * len(groups), cols=11)
    _set_borders(tbl)
    for gi, g in enumerate(groups):
        r1, r2 = tbl.rows[2 * gi], tbl.rows[2 * gi + 1]
        r1.cells[0].text, r2.cells[0].text = "Câu", "Đáp án"
        for j, (n, a) in enumerate(g):
            r1.cells[j + 1].text, r2.cells[j + 1].text = str(n), a
    marker._p.addprevious(tbl._tbl)
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
                    "Lời giải (tùy chọn). Thứ tự cột tùy ý. Furigana viết dạng {試験|しけん}. "
                    "TN: mỗi phương án một dòng bắt đầu bằng A. B. C. D.; # đầu phương án = cố định vị trí. "
                    "DS: mỗi ý một dòng bắt đầu a) b) c); đáp án Đ S Đ S. DT: nhiều đáp án cách nhau bằng |.")
    t = d.add_table(rows=1, cols=7)
    t.style = "Table Grid"
    for c, h in zip(t.rows[0].cells, ["Mã chương", "Nội dung", "Đáp án", "Loại", "Mức", "Đoạn văn", "Lời giải"]):
        c.text = h
    examples = [
        ("B01", "{試験|しけん}の{結果|けっか}が悪かった。\nA. おけばよかった\nB. おくべきだ", "A", "TN", "1", "", ""),
        ("B01", "「インターンシップを通して」と同じ意味はどれですか。\nA. 間に\nB. おかげで", "B", "TN", "2", "私は先月、インターンシップに参加しました。", ""),
        ("B01", "Câu đúng/sai:\na) 私は学生です。\nb) 本を読むです。", "ĐS", "DS", "2", "", ""),
        ("B02", "学校の___", "がっこう", "DT", "1", "", ""),
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
    """rows: [(mã câu, tỷ lệ đúng %)]. Trả về số dòng đã ghi và danh sách lỗi theo dòng."""
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


def stats():
    tree = get_tree()
    counts = {}
    for r in q("SELECT chapter_id, bloom, qtype, COUNT(*) AS n FROM questions GROUP BY chapter_id, bloom, qtype", fetch=True):
        c = counts.setdefault(r["chapter_id"], {b: 0 for b in BLOOM})
        c[r["bloom"]] += r["n"]
    by_type = {t: 0 for t in QTYPES}
    for r in q("SELECT qtype, COUNT(*) AS n FROM questions GROUP BY qtype", fetch=True):
        by_type[r["qtype"]] = r["n"]
    rows = q(f"""SELECT ch.id, ch.code, ch.title, t.id AS textbook_id, t.name AS textbook,
                 co.id AS course_id, co.name AS course, ({CHAPTER_LABEL}) AS label
                 FROM chapters ch JOIN textbooks t ON t.id=ch.textbook_id JOIN courses co ON co.id=t.course_id
                 ORDER BY co.name, t.name, ch.code""", fetch=True)
    chapters = []
    for r in rows:
        r["counts"] = counts.get(r["id"], {b: 0 for b in BLOOM})
        r["total"] = sum(r["counts"].values())
        chapters.append(r)
    return {"chapters": chapters, "blooms": BLOOM, "by_type": by_type,
            "total": sum(by_type.values()), "tree": tree}


def backup_payload():
    rows = q(QUESTION_SQL + " ORDER BY qs.id", fetch=True)
    items = []
    for r in rows:
        r = row_to_api(r)
        items.append({"course": r["course"], "textbook": r["textbook"], "chapter_code": r["chapter_code"],
                      "chapter_title": r["chapter_title"], "bloom": r["bloom"], "qtype": r["qtype"],
                      "content": r["content"], "options": r["options"], "answer": r["answer"],
                      "solution": r["solution"], "passage": r["passage"]})
    return {"app": "QBank", "version": 2, "questions": items}


BACKUP_DIR = os.path.join(BASE, "data", "backup")
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
    Lọc nhanh bằng độ dài và quick_ratio trước khi tính ratio đầy đủ."""
    from difflib import SequenceMatcher
    rows = q(QUESTION_SQL + " ORDER BY qs.id", fetch=True)
    keys = {}
    for r in rows:
        k = dup_key(r["content"], json.loads(r["options"]))
        keys[r["id"]] = k
    by_type = {}
    for r in rows:
        by_type.setdefault(r["qtype"], []).append(r)
    pairs = []
    for items in by_type.values():
        for i, a in enumerate(items):
            ka = keys[a["id"]]
            for b in items[i + 1:]:
                kb = keys[b["id"]]
                if ka == kb or not ka or not kb:
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
    rows = q(f"SELECT id, content, options FROM questions WHERE id IN ({','.join('?' * len(ids))})", ids, fetch=True)
    if len(rows) != len(set(ids)):
        raise ValueError("Có câu không còn trong ngân hàng, hãy tải lại trang")
    if chapter_id is not None:
        index, moving = dup_index(), set(ids)
        for r in rows:
            key = dup_key(r["content"], json.loads(r["options"]))
            clash = [h for h in index.get(key, []) if h[1] == int(chapter_id) and h[0] not in moving]
            if clash:
                raise ValueError(f"Câu #{r['id']} sẽ trùng với câu #{clash[0][0]} trong chương đích; không đổi câu nào")
    with closing(sqlite3.connect(DB_PATH)) as conn:
        for i in ids:
            if chapter_id is not None:
                conn.execute("UPDATE questions SET chapter_id=? WHERE id=?", (int(chapter_id), i))
            if bloom is not None:
                conn.execute("UPDATE questions SET bloom=? WHERE id=?", (int(bloom), i))
        conn.commit()
    return {"updated": len(ids)}


def restore_item(item, index, allow_dup=False):
    """Khôi phục một câu từ file sao lưu; tự tạo học phần/giáo trình/chương nếu chưa có.
    Câu trùng trong cùng chương bị bỏ qua (báo lỗi để người dùng biết)."""
    chapter_id = ensure_chapter(str(item.get("course", DEFAULT_COURSE)).strip() or DEFAULT_COURSE,
                                str(item.get("textbook", DEFAULT_COURSE)).strip() or DEFAULT_COURSE,
                                str(item.get("chapter_code", "Chung")).strip() or "Chung",
                                str(item.get("chapter_title", "")))
    payload = clean_question(dict(item, chapter_id=chapter_id))
    same, _, key = check_duplicate(index, payload)
    if same and not allow_dup:
        raise ValueError(f"trùng câu #{same[0][0]} trong cùng chương (bỏ qua)")
    new_id = insert_q(payload, keep_code=str(item.get("code") or "").strip() or None)
    index.setdefault(key, []).append((new_id, chapter_id, payload["content"]))


# ---------------------------------------------------------------- HTTP
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

    def _guard(self, fn):
        try:
            fn()
        except ValueError as e:
            self._json(400, {"error": str(e)})
        except sqlite3.IntegrityError as e:
            msg = "Đã tồn tại (trùng tên/mã trong cùng cấp)" if "UNIQUE" in str(e) else "Dữ liệu bị ràng buộc"
            self._json(400, {"error": msg})
        except Exception as e:  # lỗi bất ngờ: trả về thông báo, không làm sập server
            self._json(500, {"error": f"{type(e).__name__}: {e}"})

    def do_GET(self):
        self._guard(self._get)

    def do_POST(self):
        self._guard(self._post)

    def do_PUT(self):
        self._guard(self._put)

    def do_DELETE(self):
        self._guard(self._delete)

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
        if p == "/api/questions":
            sql, args = QUESTION_SQL + " WHERE 1=1", []
            if arg("chapter_id"):
                sql += " AND qs.chapter_id=?"; args.append(int(arg("chapter_id")))
            if arg("textbook_id"):
                sql += " AND t.id=?"; args.append(int(arg("textbook_id")))
            if arg("course_id"):
                sql += " AND co.id=?"; args.append(int(arg("course_id")))
            if arg("bloom"):
                sql += " AND qs.bloom=?"; args.append(int(arg("bloom")))
            if arg("qtype"):
                sql += " AND qs.qtype=?"; args.append(arg("qtype"))
            kw = arg("q")
            if kw:
                sql += " AND (qs.content LIKE ? OR qs.solution LIKE ? OR qs.options LIKE ? OR qs.answer LIKE ? OR qs.code LIKE ?)"
                args += [f"%{kw}%"] * 5
            sql += " ORDER BY co.name, t.name, ch.code, qs.bloom, qs.id"
            return self._json(200, [row_to_api(r) for r in q(sql, args, fetch=True)])
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
            rows = q("SELECT codes, data FROM exam_exports WHERE id=?", (int(fm.group(1)),), fetch=True)
            if not rows:
                return self._json(404, {"error": "không tìm thấy bản xuất"})
            fname = "de_thi_" + re.sub(r"[^\w.-]+", "_", rows[0]["codes"].split(",")[0].strip() or "x") + ".zip"
            return self._send(200, rows[0]["data"], "application/zip",
                              {"Content-Disposition": f'attachment; filename="{fname}"'})
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
        if p == "/api/questions":
            payload = clean_question(self._json_body())
            same, other, _ = check_duplicate(dup_index(), payload)
            if same:  # đã có câu y hệt trong cùng chương: không lưu, báo người dùng
                raise ValueError(f"Câu này đã có trong cùng chương (#{same[0][0]}). Hãy sửa nội dung nếu muốn lưu thêm.")
            new_id = insert_q(payload)
            warnings = [f"Trùng nội dung với câu #{h[0]} ở chương khác" for h in other[:3]]
            return self._json(201, {"ok": True, "id": new_id, "warnings": warnings})
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
        if p == "/api/textbooks":
            b = self._json_body()
            name = str(b.get("name", "")).strip()
            if not name or not b.get("course_id"):
                raise ValueError("Thiếu học phần hoặc tên giáo trình")
            q("INSERT INTO textbooks(course_id, name, code) VALUES(?, ?, ?)",
              (int(b["course_id"]), name, str(b.get("code", "")).strip()))
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
                q("UPDATE textbooks SET name=?, code=? WHERE id=?", (name, str(b.get("code", "")).strip(), rid))
            else:
                code = str(b.get("code", "")).strip()
                if not code:
                    raise ValueError("Thiếu mã chương")
                q("UPDATE chapters SET code=?, title=? WHERE id=?", (code, str(b.get("title", "")).strip(), rid))
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
        payload = clean_question(self._json_body())
        payload["id"] = int(m.group(1))
        q("""UPDATE questions SET chapter_id=:chapter_id, bloom=:bloom, qtype=:qtype, content=:content,
             options=:options, answer=:answer, solution=:solution, passage=:passage WHERE id=:id""", payload)
        self._json(200, {"ok": True})

    # --- DELETE: chỉ xóa khi không còn dữ liệu bên trong (khóa ngoài RESTRICT)
    def _delete(self):
        path = urlparse(self.path).path
        cfg = re.fullmatch(r"/api/configs/(\d+)", path)
        if cfg:
            q("DELETE FROM exam_configs WHERE id=?", (int(cfg.group(1)),))
            return self._json(200, {"ok": True})
        ex = re.fullmatch(r"/api/exports/(\d+)", path)
        if ex:
            q("DELETE FROM exam_exports WHERE id=?", (int(ex.group(1)),))
            return self._json(200, {"ok": True})
        m = re.fullmatch(r"/api/(questions|courses|textbooks|chapters)/(\d+)", path)
        if not m:
            return self._json(404, {"error": "không tìm thấy"})
        table = {"questions": "questions", "courses": "courses",
                 "textbooks": "textbooks", "chapters": "chapters"}[m.group(1)]
        try:
            if table == "questions":
                q("DELETE FROM exam_usage WHERE question_id=?", (int(m.group(2)),))
                q("DELETE FROM question_results WHERE question_id=?", (int(m.group(2)),))
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
        prefix = str(body.get("prefix", ""))
        answer_key = bool(body.get("answer_key", True))
        key_separate = bool(body.get("key_separate", False))
        sections = body.get("sections") or [{"title": "", "qtype": None,
                                             "shuffle": body.get("shuffle", True),
                                             "matrix": body.get("matrix", [])}]
        furigana = "paren" if body.get("furigana") == "paren" else "ruby"
        key_name = os.path.basename(body.get("key_template") or "dap_an_chuan.docx")
        key_tpl_path = os.path.join(TPL_DIR, key_name)
        usage = usage_map()
        taken = set()
        picked = body.get("picked")  # câu đã xem trước và đổi tay, aligned theo sections
        parts = []  # (tiêu đề, danh sách câu gốc, có xáo)
        for idx, s in enumerate(sections):
            if not any(int(c.get("n", 0)) > 0 for c in s.get("matrix", [])):
                continue  # phần không có số câu thì bỏ qua
            ids = picked[idx] if picked and idx < len(picked) else None
            if ids:
                want = sum(int(c.get("n", 0)) for c in s.get("matrix", []))
                if len(ids) != want or len(set(ids)) != len(ids):
                    raise ValueError("Bản xem trước không khớp số câu; hãy xem trước lại")
                base = load_items(ids, usage)
                taken.update(int(i) for i in ids)
            else:
                base = pick_base_exam(s.get("matrix", []), s.get("qtype") or None, usage, taken)
            parts.append((str(s.get("title", "")).strip(), base, bool(s.get("shuffle", True))))
        if not parts:
            raise ValueError("Chưa có phần nào có số câu")
        reused = sum(1 for _, base, _ in parts for it in base if it.get("last_used"))
        exam_name = str(body.get("meta", {}).get("ten_de", ""))
        now = datetime.now().isoformat(timespec="seconds")
        rng = random.Random()
        out = io.BytesIO()
        usage_rows = []
        codes = []
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
            for i in range(variants):
                code = f"{prefix}{start + i:03d}"
                codes.append(code)
                exam = [(t, make_variant(b, sh, rng)) for t, b, sh in parts]
                data = render_docx(tpl_path, body.get("meta", {}), exam, code,
                                   answer_key and not key_separate, furigana)
                z.writestr(f"De_{code}.docx", data)
                if key_separate:
                    if os.path.isfile(key_tpl_path):
                        z.writestr(f"DapAn_{code}.docx", render_key_doc(key_tpl_path, body.get("meta", {}), code, exam))
                    else:
                        z.writestr(f"DapAn_{code}.docx", render_key(code, exam))
                if body.get("record_usage", True):
                    usage_rows += [(it["id"], code, exam_name, now) for _, _, it in flatten(exam)]
        record_usage(usage_rows)
        data = out.getvalue()
        if body.get("save_history", True):  # lưu bản xuất để tải lại hoặc in lại đáp án
            q("INSERT INTO exam_exports(created_at, name, codes, config, data) VALUES(?,?,?,?,?)",
              (now, exam_name, ", ".join(codes),
               json.dumps({"meta": body.get("meta", {}), "variants": variants, "start": start,
                           "prefix": prefix, "template": tpl}, ensure_ascii=False), data))
        self._send(200, data, "application/zip",
                   {"X-Reused": str(reused), "X-Variants": str(variants)})


def main():
    init_storage()
    backup_loop()
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"QBank đang chạy tại http://127.0.0.1:{PORT}  (Ctrl+C để dừng)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        sys.exit(0)


if __name__ == "__main__":
    main()
