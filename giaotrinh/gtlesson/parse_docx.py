"""Đọc bản thảo giáo trình (.docx dùng các style GT_*) thành dict theo định dạng lesson.json.

Dựa vào style đoạn văn, không dựa vào vị trí trang:
GT_LessonTitle, GT_Section, GT_SubHead, GT_LessonGoal, GT_Dialogue, GT_Body,
GT_Example, GT_Note, GT_Instruction, GT_Exercise. Furigana (w:ruby) được giữ ở dạng {漢字|かな}.
Nội dung không nhận ra được giữ trong "unparsed" để không mất chữ.
"""
from __future__ import annotations

import re
from typing import Any

import docx
from docx.oxml.ns import qn

SCHEMA = "imeji-lesson/1"

SECTION_KEYS = {
    "かいわ": "kaiwa",
    "ことば": "kotoba",
    "ぶんぽう": "bunpou",
    "ひょうげんノート": "hyougen",
    "れんしゅう": "renshuu",
    "よみかき": "yomikaki",
    "まとめのれんしゅう": "matome",
    "ぶんか": "bunka",
    "じこひょうか": "jikohyouka",
}
LEVELS = {"①": 1, "②": 2, "③": 3}
ROMAN = "ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ"
AUDIO = re.compile(r"\s*♪\s*(\d+-\d+)\s*")
BLANK = re.compile(r"＿{2,}")
VI_START = re.compile(r"[　\s]+(?=[A-Za-zÀ-ỹĐđ])")


# ---------------------------------------------------------------- đọc chữ


def _is_in_ruby(el) -> bool:
    a = el.getparent()
    while a is not None:
        if a.tag == qn("w:ruby"):
            return True
        a = a.getparent()
    return False


def el_text(el) -> str:
    """Chữ của một đoạn (w:p) hoặc ô, ruby viết thành {漢字|かな}."""
    out: list[str] = []
    for node in el.iter():
        if node.tag == qn("w:ruby"):
            base = "".join(t.text or "" for t in node.find(qn("w:rubyBase")).iter(qn("w:t")))
            rt = "".join(t.text or "" for t in node.find(qn("w:rt")).iter(qn("w:t")))
            out.append(f"{{{base}|{rt}}}" if rt.strip() else base)
        elif node.tag == qn("w:t") and not _is_in_ruby(node):
            out.append(node.text or "")
        elif node.tag == qn("w:tab") and not _is_in_ruby(node):
            out.append("\t")
        elif node.tag in (qn("w:br"), qn("w:cr")) and not _is_in_ruby(node):
            out.append("\n")
    return "".join(out)


def cell_text(cell) -> str:
    return "\n".join(t for t in (el_text(p) for p in cell.findall(qn("w:p"))) if t.strip())


def table_rows(tbl) -> list[list[str]]:
    rows = []
    for tr in tbl.findall(qn("w:tr")):
        rows.append([cell_text(tc) for tc in tr.findall(qn("w:tc"))])
    return rows


def roman_no(r: str) -> int:
    return ROMAN.index(r) + 1


def take_audio(text: str) -> tuple[str, str | None]:
    m = AUDIO.search(text)
    if not m:
        return text.strip(), None
    return (text[: m.start()] + " " + text[m.end():]).strip(), m.group(1)


def split_ja_vi(text: str) -> tuple[str, str]:
    """'北海道の…　Truyền đạt…' -> (phần Nhật, phần Việt)."""
    m = VI_START.search(text)
    if not m:
        return text.strip(), ""
    return text[: m.start()].strip(), text[m.end():].strip()


def clean_blank(text: str) -> str:
    """Rút gọn chỗ trống thành ＿＿; bỏ ô trả lời ở cuối câu ('→ ＿＿') vì giao diện tự thêm ô nhập."""
    text = BLANK.sub("＿＿", text).strip()
    return re.sub(r"\s*→?\s*＿＿$", "", text).strip()


# ---------------------------------------------------------------- duyệt tài liệu


def iter_blocks(document):
    """Sinh ('p', style, text) hoặc ('t', None, rows) theo thứ tự trong thân tài liệu."""
    styles = {s.style_id: s.name for s in document.styles}
    for ch in document.element.body.iterchildren():
        if ch.tag == qn("w:p"):
            ps = ch.find(qn("w:pPr"))
            sid = None
            if ps is not None and ps.find(qn("w:pStyle")) is not None:
                sid = ps.find(qn("w:pStyle")).get(qn("w:val"))
            text = el_text(ch)
            if text.strip():
                yield "p", styles.get(sid, sid or ""), text
        elif ch.tag == qn("w:tbl"):
            yield "t", None, table_rows(ch)


def _new_lesson(no: int, title: str) -> dict[str, Any]:
    ja, vi = split_ja_vi(title)
    return {
        "no": no,
        "title_ja": ja,
        "title_vi": vi,
        "goals": [],
        "dialogues": [],
        "vocab": [],
        "grammar": [],
        "notes": [],
        "exercises": [],
        "kanji": [],
        "lesson_words": [],
        "reading": None,
        "writing": None,
        "culture": None,
        "self_check": [],
        "media": [],
        "unparsed": [],
    }


class Parser:
    def __init__(self) -> None:
        self.book: dict[str, Any] = {
            "schema": SCHEMA,
            "characters": [],
            "lessons": [],
            "audio_scripts": [],
        }
        self.mode = "front"  # front | characters | lesson | appendix
        self.lesson: dict[str, Any] | None = None
        self.section: str | None = None
        self.appendix: str | None = None
        self.part: str | None = None  # Ⅰ, Ⅱ… trong れんしゅう / よみかき
        self.part_title = ""
        self.cur: dict[str, Any] | None = None  # bài tập / mục ngữ pháp / hội thoại đang mở
        self.banner_no: int | None = None
        self.answers: dict[int, dict[str, list[str]]] = {}
        self.answer_lesson: int | None = None
        self.script_ref: dict[str, Any] | None = None

    # --- tiện ích
    def _keep(self, kind: str, data: Any) -> None:
        if self.lesson is not None:
            self.lesson["unparsed"].append({"section": self.section, kind: data})

    def _exercise(self, label: str, instruction: str) -> dict[str, Any]:
        instruction, audio = take_audio(instruction)
        level = None
        for mark, lv in LEVELS.items():
            if mark in instruction:
                level = lv
                instruction = re.sub(rf"\s*{mark}\s*[^　\s]*(\s[^　\s]+)?\s*$", "", instruction).strip()
        ex_id = f"{self.section}.{roman_no(self.part)}.{label}" if self.part else f"{self.section}.{label}"
        ex = {
            "id": ex_id,
            "section": self.section,
            "part": self.part,
            "part_title": self.part_title,
            "label": label,
            "instruction": instruction,
            "level": level,
            "audio": audio,
            "example": [],
            "passage": None,
            "items": [],
            "kind": "short_answer",
        }
        self.lesson["exercises"].append(ex)
        self.cur = ex
        return ex

    # --- xử lý từng khối
    def para(self, style: str, text: str) -> None:
        s = style.replace(" ", "_")
        if s == "GT_LessonTitle":
            return self._title(text)
        if self.mode == "characters" or self.mode == "front":
            return
        if self.mode == "appendix":
            return self._appendix_para(s, text)
        if self.lesson is None:
            return
        if s == "GT_Section":
            return self._section(text)
        if s == "GT_SubHead" and text.strip() == "Mục tiêu bài học":
            return
        if s == "GT_LessonGoal":
            self.lesson["goals"].append(text.lstrip("◆ ").strip())
            return
        handler = getattr(self, f"_p_{self.section}", None)
        if handler is None or handler(s, text) is False:
            self._keep("text", {"style": style, "text": text})

    def table(self, rows: list[list[str]]) -> None:
        if not rows:
            return
        if len(rows) == 1 and rows[0] and re.fullmatch(r"第\d+課", rows[0][0].strip()):
            self.banner_no = int(re.search(r"\d+", rows[0][0]).group())
            return
        if self.mode == "characters":
            return self._characters(rows)
        if self.mode != "lesson" or self.lesson is None:
            return
        flat = "\n".join(c for c in rows[0] if c)
        if self.section is None:
            if flat.startswith("[HÌNH") or flat.startswith("♪"):
                self.lesson["media"].append({"kind": "image" if flat.startswith("[HÌNH") else "audio_note",
                                             "text": "\n".join(" ".join(r) for r in rows).strip()})
                return
        handler = getattr(self, f"_t_{self.section}", None)
        if handler is None or handler(rows) is False:
            self._keep("table", rows)

    # --- tiêu đề, mục
    def _title(self, text: str) -> None:
        self.section, self.part, self.cur = None, None, None
        if "Nhân vật" in text or "じんぶつ" in text:
            self.mode = "characters"
        elif "付録" in text or "Phụ lục" in text:
            self.mode, self.lesson = "appendix", None
        elif "目次" in text or "Mục lục" in text:
            self.mode = "front"
        else:
            self.mode = "lesson"
            no = self.banner_no or len(self.book["lessons"]) + 1
            self.lesson = _new_lesson(no, text)
            self.book["lessons"].append(self.lesson)
            self.banner_no = None

    def _section(self, text: str) -> None:
        self.part, self.cur, self.part_title = None, None, ""
        body = re.sub(r"^[0-9A-Z]+\s+", "", text.strip())
        name = body.split("　")[0].strip()
        self.section = SECTION_KEYS.get(name, name)

    def _characters(self, rows: list[list[str]]) -> None:
        if len(rows) < 2 or "Nhân vật" not in "".join(rows[0]):
            return
        head = rows[0]
        i_name = next(i for i, h in enumerate(head) if "Nhân vật" in h)
        i_prof = next((i for i, h in enumerate(head) if "Hồ sơ" in h), None)
        i_voice = next((i for i, h in enumerate(head) if "Giọng" in h), None)
        for r in rows[1:]:
            if not r[i_name].strip():
                continue
            self.book["characters"].append({
                "name": r[i_name].strip(),
                "profile": r[i_prof].strip() if i_prof is not None else "",
                "voice_hint": r[i_voice].strip() if i_voice is not None else "",
            })

    # --- 1 かいわ
    def _p_kaiwa(self, s: str, text: str):
        if s == "GT_SubHead":
            desc, audio = take_audio(text)
            m = re.match(rf"^([{ROMAN}])[　\s]+(.*)$", desc)
            part, desc = (m.group(1), m.group(2)) if m else (None, desc)
            self.cur = {"part": part, "situation": desc, "audio": audio, "lines": []}
            self.lesson["dialogues"].append(self.cur)
            return
        if s == "GT_Dialogue" and self.cur is not None:
            cols = text.split("\t")
            if len(cols) >= 3:
                n, speaker, line = cols[0].strip(), cols[1].strip().rstrip("：:"), "\t".join(cols[2:]).strip()
            elif len(cols) == 2:
                n, speaker, line = cols[0].strip(), "", cols[1].strip()
            else:
                n, speaker, line = "", "", text.strip()
            if not speaker and self.cur["lines"]:
                speaker = self.cur["lines"][-1]["speaker"]
            self.cur["lines"].append({"n": int(n) if n.isdigit() else None, "speaker": speaker, "text": line})
            return
        return False

    # --- 2 ことば
    def _p_kotoba(self, s: str, text: str):
        if s == "GT_Body":
            _, audio = take_audio(text)
            if audio:
                self.lesson["vocab_audio"] = audio
            return
        return False

    def _t_kotoba(self, rows):
        head = rows[0]
        if "ことば" not in head:
            return False
        iw = head.index("ことば")
        ig = next((i for i, h in enumerate(head) if "Nhóm" in h), None)
        im = next((i for i, h in enumerate(head) if "Nghĩa" in h), None)
        ist = next((i for i, h in enumerate(head) if h.strip() == "＊"), None)
        for r in rows[1:]:
            if not r[iw].strip():
                continue
            self.lesson["vocab"].append({
                "word": r[iw].strip(),
                "meaning": r[im].strip() if im is not None else "",
                "group": r[ig].strip() if ig is not None else "",
                "in_dialogue": bool(ist is not None and "＊" in r[ist]),
            })

    # --- 3 ぶんぽう
    def _p_bunpou(self, s: str, text: str):
        if s == "GT_SubHead":
            m = re.match(r"^(\d+)\.[　\s]+(.*)$", text.strip())
            if not m:
                return False
            pattern, _, meaning = m.group(2).partition("　　")
            self.cur = {"no": int(m.group(1)), "pattern": pattern.strip(), "meaning": meaning.strip(),
                        "form": "", "explanation": "", "examples": [], "notes": []}
            self.lesson["grammar"].append(self.cur)
            return
        if self.cur is None:
            return False
        if s == "GT_Body":
            self.cur["explanation"] = (self.cur["explanation"] + "\n" + text.strip()).strip()
        elif s == "GT_Example":
            self.cur["examples"].append(text.strip().lstrip("・").strip())
        elif s == "GT_Note":
            self.cur["notes"].append(text.strip())
        else:
            return False

    def _t_bunpou(self, rows):
        if self.cur is None or len(rows) != 1 or len(rows[0]) != 1:
            return False
        self.cur["form"] = rows[0][0].strip()

    # --- 4 ひょうげんノート
    def _t_hyougen(self, rows):
        self.lesson["notes"].extend(c.strip() for r in rows for c in r if c.strip())

    # --- 5 れんしゅう và 7 まとめ
    def _p_renshuu(self, s: str, text: str):
        if s == "GT_Note" and self.part is None:
            return  # chú thích nhãn bậc chung
        if s == "GT_SubHead":
            m = re.match(rf"^([{ROMAN}])[　\s]+(.*)$", text.strip())
            if not m:
                return False
            self.part, self.part_title, self.cur = m.group(1), m.group(2).strip(), None
            return
        return self._exercise_para(s, text)

    def _p_matome(self, s: str, text: str):
        return self._exercise_para(s, text)

    def _exercise_para(self, s: str, text: str):
        if s == "GT_Instruction":
            m = re.match(r"^([A-Z])[.．][　\s]+(.*)$", text.strip())
            if m:
                self._exercise(m.group(1), m.group(2))
                return
            return False
        ex = self.cur
        if ex is None:
            return False
        if s == "GT_Example":
            ex["example"].append(text.strip())
            return
        if s == "GT_Exercise":
            return self._item(ex, text)
        if s == "GT_Note":
            ex.setdefault("notes", []).append(text.strip())
            return
        return False

    def _item(self, ex: dict, text: str):
        raw = text.strip()
        m = re.match(r"^(?:\((\d+)\)|(\d+)\.)\s*(.*)$", raw, re.S)
        if m:
            n = m.group(1) or m.group(2)
            ex["items"].append({"n": int(n), "prompt": clean_blank(m.group(3).replace("\t", " ")), "answer": None})
            return
        if re.fullmatch(r"[→\s]*＿+", raw):
            if not ex["items"]:
                ex["kind"] = "free"
                ex["lines"] = ex.get("lines", 0) + 1
            return
        if ex["items"]:
            ex["items"][-1]["prompt"] += " " + clean_blank(raw.replace("\t", " "))
            return
        return False

    def _t_renshuu(self, rows):
        return self._exercise_table(rows)

    def _t_matome(self, rows):
        return self._exercise_table(rows)

    def _exercise_table(self, rows):
        ex = self.cur
        if ex is None:
            return False
        if len(rows) == 1 and len(rows[0]) == 1:
            ex["passage"] = rows[0][0].strip()
            return
        if len(rows) > 1 and all(not c.strip() for r in rows[1:] for c in r):
            ex["kind"] = "survey"
            ex["columns"] = [c.strip() for c in rows[0]]
            ex["rows"] = len(rows) - 1
            return
        return False

    # --- 6 よみかき
    def _p_yomikaki(self, s: str, text: str):
        if s == "GT_SubHead":
            body, audio = take_audio(text)
            m = re.match(rf"^([{ROMAN}])[　\s]+(.*)$", body.strip())
            title = m.group(2) if m else body.strip()
            self.part = m.group(1) if m else None
            if "読みもの" in title:
                self.lesson["reading"] = {"title": title, "audio": audio, "text": "", "questions": []}
                self.cur = None
            elif "書く" in title:
                self.lesson["writing"] = {"title": title, "instruction": "", "lines": 0}
                self.cur = None
            elif "漢字" in title:
                self.lesson["kanji_range"] = title
            elif "Từ vựng của bài" in title:
                self.part = "words"
            else:
                return False
            return
        if s == "GT_Note" and self.part == "Ⅰ":
            self.lesson.setdefault("kanji_notes", []).append(text.strip())
            return
        rd, wr = self.lesson["reading"], self.lesson["writing"]
        if wr is not None:
            if s == "GT_Instruction":
                wr["instruction"] = text.strip()
            elif s == "GT_Exercise" and re.fullmatch(r"[＿\s]+", text):
                wr["lines"] += 1
            else:
                return False
            return
        if rd is not None:
            if s == "GT_Instruction":
                return
            if s == "GT_Exercise":
                fake = {"items": rd["questions"], "kind": "qa"}
                return self._item(fake, text)
        return False

    def _t_yomikaki(self, rows):
        head = [c.strip() for c in rows[0]]
        if self.lesson["reading"] is not None and len(rows) == 1 and len(rows[0]) == 1:
            self.lesson["reading"]["text"] = rows[0][0].strip()
            return
        if len(rows) == 1 and len(head) >= 2 and len(head[0]) == 1:
            info = head[1]
            hv = re.match(r"^([^\s(（/]+)\s*[(（]([^)）]*)[)）]", info)
            on = re.search(r"Âm On:\s*([^/\n]*)", info)
            kun = re.search(r"Âm Kun:\s*([^/\n]*)", info)

            def lst(m):
                if not m:
                    return []
                return [x.strip() for x in re.split(r"[,、，]", m.group(1)) if x.strip() and x.strip() != "-"]

            self.lesson["kanji"].append({
                "char": head[0],
                "han_viet": hv.group(1) if hv else "",
                "meaning": hv.group(2) if hv else "",
                "on": lst(on),
                "kun": lst(kun),
                "words": [],
            })
            return
        if head[:1] == ["Từ"]:
            if self.part == "words":
                for r in rows[1:]:
                    if r[0].strip():
                        self.lesson["lesson_words"].append({"word": r[0].strip(), "meaning": r[1].strip()})
                return
            if self.lesson["kanji"]:
                for r in rows[1:]:
                    if r[0].strip():
                        self.lesson["kanji"][-1]["words"].append(
                            {"word": r[0].strip(), "reading": r[1].strip(), "meaning": r[2].strip() if len(r) > 2 else ""})
                return
        return False

    # --- 8 ぶんか, 9 じこひょうか
    def _t_bunka(self, rows):
        text = "\n".join(c.strip() for r in rows for c in r if c.strip())
        title, _, body = text.partition("\n")
        self.lesson["culture"] = {"title": title, "text": body}

    def _p_jikohyouka(self, s: str, text: str):
        return s == "GT_Body" and None

    def _t_jikohyouka(self, rows):
        if "Tôi có thể" not in rows[0][0]:
            return False
        self.lesson["self_check"] = [r[0].strip() for r in rows[1:] if r[0].strip()]

    # --- phụ lục: E lời thoại nghe, F đáp án
    def _appendix_para(self, s: str, text: str) -> None:
        if s == "GT_Section":
            self.appendix = re.match(r"^\s*([A-Z])", text).group(1) if re.match(r"^\s*[A-Z]", text) else None
            return
        if self.appendix == "E":
            if s == "GT_SubHead":
                body, audio = take_audio(text)
                m = re.search(r"第(\d+)課\s*(.*)$", body)
                self.script_ref = {"audio": audio, "lesson": int(m.group(1)) if m else None,
                                   "ref": m.group(2).strip() if m else body, "text": ""}
                self.book["audio_scripts"].append(self.script_ref)
            elif self.script_ref is not None:
                self.script_ref["text"] = (self.script_ref["text"] + "\n" + text.strip()).strip()
        elif self.appendix == "F":
            if s == "GT_SubHead":
                m = re.match(r"^第(\d+)課", text.strip())
                self.answer_lesson = int(m.group(1)) if m else None
            elif self.answer_lesson is not None:
                key, sep, rest = text.partition(":")
                if not sep:
                    key, sep, rest = text.partition("：")
                if sep:
                    self.answers.setdefault(self.answer_lesson, {})[key.strip()] = split_answers(rest)

    # --- gắn đáp án vào bài tập
    def finish(self) -> dict[str, Any]:
        for lesson in self.book["lessons"]:
            keys = self.answers.get(lesson["no"], {})
            for key, answers in keys.items():
                target = answer_target(lesson, key)
                if target is None:
                    lesson["unparsed"].append({"section": "answers", "text": f"{key}: {answers}"})
                    continue
                for i, a in enumerate(answers):
                    if i < len(target):
                        target[i]["answer"] = a
            for ex in lesson["exercises"]:
                if ex["kind"] == "short_answer" and not ex["items"]:
                    ex["kind"] = "free" if ex.get("lines") else "task"
            if not lesson["unparsed"]:
                del lesson["unparsed"]
        return self.book


def split_answers(text: str) -> list[str]:
    """'(1) A。　(2) B。' hoặc '1. A　2. B' -> ['A。', 'B']."""
    parts = re.split(r"(?:^|(?<=[\s　]))(?:\(\d+\)|\d+\.)\s*", text.strip())
    return [p.strip() for p in parts if p.strip()]


def answer_target(lesson: dict, key: str) -> list[dict] | None:
    key = key.strip()
    m = re.match(rf"^れんしゅう\s*([{ROMAN}])-([A-Z])$", key)
    if m:
        ex_id = f"renshuu.{roman_no(m.group(1))}.{m.group(2)}"
    elif key.startswith("よみかき") and lesson.get("reading"):
        return lesson["reading"]["questions"]
    elif (m := re.match(r"^まとめのれんしゅう\s*([A-Z])$", key)):
        ex_id = f"matome.{m.group(1)}"
    else:
        return None
    ex = next((e for e in lesson["exercises"] if e["id"] == ex_id), None)
    return ex["items"] if ex else None


def parse(path: str) -> dict[str, Any]:
    document = docx.Document(path)
    p = Parser()
    for kind, style, data in iter_blocks(document):
        if kind == "p":
            p.para(style, data)
        else:
            p.table(data)
    return p.finish()
