"""lesson.json -> kịch bản textts (.txt) cho hội thoại và phần nghe.

Mỗi nhân vật trong trang Nhân vật có một số vai cố định cho cả bộ sách (vai 1 = nhân vật đầu bảng…),
nên giọng không đổi giữa các bài. Vai 0 là người dẫn (đọc lời thoại phần nghe ở phụ lục).
"""
from __future__ import annotations

import os
import re

from .furigana import plain, to_textts

NARRATOR = 0


def _norm(name: str) -> str:
    """Bỏ furigana, đổi katakana sang hiragana: マキ và まき là một người."""
    return "".join(chr(ord(c) - 0x60) if "ァ" <= c <= "ヶ" else c for c in plain(name).strip())


def role_map(book: dict) -> dict[str, int]:
    """Tên nhân vật (không furigana, bỏ họ) -> số vai."""
    roles: dict[str, int] = {}
    for i, ch in enumerate(book.get("characters", []), start=1):
        name = _norm(ch["name"]).replace("　", " ")
        roles[name] = i
        for part in name.split():
            roles.setdefault(part, i)
    return roles


def speaker_role(speaker: str, roles: dict[str, int]) -> int | None:
    sp = _norm(speaker)
    if sp in roles:
        return roles[sp]
    for name, n in roles.items():
        if sp and (sp in name or name in sp):
            return n
    return None


def dialogue_script(lesson: dict, dlg: dict, roles: dict[str, int]) -> str:
    head = f"# Bài {lesson['no']} – hội thoại {dlg.get('part') or ''} ♪{dlg.get('audio') or ''}\n"
    head += f"# {plain(dlg.get('situation', ''))}\n"
    out = [head]
    for line in dlg["lines"]:
        n = speaker_role(line["speaker"], roles)
        if n is None:
            out.append(f"# CHƯA GÁN VAI: {line['speaker']}\n")
            n = NARRATOR
        out.append(f"[{n}] {to_textts(line['text'])}\n")
    return "".join(out)


def script_text(script: dict) -> str:
    """Lời thoại phần nghe: mỗi ý (1) (2)… thành một câu của người dẫn, cách nhau 1,5 giây."""
    parts = [p.strip() for p in re.split(r"\(\d+\)", script["text"]) if p.strip()]
    lines = [f"# Bài {script.get('lesson')} – {plain(script.get('ref', ''))} ♪{script.get('audio')}\n"]
    for i, p in enumerate(parts):
        if i:
            lines.append("[pause=1500]\n")
        lines.append(f"[{NARRATOR}] {to_textts(p)}\n")
    return "".join(lines)


def roles_yaml(book: dict) -> str:
    """Khung roles.yaml: tên vai và mô tả giọng; giọng cụ thể giáo viên tự chọn trong textts."""
    out = ["# Sinh từ trang Nhân vật. Điền voice theo engine đang dùng.\nroles:\n",
           f'  "{NARRATOR}":\n    name: ナレーター\n']
    for i, ch in enumerate(book.get("characters", []), start=1):
        out.append(f'  "{i}":\n    name: {plain(ch["name"])}\n    # giọng: {ch.get("voice_hint", "")}\n')
    return "".join(out)


def write_all(book: dict, out_dir: str) -> list[str]:
    os.makedirs(out_dir, exist_ok=True)
    roles = role_map(book)
    written = []

    def put(name: str, text: str) -> None:
        path = os.path.join(out_dir, name)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        written.append(path)

    put("roles.yaml", roles_yaml(book))
    for lesson in book["lessons"]:
        for k, dlg in enumerate(lesson["dialogues"], start=1):
            name = dlg.get("audio") or f"{lesson['no']}-kaiwa{k}"
            put(f"{name}.txt", dialogue_script(lesson, dlg, roles))
    for sc in book.get("audio_scripts", []):
        if sc.get("audio") and sc.get("text"):
            put(f"{sc['audio']}.txt", script_text(sc))
    return written
