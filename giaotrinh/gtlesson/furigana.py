"""Cú pháp furigana chung: {漢字|かな}.

Đây là dạng chuẩn trong lesson.json (cũng là cú pháp của qbank và dominion).
Các hàm ở đây chuyển sang dạng mà từng công cụ cần.
"""
from __future__ import annotations

import re

RUBY = re.compile(r"\{([^{}|]+)\|([^{}|]+)\}")
KANJI = r"一-鿿々〆ヶ"
KANA = r"぀-ゟ゠-ヿー"
_ALL_KANJI = re.compile(rf"^[{KANJI}]+$")


def plain(text: str) -> str:
    """Bỏ furigana, giữ chữ gốc: {物産展|ぶっさんてん} -> 物産展."""
    return RUBY.sub(r"\1", text)


def kana(text: str) -> str:
    """Thay chữ có furigana bằng cách đọc: {物産展|ぶっさんてん} -> ぶっさんてん."""
    return RUBY.sub(r"\2", text)


def _split_okurigana(base: str, reading: str) -> list[tuple[str, str]] | None:
    """Tách từ lẫn kanji và kana thành từng cụm kanji kèm cách đọc.

    持ち主 / もちぬし -> [(持, も), (ち, ''), (主, ぬし)]. Trả None nếu không tách được.
    """
    parts = re.findall(rf"[{KANJI}]+|[^{KANJI}]+", base)
    pattern = "".join(
        "(.+?)" if re.match(rf"[{KANJI}]", p) else "(" + re.escape(p) + ")" for p in parts
    )
    m = re.fullmatch(pattern, reading)
    if not m:
        return None
    return [(p, g if re.match(rf"[{KANJI}]", p) else "") for p, g in zip(parts, m.groups())]


def to_textts(text: str) -> str:
    """Sang cú pháp textts 漢字{かな}; textts chỉ nhận furigana gắn sau chữ Hán."""

    def repl(m: re.Match) -> str:
        base, reading = m.group(1), m.group(2)
        if _ALL_KANJI.match(base):
            return f"{base}{{{reading}}}"
        parts = _split_okurigana(base, reading)
        if parts is None:
            return base
        return "".join(f"{b}{{{r}}}" if r else b for b, r in parts)

    return RUBY.sub(repl, text)


def ruby_count(text: str) -> int:
    return len(RUBY.findall(text))
