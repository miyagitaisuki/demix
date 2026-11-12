"""python -m gtlesson GIAOTRINH.docx -o out/ [--textts]"""
from __future__ import annotations

import argparse
import json
import os
import sys

from .parse_docx import parse


def summary(book: dict) -> str:
    lines = [f"{len(book['characters'])} nhân vật, {len(book['lessons'])} bài, "
             f"{len(book['audio_scripts'])} lời thoại phần nghe"]
    for L in book["lessons"]:
        items = [i for e in L["exercises"] for i in e["items"]]
        if L.get("reading"):
            items += L["reading"]["questions"]
        missing = sum(1 for i in items if not i.get("answer"))
        lines.append(
            f"  Bài {L['no']}: {len(L['dialogues'])} hội thoại, {len(L['vocab'])} từ, {len(L['grammar'])} ngữ pháp, "
            f"{len(L['exercises'])} bài tập ({len(items)} câu, {missing} thiếu đáp án), {len(L['kanji'])} chữ Hán"
            + (f", {len(L['unparsed'])} khối chưa đọc được" if L.get("unparsed") else "")
        )
    return "\n".join(lines)


def main(argv=None) -> int:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except AttributeError:
            pass
    ap = argparse.ArgumentParser(prog="gtlesson", description="Bản thảo giáo trình (.docx) -> lesson.json")
    ap.add_argument("docx")
    ap.add_argument("-o", "--out", default="out")
    ap.add_argument("--textts", action="store_true", help="xuất thêm kịch bản textts vào <out>/textts")
    a = ap.parse_args(argv)
    book = parse(a.docx)
    os.makedirs(a.out, exist_ok=True)
    path = os.path.join(a.out, "lesson.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(book, f, ensure_ascii=False, indent=1)
    print(summary(book))
    print(f"-> {path}")
    if a.textts:
        from .export_textts import write_all

        files = write_all(book, os.path.join(a.out, "textts"))
        print(f"-> {len(files)} file textts trong {os.path.join(a.out, 'textts')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
