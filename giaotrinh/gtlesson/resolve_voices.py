"""Đổi tên giọng JVNV (jvnv-F1…) trong roles.aivis.yaml thành mã style của AivisSpeech.

Đầu vào: file JSON do textts in ra:  python tts.py --list-voices aivis > voices.json
Mỗi mục có dạng {"value": "<mã style>", "label": "<tên giọng> (<tên style>)"}.
Chỉ lấy style ノーマル của giọng khớp tên.

    python -m gtlesson.resolve_voices out/textts/roles.aivis.yaml voices.json -o out/textts/roles.aivis.resolved.yaml
"""
from __future__ import annotations

import argparse
import json
import re
import sys

import yaml

NORMAL_STYLE = "ノーマル"


class VoiceError(ValueError):
    pass


def find_style(voice: str, voices: list[dict]) -> str:
    """Mã style ノーマル của giọng `voice` (vd 'jvnv-F1'). Không thấy thì báo lỗi kèm danh sách có sẵn."""
    key = voice.lower()
    short = key.split("-", 1)[-1]  # 'F1'
    token = re.compile(rf"(?<![a-z0-9]){re.escape(short)}(?![a-z0-9])")
    hits = []
    for v in voices:
        label = str(v.get("label", ""))
        low = label.lower()
        if NORMAL_STYLE not in label:
            continue
        if key in low or ("jvnv" in low and token.search(low)):
            hits.append(str(v["value"]))
    if not hits:
        available = ", ".join(sorted({str(v.get("label", "")) for v in voices})) or "(không có giọng nào)"
        raise VoiceError(f"Không tìm thấy giọng '{voice}' (style ノーマル). Giọng đang có: {available}")
    if len(set(hits)) > 1:
        raise VoiceError(f"Giọng '{voice}' khớp nhiều mục: {sorted(set(hits))}. Hãy xóa model trùng trong AivisSpeech.")
    return hits[0]


def resolve(roles_doc: dict, voices: list[dict]) -> dict:
    out = dict(roles_doc)
    out["roles"] = {}
    for rid, role in (roles_doc.get("roles") or {}).items():
        new = dict(role)
        voice = str(role.get("voice", ""))
        if voice.startswith("jvnv-"):
            new["voice"] = find_style(voice, voices)
        out["roles"][rid] = new
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="gtlesson.resolve_voices", description="Đổi tên giọng JVNV thành mã style AivisSpeech")
    ap.add_argument("roles", help="roles.aivis.yaml do gtlesson sinh ra")
    ap.add_argument("voices", help="JSON từ: python tts.py --list-voices aivis")
    ap.add_argument("-o", "--out", required=True, help="file roles đã điền mã style")
    a = ap.parse_args(argv)
    with open(a.roles, encoding="utf-8") as f:
        doc = yaml.safe_load(f)
    with open(a.voices, encoding="utf-8") as f:
        voices = json.load(f)
    try:
        resolved = resolve(doc, voices)
    except VoiceError as e:
        print(str(e), file=sys.stderr)
        return 1
    with open(a.out, "w", encoding="utf-8") as f:
        yaml.safe_dump(resolved, f, allow_unicode=True, sort_keys=False)
    print(f"-> {a.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
