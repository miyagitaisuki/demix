"""Programmatic Tool Calling (PTC) client dùng chung cho QBank trên Antigravity.

Tái sử dụng trực tiếp các tool handler và hàm gọi REST API từ ``mcp_server.py``
để giữ nguyên toàn bộ kiểm tra nghiệp vụ, ``dry_run=True`` mặc định và quy tắc
``audio_engine="aivis"`` mà không trùng lặp code.
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Iterator

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

import mcp_server  # noqa: E402
from mcp_server import ApiError  # noqa: F401

# Các quy tắc kiểm định cứng theo chuẩn biên soạn JLPT
LINGUISTIC_TERMS = (
    "主語", "目的語", "修飾語", "品詞", "活用形",
    "機能語", "受け手", "話し手", "反対語",
)
AMBIGUOUS_PRONOUNS = ("私", "彼", "彼女", "筆者")

# Khoảng trắng thường hoặc full-width nằm giữa 2 ký tự tiếng Nhật (kana/kanji/chouonpu)
_JP_CHAR = r"[\u3041-\u3096\u30A1-\u30FA\u30FC\u4E00-\u9FFF]"
_RE_JP_SPACE = re.compile(rf"({_JP_CHAR})([ \u3000]+)({_JP_CHAR})")
_RE_RUBY = re.compile(r"\{([^{}]+)\}")
_RE_KANJI = re.compile(r"[\u4E00-\u9FFF]")
_RE_BLANK = re.compile(r"_{3,}|＿{3,}")


class QBankClient:
    """Client Python mỏng cho QBank MCP/REST API phục vụ PTC."""

    def __init__(self, base_url: str | None = None):
        self.base_url = (base_url or os.environ.get("QBANK_URL") or mcp_server.BASE_URL).rstrip("/")
        if self.base_url != mcp_server.BASE_URL:
            mcp_server.BASE_URL = self.base_url

    def call(self, tool_name: str, **kwargs: Any) -> Any:
        """Gọi bất kỳ tool nào trong ``mcp_server.TOOL_BY_NAME``."""
        tool = mcp_server.TOOL_BY_NAME.get(tool_name)
        if not tool:
            raise ApiError(f"Không có công cụ QBank {tool_name!r}")
        return tool["handler"](kwargs)

    # ---- Đọc cấu trúc & ngữ cảnh ----
    def list_tree(self, include_scope: bool = False) -> dict[str, Any]:
        return self.call("list_tree", include_scope=include_scope)

    def get_chapter_context(self, chapter_id: int, sample_limit: int = 5) -> dict[str, Any]:
        return self.call("get_chapter_context", chapter_id=chapter_id, sample_limit=sample_limit)

    def list_tasks(self) -> Any:
        return self.call("list_tasks")

    def list_templates(self) -> Any:
        return self.call("list_templates")

    def bank_stats(self) -> Any:
        return self.call("bank_stats")

    def difficulty_report(self) -> Any:
        return self.call("difficulty_report")

    def list_questions(self, **filters: Any) -> dict[str, Any]:
        return self.call("list_questions", **filters)

    def get_question(self, question_id: int) -> dict[str, Any]:
        """Lấy chi tiết đầy đủ một câu hỏi qua REST API."""
        return mcp_server.api("GET", f"/api/questions/{int(question_id)}")

    def iter_questions(
        self,
        *,
        course_id: int | None = None,
        textbook_id: int | None = None,
        chapter_id: int | None = None,
        bloom: int | None = None,
        qtype: str | None = None,
        task: str | None = None,
        status: str | None = None,
        query: str | None = None,
        per: int = 50,
        full: bool = True,
    ) -> Iterator[dict[str, Any]]:
        """Duyệt qua toàn bộ câu hỏi khớp bộ lọc bằng phân trang tự động.

        Khi ``full=True``, trả về nguyên bản item từ ``/api/questions`` (không bị cắt 300 ký tự)
        để script PTC kiểm định chính xác ``content``, ``options``, ``answer``.
        """
        per = min(max(int(per), 1), 100)
        page = 0
        seen = 0
        while True:
            params = {
                "course_id": course_id,
                "textbook_id": textbook_id,
                "chapter_id": chapter_id,
                "bloom": bloom,
                "qtype": qtype,
                "task": task,
                "status": status,
                "q": query,
                "per": per,
                "page": page,
            }
            res = mcp_server.api("GET", "/api/questions", params)
            items = res.get("items") or []
            if not items:
                break
            for item in items:
                if not full:
                    yield {
                        "id": item.get("id"),
                        "code": item.get("code"),
                        "qtype": item.get("qtype"),
                        "task": item.get("task", ""),
                        "bloom": item.get("bloom"),
                        "status": item.get("status"),
                        "content": (item.get("content") or "")[:300],
                    }
                else:
                    yield item
                seen += 1
            total = int(res.get("total") or 0)
            if seen >= total or len(items) < per:
                break
            page += 1

    # ---- Kiểm trùng & ghi dữ liệu an toàn ----
    def check_duplicates(self, items: list[dict[str, Any]]) -> dict[str, Any]:
        return self.call("check_duplicates", items=items)

    def create_question(self, **kwargs: Any) -> dict[str, Any]:
        """Tạo câu hỏi mới, luôn mặc định ``status='nhap'`` và ``audio_engine='aivis'``."""
        kwargs.setdefault("status", "nhap")
        if kwargs.get("audio_script"):
            kwargs["audio_engine"] = "aivis"
        return self.call("create_question", **kwargs)

    def attach_audio(
        self,
        question_id: int,
        file_path: str,
        *,
        target: str = "question",
        replace: bool = True,
        transcript: str | None = None,
    ) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "question_id": question_id,
            "file_path": file_path,
            "target": target,
            "replace": replace,
        }
        if transcript is not None:
            payload["transcript"] = transcript
        return self.call("attach_audio", **payload)

    def update_question(self, question_id: int, dry_run: bool = True, confirm: bool = False, **fields: Any) -> dict[str, Any]:
        return self.call("update_question", question_id=question_id, dry_run=dry_run, confirm=confirm, **fields)

    def set_status(self, ids: list[int], status: str, dry_run: bool = True) -> dict[str, Any]:
        return self.call("set_status", ids=ids, status=status, dry_run=dry_run)

    def create_exam(self, **kwargs: Any) -> dict[str, Any]:
        kwargs.setdefault("dry_run", True)
        return self.call("create_exam", **kwargs)


def _strip_ruby_tags(text: str) -> str:
    """Thay ``{漢字|かな}`` bằng phần Kanji để kiểm tra dấu cách hoặc từ cấm chính xác hơn."""
    return _RE_RUBY.sub(lambda m: m.group(1).split("|", 1)[0], text)


def _check_ruby_syntax(text: str) -> list[str]:
    errors: list[str] = []
    depth = 0
    for ch in text:
        if ch == "{":
            depth += 1
            if depth > 1:
                errors.append("Ngoặc ruby lồng nhau '{...{...}'")
                break
        elif ch == "}":
            depth -= 1
            if depth < 0:
                errors.append("Dư dấu đóng ngoặc '}' trong ruby")
                break
    if depth > 0:
        errors.append("Thiếu dấu đóng ngoặc '}' trong ruby")
    for m in _RE_RUBY.finditer(text):
        inner = m.group(1)
        if "|" not in inner:
            errors.append(f"Ruby thiếu '|': {{{inner}}}")
        else:
            kanji, kana = inner.split("|", 1)
            if not kanji.strip() or not kana.strip():
                errors.append(f"Ruby rỗng vế kanji hoặc kana: {{{inner}}}")
    return errors


def _extract_unannotated_kanji(text: str) -> set[str]:
    """Lấy các chữ Kanji đứng ngoài ``{漢字|かな}``."""
    without_ruby = _RE_RUBY.sub("", text)
    return set(_RE_KANJI.findall(without_ruby))


def audit_question_rules(question: dict[str, Any], allowed_kanji: set[str] | None = None) -> list[dict[str, str]]:
    """Kiểm tra nhanh các lỗi cứng JLPT trên 1 câu hỏi (không tốn token LLM).

    Trả về danh sách issue ``{"rule", "severity", "field", "message"}``.
    """
    issues: list[dict[str, str]] = []
    qtype = str(question.get("qtype") or "").upper()
    options = question.get("options") or []

    text_fields: list[tuple[str, str]] = []
    for fname in ("content", "passage", "transcript"):
        val = question.get(fname)
        if isinstance(val, str) and val.strip():
            text_fields.append((fname, val))
    if isinstance(options, list):
        for idx, opt in enumerate(options, 1):
            if isinstance(opt, str) and opt.strip():
                text_fields.append((f"option_{idx}", opt))

    for fname, raw in text_fields:
        for msg in _check_ruby_syntax(raw):
            issues.append({"rule": "ruby_syntax", "severity": "error", "field": fname, "message": msg})

        plain = _strip_ruby_tags(raw)
        if _RE_JP_SPACE.search(plain):
            issues.append({
                "rule": "jp_space",
                "severity": "error",
                "field": fname,
                "message": "Có dấu cách (space/full-width) giữa các chữ tiếng Nhật",
            })

        for term in LINGUISTIC_TERMS:
            if term in plain:
                issues.append({
                    "rule": "linguistic_term",
                    "severity": "error",
                    "field": fname,
                    "message": f"Chứa thuật ngữ ngôn ngữ học cấm: {term}",
                })

        for pron in AMBIGUOUS_PRONOUNS:
            if pron in plain:
                issues.append({
                    "rule": "ambiguous_pronoun",
                    "severity": "warning",
                    "field": fname,
                    "message": f"Chứa đại từ dễ mơ hồ: {pron} (cần kiểm tra ngữ cảnh)",
                })

        if allowed_kanji is not None and (fname in ("content", "passage") or fname.startswith("option_")):
            extra = _extract_unannotated_kanji(raw) - allowed_kanji
            if extra:
                issues.append({
                    "rule": "scope_kanji",
                    "severity": "warning",
                    "field": fname,
                    "message": f"Kanji ngoài scope chưa gắn furigana: {''.join(sorted(extra))}",
                })

    if qtype == "DT":
        content = str(question.get("content") or "")
        blanks = _RE_BLANK.findall(content)
        if len(blanks) != 1:
            issues.append({
                "rule": "dt_blank",
                "severity": "error",
                "field": "content",
                "message": f"Câu điền từ (DT) phải có đúng 1 cụm '___' (hiện có {len(blanks)})",
            })

    if qtype == "TN":
        if not isinstance(options, list) or len(options) < 2:
            issues.append({
                "rule": "options_invalid",
                "severity": "error",
                "field": "options",
                "message": "Câu trắc nghiệm (TN) thiếu danh sách phương án hợp lệ",
            })
        else:
            norm = [str(o).strip() for o in options]
            if any(not x for x in norm):
                issues.append({
                    "rule": "options_invalid",
                    "severity": "error",
                    "field": "options",
                    "message": "Có phương án bị rỗng",
                })
            elif len(set(norm)) != len(norm):
                issues.append({
                    "rule": "options_invalid",
                    "severity": "error",
                    "field": "options",
                    "message": "Có phương án bị trùng nội dung",
                })

    return issues


def audit_bank(client: QBankClient, **filters: Any) -> dict[str, Any]:
    """Quét toàn bộ câu hỏi khớp bộ lọc và chỉ trả về các câu có issue."""
    flagged: list[dict[str, Any]] = []
    total_scanned = 0
    for q in client.iter_questions(**filters):
        total_scanned += 1
        issues = audit_question_rules(q)
        if issues:
            flagged.append({
                "id": q.get("id"),
                "code": q.get("code"),
                "qtype": q.get("qtype"),
                "status": q.get("status"),
                "issues": issues,
            })
    return {
        "scanned": total_scanned,
        "flagged_count": len(flagged),
        "flagged": flagged,
    }


def _main() -> None:
    parser = argparse.ArgumentParser(description="QBank PTC Client CLI")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_audit = sub.add_parser("audit", help="Quét lỗi cứng JLPT trên ngân hàng câu hỏi")
    p_audit.add_argument("--course-id", type=int)
    p_audit.add_argument("--textbook-id", type=int)
    p_audit.add_argument("--chapter-id", type=int)
    p_audit.add_argument("--status")
    p_audit.add_argument("--qtype")

    args = parser.parse_args()
    client = QBankClient()
    if args.cmd == "audit":
        report = audit_bank(
            client,
            course_id=args.course_id,
            textbook_id=args.textbook_id,
            chapter_id=args.chapter_id,
            status=args.status,
            qtype=args.qtype,
        )
        print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    _main()
