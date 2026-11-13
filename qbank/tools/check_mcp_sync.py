#!/usr/bin/env python3
"""Kiểm tra mcp_server.py còn khớp với API và cấu trúc câu hỏi của app.py.

Chạy từ thư mục qbank/:  python3 tools/check_mcp_sync.py          (kiểm tra; mã thoát 1 nếu lệch)
                         python3 tools/check_mcp_sync.py --update (ghi lại mcp_manifest.json sau khi đã xem lại MCP)

Ba kiểm tra:
1. Mọi đường dẫn /api/... mà mcp_server.py gọi phải còn trong app.py (route bị đổi tên/xóa thì MCP hỏng).
2. Các trường của câu hỏi (cột trong UPDATE questions) phải có trong QUESTION_FIELDS của mcp_server.py,
   nếu không attach_audio sẽ làm mất dữ liệu khi PUT.
3. Danh sách route của app.py so với mcp_manifest.json (ảnh chụp lần đồng bộ trước): có route mới/mất thì báo,
   để người sửa cân nhắc thêm/sửa công cụ MCP, rồi chạy --update.
"""
import ast
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
APP = os.path.join(BASE, "app.py")
MCP = os.path.join(BASE, "mcp_server.py")
MANIFEST = os.path.join(BASE, "mcp_manifest.json")
HANDLERS = {"_get": "GET", "_post": "POST", "_put": "PUT", "_delete": "DELETE"}
# cột của câu hỏi do app tự quản lý, MCP không cần gửi lại khi PUT
SERVER_MANAGED = {"passage_id", "media"}


def app_routes(src):
    """Tập 'METHOD /api/...' lấy từ các chuỗi trong _get/_post/_put/_delete của app.py (kể cả mẫu regex)."""
    routes = set()
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.FunctionDef) and node.name in HANDLERS:
            for c in ast.walk(node):
                if isinstance(c, ast.Constant) and isinstance(c.value, str) and c.value.startswith(("/api/", "/api")):
                    if c.value.startswith("/api/"):
                        routes.add(f"{HANDLERS[node.name]} {c.value}")
    return routes


def question_columns(src):
    """Cột được gán trong câu UPDATE questions SET ... của app.py."""
    cols = set()
    for m in re.finditer(r"UPDATE questions SET(.*?)WHERE id=:id", src, re.S):
        cols |= set(re.findall(r"(\w+)=:", m.group(1)))
    return cols


def mcp_paths(src):
    """Đường dẫn /api/... mcp_server.py gọi (bỏ phần f-string động và query)."""
    paths, inside = set(), set()
    tree = ast.parse(src)
    for node in ast.walk(tree):
        if isinstance(node, ast.JoinedStr):
            inside |= {id(v) for v in node.values}
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str) and node.value.startswith("/api/")                 and id(node) not in inside:
            paths.add(node.value)
        elif isinstance(node, ast.JoinedStr):
            head = "".join(v.value for v in node.values if isinstance(v, ast.Constant))
            if head.startswith("/api/"):
                paths.add(head.split("{")[0] + "{}")
    return paths


def mcp_question_fields(src):
    for node in ast.walk(ast.parse(src)):
        if isinstance(node, ast.Assign) and any(getattr(t, "id", "") == "QUESTION_FIELDS" for t in node.targets):
            return set(ast.literal_eval(node.value))
    return set()


def path_known(path, routes):
    r"""Đường dẫn MCP gọi có khớp một route app (route dạng regex như /api/questions/(\d+) cũng được)."""
    probe = path.replace("{}", "1")
    for r in routes:
        pat = r.split(" ", 1)[1]
        if pat == probe or re.fullmatch(pat, probe):
            return True
        if pat.endswith("/") and probe.startswith(pat):
            return True
    return False


def main():
    update = "--update" in sys.argv
    app_src, mcp_src = open(APP, encoding="utf-8").read(), open(MCP, encoding="utf-8").read()
    routes = app_routes(app_src)
    problems = []

    for p in sorted(mcp_paths(mcp_src)):
        if not path_known(p, routes):
            problems.append(f"MCP gọi {p} nhưng app.py không còn route này: sửa mcp_server.py cho khớp.")

    missing = question_columns(app_src) - SERVER_MANAGED - mcp_question_fields(mcp_src)
    if missing:
        problems.append("Cột câu hỏi mới trong app.py chưa có trong QUESTION_FIELDS của mcp_server.py: "
                        + ", ".join(sorted(missing)) + " (attach_audio sẽ làm mất dữ liệu khi PUT).")

    old = json.load(open(MANIFEST, encoding="utf-8")) if os.path.exists(MANIFEST) else {"routes": []}
    added, removed = sorted(routes - set(old["routes"])), sorted(set(old["routes"]) - routes)
    if (added or removed) and not update:
        if added:
            problems.append("Route mới trong app.py từ lần đồng bộ MCP trước: " + "; ".join(added))
        if removed:
            problems.append("Route đã mất khỏi app.py: " + "; ".join(removed))
        problems.append("Hãy xem lại mcp_server.py (thêm/sửa công cụ, cập nhật tests/test_mcp_server.py) rồi chạy: "
                        "python3 tools/check_mcp_sync.py --update")

    if update and not [p for p in problems if "MCP gọi" in p or "QUESTION_FIELDS" in p]:
        with open(MANIFEST, "w", encoding="utf-8") as f:
            json.dump({"note": "Ảnh chụp route của app.py lần đồng bộ MCP gần nhất. Cập nhật bằng tools/check_mcp_sync.py --update.",
                       "routes": sorted(routes)}, f, ensure_ascii=False, indent=1)
            f.write("\n")
        print(f"Đã ghi {os.path.basename(MANIFEST)}: {len(routes)} route (thêm {len(added)}, mất {len(removed)}).")
    if problems:
        print("MCP CHƯA ĐỒNG BỘ VỚI APP:")
        for p in problems:
            print(" -", p)
        sys.exit(1)
    if not update:
        print(f"MCP đồng bộ với app ({len(routes)} route).")


if __name__ == "__main__":
    main()
