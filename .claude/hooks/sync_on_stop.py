#!/usr/bin/env python3
"""Hook Stop của Claude Code: trước khi Claude kết thúc lượt, kiểm tra MCP và CSDL còn khớp app.

Hook PostToolUse chỉ chạy khi sửa bằng Edit/Write; sửa qua Bash (sed, script Python, git) thì bị bỏ qua.
Hook này xét theo git: file đã đổi mà chưa push (chưa commit, hoặc commit chưa có trên nhánh upstream).
- Đổi qbank/app.py hoặc qbank/mcp_server.py: chạy tools/check_mcp_sync.py.
- Đổi qbank/app.py hoặc qbank/mau_nhap/: chạy tools/check_db_sync.py.
- Đổi qbank/app.py ở đoạn xử lý route mà MCP đang gọi: nhắc xem lại công cụ MCP tương ứng (mô tả, tham số),
  vì check_mcp_sync.py chỉ bắt route mất/mới, không bắt thay đổi hành vi (ví dụ thêm định dạng nhập mới).
Lệch thì in lý do ra stderr và thoát 2 để Claude sửa tiếp. Chỉ chặn một lần mỗi lượt (stop_hook_active).
"""
import json
import os
import re
import subprocess
import sys

try:
    data = json.load(sys.stdin)
except Exception:
    data = {}
if data.get("stop_hook_active"):
    sys.exit(0)

root = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
qbank = os.path.join(root, "qbank")


def git(*args):
    r = subprocess.run(["git", *args], cwd=root, capture_output=True, text=True, encoding="utf-8")
    return r.stdout if r.returncode == 0 else ""


def base_ref():
    """Upstream của nhánh hiện tại; chưa có thì origin/main; không có nữa thì HEAD (chỉ xét phần chưa commit)."""
    for ref in ("@{upstream}", "origin/main"):
        if git("rev-parse", "--verify", "--quiet", ref).strip():
            return ref
    return "HEAD"


base = base_ref()
changed = set(git("diff", "--name-only", base).split()) | {
    ln[3:].strip() for ln in git("status", "--porcelain").splitlines() if ln.startswith("??")}
if not changed:
    sys.exit(0)

problems = []


def run_check(script):
    r = subprocess.run([sys.executable, os.path.join("tools", script)], cwd=qbank,
                       capture_output=True, text=True, encoding="utf-8")
    if r.returncode != 0:
        problems.append(f"[{script}]\n{(r.stdout + r.stderr).strip()}")


app_changed = "qbank/app.py" in changed
if app_changed or "qbank/mcp_server.py" in changed:
    run_check("check_mcp_sync.py")
if app_changed or any(p.startswith("qbank/mau_nhap/") for p in changed):
    run_check("check_db_sync.py")

if app_changed and "qbank/mcp_server.py" not in changed:
    # route MCP đang gọi
    try:
        with open(os.path.join(qbank, "mcp_server.py"), encoding="utf-8") as f:
            used = set(re.findall(r'"(/api/[a-z0-9_\-/]+)', f.read()))
    except OSError:
        used = set()
    # route có đoạn xử lý bị sửa: tìm `p == "/api/..."` gần nhất phía trên mỗi hunk của app.py
    try:
        with open(os.path.join(qbank, "app.py"), encoding="utf-8") as f:
            lines = f.read().splitlines()
    except OSError:
        lines = []
    touched = set()
    for m in re.finditer(r"^@@ -\d+(?:,\d+)? \+(\d+)(?:,(\d+))? @@", git("diff", "-U0", base, "--", "qbank/app.py"), re.M):
        start, count = int(m.group(1)), int(m.group(2) or 1)
        for i in range(min(start + max(count, 1) - 1, len(lines)) - 1, -1, -1):
            r = re.search(r'p\s*(?:==|\.startswith\()\s*"(/api/[a-z0-9_\-/]+)', lines[i])
            if r:
                touched.add(r.group(1).rstrip("/"))
                break
            if re.match(r"\s*def (_get|_post|_put|_delete)\(", lines[i]):
                break
    hit = sorted(r for r in touched if any(u.rstrip("/") == r or u.startswith(r + "/") for u in used))
    if hit:
        problems.append("app.py đã sửa phần xử lý route mà MCP đang gọi: " + ", ".join(hit) + ".\n"
                        "Xem lại công cụ tương ứng trong qbank/mcp_server.py (mô tả, tham số, định dạng nhận) "
                        "để khớp hành vi mới. Nếu không cần đổi thì bỏ qua thông báo này.")

if problems:
    sys.stderr.write("Kiểm tra đồng bộ MCP/CSDL trước khi kết thúc:\n\n" + "\n\n".join(problems) + "\n")
    sys.exit(2)
