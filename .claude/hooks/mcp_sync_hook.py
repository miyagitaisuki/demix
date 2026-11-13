#!/usr/bin/env python3
"""Hook PostToolUse của Claude Code: sau khi sửa qbank/app.py hoặc qbank/mcp_server.py, kiểm tra MCP còn khớp app.
Nếu lệch, in lý do ra stderr và thoát 2 để Claude thấy và sửa mcp_server.py (rồi chạy check_mcp_sync.py --update)."""
import json
import os
import subprocess
import sys

try:
    data = json.load(sys.stdin)
except Exception:
    sys.exit(0)
path = str((data.get("tool_input") or {}).get("file_path", "")).replace("\\", "/")
if not path.endswith(("qbank/app.py", "qbank/mcp_server.py")):
    sys.exit(0)
root = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
r = subprocess.run([sys.executable, os.path.join(root, "qbank", "tools", "check_mcp_sync.py")],
                   cwd=os.path.join(root, "qbank"), capture_output=True, text=True, encoding="utf-8")
if r.returncode != 0:
    sys.stderr.write(r.stdout + r.stderr)
    sys.exit(2)
