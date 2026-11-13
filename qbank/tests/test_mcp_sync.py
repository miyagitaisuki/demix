"""Kiểm tra công cụ đồng bộ MCP: repo hiện tại phải khớp, và phát hiện được các kiểu lệch."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tools"))

import check_mcp_sync as c  # noqa: E402


def test_repo_is_in_sync():
    app = open(c.APP, encoding="utf-8").read()
    mcp = open(c.MCP, encoding="utf-8").read()
    routes = c.app_routes(app)
    assert routes, "không đọc được route nào từ app.py"
    assert all(c.path_known(p, routes) for p in c.mcp_paths(mcp))
    assert c.question_columns(app) - c.SERVER_MANAGED <= c.mcp_question_fields(mcp)


def test_detects_removed_route_and_new_column():
    app = '''
class H:
    def _get(self):
        if p == "/api/tree": pass
        m = re.fullmatch(r"/api/questions/(\d+)", p)
    def _put(self):
        conn.execute("""UPDATE questions SET content=:content, extra_col=:extra_col WHERE id=:id""")
'''
    routes = c.app_routes(app)
    assert routes == {"GET /api/tree", "GET /api/questions/(\d+)"}
    assert c.path_known("/api/questions/{}", routes)
    assert not c.path_known("/api/passages/{}", routes)  # route bị xóa/đổi tên
    assert c.question_columns(app) == {"content", "extra_col"}


def test_mcp_paths_handles_fstrings():
    src = 'a = "/api/media"\nb = f"/api/questions/{qid}"\n'
    assert c.mcp_paths(src) == {"/api/media", "/api/questions/{}"}
