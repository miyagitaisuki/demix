"""Cập nhật từ GitHub Releases: chọn bản mới, kiểm tra nguồn tải, thay file .exe."""
import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402


def release(tag, *, draft=False, prerelease=False, asset_name="QBank.exe",
            url="https://github.com/miyagitaisuki/demix/releases/download/x/QBank.exe"):
    assets = [] if asset_name is None else [{"name": asset_name, "browser_download_url": url, "size": 10}]
    return {"tag_name": tag, "draft": draft, "prerelease": prerelease, "assets": assets, "body": f"notes {tag}"}


def test_parse_version():
    assert app.parse_version("1.2.3") == (1, 2, 3)
    assert app.parse_version("qbank-v10.0.1") == (10, 0, 1)
    assert app.parse_version("v2.0.0") == (2, 0, 0)
    assert app.parse_version("1.2") is None
    assert app.parse_version("") is None
    assert app.parse_version(None) is None


def test_safe_download_url():
    assert app.safe_download_url("https://github.com/a/b/releases/download/t/QBank.exe")
    assert app.safe_download_url("https://objects.githubusercontent.com/x/QBank.exe")
    assert not app.safe_download_url("http://github.com/a/QBank.exe")          # không phải HTTPS
    assert not app.safe_download_url("https://evil.example/QBank.exe")          # máy chủ lạ
    assert not app.safe_download_url("https://github.com.evil.example/QBank.exe")
    assert not app.safe_download_url(None)


def test_pick_update_chooses_newest_qbank_release():
    rels = [release("qbank-v1.0.0"), release("qbank-v1.2.0"), release("qbank-v1.1.0")]
    upd = app.pick_update(rels, "1.0.0")
    assert upd["version"] == "1.2.0"
    assert upd["notes"] == "notes qbank-v1.2.0"


def test_pick_update_none_when_up_to_date():
    rels = [release("qbank-v1.0.0")]
    assert app.pick_update(rels, "1.0.0") is None
    assert app.pick_update(rels, "1.1.0") is None


def test_pick_update_skips_drafts_prereleases_and_other_tags():
    rels = [
        release("qbank-v9.0.0", draft=True),
        release("qbank-v8.0.0", prerelease=True),
        release("giaotrinh-v7.0.0"),        # tag của phần khác trong repo
        release("v6.0.0"),                  # tag không có tiền tố qbank-
        release("qbank-v5.0.0", asset_name="Other.exe"),  # thiếu QBank.exe
        release("qbank-v4.0.0", url="https://evil.example/QBank.exe"),  # nguồn tải lạ
        release("qbank-v1.5.0"),
    ]
    upd = app.pick_update(rels, "1.0.0")
    assert upd["version"] == "1.5.0"


def test_pick_update_empty_list():
    assert app.pick_update([], "1.0.0") is None


def test_pick_update_unparseable_current_version_gives_nothing():
    assert app.pick_update([release("qbank-v1.0.0")], "dev") is None


def test_swap_exe_replaces_and_keeps_old(tmp_path):
    cur = tmp_path / "QBank.exe"
    new = tmp_path / "QBank.new.exe"
    cur.write_bytes(b"old")
    new.write_bytes(b"new")
    old = app.swap_exe(str(new), str(cur))
    assert cur.read_bytes() == b"new"
    assert not new.exists()
    assert open(old, "rb").read() == b"old"


def test_swap_exe_removes_previous_old_copy(tmp_path):
    cur = tmp_path / "QBank.exe"
    new = tmp_path / "QBank.new.exe"
    cur.write_bytes(b"v2")
    (tmp_path / "QBank.exe.old").write_bytes(b"v1")
    new.write_bytes(b"v3")
    app.swap_exe(str(new), str(cur))
    assert cur.read_bytes() == b"v3"
    assert (tmp_path / "QBank.exe.old").read_bytes() == b"v2"


def test_can_self_update_false_outside_frozen_windows():
    # chạy từ mã nguồn (không đóng gói) thì không được tự thay file
    assert app.can_self_update() is False


def test_install_refused_when_not_frozen():
    with pytest.raises(ValueError):
        app.install_update()


def test_read_version_matches_file():
    with open(os.path.join(os.path.dirname(__file__), "..", "VERSION"), encoding="utf-8") as f:
        assert app.APP_VERSION == f.read().strip()
