"""Phương án là ảnh đính kèm câu: [hình: tên ảnh] tra cứu theo tên ảnh đã đính kèm."""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402

MEDIA = [{"kind": "image", "file": "a" * 32 + ".png", "name": "Tranh1.png"},
         {"kind": "audio", "file": "b" * 32 + ".mp3", "name": "nghe.mp3"}]


def test_chu_khong_phai_anh():
    assert app.option_image("Tokyo", MEDIA) is None
    assert app.option_image("A. Tokyo", MEDIA) is None


def test_tim_anh_khong_phan_biet_hoa_thuong():
    assert app.option_image("[Hình: tranh1.png]", MEDIA) == "a" * 32 + ".png"
    assert app.option_image(" [hinh:Tranh1.png] ", MEDIA) == "a" * 32 + ".png"


def test_khong_co_anh_khop_tra_ve_chuoi_rong():
    assert app.option_image("[hình: khong.png]", MEDIA) == ""
    assert app.option_image("[hình: nghe.mp3]", MEDIA) == ""  # chỉ nhận ảnh, không nhận âm thanh


def test_opt_html_hien_anh():
    html = app.opt_html("[hình: Tranh1.png]", MEDIA)
    assert "/media/" + "a" * 32 + ".png" in html
    assert app.opt_html("Tokyo", MEDIA) == app.rich_html("Tokyo")
