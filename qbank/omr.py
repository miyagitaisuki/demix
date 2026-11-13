"""Phiếu trả lời trắc nghiệm để quét chấm tự động (OMR) và tờ đáp án đục lỗ.

Mọi tọa độ tính bằng mm trên khổ A4 dọc (210 × 297). Bốn ô vuông đen ở bốn góc là mốc căn chỉnh;
một ô nhỏ cạnh mốc trên-trái cho biết chiều phiếu (ảnh chụp ngược/xoay vẫn đọc được).
Trang chấm bài (static/cham-bai.html) dùng đúng bố cục do layout() trả về để tìm vị trí từng ô tô,
nên phiếu in và bộ đọc luôn khớp nhau. Hàm ở đây thuần Python, không phụ thuộc thư viện ngoài.
"""
from html import escape

PAGE_W, PAGE_H = 210.0, 297.0
FID = 7.0  # cạnh ô mốc góc
FIDUCIALS = [(8.0, 8.0), (195.0, 8.0), (8.0, 282.0), (195.0, 282.0)]  # góc trên-trái của 4 mốc: TL, TR, BL, BR
ORIENT = (17.0, 8.0, 4.0)  # ô nhỏ chỉ chiều, cạnh mốc trên-trái (x, y, cạnh)
LETTERS = "ABCDEFGH"
MAX_OPTS = 8
BUBBLE_R = 2.0
# lưới trắc nghiệm
TN_LABEL_W, TN_PITCH_X, TN_PITCH_Y, TN_GROUP_GAP, TN_COL_GAP = 9.0, 6.0, 5.5, 2.5, 5.0
# khối đúng/sai: mỗi câu 4 ý a–d, mỗi ý hai ô Đ / S
DS_PITCH_Y, DS_W, DS_ITEMS = 5.5, 25.0, 4
AREA_X0, AREA_X1, AREA_Y0, AREA_Y1 = 18.0, 192.0, 108.0, 276.0
DIGIT_PITCH_X, DIGIT_PITCH_Y = 5.5, 5.0


def code_digits(code, n):
    """Phần số của mã đề, đủ n chữ số (A001 -> 001); dùng để tô và để đối chiếu mã đề đọc được trên phiếu."""
    d = "".join(ch for ch in str(code) if ch.isdigit())
    return d[-n:].rjust(n, "0") if n else ""


def fid_centers():
    return [(x + FID / 2, y + FID / 2) for x, y in FIDUCIALS]


def norm_spec(spec):
    """spec: tn = số dòng trắc nghiệm (đánh số 1..tn), ds = danh sách số câu đúng/sai, opts = số phương án,
    sbd = số chữ số báo danh, made = số chữ số mã đề."""
    tn = int(spec.get("tn") or 0)
    ds = spec.get("ds") or []
    if isinstance(ds, (int, str)):  # số câu đúng/sai: đánh số tiếp sau phần trắc nghiệm
        ds = list(range(tn + 1, tn + 1 + int(ds or 0)))
    ds = sorted({int(x) for x in ds})
    opts = int(spec.get("opts") or 4)
    sbd = int(spec.get("sbd") or 6)
    made = int(spec.get("made") or 3)
    if not 0 <= tn <= 200:
        raise ValueError("Số câu trắc nghiệm trên phiếu phải từ 0 đến 200")
    if len(ds) > 40:
        raise ValueError("Tối đa 40 câu đúng/sai trên một phiếu")
    if not tn and not ds:
        raise ValueError("Phiếu cần ít nhất một câu trắc nghiệm hoặc đúng/sai")
    if not 2 <= opts <= MAX_OPTS:
        raise ValueError(f"Số phương án phải từ 2 đến {MAX_OPTS}")
    if not 0 <= sbd <= 10 or not 1 <= made <= 4:
        raise ValueError("Số báo danh tối đa 10 chữ số, mã đề 1–4 chữ số")
    return {"tn": tn, "ds": ds, "opts": opts, "sbd": sbd, "made": made}


def _digit_block(x0, y0, digits, kind):
    """Khối tô số (số báo danh / mã đề): mỗi cột một chữ số, 10 hàng 0–9."""
    cols = []
    for c in range(digits):
        cx = x0 + 3 + c * DIGIT_PITCH_X
        cols.append([{"x": round(cx, 2), "y": round(y0 + 14 + d * DIGIT_PITCH_Y, 2), "digit": d} for d in range(10)])
    return {"kind": kind, "x": x0, "y": y0, "w": 6 + digits * DIGIT_PITCH_X, "h": 14 + 10 * DIGIT_PITCH_Y, "cols": cols}


def layout(spec):
    """Vị trí mọi ô tô trên phiếu (mm). Lỗi nếu số câu vượt sức chứa của một trang A4."""
    s = norm_spec(spec)
    out = {"page": [PAGE_W, PAGE_H], "fiducials": fid_centers(), "fid_size": FID, "orient": ORIENT,
           "bubble_r": BUBBLE_R, "spec": s, "white": [], "tn": [], "ds": [], "tn_cols": []}
    sbd = _digit_block(20.0, 40.0, s["sbd"], "sbd") if s["sbd"] else None
    made_x = 20.0 + (sbd["w"] + 8 if sbd else 0)
    out["sbd"] = sbd
    out["made"] = _digit_block(made_x, 40.0, s["made"], "made")
    # điểm giấy trắng để lấy mức sáng tham chiếu (lề, khoảng trống giữa các khối)
    out["white"] = [[14, 40], [14, 150], [14, 250], [196, 40], [196, 150], [196, 250], [105, 4.5], [60, 280]]
    y = AREA_Y0
    if s["tn"]:
        colw = TN_LABEL_W + s["opts"] * TN_PITCH_X + TN_COL_GAP
        ncols = max(1, int((AREA_X1 - AREA_X0 + TN_COL_GAP) // colw))
        per_col = -(-s["tn"] // ncols)
        ncols = -(-s["tn"] // per_col)
        groups = -(-per_col // 5)
        h = 5 + per_col * TN_PITCH_Y + (groups - 1) * TN_GROUP_GAP
        if y + h > AREA_Y1:
            cap = ncols * int((AREA_Y1 - y - 5 + TN_GROUP_GAP) // (5 * TN_PITCH_Y + TN_GROUP_GAP)) * 5
            raise ValueError(f"Phiếu không đủ chỗ cho {s['tn']} câu trắc nghiệm {s['opts']} phương án (tối đa khoảng {cap})")
        for c in range(ncols):
            x0 = AREA_X0 + c * colw
            out["tn_cols"].append({"x": round(x0, 2), "y": y, "w": colw - TN_COL_GAP})
            for r in range(per_col):
                n = c * per_col + r + 1
                if n > s["tn"]:
                    break
                ry = y + 5 + r * TN_PITCH_Y + (r // 5) * TN_GROUP_GAP + TN_PITCH_Y / 2
                out["tn"].append({"n": n, "y": round(ry, 2), "label_x": round(x0 + TN_LABEL_W - 1.5, 2),
                                  "bubbles": [{"x": round(x0 + TN_LABEL_W + k * TN_PITCH_X + TN_PITCH_X / 2, 2),
                                               "y": round(ry, 2), "opt": LETTERS[k]} for k in range(s["opts"])]})
        y += h + 6
    if s["ds"]:
        per_row = int((AREA_X1 - AREA_X0) // DS_W)
        block_h = 6 + 4 + DS_ITEMS * DS_PITCH_Y
        rows = -(-len(s["ds"]) // per_row)
        out["ds_title_y"] = y
        y += 5
        if y + rows * (block_h + 3) > AREA_Y1 + 3:
            raise ValueError(f"Phiếu không đủ chỗ cho {len(s['ds'])} câu đúng/sai (giảm số câu trắc nghiệm hoặc đúng/sai)")
        for i, n in enumerate(s["ds"]):
            bx = AREA_X0 + (i % per_row) * DS_W
            by = y + (i // per_row) * (block_h + 3)
            items = []
            for k in range(DS_ITEMS):
                iy = by + 10 + k * DS_PITCH_Y + DS_PITCH_Y / 2
                items.append({"item": "abcd"[k], "y": round(iy, 2),
                              "bubbles": [{"x": round(bx + 9 + j * 6, 2), "y": round(iy, 2), "opt": "DS"[j]} for j in range(2)]})
            out["ds"].append({"n": n, "x": round(bx, 2), "y": round(by, 2), "h": block_h, "items": items})
    return out


# ---------------------------------------------------------------- vẽ phiếu (SVG)
def _t(x, y, text, size=3.2, anchor="start", weight="normal", fill="#000"):
    return (f'<text x="{x:.2f}" y="{y:.2f}" font-size="{size}" text-anchor="{anchor}" font-weight="{weight}" '
            f'fill="{fill}" font-family="Arial, Helvetica, sans-serif">{escape(str(text))}</text>')


def _bubble(x, y, label="", key=None, r=BUBBLE_R):
    """Ô tô. key=None: phiếu trắng; key=True: đáp án đúng (đục lỗ: tô đen); key=False: không phải đáp án."""
    fill = "#000" if key else "none"
    out = f'<circle cx="{x:.2f}" cy="{y:.2f}" r="{r}" fill="{fill}" stroke="#c2185b" stroke-width="0.25"/>'
    if label and not key:
        out += _t(x, y + 0.95, label, 2.3, "middle", fill="#c2185b")
    return out


def sheet_svg(spec, title="", subtitle="", key=None, code=""):
    """Phiếu trả lời (key=None) hoặc tờ đáp án đục lỗ của một mã đề (key={số câu: 'A'} và {số câu: 'DSDS'})."""
    L = layout(spec)
    s = L["spec"]
    pink = "#c2185b"
    parts = [f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {PAGE_W} {PAGE_H}" width="{PAGE_W}mm" height="{PAGE_H}mm">',
             f'<rect x="0" y="0" width="{PAGE_W}" height="{PAGE_H}" fill="#fff"/>']
    for x, y in FIDUCIALS:
        parts.append(f'<rect x="{x}" y="{y}" width="{FID}" height="{FID}" fill="#000"/>')
    ox, oy, os_ = ORIENT
    parts.append(f'<rect x="{ox}" y="{oy}" width="{os_}" height="{os_}" fill="#000"/>')
    head = "TỜ ĐÁP ÁN ĐỤC LỖ" if key is not None else "PHIẾU TRẢ LỜI TRẮC NGHIỆM"
    parts.append(_t(105, 15, head, 5, "middle", "bold"))
    if title:
        parts.append(_t(105, 21, title, 3.4, "middle"))
    if subtitle:
        parts.append(_t(105, 26, subtitle, 3, "middle"))
    if key is not None:
        parts.append(_t(105, 33, f"MÃ ĐỀ {code}  ·  in ra giấy cứng rồi đục lỗ ở các ô đen, đặt chồng lên phiếu của thí sinh",
                        3, "middle", fill="#444"))
    else:
        parts.append(_t(20, 33, "Họ và tên: ................................................................", 3.4))
        parts.append(_t(125, 33, "Lớp: ........................  Phòng: ..........", 3.4))
    for blk, name in ((L["sbd"], "SỐ BÁO DANH"), (L["made"], "MÃ ĐỀ")):
        if not blk:
            continue
        parts.append(f'<rect x="{blk["x"]}" y="{blk["y"]}" width="{blk["w"]:.2f}" height="{blk["h"] + 2:.2f}" '
                     f'fill="none" stroke="{pink}" stroke-width="0.35"/>')
        parts.append(_t(blk["x"] + blk["w"] / 2, blk["y"] + 4, name, 2.8, "middle", "bold", pink))
        for col in blk["cols"]:
            cx = col[0]["x"]
            parts.append(f'<rect x="{cx - 2.4:.2f}" y="{blk["y"] + 5.5:.2f}" width="4.8" height="5" fill="none" '
                         f'stroke="{pink}" stroke-width="0.25"/>')
            digits = code_digits(code, len(blk["cols"])) if (key is not None and blk["kind"] == "made") else ""
            if digits:
                parts.append(_t(cx, blk["y"] + 9.4, digits[blk["cols"].index(col)], 3.2, "middle", "bold"))
            for b in col:
                mark = (key is not None and blk["kind"] == "made" and digits and int(digits[blk["cols"].index(col)]) == b["digit"])
                parts.append(_bubble(b["x"], b["y"], str(b["digit"]), True if mark else None, 1.8))
    # hướng dẫn tô
    gx = (L["made"]["x"] + L["made"]["w"] + 8)
    if gx < 150:
        lines = ["HƯỚNG DẪN", "• Dùng bút chì 2B hoặc bút mực đen, tô kín ô tròn.",
                 "• Đúng:  ●     Sai:  ○  ⊘  ✓  ✗", "• Ghi số báo danh, mã đề vào ô rồi tô số tương ứng.",
                 "• Muốn đổi đáp án: tẩy sạch ô cũ rồi tô ô mới.", "• Không gấp, không làm bẩn bốn ô vuông ở góc."]
        for i, ln in enumerate(lines):
            parts.append(_t(gx, 46 + i * 6, ln, 3 if i else 3.2, weight="bold" if not i else "normal"))
    if s["tn"]:
        parts.append(_t(AREA_X0, AREA_Y0 - 2.5, "PHẦN TRẮC NGHIỆM (chọn một phương án)", 3, weight="bold", fill=pink))
        for col in L["tn_cols"]:
            for k in range(s["opts"]):
                parts.append(_t(col["x"] + TN_LABEL_W + k * TN_PITCH_X + TN_PITCH_X / 2, col["y"] + 3.6, LETTERS[k],
                                2.8, "middle", "bold", pink))
        for row in L["tn"]:
            parts.append(_t(row["label_x"], row["y"] + 1.1, row["n"], 3, "end", "bold"))
            ans = (key or {}).get(row["n"]) if key is not None else None
            for b in row["bubbles"]:
                parts.append(_bubble(b["x"], b["y"], b["opt"], (b["opt"] == ans) if key is not None else None))
    if s["ds"]:
        parts.append(_t(AREA_X0, L["ds_title_y"] + 2.5, "PHẦN ĐÚNG / SAI (mỗi ý tô Đ hoặc S)", 3, weight="bold", fill=pink))
        for blk in L["ds"]:
            parts.append(f'<rect x="{blk["x"]:.2f}" y="{blk["y"]:.2f}" width="{DS_W - 2}" height="{blk["h"]:.2f}" '
                         f'fill="none" stroke="{pink}" stroke-width="0.25"/>')
            parts.append(_t(blk["x"] + 2, blk["y"] + 4, f"Câu {blk['n']}", 3, weight="bold"))
            parts.append(_t(blk["x"] + 9, blk["y"] + 8.6, "Đ", 2.6, "middle", "bold", pink))
            parts.append(_t(blk["x"] + 15, blk["y"] + 8.6, "S", 2.6, "middle", "bold", pink))
            ans = str((key or {}).get(blk["n"], "")) if key is not None else ""
            for k, it in enumerate(blk["items"]):
                parts.append(_t(blk["x"] + 3.5, it["y"] + 1.1, it["item"] + ")", 3))
                want = ans[k] if k < len(ans) else ""
                for b in it["bubbles"]:
                    parts.append(_bubble(b["x"], b["y"], "", (b["opt"] == want) if key is not None else None, 1.9))
    parts.append(_t(105, 293, "QBank · phiếu chấm máy · không gấp, không viết vào bốn ô vuông góc", 2.4, "middle", fill="#777"))
    parts.append("</svg>")
    return "".join(parts)


def sheet_page(pages, title="Phiếu trả lời trắc nghiệm"):
    """Trang HTML in được: mỗi phần tử của pages là một SVG khổ A4, mỗi SVG một trang giấy."""
    body = "".join(f'<div class="pg">{svg}</div>' for svg in pages)
    return f"""<!DOCTYPE html><html lang="vi"><head><meta charset="utf-8"><title>{escape(title)}</title>
<style>@page {{ size: A4 portrait; margin: 0; }} html, body {{ margin: 0; padding: 0; background: #e5e7eb; }}
.pg {{ width: 210mm; height: 297mm; margin: 8mm auto; background: #fff; box-shadow: 0 2px 10px rgba(0,0,0,.15); page-break-after: always; }}
.pg svg {{ display: block; }} .bar {{ text-align: center; font: 14px Arial, sans-serif; padding: 10px; }}
@media print {{ html, body {{ background: #fff; }} .pg {{ margin: 0; box-shadow: none; }} .bar {{ display: none; }} }}</style></head>
<body><div class="bar">In ở khổ A4, tỷ lệ 100% (không chọn "Vừa trang" nếu làm méo ô vuông góc). <button onclick="print()">In</button></div>
{body}</body></html>"""


def keys_from_flat(flat):
    """Đáp án chấm máy của một mã đề từ flatten(exam): câu TN (≤ 8 phương án) và câu DS (≤ 4 ý).
    Các câu khác (DT, TL, nối tiêu đề nhiều hơn 8 tiêu đề) không có trên phiếu, chấm tay."""
    items = []
    for n, _, it in flat:
        opts = it.get("options") or []
        if it["qtype"] == "TN" and 2 <= len(opts) <= MAX_OPTS:
            items.append({"n": n, "qtype": "TN", "answer": it["answer"], "nopt": len(opts), "qcode": it.get("code", ""),
                          "points": it.get("points")})
        elif it["qtype"] == "DS" and 1 <= len(opts) <= DS_ITEMS:
            items.append({"n": n, "qtype": "DS", "answer": it["answer"], "nopt": len(opts), "qcode": it.get("code", ""),
                          "points": it.get("points")})
    return items


def spec_from_keys(keys, sbd=6):
    """Bố cục phiếu đủ cho mọi mã đề của một lần xuất: số dòng TN = số câu TN lớn nhất, ô DS theo số câu."""
    tn = max((i["n"] for k in keys.values() for i in k if i["qtype"] == "TN"), default=0)
    ds = sorted({i["n"] for k in keys.values() for i in k if i["qtype"] == "DS"})
    opts = max([i["nopt"] for k in keys.values() for i in k if i["qtype"] == "TN"] + [4])
    digits = max([len("".join(ch for ch in c if ch.isdigit())) for c in keys] + [1])
    return {"tn": tn, "ds": ds, "opts": opts, "sbd": sbd, "made": min(max(digits, 3), 4)}
