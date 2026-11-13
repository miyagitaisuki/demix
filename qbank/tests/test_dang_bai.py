"""Dạng bài: danh mục, tự nhận dạng, lọc khi ra đề (câu lẻ và đoạn văn), nhập/xuất, thêm dạng riêng, nâng cấp CSDL cũ."""
import os
import sqlite3
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import app  # noqa: E402

TFNG = ("正（真）", "誤（偽）", "情報なし（書かれていない）")


@pytest.mark.parametrize("args,expected", [
    (("TN", "次の漢字の読み方を選びなさい。話します", ["ききます", "はなします"]), "kanji-doc"),
    (("TN", "駅員さんに電車の乗り場を<u>案内</u>してもらいました。", ["あんない", "あない"]), "kanji-doc"),
    (("TN", "次の言葉を漢字で書いたものを選びなさい。みず", ["水", "木"]), "kanji-viet"),
    (("TN", "★　ただしい　じゅんばんに　ならべてください。", ["a", "b"]), "np-sapxep"),
    (("TN", "次の文の★に入る最もよいものを選びなさい。 / 雨が★、外出はやめよう。", ["降ったら", "降って"]), "np-chon"),
    (("TN", "6時（　）起きます。", ["に", "で", "と", "の"]), "np-trotu"),
    (("TN", "「〜はずです」の意味として正しいものはどれですか。", ["a", "b"]), "np-ynghia"),
    (("TN", "文章の（3）に入る最もよいものを選びなさい。", ["a", "b"], "", "昨日（1）…"), "np-doanvan"),
    (("TN", "料金はいくらですか。", ["0円", "500円"], "", "「■ 駅前駐輪場 ご利用案内」料金：500円"), "doc-thongtin"),
    (("TN", "筆者は何をしましたか。", ["a", "b"], "", "週末、私は図書館へ行きました。"), "doc-noidung"),
    (("TN", "筆者は料理教室に通っている。", list(TFNG), ",TFNG,", "週末は単語を覚えます。"), "doc-tfng"),
    (("TN", "段落1 に合う見出しを選びなさい。", ["a", "b"], ",noitieude,", "段落1 …"), "doc-tieude"),
    (("TN", "いちばん　いい　へんじは　どれですか。", ["a", "b"], "", "", "N5 Nghe tự do"), "nghe-sokuji"),
    (("TN", "男の人は　このあと　まず　なにを　しますか。", ["a", "b"], "", "", "", True), "nghe-kadai"),
    (("TL", "経験を書きなさい。200〜300字で書きなさい。", []), "viet-sakubun"),
    (("DT", "あの人が断るはず＿＿よ。", []), "np-dien"),
    (("DS", "Chọn đáp án đúng sai", ["a", "b"]), "np-dungsai"),
    (("TN", "稟議書で気をつけることは何ですか。", ["a", "b"], "", "", "Sổ tay tiếng Nhật thương mại"), "cn-kienthuc"),
])
def test_guess_task(args, expected):
    assert app.guess_task(*args) == expected


@pytest.fixture
def bank(tmp_path, monkeypatch):
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "data" / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    monkeypatch.setattr(app, "BACKUP_DIR", str(tmp_path / "data" / "backup"))
    os.makedirs(tmp_path / "data")
    app.init_storage()
    ch = app.ensure_chapter(["TN3"], "Try N3", "C1", "Chương 1")

    def add(content, task="", qtype="TN", options=("a", "b", "c", "d"), passage="", tags="", answer="A"):
        return app.insert_q(app.clean_question({"chapter_id": ch, "bloom": 2, "qtype": qtype, "content": content,
                                                "options": list(options), "answer": answer, "passage": passage,
                                                "tags": tags, "task": task, "status": "da_duyet"}))
    return {"ch": ch, "add": add}


def task_of(qid):
    return app.q("SELECT task FROM questions WHERE id=?", (qid,), True)[0]["task"]


def test_save_guesses_task_and_keeps_chosen_one(bank):
    auto = bank["add"]("6時（　）起きます。", options=("に", "で", "と", "の"))
    chosen = bank["add"]("6時（　）寝ます。", task="np-chon", options=("に", "で", "と", "の"))
    assert task_of(auto) == "np-trotu" and task_of(chosen) == "np-chon"
    with pytest.raises(ValueError, match="không có trong danh mục"):
        app.clean_question({"chapter_id": bank["ch"], "bloom": 1, "qtype": "TN", "content": "x",
                            "options": ["a", "b"], "answer": "A", "task": "khong-co"})


def test_task_with_tag_adds_tag(bank):
    qid = bank["add"]("筆者は旅行が好きだ。", task="doc-tfng", passage="私は旅行が好きです。")
    row = app.q("SELECT tags, options FROM questions WHERE id=?", (qid,), True)[0]
    assert ",TFNG," in row["tags"] and app.json.loads(row["options"]) == list(TFNG)


def test_section_tasks_filter_single_and_passages(bank):
    add = bank["add"]
    for k in range(3):
        add(f"文{k}（　）行きます。", task="np-trotu", options=("に", "で", "と", "の"))
        add(f"「表現{k}」の意味として正しいものはどれですか。", task="np-ynghia")
    for k in range(2):
        add(f"段落{k + 1} に合う見出しを選びなさい。", task="doc-tieude", passage="段落1 …\n段落2 …")
    add("筆者は何をしましたか。", task="doc-noidung", passage="週末、私は図書館へ行きました。")
    ch = bank["ch"]
    sec = {"matrix": [{"chapter_id": ch, "bloom": 2, "n": 3}], "tasks": ["np-trotu"]}
    items = app.preview_exam([sec])[0]["items"]
    assert len(items) == 3 and {it["task"] for it in items} == {"np-trotu"}
    with pytest.raises(ValueError, match="Không đủ câu"):
        app.preview_exam([{"matrix": [{"chapter_id": ch, "bloom": 2, "n": 4}], "tasks": ["np-trotu"]}])
    both = app.preview_exam([{"matrix": [{"chapter_id": ch, "bloom": 2, "n": 6}], "tasks": ["np-trotu", "np-ynghia"]}])
    assert len(both[0]["items"]) == 6
    # đoạn văn: chỉ lấy đoạn có câu đúng dạng bài; đoạn hiểu nội dung bị bỏ qua
    out = app.preview_exam([{"passages": [{"chapter_id": ch, "n": 1}], "tasks": ["doc-tieude"]}])[0]["items"]
    assert len(out) == 2 and all(it["task"] == "doc-tieude" for it in out)
    with pytest.raises(ValueError, match="Không đủ đoạn"):
        app.preview_exam([{"passages": [{"chapter_id": ch, "n": 2}], "tasks": ["doc-tieude"]}])


def test_stats_counts_by_task(bank):
    bank["add"]("文（　）行きます。", task="np-trotu", options=("に", "で", "と", "の"))
    bank["add"]("段落1 に合う見出しを選びなさい。", task="doc-tieude", passage="段落1 …")
    c = next(x for x in app.stats()["chapters"] if x["id"] == bank["ch"])
    assert c["counts_by_qt_task"]["TN|np-trotu"][2] == 1
    assert c["passages_by_qt_task"]["TN|doc-tieude"] == 1


def test_bulk_set_task_and_replacement_keeps_task(bank):
    a = bank["add"]("文A（　）行きます。", task="np-trotu", options=("に", "で", "と", "の"))
    b = bank["add"]("文B（　）行きます。", task="np-trotu", options=("に", "で", "と", "の"))
    app.bulk_update([b], task="np-chon")
    assert task_of(b) == "np-chon"
    with pytest.raises(ValueError, match="Không còn câu thay thế"):
        app.find_replacement({"chapter_id": bank["ch"], "bloom": 2, "qtype": "TN", "task": "np-trotu"}, {a}, {})
    app.bulk_update([b], task="_none")
    assert task_of(b) == ""


def test_custom_task_type(bank):
    code = app.add_task_type({"name": "Chọn câu trả lời phù hợp hội thoại", "skill": "hoithoai", "qtype": "TN"})
    assert code == "hoithoai-rieng1" and app.resolve_task("chọn câu trả lời phù hợp hội thoại") == code
    with pytest.raises(ValueError, match="cùng tên"):
        app.add_task_type({"name": "Chọn câu trả lời phù hợp hội thoại", "skill": "hoithoai"})
    qid = bank["add"]("A：こんにちは。B：（　）", task=code)
    with pytest.raises(ValueError, match="Còn 1 câu"):
        app.delete_task_type(code)
    with pytest.raises(ValueError, match="dựng sẵn"):
        app.delete_task_type("kanji-doc")
    app.bulk_update([qid], task="_none")
    app.delete_task_type(code)
    assert code not in app.task_map()


def test_import_export_task_column(bank):
    tb = app.q("SELECT textbook_id FROM chapters WHERE id=?", (bank["ch"],), True)[0]["textbook_id"]
    rows = [app.IMPORT_HEADERS,
            ["C1", "文（　）行きます。\nA. に\nB. で", "A", "TN", "2", "", "", "", "", "", "", "", "np-chon"],
            ["C1", "文X（　）行きます。\nA. に\nB. で", "A", "TN", "2", "", "", "", "", "", "", "", ""]]
    res = app.parse_rows([rows], tb, docx=False)
    assert res["added"] == 2, res
    tasks = sorted(r["task"] for r in app.q("SELECT task FROM questions", fetch=True))
    assert tasks == ["np-chon", "np-trotu"]
    csv_text = app.export_questions_csv(tb).decode("utf-8-sig")
    assert "Dạng bài" in csv_text.splitlines()[0] and "np-chon" in csv_text


def test_upgrade_classifies_existing_questions(tmp_path, monkeypatch):
    """CSDL cũ chưa có cột task: lần mở app đầu tiên tự gán dạng bài cho câu đã có."""
    monkeypatch.setattr(app, "DB_PATH", str(tmp_path / "data" / "qbank.db"))
    monkeypatch.setattr(app, "TPL_DIR", str(tmp_path / "tpl"))
    monkeypatch.setattr(app, "BACKUP_DIR", str(tmp_path / "data" / "backup"))
    os.makedirs(tmp_path / "data")
    app.init_storage()
    ch = app.ensure_chapter(["KJ"], "Kanji Challenge N5", "K1")
    app.insert_q(app.clean_question({"chapter_id": ch, "bloom": 1, "qtype": "TN", "content": "次の漢字の読み方を選びなさい。山",
                                     "options": ["やま", "かわ"], "answer": "A"}))
    with sqlite3.connect(app.DB_PATH) as conn:  # giả lập CSDL bản cũ: bỏ cột task
        conn.execute("ALTER TABLE questions DROP COLUMN task")
    app.init_storage()
    assert app.q("SELECT task FROM questions", fetch=True)[0]["task"] == "kanji-doc"
