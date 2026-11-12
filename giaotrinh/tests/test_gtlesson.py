"""Kiểm tra bộ đọc giáo trình trên một file Word tổng hợp (không dùng nội dung giáo trình thật)."""
import os
import sys

import docx
import pytest
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from gtlesson import furigana  # noqa: E402
from gtlesson.export_textts import role_map, speaker_role, write_all  # noqa: E402
from gtlesson.parse_docx import parse, split_answers  # noqa: E402

STYLES = ["GT_LessonTitle", "GT_Section", "GT_SubHead", "GT_LessonGoal", "GT_Dialogue", "GT_Body",
          "GT_Example", "GT_Note", "GT_Instruction", "GT_Exercise"]


def add(doc, style, text):
    """Thêm đoạn; {漢字|かな} trong text được ghi thành w:ruby thật như bản thảo."""
    p = doc.add_paragraph(style=style)
    pos = 0
    for m in furigana.RUBY.finditer(text):
        if m.start() > pos:
            p.add_run(text[pos:m.start()])
        p._p.append(parse_xml(
            f'<w:r {nsdecls("w")}><w:ruby><w:rubyPr/><w:rt><w:r><w:t>{m.group(2)}</w:t></w:r></w:rt>'
            f'<w:rubyBase><w:r><w:t>{m.group(1)}</w:t></w:r></w:rubyBase></w:ruby></w:r>'))
        pos = m.end()
    if pos < len(text):
        p.add_run(text[pos:])


def table(doc, rows):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            t.cell(r, c).text = val


@pytest.fixture()
def sample(tmp_path):
    doc = docx.Document()
    for s in STYLES:
        doc.styles.add_style(s, WD_STYLE_TYPE.PARAGRAPH)
    add(doc, "GT_LessonTitle", "とうじょうじんぶつ　Nhân vật")
    table(doc, [["Hình", "Nhân vật", "Hồ sơ", "Giọng audio"],
                ["", "ミン", "Nam, 19 tuổi", "Nam trẻ"],
                ["", "佐藤 まき", "Nữ, 24 tuổi", "Nữ trưởng thành"]])
    table(doc, [["第1課", "B À I 1"]])
    add(doc, "GT_LessonTitle", "{雨|あめ}が　{降|ふ}るそうです　Truyền đạt tin tức")
    add(doc, "GT_SubHead", "Mục tiêu bài học")
    add(doc, "GT_LessonGoal", "◆ Có thể dùng 〜そうです.")
    add(doc, "GT_Section", "1  かいわ　Hội thoại")
    add(doc, "GT_SubHead", "Ⅰ　ミン và マキ nói chuyện.  ♪ 1-01")
    add(doc, "GT_Dialogue", "1\tマキ：\t{明日|あした}は　雨だそうですよ。")
    add(doc, "GT_Dialogue", "2\tミン：\tそうですか。")
    add(doc, "GT_Section", "2  ことば　Từ vựng")
    table(doc, [["Nhóm", "ことば", "Nghĩa", "＊"], ["Danh từ", "雨", "mưa", "＊"], ["Danh từ", "雪", "tuyết", ""]])
    add(doc, "GT_Section", "3  ぶんぽう　Ngữ pháp")
    add(doc, "GT_SubHead", "1.　［Thể ngắn］＋ そうです　　Nghe nói là…")
    table(doc, [["V ＋ そうです。"]])
    add(doc, "GT_Body", "Truyền đạt thông tin.")
    add(doc, "GT_Example", "・雨が　{降|ふ}るそうです。")
    add(doc, "GT_Section", "5  れんしゅう　Luyện tập")
    add(doc, "GT_Note", "Mỗi bài tập có nhãn bậc.")
    add(doc, "GT_SubHead", "Ⅰ　〜そうです　Nghe nói")
    add(doc, "GT_Instruction", "A.　Chuyển sang 〜そうです.  ① Cơ bản  ♪ 1-04")
    add(doc, "GT_Example", "{例|れい}）雨です → 雨だそうです。")
    add(doc, "GT_Exercise", "(1)\t雪です → ＿＿＿＿＿＿＿＿")
    add(doc, "GT_Exercise", "(2)\t{暑|あつ}いです")
    add(doc, "GT_Exercise", "\t→ ＿＿＿＿＿＿＿＿")
    add(doc, "GT_Section", "6  よみかき　Đọc – Viết")
    add(doc, "GT_SubHead", "Ⅰ　{漢字|かんじ}　Chữ Hán")
    table(doc, [["雨", "VŨ (mưa)\nÂm On: ウ\nÂm Kun: あめ, あま", "", ""]])
    table(doc, [["Từ", "Cách đọc", "Nghĩa"], ["大雨", "おおあめ", "mưa to"]])
    add(doc, "GT_SubHead", "Ⅱ　読みもの　Bài đọc  ♪ 1-05")
    table(doc, [["今日は　雨です。"]])
    add(doc, "GT_Instruction", "しつもん　Trả lời câu hỏi.")
    add(doc, "GT_Exercise", "1.\t今日は　どんな　天気ですか。")
    add(doc, "GT_Exercise", "\t→ ＿＿＿＿＿＿")
    add(doc, "GT_SubHead", "Ⅲ　書くれんしゅう　Luyện viết")
    add(doc, "GT_Instruction", "Viết 3 câu.")
    add(doc, "GT_Exercise", "＿＿＿＿＿＿＿＿＿＿")
    add(doc, "GT_Exercise", "＿＿＿＿＿＿＿＿＿＿")
    add(doc, "GT_Section", "7  まとめのれんしゅう　Luyện tập tổng hợp")
    add(doc, "GT_Instruction", "A.　Đọc thông báo.  ② Vận dụng")
    table(doc, [["【お知らせ】\n午後　雨"]])
    add(doc, "GT_Exercise", "1.\t午後は　どうですか。")
    add(doc, "GT_Instruction", "B.　Hỏi bạn.  ③ Giao tiếp")
    table(doc, [["なまえ", "天気"], ["", ""], ["", ""]])
    add(doc, "GT_Section", "9  じこひょうか　Tự đánh giá")
    table(doc, [["Tôi có thể…", "◎ Tốt"], ["dùng そうです", ""]])
    add(doc, "GT_LessonTitle", "{付録|ふろく}　Phụ lục cuối sách")
    add(doc, "GT_Section", "E  スクリプト　Lời thoại phần nghe")
    add(doc, "GT_SubHead", "♪1-04　第1課 れんしゅう Ⅰ-A")
    add(doc, "GT_Example", "(1) 雪だそうです。　(2) {暑|あつ}いそうです。")
    add(doc, "GT_Section", "F  こたえ　Đáp án")
    add(doc, "GT_SubHead", "第1課　雨")
    add(doc, "GT_Example", "れんしゅう Ⅰ-A: (1) 雪だそうです。　(2) {暑|あつ}いそうです。")
    add(doc, "GT_Example", "よみかき Ⅱ しつもん: 1. 雨です。")
    add(doc, "GT_Example", "まとめのれんしゅう A: 1. 雨です。")
    path = tmp_path / "gt.docx"
    doc.save(path)
    return parse(str(path))


def test_lesson_structure(sample):
    assert [c["name"] for c in sample["characters"]] == ["ミン", "佐藤 まき"]
    (L,) = sample["lessons"]
    assert L["no"] == 1
    assert L["title_ja"] == "{雨|あめ}が　{降|ふ}るそうです"
    assert L["title_vi"] == "Truyền đạt tin tức"
    assert L["goals"] == ["Có thể dùng 〜そうです."]
    assert "unparsed" not in L


def test_dialogue_and_vocab(sample):
    L = sample["lessons"][0]
    (d,) = L["dialogues"]
    assert d["audio"] == "1-01" and d["part"] == "Ⅰ"
    assert d["lines"][0] == {"n": 1, "speaker": "マキ", "text": "{明日|あした}は　雨だそうですよ。"}
    assert [v["in_dialogue"] for v in L["vocab"]] == [True, False]


def test_grammar(sample):
    (g,) = sample["lessons"][0]["grammar"]
    assert g["pattern"] == "［Thể ngắn］＋ そうです" and g["meaning"] == "Nghe nói là…"
    assert g["form"] == "V ＋ そうです。"
    assert g["examples"] == ["雨が　{降|ふ}るそうです。"]


def test_exercises_with_answers(sample):
    ex = {e["id"]: e for e in sample["lessons"][0]["exercises"]}
    a = ex["renshuu.1.A"]
    assert a["level"] == 1 and a["audio"] == "1-04"
    assert a["instruction"] == "Chuyển sang 〜そうです."
    assert [i["prompt"] for i in a["items"]] == ["雪です", "{暑|あつ}いです"]
    assert [i["answer"] for i in a["items"]] == ["雪だそうです。", "{暑|あつ}いそうです。"]
    assert ex["matome.A"]["passage"].startswith("【お知らせ】")
    assert ex["matome.A"]["items"][0]["answer"] == "雨です。"
    assert ex["matome.B"]["kind"] == "survey" and ex["matome.B"]["columns"] == ["なまえ", "天気"]


def test_yomikaki(sample):
    L = sample["lessons"][0]
    (k,) = L["kanji"]
    assert (k["char"], k["han_viet"], k["on"], k["kun"]) == ("雨", "VŨ", ["ウ"], ["あめ", "あま"])
    assert k["words"][0]["reading"] == "おおあめ"
    assert L["reading"]["audio"] == "1-05" and L["reading"]["questions"][0]["answer"] == "雨です。"
    assert L["writing"]["lines"] == 2
    assert L["self_check"] == ["dùng そうです"]


def test_audio_script(sample):
    (s,) = sample["audio_scripts"]
    assert (s["audio"], s["lesson"], s["ref"]) == ("1-04", 1, "れんしゅう Ⅰ-A")


def test_textts_export(sample, tmp_path):
    roles = role_map(sample)
    assert speaker_role("マキ", roles) == 2  # katakana trong hội thoại, hiragana ở trang Nhân vật
    files = {os.path.basename(f): f for f in write_all(sample, str(tmp_path / "tt"))}
    text = open(files["1-01.txt"], encoding="utf-8").read()
    assert "[2] 明日{あした}は　雨だそうですよ。" in text
    assert "[1] そうですか。" in text
    script = open(files["1-04.txt"], encoding="utf-8").read()
    assert "[0] 暑{あつ}いそうです。" in script and "[pause=1500]" in script


def test_furigana_helpers():
    t = "{北海道|ほっかいどう}の　{持ち主|もちぬし}と　お{土産|みやげ}"
    assert furigana.plain(t) == "北海道の　持ち主と　お土産"
    assert furigana.kana(t) == "ほっかいどうの　もちぬしと　おみやげ"
    assert furigana.to_textts(t) == "北海道{ほっかいどう}の　持{も}ち主{ぬし}と　お土産{みやげ}"
    assert furigana.to_textts("{お土産|おみやげ}") == "お土産{みやげ}"


def test_split_answers():
    assert split_answers("(1) A。　(2) B。") == ["A。", "B。"]
    assert split_answers("1. ３，０００円です。　2. ５０食です。") == ["３，０００円です。", "５０食です。"]
