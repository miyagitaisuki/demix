# giaotrinh — định dạng bài học chung (lesson.json)

Bước đầu của "Xưởng học liệu": đọc **bản thảo giáo trình riêng (.docx)** thành một file `lesson.json`,
từ đó xuất sang các công cụ khác (textts, qbank, eduscan, imeji_dominion) mà không phải soạn lại.

Bản thảo đã dùng style `GT_*` và furigana bằng ruby thật của Word, nên đọc thẳng file Word, **không cần OCR**.

> Repo demix là công khai: **không commit bản thảo giáo trình hay `lesson.json` sinh ra từ bản thảo.**
> Thư mục `out/` đã nằm trong `.gitignore`.

## Chạy

```
pip install python-docx
cd giaotrinh
python -m gtlesson GIAOTRINH.docx -o out --textts
```

Kết quả: `out/lesson.json` và (với `--textts`) `out/textts/*.txt` + `out/textts/roles.yaml`.
Dòng tóm tắt in số hội thoại, từ, ngữ pháp, bài tập, số câu thiếu đáp án và số khối chưa đọc được của từng bài.

Kiểm thử: `python -m pytest giaotrinh/tests` (dùng file Word tổng hợp, không cần bản thảo thật).

## Quy ước bản thảo mà bộ đọc dựa vào

| Style | Dùng cho |
|---|---|
| `GT_LessonTitle` | Tên bài `phần Nhật　phần Việt`; cũng là "Nhân vật", "目次", "付録" |
| bảng 1 dòng `第N課` | Số bài (đặt ngay trước tên bài) |
| `GT_Section` | Mục `1 かいわ` … `9 じこひょうか`; phụ lục `A`…`F` |
| `GT_SubHead` | Hội thoại `Ⅰ　tình huống ♪1-01`, ngữ pháp `1.　mẫu　　nghĩa`, phần luyện tập `Ⅰ　mẫu` |
| `GT_Dialogue` | `số<TAB>người nói：<TAB>lời thoại` |
| `GT_Instruction` | Bài tập `A.　đề bài  ① Cơ bản  ♪1-04` |
| `GT_Exercise` | Câu `(1)<TAB>…` hoặc `1.<TAB>…`; dòng chỉ có `＿` là dòng viết |
| `GT_Example`, `GT_Note`, `GT_Body`, `GT_LessonGoal` | Ví dụ, chú ý, giải thích, mục tiêu |
| Phụ lục E `♪1-04　第1課 れんしゅう Ⅰ-A` | Lời thoại phần nghe |
| Phụ lục F `れんしゅう Ⅰ-A: (1) … (2) …` | Đáp án, tự gắn vào câu tương ứng |

Bỏ qua: phụ lục A–D (từ vựng, ngữ pháp, chia động từ, chữ Hán) vì đã có trong từng bài.
Khối không nhận ra được nằm trong `lessons[].unparsed` để không mất nội dung.

## Định dạng lesson.json (`schema: "imeji-lesson/1"`)

Furigana thống nhất một kiểu: **`{漢字|かな}`** (giống qbank và imeji_dominion).
`gtlesson.furigana` có `plain()`, `kana()`, `to_textts()` (sang `漢字{かな}` của textts).

```jsonc
{
  "schema": "imeji-lesson/1",
  "characters": [{"name": "ミン", "profile": "…", "voice_hint": "Nam trẻ, nói chậm, rõ"}],
  "audio_scripts": [{"audio": "1-04", "lesson": 1, "ref": "れんしゅう Ⅰ-A", "text": "(1) … (2) …"}],
  "lessons": [{
    "no": 1, "title_ja": "{北海道|ほっかいどう}の…", "title_vi": "Truyền đạt tin tức…",
    "goals": ["Có thể …"],
    "dialogues": [{"part": "Ⅰ", "situation": "…", "audio": "1-01",
                   "lines": [{"n": 1, "speaker": "ゆい", "text": "…"}]}],
    "vocab_audio": "1-03",
    "vocab": [{"word": "{物産展|ぶっさんてん}", "meaning": "…", "group": "Danh từ", "in_dialogue": true}],
    "grammar": [{"no": 1, "pattern": "…", "meaning": "…", "form": "…", "explanation": "…",
                 "examples": ["…"], "notes": ["…"]}],
    "notes": ["ひょうげんノート …"],
    "exercises": [{
      "id": "renshuu.1.A",          // renshuu.<số phần>.<nhãn> | matome.<nhãn>
      "section": "renshuu", "part": "Ⅰ", "part_title": "〜そうです…", "label": "A",
      "instruction": "…", "level": 1, "audio": "1-04",   // level: ①=1 ②=2 ③=3
      "example": ["{例|れい}）… → …"], "passage": null,    // passage: văn bản đọc kèm (thông báo…)
      "kind": "short_answer",       // short_answer | survey | free | task
      "items": [{"n": 1, "prompt": "…", "answer": "…"}]
    }],
    "kanji_range": "漢字　Chữ Hán (… #71–75)",
    "kanji": [{"char": "活", "han_viet": "HOẠT", "meaning": "…", "on": ["カツ"], "kun": [],
               "words": [{"word": "{生活|せいかつ}", "reading": "せいかつ", "meaning": "…"}]}],
    "lesson_words": [{"word": "…", "meaning": "…"}],
    "reading": {"title": "…", "audio": "1-05", "text": "…", "questions": [{"n": 1, "prompt": "…", "answer": "…"}]},
    "writing": {"title": "…", "instruction": "…", "lines": 4},
    "culture": {"title": "…", "text": "…"},
    "self_check": ["…"],
    "media": [{"kind": "image", "text": "[HÌNH 1-01] mô tả tranh…"}]
  }]
}
```

Mã `audio` (`1-01`…) là mã ♪ trong sách, dùng làm tên file âm thanh ở mọi công cụ.

## Xuất sang textts

- Mỗi nhân vật trên trang Nhân vật có **số vai cố định cho cả bộ sách** (thứ tự trong bảng), vai `0` là người dẫn.
  Tên so khớp bỏ furigana và không phân biệt katakana/hiragana (`マキ` = `まき`).
- `roles.yaml` chỉ có tên vai và mô tả giọng từ sách; giáo viên chọn giọng cụ thể trong textts một lần.
- Hội thoại: `<mã ♪>.txt`; lời thoại phần nghe (phụ lục E): mỗi ý `(1) (2)…` một câu, nghỉ 1,5 giây.

## Trạng thái và bước tiếp

- ĐÃ LÀM: docx → lesson.json; lesson.json → kịch bản textts. Đã chạy trên bản thảo Cuốn 8 (8 bài):
  đủ 8 nhân vật, 16 hội thoại, 17 mục ngữ pháp, 117 câu bài tập/đọc hiểu đều gắn được đáp án, không còn khối chưa đọc.
- CHƯA LÀM: lesson.json → qbank (câu DT/TL theo chương = bài), → eduscan (bài giao tự chấm), → dominion (game, slide).
