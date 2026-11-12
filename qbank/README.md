# QBank - ngân hàng câu hỏi offline

Ứng dụng web nhẹ chạy trên máy cục bộ, không cần internet, chỉ cần thêm thư viện `python-docx`.

## Chạy
```
pip install python-docx openpyxl pykakasi   # chỉ cần một lần (pykakasi tùy chọn, cho furigana tự động)
python3 app.py                 # rồi mở http://127.0.0.1:8765
```

## Cấu trúc phân loại
**Học phần → Giáo trình → Chương → Câu hỏi.** Mã chương chỉ cần duy nhất trong một giáo trình, nên hai giáo trình có thể cùng có "B1".
- Tạo học phần, giáo trình, chương trong tab **Cấu trúc**. Mục chỉ xóa được khi không còn dữ liệu bên trong.
- Nhập Word: chọn giáo trình ở bộ lọc trước. Cột "Mã" phải là mã chương đã có trong giáo trình đó; mã sai được báo lỗi từng dòng.
- Dữ liệu từ phiên bản cũ được chuyển sang học phần và giáo trình "Chưa phân loại", giữ nguyên mã chương cũ. Hãy gán lại học phần cho đúng.

## Trùng lặp và lịch sử dùng
- **Trùng khi nhập Word hoặc sao lưu:** so khớp nội dung sau khi bỏ hoa/thường, khoảng trắng và dấu câu. Trùng trong cùng chương thì bỏ qua; trùng ở chương khác thì vẫn thêm nhưng báo cảnh báo.
- **Thêm câu bằng tay:** câu trùng trong cùng chương bị từ chối; trùng chương khác được lưu kèm cảnh báo.
- **Nút "Kiểm tra trùng"** (ribbon tab Ngân hàng, nhóm "Rà soát") liệt kê các nhóm câu giống nhau. Với mỗi nhóm, chọn câu giữ lại rồi bấm "Gộp": các câu còn lại bị xóa, lịch sử dùng chuyển sang câu giữ lại.
- **"Tự gộp trùng trong cùng chương"**: gộp tự động mọi nhóm trùng nằm trong cùng một chương, giữ câu có mã nhỏ nhất. Trùng khác chương không bị gộp tự động.
- **Lịch sử dùng:** mỗi lần xuất đề ghi lại câu nào đã vào mã đề nào. Khi xuất, app ưu tiên câu chưa dùng, rồi đến câu dùng lâu nhất. Bảng ngân hàng hiện số lần đã dùng. Bỏ tick "Ghi lịch sử dùng câu" ở tab Xuất đề nếu không muốn ghi lại lần xuất đó.
- Xóa một câu sẽ xóa luôn lịch sử dùng của câu đó.

## Xuất đề theo tỷ lệ mức Bloom
- Ở bước "Số câu và đoạn", đổi phần sang **Theo tỷ lệ %**, nhập **tổng số câu**, chọn tỷ lệ (0–100%, bước 10%) cho từng mức Biết → Đánh giá, rồi tick các chương muốn lấy câu.
- Tổng tỷ lệ phải bằng 100% và đủ câu cho từng mức thì mới bấm **Chốt** được; chưa chốt không qua bước tiếp. Số câu được chia đều cho các chương đã tick (phần dư chia cho chương còn câu).
- Chọn câu: ưu tiên câu chưa ra đề lần nào, rồi đến câu ra đề ít lần nhất. Câu thuộc đoạn văn không được chọn lẻ.
- **Seed:** mỗi lần xuất có một seed (hiện trên thông báo sau khi tải). Lưu trong lịch sử xuất cùng danh sách câu đã chọn. Nhập lại seed cùng lịch sử dùng để chọn lại đúng các câu.
- Thống kê → "Chưa từng ra đề" cho biết số câu còn có thể chọn không trùng, theo chương × mức Bloom và dạng câu.

## Reset tình trạng ra đề
Ribbon tab Ngân hàng, nhóm "Reset ra đề": reset câu đang tick, chương đang lọc, hoặc toàn bộ CSDL (phải gõ `RESET`). Reset chỉ xóa số lần đã ra đề, không xóa câu hay file đề đã xuất. Trước khi xóa, app ghi bản sao lưu vào thư mục `data/backup` (`truoc_reset_<thời điểm>.json` chứa lịch sử bị xóa; khi reset toàn bộ còn có `qbank_<ngày>.json`).

## Định dạng file nhập (chuẩn)
Tải mẫu bằng nút **Tải mẫu nhập** (ribbon tab Ngân hàng, nhóm "Nhập / Xuất"). File là .docx, mỗi bảng là một nhóm câu:
- **Dòng 1 là tiêu đề cột**, thứ tự cột tùy ý. Cột bắt buộc: `Mã chương`, `Nội dung`, `Đáp án`, `Loại`, `Mức`.
- Cột tùy chọn: `Đoạn văn` (các câu cùng đoạn ghi **giống hệt** nhau; app in đoạn một lần cho cả nhóm), `Lời giải`.
- `Loại`: TN (trắc nghiệm), DS (đúng/sai), DT (điền từ), TL (tự luận). `Mức`: 1–5.
- Phương án TN: mỗi dòng bắt đầu `A.` `B.`…; `#` đầu phương án = cố định vị trí khi xáo.
- DS: mỗi ý một dòng `a)` `b)`…; đáp án `Đ S Đ S`. DT: nhiều đáp án cách nhau bằng `|`.
- Furigana: `{試験|しけん}`. Khi xuất đề chọn "tự thêm" thì app tự thêm furigana cho chữ Hán chưa có cú pháp (cần `pip install pykakasi`). Cách đọc tự sinh có thể sai với từ đọc đặc biệt, nên kiểm tra và viết tay `{chữ|cách đọc}` để ghi đè.
- `Mã chương` phải đã tồn tại trong giáo trình được chọn ở bộ lọc.

## Mẫu chính thức (từ file gốc của trường)
- `mau_goc/` giữ file Word gốc: `de_goc_mau.docx` (đề) và `dap_an_goc_mau.docx` (đáp án, đã đổi từ .doc).
- `tools/build_templates.py` dựng lại `templates/de_thi_chuan.docx` và `templates/dap_an_chuan.docx` từ file gốc: thay các giá trị thay đổi bằng placeholder, giữ nguyên khung, tên trường, khoa và chân trang. Chạy lại sau khi sửa file gốc: `python3 tools/build_templates.py`.
- Placeholder trong mẫu đề: `{{MA_DE}} {{HOC_KY}} {{NAM_HOC}} {{TEN_HP}} {{MA_HP}} {{SO_TC}} {{KHOA_HOC}} {{THOI_GIAN}} {{SO_TRANG}} {{CAU_HOI}}`. Mẫu đáp án thêm `{{BANG_DAP_AN}}`.
- "(Đáp án gồm 1 trang)" và "Trang 1/1" là chữ cố định trong mẫu đáp án.

## Định dạng xuất đề
- **Mẫu đề** (`de_thi_chuan.docx`) và **mẫu đáp án** (`dap_an_chuan.docx`) theo định dạng trường. Trường, khoa và bố cục là cố định trong mẫu. Các trường thay đổi được nhập khi xuất: tên học phần, mã HP, số tín chỉ, khóa, học kỳ, năm học, thời gian, mã đề.
- Mỗi phần là một khối câu theo dạng, có tiêu đề (ví dụ `I. CHỌN CÁCH SỬ DỤNG NGỮ PHÁP ĐÚNG NHẤT (CLO1)`).
- Đáp án theo lưới 10 câu mỗi dòng.
- Một câu không được chọn hai lần trong cùng một đề.
- **Furigana:** chọn "Chữ đọc phía trên (Word)" (mặc định) hoặc "Chữ(cách đọc) cùng dòng".

## Chức năng
- **Ngân hàng**: thêm/sửa/xóa câu hỏi với 4 loại: **TN** (trắc nghiệm), **DS** (đúng/sai nhiều ý), **DT** (điền từ), **TL** (tự luận). Lọc theo chương, mức Bloom, loại; ô tìm kiếm tìm trong nội dung, phương án, lời giải và đáp án. **Sao lưu** tải về file JSON; **Khôi phục** nạp lại file JSON (thêm vào ngân hàng hiện có). Nhập hàng loạt từ `.docx` có bảng 6 cột: `STT | Mã | Nội dung | Đáp án | Loại | Mức TN`. Trong ô nội dung, phương án TN bắt đầu bằng `A.`, `B.`...; ý của DS bắt đầu bằng `a)`, `b)`...; đáp án DS viết dạng `Đ S Đ S`; đáp án DT có thể có nhiều cách, cách nhau bằng `|`; `Lời giải:` là phần ẩn. Dấu `#` đầu phương án TN giữ nguyên vị trí khi xáo. Tick nhiều câu để đổi chương/mức Bloom hàng loạt, hoặc **Xóa các câu đã chọn** (có nút Hoàn tác).
- **Thống kê**: số câu theo học phần, giáo trình, chương × mức Bloom, tổng theo loại.
- **Xuất đề theo phần**: mỗi phần có tiêu đề (ví dụ `Phần I. Trắc nghiệm`), dạng câu (TN, DS, DT, TL hoặc mọi dạng), tùy chọn xáo câu và phương án, và ma trận riêng chương × mức Bloom. Bỏ tick xáo để giữ thứ tự câu trong ngân hàng. Phần không có số câu sẽ bị bỏ qua. Số câu được đánh liên tục toàn đề. Ô thiếu câu được tô đỏ. Có thể chọn template `.docx` trong `templates/`, kèm đáp án trong file hoặc tách ra `DapAn_xxx.docx`. Kết quả là một file `.zip`.

## Template
Đặt file `.docx` vào `templates/`, dùng các placeholder:

| Placeholder | Nội dung |
|---|---|
| `{{MON}}` `{{TEN_DE}}` `{{THOI_GIAN}}` | Lấy từ form xuất đề |
| `{{MA_DE}}` `{{SO_CAU}}` | Mã đề và số câu |

| `{{CAU_HOI}}` | Đoạn đánh dấu vị trí chèn câu hỏi (đoạn này sẽ bị thay thế) |

Hai template mẫu (`mau_chuan.docx`, `mau_gon.docx`) được tạo tự động lần đầu chạy.

## Mã câu hỏi
Mỗi câu có mã ổn định, dạng `HọcPhần-GiáoTrình-Chương-Số`, ví dụ `N5-MN1-B03-001`. Mã được gán khi tạo và không đổi. Đặt mã học phần (N5) và giáo trình (MN1) khi tạo để mã dễ đọc; nếu bỏ trống, app dùng `HP<id>` và `GT<id>`. Có thể tìm câu theo mã trong ô tìm kiếm.

## Dữ liệu
Lưu trong `data/qbank.db` (SQLite). Sao lưu bằng cách sao chép file này.

## Hướng cải tiến

### 1. Logic chức năng
**Nên làm sớm**
- ✅ Lưu cấu trúc đề có tên, dùng lại cho kỳ sau (chỉ cần chọn và xuất). *Lưu trong tab Xuất đề; không lưu mã đề và thông tin đề.*
- ✅ Gắn thông tin đề (tên HP, mã HP, số tín chỉ, khóa, học kỳ, năm học, thời gian) vào học phần: chọn học phần trong tab Xuất đề để tự điền, bấm "Lưu thông tin này cho học phần" để ghi lại.
- ✅ Lưu lịch sử các lần xuất: mã đề, ngày, tên đề. Tải lại file .zip (đề và đáp án) hoặc xóa khỏi lịch sử. *Chưa lưu danh sách câu theo từng mã đề.*
- ✅ Xem trước đề trên web, cho phép đổi một câu cụ thể mà không chọn lại toàn bộ. *Nút "Xem trước đề" ở tab Xuất đề; xuất dùng đúng các câu đã xem.*
- ✅ Đoạn văn là đối tượng riêng: mỗi đoạn có mã (ví dụ `N4-MN1-B01-P001`), câu gắn vào đoạn qua form câu hỏi (chọn đoạn có sẵn hoặc nhập mới). Danh sách đoạn ở tab Cấu trúc: sửa nội dung, xóa khi không còn câu. *Ma trận theo đoạn: trong từng phần, chọn số đoạn của một chương.*

**Nên làm tiếp**
- ✅ Sửa tên và mã học phần, giáo trình, chương (nút "Sửa" ở tab Cấu trúc).
- ✅ Chuyển câu sang chương khác và sửa mức Bloom hàng loạt: tick các câu ở tab Ngân hàng, chọn chương hoặc mức rồi áp dụng. Chỉ đổi khi không gây trùng trong chương đích.
- ✅ Độ khó thực tế: nhập tỷ lệ đúng từng câu sau kỳ thi ở tab Thống kê; xem tỷ lệ TB theo mức Bloom và câu bất thường (≤20% hoặc ≥95%). *Ngưỡng bất thường là cố định, chưa chỉnh được.*
- ✅ Cảnh báo câu gần giống (từ 90%): nút "Tìm câu gần giống" ở tab Ngân hàng, chỉ gợi ý, không tự xóa.
- ✅ Cân bằng đáp án khi xáo: trong mỗi mã đề, đáp án được rải đều A/B/C/D. Kiểm tra: 200 mã đề × 12 câu cho 600 lần mỗi chữ.

### 2. Giao diện (GUI)
- ✅ Tab Ngân hàng hai cột: cây học phần → giáo trình → chương bên trái (có số câu), bấm để lọc; danh sách câu bên phải. Thu gọn xuống một cột trên màn hình hẹp.
- Form sửa câu dạng panel bên phải hoặc cửa sổ nổi, có xem trước furigana khi gõ `{漢字|かな}`.
- Hiển thị furigana thật trong danh sách (thẻ `<ruby>`), không hiện cú pháp thô.
- Ma trận xuất đề gọn: chỉ hiện chương có câu đúng dạng của phần; ô hiện "đã chọn / có sẵn" (ví dụ `3/12`); có tổng số câu mỗi phần và cả đề.
- ✅ Chia tab Xuất đề thành 6 bước: phần và dạng câu → số câu và đoạn → thông tin đề → mẫu và mã đề → xem trước → xác nhận và xuất. Mỗi bước kiểm tra trước khi sang bước sau; bước 6 có tóm tắt và phải tích xác nhận. *Nhập Word: bước xem trước và xác nhận ghi rõ số câu sẽ nhập vào giáo trình nào.*
- Phân trang hoặc cuộn ảo khi ngân hàng lên đến hàng nghìn câu.

### 3. Trải nghiệm người dùng (UX)
**Giảm thao tác**
- ✅ Phím tắt: `Ctrl+S` lưu câu (khi đang sửa), `N` thêm câu mới, `/` vào ô tìm kiếm.
- ✅ Nhớ thông tin đề và mẫu xuất đề gần nhất (lưu trên trình duyệt). Còn lại: học phần, giáo trình đang lọc.
- ✅ Nhập Word: xem trước từng dòng (sẽ nhập, cảnh báo, bỏ qua, lỗi) trước khi xác nhận. *Chưa có kéo thả.* các dòng sẽ nhập (dòng lỗi tô đỏ) trước khi xác nhận.

**Giảm lỗi**
- ✅ Thông báo dạng toast (xóa, xuất, sửa, nhập). *Lỗi nhập Word hiển thị trong bảng xem trước.*
- ✅ Hoàn tác sau khi xóa câu hoặc gộp trùng từng nhóm (nút "Hoàn tác" trong vài giây, giữ lại mã câu). *Tự gộp hàng loạt chưa có hoàn tác; lịch sử dùng của câu đã xóa hoặc đã gộp không khôi phục.*
- ✅ Kiểm tra trước khi xuất: chặn khi thiếu câu hoặc chưa chọn phần nào; hỏi xác nhận khi chưa nhập tên HP, mã HP hoặc năm học. *Phần trống: còn để làm.*
- ✅ Tự sao lưu dữ liệu mỗi ngày vào `data/backup/`, giữ 7 bản gần nhất (kiểm tra mỗi giờ khi server chạy).

**Dễ làm quen**
- Màn hình bắt đầu khi ngân hàng trống: hướng dẫn 3 bước (tạo học phần → nhập câu → xuất đề), kèm nút tải mẫu nhập.
- Nút nạp dữ liệu mẫu tiếng Nhật để dùng thử.
- Gợi ý ngay dưới ô đáp án theo loại câu đang chọn.

### 4. Đóng gói và vận hành
- Đóng gói PyInstaller thành file chạy trên Windows, không cần cài Python; mở app tự mở trình duyệt.
- Đưa tên trường và khoa vào file cấu hình riêng để khoa khác dùng được.

### Thứ tự đề xuất
| Ưu tiên | Việc | Lý do |
|---|---|---|
| 1 | ✅ Lưu cấu trúc đề và thông tin học phần, lịch sử xuất |
| 2 | ✅ Xem trước và đổi từng câu | Giảm xuất sai, phải làm lại |
| 3 | ✅ Đoạn văn là đối tượng riêng | Đề tiếng Nhật dùng đoạn văn nhiều |
| 4 | ✅ Bố cục hai cột; chia xuất đề thành bước | Thao tác hằng ngày nhanh hơn |
| 5 | ✅ Xem trước khi nhập Word, toast, hoàn tác | Giảm lỗi khi nhập số lượng lớn |
| 6 | Đóng gói Windows | Để giáo viên khác dùng được |

## Đề xuất từ thử nghiệm người dùng
Thử luồng từ CSDL trống đến khi tải file đề (học phần → chương → thêm câu → nhập Word → sửa hàng loạt → wizard xuất đề → lịch sử → thống kê). Các chỗ gây khó chịu, xếp theo mức độ:

1. ✅ **Lưu câu không có phản hồi.** Bấm Lưu thì form đóng, không có thông báo. *Đã sửa: toast "Đã lưu câu N4-…" (kèm cảnh báo trùng nếu có).*
2. ✅ **Thông báo nhập Word dùng số nội bộ.** Ví dụ "trùng câu #1" trong khi giao diện chỉ hiện mã như `N4-MN1-B01-001`. *Đã sửa: mọi thông báo trùng, gộp, đổi chương, khôi phục dùng mã câu.*
3. ✅ **Lịch sử xuất khó phân biệt.** Tên đề để trống nên mỗi dòng chỉ có ngày và mã đề. *Đã sửa: tên mặc định "học phần · mã đề · thời gian", ví dụ "Tiếng Nhật N4 · mã 001–002 · 90 phút". Tên người dùng đặt vẫn được giữ.*
4. ✅ **Ô số câu ở bước 1 hiện đủ mọi chương × Bloom.** Phần lớn ô trống, người dùng dễ nhập vào ô không có câu rồi bị chặn mà không hiểu vì sao. *Đã sửa: ô không có câu bị khóa và làm mờ.*
5. ✅ **Đổi tên học phần, giáo trình, chương dùng `prompt()`.** Chương phải nhập qua hai hộp liên tiếp, hủy giữa chừng thì không rõ đã lưu gì. *Đã sửa: sửa tại chỗ trong cây, có nút Lưu/Hủy; Enter để lưu, Esc để hủy.*
6. ✅ **Đoạn văn dễ tạo trùng.** Sửa một ký tự trong ô đoạn đã chọn sẽ tạo đoạn mới mà không cảnh báo. *Đã sửa: khi đoạn mới gần giống (từ 90%) đoạn có sẵn trong chương, app hỏi xác nhận trước khi tạo. Khi so sánh bỏ qua dấu câu và hoa/thường nên có thể hiện 100% với khác biệt chỉ là dấu.*
7. ✅ **Tên file xuất luôn là `de_thi.zip`.** Xuất nhiều đề trong ngày sẽ trùng tên. *Đã sửa: tên file `De_<mã HP>_<mã đề đầu-cuối>_<yyyymmdd>.zip`, ví dụ `De_N4_001-010_20261101.zip`, dùng cho cả tải mới và tải lại từ lịch sử.*

Chưa thử trong lần này: hoàn tác trên giao diện, bố cục trên màn hình hẹp, và cây hai cột với nhiều học phần.

## Đề xuất chỉnh sửa từ kiểm thử 2026-11-02
Chi tiết từng lỗi và cách tái hiện: `docs/bao_cao_kiem_thu_2026-11-02.md`. Trạng thái sửa được ghi ở mục cuối.

### Lỗi (B01–B32)
- **Critical:** B01 ô lọc chương làm bảng trống; B02 sao lưu/khôi phục thiếu dữ liệu; B03 câu TN hơn 8 phương án làm hỏng xuất đề; B04 CSRF cho phép trang lạ xóa dữ liệu; B05 so trùng bỏ ký hiệu toán nên báo trùng sai.
- **High:** B06 "Tìm câu gần giống" chậm (42 s/3.000 câu); B07 cột Mức sai bị đoán thành 1; B08 đáp án Tự luận bị xóa; B09 sửa tạo được câu trùng; B10 lưu câu đã bị xóa vẫn báo thành công; B11 xuất nhiều mã đề đếm nhiều lần dùng; B12 xuất lần hai dùng lại câu cũ.
- **Medium:** B13 "Bắt đầu từ 0" đổi thành 1; B14 thiếu "Đề thi gồm __ trang"; B15 thông tin đề không kiểm tra; B28 "Sửa" tải cả ngân hàng; B29 màn hình hẹp bị cuộn ngang; B30 F5 mất cấu trúc đề đang soạn.
- **Low:** B16 hủy xóa vẫn báo "Đã lưu"; B17 bộ lọc không nhất quán; B18 tìm `_` `%` khớp tất cả; B19–B20 thông báo lỗi kỹ thuật; B21 chương còn đoạn văn mồ côi không xóa được; B22 đổi chương giữ mã cũ; B23 đáp án đúng/sai hiện "DDD"; B24 tên file thừa "De"; B25 tiền tố mã đề không lọc ký tự; B26 nhập kết quả không nhận dán từ Excel; B27 đoạn văn gần giống khi nhập Word không cảnh báo; B31 lịch sử xuất phình; B32 thiếu tiến trình cho thao tác lâu.

### Cải tiến UI (top 10)
1. Thanh công cụ gọn: "+ Câu mới" nổi bật, gom Nhập/Sao lưu vào menu "Dữ liệu", Rà soát vào menu "Rà soát"; thanh hàng loạt chỉ hiện khi đã chọn.
2. Form câu mở thành panel bên phải, không đẩy danh sách.
3. Furigana hiển thị bằng `<ruby>` trong danh sách, xem trước và form.
4. Đáp án TN chọn bằng radio, có thêm/bớt phương án (tối đa 8).
5. Bảng ngân hàng dạng thẻ trên màn hình hẹp, nội dung dài thu gọn.
6. Chọn file kéo thả có nhãn rõ.
7. Đáp án đúng/sai hiện "Đ S Đ"; nhãn dạng câu viết đủ.
8. Bước 2 chia tab theo phần; thanh tổng cố định "Đề: N câu".
9. Chỉ báo đang tải và khóa nút cho nhập Word, khôi phục, gần giống, xuất đề.
10. Lịch sử xuất hiện khoảng mã "001–100", thêm cột học phần.

### Cải tiến UX (top 10)
1. Màn hình bắt đầu 3 bước, kèm tải mẫu và dữ liệu mẫu.
2. Chọn giáo trình đích ngay trong luồng nhập Word.
3. Bộ lọc nhất quán (tự chạy, Enter để tìm); lưu bộ lọc lên URL.
4. "Lưu và thêm câu tiếp" giữ chương, Bloom, loại.
5. Lưu nháp wizard xuất đề và định tuyến theo URL để F5/Back không mất việc.
6. Sau khi xuất: xóa bản xem trước, bỏ tích xác nhận.
7. Mọi thông báo lỗi kỹ thuật bằng tiếng Việt dễ hiểu.
8. Mã chương khi nhập Word không phân biệt hoa/thường; gợi ý mã gần đúng.
9. Kiểm tra thông tin đề (số, khoảng hợp lệ); thêm ô số trang.
10. Nhập kết quả thi nhận dán từ Excel; ghi lại cùng đợt thì hỏi thay thế.

### Tính năng nên bổ sung (top 10)
1. Sao lưu và khôi phục đầy đủ (cả file `.db`, chế độ thay thế).
2. Thùng rác thay cho hoàn tác 8 giây.
3. Đáp án và thang điểm cho câu tự luận.
4. Xem trước bản in trước khi xuất.
5. Bảng đặc tả đề (CLO × Bloom × số câu × điểm).
6. Nhập Excel/CSV và dán từ clipboard.
7. Hình ảnh và âm thanh trong câu.
8. Thẻ (tag) tự do: chủ điểm ngữ pháp, cấp JLPT.
9. Phân tích câu hỏi từ bảng điểm sinh viên.
10. Trạng thái câu và người duyệt (Nháp → Đã duyệt → Ngưng dùng).

### Nên bỏ hoặc gộp (top 5)
1. Bộ 3 ô lọc Học phần/Giáo trình/Chương trùng với cây bên trái.
2. Tùy chọn "kèm đáp án cuối đề".
3. Hai mẫu demo tự tạo `mau_chuan.docx`, `mau_gon.docx`.
4. "Tự gộp trùng cùng chương" chạy ngay không hoàn tác; nên có xem trước.
5. Bốn phần mặc định TN/DS/DT/TL tự thêm.

### Roadmap
- **Critical:** B01–B05.
- **High:** B06–B12; nút "Sửa" lấy một câu; màn hình bắt đầu.
- **Medium:** B13–B15, B29, B30; panel sửa câu, ruby, radio đáp án; nhập Excel; thùng rác; đáp án và điểm tự luận.
- **Low:** phần còn lại của B16–B32; thẻ; bảng đặc tả; phân tích câu hỏi; trạng thái duyệt; xuất Moodle/QTI; dùng chung qua mạng nội bộ.

### Trạng thái sửa
- **Đã sửa (đợt 1):** B01–B05, B07–B13, B15–B20, B23–B25, B16, B17, B18.
- **Đã sửa (đợt 2):** B14, B21, B26 (kết quả cũ của mã đề được giữ, chỉ thêm dòng mới), B27, B28, B30, B31 (quá 100 bản xuất thì reset lịch sử về 0 trước khi lưu bản mới), B32.
- **Đã sửa (đợt 3):**
  - **Thùng rác (B-tính năng mới):** xóa câu chuyển vào thùng rác, không xóa hẳn. Khôi phục trả lại câu cùng lịch sử dùng và kết quả thi. Có thể khôi phục từng câu hoặc tất cả, xóa hẳn từng câu hoặc làm trống. Câu không khôi phục được (ví dụ trùng nội dung) vẫn nằm lại trong thùng rác. Thùng rác không tự xóa theo thời gian.
  - **Điểm và hướng dẫn chấm:** mỗi câu có thể có điểm (0–10, không bắt buộc), in kèm câu trong đề. Với câu tự luận, ô "Đáp án" là hướng dẫn chấm, được in ở cuối đề và trong file đáp án, kèm tổng điểm. Mã câu không còn bị cấp lại sau khi xóa.
- **Sửa một phần:** B06 (tìm câu gần giống: 8 giây với 3.000 câu, trước đây 42 giây; vẫn tăng theo số cặp). B29 (màn hình 375px giảm từ 608px xuống 468px, vẫn cuộn ngang).
- **Chưa sửa:** B22 (đổi chương giữ mã cũ; cần quyết định về mã khi chuyển chương).
- **Đã làm (đợt 4):**
  - **Nhập CSV và dán từ Excel:** chọn file .csv (dấu phẩy hoặc chấm phẩy, UTF-8 hoặc Windows-1258), hoặc dán các ô đã chọn trong Excel (kèm dòng tiêu đề) vào mục "Dán từ Excel". Vẫn kiểm tra trước khi ghi như nhập Word. File .xlsx chưa đọc trực tiếp: lưu lại thành .csv trong Excel.
  - **CLO:** mỗi câu có thể gắn CLO (ví dụ CLO1). Cột "CLO" trong bảng Word, CSV và dữ liệu dán là tùy chọn.
  - **Bảng đặc tả:** sau khi xem trước đề, bảng CLO × mức Bloom hiện số câu và điểm của các câu đã chọn.
  - **Xem bản in:** mở bản xem trước HTML của mã đề đầu tiên (furigana, điểm, hướng dẫn chấm, bảng đặc tả) trong cửa sổ mới, có nút In. Bố cục khác đôi chút so với file Word xuất ra. Không ghi lịch sử dùng câu.
- **Đã làm (đợt 5):**
  - **Excel:** đọc file `.xlsx` khi nhập (cần `openpyxl`). Nếu chưa cài, nhập `.xlsx` sẽ báo lỗi; CSV và dữ liệu dán không cần thư viện này.
  - **Thẻ (tag):** mỗi câu có thể có nhiều thẻ (ví dụ `ngữ pháp, trợ từ`). Ô "Thẻ" trong thanh lọc ngân hàng lọc theo thẻ và gợi ý thẻ đang dùng. Thẻ đi theo khi sao lưu và nhập. Lọc theo thẻ hiện chỉ có ở ngân hàng, chưa có khi chọn câu để ra đề.
  - **Phân tích câu hỏi từ bảng điểm** (tab Thống kê): tải bảng điểm (.csv, .xlsx hoặc dữ liệu dán). Mỗi dòng là một sinh viên; cột đầu là tên hoặc mã sinh viên; các cột sau có tiêu đề là mã câu; ô điểm là 1 (đúng), 0 (sai) hoặc tỷ lệ từ 0 đến 1. Kết quả gồm độ khó (% đúng), chỉ số phân biệt D (nhóm 27% cao trừ nhóm 27% thấp), tương quan câu–tổng và KR-20. Có thể ghi tỷ lệ đúng vào kết quả của mã đề. Các cột không phải mã câu được bỏ qua và báo lại.
  - **Hình ảnh và âm thanh trong câu:** tối đa 4 file mỗi câu. Ảnh (png, jpg, gif, webp) tối đa 5 MB, được nhúng vào đề Word. Âm thanh (mp3, wav, ogg, m4a) tối đa 15 MB, nằm trong thư mục `Am_thanh/` của file zip và được ghi tên trong đề. File lưu trong `data/media/`. Sao lưu JSON nhúng cả file (base64) nên file sao lưu lớn hơn khi có nhiều file đính kèm.
- **Đã làm (đợt 6):**
  - **Trạng thái duyệt câu:** Nháp (câu mới) → Đã duyệt (được ra đề) → Ngưng dùng (không ra đề, vẫn còn trong ngân hàng). Câu có sẵn trước đợt này được đặt là "Đã duyệt" để đề đang dùng không thiếu câu. Ma trận và xem trước chỉ chọn câu đã duyệt. Đổi trạng thái hàng loạt bằng các nút trong thanh chọn; lọc theo trạng thái ở thanh lọc ngân hàng. Sửa một câu đã duyệt không tự đưa về Nháp: người dùng tự đổi nếu cần.
  - **Xuất Moodle XML:** chọn câu rồi bấm "Xuất Moodle XML"; không chọn thì xuất mọi câu đã duyệt. Trắc nghiệm → multichoice (một đáp án); đúng/sai nhiều ý → mỗi ý là một câu truefalse; điền từ → shortanswer (mọi đáp án chấp nhận); tự luận → essay, hướng dẫn chấm đặt vào graderinfo. Danh mục Moodle theo học phần/giáo trình/chương. Ảnh và âm thanh được nhúng vào câu Moodle. Chưa thử nhập vào Moodle thật: cần kiểm tra trên máy có Moodle trước khi dùng cho lớp.
- **Chưa làm (tính năng mới):** QTI; dùng chung qua mạng nội bộ có đăng nhập; cấu hình tên trường và khoa.

## Tạo file nghe bằng textts
Câu hỏi dạng nghe có thể tạo âm thanh ngay trong form sửa câu, không cần thu âm.
1. Bật API textts trên cùng máy: mở app textts, hoặc chạy `python server.py` trong repo textts (cổng 8000). Muốn dùng VOICEVOX hoặc AivisSpeech thì bật engine đó trước; Edge cần Internet.
2. Trong form sửa câu, bấm **🔊 Tạo file nghe (textts)**. Nút "Lấy từ đoạn văn + câu hỏi" điền sẵn kịch bản (TN có kèm phương án đọc số 1、2、3…).
3. Sửa kịch bản: mỗi dòng một lượt nói `[1] ...`, `[2] ...` (dòng không ghi vai được gán `[1]`), furigana `{漢字|かな}` tự đổi sang cú pháp textts, nghỉ bằng `[pause=800]`.
4. Chọn engine, trình độ (N5 đọc chậm nhất), bấm **Tạo và đính kèm**. File mp3 được lưu vào `data/media` như file tải lên; bấm **Lưu** để gắn vào câu.

5. **Gắn vào đoạn văn:** câu có đoạn văn (ví dụ hội thoại 聴解 dùng cho nhiều câu) thì chọn "Gắn vào: Đoạn văn". File nghe lưu ở đoạn, mọi câu thuộc đoạn dùng chung; chọn đoạn có sẵn ở câu khác sẽ thấy ngay file đó. Khi xuất đề, file của đoạn được ghi một lần trước câu đầu tiên của đoạn và chép vào thư mục `Am_thanh`; bản Moodle gắn file của đoạn vào từng câu.
6. **Sửa và tạo lại:** kịch bản được lưu cùng file. Bấm ✎ cạnh file để mở lại kịch bản, sửa rồi bấm "Tạo lại" để thay file cũ (bấm Lưu để giữ thay đổi).

textts chạy ở địa chỉ khác thì đặt biến môi trường `TEXTTS_API_URL` (mặc định `http://127.0.0.1:8000`).

## Xuất đề nghe trọn bộ theo mã đề
Ở bước tùy chọn xuất đề, tick **Tạo file đề nghe** (cần bật textts bản có `/api/v1/assemble`). Mỗi mã đề có thêm `Am_thanh/De_xxx_nghe.mp3`, ghép đúng thứ tự câu trong mã đề đó:
- Câu thuộc đoạn văn có file nghe: đọc "n番からm番", phát file của đoạn, rồi đến file riêng của từng câu (nếu có), sau đó chừa thời gian làm bài cho cả nhóm.
- Câu lẻ có file nghe: đọc "n番", phát file, chừa thời gian làm bài.
- Câu không có file nghe thì bỏ qua. Chọn được số lần phát (1 hoặc 2), số giây làm mỗi câu và giọng đọc số câu.
