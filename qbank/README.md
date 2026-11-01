# QBank - ngân hàng câu hỏi offline

Ứng dụng web nhẹ chạy trên máy cục bộ, không cần internet, chỉ cần thêm thư viện `python-docx`.

## Chạy
```
pip install python-docx        # chỉ cần một lần
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
- **Nút "Kiểm tra trùng"** (tab Ngân hàng) liệt kê các nhóm câu giống nhau. Với mỗi nhóm, chọn câu giữ lại rồi bấm "Gộp": các câu còn lại bị xóa, lịch sử dùng chuyển sang câu giữ lại.
- **"Tự gộp trùng trong cùng chương"**: gộp tự động mọi nhóm trùng nằm trong cùng một chương, giữ câu có mã nhỏ nhất. Trùng khác chương không bị gộp tự động.
- **Lịch sử dùng:** mỗi lần xuất đề ghi lại câu nào đã vào mã đề nào. Khi xuất, app ưu tiên câu chưa dùng, rồi đến câu dùng lâu nhất. Bảng ngân hàng hiện số lần đã dùng. Bỏ tick "Ghi lịch sử dùng câu" ở tab Xuất đề nếu không muốn ghi lại lần xuất đó.
- Xóa một câu sẽ xóa luôn lịch sử dùng của câu đó.

## Định dạng file nhập (chuẩn)
Tải mẫu bằng nút **Tải mẫu nhập** (tab Ngân hàng). File là .docx, mỗi bảng là một nhóm câu:
- **Dòng 1 là tiêu đề cột**, thứ tự cột tùy ý. Cột bắt buộc: `Mã chương`, `Nội dung`, `Đáp án`, `Loại`, `Mức`.
- Cột tùy chọn: `Đoạn văn` (các câu cùng đoạn ghi **giống hệt** nhau; app in đoạn một lần cho cả nhóm), `Lời giải`.
- `Loại`: TN (trắc nghiệm), DS (đúng/sai), DT (điền từ), TL (tự luận). `Mức`: 1–5.
- Phương án TN: mỗi dòng bắt đầu `A.` `B.`…; `#` đầu phương án = cố định vị trí khi xáo.
- DS: mỗi ý một dòng `a)` `b)`…; đáp án `Đ S Đ S`. DT: nhiều đáp án cách nhau bằng `|`.
- Furigana: `{試験|しけん}`.
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
- **Ngân hàng**: thêm/sửa/xóa câu hỏi với 4 loại: **TN** (trắc nghiệm), **DS** (đúng/sai nhiều ý), **DT** (điền từ), **TL** (tự luận). Lọc theo chương, mức Bloom, loại; ô tìm kiếm tìm trong nội dung, phương án, lời giải và đáp án. **Sao lưu** tải về file JSON; **Khôi phục** nạp lại file JSON (thêm vào ngân hàng hiện có). Nhập hàng loạt từ `.docx` có bảng 6 cột: `STT | Mã | Nội dung | Đáp án | Loại | Mức TN`. Trong ô nội dung, phương án TN bắt đầu bằng `A.`, `B.`...; ý của DS bắt đầu bằng `a)`, `b)`...; đáp án DS viết dạng `Đ S Đ S`; đáp án DT có thể có nhiều cách, cách nhau bằng `|`; `Lời giải:` là phần ẩn. Dấu `#` đầu phương án TN giữ nguyên vị trí khi xáo.
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
