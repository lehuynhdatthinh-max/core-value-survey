# Apps Script của khảo sát VHDN

Dự án Apps Script gắn với sheet **phản hồi khảo sát** (web app `.../exec` mà `index.html` gọi). Dự án này KHÔNG tự đẩy lên như SafetyHub:
sửa xong phải dán vào trình soạn thảo Apps Script. Repo này công khai nên chỉ lưu `SurveyRetake.gs`; `Code.gs` bản 4.9.0 được gửi riêng cho HSSE.

| File | Việc |
|---|---|
| `Code.gs` (gửi riêng, không để trong repo công khai) | Toàn bộ backend (tra MSNV, câu hỏi, hàng đợi nộp bài, dashboard…). 4.9.0: `getMyResult` + ghi đè điểm khi làm lại. 4.10.0: `processQueue` nhanh hơn + tự lưu trữ Queue cũ. 4.11.0: giờ nộp theo máy chủ (để gộp theo tháng). 4.12.0: chịu tải 75–100 người cùng lúc. 4.13.0: tab Báo cáo VHDN. |
| `SurveyReport.gs` | Số liệu tab **Báo cáo VHDN**: BIẾT / HIỂU theo tháng, quý, bộ phận; số liệu nhập từ file báo cáo; điểm chuẩn HIỂU. |
| `SurveyRetake.gs` | Mỗi MSNV **một dòng mỗi tháng** trong **Responses** (luôn là lần nộp mới nhất của tháng), điểm tháng này cho trang khảo sát, gộp dòng trùng cũ. |
| `SafetyHubEmployeeRefresh.gs` | Làm mới tra cứu MSNV sau khi SafetyHub đồng bộ Sheet1 (đã có sẵn trong dự án; bản lưu ở repo SafetyHub, `survey-appscript/`). |

## Cài bản "xem điểm + làm lại ghi đè điểm cũ" (khoảng 5 phút)

1. **Sao lưu trước:** trong sheet, chuột phải tab **Responses** › **Tạo bản sao** (để lỡ cần thì còn dữ liệu cũ).
2. Mở **Tiện ích mở rộng › Apps Script**.
3. Mở `Code.gs`, **xóa hết** rồi dán toàn bộ nội dung file `Code.gs` bản 4.9.0 (gửi riêng). Bấm **Lưu**.
   (Hoặc tự sửa 5 chỗ — xem "Các thay đổi trong Code.gs" bên dưới.)
4. Bấm **＋ › Tập lệnh**, đặt tên `SurveyRetake`, dán toàn bộ `apps-script/SurveyRetake.gs`. Bấm **Lưu**.
5. Chọn hàm **`mergeDuplicateResponsesByMsnv`** › **Chạy** (cấp quyền nếu được hỏi). Nhật ký sẽ báo đã gộp bao nhiêu MSNV.
   Từ giờ mỗi MSNV chỉ còn một dòng, cột mới **Số lần làm** ghi số lần đã làm.
6. **Triển khai › Quản lý triển khai** › bút chì › **Phiên bản: Phiên bản mới** › **Triển khai**. Giữ nguyên URL `/exec`.
7. Thử: chọn hàm `testGetMyResult` (đổi MSNV trong hàm thành một MSNV đã làm) › **Chạy**; rồi mở trang khảo sát, nhập MSNV đó — phải hiện điểm lần trước và nút **Làm lại khảo sát**.

Không cần tạo trigger mới. Trigger `processQueue` (mỗi phút) và trigger làm mới MSNV đang có vẫn giữ nguyên.

## Hoạt động thế nào

- Người làm bài nộp → vào **Queue** → `processQueue` ghi vào **Responses** như cũ, nhưng `appendResponse_` giờ gọi `upsertResponseRow_`:
  MSNV đã có dòng → **ghi đè ngay dòng đó** bằng lần nộp mới nhất (kể cả khi điểm thấp hơn), **Số lần làm** +1; MSNV mới → thêm dòng.
- Dòng **Stress Test = YES** không ghi đè và không bị ghi đè, nên dọn stress test vẫn chạy như cũ.
- Sau khi bấm **Kiểm tra**, trang gọi `getMyResult` → hiện điểm gần nhất, giờ nộp, số lần làm. Kết quả được nhớ 10 phút và xóa ngay khi có lần nộp mới.
- Dashboard quản trị giờ đếm mỗi người một lần (dòng mới nhất).

## Cập nhật 4.13.0: tab "Báo cáo VHDN" (giống file báo cáo PowerPoint)

Dashboard quản trị có thêm tab **Báo cáo VHDN**:
- Chọn **tháng / quý / cả năm**. Ô tổng hợp: Tổng nhân sự khảo sát, Nhân sự đạt "BIẾT" (đúng **tất cả** câu BIẾT), Tỷ lệ đạt, Điểm chuẩn "HIỂU", Điểm TB "HIỂU" tất cả bộ phận (trung bình điểm TB các bộ phận, như báo cáo).
- Biểu đồ theo bộ phận: BIẾT (tổng / đạt / tỷ lệ) và HIỂU (điểm TB so với đường điểm chuẩn); xu hướng theo tháng; bảng số liệu.
- Bộ phận gộp theo nhóm của báo cáo (KHO 7-8 → KHO 7-8-DEPOT, XNK ICD / XNK TCY → XNK, BẢO VỆ ICD → BẢO VỆ ICD 30, BẢO VỆ ST → BẢO VỆ KHO ST, Giao nhận quốc tế → GNQT…); tên khác giữ nguyên.
- **Điểm chuẩn HIỂU tự động** = điểm tối đa phần HIỂU của tháng (đổi theo số câu HIỂU); có thể **đặt tay** cho từng tháng. Quý = trung bình các tháng.
- **Nhập Excel / PowerPoint**: đọc file báo cáo như "Báo cáo khảo sát VHDN - 2026.pptx" (biểu đồ "… của các bộ phận tháng M/YYYY") hoặc file Excel xuất từ tab này; xem trước rồi lưu vào tab **Báo cáo VHDN (nhập)** của sheet. Tháng có bài khảo sát trực tuyến luôn dùng số liệu trực tuyến.
- **Nhân viên chưa làm khảo sát** (chọn một tháng): ai có trong danh sách nhân sự (Sheet1, đồng bộ từ SafetyHub) mà chưa có bài khảo sát trực tuyến tháng đó, theo từng bộ phận (STT, MSNV, Họ và tên), có ô tìm. **Xuất Excel danh sách chưa làm**: sheet "Tổng hợp" + mỗi bộ phận một sheet, tiêu đề "DANH SÁCH NHÂN VIÊN CHƯA LÀM VHDN THÁNG mm/yyyy", màu TBS (cam #E45620, xám than #4D4D4F), font Arial, sẵn in A4. Tháng hiện tại luôn có trong danh sách kỳ báo cáo, kể cả khi chưa ai làm.
- **Xuất PowerPoint / Excel** cả năm: trang bìa, tổng hợp BIẾT / HIỂU (3 kỳ mỗi trang, có quý), biểu đồ theo bộ phận từng tháng và từng quý. File xuất nhập lại được.

Cài: dán đè `Code.gs` 4.13.0 (gửi riêng) › thêm tệp `SurveyReport.gs` (**＋ › Tập lệnh**, dán `apps-script/SurveyReport.gs`) › **Lưu** › **Triển khai › Quản lý triển khai › Phiên bản mới**. Trang khảo sát (GitHub Pages) đã có sẵn tab mới.
Thử: chọn hàm `testReportData` › **Chạy** — Nhật ký liệt kê từng tháng có số liệu; `testMissingEmployees` — số người chưa làm tháng này theo bộ phận.

## Cập nhật 4.12.0: chịu được 75–100 người cùng lúc

Đo trên bản 4.11.0 (100 người cùng lúc): tra MSNV ~2 giây (ổn), nhưng tra điểm tháng này mất 5,7–10,5 giây và nộp bài phải xếp hàng qua một khóa (chờ tối đa 7 giây rồi báo lỗi).
- **Tra MSNV trả kèm điểm tháng này** (một lần gọi thay vì hai).
- **Bộ nhớ điểm của tháng** (CacheService, 32 ngăn): dựng một lần từ Responses, cập nhật ngay mỗi lần ghi kết quả → tra điểm không phải đọc sheet.
- **Nộp bài không cần khóa**: bài mới thêm vào Queue bằng `appendRow` (nhiều người cùng nộp không đè nhau); chống trùng bằng bộ nhớ mã lần nộp.
- **processQueue ghi cả đợt một lần** (tối đa 200 bài/phút), đọc cột khóa của Responses một lần cho cả đợt.
- **Trang khảo sát tự thử lại** khi máy chủ bận / quá giờ (tra MSNV và nộp bài; cùng mã lần nộp nên không tạo bài trùng).

Cài: dán đè `Code.gs` 4.12.0 và `SurveyRetake.gs` mới › **Lưu** › chạy **`warmMonthResultIndex`** (dựng sẵn bộ nhớ điểm) › **Triển khai › Phiên bản mới**.
Sau khi Claude thử tải bằng dữ liệu đánh dấu Stress Test (run ID `claude-load-test`), chạy **`cleanupClaudeLoadTest`** để xóa dữ liệu thử.

## Cập nhật 4.11.0: gộp MSNV theo TỪNG THÁNG

Khảo sát làm hằng tháng, nên chỉ gộp các lần làm của cùng một MSNV **trong cùng tháng**:
- Làm 3 lần trong tháng 8 → 1 dòng tháng 8 (lần mới nhất, Số lần làm = 3). Sang tháng 10 làm → **thêm dòng tháng 10**, dòng tháng 8 giữ nguyên. Làm lại trong tháng 10 → ghi đè dòng tháng 10.
- Tháng = tháng của cột **Timestamp** theo giờ Việt Nam. Từ 4.11.0 Timestamp lấy theo **giờ máy chủ lúc nhận bài** (không theo đồng hồ điện thoại; giờ điện thoại vẫn lưu trong payload là `clientTimestamp`).
- Trang khảo sát chỉ hiện điểm của **tháng hiện tại** ("Đã làm khảo sát tháng 10/2026"); tháng mới chưa làm thì hiện **Bắt đầu khảo sát**.
- `mergeDuplicateResponsesByMsnv` cũng gộp theo MSNV + tháng.

Cài: dán đè `Code.gs` 4.11.0 và `SurveyRetake.gs` mới › **Lưu** › nếu đã lỡ gộp các tháng với nhau bằng bản trước, khôi phục từ tab sao lưu "Responses (sao lưu trước gộp …)" rồi chạy lại `mergeDuplicateResponsesByMsnv` › **Triển khai › Quản lý triển khai › Phiên bản mới**.

## Cập nhật 4.10.0: processQueue nhanh hơn + lưu trữ Queue cũ

Trước đây `processQueue` (chạy mỗi phút) đọc 6 cột của MỌI dòng Queue, kể cả nội dung bài làm (payloadJson), và Queue không bao giờ được dọn — sheet càng dài, mỗi lần chạy càng chậm và càng tốn hạn mức thời gian chạy trigger mỗi ngày.

- Mỗi phút chỉ đọc **cột status**; chỉ đọc bài làm của đúng các dòng NEW được nhận (tối đa 80/lần); đánh dấu PROCESSING bằng một lần ghi.
- Ghi DONE / ERROR theo **queueId** (không theo số dòng), nên vẫn đúng khi dòng Queue được dời đi.
- **Mỗi giờ**, dòng **DONE đã xử lý quá 24 giờ** được chuyển sang tab mới **Queue Archive** (giữ nguyên cột). Dòng NEW / PROCESSING / ERROR giữ lại trong Queue.

Cài: dán đè `Code.gs` bản 4.10.0 › **Lưu** › chạy **`archiveQueueNow`** một lần (dọn ngay Queue cũ, nhật ký báo số dòng đã chuyển) › **Triển khai › Quản lý triển khai › Phiên bản mới**. Không cần trigger mới.

## Các thay đổi trong Code.gs (4.8.0 → 4.9.0)

1. `APP_VERSION = '4.9.0'`.
2. `doGet`, nhánh `lookupEmployee`: thông báo không tìm thấy → "Không tìm thấy MSNV này. Vui lòng kiểm tra lại số đã nhập hoặc liên hệ phòng HSSE."
3. `doGet`: thêm nhánh `getMyResult` → `getOutput_(e, getMyResult_(msnv))`.
4. `buildResponseObject_`: thêm cột `'Số lần làm': 1`.
5. `appendResponse_`: cuối hàm thay đoạn ghi dòng mới bằng `upsertResponseRow_(sheet, finalHeaders, rowObject)`; nhánh sheet trống xóa bộ nhớ điểm của MSNV đó.

## Lưu ý

- Ai biết MSNV của người khác cũng xem được điểm của người đó (giống tra họ tên hiện nay).
- Từ bản trang mới, **bắt buộc có MSNV đúng** mới làm được khảo sát.
