# Apps Script của khảo sát VHDN

Dự án Apps Script gắn với sheet **phản hồi khảo sát** (web app `.../exec` mà `index.html` gọi). Dự án này KHÔNG tự đẩy lên như SafetyHub:
sửa xong phải dán vào trình soạn thảo Apps Script. Repo này công khai nên chỉ lưu `SurveyRetake.gs`; `Code.gs` bản 4.9.0 được gửi riêng cho HSSE.

| File | Việc |
|---|---|
| `Code.gs` (gửi riêng, không để trong repo công khai) | Toàn bộ backend (tra MSNV, câu hỏi, hàng đợi nộp bài, dashboard…). 4.9.0: `getMyResult` + ghi đè điểm khi làm lại. 4.10.0: `processQueue` nhanh hơn + tự lưu trữ Queue cũ. 4.11.0: giờ nộp theo máy chủ (để gộp theo tháng). |
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
