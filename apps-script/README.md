# Apps Script của khảo sát — phần "Xem điểm + Làm lại"

Trang khảo sát (`index.html`) giờ đi theo thứ tự: **nhập MSNV → Kiểm tra → xem họ tên (và điểm lần trước nếu đã làm) → Bắt đầu / Làm lại khảo sát**.
Phần trang đã xong. Để **xem điểm lần trước** và **điểm lần làm lại ghi đè điểm cũ trên sheet Responses**, cần thêm file `SurveyRetake.gs` vào dự án Apps Script của sheet "phản hồi khảo sát" (dự án đang chạy web app `.../exec` của khảo sát).

Khi chưa làm các bước dưới đây, trang vẫn chạy bình thường: chỉ không hiện điểm lần trước, và mỗi lần nộp vẫn thêm một dòng mới.

## Cách cài (khoảng 5 phút)

1. Mở sheet **phản hồi khảo sát** › **Tiện ích mở rộng** › **Apps Script**.
2. Bấm **＋** › **Tập lệnh**, đặt tên `SurveyRetake`, dán toàn bộ nội dung `SurveyRetake.gs`, bấm **Lưu**.
3. Trong `Code.gs`, tìm chỗ `doGet` xử lý `lookupEmployee` (tìm chữ `'lookupEmployee'`). Thêm một nhánh giống hệt bên cạnh, cho `getMyResult`, gọi:
   ```js
   surveyRetakeGetMyResult_(e.parameter.msnv)
   ```
   Hàm này trả về `{ ok: true, result: {...} }` (hoặc `result: null` nếu MSNV chưa làm) — cùng kiểu với `lookupEmployee`, nên trả về theo đúng cách nhánh `lookupEmployee` đang trả (iframe / JSONP).
   Ví dụ nếu đang có `case 'lookupEmployee': result = lookupEmployee_(e.parameter.msnv); break;` thì thêm
   `case 'getMyResult': result = surveyRetakeGetMyResult_(e.parameter.msnv); break;`
4. Chọn hàm `setupSurveyRetakeTrigger` › **Chạy** (cấp quyền nếu được hỏi). Xem **Nhật ký thực thi**: phải thấy dòng `Sheet "Responses", tiêu đề ở dòng ..., cột: {...}` với `msnv` khác `-1`. Nếu báo không tìm thấy sheet, đặt Script Property `SURVEY_RESPONSES_SHEET` = tên sheet kết quả.
5. **Triển khai** › **Quản lý triển khai** › bút chì › **Phiên bản mới** › **Triển khai** (giữ nguyên URL `/exec`).
6. Thử: chạy `testSurveyRetakeGetMyResult` (đổi MSNV trong hàm nếu cần), rồi mở trang khảo sát, nhập MSNV đã từng làm — phải hiện điểm lần trước và nút **Làm lại khảo sát**.

## Ghi đè điểm cũ hoạt động thế nào

- `setupSurveyRetakeTrigger` tạo trigger **5 phút** chạy `dedupeSurveyResponsesByMsnv`: MSNV nào có nhiều dòng trong Responses thì **giữ vị trí dòng cũ, ghi đè bằng nội dung lần nộp mới nhất, xóa các dòng thừa**. Dòng không có MSNV không bị đụng.
- Người làm bài luôn thấy đúng điểm mới nhất ngay (getMyResult lấy dòng dưới cùng), trigger chỉ dọn sheet.
- Muốn sheet ghi đè **ngay lúc nộp**: trong hàm chuyển Queue → Responses của `Code.gs`, thay `sheet.appendRow(values)` bằng `surveyRetakeUpsertResponseRow_(sheet, values)`.
- File không thêm / đổi cột: tìm cột theo tên tiêu đề (MSNV, Thời gian, Điểm Biết, Điểm Hiểu, Điểm tổng, Tỷ lệ…).
- Dùng `LockService.getScriptLock()`: nếu hàm ghi Responses trong `Code.gs` cũng dùng script lock thì hai bên không ghi chồng lên nhau.

## Lưu ý

- Ai biết MSNV của người khác cũng xem được điểm của người đó (giống tra cứu họ tên hiện nay).
- Từ nay **bắt buộc có MSNV đúng** mới làm được khảo sát. MSNV không có trong Sheet1 sẽ thấy: "Không tìm thấy MSNV này… liên hệ phòng HSSE".
