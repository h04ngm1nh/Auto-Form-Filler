# Hướng Dẫn Reverse Engineering Google Forms

Tài liệu này hướng dẫn chi tiết cách đảo ngược giao thức (reverse engineer) của Google Forms để trích xuất chính xác mã biểu mẫu (`FORM_ID`) và các mã định danh câu hỏi (`entry.xxxx`).

---

## 1. Trích xuất FORM_ID (Form Identifier)

Khi bạn truy cập một Google Form để điền thông tin, URL trên trình duyệt sẽ có dạng:
```text
https://docs.google.com/forms/d/e/1FAIpQLSfD_YOUR_MOCK_FORM_ID/viewform
```

* **FORM_ID** chính là chuỗi ký tự nằm giữa `/e/` và `/viewform`. 
* Ví dụ trong URL trên: `1FAIpQLSfD_YOUR_MOCK_FORM_ID`.
* Endpoint để gửi payload HTTP POST của chúng ta sẽ là:
  ```text
  https://docs.google.com/forms/d/e/[FORM_ID]/formResponse
  ```

---

## 2. Tìm mã định danh trường dữ liệu (`entry.xxxx`)

Có 2 phương pháp chính để tìm các mã định danh `entry.xxxx` tương ứng với mỗi câu hỏi.

### Cách 1: Sử dụng tính năng "Get pre-filled link" (Khuyên dùng - nếu bạn sở hữu Form)

Nếu bạn là người tạo hoặc có quyền chỉnh sửa biểu mẫu:
1. Mở trang quản trị chỉnh sửa biểu mẫu Google Forms.
2. Click vào biểu tượng **Ba dấu chấm** (Thêm) ở góc trên bên phải màn hình.
3. Chọn **Nhận liên kết được điền trước** (Get pre-filled link).
4. Điền các giá trị giả lập đặc trưng vào từng ô câu hỏi (ví dụ: câu hỏi họ tên điền `HO_TEN_TEST`, câu hỏi checkbox chọn vài mục cụ thể).
5. Click nút **Nhận liên kết** (Get link) ở cuối trang, sau đó nhấn **Sao chép liên kết** (Copy link).
6. Dán liên kết đó vào trình chỉnh sửa văn bản. Liên kết sẽ có dạng tương tự:
   ```text
   https://docs.google.com/forms/d/e/1FAIpQLSfD_YOUR_MOCK_FORM_ID/viewform?entry.1000001=HO_TEN_TEST&entry.1000002=test@email.com&entry.1000005=Sở+thích+A&entry.1000005=Sở+thích+B
   ```
7. Từ chuỗi query parameters trên, bạn có thể dễ dàng map:
   * `entry.1000001` -> Họ và Tên (`HO_TEN_TEST`)
   * `entry.1000002` -> Email (`test@email.com`)
   * `entry.1000005` -> Hộp kiểm (Checkbox) chứa nhiều giá trị `Sở thích A` và `Sở thích B`.

---

### Cách 2: Inspect Source HTML (Dành cho Form công khai bất kỳ)

Nếu bạn chỉ có link điền form công khai và không có quyền chỉnh sửa:

#### Lựa chọn A: Inspect từng Element bằng F12
1. Truy cập vào biểu mẫu trên trình duyệt (Chrome/Firefox/Edge).
2. Nhấn `F12` hoặc Click chuột phải chọn **Kiểm tra** (Inspect) để mở DevTools.
3. Sử dụng công cụ trỏ chuột (Select Element) nhấp vào trường nhập liệu của câu hỏi.
4. Tìm thẻ `<input>` hoặc `<textarea>` có thuộc tính `name` bắt đầu bằng `entry.xxxx`.
   * Ví dụ: `<input type="text" name="entry.1000001" ...>`
   * Số sau dấu chấm (`1000001`) chính là mã bạn cần tìm.

#### Lựa chọn B: Phân tích biến Javascript `FB_PUBLIC_APP_DATA` (Nhanh nhất cho Forms lớn)
1. Nhấp chuột phải vào trang web điền form và chọn **Xem nguồn trang** (View Page Source) hoặc nhấn `Ctrl + U`.
2. Nhấn `Ctrl + F` tìm kiếm cụm từ khóa: `FB_PUBLIC_APP_DATA`.
3. Bạn sẽ thấy một khối Javascript tương tự:
   ```javascript
   var FB_PUBLIC_APP_DATA = [..., [0, 1000001, "Họ và Tên", ...], [0, 1000002, "Email", ...]];
   ```
4. Ở đây, các mảng con chứa tiêu đề câu hỏi trực quan đi kèm trực tiếp với mã số `entry_id` nằm ngay bên cạnh.

---

## 3. Cấu trúc Payload Đặc Biệt Cần Lưu Ý

### Câu hỏi Hộp kiểm (Checkbox - Chọn nhiều đáp án)
Nếu người dùng tích chọn nhiều ô (ví dụ: "Đá bóng" và "Đọc sách"), HTTP POST yêu cầu gửi 2 cặp tham số độc lập có **cùng key**:
* `entry.1000005=Đá bóng`
* `entry.1000005=Đọc sách`

Mã nguồn Python của chúng ta xử lý việc này bằng cách chuyển đổi giá trị dạng mảng/list thành nhiều tuple có cùng tên tham số khi gửi bằng `requests.post(..., data=payload)`.

### Câu hỏi Ngày tháng (Date)
Đối với câu hỏi ngày tháng (ví dụ: `Ngày sinh` với mã định danh gốc là `entry.1000007`), Google Forms **không** nhận chuỗi ngày dạng `2023-10-27` trực tiếp. Nó chia nhỏ thành 3 tham số POST:
* `entry.1000007_year=2023`
* `entry.1000007_month=10`
* `entry.1000007_day=27`

Công cụ của chúng ta sẽ tự động phát hiện kiểu trường dữ liệu `"date"` trong cấu hình và thực hiện chia nhỏ thành 3 tham số trên trước khi gửi.
