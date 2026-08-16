# Auto-Form-Filler

[![Python Version](https://img.shields.io/badge/python-3.12%2B-blue.svg)](https://www.python.org/)
[![Requests](https://img.shields.io/badge/requests-2.31.0%2B-green.svg)](https://requests.readthedocs.io/)
[![Pandas](https://img.shields.io/badge/pandas-2.0.0%2B-orange.svg)](https://pandas.pydata.org/)
[![OpenpyXL](https://img.shields.io/badge/openpyxl-3.1.0%2B-red.svg)](https://openpyxl.readthedocs.io/)

**Auto-Form-Filler** là một công cụ dòng lệnh (CLI Tool) viết bằng Python giúp tự động gửi dữ liệu hàng loạt từ file Excel/CSV vào Google Forms thông qua giao thức **HTTP POST Request trực tiếp (Bypass giao diện UI/Headless)**. Công cụ được thiết kế dành riêng cho các kỹ sư tự động hóa và quản trị viên dữ liệu cần xử lý luồng điền form quy mô lớn với độ ổn định tuyệt đối và khả năng kháng bot mạnh mẽ.

---

## 🚀 Các Tính Năng Nổi Bật

*   **Cơ chế kháng chặn (Anti-bot Evasion)**: 
    *   **Nhất quán Identity**: Cố định một `User-Agent` duy nhất cho mỗi phiên kết nối HTTP (`requests.Session`), tránh việc thay đổi liên tục gây nghi ngờ cho các bộ lọc của Google.
    *   **Giả lập Browser Hints**: Tích hợp các Client Hints hiện đại (`sec-ch-ua`, `sec-ch-ua-platform`, `sec-ch-ua-mobile`) đồng bộ hóa với vân tay trình duyệt giả lập.
    *   **Xoay vòng Session (Session Rotation)**: Tự động hủy phiên kết nối cũ và tạo phiên kết nối mới có IP/User-Agent mới sau mỗi $N$ requests hoặc khi gặp mã lỗi mạng (như 429 Too Many Requests).
*   **Jitter Delay phi tuyến tính**: Sử dụng thuật toán kết hợp giữa `random.uniform` và phân phối Log-Normal để giả lập hành vi ngẫu nhiên có độ dài nghỉ bất thường (thời gian người dùng suy nghĩ và điền thông tin), hạn chế tối đa việc bị block IP.
*   **Xác thực phản hồi kép (Robust Response Parser)**: Loại bỏ các nhận diện thành công giả (False Positive) bằng cách kiểm tra điều kiện thành công kép: HTTP status code 200, tìm thấy chuỗi xác nhận (hỗ trợ tùy biến `success_keywords` đa ngôn ngữ/custom message), và không chứa các cấu trúc lỗi validation/DOM của Google Forms.
*   **Lưu vết checkpoint & Khôi phục tự động (Fault-tolerant Resume)**: Ghi log tiến trình gửi của mỗi dòng dữ liệu thời gian thực vào file tạm `checkpoint_log.csv`. Nếu chương trình bị gián đoạn (mất điện, lỗi OS...), khi khởi động lại, công cụ tự động hỏi và cho phép chạy tiếp (Resume) từ dòng bị lỗi mà không cần điền lại từ đầu.
*   **Kháng lỗi khóa file Excel (File Lock Handler)**: Tự động bắt ngoại lệ `PermissionError` khi file dữ liệu Excel đầu vào hoặc file báo cáo kết quả đang bị mở và khóa bởi Microsoft Excel. Công cụ sẽ hướng dẫn người dùng đóng file hoặc xuất ra file backup copy để tránh crash kịch bản giữa chừng.

---

## 🛠️ Hướng Dẫn Cài Đặt & Chuẩn Bị Môi Trường

### Bước 1: Clone Repository
```bash
git clone https://github.com/h04ngm1nh/Auto-Form-Filler.git
cd Auto-Form-Filler
```

### Bước 2: Tạo và Kích hoạt Môi trường ảo (Khuyên dùng)
*   **Trên Windows (PowerShell)**:
    ```powershell
    python -m venv venv
    .\venv\Scripts\Activate.ps1
    ```
*   **Trên macOS/Linux**:
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```

### Bước 3: Cài đặt các thư viện phụ thuộc
```bash
pip install -r requirements.txt
```

---

## 📋 Hướng Dẫn Lấy Thông Tin Google Form (Entry IDs)

Để điền dữ liệu tự động qua HTTP POST, bạn cần trích xuất chính xác URL của endpoint phản hồi và các mã định danh của từng câu hỏi (`entry.xxxxxxxxx`).

### Phương pháp: Sử dụng "Get pre-filled link" (Lấy liên kết điền sẵn)

1. Mở trang quản trị chỉnh sửa biểu mẫu Google Forms của bạn.
2. Click vào biểu tượng **Ba dấu chấm** ở góc trên bên phải màn hình và chọn **Nhận liên kết được điền trước** (Get pre-filled link).
3. Nhập các giá trị mẫu dễ nhận biết vào từng ô câu hỏi (ví dụ: họ tên nhập `NGUYEN_VAN_A`, số điện thoại nhập `0912345678`).
4. Click nút **Nhận liên kết** (Get link) ở cuối trang, sau đó nhấn **Sao chép liên kết** (Copy link).
5. Dán liên kết đó vào trình chỉnh sửa văn bản. Liên kết sẽ có dạng tương tự:
   ```text
   https://docs.google.com/forms/d/e/1FAIpQLSfD_MOCK_ID/viewform?usp=pp_url&entry.2005620554=NGUYEN_VAN_A&entry.1166974658=0912345678
   ```
6. Bóc tách thông tin:
   *   URL phản hồi (`form_url`): Thay thế cụm `/viewform` bằng `/formResponse`.
       *   *Ví dụ*: `https://docs.google.com/forms/d/e/1FAIpQLSfD_MOCK_ID/formResponse`
   *   Các tham số mapping:
       *   `entry.2005620554` -> Trường tương ứng với `NGUYEN_VAN_A` (Họ và Tên)
       *   `entry.1166974658` -> Trường tương ứng với `0912345678` (Số điện thoại)

> [!TIP]
> Bạn có thể sử dụng tính năng **Tự động khởi tạo (Init Bootstrap)** có sẵn của công cụ để tự tạo file config và file Excel mẫu từ URL pre-filled chỉ với một câu lệnh (Xem chi tiết ở mục bên dưới).

---

## ⚙️ Cấu Hình (config.json) & Chuẩn Bị Excel

Công cụ hỗ trợ cả 2 định dạng cấu hình metadata: dạng chuỗi trực tiếp và dạng Object nâng cao.

### 1. File cấu hình mẫu (`config.json`)
```json
{
  "form_url": "https://docs.google.com/forms/d/e/1FAIpQLSe2USbGhItHRiU5R4JQ4oz32wJ1VMFNMjcW5Xa9NJVvDx6g0g/formResponse",
  "session_rotation_limit": 15,
  "success_keywords": [
    "Your response has been recorded",
    "Câu trả lời của bạn đã được ghi lại",
    "Thanks for submitting your contact info!"
  ],
  "field_mappings": {
    "entry.2005620554": {
      "column_name": "Họ và Tên",
      "required": true,
      "type": "string"
    },
    "entry.1166974658": {
      "column_name": "Số điện thoại",
      "required": true,
      "type": "string"
    },
    "entry.840782606": {
      "column_name": "Email",
      "required": true,
      "type": "string"
    },
    "entry.284884374": {
      "column_name": "Xác nhận",
      "required": false,
      "type": "string"
    }
  }
}
```

### 2. Cấu trúc bảng Excel đầu vào (`sample_data.xlsx`)
File Excel đầu vào phải chứa các Header cột trùng khớp 100% (từng ký tự, hoa/thường, khoảng trắng) với giá trị `"column_name"` được cấu hình trong `field_mappings`.

| Họ và Tên | Số điện thoại | Email | Xác nhận |
| :--- | :--- | :--- | :--- |
| Nguyễn Văn A | 0987654321 | nguyenvana@example.com | Đồng ý |
| Trần Thị B | 0901234567 | tranthib@example.com | |

*Lưu ý: Nếu một cột được cấu hình `"required": false`, bạn có thể bỏ trống ô dữ liệu tương ứng trong Excel.*

---

## 💻 Hướng Dẫn Thực Thi CLI

### Lệnh 1: Khởi tạo nhanh cấu hình & File dữ liệu mẫu (Init Bootstrap)
Nếu bạn đã lấy được đường link Pre-filled của Google Form, hãy chạy lệnh này để sinh tự động file Excel mẫu và file cấu hình JSON:
```bash
python main.py --init-from-url "SỐ_ĐƯỜNG_LINK_PREFILLED_CỦA_BẠN" --output-excel "sample_data.xlsx" --output-config "config.json"
```

### Lệnh 2: Gửi dữ liệu hàng loạt (Chế độ chạy chính)
```bash
python main.py -i sample_data.xlsx -c config.json -o result_log.xlsx
```
*   `-i` / `--input`: Đường dẫn file dữ liệu Excel/CSV (mặc định: `sample_data.xlsx`).
*   `-c` / `--config`: Đường dẫn file cấu hình JSON (mặc định: `config.json`).
*   `-o` / `--output`: Đường dẫn lưu file Excel báo cáo kết quả (mặc định: `result_log.xlsx`).

### Chế độ phục hồi lỗi (Resume State)
Nếu chương trình bị gián đoạn giữa chừng, khi chạy lại lệnh chính trên, màn hình sẽ hiển thị prompt:
```text
[!] CẢNH BÁO: Phát hiện file checkpoint `checkpoint_log.csv`.
    Tìm thấy 10 dòng dữ liệu đã được gửi thành công trước đó.
    Bạn có muốn tiếp tục (Resume) chạy tiếp từ vị trí bị gián đoạn không? [Y/n]: 
```
*   Nhấn **Enter** hoặc nhập `y`/`yes`: Hệ thống sẽ bỏ qua toàn bộ các bản ghi đã gửi thành công trước đó và tiếp tục điền các bản ghi còn lại.
*   Nhập `n`/`no`: Xóa checkpoint cũ và chạy lại từ đầu.

---

## 📁 Quản Lý Log & Báo Cáo Đầu Ra

*   **`checkpoint_log.csv`**: File log tạm thời dạng CSV, cập nhật trạng thái ngay lập tức sau mỗi lượt gửi (chứa: index dòng, định danh, trạng thái Success/Failed, timestamp và mô tả chi tiết lỗi).
*   **`result_log.xlsx`**: File Excel báo cáo đầu ra chính thức được tạo ra sau khi toàn bộ quy trình hoàn tất. Dữ liệu gốc sẽ được bổ sung thêm 2 cột: `Submission_Status` và `Submission_Error_Detail`.
*   **Thư mục `./logs/error_responses/`**: Khi phát hiện Google Forms từ chối dữ liệu (trả về trang lỗi validation, CAPTCHA, hoặc lỗi HTTP khác), toàn bộ mã nguồn HTML phản hồi từ Google sẽ được lưu lại dưới dạng file `.html` (ví dụ: `form_error_20260817_041016_line_1.html`) để phục vụ quá trình debug và kiểm toán dữ liệu.
