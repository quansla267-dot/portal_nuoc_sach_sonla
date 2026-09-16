====================================================================
HƯỚNG DẪN TRIỂN KHAI TỪ A → Z
====================================================================

Bước 1: Tạo Google Cloud Project
====================================================================
1. Truy cập: https://console.cloud.google.com/
2. Đăng nhập tài khoản Google
3. Nhấn tên dự án → NEW PROJECT → Đặt tên → CREATE
4. Chờ tạo xong → chọn dự án vừa tạo

Bước 2: Bật Google Sheets API
====================================================================
1. Tìm "Google Sheets API" trong ô tìm kiếm → Enable
2. Tương tự: bật "Google Drive API"

Bước 3: Tạo Service Account & Tải file JSON
====================================================================
1. Vào: APIs & Services → Credentials (Thông tin xác thực)
2. Nhấn + CREATE CREDENTIALS → Service account
3. Điền tên → CREATE AND CONTINUE → Bỏ qua các bước → DONE
4. Nhấn vào tên Service Account vừa tạo
5. Vào tab Keys → Add Key → Create new key → JSON → CREATE
6. File JSON tự động tải về → Lưu lại (quan trọng!)

Bước 4: Tạo Google Sheets Master & Chia sẻ quyền
====================================================================
1. Tạo file mới tại: https://sheets.google.com
2. Copy ID trên thanh địa chỉ (giữa /d/ và /edit):
   https://docs.google.com/spreadsheets/d/**ID_CUA_BAN**/edit
3. Nhấn Share → Dán email của Service Account (trong file JSON: "client_email")
   → Quyền: Editor → Send

Bước 5: Điền thông tin vào .streamlit/secrets.toml
====================================================================
1. Mở file JSON vừa tải về → copy từng giá trị vào secrets.toml
2. google_sheets_master_id = "ID lấy ở Bước 4.2"
3. [gcp_service_account] → copy toàn nội dung file JSON vào đây
   (Lưu ý: thay dấu xuống dòng "\n" trong private_key thành "\n" nếu cần)

Bước 6: Cập nhật danh sách tài khoản
====================================================================
Mở: config/Config_Users.xlsx
→ Thêm/xóa các xã, phường, đặt mật khẩu
→ Lưu

Bước 7: Đẩy lên GitHub
====================================================================
1. Tạo kho mới tại: https://github.com/new
2. Đặt tên → Public hoặc Private → Create
3. Làm theo hướng dẫn đẩy code lên (git add → commit → push)

Bước 8: Triển khai trên Streamlit Cloud
====================================================================
1. Truy cập: https://share.streamlit.io/ → Đăng nhập
2. Nhấn New app → Kết nối GitHub
3. Chọn kho → Branch: main → File: app.py
4. Phần Advanced settings → Secrets:
   Copy toàn nội dung .streamlit/secrets.toml dán vào đây
5. Nhấn Deploy → Chờ vài phút → Hoàn thành!

Bước 9: Sử dụng
====================================================================
- Link công khai xuất hiện sau khi Deploy xong
- Đăng nhập với tài khoản ADMIN trước để kiểm tra
- Các xã/phường tải file mẫu → điền → tải lên → nộp
- ADMIN xem tiến độ toàn tỉnh, tải dữ liệu tổng hợp

Lưu ý quan trọng
====================================================================
- File JSON khóa bảo mật KHÔNG chia sẻ cho ai khác
- Nếu dùng GitHub công khai → KHÔNG đẩy file secrets.toml lên
  → Dán trực tiếp vào phần Secrets trên Streamlit Cloud
