# Hệ thống chứng từ đa pháp nhân

Ứng dụng Streamlit tạo Đơn đặt hàng và Biên bản giao nhận từ file hóa đơn, dùng chung mẫu TLO cho 7 pháp nhân.

## Triển khai
1. Giữ repository `multi-company-documents` ở chế độ Public nếu cần, nhưng **chỉ đưa lên** `app.py`, `requirements.txt`, `README.md` và `mau/Mau_TLO.xlsx`.
2. Vào Streamlit Cloud → ứng dụng mới → Settings → Secrets, nhập `APP_PASSWORD = "mat-khau-rieng-manh"` (thay bằng mật khẩu thật). **Không lưu mật khẩu trong GitHub.**
3. Reboot ứng dụng nếu cần.
4. Truy cập app, nhập mật khẩu, tải lên `THÔNG TIN CÔNG TY.xlsx` từ máy tính, chọn công ty và tải danh sách hóa đơn.
5. Tạo chứng từ và tải ZIP.

## Lưu ý
- **Không commit** `THÔNG TIN CÔNG TY.xlsx`, hóa đơn thật, mật khẩu hoặc dữ liệu khách hàng lên GitHub Public.
- File thông tin công ty được tải lên khi dùng, không được lưu vào repository. Khi phiên Streamlit làm mới, cần tải lại file.
- Mẫu TLO có giới hạn 15 dòng hàng mỗi hóa đơn. Các hóa đơn vượt giới hạn được cảnh báo và không xuất.
- Mẫu chưa có vị trí riêng cho tài khoản ngân hàng, nên chọn tài khoản chỉ để tham khảo, không được in lên chứng từ.
- Đại diện bên mua lấy từ cột `Người mua` (nếu có), chức vụ bên mua và bên bán vẫn theo mẫu; cần kiểm tra trước khi ký.
- Excel tạo ra giữ công thức tổng tiền; ứng dụng không tự tính lại cache công thức Excel.
- File tải lên được xử lý trong phiên ứng dụng Streamlit Cloud. Cần có phê duyệt nội bộ trước khi xử lý dữ liệu thật.
