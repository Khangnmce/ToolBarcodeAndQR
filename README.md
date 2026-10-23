# Công cụ tạo QR và Barcode hàng loạt

Ứng dụng web giúp tạo mã QR và mã vạch hàng loạt từ dữ liệu trong file Excel hoặc CSV, sau đó xuất toàn bộ hình ảnh thành một file ZIP để tải xuống.

**Bản dùng thử:** https://toolbarcodeandqr-pbv92vy4sdarxhhqgunsw9.streamlit.app/

---

## Tính năng

- Tải lên file Excel (`.xlsx`, `.xls`, `.xlsm`) hoặc CSV.
- Tự động phát hiện cột chứa dữ liệu mã cần tạo.
- Tạo mã QR hoặc mã vạch Code 128 hàng loạt.
- Xuất tất cả hình ảnh mã vào một file ZIP để tải xuống trực tiếp.
- Hỗ trợ xử lý đa luồng với tối đa 16 luồng.
- Xử lý dữ liệu trong bộ nhớ, không ghi tạm hình ảnh ra ổ đĩa.
- Giao diện hiện đại với chế độ nền tối.

## Chạy ứng dụng trên máy local

### 1. Cài đặt các thư viện cần thiết

```bash
pip install -r requirements.txt
```

### 2. Khởi động ứng dụng

```bash
streamlit run streamlit_app.py
```

Sau khi khởi động, Streamlit sẽ cung cấp địa chỉ truy cập ứng dụng trên trình duyệt.

## Triển khai lên Streamlit Community Cloud

1. Đẩy mã nguồn lên một repository trên GitHub.
2. Truy cập [Streamlit Community Cloud](https://share.streamlit.io).
3. Chọn repository cần triển khai.
4. Chọn nhánh `main`.
5. Chọn file khởi chạy `streamlit_app.py`.
6. Nhấn **Deploy**.

## Công nghệ sử dụng

- [Streamlit](https://streamlit.io): Xây dựng giao diện web.
- [pandas](https://pandas.pydata.org): Đọc và xử lý dữ liệu từ Excel hoặc CSV.
- [qrcode](https://github.com/lincolnloop/python-qrcode): Tạo mã QR.
- [python-barcode](https://github.com/WhyNotHugo/python-barcode): Tạo mã vạch Code 128.
- [Pillow](https://python-pillow.org): Xử lý hình ảnh.
