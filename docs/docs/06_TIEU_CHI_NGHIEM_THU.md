# Tiêu chí nghiệm thu

## 1. Chạy được

- Cài thư viện bằng `pip install -r requirements.txt`.
- Chạy được bằng `python main.py`.
- Giao diện mở không lỗi.

## 2. Chọn file được

- Chọn được file `.xlsx`.
- Hiển thị đúng đường dẫn file.
- Nếu chọn file sai định dạng thì báo lỗi.

## 3. Xử lý file mẫu được

Dùng file:

```text
samples/hovtoi-hang18.xlsx
```

Nhập:

```text
Độ dài mũi khoan = 3.00
```

Kết quả:

- Không crash.
- Bỏ qua sheet trống.
- Xử lý các sheet dữ liệu.
- Tạo file output mới.

## 4. Output Excel đúng

File output phải có:

- Dữ liệu gốc vẫn còn.
- Kết quả CĐ đỉnh, LLV, VX, VL ở từng sheet.
- Vùng VX/VL/LLV được tô màu nếu bật tùy chọn.
- Có sheet `Tong_hop`.

## 5. Sheet tổng hợp đúng

Sheet `Tong_hop` có các cột:

- STT
- Sheet
- Số cọc
- Độ sâu mục tiêu
- Khoan sâu header
- Khoan sâu thực tế
- Độ dài mũi khoan
- CĐ đỉnh
- LLV
- VX
- VL
- VX range
- VL range
- LLV range
- Trạng thái
- Ghi chú

## 6. Sai số chấp nhận

Vì một số công thức thủ công trong file mẫu chưa nhất quán, chia làm hai mức kiểm tra:

### 6.1 CĐ đỉnh

Phải đúng theo công thức:

```text
Độ sâu mục tiêu - Độ dài mũi khoan
```

Sai số chấp nhận: `0.01`.

### 6.2 VX/VL

So với cách tính tay, sai số mong muốn nhỏ hơn `0.05` nếu thuật toán chọn cùng vùng.

Nếu khác vùng, phần mềm phải ghi rõ range đã chọn để kiểm tra.

### 6.3 LLV

Do có tranh luận về việc có tính giá trị `1` hay không, cần có cấu hình:

- Mặc định: không tính 0 và 1.
- Nếu bật `Tính cả giá trị 1 trong range`, có thể khớp hơn với file thủ công.

## 7. Lỗi không được xảy ra

- Không làm mất sheet gốc.
- Không ghi đè file gốc khi chưa hỏi.
- Không crash khi gặp sheet lỗi.
- Không lấy VX từ cột `nghiêngX`.
- Không hard-code cứng một sheet hoặc một dòng cố định.

## 8. Đóng gói

Phải có hướng dẫn đóng gói:

```bash
pyinstaller --onefile --windowed main.py
```

Sau khi đóng gói, file `.exe` chạy được trên Windows.
