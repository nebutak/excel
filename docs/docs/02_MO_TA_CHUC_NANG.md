# Mô tả chức năng phần mềm

## 1. Tên tạm

Phần mềm tính chỉ số khoan/phun từ Excel nhiều sheet.

## 2. Bài toán

Người dùng có một file Excel `.xlsx` gồm nhiều sheet. Mỗi sheet chứa dữ liệu quá trình khoan/phun của một cọc.

Hiện tại người dùng đang tính thủ công trong Excel bằng các công thức như:

```excel
=SUM(F22:F27)
=AVERAGE(H21:H26)
=AVERAGE(H27:H30)
```

Việc chọn vùng tính bằng tay mất thời gian và dễ sai. Phần mềm cần tự động nhận diện vùng tính và xuất kết quả.

## 3. Người dùng cuối

Người dùng cuối không cần biết code. Họ chỉ cần:

1. Mở phần mềm.
2. Chọn file Excel.
3. Nhập độ dài mũi khoan.
4. Bấm xử lý.
5. Nhận file Excel kết quả.

## 4. Input

- File Excel `.xlsx`.
- Có nhiều sheet.
- Có thể có sheet trống.
- Mỗi sheet có cùng kiểu dữ liệu nhưng vị trí kết quả có thể khác nhau.
- Độ dài mũi khoan nhập tay, mặc định 3m.

## 5. Output

Một file Excel mới, ví dụ:

```text
hovtoi-hang18_da_tinh.xlsx
```

Trong file output:

- Mỗi sheet được ghi kết quả CĐ đỉnh, LLV, VX, VL.
- Vùng tính VX/VL/LLV được tô màu.
- Có sheet tổng hợp `Tong_hop`.
- Không làm mất dữ liệu gốc.

## 6. Chỉ số cần tính

### CĐ đỉnh

```text
CĐ đỉnh = Độ sâu mục tiêu - Độ dài mũi khoan
```

### LLV

Tổng lưu lượng đoạn trong vùng phun vữa chính.

### VX

Trung bình tốc độ Tời chính khi khoan xuống, từ gần CĐ đỉnh đến vùng khoan sâu.

### VL

Trung bình tốc độ Tời chính khi kéo lên, từ vùng khoan sâu về gần CĐ đỉnh.

## 7. Các chức năng chính

### 7.1 Chọn file Excel

Người dùng chọn file `.xlsx`.

### 7.2 Nhập độ dài mũi khoan

- Mặc định: 3.00m.
- Cho phép nhập số thập phân.
- Có thể phát triển thêm bảng cấu hình riêng cho từng sheet/cọc.

### 7.3 Phân tích file

Phần mềm duyệt từng sheet:

- Bỏ qua sheet trống.
- Đọc metadata.
- Dò header bảng.
- Tách dữ liệu.
- Tính chỉ số.
- Ghi kết quả.

### 7.4 Xuất file

- Không ghi đè file gốc.
- Người dùng chọn thư mục lưu.
- Tên file tự động thêm hậu tố `_da_tinh`.

### 7.5 Bảng kết quả

Hiển thị kết quả tất cả sheet trên giao diện.

### 7.6 Log/cảnh báo

Hiển thị lỗi hoặc cảnh báo như:

- Sheet thiếu cột.
- Không tìm thấy CĐ đỉnh.
- Không đủ dữ liệu tính VL.
- CĐ đỉnh trong file cũ khác với CĐ đỉnh tính lại.

## 8. Chức năng nên có thêm

- Nút mở file kết quả.
- Nút xuất log `.txt`.
- Tùy chọn bật/tắt tô màu vùng tính.
- Tùy chọn tạo sheet tổng hợp.
- Tùy chọn ngưỡng LLV.
- Tùy chọn kiểm tra với kết quả thủ công trong file mẫu.
