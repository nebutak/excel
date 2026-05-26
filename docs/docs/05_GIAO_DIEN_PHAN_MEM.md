# Yêu cầu giao diện phần mềm

## 1. Mục tiêu giao diện

Giao diện phải đơn giản để người không biết kỹ thuật vẫn dùng được.

## 2. Layout đề xuất

```text
+------------------------------------------------------------+
| PHẦN MỀM TÍNH CHỈ SỐ KHOAN/PHUN EXCEL                     |
+------------------------------------------------------------+

File Excel:
[ C:\...\hovtoi-hang18.xlsx                         ] [Chọn file]

Độ dài mũi khoan:
[ 3.00 ] m

Thư mục lưu:
[ C:\...\output                                    ] [Chọn thư mục]

Tùy chọn:
[x] Tạo sheet tổng hợp
[x] Tô màu vùng tính
[x] Không ghi đè file gốc
[ ] Tính cả giá trị lưu lượng đoạn = 1

[ Phân tích & Xuất file ]   [ Mở file kết quả ]

Progress: [======================          ] 70%

Bảng kết quả:
| Sheet | Số cọc | CĐ đỉnh | LLV | VX | VL | Trạng thái | Ghi chú |

Log:
- Đang xử lý A15-18-PK34...
- OK
- Cảnh báo A5-18-PK34: CĐ đỉnh cũ khác giá trị tính lại.
```

## 3. Thành phần bắt buộc

- QLineEdit hiển thị file Excel.
- QPushButton chọn file.
- QDoubleSpinBox nhập độ dài mũi khoan.
- QLineEdit hiển thị thư mục lưu.
- QPushButton chọn thư mục lưu.
- QPushButton xử lý.
- QProgressBar.
- QTableWidget hoặc QTableView hiển thị kết quả.
- QTextEdit hiển thị log.
- QMessageBox khi thành công hoặc lỗi.

## 4. Validation

Trước khi xử lý:

- File không được rỗng.
- File phải có đuôi `.xlsx`.
- Độ dài mũi khoan > 0.
- Thư mục lưu tồn tại.

## 5. Trạng thái sheet

Các trạng thái:

| Trạng thái | Ý nghĩa |
|---|---|
| OK | Tính thành công |
| WARNING | Tính được nhưng có cảnh báo |
| SKIP | Bỏ qua sheet trống/không liên quan |
| ERROR | Sheet lỗi, không tính được |

## 6. Không làm treo giao diện

Nếu xử lý nhiều sheet, nên dùng QThread hoặc worker thread để giao diện không bị đứng.

Bản MVP có thể xử lý trực tiếp nếu file nhỏ, nhưng nên thiết kế sẵn worker để mở rộng.
