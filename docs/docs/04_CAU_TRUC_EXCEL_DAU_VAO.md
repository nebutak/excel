# Cấu trúc Excel đầu vào

## 1. Workbook

- File `.xlsx`.
- Có nhiều sheet.
- Mỗi sheet thường tương ứng một cọc.
- Có thể có sheet trống, ví dụ `Sheet1`.

## 2. Tên sheet mẫu

File mẫu `hovtoi-hang18.xlsx` có các sheet như:

- A15-18-PK34
- A13-18-PK34
- A12-18-PK34
- A11-18-PK34
- ...
- A14-18-PK34
- Sheet1

Sheet `Sheet1` là sheet trống, cần bỏ qua.

## 3. Phần thông tin đầu sheet

Ví dụ trong file mẫu:

| Nhãn | Giá trị |
|---|---|
| Số cọc | 1518 |
| Đường kính cọc | 1800mm |
| Độ sâu mục tiêu | 13.66m |
| Khoan sâu | 13.68m |
| Lưu lượng tích lũy | 2282L |
| Thời gian tích lũy | 0:21:56 |
| Thời gian bắt đầu | 2026-05-24 22:22 |
| Thời gian kết thúc | 2026-05-24 22:44 |

Vị trí có thể thay đổi nhẹ, nên cần dò nhãn thay vì hard-code ô.

## 4. Header bảng chính

Các cột thường là:

| Tên cột | Ý nghĩa | Dùng tính |
|---|---|---|
| TT | Số thứ tự | Không |
| Thời gian | Thời gian ngắn | Không |
| Độ sâu | Độ sâu khoan | Có |
| Lưu lượng tức thì | Lưu lượng tức thời | Không |
| Lưu lượng tích lũy | Tổng tích lũy | Không bắt buộc |
| Lưu lượng đoạn | Lưu lượng từng đoạn | Có, dùng tính LLV |
| Vòng quay | Vòng quay | Không |
| Tốc độ Tời chính | Tốc độ tời | Có, dùng tính VX/VL |
| Đầu động lực dòng điện | Dòng điện | Không |
| nghiêngX | Nghiêng X | Không |
| nghiêngY | Nghiêng Y | Không |
| Số cọc | Mã cọc | Không |
| Thời gian | Thời gian đầy đủ | Không |

## 5. Cột quan trọng theo file mẫu

Nếu dùng fallback:

| Cột Excel | Dữ liệu |
|---|---|
| C | Độ sâu |
| F | Lưu lượng đoạn |
| H | Tốc độ Tời chính |

## 6. Khu vực kết quả hiện tại

Trong file mẫu, kết quả thủ công thường được đặt ở bên phải bảng, ví dụ:

| Label | Value |
|---|---|
| CĐ ĐỈNH | 10.66 |
| LLV | 2273 |
| VX | 0.9167 |
| VL | 0.725 |

Vị trí có thể là:

- O/P
- N/O

Không nên hard-code chỉ một vị trí. Hãy tìm khu vực trống bên phải bảng hoặc tìm label cũ để cập nhật.

## 7. Các vấn đề trong file mẫu

Có một số sheet có công thức thủ công chưa nhất quán:

- `A5-18-PK34`: CĐ đỉnh hiện có thể đang hiển thị bằng độ sâu mục tiêu. Nếu độ dài mũi khoan là 3m thì CĐ đỉnh phải bằng `Độ sâu mục tiêu - 3`.
- `A14-18-PK34`: VX thủ công đang dùng `AVERAGE(J21:J25)`, nhưng cột J là `nghiêngX`. Theo nghiệp vụ, VX phải tính từ cột H `Tốc độ Tời chính`.

Vì vậy test cần phân biệt:

- Khớp công thức thủ công cũ.
- Đúng nghiệp vụ mới.

Ưu tiên đúng nghiệp vụ mới.
