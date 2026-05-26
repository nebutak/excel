# PROMPT CHÍNH ĐỂ ĐƯA CHO AI CODE

Bạn hãy đóng vai trò là một senior Python desktop developer. Hãy code đầy đủ một phần mềm desktop xử lý file Excel nhiều sheet để tính chỉ số khoan/phun theo mô tả dưới đây.

## 1. Công nghệ bắt buộc/khuyến nghị

Hãy dùng:

- Python 3.10+
- PyQt5 hoặc PySide6 để làm giao diện desktop
- openpyxl để đọc/ghi Excel `.xlsx`
- pandas nếu cần xử lý bảng
- PyInstaller để đóng gói `.exe`

Không cần dùng database ở bản MVP.

## 2. Mục tiêu phần mềm

Phần mềm nhận vào 1 file Excel `.xlsx` có nhiều sheet. Mỗi sheet là dữ liệu của một cọc/đoạn khoan/phun.

Người dùng chọn file Excel, nhập độ dài mũi khoan, bấm xử lý. Phần mềm sẽ:

1. Duyệt tất cả sheet có dữ liệu.
2. Đọc thông tin đầu sheet.
3. Đọc bảng dữ liệu chính.
4. Tính:
   - CĐ đỉnh
   - LLV
   - VX
   - VL
5. Ghi kết quả vào từng sheet.
6. Tô màu vùng dữ liệu đã dùng để tính.
7. Tạo thêm sheet tổng hợp kết quả.
8. Xuất ra file Excel mới, không ghi đè file gốc.
9. Hiển thị bảng kết quả trên giao diện.

## 3. File mẫu

Trong gói bàn giao có file:

` samples/hovtoi-hang18.xlsx `

Đây là file mẫu đã được tính thủ công. Hãy dùng file này để test. Các kết quả thủ công hiện có nằm trong:

` docs/expected_results_from_sample.csv `

Các công thức thủ công đang có trong file mẫu nằm trong:

` docs/manual_formula_ranges_from_sample.csv `

## 4. Format Excel đầu vào

Mỗi sheet thường có:

### 4.1 Phần thông tin đầu sheet

Các nhãn thường nằm ở những dòng đầu:

- Số cọc
- Độ sâu mục tiêu
- Đường kính cọc
- Khoan sâu
- Lưu lượng tích lũy
- Thời gian tích lũy
- Thời gian bắt đầu
- Thời gian kết thúc

Giá trị thường nằm ở ô bên phải nhãn.

Ví dụ:

- `Độ sâu mục tiêu`: `13.66m`
- `Khoan sâu`: `13.68m`
- `Đường kính cọc`: `1800mm`

Cần viết hàm parse số để lấy được `13.66` từ `13.66m`, `1800` từ `1800mm`, `2282` từ `2282L`.

### 4.2 Bảng dữ liệu chính

Tìm dòng header bằng cách dò các cột có tên:

- TT
- Thời gian
- Độ sâu
- Lưu lượng tức thì
- Lưu lượng tích lũy
- Lưu lượng đoạn
- Vòng quay
- Tốc độ Tời chính
- Đầu động lực/dòng điện
- nghiêngX
- nghiêngY
- Số cọc
- Thời gian

Trong file mẫu, header thường ở dòng 10. Nhưng code không được hard-code dòng 10; hãy dò header tự động.

Các cột quan trọng:

- `Độ sâu`
- `Lưu lượng đoạn`
- `Tốc độ Tời chính`

Nếu dò header thất bại, fallback theo file mẫu:

- Cột C = Độ sâu
- Cột F = Lưu lượng đoạn
- Cột H = Tốc độ Tời chính

## 5. Công thức nghiệp vụ

### 5.1 CĐ đỉnh

```text
CĐ đỉnh = Độ sâu mục tiêu - Độ dài mũi khoan
```

Độ dài mũi khoan do người dùng nhập trên giao diện. Mặc định là `3.00m`.

Ví dụ:

```text
Độ sâu mục tiêu = 13.66
Độ dài mũi khoan = 3.00
CĐ đỉnh = 10.66
```

### 5.2 LLV

LLV là tổng lưu lượng vữa trong vùng phun chính.

Theo nghiệp vụ:

- `Lưu lượng đoạn = 0` là không phun.
- `Lưu lượng đoạn = 1` thường được xem là rỉ, mặc định không tính.
- Giá trị lớn hơn `1` được xem là phun chính.

Quy tắc đề xuất:

1. Xác định vùng quanh CĐ đỉnh đến qua điểm khoan sâu nhất.
2. Tìm dòng đầu tiên trong vùng này có `Lưu lượng đoạn > ngưỡng`.
3. Cộng các giá trị `Lưu lượng đoạn > ngưỡng` liên tục cho đến khi sau giai đoạn khoan sâu/kéo lên, lưu lượng giảm về `<= ngưỡng`.
4. Ngưỡng mặc định là `1`.

Cần cho phép cấu hình:

- `flow_min_for_llv = 1`
- `sum_only_values_greater_than_threshold = true`

Ghi chú: file thủ công có thể có vài sheet lệch 1 đơn vị do công thức tay cộng luôn một ô có giá trị `1`. Nhưng nghiệp vụ ưu tiên là không tính 0 và 1.

### 5.3 VX

VX là vận tốc xuống.

```text
VX = trung bình cột "Tốc độ Tời chính" trong giai đoạn khoan xuống
```

Vùng tính VX:

- Bắt đầu từ dòng gần CĐ đỉnh khi độ sâu đang tăng.
- Nếu không có đúng giá trị CĐ đỉnh trong cột Độ sâu, lấy dòng có độ sâu thấp hơn gần nhất trước khi vượt qua CĐ đỉnh.

Ví dụ:

```text
CĐ đỉnh = 10.61
Cột Độ sâu có 10.07 và 11.07
=> bắt đầu từ dòng 10.07
```

- Kết thúc tại điểm chuyển pha ở vùng khoan sâu nhất.

Điểm chuyển pha:

- Độ sâu đạt lớn nhất hoặc gần lớn nhất.
- Tốc độ Tời chính thường giảm về 0.
- Lưu lượng đoạn giảm đột ngột.
- Sau đó độ sâu bắt đầu giảm.

Quy tắc tự động đề xuất:

1. Tìm `max_depth = max(cột Độ sâu)`.
2. Tìm các dòng quanh `max_depth`.
3. Nếu `max_depth` xuất hiện nhiều dòng liên tiếp:
   - VX kết thúc tại dòng max_depth đầu tiên.
   - VL bắt đầu tại dòng max_depth tiếp theo.
4. Nếu `max_depth` chỉ xuất hiện một dòng:
   - Nếu tốc độ tại dòng max_depth bằng 0 hoặc gần 0, xem dòng đó là điểm chuyển pha và đưa vào VL.
   - VX kết thúc ở dòng trước đó.
   - Nếu tốc độ tại dòng max_depth vẫn > 0, có thể đưa dòng max_depth vào VX.
5. Có cấu hình `speed_zero_threshold`, mặc định `0.05`.

### 5.4 VL

VL là vận tốc lên.

```text
VL = trung bình cột "Tốc độ Tời chính" trong giai đoạn kéo lên
```

Vùng tính VL:

- Bắt đầu sau điểm kết thúc VX.
- Kết thúc khi độ sâu giảm về gần CĐ đỉnh.
- Nếu có đúng CĐ đỉnh, lấy đến dòng đó.
- Nếu không có đúng CĐ đỉnh, lấy đến dòng đầu tiên có `Độ sâu <= CĐ đỉnh` khi đang kéo lên.

## 6. Yêu cầu ghi kết quả

Với mỗi sheet:

- Ghi kết quả ở khu vực trống bên phải bảng, ví dụ cột O/P hoặc N/O tùy sheet.
- Không được phá dữ liệu gốc.
- Nếu vùng kết quả cũ đã có thì cập nhật.
- Nên ghi thêm các range debug:
  - `VX range`
  - `VL range`
  - `LLV range`
  - `Ghi chú`
- Tô màu vùng dùng để tính:
  - VX: màu vàng nhạt
  - VL: màu xanh hoặc xám nhạt
  - LLV: màu cam nhạt hoặc viền nổi bật

Tạo thêm sheet `Tong_hop` gồm các cột:

- STT
- Sheet
- Số cọc
- Độ sâu mục tiêu
- Khoan sâu trong header
- Khoan sâu thực tế = max(Độ sâu)
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

## 7. Giao diện phần mềm

Giao diện cần có:

- Nút chọn file Excel
- Ô hiển thị đường dẫn file
- Ô nhập độ dài mũi khoan mặc định `3.00`
- Nút chọn thư mục lưu
- Nút `Phân tích & Xuất file`
- Thanh progress
- Bảng kết quả sau khi xử lý
- Khu vực log lỗi/cảnh báo
- Nút mở file kết quả

Bảng kết quả trên app có các cột:

- Sheet
- Số cọc
- CĐ đỉnh
- LLV
- VX
- VL
- Trạng thái
- Ghi chú

## 8. Yêu cầu kỹ thuật

Hãy chia project thành các file:

```text
main.py
ui_main.py
excel_processor.py
calculator.py
models.py
config.py
utils.py
requirements.txt
README.md
```

Trong đó:

- `main.py`: chạy app.
- `ui_main.py`: giao diện PyQt5/PySide6.
- `excel_processor.py`: đọc/ghi Excel, duyệt sheet, tô màu, xuất file.
- `calculator.py`: thuật toán tính CĐ đỉnh, LLV, VX, VL.
- `models.py`: dataclass cho kết quả từng sheet.
- `config.py`: cấu hình mặc định.
- `utils.py`: hàm parse số, dò header, log.
- `requirements.txt`: thư viện cần cài.
- `README.md`: hướng dẫn chạy và đóng gói.

## 9. Xử lý lỗi

Phần mềm không được crash nếu một sheet lỗi.

Nếu một sheet thiếu dữ liệu:

- Bỏ qua sheet đó.
- Ghi trạng thái `SKIP` hoặc `ERROR`.
- Ghi lý do vào cột Ghi chú.

Các lỗi cần xử lý:

- Sheet trống.
- Không tìm thấy `Độ sâu mục tiêu`.
- Không tìm thấy bảng dữ liệu.
- Không tìm thấy cột `Độ sâu`.
- Không tìm thấy cột `Lưu lượng đoạn`.
- Không tìm thấy cột `Tốc độ Tời chính`.
- Không đủ dữ liệu để tính VX/VL.
- File Excel đang mở hoặc không thể ghi.

## 10. Test và nghiệm thu

Dùng file `samples/hovtoi-hang18.xlsx` để test.

Yêu cầu:

1. Chạy app chọn file mẫu.
2. Nhập độ dài mũi khoan `3.00`.
3. Xuất file mới.
4. Kiểm tra sheet tổng hợp.
5. Kết quả phải gần với file thủ công ở `docs/expected_results_from_sample.csv`.
6. Với sheet có công thức thủ công sai, app vẫn ưu tiên tính theo thuật toán đúng.

Lưu ý đặc biệt:

- Sheet `A5-18-PK34` trong file mẫu có thể đang có CĐ đỉnh thủ công không đúng. Nếu độ sâu mục tiêu là `13.76m` và mũi khoan là `3m`, CĐ đỉnh đúng phải là `10.76`, không phải `13.76`.
- Sheet `A14-18-PK34` có công thức VX thủ công đang lấy cột J, trong khi đúng nghiệp vụ phải lấy cột H `Tốc độ Tời chính`.

## 11. Hãy trả về

Hãy viết đầy đủ toàn bộ source code project theo cấu trúc trên, có thể chạy được ngay. Sau đó hướng dẫn:

- Cách cài thư viện.
- Cách chạy app.
- Cách đóng gói `.exe`.
- Cách test bằng file mẫu.
