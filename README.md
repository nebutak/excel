# Phần mềm tính chỉ số khoan/phun Excel

Ứng dụng desktop Python đọc file Excel `.xlsx` nhiều sheet, tự động tính:

- `CĐ đỉnh`
- `LLV`
- `VX`
- `VL`
- `Cao độ đỉnh thực tế`
- `Cao độ đáy thực tế`
- `Chiều dài cọc thực tế`

Sau khi xử lý, app sẽ:

- ghi kết quả lại vào từng sheet
- tô màu range dùng để tính
- tạo sheet tổng hợp `Tong_hop`
- xuất ra file Excel mới, không ghi đè file gốc
- hỗ trợ xử lý trực tiếp 1 hoặc nhiều file CSV thô và gom ra 1 file Excel kết quả
- hỗ trợ chọn sheet trong 1 file Excel và xuất từng sheet thành PDF riêng để in

## Cấu trúc project

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

## Cài thư viện

```bash
python -m venv venv
```

Windows:

```bash
venv\Scripts\activate
```

Sau đó cài dependency:

```bash
pip install -r requirements.txt
```

## Chạy app

```bash
python main.py
```

## Cách dùng

1. Chọn file Excel `.xlsx`, hoặc chọn nhiều file CSV, hoặc chọn thư mục chứa CSV.
2. Nhập độ dài mũi khoan, mặc định `3.00m`.
3. Nhập cao độ đỉnh thiết kế, cao độ đáy thiết kế và ngưỡng LLV nếu cần.
4. Chọn thư mục lưu.
5. Bấm `Phân tích & Xuất file`.
6. Kiểm tra bảng kết quả trên giao diện và file Excel output.

## Xuất PDF theo sheet

1. Ở khu vực `Tool xuất PDF theo sheet`, chọn file Excel `.xlsx`.
2. Bấm `Tải sheet` nếu danh sách sheet chưa hiện.
3. Tick các sheet cần xuất.
4. Chọn thư mục PDF.
5. Bấm `Xuất PDF`.

Tên file PDF được đặt theo tên sheet. Nút `Xem trước sheet` dùng Microsoft Excel để mở print preview cho sheet đầu tiên đang được tick. Nút `In sheet đã chọn` gửi các sheet đang tick tới máy in mặc định.

Lưu ý:

- Trên Windows, chức năng xuất PDF/xem trước/in trực tiếp ưu tiên Microsoft Excel qua `pywin32`.
- Nếu không có Excel COM, app thử xuất PDF bằng LibreOffice headless nếu máy có lệnh `soffice` hoặc `libreoffice`.

## Xử lý CSV thô

Nếu dữ liệu gốc là nhiều file `.csv`, không cần dùng Excel Get Data để gom thủ công.

- Bấm `Chọn file` để chọn một hoặc nhiều file `.csv`.
- Hoặc bấm `Chọn thư mục CSV` để lấy tất cả file `.csv` trong thư mục đó.
- App sẽ tạo một workbook `.xlsx`, mỗi CSV thành một sheet riêng.
- Sheet `Tong_hop` tổng hợp kết quả tất cả CSV.

Lưu ý:

- Không trộn file `.xlsx` và `.csv` trong cùng một lần xử lý.
- Nếu chọn Excel thì chỉ chọn một file `.xlsx`.
- CSV cần có metadata và header giống mẫu: `Số cọc`, `Độ sâu mục tiêu`, `Độ sâu`, `Lưu lượng đoạn`, `Tốc độ Tời chính`.

## Test bằng file mẫu

File mẫu thật trong repo hiện nằm ở:

```text
docs/samples/hovtoi-hang18.xlsx
```

File đối chiếu kết quả:

```text
docs/docs/expected_results_from_sample.csv
docs/docs/manual_formula_ranges_from_sample.csv
```

Quy trình test:

1. Chạy app bằng `python main.py`
2. Chọn file `docs/samples/hovtoi-hang18.xlsx`
3. Nhập `3.00`
4. Chọn thư mục output
5. Bấm xử lý
6. Mở file kết quả và kiểm tra:
   - từng sheet có `CĐ đỉnh`, `LLV`, `VX`, `VL`
   - có `VX range`, `VL range`, `LLV range`
   - có sheet `Tong_hop`

Lưu ý:

- `A5-18-PK34`: app sẽ ưu tiên `CĐ đỉnh = Độ sâu mục tiêu - 3`, nên giá trị đúng là `10.76`.
- `A14-18-PK34`: app tính `VX` từ cột `H - Tốc độ Tời chính`, không dùng cột `J`.
- Một vài sheet có thể lệch `LLV` so với công thức tay cũ vì app mặc định **không cộng giá trị 0 và 1**.

## Đóng gói `.exe`

```bash
pyinstaller --onefile --windowed main.py
```

File build xong sẽ nằm trong:

```text
dist/
```

## Ghi chú thuật toán

- `CĐ đỉnh = Độ sâu mục tiêu - Độ dài mũi khoan`
- `VX`: trung bình `Tốc độ Tời chính` trong pha đi xuống
- `VL`: trung bình `Tốc độ Tời chính` trong pha đi lên
- `LLV`: tổng `Lưu lượng đoạn` trong vùng phun chính
- Mặc định LLV tính các giá trị `Lưu lượng đoạn >= 10`, vì giá trị nhỏ hơn 10 được xem là rỉ theo tài liệu V3.
- `Cao độ đỉnh thực tế`: nếu bắt đầu phun sớm hơn CĐ đỉnh thì cộng phần chênh vào cao độ đỉnh thiết kế, ngược lại giữ cao độ đỉnh thiết kế.
- `Cao độ đáy thực tế = Cao độ đáy thiết kế - (Khoan sâu thực tế - Độ sâu mục tiêu)`
- `Chiều dài cọc thực tế = Cao độ đỉnh thực tế - Cao độ đáy thực tế`

App tự dò header bảng dữ liệu. Nếu không dò được đầy đủ, app fallback theo format mẫu:

- `C = Độ sâu`
- `F = Lưu lượng đoạn`
- `H = Tốc độ Tời chính`
