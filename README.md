# Phần mềm tính chỉ số khoan/phun Excel

Ứng dụng desktop Python đọc file Excel `.xlsx` nhiều sheet, tự động tính:

- `CĐ đỉnh`
- `LLV`
- `VX`
- `VL`

Sau khi xử lý, app sẽ:

- ghi kết quả lại vào từng sheet
- tô màu range dùng để tính
- tạo sheet tổng hợp `Tong_hop`
- xuất ra file Excel mới, không ghi đè file gốc

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

1. Chọn file Excel `.xlsx`.
2. Nhập độ dài mũi khoan, mặc định `3.00m`.
3. Chọn thư mục lưu.
4. Bấm `Phân tích & Xuất file`.
5. Kiểm tra bảng kết quả trên giao diện và file Excel output.

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

App tự dò header bảng dữ liệu. Nếu không dò được đầy đủ, app fallback theo format mẫu:

- `C = Độ sâu`
- `F = Lưu lượng đoạn`
- `H = Tốc độ Tời chính`
