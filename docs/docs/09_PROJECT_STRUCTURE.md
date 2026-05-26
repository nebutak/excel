# Cấu trúc project đề xuất

```text
excel_khoan_app/
│
├── main.py
├── ui_main.py
├── excel_processor.py
├── calculator.py
├── models.py
├── config.py
├── utils.py
├── requirements.txt
├── README.md
│
├── assets/
│   └── app_icon.ico
│
├── tests/
│   ├── test_calculator.py
│   └── test_parse_excel.py
│
└── dist/
```

## main.py

- Khởi tạo QApplication.
- Mở MainWindow.

## ui_main.py

- Giao diện chính.
- Chọn file.
- Nhập độ dài mũi khoan.
- Chạy xử lý.
- Hiển thị bảng kết quả.
- Hiển thị log.

## excel_processor.py

- Mở workbook.
- Duyệt sheet.
- Dò metadata.
- Dò bảng.
- Gọi calculator.
- Ghi kết quả.
- Tô màu.
- Tạo sheet tổng hợp.
- Save as output.

## calculator.py

- Tính CĐ đỉnh.
- Tìm VX range.
- Tìm VL range.
- Tính LLV range.
- Return dataclass kết quả.

## models.py

Dataclass đề xuất:

```python
@dataclass
class SheetResult:
    sheet_name: str
    pile_no: str | None
    target_depth: float | None
    header_drill_depth: float | None
    actual_drill_depth: float | None
    drill_length: float
    cd_dinh: float | None
    llv: float | None
    vx: float | None
    vl: float | None
    vx_range: str | None
    vl_range: str | None
    llv_range: str | None
    status: str
    note: str
```

## config.py

Chứa cấu hình mặc định.

## utils.py

- parse_number
- normalize_text
- find_header_row
- safe_average
- excel_col_letter
- build_range
