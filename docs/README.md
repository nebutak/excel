# Gói bàn giao yêu cầu cho AI code phần mềm tính chỉ số khoan/phun Excel

Mục đích của gói này: cung cấp đủ thông tin để một AI/coder có thể viết phần mềm desktop xử lý file Excel nhiều sheet, tự động tính các chỉ số:

- CĐ đỉnh
- LLV
- VX
- VL

## Cách dùng gói này

Khi đưa cho AI/coder, hãy gửi toàn bộ thư mục này hoặc file `.zip`.

Quan trọng nhất là các file:

1. `docs/01_PROMPT_CHINH_CHO_AI.md`  
   Prompt chính, có thể copy nguyên văn đưa cho AI để yêu cầu code.

2. `docs/02_MO_TA_CHUC_NANG.md`  
   Mô tả nghiệp vụ và chức năng phần mềm.

3. `docs/03_THUAT_TOAN_TINH_TOAN.md`  
   Quy tắc đọc Excel và thuật toán tính CĐ đỉnh, LLV, VX, VL.

4. `docs/04_CAU_TRUC_EXCEL_DAU_VAO.md`  
   Mô tả format file Excel đầu vào.

5. `docs/05_GIAO_DIEN_PHAN_MEM.md`  
   Yêu cầu giao diện phần mềm.

6. `docs/06_TIEU_CHI_NGHIEM_THU.md`  
   Checklist để kiểm tra phần mềm đã đúng hay chưa.

7. `docs/expected_results_from_sample.csv`  
   Bảng kết quả đang có trong file mẫu, dùng để đối chiếu khi test.

8. `docs/manual_formula_ranges_from_sample.csv`  
   Các công thức thủ công hiện có trong file mẫu, giúp AI hiểu vùng tính tay.

9. `config/config.example.json`  
   Cấu hình mẫu cho thuật toán.

10. `samples/hovtoi-hang18.xlsx`  
    File Excel mẫu đã tính thủ công.

11. `screenshots/`  
    Hình minh họa vùng tính và layout Excel.

## Công nghệ khuyến nghị

- Python 3.10+
- PyQt5 hoặc PySide6
- openpyxl
- pandas, dùng tùy chọn
- PyInstaller để đóng gói `.exe`

## Ghi chú quan trọng

File Excel mẫu có một số sheet đang có công thức thủ công. Phần mềm cần tự động nhận diện vùng tính, nhưng vẫn nên có chế độ debug/đối chiếu để người dùng kiểm tra vùng VX, VL, LLV mà phần mềm đã chọn.

Có một vài điểm nghiệp vụ có thể phải cho phép cấu hình, ví dụ:

- Độ dài mũi khoan mặc định là 3m nhưng cần cho phép nhập tay.
- LLV mặc định không tính giá trị 0 và 1 vì được xem là rỉ.
- VX/VL cần xác định theo CĐ đỉnh và điểm chuyển pha ở độ sâu lớn nhất.
