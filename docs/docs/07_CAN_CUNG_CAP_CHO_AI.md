# Cần cung cấp gì cho AI/coder để làm chính xác?

Khi thuê/cung cấp yêu cầu cho AI/coder, hãy gửi các thứ sau:

## 1. File Excel mẫu

Bắt buộc gửi:

```text
samples/hovtoi-hang18.xlsx
```

Đây là file mẫu thật, có nhiều sheet và có kết quả tính tay để đối chiếu.

## 2. Ảnh minh họa

Gửi thư mục:

```text
screenshots/
```

Ảnh giúp AI/coder hiểu:

- Vùng dữ liệu Excel.
- Vùng bôi vàng/xám đang dùng để tính.
- Vị trí kết quả CĐ đỉnh, LLV, VX, VL.

## 3. Mô tả nghiệp vụ

Gửi các file:

```text
docs/02_MO_TA_CHUC_NANG.md
docs/03_THUAT_TOAN_TINH_TOAN.md
docs/04_CAU_TRUC_EXCEL_DAU_VAO.md
```

Đây là phần quan trọng nhất để AI/coder hiểu cách tính.

## 4. Kết quả kỳ vọng

Gửi:

```text
docs/expected_results_from_sample.csv
docs/manual_formula_ranges_from_sample.csv
```

Hai file này dùng để test:

- Kết quả hiện đang có trong file mẫu.
- Công thức thủ công hiện tại đang lấy range nào.

## 5. Các điểm cần nói rõ với AI/coder

Bạn cần nói rõ:

1. Độ dài mũi khoan mặc định là bao nhiêu?  
   Hiện đề xuất: `3.00m`.

2. Độ dài mũi khoan dùng chung cho tất cả sheet hay mỗi sheet khác nhau?  
   Bản MVP: dùng chung cho tất cả sheet.  
   Bản nâng cấp: cho phép nhập riêng từng sheet.

3. LLV có tính giá trị `1` không?  
   Theo nghiệp vụ bạn mô tả: không tính `0` và `1`, vì đó là rỉ.  
   Tuy nhiên một vài công thức tay có thể cộng cả `1`. Vì vậy nên làm option bật/tắt.

4. VX có tính dòng tốc độ bằng `0` tại điểm chuyển pha không?  
   Nên giữ configurable. Mặc định: không loại số 0 nếu nó nằm trong range được nhận diện.

5. Có cần tô màu vùng tính không?  
   Nên có để kiểm tra.

6. Có cần sheet tổng hợp không?  
   Nên có.

7. Có cần xuất `.exe` không?  
   Nếu giao cho người khác dùng trên Windows thì có.

## 6. Những câu nên hỏi lại khách trước khi chốt code

Nếu muốn phần mềm chính xác hơn, hãy hỏi khách:

1. Mỗi file có phải luôn cùng format như file mẫu không?
2. Độ dài mũi khoan có luôn là 3m không?
3. LLV có loại giá trị `1` hoàn toàn không?
4. Nếu vùng phun bị ngắt quãng, có cộng tiếp không?
5. Khi có nhiều dòng cùng độ sâu lớn nhất, dòng nào thuộc VX và dòng nào thuộc VL?
6. Kết quả làm tròn mấy chữ số?
7. Muốn ghi kết quả vào cột cố định hay vị trí trống bên phải bảng?
8. Có muốn lưu lịch sử xử lý không?
9. Có cần phân quyền/người dùng không?
10. Có cần hỗ trợ `.xls` hoặc chỉ `.xlsx`?

## 7. Cách đưa prompt cho AI

Cách đơn giản nhất:

1. Upload toàn bộ file `.zip` này cho AI.
2. Nói: “Hãy đọc toàn bộ gói yêu cầu này”.
3. Copy nội dung `docs/01_PROMPT_CHINH_CHO_AI.md`.
4. Yêu cầu AI code full project.
