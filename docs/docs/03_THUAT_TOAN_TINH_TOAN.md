# Thuật toán tính toán

## 1. Hàm parse số

Cần parse được các dạng:

```text
13.66m -> 13.66
1800mm -> 1800
2282L -> 2282
0:21:56 -> giữ dạng text hoặc time, không dùng cho tính chính
```

Quy tắc:

- Nếu giá trị đã là int/float thì dùng trực tiếp.
- Nếu là string thì dùng regex lấy số đầu tiên.
- Hỗ trợ dấu phẩy thập phân nếu có, ví dụ `13,66m`.

## 2. Dò sheet hợp lệ

Một sheet được xem là hợp lệ nếu:

- Có `Độ sâu mục tiêu`, hoặc
- Có header bảng chứa `Độ sâu`, `Lưu lượng đoạn`, `Tốc độ Tời chính`.

Sheet trống thì bỏ qua.

## 3. Dò metadata

Trong 10 dòng đầu, tìm các nhãn:

- `Số cọc`
- `Độ sâu mục tiêu`
- `Đường kính cọc`
- `Khoan sâu`
- `Lưu lượng tích lũy`
- `Thời gian bắt đầu`
- `Thời gian kết thúc`

Giá trị thường nằm ở ô bên phải nhãn.

Nếu không tìm thấy `Khoan sâu` thì dùng:

```text
Khoan sâu thực tế = max(cột Độ sâu)
```

Khuyến nghị luôn lưu cả hai:

- `khoan_sau_header`
- `khoan_sau_thuc_te`

## 4. Dò header bảng

Không hard-code dòng 10.

Cách dò:

1. Quét 30 dòng đầu.
2. Dòng nào có ít nhất 3 nhãn quan trọng thì chọn là header.
3. Nhãn quan trọng:
   - `TT`
   - `Độ sâu`
   - `Lưu lượng đoạn`
   - `Tốc độ Tời chính`

Nếu không dò được, fallback:

- Header row = 10
- Depth col = C
- Flow segment col = F
- Speed col = H

## 5. Chuẩn hóa dữ liệu bảng

Sau khi xác định header:

- Data bắt đầu từ dòng sau header.
- Dừng khi cột `Độ sâu` trống liên tiếp hoặc không còn số.
- Chỉ giữ các dòng có `Độ sâu` parse được số.
- Parse `Lưu lượng đoạn` và `Tốc độ Tời chính` thành số.
- Nếu ô trống ở flow/speed thì xem là 0 hoặc bỏ qua tùy ngữ cảnh; khuyến nghị xem là 0 nhưng ghi warning.

## 6. Tính CĐ đỉnh

```python
cd_dinh = target_depth - drill_length
cd_dinh = round(cd_dinh, 2)
```

## 7. Tìm vùng khoan xuống

Dữ liệu độ sâu thường tăng từ 0 đến độ sâu lớn nhất, sau đó giảm.

### 7.1 Tìm `vx_start_idx`

Tìm dòng `i` trước vùng sâu nhất sao cho:

```text
depth[i] <= cd_dinh <= depth[i + 1]
```

Nếu có đúng dòng `depth[i] == cd_dinh`, chọn dòng đó.

Nếu không có, chọn dòng thấp hơn gần nhất trước khi vượt qua CĐ đỉnh.

Ví dụ:

```text
CĐ đỉnh = 10.61
Độ sâu có 10.07 và 11.07
=> vx_start_idx là dòng 10.07
```

## 8. Tìm điểm chuyển pha

### 8.1 Tìm max depth

```python
max_depth = max(depths)
```

Dùng tolerance:

```python
abs(depth - max_depth) <= depth_tolerance
```

Mặc định:

```text
depth_tolerance = 0.02
```

### 8.2 Tìm các dòng max depth liên tiếp

Tìm nhóm dòng gần `max_depth` quanh peak.

Ví dụ:

```text
13.68
13.68
12.66
```

Ở đây có 2 dòng max depth liên tiếp.

### 8.3 Quy tắc chia VX/VL

Trường hợp 1: max depth xuất hiện nhiều dòng liên tiếp

```text
VX kết thúc tại dòng max depth đầu tiên.
VL bắt đầu tại dòng max depth tiếp theo.
```

Trường hợp 2: max depth chỉ xuất hiện 1 dòng

Nếu tại dòng max depth:

```text
speed <= speed_zero_threshold
```

thì dòng này là điểm chuyển pha:

```text
VX kết thúc ở dòng trước max depth.
VL bắt đầu tại dòng max depth.
```

Nếu speed vẫn lớn hơn ngưỡng:

```text
VX kết thúc tại dòng max depth.
VL bắt đầu ở dòng sau đó.
```

Mặc định:

```text
speed_zero_threshold = 0.05
```

## 9. Tính VX

```python
vx_values = speed[vx_start_idx : vx_end_idx + 1]
vx = average(vx_values)
```

Có thể có option:

```text
exclude_zero_speed_from_average = false
```

Mặc định không loại số 0 vì file thủ công có lúc tính cả dòng tốc độ 0 tại điểm chuyển pha.

## 10. Tìm vùng kéo lên và tính VL

```python
vl_start_idx = vx_end_idx + 1 hoặc theo quy tắc chuyển pha
```

Tìm `vl_end_idx`:

- Đi từ `vl_start_idx` về sau.
- Chọn dòng đầu tiên khi `depth <= cd_dinh`.
- Nếu không có, chọn dòng có độ sâu gần CĐ đỉnh nhất trong pha kéo lên.

```python
vl_values = speed[vl_start_idx : vl_end_idx + 1]
vl = average(vl_values)
```

## 11. Tính LLV

LLV tính theo vùng phun chính.

### 11.1 Vùng tìm kiếm

Vùng tìm kiếm bắt đầu từ quanh `vx_start_idx` đến sau `vl_start_idx`, thường kết thúc trước khi flow giảm về 0/1 sau khi kéo lên.

### 11.2 Quy tắc đề xuất

```python
threshold = flow_min_for_llv  # mặc định 1
```

1. Từ `vx_start_idx`, tìm dòng đầu tiên có `flow > threshold`.
2. Từ dòng đó, tiếp tục đi xuống.
3. Cộng các giá trị `flow > threshold`.
4. Dừng khi đã qua peak/chuyển pha và gặp dòng có `flow <= threshold`.
5. LLV = tổng các giá trị đã cộng.

### 11.3 Cấu hình

```json
{
  "flow_min_for_llv": 1,
  "sum_only_values_greater_than_threshold": true,
  "include_flow_equal_threshold": false
}
```

Nếu khách muốn khớp tuyệt đối file tính tay, có thể bật chế độ:

```json
{
  "sum_raw_selected_llv_range": true
}
```

Chế độ này cộng toàn bộ giá trị trong range đã chọn, kể cả số 1.

## 12. Debug range

Phần mềm cần lưu lại:

```text
VX range: ví dụ H21:H26
VL range: ví dụ H27:H30
LLV range: ví dụ F22:F27
```

Các range này rất quan trọng để người dùng kiểm tra phần mềm chọn đúng vùng hay chưa.

## 13. Pseudocode

```python
def calculate_sheet(rows, target_depth, drill_length, config):
    cd_dinh = round(target_depth - drill_length, 2)

    depths = [row.depth for row in rows]
    flows = [row.flow_segment for row in rows]
    speeds = [row.speed for row in rows]

    max_depth = max(depths)
    peak_indexes = find_peak_indexes(depths, max_depth, config.depth_tolerance)

    vx_start = find_vx_start(depths, cd_dinh, peak_indexes[0])

    vx_end, vl_start = split_down_up_phase(
        depths=depths,
        flows=flows,
        speeds=speeds,
        peak_indexes=peak_indexes,
        speed_zero_threshold=config.speed_zero_threshold
    )

    vl_end = find_vl_end(depths, cd_dinh, vl_start)

    vx = average(speeds[vx_start:vx_end + 1])
    vl = average(speeds[vl_start:vl_end + 1])

    llv_start, llv_end, llv = calculate_llv_range(
        depths=depths,
        flows=flows,
        vx_start=vx_start,
        peak_index=peak_indexes[0],
        vl_start=vl_start,
        config=config
    )

    return Result(
        cd_dinh=cd_dinh,
        llv=llv,
        vx=vx,
        vl=vl,
        ranges={
            "vx": speed_col_range(vx_start, vx_end),
            "vl": speed_col_range(vl_start, vl_end),
            "llv": flow_col_range(llv_start, llv_end)
        }
    )
```
