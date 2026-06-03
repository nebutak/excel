from __future__ import annotations

from math import isclose

from config import RuntimeConfig
from models import SheetCalculation, TableRow
from utils import safe_average


class CalculationError(ValueError):
    pass


def calculate_sheet(
    rows: list[TableRow],
    target_depth: float,
    drill_length: float,
    config: RuntimeConfig,
) -> SheetCalculation:
    if not rows:
        raise CalculationError("Không có dữ liệu dòng để tính.")
    if target_depth is None:
        raise CalculationError("Không tìm thấy Độ sâu mục tiêu.")

    cd_dinh = round(target_depth - drill_length, 2)
    actual_drill_depth = max(row.depth for row in rows)
    vx_start_idx = find_vx_start(rows, cd_dinh, config)
    vx_end_idx, vl_start_idx = split_down_up_phase(rows, actual_drill_depth, config)
    vl_end_idx = find_vl_end(rows, cd_dinh, vl_start_idx, config)

    if vx_end_idx < vx_start_idx:
        raise CalculationError("Không xác định được vùng VX hợp lệ.")
    if vl_end_idx < vl_start_idx:
        raise CalculationError("Không xác định được vùng VL hợp lệ.")

    vx_values = [rows[idx].main_hoist_speed for idx in range(vx_start_idx, vx_end_idx + 1)]
    vl_values = [rows[idx].main_hoist_speed for idx in range(vl_start_idx, vl_end_idx + 1)]
    vx = safe_average(vx_values, exclude_zero=config.exclude_zero_speed_from_average)
    vl = safe_average(vl_values, exclude_zero=config.exclude_zero_speed_from_average)

    if vx is None:
        raise CalculationError("Không đủ dữ liệu để tính VX.")
    if vl is None:
        raise CalculationError("Không đủ dữ liệu để tính VL.")

    llv_start_idx, llv_end_idx, llv = calculate_llv_range(rows, vx_start_idx, vl_start_idx, config)
    injection_start_depth = rows[llv_start_idx].depth if llv_end_idx >= llv_start_idx else None
    actual_top_elevation = calculate_actual_top_elevation(
        design_top_elevation=config.default_design_top_elevation,
        cd_dinh=cd_dinh,
        injection_start_depth=injection_start_depth,
    )
    actual_bottom_elevation = calculate_actual_bottom_elevation(
        design_bottom_elevation=config.default_design_bottom_elevation,
        actual_drill_depth=actual_drill_depth,
        target_depth=target_depth,
    )
    actual_column_length = actual_top_elevation - actual_bottom_elevation

    notes: list[str] = []
    if llv_end_idx < llv_start_idx:
        notes.append("Không nhận diện được range LLV, giá trị LLV đặt về 0.")

    return SheetCalculation(
        cd_dinh=cd_dinh,
        llv=llv,
        vx=round(vx, 6),
        vl=round(vl, 6),
        actual_top_elevation=round(actual_top_elevation, 6),
        actual_bottom_elevation=round(actual_bottom_elevation, 6),
        actual_column_length=round(actual_column_length, 6),
        vx_start_idx=vx_start_idx,
        vx_end_idx=vx_end_idx,
        vl_start_idx=vl_start_idx,
        vl_end_idx=vl_end_idx,
        llv_start_idx=llv_start_idx,
        llv_end_idx=llv_end_idx,
        actual_drill_depth=actual_drill_depth,
        injection_start_depth=injection_start_depth,
        notes=notes,
    )


def find_vx_start(rows: list[TableRow], cd_dinh: float, config: RuntimeConfig) -> int:
    for idx, row in enumerate(rows):
        if _is_main_instant_flow(row.instant_flow, config):
            return idx

    vx_start = 0
    for idx in range(len(rows)):
        current_depth = rows[idx].depth
        if isclose(current_depth, cd_dinh, abs_tol=0.02):
            return idx
        if idx < len(rows) - 1 and current_depth <= cd_dinh <= rows[idx + 1].depth:
            return idx
        if current_depth <= cd_dinh:
            vx_start = idx
    return vx_start


def split_down_up_phase(
    rows: list[TableRow],
    max_depth: float,
    config: RuntimeConfig,
) -> tuple[int, int]:
    exact_peak_indexes = [
        idx for idx, row in enumerate(rows) if isclose(row.depth, max_depth, abs_tol=1e-9)
    ]
    if not exact_peak_indexes:
        peak_index = max(range(len(rows)), key=lambda idx: rows[idx].depth)
        return split_peak_cluster(rows, [peak_index], config)

    peak_cluster = [exact_peak_indexes[0]]
    for idx in exact_peak_indexes[1:]:
        if idx == peak_cluster[-1] + 1:
            peak_cluster.append(idx)
        else:
            break

    return split_peak_cluster(rows, peak_cluster, config)


def split_peak_cluster(
    rows: list[TableRow],
    peak_cluster: list[int],
    config: RuntimeConfig,
) -> tuple[int, int]:
    first_peak = peak_cluster[0]
    last_peak = peak_cluster[-1]

    vx_end = last_peak
    for idx in peak_cluster:
        if rows[idx].main_hoist_speed <= config.speed_zero_threshold:
            vx_end = max(idx - 1, 0)
            break

    vl_start = min(last_peak + 1, len(rows) - 1)
    while (
        vl_start < len(rows) - 1
        and rows[vl_start].main_hoist_speed <= config.speed_zero_threshold
    ):
        vl_start += 1

    if vl_start <= vx_end and last_peak < len(rows) - 1:
        vl_start = last_peak + 1
    if vx_end < first_peak and first_peak > 0:
        vx_end = first_peak - 1
    return vx_end, min(vl_start, len(rows) - 1)


def find_vl_end(
    rows: list[TableRow],
    cd_dinh: float,
    vl_start_idx: int,
    config: RuntimeConfig,
) -> int:
    best_idx = vl_start_idx
    for idx in range(vl_start_idx, len(rows)):
        depth = rows[idx].depth
        if isclose(depth, cd_dinh, abs_tol=0.01) or depth <= cd_dinh:
            return idx
        if idx > vl_start_idx:
            previous_speed = rows[idx - 1].main_hoist_speed
            current_speed = rows[idx].main_hoist_speed
            if (
                previous_speed <= config.vl_speed_increase_threshold
                and current_speed > config.vl_speed_increase_threshold
            ):
                return idx - 1
        if abs(depth - cd_dinh) < abs(rows[best_idx].depth - cd_dinh):
            best_idx = idx
    return best_idx


def _is_main_instant_flow(value: float, config: RuntimeConfig) -> bool:
    if config.include_flow_equal_threshold:
        return value >= config.flow_min_for_llv
    return value > config.flow_min_for_llv


def calculate_llv_range(
    rows: list[TableRow],
    vx_start_idx: int,
    vl_start_idx: int,
    config: RuntimeConfig,
) -> tuple[int, int, float]:
    threshold = config.flow_min_for_llv

    def is_counted(flow_value: float) -> bool:
        if config.include_flow_equal_threshold:
            return flow_value >= threshold
        return flow_value > threshold

    llv_start_idx = -1
    llv_end_idx = -1
    llv_total = 0.0
    started = False

    for idx in range(vx_start_idx, len(rows)):
        flow = rows[idx].flow_segment
        if not started:
            if is_counted(flow):
                started = True
                llv_start_idx = idx
                llv_end_idx = idx
                llv_total += flow if config.sum_only_values_greater_than_threshold else max(flow, 0.0)
            continue

        if is_counted(flow):
            llv_end_idx = idx
            llv_total += flow if config.sum_only_values_greater_than_threshold else max(flow, 0.0)
            continue

        if idx >= vl_start_idx:
            break

    if llv_start_idx == -1:
        return 0, -1, 0.0
    return llv_start_idx, llv_end_idx, round(llv_total, 6)


def calculate_actual_top_elevation(
    design_top_elevation: float,
    cd_dinh: float,
    injection_start_depth: float | None,
) -> float:
    if injection_start_depth is not None and injection_start_depth < cd_dinh:
        return design_top_elevation + (cd_dinh - injection_start_depth)
    return design_top_elevation


def calculate_actual_bottom_elevation(
    design_bottom_elevation: float,
    actual_drill_depth: float,
    target_depth: float,
) -> float:
    return design_bottom_elevation - (actual_drill_depth - target_depth)
