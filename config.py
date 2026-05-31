from __future__ import annotations

from dataclasses import asdict, dataclass, field


@dataclass(slots=True)
class RuntimeConfig:
    default_drill_length_m: float = 3.0
    default_design_top_elevation: float = -10.5
    default_design_bottom_elevation: float = -13.5
    depth_tolerance: float = 0.02
    speed_zero_threshold: float = 0.05
    flow_min_for_llv: float = 10.0
    sum_only_values_greater_than_threshold: bool = True
    include_flow_equal_threshold: bool = True
    exclude_zero_speed_from_average: bool = False
    create_summary_sheet: bool = True
    highlight_ranges: bool = True
    do_not_overwrite_original: bool = True
    output_suffix: str = "_da_tinh"
    result_labels: dict[str, str] = field(
        default_factory=lambda: {
            "cd_dinh": "CĐ đỉnh",
            "llv": "LLV",
            "vx": "VX",
            "vl": "VL",
            "actual_column_length": "Chiều dài cọc thực tế",
            "actual_top_elevation": "Cao độ đỉnh thực tế",
            "actual_bottom_elevation": "Cao độ đáy thực tế",
            "vx_range": "VX range",
            "vl_range": "VL range",
            "llv_range": "LLV range",
            "status": "Trạng thái",
            "note": "Ghi chú",
        }
    )
    fallback_columns: dict[str, str] = field(
        default_factory=lambda: {
            "depth": "C",
            "flow_segment": "F",
            "main_hoist_speed": "H",
        }
    )
    highlight_colors: dict[str, str] = field(
        default_factory=lambda: {
            "vx": "FFF2CC",
            "vl": "D9EAD3",
            "llv": "FCE4D6",
        }
    )


DEFAULT_CONFIG = RuntimeConfig()


def clone_config(config: RuntimeConfig | None = None) -> RuntimeConfig:
    source = config or DEFAULT_CONFIG
    return RuntimeConfig(**asdict(source))

SUMMARY_HEADERS = [
    "STT",
    "Sheet",
    "Số cọc",
    "Độ sâu mục tiêu",
    "Khoan sâu header",
    "Khoan sâu thực tế",
    "Độ dài mũi khoan",
    "Cao độ đỉnh thiết kế",
    "Cao độ đáy thiết kế",
    "CĐ đỉnh",
    "LLV",
    "VX",
    "VL",
    "Cao độ đỉnh thực tế",
    "Cao độ đáy thực tế",
    "Chiều dài cọc thực tế",
    "VX range",
    "VL range",
    "LLV range",
    "Trạng thái",
    "Ghi chú",
]
