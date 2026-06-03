from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(slots=True)
class TableRow:
    excel_row: int
    depth: float
    instant_flow: float
    flow_segment: float
    main_hoist_speed: float


@dataclass(slots=True)
class SheetMetadata:
    pile_no: str | None = None
    target_depth: float | None = None
    target_depth_raw: str | None = None
    diameter: float | None = None
    diameter_raw: str | None = None
    header_drill_depth: float | None = None
    header_drill_depth_raw: str | None = None
    accumulated_flow_raw: str | None = None
    start_time_raw: str | None = None
    end_time_raw: str | None = None


@dataclass(slots=True)
class TableLayout:
    header_row: int
    depth_col: int
    instant_flow_col: int
    flow_col: int
    speed_col: int
    last_data_col: int
    used_fallback: bool = False


@dataclass(slots=True)
class SheetCalculation:
    cd_dinh: float
    llv: float
    vx: float
    vl: float
    actual_top_elevation: float
    actual_bottom_elevation: float
    actual_column_length: float
    vx_start_idx: int
    vx_end_idx: int
    vl_start_idx: int
    vl_end_idx: int
    llv_start_idx: int
    llv_end_idx: int
    actual_drill_depth: float
    injection_start_depth: float | None = None
    notes: list[str] = field(default_factory=list)


@dataclass(slots=True)
class SheetResult:
    sheet_name: str
    pile_no: str | None
    target_depth: float | None
    header_drill_depth: float | None
    actual_drill_depth: float | None
    drill_length: float
    design_top_elevation: float
    design_bottom_elevation: float
    cd_dinh: float | None
    llv: float | None
    vx: float | None
    vl: float | None
    actual_top_elevation: float | None
    actual_bottom_elevation: float | None
    actual_column_length: float | None
    vx_range: str | None
    vl_range: str | None
    llv_range: str | None
    status: str
    note: str = ""

    def as_table_row(self) -> list[str]:
        return [
            self.sheet_name,
            self.pile_no or "",
            "" if self.cd_dinh is None else f"{self.cd_dinh:.2f}",
            "" if self.llv is None else f"{self.llv:.2f}",
            "" if self.vx is None else f"{self.vx:.4f}",
            "" if self.vl is None else f"{self.vl:.4f}",
            "" if self.actual_top_elevation is None else f"{self.actual_top_elevation:.2f}",
            "" if self.actual_bottom_elevation is None else f"{self.actual_bottom_elevation:.2f}",
            "" if self.actual_column_length is None else f"{self.actual_column_length:.2f}",
            self.status,
            self.note,
        ]
