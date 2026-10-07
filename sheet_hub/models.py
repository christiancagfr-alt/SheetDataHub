from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any


@dataclass
class SourceConfig:
    id: str
    name: str
    url: str
    include_sheets: list[str] = field(default_factory=list)
    exclude_sheets: list[str] = field(default_factory=list)
    header_row: int = 1
    enabled: bool = True
    credential_path: str = ""
    column_schema_enabled: bool = False
    column_schema: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, value: dict[str, Any]) -> "SourceConfig":
        allowed = {item.name for item in cls.__dataclass_fields__.values()}
        return cls(**{key: val for key, val in value.items() if key in allowed})


@dataclass
class Record:
    source_id: str
    source_name: str
    spreadsheet_id: str
    sheet_name: str
    row_number: int
    values: dict[str, str]
    row_hash: str

    def display(self) -> dict[str, str]:
        result = dict(self.values)
        result.update(
            {
                "数据源": self.source_name,
                "子Sheet": self.sheet_name,
                "原始行号": str(self.row_number),
            }
        )
        return result


@dataclass
class DailyPoint:
    day: str
    count: int
    total: float
    average: float | None
    growth: float | None
    growth_pct: float | None


@dataclass
class PeriodSnapshot:
    label: str
    start: str
    end: str
    count: int
    total: float
    average: float | None
    daily: list[DailyPoint] = field(default_factory=list)


@dataclass
class PersonSnapshot:
    name: str
    count: int
    total: float
    average: float | None
    previous_count: int = 0
    previous_total: float = 0.0


@dataclass
class HeaderStat:
    name: str
    count: int
    previous_count: int = 0
    total: float | None = None
    previous_total: float | None = None
    average: float | None = None
    numeric: bool = False


@dataclass
class BreakdownSeries:
    label: str
    daily: list[int] = field(default_factory=list)
    previous_daily: list[int] = field(default_factory=list)
    count: int = 0
    previous_count: int = 0


@dataclass
class HeaderBreakdown:
    header: str
    series: list[BreakdownSeries] = field(default_factory=list)


@dataclass
class AnalysisResult:
    metric_field: str
    metric_is_count: bool
    date_field: str
    name_field: str
    team_field: str
    team: str
    scope: str
    compare_mode: str
    names_found: list[str]
    headers: list[str]
    current: PeriodSnapshot
    previous: PeriodSnapshot
    delta_count: int
    delta_total: float
    delta_average: float | None
    delta_count_pct: float | None
    delta_total_pct: float | None
    people: list[PersonSnapshot] = field(default_factory=list)
    header_stats: list[HeaderStat] = field(default_factory=list)
    compare_enabled: bool = True
    breakdowns: list[HeaderBreakdown] = field(default_factory=list)

