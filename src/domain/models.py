from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal
from enum import Enum


class StatementType(str, Enum):
    BALANCE_SHEET_ASSETS = "Bảng cân đối kế toán - Tài sản"
    BALANCE_SHEET_EQUITY = "Bảng cân đối kế toán - Nguồn vốn"
    INCOME_STATEMENT = "Báo cáo kết quả hoạt động kinh doanh"
    CASH_FLOW = "Báo cáo lưu chuyển tiền tệ"
    NOTE = "Thuyết minh"
    GENERAL_INFO = "Thông tin chung"
    SIGNATURE = "Chữ ký"
    UNKNOWN = "Chưa phân loại"


class AccountingRegime(str, Enum):
    TT200_2014 = "200/2014/TT-BTC"
    TT133_2016 = "133/2016/TT-BTC"
    TT99_2025 = "99/2025/TT-BTC"
    UNKNOWN = "unknown"
    AMBIGUOUS = "ambiguous"


class DetectionConfidence(str, Enum):
    HIGH = "Cao"
    MEDIUM = "Trung bình"
    LOW = "Thấp"
    NONE = "Không xác định"


@dataclass(frozen=True)
class AccountingRegimeDetection:
    regime: AccountingRegime
    confidence: DetectionConfidence
    evidence: tuple[str, ...] = ()
    message: str = ""

    @property
    def display_name(self) -> str:
        if self.regime == AccountingRegime.UNKNOWN:
            return "Chưa xác định"
        if self.regime == AccountingRegime.AMBIGUOUS:
            return "Có nhiều Thông tư có thể áp dụng"
        return f"Thông tư {self.regime.value}"


class TableValueStatus(str, Enum):
    NUMBER = "number"
    ZERO = "zero"
    DASH_ZERO = "dash_zero"
    MISSING = "missing"
    INVALID = "invalid"


class WordCommentCategory(str, Enum):
    GENERAL = "general"
    DIFFERENCE = "difference"
    REVIEW = "review"


@dataclass(frozen=True)
class TableValue:
    status: TableValueStatus
    value: Decimal | None = None


@dataclass(frozen=True)
class CellBorder:
    style: str = ""
    color: str = ""


@dataclass(frozen=True)
class CellPresentation:
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: str = ""
    fill_color: str = ""
    horizontal_alignment: str = ""
    vertical_alignment: str = ""
    top_border: CellBorder = field(default_factory=CellBorder)
    right_border: CellBorder = field(default_factory=CellBorder)
    bottom_border: CellBorder = field(default_factory=CellBorder)
    left_border: CellBorder = field(default_factory=CellBorder)


@dataclass(frozen=True)
class ExtractedTable:
    index: int
    rows: list[list[str]]
    statement_type: StatementType
    title_hint: str = ""
    context_hints: tuple[str, ...] = ()
    cell_presentations: dict[tuple[int, int], CellPresentation] = field(default_factory=dict)
    merged_ranges: tuple[tuple[int, int, int, int], ...] = ()
    column_widths: tuple[float, ...] = ()

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def column_count(self) -> int:
        return max((len(row) for row in self.rows), default=0)


@dataclass(frozen=True)
class TableCheckResult:
    table_index: int
    status: str
    check_count: int = 0
    issue_count: int = 0
    note: str = ""
    formula_cells: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class WordCommentTarget:
    table_index: int
    row_index: int
    column_index: int
    message: str
    category: WordCommentCategory = WordCommentCategory.GENERAL


@dataclass(frozen=True)
class NoteMatchResult:
    note_table_index: int
    note_title: str
    statement_table_index: int | None = None
    statement_current_cell: str = ""
    note_current_cell: str = ""
    statement_prior_cell: str = ""
    note_prior_cell: str = ""
    source: str = ""
    statement_code: str = ""
    note_code: str = ""
    item_name: str = ""
    statement_current: Decimal | None = None
    note_current: Decimal | None = None
    current_difference: Decimal | None = None
    statement_prior: Decimal | None = None
    note_prior: Decimal | None = None
    prior_difference: Decimal | None = None
    status: str = "Statement not found"
    match_type: str = ""
    match_score: float = 0.0
