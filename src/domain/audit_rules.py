from __future__ import annotations

from dataclasses import dataclass, field
from typing import Mapping

from src.domain.models import AccountingRegime, StatementType


@dataclass(frozen=True)
class RuleTerm:
    code: str
    coefficient: int = 1
    source: StatementType | None = None


@dataclass(frozen=True)
class ArithmeticRule:
    target_code: str
    terms: tuple[RuleTerm, ...]
    note: str = "Khác số chi tiết"


@dataclass(frozen=True)
class SignRule:
    code: str
    expected: str
    note: str


@dataclass(frozen=True)
class CrossPeriodRule:
    target_code: str
    source_code: str
    source: StatementType
    target_period: str
    source_period: str
    note: str


@dataclass(frozen=True)
class RulePack:
    regime: AccountingRegime
    display_name: str
    version: str
    arithmetic_rules: Mapping[StatementType, tuple[ArithmeticRule, ...]] = field(default_factory=dict)
    sign_rules: Mapping[StatementType, tuple[SignRule, ...]] = field(default_factory=dict)
    cross_period_rules: Mapping[StatementType, tuple[CrossPeriodRule, ...]] = field(default_factory=dict)


ASSET_RULES: tuple[ArithmeticRule, ...] = (
    ArithmeticRule("100", (RuleTerm("110"), RuleTerm("120"), RuleTerm("130"), RuleTerm("140"), RuleTerm("150"))),
    ArithmeticRule("110", (RuleTerm("111"), RuleTerm("112"))),
    ArithmeticRule("120", (RuleTerm("121"), RuleTerm("122"), RuleTerm("123"))),
    ArithmeticRule("130", (RuleTerm("131"), RuleTerm("132"), RuleTerm("133"), RuleTerm("134"), RuleTerm("135"), RuleTerm("136"), RuleTerm("137"), RuleTerm("139"))),
    ArithmeticRule("140", (RuleTerm("141"), RuleTerm("149"))),
    ArithmeticRule("150", (RuleTerm("151"), RuleTerm("152"), RuleTerm("153"), RuleTerm("154"), RuleTerm("155"))),
    ArithmeticRule("200", (RuleTerm("210"), RuleTerm("220"), RuleTerm("230"), RuleTerm("240"), RuleTerm("250"), RuleTerm("260"))),
    ArithmeticRule("210", (RuleTerm("211"), RuleTerm("212"), RuleTerm("213"), RuleTerm("214"), RuleTerm("215"), RuleTerm("216"), RuleTerm("219"))),
    ArithmeticRule("220", (RuleTerm("221"), RuleTerm("224"), RuleTerm("227"))),
    ArithmeticRule("221", (RuleTerm("222"), RuleTerm("223"))),
    ArithmeticRule("224", (RuleTerm("225"), RuleTerm("226"))),
    ArithmeticRule("227", (RuleTerm("228"), RuleTerm("229"))),
    ArithmeticRule("230", (RuleTerm("231"), RuleTerm("232"))),
    ArithmeticRule("240", (RuleTerm("241"), RuleTerm("242"))),
    ArithmeticRule("250", (RuleTerm("251"), RuleTerm("252"), RuleTerm("253"), RuleTerm("254"), RuleTerm("255"))),
    ArithmeticRule("260", (RuleTerm("261"), RuleTerm("262"), RuleTerm("263"), RuleTerm("268"))),
    ArithmeticRule("270", (RuleTerm("100"), RuleTerm("200"))),
    ArithmeticRule("270", (RuleTerm("440", source=StatementType.BALANCE_SHEET_EQUITY),), "Tổng Tài sản khác Tổng Nguồn vốn"),
)


EQUITY_RULES: tuple[ArithmeticRule, ...] = (
    ArithmeticRule("300", (RuleTerm("310"), RuleTerm("330"))),
    ArithmeticRule("310", (RuleTerm("311"), RuleTerm("312"), RuleTerm("313"), RuleTerm("314"), RuleTerm("315"), RuleTerm("316"), RuleTerm("317"), RuleTerm("318"), RuleTerm("319"), RuleTerm("320"), RuleTerm("321"), RuleTerm("322"), RuleTerm("323"), RuleTerm("324"), RuleTerm("325"))),
    ArithmeticRule("330", (RuleTerm("331"), RuleTerm("332"), RuleTerm("333"), RuleTerm("334"), RuleTerm("335"), RuleTerm("336"), RuleTerm("337"), RuleTerm("338"), RuleTerm("339"), RuleTerm("340"), RuleTerm("341"), RuleTerm("342"), RuleTerm("343"))),
    ArithmeticRule("400", (RuleTerm("411"), RuleTerm("412"), RuleTerm("413"), RuleTerm("414"), RuleTerm("415"), RuleTerm("416"), RuleTerm("417"), RuleTerm("418"), RuleTerm("419"), RuleTerm("420"), RuleTerm("421"), RuleTerm("422"), RuleTerm("431"), RuleTerm("432"))),
    ArithmeticRule("410", (RuleTerm("411"), RuleTerm("412"), RuleTerm("413"), RuleTerm("414"), RuleTerm("415"), RuleTerm("416"), RuleTerm("417"), RuleTerm("418"), RuleTerm("419"), RuleTerm("420"), RuleTerm("421"), RuleTerm("422"))),
    ArithmeticRule("421", (RuleTerm("421a"), RuleTerm("421b")), "Khác với lũy kế kỳ trước và kỳ này"),
    ArithmeticRule("421b", (RuleTerm("60", source=StatementType.INCOME_STATEMENT),), "Khác với lợi nhuận sau thuế trên PL"),
    ArithmeticRule("430", (RuleTerm("431"), RuleTerm("432"))),
    ArithmeticRule("440", (RuleTerm("300"), RuleTerm("400"))),
)


INCOME_RULES: tuple[ArithmeticRule, ...] = (
    ArithmeticRule("10", (RuleTerm("1"), RuleTerm("2", -1))),
    ArithmeticRule("20", (RuleTerm("10"), RuleTerm("11", -1))),
    ArithmeticRule("30", (RuleTerm("20"), RuleTerm("21"), RuleTerm("22", -1), RuleTerm("25", -1), RuleTerm("26", -1))),
    ArithmeticRule("40", (RuleTerm("31"), RuleTerm("32", -1))),
    ArithmeticRule("50", (RuleTerm("30"), RuleTerm("40"))),
    ArithmeticRule("60", (RuleTerm("50"), RuleTerm("51", -1), RuleTerm("52", -1))),
)


CASH_FLOW_RULES: tuple[ArithmeticRule, ...] = (
    ArithmeticRule("1", (RuleTerm("50", source=StatementType.INCOME_STATEMENT),), "Khác với lợi nhuận trước thuế trên PL"),
    ArithmeticRule("6", (RuleTerm("23", source=StatementType.INCOME_STATEMENT),), "Khác với chi phí lãi vay trên PL"),
    ArithmeticRule("8", (RuleTerm("1"), RuleTerm("2"), RuleTerm("3"), RuleTerm("4"), RuleTerm("5"), RuleTerm("6"), RuleTerm("7"))),
    ArithmeticRule("20", (RuleTerm("8"), RuleTerm("9"), RuleTerm("10"), RuleTerm("11"), RuleTerm("12"), RuleTerm("13"), RuleTerm("14"), RuleTerm("15"), RuleTerm("16"), RuleTerm("17"))),
    ArithmeticRule("30", (RuleTerm("21"), RuleTerm("22"), RuleTerm("23"), RuleTerm("24"), RuleTerm("25"), RuleTerm("26"), RuleTerm("27"))),
    ArithmeticRule("40", (RuleTerm("31"), RuleTerm("32"), RuleTerm("33"), RuleTerm("34"), RuleTerm("35"), RuleTerm("36"))),
    ArithmeticRule("50", (RuleTerm("20"), RuleTerm("30"), RuleTerm("40"))),
    ArithmeticRule("70", (RuleTerm("50"), RuleTerm("60"), RuleTerm("61"))),
)


CASH_FLOW_SIGN_RULES: tuple[SignRule, ...] = (
    SignRule("2", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("14", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("15", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("16", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("17", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("21", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("22", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("23", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("24", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("25", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("26", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("27", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("31", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("32", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("33", "positive", "Chỉ tiêu này của CF phải là số dương (+)"),
    SignRule("34", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("35", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
    SignRule("36", "negative", "Chỉ tiêu này của CF phải là số âm (-)"),
)


CROSS_PERIOD_RULES: dict[StatementType, tuple[CrossPeriodRule, ...]] = {
    StatementType.BALANCE_SHEET_EQUITY: (
        CrossPeriodRule("421a", "421", StatementType.BALANCE_SHEET_EQUITY, "current", "prior", "Khác với kỳ trước, cần xem lại có chia lãi hay không"),
    ),
    StatementType.CASH_FLOW: (
        CrossPeriodRule("60", "110", StatementType.BALANCE_SHEET_ASSETS, "current", "prior", "Tiền đầu năm khác với BS"),
        CrossPeriodRule("70", "110", StatementType.BALANCE_SHEET_ASSETS, "current", "current", "Tiền cuối năm khác với BS"),
        CrossPeriodRule("70", "110", StatementType.BALANCE_SHEET_ASSETS, "prior", "prior", "Tiền cuối kỳ so sánh khác với BS"),
    ),
}


ARITHMETIC_RULES_BY_TYPE: dict[StatementType, tuple[ArithmeticRule, ...]] = {
    StatementType.BALANCE_SHEET_ASSETS: ASSET_RULES,
    StatementType.BALANCE_SHEET_EQUITY: EQUITY_RULES,
    StatementType.INCOME_STATEMENT: INCOME_RULES,
    StatementType.CASH_FLOW: CASH_FLOW_RULES,
}


TT200_RULE_PACK = RulePack(
    regime=AccountingRegime.TT200_2014,
    display_name="Thông tư 200/2014/TT-BTC",
    version="tt200-v1",
    arithmetic_rules=ARITHMETIC_RULES_BY_TYPE,
    sign_rules={StatementType.CASH_FLOW: CASH_FLOW_SIGN_RULES},
    cross_period_rules=CROSS_PERIOD_RULES,
)

TT133_RULE_PACK = RulePack(
    regime=AccountingRegime.TT133_2016,
    display_name="Thông tư 133/2016/TT-BTC",
    version="tt133-v1",
    arithmetic_rules=ARITHMETIC_RULES_BY_TYPE,
    sign_rules={StatementType.CASH_FLOW: CASH_FLOW_SIGN_RULES},
    cross_period_rules=CROSS_PERIOD_RULES,
)

EMPTY_RULE_PACK = RulePack(
    regime=AccountingRegime.UNKNOWN,
    display_name="Chưa áp dụng bộ rule nghiệp vụ",
    version="none",
)

RULE_PACKS: dict[AccountingRegime, RulePack] = {
    AccountingRegime.TT200_2014: TT200_RULE_PACK,
    AccountingRegime.TT133_2016: TT133_RULE_PACK,
}


def get_rule_pack(regime: AccountingRegime) -> RulePack | None:
    return RULE_PACKS.get(regime)
