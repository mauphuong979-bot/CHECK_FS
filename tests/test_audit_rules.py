import unittest

from src.domain.audit_rules import (
    ARITHMETIC_RULES_BY_TYPE,
    CASH_FLOW_SIGN_RULES,
    CROSS_PERIOD_RULES,
)
from src.domain.models import StatementType


class AuditRuleRegistryTest(unittest.TestCase):
    def test_arithmetic_rule_targets_are_preserved(self):
        expected_targets = {
            StatementType.BALANCE_SHEET_ASSETS: (
                "100", "110", "120", "130", "140", "150", "200", "210",
                "220", "221", "224", "227", "230", "240", "250", "260",
                "270", "270",
            ),
            StatementType.BALANCE_SHEET_EQUITY: (
                "300", "310", "330", "400", "410", "421", "421b", "430", "440",
            ),
            StatementType.INCOME_STATEMENT: ("10", "20", "30", "40", "50", "60"),
            StatementType.CASH_FLOW: ("1", "6", "8", "20", "30", "40", "50", "70"),
        }

        actual_targets = {
            statement_type: tuple(rule.target_code for rule in rules)
            for statement_type, rules in ARITHMETIC_RULES_BY_TYPE.items()
        }

        self.assertEqual(actual_targets, expected_targets)

    def test_cross_statement_rule_sources_and_coefficients_are_preserved(self):
        asset_cross_rule = ARITHMETIC_RULES_BY_TYPE[
            StatementType.BALANCE_SHEET_ASSETS
        ][-1]
        retained_earnings_rule = ARITHMETIC_RULES_BY_TYPE[
            StatementType.BALANCE_SHEET_EQUITY
        ][6]
        income_rule = ARITHMETIC_RULES_BY_TYPE[StatementType.INCOME_STATEMENT][0]

        self.assertEqual(
            tuple((term.code, term.coefficient, term.source) for term in asset_cross_rule.terms),
            (("440", 1, StatementType.BALANCE_SHEET_EQUITY),),
        )
        self.assertEqual(
            tuple((term.code, term.coefficient, term.source) for term in retained_earnings_rule.terms),
            (("60", 1, StatementType.INCOME_STATEMENT),),
        )
        self.assertEqual(
            tuple((term.code, term.coefficient) for term in income_rule.terms),
            (("1", 1), ("2", -1)),
        )

    def test_cash_flow_sign_rules_are_preserved(self):
        expected = {
            "2": "positive",
            "14": "negative",
            "15": "negative",
            "16": "positive",
            "17": "negative",
            "21": "negative",
            "22": "positive",
            "23": "negative",
            "24": "positive",
            "25": "negative",
            "26": "positive",
            "27": "positive",
            "31": "positive",
            "32": "negative",
            "33": "positive",
            "34": "negative",
            "35": "negative",
            "36": "negative",
        }

        self.assertEqual(
            {rule.code: rule.expected for rule in CASH_FLOW_SIGN_RULES},
            expected,
        )

    def test_cross_period_rules_are_preserved(self):
        actual = {
            statement_type: tuple(
                (
                    rule.target_code,
                    rule.source_code,
                    rule.source,
                    rule.target_period,
                    rule.source_period,
                )
                for rule in rules
            )
            for statement_type, rules in CROSS_PERIOD_RULES.items()
        }

        self.assertEqual(
            actual,
            {
                StatementType.BALANCE_SHEET_EQUITY: (
                    ("421a", "421", StatementType.BALANCE_SHEET_EQUITY, "current", "prior"),
                ),
                StatementType.CASH_FLOW: (
                    ("60", "110", StatementType.BALANCE_SHEET_ASSETS, "current", "prior"),
                    ("70", "110", StatementType.BALANCE_SHEET_ASSETS, "current", "current"),
                    ("70", "110", StatementType.BALANCE_SHEET_ASSETS, "prior", "prior"),
                ),
            },
        )


    def test_tt133_rule_pack_registration(self):
        from src.domain.audit_rules import TT133_RULE_PACK, get_rule_pack
        from src.domain.models import AccountingRegime

        pack = get_rule_pack(AccountingRegime.TT133_2016)
        self.assertIsNotNone(pack)
        self.assertEqual(pack, TT133_RULE_PACK)
        self.assertEqual(pack.regime, AccountingRegime.TT133_2016)
        self.assertIn("133", pack.display_name)


if __name__ == "__main__":
    unittest.main()

