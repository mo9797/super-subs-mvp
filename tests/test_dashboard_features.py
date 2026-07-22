import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HTML = (ROOT / "index.html").read_text(encoding="utf-8")
API = (ROOT / "Code.gs").read_text(encoding="utf-8")


class DashboardFeatureTests(unittest.TestCase):
    def test_financial_cards_have_semantic_classes(self):
        for cls in ("purchase-stat", "sale-stat", "ad-stat", "profit-stat"):
            self.assertIn(cls, HTML)
        self.assertIn("loss", HTML)

    def test_transaction_history_has_edit_and_delete_actions(self):
        self.assertIn("<th>إجراءات</th>", HTML)
        self.assertIn("data-action", HTML)
        self.assertIn("editTransaction", HTML)
        self.assertIn("deleteTransaction", HTML)
        self.assertIn("confirm(", HTML)

    def test_api_supports_id_based_update_and_delete(self):
        self.assertIn("body.action === 'update'", API)
        self.assertIn("body.action === 'delete'", API)
        self.assertIn("updateTransaction_", API)
        self.assertIn("deleteTransaction_", API)
        self.assertRegex(API, r"findTransactionRow_\s*\(")
        self.assertIn("LockService.getScriptLock()", API)

    def test_api_validates_financial_party_allowlist(self):
        self.assertIn("VALID_PURCHASE_PARTIES", API)
        self.assertIn("VALID_SALE_PARTIES", API)
        self.assertIn("validateParty_", API)


if __name__ == "__main__":
    unittest.main()
