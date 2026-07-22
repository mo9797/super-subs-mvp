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

    def test_partner_cash_settlement_is_visible(self):
        for partner in ("حمدي", "علي"):
            self.assertIn(f'data-partner="{partner}"', HTML)
        for field in ("paid", "received", "share", "due", "owes"):
            self.assertIn(f'data-field="{field}"', HTML)
        self.assertIn("function partnerSettlement", HTML)
        self.assertIn("profitShare=projectNet/2", HTML)
        self.assertIn("finalBalance=paid+profitShare-received", HTML)
        self.assertIn("Math.max(finalBalance,0)", HTML)
        self.assertIn("Math.max(-finalBalance,0)", HTML)
        self.assertIn("المكسب أو الخسارة مقسوم بالتساوي 50% لكل واحد", HTML)

    def test_advertising_expense_requires_a_payer(self):
        self.assertIn('id="adBy"', HTML)
        self.assertIn("اختار مين اللي دفع الإعلان", HTML)
        self.assertIn("type === 'ad' && VALID_PURCHASE_PARTIES", API)

    def test_current_session_and_overall_views_exist(self):
        self.assertIn('data-view="today"', HTML)
        self.assertIn('data-view="all"', HTML)
        self.assertIn("الجلسة الحالية", HTML)
        self.assertIn("الإجمالي العام", HTML)
        self.assertIn("function transactionDay", HTML)
        self.assertIn("viewTransactions", HTML)
        self.assertIn("day", HTML)
        self.assertIn("day", API)

    def test_named_sessions_can_be_created_and_selected(self):
        self.assertIn('id="sessionName"', HTML)
        self.assertIn('id="createSession"', HTML)
        self.assertIn('id="sessionSelect"', HTML)
        self.assertIn("createSession", HTML)
        self.assertIn("sessionId", HTML)
        self.assertIn("body.action === 'session'", API)
        self.assertIn("SESSIONS_SHEET", API)
        self.assertIn("readSessions_", API)

    def test_new_transactions_are_attached_to_the_active_session(self):
        self.assertIn("sessionId:currentSessionId", HTML)
        self.assertIn("sessionId", API)
        self.assertIn("assertSession_", API)
        self.assertIn("Session ID", API)
        self.assertIn("غير مصنفة", API)


if __name__ == "__main__":
    unittest.main()
