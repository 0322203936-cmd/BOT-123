import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from bot.facturas_komet import (
    export_details_email_config,
    load_sent_order_keys,
    order_key,
    process_export_details,
    save_sent_order_keys,
)


class FacturasKometStateTests(unittest.TestCase):
    def test_order_key_uses_order_and_date(self):
        self.assertEqual(
            order_key({"order": "W000308", "date": "10/02/2026", "internal_id": "42"}),
            "W000308|10/02/2026",
        )

    def test_missing_state_is_empty(self):
        with tempfile.TemporaryDirectory() as directory:
            self.assertEqual(load_sent_order_keys(Path(directory) / "missing.json"), set())

    def test_state_round_trip(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "facturas_enviadas.json"
            save_sent_order_keys({"W000308|10/02/2026", "000312|10/09/2026"}, path)
            self.assertEqual(
                load_sent_order_keys(path),
                {"W000308|10/02/2026", "000312|10/09/2026"},
            )
            self.assertEqual(json.loads(path.read_text(encoding="utf-8"))["version"], 1)


class ExportDetailsTests(unittest.TestCase):
    def setUp(self):
        self.order = {
            "internal_id": "42",
            "checkbox_id": "jqg_gridResults_42",
            "order": "000316",
            "date": "10/08/2026",
        }
        self.details = {
            "Order Number": "000316",
            "Customer": "WFS - SPOKANE",
            "Ship Date": "10/15/2026",
            "Carrier": "Prime Floral",
            "Location": "Pacifica Farms (PPFF)",
            "Created on": "10/08/2026 15:02:27",
        }

    def test_email_contains_all_requested_fields_and_escapes_customer(self):
        details = {**self.details, "Customer": "A&B <Buyer>"}
        config = export_details_email_config(details, "Irene.Machain@Pacifica-farms.com")
        self.assertEqual(config["subject"], "Nueva Orden Komet: 000316")
        self.assertEqual(config["recipients"], ["Irene.Machain@Pacifica-farms.com"])
        self.assertIn("A&amp;B &lt;Buyer&gt;", config["bodyHtml"])
        self.assertNotIn("Sales Person", config["bodyHtml"])
        for field in self.details:
            self.assertIn(field, config["bodyHtml"])

    def test_order_details_status_email_sends_to_irene_and_jesus(self):
        config = export_details_email_config(
            self.details, "Irene.Machain@Pacifica-farms.com; jesus.sandoval@cfbc.co",
            posco_status="applied_confirmed",
        )
        self.assertEqual(config["recipients"], [
            "Irene.Machain@Pacifica-farms.com", "jesus.sandoval@cfbc.co",
        ])
        self.assertIn("Orden actualizada y confirmada en POSCO", config["bodyHtml"])

    def test_posco_review_is_amber_and_manual_failure_is_red(self):
        review = export_details_email_config(
            self.details, "irene@example.com", posco_status="staged_for_review",
        )
        self.assertIn("#9A6700", review["bodyHtml"])
        self.assertIn("pendiente de Actualizar", review["bodyHtml"])
        self.assertNotIn("actualizada y confirmada", review["bodyHtml"])

        manual = export_details_email_config(
            self.details, "irene@example.com", posco_status="error",
            posco_reason="Pack < 10 & Stems no coincide",
        )
        self.assertIn("#B42318", manual["bodyHtml"])
        self.assertIn("Pack &lt; 10 &amp; Stems", manual["bodyHtml"])

    def test_green_status_requires_explicit_final_confirmation(self):
        confirmed = export_details_email_config(
            self.details, "irene@example.com", posco_status="applied_confirmed",
        )
        self.assertIn("#137333", confirmed["bodyHtml"])
        self.assertIn("Orden actualizada y confirmada en POSCO", confirmed["bodyHtml"])

    def test_existing_original_email_does_not_block_one_new_detail_email(self):
        with tempfile.TemporaryDirectory() as directory:
            ledger = Path(directory) / "details.json"
            key = order_key(self.order)
            sent = set()
            with (
                patch("bot.facturas_komet.SENT_EXPORT_DETAILS_PATH", ledger),
                patch("bot.facturas_komet.download_export_details", return_value=Path("order.xls")) as download,
                patch("bot.facturas_komet.parse_export_details", return_value=self.details),
                patch("bot.facturas_komet.send_export_details_email") as send,
            ):
                first = process_export_details(None, self.order, 1, key, "irene@example.com", sent)
                second = process_export_details(None, self.order, 1, key, "irene@example.com", sent)
            self.assertEqual(first["status"], "enviado")
            self.assertEqual(second["status"], "omitido_ya_enviado")
            self.assertEqual(load_sent_order_keys(ledger), {key})
            download.assert_called_once()
            send.assert_called_once()

    def test_failed_send_does_not_mark_detail_as_sent(self):
        with tempfile.TemporaryDirectory() as directory:
            ledger = Path(directory) / "details.json"
            key = order_key(self.order)
            sent = set()
            with (
                patch("bot.facturas_komet.SENT_EXPORT_DETAILS_PATH", ledger),
                patch("bot.facturas_komet.download_export_details", return_value=Path("order.xls")),
                patch("bot.facturas_komet.parse_export_details", return_value=self.details),
                patch("bot.facturas_komet.send_export_details_email", side_effect=RuntimeError("Graph failed")),
            ):
                with self.assertRaisesRegex(RuntimeError, "Graph failed"):
                    process_export_details(None, self.order, 1, key, "irene@example.com", sent)
            self.assertEqual(sent, set())
            self.assertFalse(ledger.exists())

    def test_wrong_downloaded_order_is_never_emailed(self):
        with (
            patch("bot.facturas_komet.download_export_details", return_value=Path("order.xls")),
            patch("bot.facturas_komet.parse_export_details", return_value={**self.details, "Order Number": "000317"}),
            patch("bot.facturas_komet.send_export_details_email") as send,
        ):
            with self.assertRaisesRegex(RuntimeError, "000317"):
                process_export_details(None, self.order, 1, order_key(self.order), "irene@example.com", set())
        send.assert_not_called()

    def test_k2k_prefix_matches_six_digit_xls_order(self):
        order = {**self.order, "order": "K2K 000316"}
        with tempfile.TemporaryDirectory() as directory:
            ledger = Path(directory) / "details.json"
            with (
                patch("bot.facturas_komet.SENT_EXPORT_DETAILS_PATH", ledger),
                patch("bot.facturas_komet.download_export_details", return_value=Path("order.xls")),
                patch("bot.facturas_komet.parse_export_details", return_value=self.details),
                patch("bot.facturas_komet.send_export_details_email") as send,
            ):
                result = process_export_details(None, order, 1, order_key(order), "irene@example.com", set())
        self.assertEqual(result["status"], "enviado")
        send.assert_called_once()


if __name__ == "__main__":
    unittest.main()
