import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import Mock, patch

from bot import posco_order_artifacts as artifacts
from bot import posco_order_import as posco_import


class PoscoOrderImportTests(unittest.TestCase):
    def test_login_uses_importar_ordenes_and_not_new_format(self):
        page = Mock()
        posco_import.open_import_screen(page, "user", "password")
        page.goto.assert_called_once_with(posco_import.POSCO_URL, wait_until="domcontentloaded", timeout=60_000)
        page.get_by_role.assert_any_call("link", name="Importar Ordenes", exact=True)
        page.get_by_role.assert_any_call("button", name="Revisar Archivo")
        self.assertNotIn("Actualizar", str(page.mock_calls))

    def test_stage_selects_pacifica_andres_and_never_updates(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "new-order.xlsx"
            workbook.write_bytes(b"xlsx")
            page = Mock()
            page.get_by_text.return_value.is_visible.return_value = False
            result = posco_import.stage_workbook(page, workbook, "W000400", Path(directory) / "review.png")
            page.locator.assert_any_call('input[type="file"][name="file"]')
            page.locator.assert_any_call('select[name="upload_mode"]')
            page.locator('input[type="file"][name="file"]').set_input_files.assert_called_with(str(workbook))
            page.locator('select[name="upload_mode"]').select_option.assert_called_with("pacifica2")
            page.get_by_role.assert_any_call("button", name="Upload", exact=True)
            page.get_by_role("heading", name="Import excel").wait_for.assert_called_with(state="hidden", timeout=600_000)
            self.assertEqual(result, "uploaded")
            self.assertNotIn("No elements found", str(page.mock_calls))
            page.screenshot.assert_called_once_with(path=str(Path(directory) / "review.png"), full_page=True)
            self.assertNotIn('name="Actualizar"', str(page.mock_calls))

    def test_sin_cambios_is_a_successful_upload_result(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "existing-order.xlsx"
            workbook.write_bytes(b"xlsx")
            page = Mock()
            page.get_by_text("Sin Cambios", exact=True).is_visible.return_value = True
            result = posco_import.stage_workbook(page, workbook, "W000317", Path(directory) / "review.png")
            self.assertEqual(result, "no_changes")
            page.screenshot.assert_called_once()
            self.assertNotIn('name="Actualizar"', str(page.mock_calls))

    def test_stage_ledger_preserves_legacy_and_uncertain_orders(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "stage.json"
            state = {
                "legacy_orders": {"W000317|10/12/2026"},
                "staged_orders": {"W000400|10/20/2026"},
                "needs_review": {"W000401|10/20/2026"},
            }
            posco_import.save_stage_state(path, state)
            self.assertEqual(posco_import.load_stage_state(path), state)

    def test_cleared_stage_ledger_regenerates_old_orders_but_skips_manual(self):
        orders = [
            {"order": "000312", "date": "10/09/2026", "internal_id": "312"},
            {"order": "W000317", "date": "10/12/2026", "internal_id": "317"},
            {"order": "W000308", "date": "10/09/2026", "internal_id": "308"},
        ]
        legacy_key = "000312|10/09/2026"
        staged_key = "W000317|10/12/2026"
        manual_key = "W000308|10/09/2026"
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            ledger_path = root / "transformed.json"
            stage_path = root / "stage.json"
            ledger_path.write_text("existing transformation ledger", encoding="utf-8")
            state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
            posco_import.save_stage_state(stage_path, state)
            staged_orders = []

            def stage_workbook(_page, _workbook, order_number, _screenshot):
                staged_orders.append(order_number)
                return "no_changes"

            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", ledger_path),
                patch.object(artifacts, "STAGE_LEDGER_PATH", stage_path),
                patch.object(artifacts, "collect_filtered_orders", return_value=orders),
                patch.object(artifacts, "download_with_recovery", side_effect=lambda _p, order, *_: Path(order["order"] + ".xls")) as download,
                patch.object(artifacts.transform, "parse_order_xls", side_effect=lambda path: ({"order_number": path.stem[-6:]}, [{}])),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, path: path.write_bytes(b"xlsx")),
                patch.object(artifacts.transform, "save_transformed_keys") as save_transformed,
                patch.object(artifacts.posco_import, "stage_workbook", side_effect=stage_workbook),
            ):
                report = {"orders": []}
                artifacts.transform_visible_orders(
                    None, {legacy_key, staged_key}, {manual_key}, {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 19),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                )

            self.assertEqual(staged_orders, ["000312", "W000317"])
            self.assertEqual(download.call_count, 2)
            self.assertEqual([item["status"] for item in report["orders"]], [
                "already_in_posco", "already_in_posco", "manual",
            ])
            self.assertEqual(len(list(root.glob("*.xlsx"))), 2)
            self.assertEqual(ledger_path.read_text(encoding="utf-8"), "existing transformation ledger")
            self.assertEqual(state, {"legacy_orders": set(), "staged_orders": {legacy_key, staged_key}, "needs_review": set()})
            self.assertEqual(posco_import.load_stage_state(stage_path), state)
            save_transformed.assert_not_called()

    def test_new_order_stages_but_old_order_is_skipped(self):
        orders = [
            {"order": "W000317", "date": "10/12/2026", "internal_id": "317"},
            {"order": "W000400", "date": "10/20/2026", "internal_id": "400"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": {"W000317|10/12/2026"}, "staged_orders": set(), "needs_review": set()}
            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=orders),
                patch.object(artifacts, "download_with_recovery", return_value=Path("W000400.xls")) as download,
                patch.object(artifacts.transform, "parse_order_xls", return_value=(
                    {"order_number": "000400", "customer": "Test", "ship_date_value": date(2026, 10, 20)}, [{}],
                )),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, p: p.write_bytes(b"xlsx")),
                patch.object(artifacts.posco_import, "stage_workbook") as stage,
            ):
                report = {"orders": []}
                artifacts.transform_visible_orders(
                    None, {"W000317|10/12/2026"}, set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                )
            self.assertEqual([item["status"] for item in report["orders"]], ["legacy", "staged_for_review"])
            self.assertEqual(download.call_count, 1)
            stage.assert_called_once()
            self.assertEqual(state["staged_orders"], {"W000400|10/20/2026"})
            self.assertEqual(state["needs_review"], set())
            self.assertEqual(posco_import.load_stage_state(root / "stage.json"), state)

    def test_uncertain_upload_does_not_retry(self):
        orders = [{"order": "W000400", "date": "10/20/2026", "internal_id": "400"}]
        state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": {"W000400|10/20/2026"}}
        with (
            patch.object(artifacts, "collect_filtered_orders", return_value=orders),
            patch.object(artifacts, "download_with_recovery") as download,
        ):
            report = {"orders": []}
            artifacts.transform_visible_orders(
                None, {"W000400|10/20/2026"}, set(), {}, b"template", report,
                date(2026, 10, 9), date(2026, 10, 20),
                stage_mode="all", stage_page=Mock(), stage_state=state,
            )
        self.assertEqual(report["orders"][0]["status"], "needs_manual_review")
        download.assert_not_called()

    def test_new_order_mails_original_details_after_posco_review_once(self):
        order = {"order": "W000400", "date": "10/20/2026", "internal_id": "400"}
        key = "W000400|10/20/2026"
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
            sent_details = set()
            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts.komet, "SENT_EXPORT_DETAILS_PATH", root / "details.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=[order]),
                patch.object(artifacts, "download_with_recovery", return_value=root / "order.xls") as download,
                patch.object(artifacts.transform, "parse_order_xls", return_value=({"order_number": "000400"}, [{}])),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, path: path.write_bytes(b"xlsx")),
                patch.object(artifacts.komet, "parse_export_details", return_value={"Order Number": "000400"}),
                patch.object(artifacts.komet, "send_export_details_email") as send,
                patch.object(artifacts.posco_import, "stage_workbook", return_value="uploaded") as stage,
            ):
                report = {"orders": [], "email_errors": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                    mail_recipient="irene@example.com", sent_detail_keys=sent_details,
                    sent_invoice_keys={key},
                )
                self.assertEqual(report["orders"][0]["status"], "staged_for_review")
                self.assertEqual(report["email_errors"], [])
                self.assertEqual(report["details_email_sent"], [
                    {"order": "W000400", "status": "staged_for_review"},
                ])
                self.assertEqual(download.call_count, 1)
                stage.assert_called_once()
                send.assert_called_once_with(
                    root / "order.xls", {"Order Number": "000400"}, "irene@example.com",
                    posco_status="staged_for_review", posco_reason="",
                )
                self.assertEqual(artifacts.komet.load_sent_order_keys(root / "details.json"), {key})

                # A later run neither uploads nor sends this order again.
                artifacts.transform_visible_orders(
                    None, {key}, set(), {}, b"template", {"orders": [], "email_errors": []},
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                    mail_recipient="irene@example.com", sent_detail_keys=sent_details,
                    sent_invoice_keys={key},
                )
                self.assertEqual(download.call_count, 1)
                send.assert_called_once()
                stage.assert_called_once()

    def test_failed_status_mail_retries_without_restaging_order(self):
        order = {"order": "W000400", "date": "10/20/2026", "internal_id": "400"}
        key = "W000400|10/20/2026"
        state = {"legacy_orders": set(), "staged_orders": {key}, "needs_review": set()}
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with (
                patch.object(artifacts.komet, "SENT_EXPORT_DETAILS_PATH", root / "details.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=[order]),
                patch.object(artifacts, "download_with_recovery", return_value=root / "order.xls") as download,
                patch.object(artifacts.komet, "parse_export_details", return_value={"Order Number": "000400"}),
                patch.object(artifacts.komet, "send_export_details_email", side_effect=[RuntimeError("Graph failed"), None]) as send,
                patch.object(artifacts.posco_import, "stage_workbook") as stage,
            ):
                sent_details = set()
                first = {"orders": [], "email_errors": []}
                artifacts.transform_visible_orders(
                    None, {key}, set(), {}, b"template", first,
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                    mail_recipient="irene@example.com", sent_detail_keys=sent_details,
                    sent_invoice_keys={key},
                )
                self.assertEqual(len(first["email_errors"]), 1)
                self.assertNotIn(key, sent_details)
                second = {"orders": [], "email_errors": []}
                artifacts.transform_visible_orders(
                    None, {key}, set(), {}, b"template", second,
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                    mail_recipient="irene@example.com", sent_detail_keys=sent_details,
                    sent_invoice_keys={key},
                )
                self.assertEqual(second["email_errors"], [])
                self.assertEqual(send.call_count, 2)
                self.assertEqual(download.call_count, 2)
                stage.assert_not_called()
                self.assertEqual(artifacts.komet.load_sent_order_keys(root / "details.json"), {key})

    def test_uncertain_posco_upload_sends_red_status_once_and_stops(self):
        orders = [
            {"order": "W000400", "date": "10/20/2026", "internal_id": "400"},
            {"order": "W000401", "date": "10/21/2026", "internal_id": "401"},
        ]
        key = "W000400|10/20/2026"
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts.komet, "SENT_EXPORT_DETAILS_PATH", root / "details.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=orders),
                patch.object(artifacts, "download_with_recovery", return_value=root / "order.xls") as download,
                patch.object(artifacts.transform, "parse_order_xls", return_value=({"order_number": "000400"}, [{}])),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, path: path.write_bytes(b"xlsx")),
                patch.object(artifacts.komet, "parse_export_details", return_value={"Order Number": "000400"}),
                patch.object(artifacts.komet, "send_export_details_email") as send,
                patch.object(artifacts.posco_import, "stage_workbook", side_effect=RuntimeError("timeout")) as stage,
            ):
                report = {"orders": [], "email_errors": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 21),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                    mail_recipient="irene@example.com", sent_detail_keys=set(),
                    sent_invoice_keys={key},
                )
                self.assertEqual(len(report["orders"]), 1)
                self.assertEqual(download.call_count, 1)
                stage.assert_called_once()
                send.assert_called_once_with(
                    root / "order.xls", {"Order Number": "000400"}, "irene@example.com",
                    posco_status="needs_manual_review",
                    posco_reason="POSCO no confirmó la carga del archivo.",
                )
                self.assertEqual(state["needs_review"], {key})
                self.assertEqual(artifacts.komet.load_sent_order_keys(root / "details.json"), {key})

    def test_transient_download_error_leaves_details_mail_pending(self):
        order = {"order": "W000400", "date": "10/20/2026", "internal_id": "400"}
        key = "W000400|10/20/2026"
        state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
        with (
            patch.object(artifacts, "collect_filtered_orders", return_value=[order]),
            patch.object(artifacts, "download_with_recovery", side_effect=RuntimeError("temporary network error")),
            patch.object(artifacts.komet, "send_export_details_email") as send,
        ):
            report = {"orders": [], "email_errors": []}
            sent_details = set()
            artifacts.transform_visible_orders(
                None, set(), set(), {}, b"template", report,
                date(2026, 10, 9), date(2026, 10, 20),
                stage_mode="all", stage_page=Mock(), stage_state=state,
                mail_recipient="irene@example.com", sent_detail_keys=sent_details,
                sent_invoice_keys={key},
            )
        self.assertEqual(report["orders"][0]["status"], "error")
        self.assertEqual(report["email_errors"], [])
        self.assertEqual(sent_details, set())
        send.assert_not_called()

    def test_homologation_error_mails_manual_status_with_reason(self):
        order = {"order": "W000400", "date": "10/20/2026", "internal_id": "400"}
        key = "W000400|10/20/2026"
        state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with (
                patch.object(artifacts.komet, "SENT_EXPORT_DETAILS_PATH", root / "details.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=[order]),
                patch.object(artifacts, "download_with_recovery", return_value=root / "order.xls") as download,
                patch.object(artifacts.transform, "parse_order_xls", side_effect=artifacts.transform.TransformationError("Pack distinto")),
                patch.object(artifacts.komet, "parse_export_details", return_value={"Order Number": "000400"}),
                patch.object(artifacts.komet, "send_export_details_email") as send,
                patch.object(artifacts.posco_import, "stage_workbook") as stage,
            ):
                report = {"orders": [], "email_errors": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                    mail_recipient="irene@example.com", sent_detail_keys=set(),
                    sent_invoice_keys={key},
                )
                self.assertEqual(report["orders"][0]["status"], "error")
                self.assertEqual(download.call_count, 1)
                stage.assert_not_called()
                send.assert_called_once_with(
                    root / "order.xls", {"Order Number": "000400"}, "irene@example.com",
                    posco_status="error", posco_reason="Pack distinto",
                )
                self.assertEqual(artifacts.komet.load_sent_order_keys(root / "details.json"), {key})

    def test_details_mail_waits_for_original_komet_documents(self):
        order = {"order": "W000400", "date": "10/20/2026", "internal_id": "400"}
        key = "W000400|10/20/2026"
        state = {"legacy_orders": set(), "staged_orders": {key}, "needs_review": set()}
        with (
            patch.object(artifacts, "collect_filtered_orders", return_value=[order]),
            patch.object(artifacts, "download_with_recovery") as download,
            patch.object(artifacts.komet, "send_export_details_email") as send,
        ):
            artifacts.transform_visible_orders(
                None, {key}, set(), {}, b"template", {"orders": [], "email_errors": []},
                date(2026, 10, 9), date(2026, 10, 20),
                stage_mode="all", stage_page=Mock(), stage_state=state,
                mail_recipient="irene@example.com", sent_detail_keys=set(), sent_invoice_keys=set(),
            )
        download.assert_not_called()
        send.assert_not_called()

    def test_all_mode_uploads_two_new_orders_sequentially(self):
        orders = [
            {"order": "W000400", "date": "10/20/2026", "internal_id": "400"},
            {"order": "W000401", "date": "10/21/2026", "internal_id": "401"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
            staged_in_order = []

            def record_stage(_page, _workbook, order_number, _screenshot):
                staged_in_order.append(order_number)
                return "no_changes" if order_number == "W000400" else "uploaded"

            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=orders),
                patch.object(artifacts, "download_with_recovery", return_value=Path("order.xls")),
                patch.object(artifacts.transform, "parse_order_xls", side_effect=[
                    ({"order_number": "000400"}, [{}]),
                    ({"order_number": "000401"}, [{}]),
                ]),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, p: p.write_bytes(b"xlsx")),
                patch.object(artifacts.posco_import, "stage_workbook", side_effect=record_stage),
            ):
                report = {"orders": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 21),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                )
            self.assertEqual(staged_in_order, ["W000400", "W000401"])
            self.assertEqual(
                [item["status"] for item in report["orders"]],
                ["already_in_posco", "staged_for_review"],
            )
            self.assertEqual(state["staged_orders"], {"W000400|10/20/2026", "W000401|10/21/2026"})
            self.assertEqual(posco_import.load_stage_state(root / "stage.json"), state)
            self.assertEqual(
                [item["review_image"] for item in report["orders"]],
                ["001_W000400_posco_review.png", "002_W000401_posco_review.png"],
            )

    def test_failed_stage_is_held_for_manual_review_instead_of_retried(self):
        order = {"order": "W000400", "date": "10/20/2026", "internal_id": "400"}
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=[order]),
                patch.object(artifacts, "download_with_recovery", return_value=Path("W000400.xls")),
                patch.object(artifacts.transform, "parse_order_xls", return_value=(
                    {"order_number": "000400"}, [{}],
                )),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, p: p.write_bytes(b"xlsx")),
                patch.object(artifacts.posco_import, "stage_workbook", side_effect=RuntimeError("Sin confirmación")),
            ):
                report = {"orders": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 20),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                )
            self.assertEqual(report["orders"][0]["status"], "error")
            self.assertEqual(state["needs_review"], {"W000400|10/20/2026"})
            self.assertEqual(posco_import.load_stage_state(root / "stage.json"), state)

    def test_all_mode_stops_after_an_uncertain_upload(self):
        orders = [
            {"order": "W000400", "date": "10/20/2026", "internal_id": "400"},
            {"order": "W000401", "date": "10/21/2026", "internal_id": "401"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": set(), "staged_orders": set(), "needs_review": set()}
            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=orders),
                patch.object(artifacts, "download_with_recovery", return_value=Path("order.xls")) as download,
                patch.object(artifacts.transform, "parse_order_xls", return_value=(
                    {"order_number": "000400"}, [{}],
                )),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, p: p.write_bytes(b"xlsx")),
                patch.object(artifacts.posco_import, "stage_workbook", side_effect=RuntimeError("Sin confirmación")) as stage,
            ):
                report = {"orders": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 21),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                )
            self.assertEqual(download.call_count, 1)
            stage.assert_called_once()
            self.assertEqual(len(report["orders"]), 1)
            self.assertEqual(state["needs_review"], {"W000400|10/20/2026"})

    def test_automatic_mode_skips_legacy_and_stages_pilot_order(self):
        orders = [
            {"order": "W000400", "date": "10/20/2026", "internal_id": "400"},
            {"order": "W000317", "date": "10/12/2026", "internal_id": "317"},
        ]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            state = {"legacy_orders": {"W000400|10/20/2026"}, "staged_orders": set(), "needs_review": set()}
            with (
                patch.object(artifacts, "ARTIFACTS_DIR", root),
                patch.object(artifacts, "LEDGER_PATH", root / "transformed.json"),
                patch.object(artifacts, "STAGE_LEDGER_PATH", root / "stage.json"),
                patch.object(artifacts, "collect_filtered_orders", return_value=orders),
                patch.object(artifacts, "download_with_recovery", return_value=Path("W000317.xls")) as download,
                patch.object(artifacts.transform, "parse_order_xls", return_value=(
                    {"order_number": "000317"}, [{}],
                )),
                patch.object(artifacts.transform, "build_order_rows", return_value=[["box"]]),
                patch.object(artifacts.transform, "write_order_workbook", side_effect=lambda _t, _r, p: p.write_bytes(b"xlsx")),
                patch.object(artifacts.posco_import, "stage_workbook") as stage,
            ):
                report = {"orders": []}
                artifacts.transform_visible_orders(
                    None, set(), set(), {}, b"template", report,
                    date(2026, 10, 9), date(2026, 10, 21),
                    stage_mode="all", stage_page=Mock(), stage_state=state,
                )
            self.assertEqual([item["status"] for item in report["orders"]], ["legacy", "staged_for_review"])
            self.assertEqual(download.call_count, 1)
            self.assertEqual(stage.call_args.args[2], "W000317")

    def test_pilot_mail_ledgers_remain_separate_from_stage_ledger(self):
        from bot import posco_order_transform as transform

        root = Path(__file__).parent / "data"
        transformed, ignored = transform.load_order_state(root / "posco_transformados.json")
        posco_import.load_stage_state(root / "posco_cargas.json")
        pilot = "W000317|10/12/2026"
        self.assertIn(pilot, transformed)
        self.assertNotIn(pilot, ignored)
        self.assertIn(pilot, (root / "facturas_enviadas.json").read_text(encoding="utf-8"))
        self.assertIn(pilot, (root / "detalles_exportacion_enviados.json").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
