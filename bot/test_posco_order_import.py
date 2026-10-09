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
            posco_import.stage_workbook(page, workbook, "W000400", Path(directory) / "review.png")
            page.locator.assert_any_call('input[type="file"][name="file"]')
            page.locator.assert_any_call('select[name="upload_mode"]')
            page.locator('input[type="file"][name="file"]').set_input_files.assert_called_with(str(workbook))
            page.locator('select[name="upload_mode"]').select_option.assert_called_with("pacifica2")
            page.get_by_role.assert_any_call("button", name="Upload", exact=True)
            page.get_by_role("heading", name="Import excel").wait_for.assert_called_with(state="hidden", timeout=600_000)
            page.get_by_text("No elements found", exact=True).wait_for.assert_called_with(state="hidden", timeout=600_000)
            page.screenshot.assert_called_once_with(path=str(Path(directory) / "review.png"), full_page=True)
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
                ["staged_for_review", "staged_for_review"],
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

    def test_pilot_order_is_removed_only_from_posco_ledgers(self):
        from bot import posco_order_transform as transform

        root = Path(__file__).parent / "data"
        transformed, ignored = transform.load_order_state(root / "posco_transformados.json")
        stage_state = posco_import.load_stage_state(root / "posco_cargas.json")
        pilot = "W000317|10/12/2026"
        self.assertNotIn(pilot, transformed | ignored | set().union(*stage_state.values()))
        self.assertIn(pilot, (root / "facturas_enviadas.json").read_text(encoding="utf-8"))
        self.assertIn(pilot, (root / "detalles_exportacion_enviados.json").read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
