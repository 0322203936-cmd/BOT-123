import sys
import types
import unittest
from unittest.mock import patch

requests_stub = types.ModuleType("requests")
requests_stub.Response = object
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)

from data_proy_refresh import (
    combine_like_power_query,
    excel_rows_equal,
    refresh_weeks_pivot,
)
from data_proy import clear_and_write


class FakeResponse:
    def __init__(self, payload=None) -> None:
        self.payload = payload or {}

    def json(self):
        return self.payload


class DataProyRefreshTests(unittest.TestCase):
    def test_combines_tables_in_power_query_order(self) -> None:
        proy = [
            ["FLOR", "Semana", "Extra"],
            ["ASTER", 31, "P"],
            ["ROSA", 32, "P2"],
        ]
        req = [
            ["FLOR", "Semana"],
            ["VERONICA", 31],
        ]
        headers = ["FLOR", "Semana", "Extra"]

        self.assertEqual(
            combine_like_power_query(proy, req, headers),
            [
                ["ASTER", 31, "P"],
                ["ROSA", 32, "P2"],
                ["VERONICA", 31, None],
            ],
        )

    def test_rejects_unknown_source_columns(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "no reconoce"):
            combine_like_power_query(
                [["FLOR", "Columna inesperada"], ["ASTER", 1]],
                [["FLOR"], ["ROSA"]],
                ["FLOR"],
            )

    def test_rejects_duplicate_target_headers(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "duplicados"):
            combine_like_power_query(
                [["FLOR"], ["ASTER"]],
                [["FLOR"], ["ROSA"]],
                ["FLOR", "FLOR"],
            )

    def test_excel_verification_treats_none_and_blank_as_equal(self) -> None:
        self.assertTrue(
            excel_rows_equal(
                ["VERONICA", 31, None],
                ["VERONICA", 31, ""],
            )
        )

    def test_excel_verification_still_detects_real_differences(self) -> None:
        self.assertFalse(
            excel_rows_equal(
                ["VERONICA", 31, None],
                ["VERONICA", 32, ""],
            )
        )

    def test_refreshes_all_pivots_on_expected_sheet_and_recalculates(self) -> None:
        calls = []

        def fake_graph_request(method, url, headers, **kwargs):
            calls.append((method, url, kwargs))
            if method == "GET":
                return FakeResponse({"name": "PivotTable6"})
            return FakeResponse()

        refresh_weeks_pivot(fake_graph_request, "https://graph/workbook", {})

        self.assertEqual(len(calls), 3)
        self.assertEqual(calls[0][0], "GET")
        self.assertTrue(
            calls[0][1].endswith(
                "/worksheets/Weeks%20x%20FechaProduccion/"
                "pivotTables/PivotTable6"
            )
        )
        self.assertEqual(calls[1][0], "POST")
        self.assertTrue(calls[1][1].endswith("/pivotTables/refreshAll"))
        self.assertEqual(calls[2][0], "POST")
        self.assertTrue(calls[2][1].endswith("/application/calculate"))
        self.assertEqual(
            calls[2][2]["json"],
            {"calculationType": "Full"},
        )

    def test_data_proy_writes_firme_in_column_u(self) -> None:
        with patch("data_proy.graph_request") as request:
            clear_and_write(
                "https://graph/workbook",
                {},
                2,
                3,
                2,
                [["Proyeccion", "CORTE", "CORTE", "CORTE"]] * 2,
                [["ASTER", "WHITE", "CORTE"]] * 2,
                [[100], [200]],
                [[31], [31]],
                [["FIRME"], ["FIRME"]],
            )

        urls_and_values = [
            (call.args[1], call.kwargs["json"]["values"])
            for call in request.call_args_list
        ]
        self.assertIn(
            (
                "https://graph/workbook/worksheets/DataProy/"
                "range(address='U2:U3')",
                [["FIRME"], ["FIRME"]],
            ),
            urls_and_values,
        )


if __name__ == "__main__":
    unittest.main()
