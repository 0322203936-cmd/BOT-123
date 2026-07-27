import sys
import types
import unittest

requests_stub = types.ModuleType("requests")
requests_stub.Response = object
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)

from data_proy_refresh import combine_like_power_query, refresh_weeks_pivot


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

    def test_refreshes_only_expected_pivot(self) -> None:
        calls = []

        def fake_graph_request(method, url, headers, **kwargs):
            calls.append((method, url, kwargs))
            if method == "GET":
                return FakeResponse({"name": "PivotTable6"})
            return FakeResponse()

        refresh_weeks_pivot(fake_graph_request, "https://graph/workbook", {})

        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[0][0], "GET")
        self.assertTrue(
            calls[0][1].endswith(
                "/worksheets/Weeks%20x%20FechaProduccion/"
                "pivotTables/PivotTable6"
            )
        )
        self.assertEqual(calls[1][0], "POST")
        self.assertTrue(calls[1][1].endswith("/PivotTable6/refresh"))


if __name__ == "__main__":
    unittest.main()
