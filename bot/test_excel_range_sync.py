import unittest
import sys
import types

requests_stub = types.ModuleType("requests")
requests_stub.Response = object
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)

openpyxl_stub = types.ModuleType("openpyxl")
openpyxl_stub.load_workbook = lambda *args, **kwargs: None
sys.modules.setdefault("openpyxl", openpyxl_stub)

sharepoint_stub = types.ModuleType("sharepoint_sync")
sharepoint_stub.GRAPH_URL = ""
sharepoint_stub.graph_headers = lambda token: {}
sys.modules.setdefault("sharepoint_sync", sharepoint_stub)


from excel_range_sync import calculated_formulas


class ExcelRangeSyncTests(unittest.TestCase):
    def test_calculated_columns_have_five_formulas(self) -> None:
        formulas = calculated_formulas(5932)
        self.assertEqual(len(formulas), 5)
        self.assertTrue(all(formula.startswith("=") for formula in formulas))

    def test_row_formula_uses_its_own_row(self) -> None:
        self.assertEqual(
            calculated_formulas(5932)[4],
            '=E5932 & " " & " X " & T5932',
        )

    def test_structured_formulas_match_table_columns(self) -> None:
        formulas = calculated_formulas(2)
        self.assertIn("[FLOR]", formulas[0])
        self.assertIn("[txr_orden]", formulas[1])
        self.assertIn("[Flor Color]", formulas[2])
        self.assertIn("[TxR2]", formulas[3])


if __name__ == "__main__":
    unittest.main()
