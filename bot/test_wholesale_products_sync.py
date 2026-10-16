import unittest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from wholesale_products_sync import deduplicate_against_products, make_product_row


class WholesaleProductsSyncTests(unittest.TestCase):
    def test_new_description_is_appended_with_fixed_identity_and_dimension(self):
        product = {
            "categoria": "FLOWER",
            "flor": "Rose",
            "flor_color": "Red",
            "descripcion": "Rose Red 60 cm",
            "ramos_cajas": 10,
            "txr_orden": 10,
            "caja": "F4P",
        }
        rows = deduplicate_against_products(
            [product],
            set(),
            {"f4p": "40x8x8"},
        )
        self.assertEqual(rows, [[
            "PACIFICA PRODUCE FARMS", "WHOLESALE", "FLOWER", "Rose", "Red",
            "Rose Red 60 cm", 10, 10, "F4P", "", "F4P", "40x8x8",
        ]])

    def test_existing_and_repeated_descriptions_are_not_appended(self):
        product = {
            "descripcion": "  Rose   Red 60 cm ",
            "caja": "F4P",
        }
        rows = deduplicate_against_products(
            [product, product],
            {"rose red 60 cm"},
            {},
        )
        self.assertEqual(rows, [])

    def test_missing_dimension_stays_blank(self):
        row = make_product_row({"descripcion": "New Product", "caja": "NEW"}, {})
        self.assertEqual(row[10:], ["NEW", ""])


if __name__ == "__main__":
    unittest.main()
