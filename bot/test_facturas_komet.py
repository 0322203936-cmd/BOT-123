import json
import tempfile
import unittest
from pathlib import Path

from bot.facturas_komet import load_sent_order_keys, order_key, save_sent_order_keys


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


if __name__ == "__main__":
    unittest.main()
