import unittest
from unittest.mock import patch

from sharepoint_sync import read_calculated_worksheet_values


class FakeResponse:
    def __init__(self, payload=None, status_code=200):
        self._payload = payload or {}
        self.status_code = status_code
        self.ok = status_code < 400
        self.headers = {}
        self.text = ""

    def json(self):
        return self._payload


class SharePointSyncTests(unittest.TestCase):
    @patch("sharepoint_sync.requests.post")
    @patch("sharepoint_sync.requests.request")
    def test_reads_excel_online_values_by_cell_address(self, request, post):
        request.side_effect = [
            FakeResponse({"id": "session-1"}),
            FakeResponse(),
            FakeResponse(
                {
                    "address": "'Customer View'!$B$8:$C$9",
                    "values": [["2026-10-07T00:00:00", 4], ["Rose", 5]],
                }
            ),
        ]

        values = read_calculated_worksheet_values(
            "token",
            {"parentReference": {"driveId": "drive-1"}, "id": "item-1"},
            "Customer View",
        )

        self.assertEqual(values["B8"].year, 2026)
        self.assertEqual(values["C8"], 4)
        self.assertEqual(values["B9"], "Rose")
        self.assertEqual(values["C9"], 5)
        self.assertEqual(post.call_count, 1)
        self.assertIn("closeSession", post.call_args.args[0])


if __name__ == "__main__":
    unittest.main()
