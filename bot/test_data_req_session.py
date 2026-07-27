import sys
import types
import unittest

requests_stub = types.ModuleType("requests")
requests_stub.Response = object
requests_stub.request = lambda *args, **kwargs: None
sys.modules.setdefault("requests", requests_stub)

playwright_stub = types.ModuleType("playwright")
playwright_sync_stub = types.ModuleType("playwright.sync_api")
playwright_sync_stub.TimeoutError = TimeoutError
playwright_sync_stub.sync_playwright = lambda: None
sys.modules.setdefault("playwright", playwright_stub)
sys.modules.setdefault("playwright.sync_api", playwright_sync_stub)

from data_req import (
    DATA_REQ_CHUNK_SIZE,
    ResilientWorkbookSession,
    excel_rows_equal,
)


class FakeResponse:
    def __init__(self, payload=None):
        self.payload = payload or {}

    def json(self):
        return self.payload


class DataReqSessionTests(unittest.TestCase):
    def test_uses_smaller_chunks(self):
        self.assertEqual(DATA_REQ_CHUNK_SIZE, 100)

    def test_excel_verification_normalizes_blank_cells(self):
        self.assertTrue(excel_rows_equal(["VERONICA", None], ["VERONICA", ""]))
        self.assertFalse(excel_rows_equal(["VERONICA", 8], ["VERONICA", 10]))

    def test_excel_verification_accepts_excel_date_serials(self):
        self.assertTrue(
            excel_rows_equal(
                ["2026-09-11T00:00:00", 8],
                [46276, 8.0],
            )
        )

    def test_excel_verification_rejects_a_different_date(self):
        self.assertFalse(
            excel_rows_equal(
                ["2026-09-11T00:00:00"],
                [46275],
            )
        )

    def test_renews_invalid_session_and_retries_pending_request(self):
        opened = 0
        patch_sessions = []

        def fake_request(method, url, headers, **kwargs):
            nonlocal opened
            if url.endswith("/createSession"):
                opened += 1
                return FakeResponse({"id": f"session-{opened}"})
            if url.endswith("/closeSession"):
                return FakeResponse()
            if method == "PATCH":
                session_id = headers["workbook-session-id"]
                patch_sessions.append(session_id)
                if session_id == "session-1":
                    raise RuntimeError(
                        "HTTP 400: The target session is invalid"
                    )
                return FakeResponse({"ok": True})
            raise AssertionError(f"Solicitud inesperada: {method} {url}")

        session = ResilientWorkbookSession(
            "https://graph.example/workbook",
            {"Authorization": "Bearer token"},
            request_func=fake_request,
        )

        response = session.request(
            "PATCH",
            "https://graph.example/workbook/worksheets/DataReq/range",
            json={"values": [[1]]},
        )

        self.assertEqual(response.json(), {"ok": True})
        self.assertEqual(patch_sessions, ["session-1", "session-2"])

    def test_does_not_hide_unrelated_errors(self):
        def fake_request(method, url, headers, **kwargs):
            if url.endswith("/createSession"):
                return FakeResponse({"id": "session-1"})
            raise RuntimeError("HTTP 403: acceso denegado")

        session = ResilientWorkbookSession(
            "https://graph.example/workbook",
            {"Authorization": "Bearer token"},
            request_func=fake_request,
        )

        with self.assertRaisesRegex(RuntimeError, "403"):
            session.request(
                "PATCH",
                "https://graph.example/workbook/worksheets/DataReq/range",
                json={"values": [[1]]},
            )


if __name__ == "__main__":
    unittest.main()
