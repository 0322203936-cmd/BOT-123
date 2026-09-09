import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from email_sender import (
    EmailConfigError,
    UPLOAD_CHUNK_SIZE,
    build_graph_message,
    load_email_config,
    send_report_email,
)


class EmailSenderTests(unittest.TestCase):
    def test_loads_and_normalizes_email_configuration(self):
        config = load_email_config(
            json.dumps(
                {
                    "recipients": [" first@example.com ", "second@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": " Reporte de cajas ",
                    "bodyHtml": "<p>Listo</p>",
                    "logoUrl": " https://example.com/logo.png ",
                }
            )
        )

        self.assertEqual(config["recipients"], ["first@example.com", "second@example.com"])
        self.assertEqual(config["subject"], "Reporte de cajas")
        self.assertEqual(config["logoUrl"], "https://example.com/logo.png")

    def test_requires_at_least_one_recipient(self):
        with self.assertRaises(EmailConfigError):
            load_email_config(json.dumps({"recipients": [], "subject": "Reporte", "bodyHtml": "<p>Hola</p>"}))

    def test_builds_graph_message_with_inline_logo(self):
        message = build_graph_message(
            {
                "recipients": ["destino@example.com"],
                "cc": ["copia@example.com"],
                "bcc": [],
                "subject": "Reporte de cajas",
                "bodyHtml": "<p>El archivo está listo.</p>",
                "logoUrl": "https://example.com/logo.png",
            }
        )

        self.assertEqual(message["subject"], "Reporte de cajas")
        self.assertEqual(message["toRecipients"][0]["emailAddress"]["address"], "destino@example.com")
        self.assertEqual(message["ccRecipients"][0]["emailAddress"]["address"], "copia@example.com")
        self.assertIn("https://example.com/logo.png", message["body"]["content"])
        self.assertIn("El archivo está listo.", message["body"]["content"])

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_uploads_large_report_in_graph_blocks(self, post, put, delete):
        self.assertLess(UPLOAD_CHUNK_SIZE, 4 * 1024 * 1024)
        self.assertEqual(UPLOAD_CHUNK_SIZE % (320 * 1024), 0)

        class Response:
            ok = True
            status_code = 200
            text = ""

            def __init__(self, payload=None):
                self.payload = payload or {}

            def json(self):
                return self.payload

        post.side_effect = [
            Response({"id": "draft-1"}),
            Response({"uploadUrl": "https://upload.example/session"}),
            Response(),
        ]
        put.return_value = Response()

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"x" * (UPLOAD_CHUNK_SIZE + 1))
            send_report_email(
                {
                    "recipients": ["destino@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": "Reporte",
                    "bodyHtml": "<p>Listo</p>",
                    "logoUrl": "",
                },
                "remitente@example.com",
                report,
                token="token",
            )

        self.assertEqual(put.call_count, 2)
        self.assertEqual(put.call_args_list[0].kwargs["headers"]["Content-Range"], f"bytes 0-{UPLOAD_CHUNK_SIZE - 1}/{UPLOAD_CHUNK_SIZE + 1}")
        self.assertEqual(put.call_args_list[1].kwargs["headers"]["Content-Range"], f"bytes {UPLOAD_CHUNK_SIZE}-{UPLOAD_CHUNK_SIZE}/{UPLOAD_CHUNK_SIZE + 1}")
        self.assertFalse(delete.called)


if __name__ == "__main__":
    unittest.main()
