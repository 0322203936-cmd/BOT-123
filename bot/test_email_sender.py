import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from email_sender import (
    EmailConfigError,
    build_direct_send_payload,
    build_graph_message,
    load_email_config,
    send_report_email,
)


class EmailSenderTests(unittest.TestCase):
    def test_includes_sharepoint_pdf_with_report_attachment(self):
        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            pdf = Path(directory) / "guide.pdf"
            report.write_bytes(b"report")
            pdf.write_bytes(b"pdf")
            payload = build_direct_send_payload(
                {
                    "recipients": ["destino@example.com"], "cc": [], "bcc": [],
                    "subject": "Reporte", "bodyHtml": "<p>Listo</p>",
                    "logoData": "", "logoSharePoint": None, "logoUrl": "",
                    "pdfSharePointPath": pdf,
                },
                report,
                None,
            )

        self.assertEqual([item["name"] for item in payload["message"]["attachments"]], ["inventory.xlsx", "guide.pdf"])

    def test_loads_and_normalizes_inline_logo_configuration(self):
        logo_data = f"data:image/png;base64,{base64.b64encode(b'logo').decode('ascii')}"
        config = load_email_config(
            json.dumps(
                {
                    "recipients": ["destino@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": "Reporte de cajas",
                    "bodyHtml": "<p>Listo</p>",
                    "logoData": logo_data,
                    "logoName": "logo.png",
                    "logoContentType": "image/png",
                }
            )
        )

        self.assertEqual(config["logoData"], logo_data)
        self.assertEqual(config["logoName"], "logo.png")
        self.assertEqual(config["logoContentType"], "image/png")

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

    def test_builds_graph_message_with_embedded_logo_reference(self):
        message = build_graph_message(
            {
                "recipients": ["destino@example.com"],
                "cc": [],
                "bcc": [],
                "subject": "Reporte de cajas",
                "bodyHtml": "<p>El archivo está listo.</p>",
                "logoData": "data:image/png;base64,bG9nbw==",
                "logoName": "logo.png",
                "logoContentType": "image/png",
            }
        )

        self.assertIn('src="cid:cajas-logo"', message["body"]["content"])
        self.assertNotIn("data:image/png", message["body"]["content"])
        self.assertTrue(message["body"]["content"].rstrip().endswith("</p>"))
        self.assertGreater(
            message["body"]["content"].rfind('src="cid:cajas-logo"'),
            message["body"]["content"].find("El archivo está listo."),
        )

    def test_loads_sharepoint_logo_reference(self):
        config = load_email_config(
            json.dumps(
                {
                    "recipients": ["destino@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": "Reporte",
                    "bodyHtml": "<p>Listo</p>",
                    "logoSharePoint": {
                        "driveId": "drive-1",
                        "itemId": "item-1",
                        "name": "cajas-email-logo.png",
                        "contentType": "image/png",
                    },
                }
            )
        )

        self.assertEqual(config["logoSharePoint"]["driveId"], "drive-1")
        self.assertEqual(config["logoSharePoint"]["itemId"], "item-1")

    def test_builds_graph_message_with_sharepoint_logo_reference(self):
        message = build_graph_message(
            {
                "recipients": ["destino@example.com"],
                "cc": [],
                "bcc": [],
                "subject": "Reporte de cajas",
                "bodyHtml": "<p>Listo</p>",
                "logoSharePoint": {
                    "driveId": "drive-1",
                    "itemId": "item-1",
                    "name": "cajas-email-logo.png",
                    "contentType": "image/png",
                },
            }
        )

        self.assertIn('src="cid:cajas-logo"', message["body"]["content"])
        self.assertGreater(
            message["body"]["content"].rfind('src="cid:cajas-logo"'),
            message["body"]["content"].find("Listo"),
        )

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_sends_one_individual_message_per_recipient(self, post, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"report")
            send_report_email(
                {
                    "recipients": ["uno@example.com", "dos@example.com", "tres@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": "Reporte",
                    "bodyHtml": "<p>Listo</p>",
                },
                "remitente@example.com",
                report,
                token="token",
            )

        self.assertEqual(post.call_count, 3)
        sent_to = [
            call.kwargs["json"]["message"]["toRecipients"][0]["emailAddress"]["address"]
            for call in post.call_args_list
        ]
        self.assertEqual(sent_to, ["uno@example.com", "dos@example.com", "tres@example.com"])
        for call in post.call_args_list:
            self.assertEqual(len(call.kwargs["json"]["message"]["toRecipients"]), 1)
            self.assertEqual(call.kwargs["json"]["message"]["ccRecipients"], [])
            self.assertEqual(call.kwargs["json"]["message"]["bccRecipients"], [])

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_sends_cc_only_on_the_first_individual_message(self, post, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"report")
            send_report_email(
                {
                    "recipients": ["uno@example.com", "dos@example.com", "tres@example.com"],
                    "cc": ["roberto@example.com"],
                    "bcc": [],
                    "subject": "Reporte",
                    "bodyHtml": "<p>Listo</p>",
                },
                "remitente@example.com",
                report,
                token="token",
            )

        self.assertEqual(post.call_count, 3)
        self.assertEqual(
            post.call_args_list[0].kwargs["json"]["message"]["ccRecipients"],
            [{"emailAddress": {"address": "roberto@example.com"}}],
        )
        self.assertEqual(post.call_args_list[1].kwargs["json"]["message"]["ccRecipients"], [])
        self.assertEqual(post.call_args_list[2].kwargs["json"]["message"]["ccRecipients"], [])

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_sends_logo_and_report_directly_without_creating_a_draft(self, post, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""

            def __init__(self, payload=None):
                self.payload = payload or {}

            def json(self):
                return self.payload

        post.side_effect = [Response({"id": "draft-1"}), Response(), Response(), Response()]

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"report")
            send_report_email(
                {
                    "recipients": ["destino@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": "Reporte",
                    "bodyHtml": "<p>Listo</p>",
                    "logoData": "data:image/png;base64,bG9nbw==",
                    "logoName": "logo.png",
                    "logoContentType": "image/png",
                },
                "remitente@example.com",
                report,
                token="token",
            )

        self.assertEqual(post.call_count, 1)
        self.assertTrue(post.call_args.args[0].endswith("/sendMail"))
        attachments = post.call_args.kwargs["json"]["message"]["attachments"]
        self.assertEqual(len(attachments), 2)
        inline_payload = attachments[0]
        self.assertTrue(inline_payload["isInline"])
        self.assertEqual(inline_payload["contentId"], "cajas-logo")
        self.assertEqual(inline_payload["contentBytes"], "bG9nbw==")
        self.assertFalse(put.called)
        self.assertFalse(delete.called)

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.get", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_sends_downloaded_sharepoint_logo_inline_in_direct_message(self, post, get, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""
            content = b"logo"

            def __init__(self, payload=None):
                self.payload = payload or {}

            def json(self):
                return self.payload

            def iter_content(self, chunk_size=None):
                yield self.content

            def close(self):
                pass

        post.side_effect = [Response({"id": "draft-1"}), Response(), Response(), Response()]
        get.return_value = Response()

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"report")
            send_report_email(
                {
                    "recipients": ["destino@example.com"],
                    "cc": [],
                    "bcc": [],
                    "subject": "Reporte",
                    "bodyHtml": "<p>Listo</p>",
                    "logoSharePoint": {
                        "driveId": "drive-1",
                        "itemId": "item-1",
                        "name": "cajas-email-logo.png",
                        "contentType": "image/png",
                    },
                },
                "remitente@example.com",
                report,
                token="token",
            )

        self.assertEqual(get.call_count, 1)
        self.assertEqual(post.call_count, 1)
        self.assertTrue(post.call_args.args[0].endswith("/sendMail"))
        attachments = post.call_args.kwargs["json"]["message"]["attachments"]
        self.assertEqual(len(attachments), 2)
        inline_payload = attachments[0]
        self.assertTrue(inline_payload["isInline"])
        self.assertEqual(inline_payload["contentId"], "cajas-logo")
        self.assertEqual(inline_payload["contentBytes"], "bG9nbw==")
        self.assertFalse(put.called)
        self.assertFalse(delete.called)

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.get", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_rejects_sharepoint_logo_over_10_mb_without_creating_a_draft(self, post, get, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""

            def iter_content(self, chunk_size=None):
                yield b"x" * (10 * 1024 * 1024)
                yield b"x"

            def close(self):
                pass

        get.return_value = Response()

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"report")
            with self.assertRaises(EmailConfigError):
                send_report_email(
                    {
                        "recipients": ["destino@example.com"],
                        "cc": [],
                        "bcc": [],
                        "subject": "Reporte",
                        "bodyHtml": "<p>Listo</p>",
                        "logoSharePoint": {
                            "driveId": "drive-1",
                            "itemId": "item-1",
                            "name": "cajas-email-logo.png",
                            "contentType": "image/png",
                        },
                    },
                    "remitente@example.com",
                    report,
                    token="token",
                )

        self.assertFalse(post.called)
        self.assertFalse(put.called)
        self.assertFalse(delete.called)

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.get", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_rejects_sharepoint_logo_too_large_for_direct_send(self, post, get, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""

            def __init__(self, payload=None):
                self.payload = payload or {}

            def json(self):
                return self.payload

            def iter_content(self, chunk_size=None):
                yield b"x" * (3 * 1024 * 1024 + 1)

            def close(self):
                pass

        post.return_value = Response()
        get.return_value = Response()
        put.return_value = Response()

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"report")
            with self.assertRaises(EmailConfigError):
                send_report_email(
                    {
                        "recipients": ["destino@example.com"],
                        "cc": [],
                        "bcc": [],
                        "subject": "Reporte",
                        "bodyHtml": "<p>Listo</p>",
                        "logoSharePoint": {
                            "driveId": "drive-1",
                            "itemId": "item-1",
                            "name": "cajas-email-logo.png",
                            "contentType": "image/png",
                        },
                    },
                    "remitente@example.com",
                    report,
                    token="token",
                )

        self.assertFalse(post.called)
        self.assertFalse(put.called)
        self.assertFalse(delete.called)

    @patch("email_sender.requests.delete", create=True)
    @patch("email_sender.requests.put", create=True)
    @patch("email_sender.requests.post", create=True)
    def test_rejects_report_too_large_for_direct_send(self, post, put, delete):
        class Response:
            ok = True
            status_code = 200
            text = ""

            def __init__(self, payload=None):
                self.payload = payload or {}

            def json(self):
                return self.payload

        post.return_value = Response()
        put.return_value = Response()

        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "inventory.xlsx"
            report.write_bytes(b"x" * (3 * 1024 * 1024))
            with self.assertRaises(EmailConfigError):
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

        self.assertFalse(post.called)
        self.assertFalse(put.called)
        self.assertFalse(delete.called)


if __name__ == "__main__":
    unittest.main()
