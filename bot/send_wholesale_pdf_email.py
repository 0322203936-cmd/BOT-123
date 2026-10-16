"""Send a completed Wholesale PDF artifact in a separate GitHub workflow."""

from __future__ import annotations

import html
import json
import os
import re
from pathlib import Path

from email_sender import send_report_email


def required_value(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        raise RuntimeError(f"Falta configurar {name}.")
    return value


def find_pdf(directory: Path) -> Path:
    candidates = sorted(path for path in directory.rglob("*.pdf") if path.is_file())
    if len(candidates) != 1:
        raise RuntimeError(
            "Se esperaba exactamente un PDF del flujo Wholesale, "
            f"pero se encontraron {len(candidates)}."
        )
    return candidates[0]


def load_order_metadata(directory: Path) -> dict[str, str]:
    candidates = sorted(path for path in directory.rglob("summary.json") if path.is_file())
    if len(candidates) != 1:
        raise RuntimeError(
            "Se esperaba exactamente un resumen del PDF Wholesale, "
            f"pero se encontraron {len(candidates)}."
        )
    try:
        value = json.loads(candidates[0].read_text(encoding="utf-8"))
        metadata = value["order_email"]
        customer = str(metadata["customer"]).strip()
        load_date = str(metadata["load_date"]).strip()
        document_kind = str(metadata["document_kind"]).strip()
    except (KeyError, TypeError, ValueError, json.JSONDecodeError) as error:
        raise RuntimeError("El resumen del PDF no contiene los datos del cliente.") from error
    if not customer or not load_date or not document_kind:
        raise RuntimeError("El resumen del PDF tiene datos incompletos del cliente.")
    return {
        "customer": customer,
        "load_date": load_date,
        "document_kind": document_kind,
    }


def safe_attachment_name(document_kind: str, customer: str) -> str:
    clean_customer = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', " ", customer)
    clean_customer = re.sub(r"\s+", " ", clean_customer).strip(" .")
    return f"{document_kind} Wholesale - {clean_customer or 'Cliente'}.pdf"


def main() -> None:
    artifact_directory = Path(required_value("WHOLESALE_PDF_ARTIFACT_DIR"))
    recipient = required_value("WHOLESALE_PDF_RECIPIENT")
    sender = required_value("MAIL_SENDER")
    pdf_path = find_pdf(artifact_directory)
    metadata = load_order_metadata(artifact_directory)
    document_kind = metadata["document_kind"]
    customer = metadata["customer"]
    load_date = metadata["load_date"]
    attachment_path = pdf_path.with_name(safe_attachment_name(document_kind, customer))
    pdf_path.replace(attachment_path)
    safe_load_date = html.escape(load_date)
    if document_kind == "Orden":
        subject = f"Orden Wholesale - {customer}"
        message = f"<p>Adjunto Orden generada - Load Date: {safe_load_date}.</p>"
    else:
        subject = f"Passing Sheet Wholesale - {customer}"
        message = f"<p>Adjunto Passing Sheet generado - Load Date: {safe_load_date}.</p>"

    config = {
        "recipients": [recipient],
        "cc": [],
        "bcc": [],
        "subject": subject,
        "bodyHtml": "<p>Hola Jesús,</p>" + message + "<p>Saludos.</p>",
        "logoUrl": "",
        "logoData": "",
        "logoName": "",
        "logoSharePoint": None,
        "pdfSharePoint": None,
    }
    send_report_email(config, sender, attachment_path)
    print(f"PDF Wholesale enviado a {recipient}: {attachment_path.name}", flush=True)


if __name__ == "__main__":
    main()
