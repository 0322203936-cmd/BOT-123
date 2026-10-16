"""Send a completed Wholesale PDF artifact in a separate GitHub workflow."""

from __future__ import annotations

import os
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


def main() -> None:
    artifact_directory = Path(required_value("WHOLESALE_PDF_ARTIFACT_DIR"))
    recipient = required_value("WHOLESALE_PDF_RECIPIENT")
    sender = required_value("MAIL_SENDER")
    label = required_value("WHOLESALE_PDF_LABEL")
    pdf_path = find_pdf(artifact_directory)

    config = {
        "recipients": [recipient],
        "cc": [],
        "bcc": [],
        "subject": f"Wholesale - {label}",
        "bodyHtml": (
            "<p>Hola Jesús,</p>"
            f"<p>Adjunto encontrarás el PDF generado por <strong>{label}</strong>.</p>"
            "<p>Saludos.</p>"
        ),
        "logoUrl": "",
        "logoData": "",
        "logoName": "",
        "logoSharePoint": None,
        "pdfSharePoint": None,
    }
    send_report_email(config, sender, pdf_path)
    print(f"PDF Wholesale enviado a {recipient}: {pdf_path.name}", flush=True)


if __name__ == "__main__":
    main()
