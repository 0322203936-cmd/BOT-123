import base64
import html
import json
import mimetypes
import os
from pathlib import Path
from urllib.parse import quote

import requests

from sharepoint_sync import GRAPH_URL, graph_headers, graph_token


# Graph upload-session chunks must be below 4 MB and use a 320 KiB multiple.
UPLOAD_CHUNK_SIZE = 12 * 320 * 1024
SIMPLE_ATTACHMENT_LIMIT = 3 * 1024 * 1024
MAX_ATTACHMENT_SIZE = 150 * 1024 * 1024


class EmailConfigError(ValueError):
    """La configuración del correo no cumple el contrato del panel."""


def _clean_recipients(value: object, field: str) -> list[str]:
    if not isinstance(value, list):
        raise EmailConfigError(f"El campo {field} debe ser una lista de correos.")

    recipients: list[str] = []
    for raw_address in value:
        if not isinstance(raw_address, str):
            raise EmailConfigError(f"El campo {field} contiene un valor inválido.")
        address = raw_address.strip()
        if not address or "@" not in address or any(character.isspace() for character in address):
            raise EmailConfigError(f"El campo {field} contiene un correo inválido.")
        if address.lower() not in {item.lower() for item in recipients}:
            recipients.append(address)
    if len(recipients) > 100:
        raise EmailConfigError(f"El campo {field} no puede tener más de 100 correos.")
    return recipients


def load_email_config(raw_config: str | None) -> dict | None:
    if not raw_config or not raw_config.strip():
        return None
    try:
        value = json.loads(raw_config)
    except json.JSONDecodeError as exc:
        raise EmailConfigError("CAJAS_EMAIL_CONFIG no contiene un JSON válido.") from exc
    if not isinstance(value, dict):
        raise EmailConfigError("CAJAS_EMAIL_CONFIG debe ser un objeto JSON.")

    recipients = _clean_recipients(value.get("recipients"), "recipients")
    if not recipients:
        raise EmailConfigError("CAJAS_EMAIL_CONFIG debe tener al menos un destinatario.")
    cc = _clean_recipients(value.get("cc", []), "cc")
    bcc = _clean_recipients(value.get("bcc", []), "bcc")
    subject = value.get("subject")
    body_html = value.get("bodyHtml")
    logo_url = value.get("logoUrl", "")
    if not isinstance(subject, str) or not subject.strip() or len(subject.strip()) > 200:
        raise EmailConfigError("El asunto es obligatorio y debe tener hasta 200 caracteres.")
    if not isinstance(body_html, str) or not body_html.strip() or len(body_html) > 100_000:
        raise EmailConfigError("El contenido HTML es obligatorio y demasiado grande.")
    if not isinstance(logo_url, str):
        raise EmailConfigError("La URL del logo no es válida.")
    logo_url = logo_url.strip()
    if logo_url and not logo_url.lower().startswith(("https://", "http://")):
        raise EmailConfigError("La URL del logo debe comenzar con http:// o https://.")

    return {
        "recipients": recipients,
        "cc": cc,
        "bcc": bcc,
        "subject": subject.strip(),
        "bodyHtml": body_html,
        "logoUrl": logo_url,
    }


def _graph_recipients(addresses: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": address}} for address in addresses]


def build_graph_message(config: dict) -> dict:
    logo_url = config.get("logoUrl", "")
    logo_html = ""
    if logo_url:
        safe_url = html.escape(logo_url, quote=True)
        logo_html = (
            '<p style="margin:0 0 20px;text-align:left;">'
            f'<img src="{safe_url}" alt="Logo" style="max-width:240px;height:auto;">'
            "</p>"
        )

    return {
        "subject": config["subject"],
        "body": {
            "contentType": "HTML",
            "content": logo_html + config["bodyHtml"],
        },
        "toRecipients": _graph_recipients(config["recipients"]),
        "ccRecipients": _graph_recipients(config["cc"]),
        "bccRecipients": _graph_recipients(config["bcc"]),
    }


def _user_path(sender: str, suffix: str) -> str:
    return f"{GRAPH_URL}/users/{quote(sender, safe='')}/{suffix}"


def _raise_for_graph(response: requests.Response, action: str) -> None:
    if response.ok:
        return
    detail = response.text[:500].strip()
    raise RuntimeError(f"Graph no pudo {action} (HTTP {response.status_code}). {detail}")


def _create_draft(token: str, sender: str, message: dict) -> str:
    response = requests.post(
        _user_path(sender, "messages"),
        headers={**graph_headers(token), "Content-Type": "application/json"},
        json=message,
        timeout=30,
    )
    _raise_for_graph(response, "crear el borrador del correo")
    message_id = response.json().get("id")
    if not message_id:
        raise RuntimeError("Graph creó el borrador sin devolver su identificador.")
    return message_id


def _upload_small_attachment(token: str, sender: str, message_id: str, attachment_path: Path) -> None:
    content_type = mimetypes.guess_type(attachment_path.name)[0] or "application/octet-stream"
    payload = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": attachment_path.name,
        "contentType": content_type,
        "contentBytes": base64.b64encode(attachment_path.read_bytes()).decode("ascii"),
    }
    response = requests.post(
        _user_path(sender, f"messages/{quote(message_id, safe='')}/attachments"),
        headers={**graph_headers(token), "Content-Type": "application/json"},
        json=payload,
        timeout=120,
    )
    _raise_for_graph(response, "adjuntar el XLS")


def _upload_large_attachment(token: str, sender: str, message_id: str, attachment_path: Path) -> None:
    size = attachment_path.stat().st_size
    response = requests.post(
        _user_path(sender, f"messages/{quote(message_id, safe='')}/attachments/createUploadSession"),
        headers={**graph_headers(token), "Content-Type": "application/json"},
        json={
            "AttachmentItem": {
                "attachmentType": "file",
                "name": attachment_path.name,
                "size": size,
            }
        },
        timeout=30,
    )
    _raise_for_graph(response, "crear la sesión para el XLS grande")
    upload_url = response.json().get("uploadUrl")
    if not upload_url:
        raise RuntimeError("Graph no devolvió la URL de carga del XLS grande.")

    with attachment_path.open("rb") as stream:
        start = 0
        while start < size:
            chunk = stream.read(UPLOAD_CHUNK_SIZE)
            if not chunk:
                raise RuntimeError("El XLS cambió o quedó incompleto durante el envío.")
            end = start + len(chunk) - 1
            upload_response = requests.put(
                upload_url,
                headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{size}",
                    "Content-Type": "application/octet-stream",
                },
                data=chunk,
                timeout=180,
            )
            _raise_for_graph(upload_response, f"subir el bloque {end + 1} de {size} bytes")
            start = end + 1
            print(f"Adjunto de correo: {start}/{size} bytes enviados.", flush=True)


def send_report_email(
    config: dict,
    sender: str,
    attachment_path: Path,
    token: str | None = None,
) -> None:
    if not sender.strip():
        raise RuntimeError("Falta configurar el secreto MAIL_SENDER.")
    if not attachment_path.is_file():
        raise RuntimeError(f"No se encontró el XLS para adjuntar: {attachment_path}.")
    attachment_size = attachment_path.stat().st_size
    if attachment_size > MAX_ATTACHMENT_SIZE:
        raise RuntimeError("El XLS supera el límite de 150 MB permitido por Graph.")

    graph_access_token = token or graph_token()
    clean_sender = sender.strip()
    message_id = _create_draft(graph_access_token, clean_sender, build_graph_message(config))
    try:
        if attachment_size < SIMPLE_ATTACHMENT_LIMIT:
            _upload_small_attachment(graph_access_token, clean_sender, message_id, attachment_path)
        else:
            _upload_large_attachment(graph_access_token, clean_sender, message_id, attachment_path)
    except Exception:
        # El borrador no debe quedar guardado si falla la carga del adjunto.
        try:
            requests.delete(
                _user_path(clean_sender, f"messages/{quote(message_id, safe='')}"),
                headers=graph_headers(graph_access_token),
                timeout=30,
            )
        except Exception:
            pass
        raise

    try:
        response = requests.post(
            _user_path(clean_sender, f"messages/{quote(message_id, safe='')}/send"),
            headers=graph_headers(graph_access_token),
            timeout=30,
        )
        _raise_for_graph(response, "enviar el correo")
    except requests.RequestException as exc:
        raise RuntimeError(
            "Graph no confirmó el envío. El borrador se conserva para revisión; "
            "verifica Borradores y Elementos enviados antes de reejecutar."
        ) from exc
    print(f"Correo enviado con {attachment_path.name} a {len(config['recipients'])} destinatario(s).", flush=True)
