import base64
import html
import json
import mimetypes
import os
import re
import tempfile
from pathlib import Path
from urllib.parse import quote

import requests

from sharepoint_sync import GRAPH_URL, graph_headers, graph_token


# Graph upload-session chunks must be below 4 MB and use a 320 KiB multiple.
UPLOAD_CHUNK_SIZE = 12 * 320 * 1024
SIMPLE_ATTACHMENT_LIMIT = 3 * 1024 * 1024
MAX_ATTACHMENT_SIZE = 150 * 1024 * 1024
MAX_INLINE_LOGO_SIZE = 24 * 1024
MAX_SHAREPOINT_LOGO_SIZE = 10 * 1024 * 1024
DIRECT_SEND_PAYLOAD_LIMIT = 3 * 1024 * 1024
INLINE_LOGO_CONTENT_ID = "cajas-logo"


class EmailConfigError(ValueError):
    """La configuración del correo no cumple el contrato del panel."""


def _parse_inline_logo(value: object) -> tuple[str, bytes, str]:
    if value in (None, ""):
        return "", b"", ""
    if not isinstance(value, str):
        raise EmailConfigError("Los datos del logo no son válidos.")

    match = re.fullmatch(
        r"data:(image/(?:png|jpeg|gif|bmp));base64,([A-Za-z0-9+/]*={0,2})",
        value.strip(),
        re.IGNORECASE,
    )
    if not match:
        raise EmailConfigError("El logo debe ser una imagen PNG, JPG, GIF o BMP.")

    try:
        content = base64.b64decode(match.group(2), validate=True)
    except ValueError as exc:
        raise EmailConfigError("Los datos del logo no son válidos.") from exc
    if not content or len(content) > MAX_INLINE_LOGO_SIZE:
        raise EmailConfigError("El logo debe pesar como máximo 24 KB.")
    return value.strip(), content, match.group(1).lower()


def _clean_sharepoint_logo(value: object) -> dict | None:
    if value in (None, ""):
        return None
    if not isinstance(value, dict):
        raise EmailConfigError("La referencia del logo en SharePoint no es válida.")
    drive_id = value.get("driveId")
    item_id = value.get("itemId")
    name = value.get("name")
    content_type = value.get("contentType")
    if not all(isinstance(item, str) and item.strip() for item in (drive_id, item_id, name, content_type)):
        raise EmailConfigError("La referencia del logo en SharePoint está incompleta.")
    drive_id = drive_id.strip()
    item_id = item_id.strip()
    name = name.strip()
    content_type = content_type.strip().lower()
    if content_type not in {"image/png", "image/jpeg", "image/gif", "image/bmp"}:
        raise EmailConfigError("El formato del logo no está permitido.")
    if len(name) > 100 or any(character in name for character in ("/", "\\", "\r", "\n")):
        raise EmailConfigError("El nombre del logo no es válido.")
    return {
        "driveId": drive_id,
        "itemId": item_id,
        "name": name,
        "contentType": content_type,
    }


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
    logo_data, _logo_bytes, logo_content_type = _parse_inline_logo(value.get("logoData", ""))
    logo_name = value.get("logoName", "")
    logo_sharepoint = _clean_sharepoint_logo(value.get("logoSharePoint"))
    if not isinstance(subject, str) or not subject.strip() or len(subject.strip()) > 200:
        raise EmailConfigError("El asunto es obligatorio y debe tener hasta 200 caracteres.")
    if not isinstance(body_html, str) or not body_html.strip() or len(body_html) > 100_000:
        raise EmailConfigError("El contenido HTML es obligatorio y demasiado grande.")
    if not isinstance(logo_url, str):
        raise EmailConfigError("La URL del logo no es válida.")
    logo_url = logo_url.strip()
    if logo_url and not logo_url.lower().startswith(("https://", "http://")):
        raise EmailConfigError("La URL del logo debe comenzar con http:// o https://.")
    if not isinstance(logo_name, str) or len(logo_name.strip()) > 100:
        raise EmailConfigError("El nombre del logo no es válido.")
    logo_name = logo_name.strip()
    if any(character in logo_name for character in ("/", "\\", "\r", "\n")):
        raise EmailConfigError("El nombre del logo no es válido.")
    if logo_data and not logo_name:
        logo_name = {
            "image/png": "logo.png",
            "image/jpeg": "logo.jpg",
            "image/gif": "logo.gif",
            "image/bmp": "logo.bmp",
        }[logo_content_type]
    if logo_data and not logo_sharepoint:
        logo_url = ""
    if logo_sharepoint:
        logo_url = ""
        logo_data = ""
        logo_name = ""
        logo_content_type = ""

    return {
        "recipients": recipients,
        "cc": cc,
        "bcc": bcc,
        "subject": subject.strip(),
        "bodyHtml": body_html,
        "logoUrl": logo_url,
        "logoData": logo_data,
        "logoName": logo_name,
        "logoContentType": logo_content_type,
        "logoSharePoint": logo_sharepoint,
    }


def _graph_recipients(addresses: list[str]) -> list[dict]:
    return [{"emailAddress": {"address": address}} for address in addresses]


def build_graph_message(config: dict) -> dict:
    logo_data = config.get("logoData", "")
    logo_sharepoint = config.get("logoSharePoint")
    logo_url = config.get("logoUrl", "")
    logo_html = ""
    if logo_data or logo_sharepoint:
        logo_html = (
            '<p style="margin:0 0 20px;text-align:left;">'
            f'<img src="cid:{INLINE_LOGO_CONTENT_ID}" alt="Logo" style="max-width:240px;height:auto;">'
            "</p>"
        )
    elif logo_url:
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
            "content": config["bodyHtml"] + logo_html,
        },
        "toRecipients": _graph_recipients(config["recipients"]),
        "ccRecipients": _graph_recipients(config["cc"]),
        "bccRecipients": _graph_recipients(config["bcc"]),
    }


def build_inline_logo_attachment(config: dict) -> dict | None:
    logo_data = config.get("logoData", "")
    if not logo_data:
        return None
    _, content, content_type = _parse_inline_logo(logo_data)
    extension = {
        "image/png": ".png",
        "image/jpeg": ".jpg",
        "image/gif": ".gif",
        "image/bmp": ".bmp",
    }[content_type]
    name = str(config.get("logoName") or f"logo{extension}").strip()
    if any(character in name for character in ("/", "\\", "\r", "\n")):
        raise EmailConfigError("El nombre del logo no es válido.")
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentType": content_type,
        "contentBytes": base64.b64encode(content).decode("ascii"),
        "contentId": INLINE_LOGO_CONTENT_ID,
        "isInline": True,
    }


def build_file_attachment(
    attachment_path: Path,
    *,
    content_id: str | None = None,
    content_type: str | None = None,
) -> dict:
    payload = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": attachment_path.name,
        "contentType": content_type
        or mimetypes.guess_type(attachment_path.name)[0]
        or "application/octet-stream",
        "contentBytes": base64.b64encode(attachment_path.read_bytes()).decode("ascii"),
    }
    if content_id:
        payload["contentId"] = content_id
        payload["isInline"] = True
    return payload


def build_direct_send_payload(
    config: dict,
    attachment_path: Path,
    sharepoint_logo_path: Path | None,
) -> dict:
    message = build_graph_message(config)
    attachments: list[dict] = []
    if sharepoint_logo_path:
        attachments.append(
            build_file_attachment(
                sharepoint_logo_path,
                content_id=INLINE_LOGO_CONTENT_ID,
                content_type=config["logoSharePoint"]["contentType"],
            )
        )
    else:
        inline_logo = build_inline_logo_attachment(config)
        if inline_logo:
            attachments.append(inline_logo)
    attachments.append(build_file_attachment(attachment_path))
    message["attachments"] = attachments
    payload = {"message": message, "saveToSentItems": True}
    payload_size = len(json.dumps(payload, ensure_ascii=True).encode("utf-8"))
    if payload_size > DIRECT_SEND_PAYLOAD_LIMIT:
        raise EmailConfigError(
            "El correo con el Excel y el logo supera el límite para enviarlo solo con "
            "Mail.Send. Reduce el tamaño del logo o habilita Mail.ReadWrite."
        )
    return payload


def _download_sharepoint_logo(token: str, config: dict) -> Path | None:
    logo = config.get("logoSharePoint")
    if not logo:
        return None
    drive_id = quote(str(logo["driveId"]), safe="")
    item_id = quote(str(logo["itemId"]), safe="")
    response = requests.get(
        f"{GRAPH_URL}/drives/{drive_id}/items/{item_id}/content",
        headers=graph_headers(token),
        timeout=120,
        stream=True,
    )
    _raise_for_graph(response, "descargar el logo desde SharePoint")
    suffix = Path(str(logo["name"])).suffix or ".img"
    temporary = tempfile.NamedTemporaryFile(prefix="cajas-logo-", suffix=suffix, delete=False)
    temporary_path = Path(temporary.name)
    total = 0
    try:
        for chunk in response.iter_content(chunk_size=64 * 1024):
            if not chunk:
                continue
            total += len(chunk)
            if total > MAX_SHAREPOINT_LOGO_SIZE:
                raise EmailConfigError("El logo de SharePoint debe pesar como máximo 10 MB.")
            temporary.write(chunk)
        if not total:
            raise EmailConfigError("El logo de SharePoint está vacío.")
    except Exception:
        temporary.close()
        close_response = getattr(response, "close", None)
        if callable(close_response):
            close_response()
        temporary_path.unlink(missing_ok=True)
        raise
    else:
        temporary.close()
        close_response = getattr(response, "close", None)
        if callable(close_response):
            close_response()
    return temporary_path


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


def _upload_small_attachment(
    token: str,
    sender: str,
    message_id: str,
    attachment_path: Path,
    *,
    content_id: str | None = None,
    content_type: str | None = None,
) -> None:
    content_type = content_type or mimetypes.guess_type(attachment_path.name)[0] or "application/octet-stream"
    payload = {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": attachment_path.name,
        "contentType": content_type,
        "contentBytes": base64.b64encode(attachment_path.read_bytes()).decode("ascii"),
    }
    if content_id:
        payload["contentId"] = content_id
        payload["isInline"] = True
    response = requests.post(
        _user_path(sender, f"messages/{quote(message_id, safe='')}/attachments"),
        headers={**graph_headers(token), "Content-Type": "application/json"},
        json=payload,
        timeout=120,
    )
    _raise_for_graph(response, "adjuntar el XLS")


def _upload_inline_logo(token: str, sender: str, message_id: str, config: dict) -> None:
    payload = build_inline_logo_attachment(config)
    if payload is None:
        return
    response = requests.post(
        _user_path(sender, f"messages/{quote(message_id, safe='')}/attachments"),
        headers={**graph_headers(token), "Content-Type": "application/json"},
        json=payload,
        timeout=120,
    )
    _raise_for_graph(response, "adjuntar el logo")


def _upload_inline_logo_file(
    token: str,
    sender: str,
    message_id: str,
    logo_path: Path,
    content_type: str,
) -> None:
    if logo_path.stat().st_size < SIMPLE_ATTACHMENT_LIMIT:
        _upload_small_attachment(
            token,
            sender,
            message_id,
            logo_path,
            content_id=INLINE_LOGO_CONTENT_ID,
            content_type=content_type,
        )
    else:
        _upload_large_attachment(
            token,
            sender,
            message_id,
            logo_path,
            content_id=INLINE_LOGO_CONTENT_ID,
            content_type=content_type,
        )


def _upload_large_attachment(
    token: str,
    sender: str,
    message_id: str,
    attachment_path: Path,
    *,
    content_id: str | None = None,
    content_type: str | None = None,
) -> None:
    size = attachment_path.stat().st_size
    attachment_item = {
        "attachmentType": "file",
        "name": attachment_path.name,
        "size": size,
    }
    if content_id:
        attachment_item["contentId"] = content_id
        attachment_item["isInline"] = True
        attachment_item["contentType"] = content_type or mimetypes.guess_type(attachment_path.name)[0] or "application/octet-stream"
    response = requests.post(
        _user_path(sender, f"messages/{quote(message_id, safe='')}/attachments/createUploadSession"),
        headers={**graph_headers(token), "Content-Type": "application/json"},
        json={"AttachmentItem": attachment_item},
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
    sharepoint_logo_path = _download_sharepoint_logo(graph_access_token, config)
    try:
        payload = build_direct_send_payload(config, attachment_path, sharepoint_logo_path)
        response = requests.post(
            _user_path(clean_sender, "sendMail"),
            headers={**graph_headers(graph_access_token), "Content-Type": "application/json"},
            json=payload,
            timeout=120,
        )
        _raise_for_graph(response, "enviar el correo")
        print(f"Correo enviado con {attachment_path.name} a {len(config['recipients'])} destinatario(s).", flush=True)
    finally:
        if sharepoint_logo_path:
            sharepoint_logo_path.unlink(missing_ok=True)
