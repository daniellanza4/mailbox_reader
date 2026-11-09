"""Integración de correo con Microsoft Graph para el buzón compartido de soporte.

Nota: se usa IdType="ImmutableId" de forma consistente en las peticiones a Graph.
Los identificadores guardados anteriormente con el formato predeterminado de Graph
pueden no coincidir con los nuevos IDs inmutables; no borres registros históricos
sin revisar primero la base de datos.
"""
import threading
import time
import msal
import requests

CLIENT_ID = "c260224f-67cd-4da5-b48a-9a5d4f32c50f"
TENANT_ID = "29663921-3da7-4330-913f-c36f76fbe809"
SHARED_MAILBOX = "soporte@experenta.com"

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "User.Read",
    "Mail.Read.Shared",
    "Mail.ReadWrite.Shared",
    "Mail.Send.Shared",
]
GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
MESSAGE_SELECT = (
    "id,internetMessageId,conversationId,conversationIndex,"
    "subject,body,uniqueBody,bodyPreview,from,sender,toRecipients,"
    "ccRecipients,replyTo,receivedDateTime,sentDateTime,createdDateTime,"
    "lastModifiedDateTime,hasAttachments,importance,isRead,isDraft,"
    "parentFolderId"
)

_msal_app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
_access_token = None
_auth_lock = threading.Lock()


def get_access_token():
    """Obtiene un token válido usando caché silenciosa o Device Code Flow."""
    global _access_token

    with _auth_lock:
        accounts = _msal_app.get_accounts()
        result = None

        if accounts:
            result = _msal_app.acquire_token_silent(
                SCOPES, account=accounts[0]
            )

        if result and result.get("access_token"):
            _access_token = result["access_token"]
            return _access_token

        _access_token = None
        flow = _msal_app.initiate_device_flow(scopes=SCOPES)

        if not flow or "user_code" not in flow:
            raise RuntimeError(
                "No se pudo iniciar la autenticación con Microsoft."
            )

        print(flow["message"], flush=True)
        result = _msal_app.acquire_token_by_device_flow(flow)

        if not result or not result.get("access_token"):
            detail = result.get("error_description") if result else None
            raise RuntimeError(
                detail or "Error de autenticación con Microsoft."
            )

        _access_token = result["access_token"]
        return _access_token


def _headers(json_content=False, text_body=False):
    """Construye encabezados comunes, manteniendo IDs inmutables de forma consistente."""
    headers = {
        "Authorization": f"Bearer {get_access_token()}",
        "Prefer": 'IdType="ImmutableId"',
    }
    if json_content:
        headers["Content-Type"] = "application/json"
    if text_body:
        headers["Prefer"] += ', outlook.body-content-type="text"'
    return headers


def _request_with_retries(method, url, *, max_attempts=3, **kwargs):
    """Reintenta errores transitorios de Graph; no reintenta envíos de correo."""
    response = None

    for attempt in range(max_attempts):
        response = requests.request(method, url, **kwargs)

        if response.status_code not in (429, 502, 503, 504):
            return response

        if attempt == max_attempts - 1:
            return response

        retry_after = response.headers.get("Retry-After")
        try:
            wait_seconds = (
                min(max(int(retry_after), 1), 30)
                if retry_after
                else min(2 ** (attempt + 1), 10)
            )
        except (TypeError, ValueError):
            wait_seconds = min(2 ** (attempt + 1), 10)

        print(
            f"Microsoft Graph devolvió HTTP {response.status_code}; "
            f"reintentando en {wait_seconds} segundos.",
            flush=True,
        )
        time.sleep(wait_seconds)

    return response


def get_emails_from_folder(folder_name, top=10):
    """Obtiene mensajes recientes de una carpeta del buzón compartido."""
    if not folder_name or "/" in folder_name:
        raise ValueError("Nombre de carpeta no válido.")

    url = (
        f"{GRAPH_BASE_URL}/users/{SHARED_MAILBOX}/"
        f"mailFolders/{folder_name}/messages"
    )
    params = {
        "$top": max(1, min(int(top), 100)),
        "$select": MESSAGE_SELECT,
        "$orderby": (
            "sentDateTime desc"
            if folder_name.lower() in ("sentitems", "sent items")
            else "receivedDateTime desc"
        ),
    }

    response = _request_with_retries(
        "GET",
        url,
        headers=_headers(text_body=True),
        params=params,
        timeout=30,
    )
    response.raise_for_status()
    return response.json().get("value", [])


def get_emails():
    """Obtiene mensajes recibidos en Inbox."""
    return get_emails_from_folder("inbox")


def get_sent_emails():
    """Obtiene mensajes de Sent Items."""
    return get_emails_from_folder("sentitems")


def get_recent_sent_emails(top=20):
    """Obtiene los mensajes enviados más recientes."""
    return get_emails_from_folder("sentitems", top=top)


def find_sent_email(subject, recipient, max_attempts=5):
    """Busca por asunto y destinatario; no debe usarse como identificador de hilo."""
    if not subject or not recipient:
        return None

    wanted_subject = subject.strip().casefold()
    wanted_recipient = recipient.strip().casefold()

    for attempt in range(max_attempts):
        for email in get_recent_sent_emails(top=50):
            email_subject = (email.get("subject") or "").strip().casefold()
            recipients = email.get("toRecipients") or []
            addresses = [
                ((item.get("emailAddress") or {}).get("address") or "")
                .strip()
                .casefold()
                for item in recipients
            ]
            if (
                email_subject == wanted_subject
                and wanted_recipient in addresses
                and email.get("id")
            ):
                return email

        if attempt < max_attempts - 1:
            time.sleep(2)

    return None


def get_message(message_id):
    """Recupera un mensaje concreto por su ID del buzón compartido."""
    if not message_id:
        raise ValueError("El identificador del mensaje es obligatorio.")

    url = f"{GRAPH_BASE_URL}/users/{SHARED_MAILBOX}/messages/{message_id}"
    params = {"$select": MESSAGE_SELECT}
    response = _request_with_retries(
        "GET",
        url,
        headers=_headers(text_body=True),
        params=params,
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


def send_email(to_email, subject, body, original_message_id=None):
    """Envía un correo nuevo o una respuesta vinculada a un mensaje existente.

    Para una respuesta, original_message_id debe ser un ID real de Graph
    perteneciente al buzón compartido. No se fabrican conversationId ni IDs.
    """
    if not to_email or not to_email.strip():
        raise ValueError("El destinatario es obligatorio.")
    if not subject or not subject.strip():
        raise ValueError("El asunto es obligatorio.")
    if not body or not body.strip():
        raise ValueError("El contenido del correo es obligatorio.")

    to_email = to_email.strip()
    subject = subject.strip()
    body = body.strip()
    base_url = f"{GRAPH_BASE_URL}/users/{SHARED_MAILBOX}"
    headers = _headers(json_content=True)
    draft_subject = None

    # 1. Crear un borrador vinculado al mensaje original, o uno nuevo.
    if original_message_id:
        create_url = (
            f"{base_url}/messages/{original_message_id}/createReply"
        )
        response = _request_with_retries(
            "POST",
            create_url,
            headers=headers,
            json={"comment": body},
            timeout=30,
        )
        if response.status_code != 201:
            raise RuntimeError(
                "Microsoft Graph no pudo crear la respuesta al mensaje "
                f"original. HTTP {response.status_code}: {response.text}"
            )

        draft = response.json()
        message_id = draft.get("id")
        draft_subject = draft.get("subject")
        if not message_id:
            raise RuntimeError(
                "Graph creó el borrador de respuesta sin devolver un ID. "
                "Verifica Outlook antes de volver a intentar."
            )

        # No sobrescribir el asunto ni el cuerpo de createReply: conservar
        # la relación y el historial que Graph generó para la respuesta.
        patch_url = f"{base_url}/messages/{message_id}"
        response = _request_with_retries(
            "PATCH",
            patch_url,
            headers=headers,
            json={
                "toRecipients": [
                    {"emailAddress": {"address": to_email}}
                ]
            },
            timeout=30,
        )
        if response.status_code != 200:
            raise RuntimeError(
                "Se creó el borrador de respuesta, pero no se pudo "
                "actualizar el destinatario. Revisa Borradores en Outlook. "
                f"HTTP {response.status_code}: {response.text}. "
                f"ID del borrador: {message_id}"
            )
    else:
        # Solo usar este camino cuando de verdad se necesita un correo nuevo.
        payload = {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [
                {"emailAddress": {"address": to_email}}
            ],
        }
        response = _request_with_retries(
            "POST",
            f"{base_url}/messages",
            headers=headers,
            json=payload,
            timeout=30,
        )
        if response.status_code != 201:
            raise RuntimeError(
                "Microsoft Graph no pudo crear el borrador. "
                f"HTTP {response.status_code}: {response.text}"
            )

        draft = response.json()
        message_id = draft.get("id")
        draft_subject = draft.get("subject") or subject
        if not message_id:
            raise RuntimeError(
                "Graph creó el borrador sin devolver un identificador."
            )

    # 2. Enviar una sola vez. No reintentar automáticamente un POST /send,
    # porque un timeout podría ocurrir después de que Graph aceptara el envío.
    response = requests.post(
        f"{base_url}/messages/{message_id}/send",
        headers=headers,
        timeout=30,
    )
    if response.status_code != 202:
        raise RuntimeError(
            "Microsoft Graph no confirmó la aceptación del envío. "
            "No reenvíes sin verificar Sent Items. "
            f"HTTP {response.status_code}: {response.text}. "
            f"ID del mensaje: {message_id}"
        )

    # 3. Recuperar metadatos reales del mensaje enviado. Nunca volver a enviar
    # si esta lectura falla: el envío ya fue aceptado.
    sent_message = None
    retrieval_error = None
    message_url = f"{base_url}/messages/{message_id}"
    params = {
        "$select": (
            "id,internetMessageId,conversationId,conversationIndex,"
            "subject,body,from,sender,toRecipients,ccRecipients,"
            "sentDateTime,receivedDateTime,parentFolderId"
        )
    }

    for attempt in range(3):
        try:
            read_response = requests.get(
                message_url,
                headers=_headers(text_body=True),
                params=params,
                timeout=30,
            )
            if read_response.status_code == 200:
                sent_message = read_response.json()
                if sent_message.get("id"):
                    break
                retrieval_error = "Graph no devolvió el ID del mensaje."
            else:
                retrieval_error = (
                    f"HTTP {read_response.status_code}: "
                    f"{read_response.text}"
                )
        except requests.RequestException as exc:
            retrieval_error = str(exc)

        if attempt < 2:
            time.sleep(2 ** (attempt + 1))

    details_retrieved = bool(sent_message and sent_message.get("id"))

    if not details_retrieved:
        print(
            "ADVERTENCIA: Graph aceptó el envío, pero no se pudieron "
            f"recuperar sus detalles. ID={message_id}; error={retrieval_error}",
            flush=True,
        )
        # No inventar conversationId. ticket_service puede usar el ID del
        # mensaje original como referencia de respaldo para su registro local.
        sent_message = {
            "id": message_id,
            "subject": draft_subject or subject,
            "body": {"contentType": "text", "content": body},
            "from": {"emailAddress": {"address": SHARED_MAILBOX}},
            "toRecipients": [
                {"emailAddress": {"address": to_email}}
            ],
            "sentDateTime": None,
        }

    return {
        "sent": True,
        "message_id": message_id,
        "recipient": to_email,
        "subject": sent_message.get("subject") or draft_subject or subject,
        "message": sent_message,
        "details_retrieved": details_retrieved,
    }
