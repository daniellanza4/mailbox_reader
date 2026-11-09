
"""Lógica de tickets y asociación de correos con PostgreSQL."""
import re
import secrets
import string

from app.services.db_service import get_connection
from app.services.mail_service import send_email

# Formato: [ABC123-XYZ789]
TICKET_CODE_PATTERN = re.compile(
    r"\[([A-Z0-9]{6}-[A-Z0-9]{6})\]",
    re.IGNORECASE,
)


def extract_ticket_code(subject="", body=""):
    """Extrae el primer código de ticket presente en asunto o cuerpo."""
    for text in (subject or "", body or ""):
        match = TICKET_CODE_PATTERN.search(text)
        if match:
            return match.group(1).upper()
    return None

def get_ticket_by_code(ticket_code):
    """Busca un ticket por su código en PostgreSQL."""
    if not ticket_code:
        return None

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT id, ticket_code, subject, requester_name,
                       requester_email, description, status,
                       created_at, updated_at
                FROM tickets
                WHERE ticket_code = %s
                """,
                (ticket_code.upper(),),
            )
            row = cursor.fetchone()
            if row is None:
                return None

            columns = [column.name for column in cursor.description]
            return dict(zip(columns, row))


def get_email_address(email_field):
    """Obtiene nombre y dirección de un campo from/sender de Graph."""
    if not isinstance(email_field, dict):
        return None, None

    address_data = email_field.get("emailAddress") or {}
    if not isinstance(address_data, dict):
        return None, None

    return address_data.get("name"), address_data.get("address")


def get_email_body(email):
    """Obtiene el cuerpo del correo de Microsoft Graph."""
    body = email.get("body")

    if isinstance(body, dict):
        return body.get("content", "") or ""

    if isinstance(body, str):
        return body

    return email.get("bodyPreview", "") or ""


def _insert_ticket_message(
    connection,
    ticket_id,
    email,
    message_type,
    sender_name=None,
    sender_email=None,
):
    """Inserta un mensaje y evita duplicados por Graph message ID."""
    if not email.get("id"):
        raise ValueError(
            "El mensaje no contiene un ID de Microsoft Graph."
        )

    if sender_name is None or sender_email is None:
        sender_name, sender_email = get_email_address(
            email.get("from") or email.get("sender")
        )

    with connection.cursor() as cursor:
        cursor.execute(
            """
            INSERT INTO ticket_messages (
                ticket_id,
                graph_message_id,
                internet_message_id,
                conversation_id,
                sender_name,
                sender_email,
                message_type,
                subject,
                body,
                received_at,
                sent_at
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON CONFLICT (graph_message_id) DO NOTHING
            RETURNING id
            """,
            (
                ticket_id,
                email.get("id"),
                email.get("internetMessageId"),
                email.get("conversationId"),
                sender_name,
                sender_email,
                message_type,
                email.get("subject") or "",
                get_email_body(email),
                email.get("receivedDateTime"),
                email.get("sentDateTime"),
            ),
        )
        return cursor.fetchone() is not None


def save_ticket_message(ticket_id, email, message_type):
    """Guarda un mensaje entrante relacionado con un ticket."""
    with get_connection() as connection:
        return _insert_ticket_message(
            connection,
            ticket_id,
            email,
            message_type,
        )


def save_outgoing_ticket_message(ticket_id, email, connection=None):
    """Guarda una respuesta saliente relacionada con un ticket."""
    if connection is None:
        with get_connection() as new_connection:
            return _insert_ticket_message(
                new_connection,
                ticket_id,
                email,
                "outgoing",
            )

    return _insert_ticket_message(
        connection,
        ticket_id,
        email,
        "outgoing",
    )


def generate_ticket_code():
    """Genera un código aleatorio con formato ABC123-XYZ789."""
    characters = string.ascii_uppercase + string.digits

    first_part = "".join(
        secrets.choice(characters) for _ in range(6)
    )
    second_part = "".join(
        secrets.choice(characters) for _ in range(6)
    )

    return f"{first_part}-{second_part}"


def create_ticket_with_message(email):
    """Crea un ticket nuevo y guarda el primer mensaje en una transacción."""
    sender_name, sender_email = get_email_address(
        email.get("from") or email.get("sender")
    )

    if not sender_email:
        raise ValueError(
            "El correo no contiene una dirección de remitente válida."
        )

    if not email.get("id"):
        raise ValueError(
            "El correo no contiene el identificador de Microsoft Graph."
        )

    subject = (email.get("subject") or "").strip() or "(Sin asunto)"
    body = get_email_body(email)

    with get_connection() as connection:
        with connection.cursor() as cursor:
            # Evitar crear otro ticket para el mismo mensaje.
            cursor.execute(
                """
                SELECT ticket_id
                FROM ticket_messages
                WHERE graph_message_id = %s
                """,
                (email["id"],),
            )
            existing = cursor.fetchone()

            if existing:
                return {
                    "created": False,
                    "reason": "duplicate",
                    "ticket_id": existing[0],
                }

            # Generar un código que todavía no exista.
            ticket_code = None

            for _ in range(10):
                candidate = generate_ticket_code()

                cursor.execute(
                    """
                    SELECT 1
                    FROM tickets
                    WHERE ticket_code = %s
                    """,
                    (candidate,),
                )

                if cursor.fetchone() is None:
                    ticket_code = candidate
                    break

            if ticket_code is None:
                raise RuntimeError(
                    "No se pudo generar un código único."
                )

            cursor.execute(
                """
                INSERT INTO tickets (
                    ticket_code,
                    subject,
                    requester_name,
                    requester_email,
                    description
                )
                VALUES (%s, %s, %s, %s, %s)
                RETURNING id
                """,
                (
                    ticket_code,
                    subject,
                    sender_name,
                    sender_email,
                    body,
                ),
            )

            ticket_id = cursor.fetchone()[0]

            _insert_ticket_message(
                connection,
                ticket_id,
                email,
                "incoming",
                sender_name=sender_name,
                sender_email=sender_email,
            )

    return {
        "created": True,
        "ticket_id": ticket_id,
        "ticket_code": ticket_code,
        "subject": subject,
    }


def process_ticket_reply(email):
    """Asocia un correo entrante a un ticket mediante su código."""
    subject = email.get("subject") or ""
    body = get_email_body(email)
    ticket_code = extract_ticket_code(subject, body)

    if not ticket_code:
        return {
            "processed": False,
            "reason": "missing_ticket_code",
        }

    ticket = get_ticket_by_code(ticket_code)

    if ticket is None:
        return {
            "processed": False,
            "reason": "ticket_not_found",
            "ticket_code": ticket_code,
        }

    saved = save_ticket_message(
        ticket_id=ticket["id"],
        email=email,
        message_type="incoming",
    )

    if not saved:
        return {
            "processed": False,
            "reason": "duplicate_message",
            "ticket_id": ticket["id"],
            "ticket_code": ticket_code,
        }

    return {
        "processed": True,
        "reason": "reply_saved",
        "ticket_id": ticket["id"],
        "ticket_code": ticket_code,
    }


def _find_ticket_by_conversation(conversation_id):
    """Devuelve los tickets únicos asociados a un conversationId."""
    if not conversation_id:
        return []

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT DISTINCT tm.ticket_id, t.ticket_code
                FROM ticket_messages AS tm
                INNER JOIN tickets AS t ON t.id = tm.ticket_id
                WHERE tm.conversation_id = %s
                """,
                (conversation_id,),
            )
            return cursor.fetchall()


def process_incoming_email(email):
    """Procesa un correo entrante evitando duplicados y tickets innecesarios."""
    graph_message_id = email.get("id")

    if not graph_message_id:
        return {
            "processed": False,
            "reason": "missing_graph_message_id",
        }

    # 1. Comprobar si el mensaje ya está registrado.
    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT ticket_id
                FROM ticket_messages
                WHERE graph_message_id = %s
                """,
                (graph_message_id,),
            )
            existing = cursor.fetchone()

    if existing:
        return {
            "processed": False,
            "reason": "duplicate_message",
            "ticket_id": existing[0],
        }

    subject = email.get("subject") or ""
    body = get_email_body(email)
    ticket_code = extract_ticket_code(subject, body)

    # 2. Primero, buscar el código explícito del ticket.
    if ticket_code:
        result = process_ticket_reply(email)

        if result.get("reason") != "ticket_not_found":
            return result

        # Si el código no existe, intentar relacionarlo por conversación
        # antes de crear un ticket nuevo.

    # 3. Buscar un ticket relacionado mediante conversationId.
    conversation_id = email.get("conversationId")
    matching_tickets = _find_ticket_by_conversation(conversation_id)

    if len(matching_tickets) == 1:
        ticket_id, existing_ticket_code = matching_tickets[0]

        saved = save_ticket_message(
            ticket_id=ticket_id,
            email=email,
            message_type="incoming",
        )

        if not saved:
            return {
                "processed": False,
                "reason": "duplicate_message",
                "ticket_id": ticket_id,
                "ticket_code": existing_ticket_code,
            }

        return {
            "processed": True,
            "reason": "conversation_reply_saved",
            "ticket_id": ticket_id,
            "ticket_code": existing_ticket_code,
        }

    if len(matching_tickets) > 1:
        return {
            "processed": False,
            "reason": "ambiguous_conversation",
            "conversation_id": conversation_id,
            "ticket_ids": [
                row[0] for row in matching_tickets
            ],
        }

    # 4. Solo crear un ticket si no se encontró relación existente.
    result = create_ticket_with_message(email)

    if result.get("created"):
        return {
            "processed": True,
            "reason": "ticket_created",
            "ticket_id": result["ticket_id"],
            "ticket_code": result["ticket_code"],
            "subject": result["subject"],
        }

    if result.get("reason") == "duplicate":
        return {
            "processed": False,
            "reason": "duplicate_message",
            "ticket_id": result.get("ticket_id"),
        }

    return {
        "processed": False,
        "reason": "not_processed",
    }


def process_incoming_emails(emails):
    """Procesa varios mensajes y captura errores individualmente."""
    results = []

    for email in emails:
        try:
            result = process_incoming_email(email)
        except Exception as exc:
            result = {
                "processed": False,
                "reason": "processing_error",
                "error": str(exc),
                "graph_message_id": email.get("id"),
            }

        results.append(result)

    return results


def reply_to_ticket(ticket_id, body):
    """Responde al mensaje entrante más reciente del cliente y guarda la salida."""
    if not body or not body.strip():
        raise ValueError("La respuesta no puede estar vacía.")

    with get_connection() as connection:
        with connection.cursor() as cursor:
            cursor.execute(
                """
                SELECT id, ticket_code, subject, requester_email, status
                FROM tickets
                WHERE id = %s
                """,
                (ticket_id,),
            )
            row = cursor.fetchone()

            if row is None:
                raise ValueError("No se encontró el ticket.")

            columns = [
                column.name for column in cursor.description
            ]
            ticket = dict(zip(columns, row))

            if not ticket["requester_email"]:
                raise ValueError(
                    "El ticket no tiene un correo de solicitante."
                )

            if ticket["status"] != "open":
                raise ValueError(
                    "El ticket no está abierto y no puede recibir respuestas."
                )

            # Elegir el mensaje entrante más reciente ya guardado.
            cursor.execute(
                """
                SELECT graph_message_id, conversation_id
                FROM ticket_messages
                WHERE ticket_id = %s
                  AND message_type = 'incoming'
                  AND graph_message_id IS NOT NULL
                ORDER BY received_at DESC NULLS LAST, id DESC
                LIMIT 1
                """,
                (ticket_id,),
            )
            original_row = cursor.fetchone()

            if original_row is None:
                raise ValueError(
                    "El ticket no tiene un mensaje entrante con ID de Graph. "
                    "No se enviará una respuesta independiente."
                )

            original_message_id, original_conversation_id = original_row

            reply_body = body.strip()
            ticket_reference = (
                f"\n\nReferencia del ticket: [{ticket['ticket_code']}]"
            )

            if ticket["ticket_code"].upper() not in reply_body.upper():
                reply_body += ticket_reference

            # createReply es quien debe vincular la respuesta al mensaje original.
            result = send_email(
                to_email=ticket["requester_email"],
                subject=ticket["subject"] or "",
                body=reply_body,
                original_message_id=original_message_id,
            )

            if not result.get("sent") or not result.get("message"):
                raise RuntimeError(
                    "No se confirmó el envío. Verifica Graph antes de reintentar."
                )

            message = result["message"]

            if not message.get("id"):
                raise RuntimeError(
                    "El correo pudo enviarse, pero no se obtuvo su ID. "
                    "Verifica Sent Items antes de reintentar."
                )

            # Este respaldo solo ayuda al historial local; no modifica el hilo real
            # de Outlook, que se determina al crear la respuesta en Graph.
            if (
                not result.get("details_retrieved", True)
                and not message.get("conversationId")
            ):
                message["conversationId"] = original_conversation_id

            saved = save_outgoing_ticket_message(
                ticket["id"],
                message,
                connection=connection,
            )

            if not saved:
                with connection.cursor() as cursor:
                    cursor.execute(
                        """
                        SELECT id
                        FROM ticket_messages
                        WHERE graph_message_id = %s
                          AND ticket_id = %s
                        """,
                        (message["id"], ticket["id"]),
                    )

                    if cursor.fetchone() is None:
                        raise RuntimeError(
                            "Graph aceptó el envío, pero no se confirmó "
                            "su registro en el ticket. No reenvíes sin verificar."
                        )

    return {
        "sent": result["sent"],
        "ticket_id": ticket["id"],
        "ticket_code": ticket["ticket_code"],
        "recipient": ticket["requester_email"],
        "subject": result.get("subject") or ticket["subject"],
        "message_id": result["message_id"],
        "details_retrieved": result.get("details_retrieved", True),
    }
