
from flask import Blueprint, jsonify, render_template, request

from app.services.mail_service import get_emails
from app.services.ticket_service import (
    get_email_address,
    process_incoming_email,
    process_incoming_emails,
    reply_to_ticket,
)

main_bp = Blueprint("main", __name__)

# Datos exactos del correo autorizado para la prueba.
TEST_SENDER = "developer6@experenta.com"
TEST_SUBJECT = "Prueba de ticket - acceso al sistema"


@main_bp.get("/")
def inbox():
    error = None
    emails = []

    try:
        emails = get_emails()
    except Exception as exc:
        error = str(exc)

    return render_template(
        "inbox.html",
        emails=emails,
        error=error,
    )


@main_bp.post("/api/test/process-email")
def process_test_email():
    """Procesa exclusivamente el correo de prueba autorizado."""

    try:
        emails = get_emails()
        test_emails = []

        for email in emails:
            _, sender_email = get_email_address(
                email.get("from") or email.get("sender")
            )

            subject = (email.get("subject") or "").strip()

            if (
                sender_email
                and sender_email.lower() == TEST_SENDER.lower()
                and subject == TEST_SUBJECT
            ):
                test_emails.append(email)

        if not test_emails:
            return jsonify({
                "processed": False,
                "reason": "test_email_not_found",
                "message": (
                    "No se encontró el correo autorizado entre "
                    "los 10 mensajes más recientes de la bandeja."
                ),
            }), 404

        # Conservamos el comportamiento de tu prueba inicial.
        email = test_emails[0]
        result = process_incoming_email(email)

        return jsonify({
            "sender": TEST_SENDER,
            "subject": TEST_SUBJECT,
            **result,
        })

    except Exception as exc:
        return jsonify({
            "processed": False,
            "reason": "processing_error",
            "error": str(exc),
        }), 500


@main_bp.post("/api/test/reply-ticket")
def test_reply_ticket():
    """Envía una respuesta de prueba a un ticket existente."""

    data = request.get_json(silent=True) or {}

    ticket_id = data.get("ticket_id")
    body = data.get("body")

    if (
        ticket_id is None
        or not isinstance(body, str)
        or not body.strip()
    ):
        return jsonify({
            "sent": False,
            "reason": "invalid_parameters",
            "message": "Debes indicar ticket_id y body.",
        }), 400

    try:
        ticket_id = int(ticket_id)

        if ticket_id <= 0:
            raise ValueError("El ID del ticket debe ser positivo.")

        result = reply_to_ticket(
            ticket_id=ticket_id,
            body=body,
        )

        return jsonify(result), 200

    except ValueError as exc:
        return jsonify({
            "sent": False,
            "reason": "validation_error",
            "message": str(exc),
        }), 400

    except Exception as exc:
        return jsonify({
            "sent": False,
            "reason": "reply_error",
            "error": str(exc),
        }), 500

@main_bp.post("/api/test/process-latest-reply")
def process_latest_test_reply():
    """Procesa la respuesta más reciente de la prueba."""

    try:
        emails = get_emails()
        candidates = []

        for email in emails:
            _, sender_email = get_email_address(
                email.get("from") or email.get("sender")
            )

            subject = (email.get("subject") or "").strip()

            if (
                sender_email
                and sender_email.lower() == TEST_SENDER.lower()
                and subject.casefold().startswith(
                    "re: " + TEST_SUBJECT.casefold()
                )
            ):
                candidates.append(email)

        if not candidates:
            return jsonify({
                "processed": False,
                "reason": "reply_not_found",
                "message": (
                    "No se encontró la respuesta en los mensajes "
                    "recientes de la bandeja de soporte."
                ),
            }), 404

        candidates.sort(
            key=lambda email: (
                email.get("receivedDateTime") or ""
            ),
            reverse=True,
        )

        email = candidates[0]
        results = process_incoming_emails([email])

        return jsonify({
            "subject": email.get("subject"),
            "result": results[0],
        }), 200

    except Exception as exc:
        return jsonify({
            "processed": False,
            "reason": "processing_error",
            "error": str(exc),
        }), 500
