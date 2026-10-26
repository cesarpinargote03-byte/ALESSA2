from __future__ import annotations

import smtplib
import time
import uuid
from email.message import EmailMessage
from typing import Any

from app.config import Settings


class ZimbraSMTPError(RuntimeError):
    def __init__(self, message: str, *, batch_id: str | None = None) -> None:
        self.message = message
        self.batch_id = batch_id
        super().__init__(message)


class ZimbraSMTPService:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.batches: dict[str, dict[str, Any]] = {}

    def submit_batch(
        self,
        *,
        message_ids: list[str],
        recipients: list[str],
        subject: str,
        body: str,
    ) -> dict[str, Any]:
        if not message_ids:
            raise ZimbraSMTPError("message_ids_required")
        if not recipients:
            raise ZimbraSMTPError("recipients_required")

        recipients = [recipient.strip() for recipient in recipients if recipient.strip()]
        if not recipients:
            raise ZimbraSMTPError("recipients_required")

        batch_id = str(uuid.uuid4())
        batch = {
            "batch_id": batch_id,
            "status": "queued",
            "created_at": time.time(),
            "total_messages": len(message_ids),
            "total_recipients": len(recipients),
            "results": [],
        }
        self.batches[batch_id] = batch

        try:
            results = self._send_batch(batch_id, subject=subject, body=body, recipients=recipients)
        except Exception as exc:  # pragma: no cover - defensive path
            batch["status"] = "failed"
            batch["error"] = str(exc)
            raise ZimbraSMTPError("zimbra_send_failed", batch_id=batch_id) from exc

        batch["status"] = "completed"
        batch["results"] = results
        return {"batch_id": batch_id, "status": batch["status"]}

    def _send_batch(
        self,
        batch_id: str,
        *,
        subject: str,
        body: str,
        recipients: list[str],
    ) -> list[dict[str, Any]]:
        results: list[dict[str, Any]] = []
        for index, recipient in enumerate(recipients, start=1):
            try:
                self._send_message(recipient=recipient, subject=subject, body=body)
                status = "sent"
                error = None
            except Exception as exc:  # pragma: no cover - simple framework for retries
                status = "failed"
                error = str(exc)
            results.append(
                {
                    "recipient": recipient,
                    "index": index,
                    "status": status,
                    "error": error,
                }
            )

        failure_count = sum(1 for item in results if item["status"] == "failed")
        batch = self.batches[batch_id]
        batch["results"] = results
        batch["failed_recipients"] = failure_count
        if failure_count:
            batch["status"] = "partial_failure"
        return results

    def _send_message(self, *, recipient: str, subject: str, body: str) -> None:
        message = EmailMessage()
        message["From"] = self.settings.zimbra_from_email
        message["To"] = recipient
        message["Subject"] = subject
        message.set_content(body)

        for attempt in range(1, self.settings.zimbra_retry_count + 1):
            try:
                with smtplib.SMTP(self.settings.zimbra_host, self.settings.zimbra_port, timeout=30) as smtp:
                    if self.settings.zimbra_use_tls:
                        smtp.starttls()
                    if self.settings.zimbra_username and self.settings.zimbra_password:
                        smtp.login(
                            self.settings.zimbra_username,
                            self.settings.zimbra_password.get_secret_value(),
                        )
                    smtp.send_message(message)
                    return
            except (smtplib.SMTPException, OSError) as exc:
                if attempt == self.settings.zimbra_retry_count:
                    raise exc
                time.sleep(self.settings.zimbra_retry_delay_seconds)

    def get_batch(self, batch_id: str) -> dict[str, Any]:
        batch = self.batches.get(batch_id)
        if batch is None:
            raise ZimbraSMTPError("batch_not_found", batch_id=batch_id)
        return {
            "batch_id": batch["batch_id"],
            "status": batch["status"],
            "created_at": batch["created_at"],
            "total_messages": batch["total_messages"],
            "total_recipients": batch["total_recipients"],
            "results": batch.get("results", []),
            "error": batch.get("error"),
        }
