from __future__ import annotations

import re
from typing import Annotated, Any

from fastapi import APIRouter, Depends, HTTPException, Request, status
from pydantic import BaseModel, Field, field_validator

from app.auth.routes import require_graph_access_token
from app.zimbra.service import ZimbraSMTPError, ZimbraSMTPService

router = APIRouter(prefix="/api/v1/zimbra", tags=["zimbra"])

# Dependency type alias for a verified session token (enforces authentication).
_AuthRequired = Annotated[str, Depends(require_graph_access_token)]

_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


class BatchRequest(BaseModel):
    message_ids: list[str] = Field(min_length=1)
    recipients: list[str] = Field(min_length=1)
    subject: str = Field(min_length=1, max_length=255)
    body: str = Field(min_length=1)

    @field_validator("recipients")
    @classmethod
    def validate_recipient_emails(cls, value: list[str]) -> list[str]:
        cleaned: list[str] = []
        for raw in value:
            addr = raw.strip()
            if not addr:
                continue
            if not _EMAIL_RE.match(addr):
                raise ValueError(f"Invalid email address: {addr}")
            cleaned.append(addr)
        if not cleaned:
            raise ValueError("recipients_required")
        return cleaned


def get_zimbra_service(request: Request) -> ZimbraSMTPService:
    service = getattr(request.app.state, "zimbra_service", None)
    if service is None:
        service = ZimbraSMTPService(request.app.state.settings)
        request.app.state.zimbra_service = service
    return service


ZimbraServiceDependency = Annotated[ZimbraSMTPService, Depends(get_zimbra_service)]


@router.post("/reenvio-masivo")
def submit_batch(
    payload: BatchRequest,
    service: ZimbraServiceDependency,
    _token: _AuthRequired,
) -> dict[str, Any]:
    try:
        return service.submit_batch(
            message_ids=payload.message_ids,
            recipients=payload.recipients,
            subject=payload.subject,
            body=payload.body,
        )
    except ZimbraSMTPError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=exc.message,
        ) from exc


@router.get("/reenvio-masivo/{batch_id}")
def get_batch_status(
    batch_id: str,
    service: ZimbraServiceDependency,
    _token: _AuthRequired,
) -> dict[str, Any]:
    try:
        return service.get_batch(batch_id)
    except ZimbraSMTPError as exc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=exc.message,
        ) from exc
