from __future__ import annotations

from datetime import date
from typing import Annotated, Any, Literal

from fastapi import APIRouter, Depends, HTTPException, Query, Request, Response, status
from pydantic import BaseModel, Field

from app.auth.routes import require_graph_access_token
from app.outlook.service import OutlookGraphError, OutlookGraphService

router = APIRouter(prefix="/api/v1", tags=["outlook"])


class AssignmentPayload(BaseModel):
    message_id: str = Field(min_length=1)
    start_date: date
    end_date: date
    time_zone: str = "America/Santiago"
    subject: str = "Planificación semanal"
    body: str | None = None
    postergado_por_proyecto: str = "No"


class ExcelPlanningRow(BaseModel):
    semana: str
    postergado_por_proyecto: str = "No"
    usuario: str
    correo: str
    asunto: str


class ExcelDownloadPayload(BaseModel):
    filename: str = "Planificacion_Helpdesk.xlsx"
    rows: list[ExcelPlanningRow]


class ExcelSendPayload(BaseModel):
    recipients: list[str]
    cc_recipients: list[str] = Field(default_factory=list)
    subject: str = "Planificación Semanal de Solicitudes - Helpdesk"
    intro_text: str = (
        "Estimados usuarios,\n"
        "Adjunto encontrarán el cuadro de planificación semanal de solicitudes registradas en Helpdesk."
    )
    filename: str = "Planificacion_Helpdesk.xlsx"
    mode: Literal["send", "display"] = "send"
    rows: list[ExcelPlanningRow]


def get_outlook_service(request: Request) -> OutlookGraphService:
    service = getattr(request.app.state, "outlook_service", None)
    if service is None:
        settings = getattr(request.app.state, "settings", None)
        admin_email = getattr(settings, "microsoft_admin_email", None)
        service = OutlookGraphService(admin_email=admin_email)
        request.app.state.outlook_service = service
    return service


OutlookServiceDependency = Annotated[OutlookGraphService, Depends(get_outlook_service)]


@router.post("/outlook/sync")
def sync_outlook_messages(
    request: Request,
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
    limit: int = Query(default=25, ge=1, le=5000),
    folder: str | None = Query(default=None),
) -> dict[str, Any]:
    try:
        return service.sync_messages(access_token, limit=limit, folder=folder)
    except OutlookGraphError as exc:
        status_code = status.HTTP_503_SERVICE_UNAVAILABLE
        if exc.status_code == 429:
            status_code = status.HTTP_429_TOO_MANY_REQUESTS
        raise HTTPException(status_code=status_code, detail=exc.message) from exc


@router.get("/outlook/filtros")
def get_outlook_filters(
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
) -> dict[str, Any]:
    del access_token
    return service.get_filter_summary()


@router.get("/correos")
def list_messages(
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
    categoria: str | None = Query(default=None),
    ano: int | None = Query(default=None, ge=2000, le=2100),
    folder: str | None = Query(default=None, alias="folder"),
    remitente: str | None = Query(default=None),
    page: int = Query(default=1, ge=1),
    page_size: int = Query(default=500, ge=1, le=10000),
) -> dict[str, Any]:
    del access_token
    return service.list_messages(
        categoria=categoria,
        ano=ano,
        folder=folder,
        remitente=remitente,
        page=page,
        page_size=page_size,
    )


@router.get("/planificacion")
def list_planned_messages(
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
) -> dict[str, Any]:
    del access_token
    items = service.list_planned_messages()
    return {"count": len(items), "items": items}


@router.post("/planificacion/asignar")
def assign_week_plan(
    payload: AssignmentPayload,
    service: OutlookGraphServiceDependency if False else OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
) -> dict[str, Any]:
    try:
        return service.create_event_for_message(
            access_token,
            message_id=payload.message_id,
            start_date=payload.start_date,
            end_date=payload.end_date,
            time_zone=payload.time_zone,
            subject=payload.subject,
            body=payload.body,
            postergado_por_proyecto=payload.postergado_por_proyecto,
        )
    except OutlookGraphError as exc:
        status_code = status.HTTP_400_BAD_REQUEST
        if exc.status_code in {409, 429}:
            status_code = exc.status_code
        elif exc.status_code == 503:
            status_code = status.HTTP_503_SERVICE_UNAVAILABLE
        raise HTTPException(status_code=status_code, detail=exc.message) from exc


@router.post("/planificacion/excel")
def download_planning_excel(
    payload: ExcelDownloadPayload,
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
) -> Response:
    del access_token
    rows_dicts = [r.model_dump() for r in payload.rows]
    xlsx_bytes = service.build_planning_excel_bytes(rows_dicts)
    filename = payload.filename or "Planificacion_Helpdesk.xlsx"
    if not filename.endswith(".xlsx"):
        filename += ".xlsx"
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
        },
    )


@router.post("/planificacion/enviar-excel")
def send_planning_excel(
    payload: ExcelSendPayload,
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
) -> dict[str, Any]:
    del access_token
    try:
        rows_dicts = [r.model_dump() for r in payload.rows]
        return service.send_planning_excel_email(
            recipients=payload.recipients,
            cc_recipients=payload.cc_recipients,
            subject=payload.subject,
            intro_text=payload.intro_text,
            rows=rows_dicts,
            filename=payload.filename,
            mode=payload.mode,
        )
    except OutlookGraphError as exc:
        status_code = exc.status_code or status.HTTP_400_BAD_REQUEST
        raise HTTPException(status_code=status_code, detail=exc.message) from exc


@router.get("/planificacion/envios")
def list_sent_planning_batches(
    service: OutlookServiceDependency,
    access_token: Annotated[str, Depends(require_graph_access_token)],
) -> dict[str, Any]:
    del access_token
    items = service.list_sent_batches()
    return {"count": len(items), "items": items}
