from __future__ import annotations

from typing import Annotated

from fastapi import APIRouter, Depends, HTTPException, Request, status
from fastapi.responses import RedirectResponse

from app.auth.service import AuthenticationError, MicrosoftAuthService
from app.config import Settings

router = APIRouter(prefix="/api/v1/auth", tags=["authentication"])


def get_auth_service(request: Request) -> MicrosoftAuthService:
    return request.app.state.auth_service


def get_settings(request: Request) -> Settings:
    return request.app.state.settings


AuthServiceDependency = Annotated[MicrosoftAuthService, Depends(get_auth_service)]
SettingsDependency = Annotated[Settings, Depends(get_settings)]


def require_graph_access_token(
    request: Request,
    auth_service: AuthServiceDependency,
) -> str:
    account_id = request.session.get("account_id")
    if not isinstance(account_id, str):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="authentication_required",
        )

    access_token = auth_service.acquire_graph_access_token(account_id)
    if access_token is None:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="reauthentication_required",
        )
    return access_token


@router.get("/login")
def login(
    request: Request,
    auth_service: AuthServiceDependency,
    settings: SettingsDependency,
) -> RedirectResponse:
    if settings.outlook_mode == "local":
        try:
            account_id = auth_service.authenticate_local_outlook()
        except AuthenticationError as exc:
            raise HTTPException(
                status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
                detail=exc.code,
            ) from exc

        request.session.clear()
        request.session["account_id"] = account_id
        return RedirectResponse(
            settings.frontend_redirect_url,
            status_code=status.HTTP_303_SEE_OTHER,
        )

    try:
        flow = auth_service.create_auth_flow()
    except AuthenticationError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=exc.code,
        ) from exc

    request.session["auth_flow"] = flow
    return RedirectResponse(flow["auth_uri"], status_code=status.HTTP_302_FOUND)


@router.post("/login-local")
def login_local(
    request: Request,
    auth_service: AuthServiceDependency,
    settings: SettingsDependency,
) -> dict[str, str | bool]:
    try:
        account_id = auth_service.authenticate_local_outlook()
    except AuthenticationError as exc:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail=exc.code,
        ) from exc

    request.session.clear()
    request.session["account_id"] = account_id
    email = (
        account_id.removeprefix("local:")
        if account_id.startswith("local:")
        else settings.microsoft_admin_email
    )
    return {
        "authenticated": True,
        "email": email,
        "mode": "local",
    }


@router.get("/callback")
def auth_callback(
    request: Request,
    auth_service: AuthServiceDependency,
    settings: SettingsDependency,
) -> RedirectResponse:
    flow = request.session.pop("auth_flow", None)
    if not isinstance(flow, dict):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="authorization_session_missing",
        )

    try:
        account_id = auth_service.complete_auth_flow(flow, dict(request.query_params))
    except AuthenticationError as exc:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=exc.code,
        ) from exc

    request.session.clear()
    request.session["account_id"] = account_id
    return RedirectResponse(settings.frontend_redirect_url, status_code=status.HTTP_303_SEE_OTHER)


@router.get("/status")
def auth_status(request: Request, settings: SettingsDependency) -> dict[str, bool | str]:
    return {
        "authenticated": isinstance(request.session.get("account_id"), str),
    }


@router.get("/me")
def current_admin(
    request: Request,
    settings: SettingsDependency,
) -> dict[str, str]:
    account_id = request.session.get("account_id")
    if not isinstance(account_id, str):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="authentication_required",
        )
    email = (
        account_id.removeprefix("local:")
        if account_id.startswith("local:") and not settings.microsoft_admin_email
        else settings.microsoft_admin_email or account_id.removeprefix("local:")
    )
    return {"email": email}


@router.post("/logout", status_code=status.HTTP_204_NO_CONTENT)
def logout(request: Request) -> None:
    request.session.clear()
