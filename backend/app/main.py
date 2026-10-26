from __future__ import annotations

from fastapi import FastAPI
from starlette.middleware.cors import CORSMiddleware
from starlette.middleware.sessions import SessionMiddleware

from app.auth.routes import router as auth_router
from app.auth.service import MicrosoftAuthService
from app.config import Settings, get_settings
from app.outlook.routes import router as outlook_router
from app.outlook.service import OutlookGraphService
from app.zimbra.routes import router as zimbra_router
from app.zimbra.service import ZimbraSMTPService


def create_app(
    settings: Settings | None = None,
    auth_service: MicrosoftAuthService | None = None,
) -> FastAPI:
    configured_settings = settings or get_settings()
    app = FastAPI(title="Outlook-Zimbra API", version="0.1.0")
    app.state.settings = configured_settings
    app.state.auth_service = auth_service or MicrosoftAuthService(configured_settings)
    app.state.outlook_service = OutlookGraphService(
        admin_email=configured_settings.microsoft_admin_email
    )
    app.state.zimbra_service = ZimbraSMTPService(configured_settings)

    app.add_middleware(
        SessionMiddleware,
        secret_key=configured_settings.session_secret_key.get_secret_value(),
        session_cookie="outlook_zimbra_session",
        max_age=configured_settings.session_max_age_seconds,
        same_site="lax",
        https_only=configured_settings.cookie_secure,
    )
    allowed_origins = list(
        {
            configured_settings.frontend_origin,
            configured_settings.frontend_origin.replace("localhost", "127.0.0.1"),
            configured_settings.frontend_origin.replace("127.0.0.1", "localhost"),
        }
    )
    app.add_middleware(
        CORSMiddleware,
        allow_origins=allowed_origins,
        allow_credentials=True,
        allow_methods=["GET", "POST"],
        allow_headers=["Content-Type"],
    )
    app.include_router(auth_router)
    app.include_router(outlook_router)
    app.include_router(zimbra_router)

    @app.get("/health", tags=["health"])
    def health() -> dict[str, str]:
        return {"status": "ok"}

    return app


app = create_app()
