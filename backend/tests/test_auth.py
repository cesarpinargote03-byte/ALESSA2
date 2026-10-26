from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest
from cryptography.fernet import Fernet
from fastapi.testclient import TestClient

from app.auth.cache import EncryptedTokenCache, TokenCacheError
from app.auth.service import MicrosoftAuthService
from app.config import Settings
from app.main import create_app


def make_settings(tmp_path: Path) -> Settings:
    return Settings(
        microsoft_client_id="test-client",
        microsoft_client_secret="test-client-secret",
        microsoft_admin_email="admin@example.com",
        token_cache_encryption_key=Fernet.generate_key().decode("ascii"),
        token_cache_path=tmp_path / "tokens.bin",
        session_secret_key="test-only-session-secret-key-at-least-32",
        frontend_origin="http://localhost:8100",
        frontend_redirect_url="http://localhost:8100/",
    )


def test_encrypted_cache_round_trip(tmp_path: Path) -> None:
    key = Fernet.generate_key().decode("ascii")
    cache_path = tmp_path / "msal" / "tokens.bin"
    cache_store = EncryptedTokenCache(cache_path, key)

    with cache_store.session() as cache:
        cache.add(
            {
                "client_id": "test-client",
                "scope": ["User.Read"],
                "environment": "login.microsoftonline.com",
                "response": {
                    "token_type": "Bearer",
                    "access_token": "not-a-real-token",
                    "expires_in": 3600,
                },
            }
        )

    encrypted_contents = cache_path.read_bytes()
    assert b"not-a-real-token" not in encrypted_contents

    with cache_store.session() as restored_cache:
        assert restored_cache.find("AccessToken")[0]["secret"] == "not-a-real-token"


def test_encrypted_cache_rejects_wrong_key(tmp_path: Path) -> None:
    cache_path = tmp_path / "tokens.bin"
    first_store = EncryptedTokenCache(
        cache_path,
        Fernet.generate_key().decode("ascii"),
    )
    with first_store.session() as cache:
        cache.add(
            {
                "client_id": "test-client",
                "scope": ["User.Read"],
                "environment": "login.microsoftonline.com",
                "response": {
                    "token_type": "Bearer",
                    "access_token": "not-a-real-token",
                    "expires_in": 3600,
                },
            }
        )

    wrong_key_store = EncryptedTokenCache(
        cache_path,
        Fernet.generate_key().decode("ascii"),
    )
    with pytest.raises(TokenCacheError):
        with wrong_key_store.session():
            pass


def test_login_and_callback_create_admin_session(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    settings = make_settings(tmp_path)

    class FakeMsalClient:
        def __init__(self, **kwargs: Any) -> None:
            assert kwargs["authority"] == "https://login.microsoftonline.com/consumers"
            self.cache = kwargs["token_cache"]

        def initiate_auth_code_flow(self, **kwargs: Any) -> dict[str, Any]:
            return {
                "auth_uri": "https://login.example.test/authorize",
                "state": "test-state",
            }

        def acquire_token_by_auth_code_flow(
            self,
            flow: dict[str, Any],
            auth_response: dict[str, str],
        ) -> dict[str, Any]:
            assert flow["state"] == auth_response["state"]
            self.cache.add(
                {
                    "client_id": "test-client",
                    "scope": ["User.Read"],
                    "environment": "login.microsoftonline.com",
                    "response": {
                        "token_type": "Bearer",
                        "access_token": "test-token",
                        "expires_in": 3600,
                    },
                }
            )
            return {
                "access_token": "test-token",
                "id_token_claims": {
                    "preferred_username": "admin@example.com",
                },
            }

        def get_accounts(self) -> list[dict[str, str]]:
            return [
                {
                    "home_account_id": "admin-account",
                    "username": "admin@example.com",
                }
            ]

    monkeypatch.setattr(
        "app.auth.service.msal.ConfidentialClientApplication",
        FakeMsalClient,
    )
    auth_service = MicrosoftAuthService(settings)
    client = TestClient(create_app(settings, auth_service))

    login_response = client.get("/api/v1/auth/login", follow_redirects=False)
    assert login_response.status_code == 302
    assert login_response.headers["location"] == "https://login.example.test/authorize"

    callback_response = client.get(
        "/api/v1/auth/callback",
        params={"code": "test-code", "state": "test-state"},
        follow_redirects=False,
    )
    assert callback_response.status_code == 303
    assert callback_response.headers["location"] == settings.frontend_redirect_url
    session_cookie = callback_response.headers["set-cookie"]
    assert "outlook_zimbra_session=" in session_cookie
    assert "httponly" in session_cookie.lower()
    assert "test-token" not in session_cookie
    assert "admin-account" not in session_cookie
    assert client.get("/api/v1/auth/status").json() == {"authenticated": True}
    assert client.get("/api/v1/auth/me").json() == {"email": "admin@example.com"}


def test_callback_rejects_different_personal_account(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    settings = make_settings(tmp_path)

    class DifferentAccountClient:
        def __init__(self, **kwargs: Any) -> None:
            assert kwargs["authority"] == "https://login.microsoftonline.com/consumers"
            self.cache = kwargs["token_cache"]

        def initiate_auth_code_flow(self, **kwargs: Any) -> dict[str, Any]:
            return {"auth_uri": "https://login.example.test/authorize", "state": "test-state"}

        def acquire_token_by_auth_code_flow(
            self,
            flow: dict[str, Any],
            auth_response: dict[str, str],
        ) -> dict[str, Any]:
            return {
                "id_token_claims": {
                    "preferred_username": "someone-else@example.com",
                }
            }

        def get_accounts(self) -> list[dict[str, str]]:
            return []

    monkeypatch.setattr(
        "app.auth.service.msal.ConfidentialClientApplication",
        DifferentAccountClient,
    )
    client = TestClient(create_app(settings, MicrosoftAuthService(settings)))
    client.get("/api/v1/auth/login", follow_redirects=False)

    response = client.get(
        "/api/v1/auth/callback",
        params={"code": "test-code", "state": "test-state"},
    )
    assert response.status_code == 401
    assert response.json()["detail"] == "administrator_not_allowed"
    assert client.get("/api/v1/auth/status").json() == {"authenticated": False}


def test_missing_auth_flow_does_not_authenticate(tmp_path: Path) -> None:
    settings = make_settings(tmp_path)
    client = TestClient(create_app(settings, MicrosoftAuthService(settings)))

    response = client.get("/api/v1/auth/callback", params={"code": "x", "state": "x"})
    assert response.status_code == 400
    assert client.get("/api/v1/auth/status").json() == {"authenticated": False}
    assert client.get("/api/v1/auth/me").status_code == 401


def test_outlook_graph_sync_and_filter_messages(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    settings = make_settings(tmp_path)

    class FakeMsalClient:
        def __init__(self, **kwargs: Any) -> None:
            self.cache = kwargs["token_cache"]

        def initiate_auth_code_flow(self, **kwargs: Any) -> dict[str, Any]:
            return {"auth_uri": "https://login.example.test/authorize", "state": "test-state"}

        def acquire_token_by_auth_code_flow(
            self,
            flow: dict[str, Any],
            auth_response: dict[str, str],
        ) -> dict[str, Any]:
            self.cache.add(
                {
                    "client_id": "test-client",
                    "scope": ["User.Read"],
                    "environment": "login.microsoftonline.com",
                    "response": {
                        "token_type": "Bearer",
                        "access_token": "test-token",
                        "expires_in": 3600,
                    },
                }
            )
            return {
                "access_token": "test-token",
                "id_token_claims": {"preferred_username": "admin@example.com"},
            }

        def get_accounts(self) -> list[dict[str, str]]:
            return [{"home_account_id": "admin-account", "username": "admin@example.com"}]

        def acquire_token_silent(self, *args: Any, **kwargs: Any) -> dict[str, str]:
            return {"access_token": "test-token"}

    monkeypatch.setattr(
        "app.auth.service.msal.ConfidentialClientApplication",
        FakeMsalClient,
    )
    auth_service = MicrosoftAuthService(settings)
    client = TestClient(create_app(settings, auth_service))

    client.get("/api/v1/auth/login", follow_redirects=False)
    client.get(
        "/api/v1/auth/callback",
        params={"code": "test-code", "state": "test-state"},
        follow_redirects=False,
    )

    class FakeGraphResponse:
        status_code = 200
        headers = {}

        def raise_for_status(self) -> None:
            return None

        @staticmethod
        def json() -> dict[str, list[dict[str, Any]]]:
            return {
                "value": [
                    {
                        "id": "mail-1",
                        "subject": "Revisión mensual",
                        "from": {"emailAddress": {"address": "jefe@empresa.test"}},
                        "receivedDateTime": "2025-03-12T09:00:00Z",
                        "categories": ["Finanzas"],
                        "parentFolderId": "inbox",
                    },
                    {
                        "id": "mail-2",
                        "subject": "Proyecto interno",
                        "from": {"emailAddress": {"address": "otro@empresa.test"}},
                        "receivedDateTime": "2024-07-03T15:00:00Z",
                        "categories": ["Operaciones"],
                        "parentFolderId": "safebox",
                    },
                ]
            }

    def fake_graph_get(
        url: str,
        headers: dict[str, str],
        params: dict[str, Any],
        timeout: float,
    ) -> FakeGraphResponse:
        assert url == "https://graph.microsoft.com/v1.0/me/messages"
        assert headers["Authorization"].startswith("Bearer ")
        assert params["$top"] == 25
        return FakeGraphResponse()

    monkeypatch.setattr("app.outlook.service.httpx.get", fake_graph_get)

    sync_response = client.post("/api/v1/outlook/sync")
    assert sync_response.status_code == 200
    assert sync_response.json()["count"] == 2

    filtered_response = client.get(
        "/api/v1/correos",
        params={
            "categoria": "Finanzas",
            "ano": 2025,
            "remitente": "jefe",
            "page": 1,
            "page_size": 10,
        },
    )
    assert filtered_response.status_code == 200
    payload = filtered_response.json()
    assert payload["count"] == 1
    assert payload["items"][0]["subject"] == "Revisión mensual"


def test_outlook_plan_assignment_is_idempotent(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    settings = make_settings(tmp_path)

    class FakeMsalClient:
        def __init__(self, **kwargs: Any) -> None:
            self.cache = kwargs["token_cache"]

        def initiate_auth_code_flow(self, **kwargs: Any) -> dict[str, Any]:
            return {"auth_uri": "https://login.example.test/authorize", "state": "test-state"}

        def acquire_token_by_auth_code_flow(
            self,
            flow: dict[str, Any],
            auth_response: dict[str, str],
        ) -> dict[str, Any]:
            self.cache.add(
                {
                    "client_id": "test-client",
                    "scope": ["User.Read"],
                    "environment": "login.microsoftonline.com",
                    "response": {
                        "token_type": "Bearer",
                        "access_token": "test-token",
                        "expires_in": 3600,
                    },
                }
            )
            return {
                "access_token": "test-token",
                "id_token_claims": {"preferred_username": "admin@example.com"},
            }

        def get_accounts(self) -> list[dict[str, str]]:
            return [{"home_account_id": "admin-account", "username": "admin@example.com"}]

        def acquire_token_silent(self, *args: Any, **kwargs: Any) -> dict[str, str]:
            return {"access_token": "test-token"}

    monkeypatch.setattr(
        "app.auth.service.msal.ConfidentialClientApplication",
        FakeMsalClient,
    )

    fake_calls = {"count": 0}

    class FakeGraphResponse:
        def __init__(self, payload: dict[str, Any]) -> None:
            self._payload = payload
            self.status_code = 200
            self.headers = {}

        def raise_for_status(self) -> None:
            return None

        def json(self) -> dict[str, Any]:
            return self._payload

    def fake_graph_post(url: str, headers: dict[str, str], json: dict[str, Any], timeout: float) -> FakeGraphResponse:
        fake_calls["count"] += 1
        assert url == "https://graph.microsoft.com/v1.0/me/events"
        assert headers["Authorization"].startswith("Bearer ")
        return FakeGraphResponse({"id": "evt-123"})

    monkeypatch.setattr("app.outlook.service.httpx.post", fake_graph_post)

    auth_service = MicrosoftAuthService(settings)
    client = TestClient(create_app(settings, auth_service))
    client.get("/api/v1/auth/login", follow_redirects=False)
    client.get(
        "/api/v1/auth/callback",
        params={"code": "test-code", "state": "test-state"},
        follow_redirects=False,
    )

    payload = {
        "message_id": "msg-42",
        "start_date": "2025-04-01",
        "end_date": "2025-04-05",
        "time_zone": "America/Santiago",
        "subject": "Planificación de prueba",
        "body": "<p>Hola</p>",
    }

    first = client.post("/api/v1/planificacion/asignar", json=payload)
    assert first.status_code == 200
    assert first.json()["status"] == "created"

    second = client.post("/api/v1/planificacion/asignar", json=payload)
    assert second.status_code == 200
    assert second.json()["status"] == "existing"
    assert fake_calls["count"] == 1

    invalid = client.post(
        "/api/v1/planificacion/asignar",
        json={
            "message_id": "msg-99",
            "start_date": "2025-04-10",
            "end_date": "2025-04-05",
            "time_zone": "America/Santiago",
            "subject": "Inválido",
        },
    )
    assert invalid.status_code == 400


def _make_authenticated_zimbra_client(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> TestClient:
    """Helper: creates a TestClient with an active admin session."""

    class FakeMsalClient:
        def __init__(self, **kwargs: Any) -> None:
            self.cache = kwargs["token_cache"]

        def initiate_auth_code_flow(self, **kwargs: Any) -> dict[str, Any]:
            return {"auth_uri": "https://login.example.test/authorize", "state": "test-state"}

        def acquire_token_by_auth_code_flow(
            self, flow: dict[str, Any], auth_response: dict[str, str]
        ) -> dict[str, Any]:
            self.cache.add({
                "client_id": "test-client",
                "scope": ["User.Read"],
                "environment": "login.microsoftonline.com",
                "response": {
                    "token_type": "Bearer",
                    "access_token": "test-token",
                    "expires_in": 3600,
                },
            })
            return {
                "access_token": "test-token",
                "id_token_claims": {"preferred_username": "admin@example.com"},
            }

        def get_accounts(self) -> list[dict[str, str]]:
            return [{"home_account_id": "admin-account", "username": "admin@example.com"}]

        def acquire_token_silent(self, *args: Any, **kwargs: Any) -> dict[str, str]:
            return {"access_token": "test-token"}

    monkeypatch.setattr(
        "app.auth.service.msal.ConfidentialClientApplication",
        FakeMsalClient,
    )
    settings = make_settings(tmp_path)
    auth_service = MicrosoftAuthService(settings)
    client = TestClient(create_app(settings, auth_service))
    client.get("/api/v1/auth/login", follow_redirects=False)
    client.get(
        "/api/v1/auth/callback",
        params={"code": "test-code", "state": "test-state"},
        follow_redirects=False,
    )
    return client


def test_zimbra_batch_submission_and_status(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    client = _make_authenticated_zimbra_client(tmp_path, monkeypatch)

    def fake_send_message(self: Any, **kwargs: Any) -> None:
        return None

    monkeypatch.setattr("app.zimbra.service.ZimbraSMTPService._send_message", fake_send_message)

    response = client.post(
        "/api/v1/zimbra/reenvio-masivo",
        json={
            "message_ids": ["msg-1"],
            "recipients": ["destino1@test.com", "destino2@test.com"],
            "subject": "Asunto de prueba",
            "body": "Cuerpo del mensaje",
        },
    )
    assert response.status_code == 200
    payload = response.json()
    assert "batch_id" in payload
    batch_id = payload["batch_id"]

    status_response = client.get(f"/api/v1/zimbra/reenvio-masivo/{batch_id}")
    assert status_response.status_code == 200
    assert status_response.json()["status"] == "completed"


def test_zimbra_batch_requires_authentication(tmp_path: Path) -> None:
    """Unauthenticated requests to Zimbra endpoints must return 401."""
    settings = make_settings(tmp_path)
    client = TestClient(create_app(settings, MicrosoftAuthService(settings)))

    response = client.post(
        "/api/v1/zimbra/reenvio-masivo",
        json={
            "message_ids": ["msg-1"],
            "recipients": ["dest@test.com"],
            "subject": "Test",
            "body": "Test body",
        },
    )
    assert response.status_code == 401

    status_response = client.get("/api/v1/zimbra/reenvio-masivo/some-batch-id")
    assert status_response.status_code == 401


def test_requires_secure_cookie_in_production(tmp_path: Path) -> None:
    settings_values = {
        "microsoft_client_id": "test-client",
        "microsoft_client_secret": "test-client-secret",
        "microsoft_admin_email": "admin@example.com",
        "token_cache_encryption_key": Fernet.generate_key().decode("ascii"),
        "token_cache_path": tmp_path / "tokens.bin",
        "session_secret_key": "test-only-session-secret-key-at-least-32",
        "app_env": "production",
        "cookie_secure": False,
    }
    with pytest.raises(ValueError, match="COOKIE_SECURE"):
        Settings(**settings_values)


def test_local_outlook_mode_syncs_all_folders_categories_and_years(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """In OUTLOOK_MODE=local, no Entra ID credentials are required and all folders/categories/years load."""
    settings = Settings(
        outlook_mode="local",
        microsoft_admin_email="admin@empresa.com",
        token_cache_encryption_key=Fernet.generate_key().decode("ascii"),
        token_cache_path=tmp_path / "tokens.bin",
        session_secret_key="test-only-session-secret-key-at-least-32",
    )
    auth_service = MicrosoftAuthService(settings)
    monkeypatch.setattr(
        auth_service,
        "authenticate_local_outlook",
        lambda: "local:admin@empresa.com",
    )

    app = create_app(settings, auth_service)

    def fake_sync_local(
        *,
        limit_per_folder: int = 100,
        target_folder: str | None = None,
    ) -> dict[str, Any]:
        messages = [
            {
                "id": "local-1",
                "subject": "Presupuesto 2026",
                "remitente_email": "finanzas@empresa.com",
                "remitente_nombre": "Área Finanzas",
                "carpeta_origen": "Bandeja de entrada/Finanzas",
                "carpeta_nombre": "Finanzas",
                "bandeja_cuenta": "admin@empresa.com",
                "categoria": "Urgente",
                "categorias": ["Urgente", "Presupuesto"],
                "ano_fiscal": 2026,
                "received_at": "2026-02-10T10:00:00",
            },
            {
                "id": "local-2",
                "subject": "Reporte anual 2025",
                "remitente_email": "operaciones@empresa.com",
                "remitente_nombre": "Operaciones",
                "carpeta_origen": "Archivo/2025",
                "carpeta_nombre": "2025",
                "bandeja_cuenta": "admin@empresa.com",
                "categoria": "Operaciones",
                "categorias": ["Operaciones"],
                "ano_fiscal": 2025,
                "received_at": "2025-11-20T14:30:00",
            },
        ]
        app.state.outlook_service.messages = messages
        return {
            "count": len(messages),
            "messages": messages,
            "summary": app.state.outlook_service.get_filter_summary(),
        }

    monkeypatch.setattr(
        app.state.outlook_service,
        "_sync_local_outlook_messages",
        fake_sync_local,
    )

    client = TestClient(app)
    login_resp = client.get("/api/v1/auth/login", follow_redirects=False)
    assert login_resp.status_code == 303

    sync_resp = client.post("/api/v1/outlook/sync")
    assert sync_resp.status_code == 200
    assert sync_resp.json()["count"] == 2

    filtros_resp = client.get("/api/v1/outlook/filtros")
    assert filtros_resp.status_code == 200
    summary = filtros_resp.json()
    assert len(summary["carpetas"]) == 2
    assert {a["ano"] for a in summary["anos"]} == {2025, 2026}

    by_subcat = client.get("/api/v1/correos", params={"categoria": "Presupuesto", "ano": 2026})
    assert by_subcat.status_code == 200
    assert by_subcat.json()["count"] == 1
    assert by_subcat.json()["items"][0]["id"] == "local-1"
