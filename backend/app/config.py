from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Literal

from cryptography.fernet import Fernet
from pydantic import Field, SecretStr, field_validator, model_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


_BACKEND_DIR = Path(__file__).resolve().parent.parent


class Settings(BaseSettings):
    model_config = SettingsConfigDict(
        env_file=(str(_BACKEND_DIR / ".env"), ".env"),
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )

    app_env: Literal["development", "production"] = "development"
    outlook_mode: Literal["local", "graph"] | None = None

    microsoft_client_id: str = ""
    microsoft_client_secret: SecretStr = Field(default=SecretStr(""))
    microsoft_redirect_uri: str = "http://localhost:8000/api/v1/auth/callback"
    microsoft_admin_email: str = ""
    microsoft_graph_scopes: str = "User.Read Mail.Read Calendars.ReadWrite"

    token_cache_encryption_key: SecretStr
    token_cache_path: Path = Path("data/msal_cache.bin")
    session_secret_key: SecretStr
    session_max_age_seconds: int = Field(default=28_800, ge=300, le=2_592_000)
    cookie_secure: bool = False

    zimbra_host: str = "localhost"
    zimbra_port: int = Field(default=587, ge=1, le=65535)
    zimbra_username: str = ""
    zimbra_password: SecretStr | None = None
    zimbra_from_email: str = "noreply@example.com"
    zimbra_use_tls: bool = True
    zimbra_retry_count: int = Field(default=3, ge=1, le=10)
    zimbra_retry_delay_seconds: int = Field(default=5, ge=0, le=300)

    frontend_origin: str = "http://localhost:8100"
    frontend_redirect_url: str = "http://localhost:8100/"

    @field_validator("token_cache_encryption_key")
    @classmethod
    def validate_fernet_key(cls, value: SecretStr) -> SecretStr:
        try:
            Fernet(value.get_secret_value().encode("ascii"))
        except (ValueError, UnicodeEncodeError) as exc:
            raise ValueError("TOKEN_CACHE_ENCRYPTION_KEY must be a valid Fernet key") from exc
        return value

    @field_validator("session_secret_key")
    @classmethod
    def validate_session_secret(cls, value: SecretStr) -> SecretStr:
        if len(value.get_secret_value()) < 32:
            raise ValueError("SESSION_SECRET_KEY must contain at least 32 characters")
        return value

    @field_validator("zimbra_from_email")
    @classmethod
    def validate_zimbra_from_email(cls, value: str) -> str:
        if "@" not in value:
            raise ValueError("ZIMBRA_FROM_EMAIL must be a valid email address")
        return value

    @field_validator("microsoft_redirect_uri", "frontend_origin", "frontend_redirect_url")
    @classmethod
    def validate_http_url(cls, value: str) -> str:
        from urllib.parse import urlsplit

        parsed = urlsplit(value)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            raise ValueError("Application URLs must be absolute HTTP or HTTPS URLs")
        return value

    @field_validator("microsoft_graph_scopes")
    @classmethod
    def validate_graph_scopes(cls, value: str) -> str:
        scopes = [scope.strip() for scope in value.split()]
        if not scopes or any(not scope for scope in scopes):
            raise ValueError("MICROSOFT_GRAPH_SCOPES must contain at least one scope")
        return " ".join(scopes)

    @model_validator(mode="after")
    def validate_mode_and_production_settings(self) -> Settings:
        if self.outlook_mode is None:
            self.outlook_mode = "graph" if self.microsoft_client_id.strip() else "local"

        if self.outlook_mode == "graph":
            if not self.microsoft_client_id.strip():
                raise ValueError("MICROSOFT_CLIENT_ID is required when OUTLOOK_MODE=graph")
            if not self.microsoft_client_secret.get_secret_value().strip():
                raise ValueError("MICROSOFT_CLIENT_SECRET is required when OUTLOOK_MODE=graph")
            if len(self.microsoft_admin_email.strip()) < 3:
                raise ValueError("MICROSOFT_ADMIN_EMAIL is required when OUTLOOK_MODE=graph")

        if self.app_env == "production" and not self.cookie_secure:
            raise ValueError("COOKIE_SECURE must be true in production")
        if self.app_env == "production":
            urls = (
                self.microsoft_redirect_uri,
                self.frontend_origin,
                self.frontend_redirect_url,
            )
            if any(not url.startswith("https://") for url in urls):
                raise ValueError("Application URLs must use HTTPS in production")
        return self

    @property
    def graph_scopes(self) -> list[str]:
        return self.microsoft_graph_scopes.split()


@lru_cache
def get_settings() -> Settings:
    return Settings()
