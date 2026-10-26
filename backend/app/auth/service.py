from __future__ import annotations

from typing import Any

import msal

from app.auth.cache import EncryptedTokenCache
from app.config import Settings


class AuthenticationError(RuntimeError):
    def __init__(self, code: str) -> None:
        self.code = code
        super().__init__(code)


class MicrosoftAuthService:
    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.token_cache = EncryptedTokenCache(
            settings.token_cache_path,
            settings.token_cache_encryption_key.get_secret_value(),
        )

    def authenticate_local_outlook(self) -> str:
        """Verify local Outlook MAPI session on Windows without Entra ID."""
        admin_email = self.settings.microsoft_admin_email.strip().casefold()
        try:
            import pythoncom  # type: ignore[import-not-found,import-untyped]
            import win32com.client  # type: ignore[import-not-found,import-untyped]
        except ImportError as exc:
            raise AuthenticationError("pywin32_not_installed") from exc

        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as exc:
                raise AuthenticationError("local_outlook_unavailable") from exc

            matched_email: str | None = None
            try:
                accounts = namespace.Accounts
                account_count = int(accounts.Count)
                discovered_emails: list[str] = []
                for idx in range(1, account_count + 1):
                    account = accounts.Item(idx)
                    smtp = getattr(account, "SmtpAddress", None) or getattr(
                        account, "DisplayName", ""
                    )
                    if isinstance(smtp, str) and smtp.strip():
                        discovered_emails.append(smtp.strip().casefold())
                        if smtp.strip().casefold() == admin_email:
                            matched_email = admin_email
                            break
                # Also inspect Store display names (often the email address in Outlook)
                store_names: list[str] = []
                try:
                    stores = namespace.Stores
                    for idx in range(1, int(stores.Count) + 1):
                        store_name = str(getattr(stores.Item(idx), "DisplayName", "")).strip()
                        if store_name:
                            store_names.append(store_name)
                except Exception:
                    pass

                # If admin_email was left blank in .env, auto-detect the primary Outlook account
                if not admin_email:
                    if discovered_emails:
                        matched_email = discovered_emails[0]
                    elif store_names:
                        matched_email = store_names[0]
                    else:
                        current_user = getattr(namespace, "CurrentUser", None)
                        matched_email = (
                            str(getattr(current_user, "Name", "") or "").strip()
                            or "outlook-local"
                        )
                    self.settings.microsoft_admin_email = matched_email
                elif matched_email is None and not discovered_emails:
                    matched_email = admin_email
                elif matched_email is None and discovered_emails:
                    for store_name in store_names:
                        if admin_email in store_name.casefold() or store_name.casefold() in admin_email:
                            matched_email = admin_email
                            break
            except Exception:
                matched_email = admin_email

            if matched_email is None:
                raise AuthenticationError("administrator_not_allowed")

            return f"local:{matched_email}"
        finally:
            pythoncom.CoUninitialize()

    def create_auth_flow(self) -> dict[str, Any]:
        with self.token_cache.session() as cache:
            client = self._create_client(cache)
            flow = client.initiate_auth_code_flow(
                scopes=self.settings.graph_scopes,
                redirect_uri=self.settings.microsoft_redirect_uri,
                prompt="select_account",
            )
        if not isinstance(flow, dict) or "auth_uri" not in flow:
            raise AuthenticationError("authorization_start_failed")
        return flow

    def complete_auth_flow(
        self,
        flow: dict[str, Any],
        auth_response: dict[str, str],
    ) -> str:
        with self.token_cache.session() as cache:
            client = self._create_client(cache)
            try:
                result = client.acquire_token_by_auth_code_flow(flow, auth_response)
            except ValueError as exc:
                raise AuthenticationError("authorization_state_invalid") from exc

            if not isinstance(result, dict) or result.get("error"):
                raise AuthenticationError("authorization_failed")

            claims = result.get("id_token_claims")
            if not isinstance(claims, dict):
                raise AuthenticationError("identity_claims_missing")

            email = (
                claims.get("preferred_username")
                or claims.get("upn")
                or claims.get("email")
                or ""
            )
            if (
                not isinstance(email, str)
                or email.casefold() != self.settings.microsoft_admin_email.casefold()
            ):
                raise AuthenticationError("administrator_not_allowed")

            accounts = client.get_accounts()
            account = next(
                (
                    item
                    for item in accounts
                    if item.get("username", "").casefold() == email.casefold()
                    and item.get("home_account_id")
                ),
                None,
            )
            if account is None:
                raise AuthenticationError("account_not_cached")
            return account["home_account_id"]

    def acquire_graph_access_token(self, account_id: str) -> str | None:
        if self.settings.outlook_mode == "local" or account_id.startswith("local:"):
            return "local-outlook-mapi"

        with self.token_cache.session() as cache:
            client = self._create_client(cache)
            account = next(
                (
                    item
                    for item in client.get_accounts()
                    if item.get("home_account_id") == account_id
                ),
                None,
            )
            if account is None:
                return None

            result = client.acquire_token_silent(self.settings.graph_scopes, account)
            if not isinstance(result, dict):
                return None
            access_token = result.get("access_token")
            return access_token if isinstance(access_token, str) and access_token else None

    def _create_client(self, cache: Any) -> msal.ConfidentialClientApplication:
        return msal.ConfidentialClientApplication(
            client_id=self.settings.microsoft_client_id,
            client_credential=self.settings.microsoft_client_secret.get_secret_value(),
            authority="https://login.microsoftonline.com/consumers",
            token_cache=cache,
        )
