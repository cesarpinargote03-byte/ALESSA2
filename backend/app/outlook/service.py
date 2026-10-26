from __future__ import annotations

import html
import io
import os
import re
import tempfile
import uuid
import zipfile
from collections import Counter
from datetime import date, datetime, time
from typing import Any
from zoneinfo import ZoneInfo

import httpx

# Folders generated internally by Outlook/Exchange or trash/junk that should not be mixed with user mail
_IGNORED_FOLDER_NAMES = frozenset(
    {
        "sync issues",
        "problemas de sincronización",
        "conflicts",
        "conflictos",
        "local failures",
        "errores locales",
        "server failures",
        "errores del servidor",
        "conversation action settings",
        "configuración de acciones de conversación",
        "quick step settings",
        "configuración de pasos rápidos",
        "rss feeds",
        "fuentes rss",
        "externalcontacts",
        "files",
        "archivos",
        "deleted items",
        "elementos eliminados",
        "junk email",
        "correo electrónico no deseado",
        "correo no deseado",
        "outbox",
        "bandeja de salida",
        "drafts",
        "borradores",
        "conversation history",
        "historial de conversaciones",
        "yammer root",
    }
)

# Cached SMTP address on the message header itself (no network call to Active Directory!)
_PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"


class OutlookGraphError(RuntimeError):
    def __init__(self, message: str, *, status_code: int | None = None) -> None:
        self.message = message
        self.status_code = status_code
        super().__init__(message)


class OutlookGraphService:
    MAX_REQUESTS_PER_WEEK = 4

    def __init__(self, admin_email: str | None = None) -> None:
        self.admin_email = (admin_email or "").strip().casefold()
        self.messages: list[dict[str, Any]] = []
        self.planned_events: dict[str, dict[str, Any]] = {}
        self.planned_by_message: dict[str, dict[str, Any]] = {}
        self.sent_batches: list[dict[str, Any]] = []

    def sync_messages(
        self,
        access_token: str,
        *,
        limit: int = 2000,
        folder: str | None = None,
    ) -> dict[str, Any]:
        if access_token == "local-outlook-mapi":
            return self._sync_local_outlook_messages(limit_per_folder=limit, target_folder=folder)

        params = {
            "$select": "id,subject,from,receivedDateTime,parentFolderId,categories",
            "$orderby": "receivedDateTime desc",
            "$top": max(1, min(int(limit), 100)),
        }
        url = "https://graph.microsoft.com/v1.0/me/messages"
        headers = {"Authorization": f"Bearer {access_token}"}

        try:
            response = httpx.get(url, headers=headers, params=params, timeout=30.0)
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            retry_after = exc.response.headers.get("Retry-After") if exc.response else None
            if exc.response is not None and exc.response.status_code == 429:
                raise OutlookGraphError(
                    f"graph_rate_limited_retry_after_{retry_after or 0}s",
                    status_code=429,
                ) from exc
            raise OutlookGraphError(
                "graph_request_failed",
                status_code=exc.response.status_code if exc.response else None,
            ) from exc
        except httpx.RequestError as exc:
            raise OutlookGraphError("graph_unavailable") from exc

        payload = response.json()
        values = payload.get("value", [])
        normalized: list[dict[str, Any]] = []
        for item in values:
            sender = item.get("from") or {}
            email_address = (sender.get("emailAddress") or {}).get("address") or ""
            received_at = item.get("receivedDateTime")
            categories = [
                str(cat).strip()
                for cat in (item.get("categories") or [])
                if isinstance(cat, str) and str(cat).strip()
            ]
            category = categories[0] if categories else None
            folder_path = folder or item.get("parentFolderId") or "Inbox"
            normalized.append(
                {
                    "id": item.get("id"),
                    "subject": item.get("subject") or "",
                    "remitente_email": email_address,
                    "carpeta_origen": folder_path,
                    "carpeta_nombre": folder_path.split("/")[-1],
                    "categoria": category,
                    "categorias": categories,
                    "ano_fiscal": (
                        int(received_at[:4])
                        if isinstance(received_at, str) and len(received_at) >= 4 and received_at[:4].isdigit()
                        else None
                    ),
                    "received_at": received_at,
                    "raw": item,
                }
            )

        self._apply_planned_metadata(normalized)
        self.messages = normalized
        return {
            "count": len(normalized),
            "messages": normalized,
            "summary": self.get_filter_summary(),
        }

    def _sync_local_outlook_messages(
        self,
        *,
        limit_per_folder: int = 2000,
        target_folder: str | None = None,
    ) -> dict[str, Any]:
        """Ultra-fast local Outlook MAPI sync using Folder.GetTable().GetArray() batching."""
        try:
            import pythoncom  # type: ignore[import-not-found,import-untyped]
            import win32com.client  # type: ignore[import-not-found,import-untyped]
        except ImportError as exc:
            raise OutlookGraphError("pywin32_not_installed", status_code=503) from exc

        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
            except Exception as exc:
                raise OutlookGraphError("local_outlook_unavailable", status_code=503) from exc

            roots = self._select_mailbox_roots(namespace)
            normalized: list[dict[str, Any]] = []
            seen_ids: set[str] = set()
            target_cf = target_folder.strip().casefold() if target_folder else None
            max_per_folder = max(1, min(int(limit_per_folder), 5000))
            max_total = 10000

            for store_name, root_folder in roots:
                for folder_path, folder_name, outlook_folder in self._walk_mail_folders(root_folder):
                    if len(normalized) >= max_total:
                        break

                    if target_cf and (
                        target_cf != folder_path.casefold()
                        and target_cf != folder_name.casefold()
                        and not folder_path.casefold().startswith(f"{target_cf}/")
                    ):
                        continue

                    folder_msgs = self._read_folder_fast(
                        outlook_folder,
                        folder_path=folder_path,
                        folder_name=folder_name,
                        store_name=store_name,
                        max_items=min(max_per_folder, max_total - len(normalized)),
                        seen_ids=seen_ids,
                    )
                    normalized.extend(folder_msgs)

            normalized.sort(
                key=lambda msg: str(msg.get("received_at") or ""),
                reverse=True,
            )
            self._apply_planned_metadata(normalized)
            self.messages = normalized
            return {
                "count": len(normalized),
                "messages": normalized,
                "summary": self.get_filter_summary(),
            }
        finally:
            pythoncom.CoUninitialize()

    def _read_folder_fast(
        self,
        outlook_folder: Any,
        *,
        folder_path: str,
        folder_name: str,
        store_name: str,
        max_items: int,
        seen_ids: set[str],
    ) -> list[dict[str, Any]]:
        """Read all mail items in a folder using MAPI GetTable().GetArray() (1 COM call per folder)."""
        results: list[dict[str, Any]] = []
        try:
            table = outlook_folder.GetTable()
            table.Columns.RemoveAll()
            table.Columns.Add("EntryID")
            table.Columns.Add("Subject")
            table.Columns.Add("ReceivedTime")
            table.Columns.Add("Categories")
            table.Columns.Add("SenderName")
            table.Columns.Add("SenderEmailAddress")
            table.Columns.Add("MessageClass")
            table.Columns.Add(_PR_SENDER_SMTP_ADDRESS)
            try:
                table.Sort("[ReceivedTime]", True)
            except Exception:
                pass

            if table.EndOfTable:
                return results

            rows = table.GetArray(max_items)
            for row in rows:
                try:
                    entry_id = str(row[0] or "")
                    if not entry_id or entry_id in seen_ids:
                        continue
                    msg_class = str(row[6] or "")
                    if msg_class and not msg_class.upper().startswith("IPM.NOTE"):
                        continue

                    seen_ids.add(entry_id)
                    subject = str(row[1] or "")
                    received_time = row[2]
                    raw_categories = str(row[3] or "").strip()
                    sender_name = str(row[4] or "").strip()
                    sender_addr = str(row[5] or "").strip()
                    smtp_prop = str(row[7] or "").strip() if len(row) > 7 else ""

                    if smtp_prop and "@" in smtp_prop:
                        sender_email = smtp_prop
                    elif sender_addr and not sender_addr.startswith("/O=") and "@" in sender_addr:
                        sender_email = sender_addr
                    else:
                        sender_email = sender_name or sender_addr or "desconocido@outlook"

                    categories = (
                        [cat.strip() for cat in re.split(r"[;,]", raw_categories) if cat.strip()]
                        if raw_categories
                        else []
                    )
                    primary_category = categories[0] if categories else None

                    received_iso: str | None = None
                    fiscal_year: int | None = None
                    if received_time is not None:
                        try:
                            fiscal_year = int(received_time.year)
                            received_iso = datetime(
                                received_time.year,
                                received_time.month,
                                received_time.day,
                                received_time.hour,
                                received_time.minute,
                                received_time.second,
                            ).isoformat()
                        except Exception:
                            received_iso = str(received_time)

                    results.append(
                        {
                            "id": entry_id,
                            "subject": subject,
                            "remitente_email": sender_email,
                            "remitente_nombre": sender_name,
                            "carpeta_origen": folder_path,
                            "carpeta_nombre": folder_name,
                            "bandeja_cuenta": store_name,
                            "categoria": primary_category,
                            "categorias": categories,
                            "ano_fiscal": fiscal_year,
                            "received_at": received_iso,
                        }
                    )
                except Exception:
                    continue
            return results
        except Exception:
            # Fallback to item-by-item reading if GetTable is unsupported on a special folder
            try:
                items = outlook_folder.Items
                total_in_folder = int(items.Count)
                if total_in_folder == 0:
                    return results
                items.Sort("[ReceivedTime]", True)
                to_read = min(total_in_folder, max_items)
                for idx in range(1, to_read + 1):
                    try:
                        item = items.Item(idx)
                        if getattr(item, "Class", None) != 43:
                            continue
                        entry_id = str(getattr(item, "EntryID", "") or "")
                        if not entry_id or entry_id in seen_ids:
                            continue
                        seen_ids.add(entry_id)
                        results.append(
                            self._normalize_local_mail_item(
                                item,
                                folder_path=folder_path,
                                folder_name=folder_name,
                                store_name=store_name,
                            )
                        )
                    except Exception:
                        continue
            except Exception:
                pass
            return results

    def _select_mailbox_roots(self, namespace: Any) -> list[tuple[str, Any]]:
        """Return (store_name, root_folder) for all mail stores (primary + archives)."""
        all_stores: list[tuple[str, Any]] = []
        seen_roots: set[str] = set()

        try:
            stores = namespace.Stores
            for idx in range(1, int(stores.Count) + 1):
                try:
                    store = stores.Item(idx)
                    display_name = str(getattr(store, "DisplayName", f"Buzón {idx}"))
                    root_folder = store.GetRootFolder()
                    root_id = str(getattr(root_folder, "EntryID", display_name))
                    if root_id not in seen_roots:
                        seen_roots.add(root_id)
                        all_stores.append((display_name, root_folder))
                except Exception:
                    continue
        except Exception:
            pass

        if all_stores:
            return all_stores

        try:
            inbox = namespace.GetDefaultFolder(6)
            parent = getattr(inbox, "Parent", inbox)
            store_name = str(getattr(parent, "Name", "") or self.admin_email or "Outlook")
            return [(store_name, parent)]
        except Exception:
            return []

    def _walk_mail_folders(
        self,
        parent_folder: Any,
        prefix: str = "",
        depth: int = 0,
    ) -> list[tuple[str, str, Any]]:
        """Recursively walk mail folders and subfolders up to depth 5."""
        if depth > 5:
            return []

        discovered: list[tuple[str, str, Any]] = []
        try:
            subfolders = parent_folder.Folders
            count = int(subfolders.Count)
        except Exception:
            return discovered

        for idx in range(1, count + 1):
            try:
                folder = subfolders.Item(idx)
                name = str(getattr(folder, "Name", "") or "").strip()
                if not name or name.casefold() in _IGNORED_FOLDER_NAMES:
                    continue

                default_item_type = getattr(folder, "DefaultItemType", 0)
                folder_path = f"{prefix}/{name}" if prefix else name

                if default_item_type == 0:
                    discovered.append((folder_path, name, folder))
                    discovered.extend(
                        self._walk_mail_folders(folder, prefix=folder_path, depth=depth + 1)
                    )
            except Exception:
                continue

        return discovered

    @staticmethod
    def _extract_sender_email(item: Any) -> str:
        """Extract sender email purely from local MAPI cache (0ms, no AD network lookup)."""
        try:
            address = str(getattr(item, "SenderEmailAddress", "") or "").strip()
            if address and not address.startswith("/O=") and "@" in address:
                return address

            prop_accessor = getattr(item, "PropertyAccessor", None)
            if prop_accessor is not None:
                try:
                    smtp = prop_accessor.GetProperty(_PR_SENDER_SMTP_ADDRESS)
                    if smtp and "@" in str(smtp):
                        return str(smtp).strip()
                except Exception:
                    pass

            sender_name = str(getattr(item, "SenderName", "") or "").strip()
            return sender_name or address or "desconocido@outlook"
        except Exception:
            return "desconocido@outlook"

    @classmethod
    def _normalize_local_mail_item(
        cls,
        item: Any,
        *,
        folder_path: str,
        folder_name: str,
        store_name: str,
    ) -> dict[str, Any]:
        raw_categories = str(getattr(item, "Categories", "") or "").strip()
        categories = (
            [cat.strip() for cat in re.split(r"[;,]", raw_categories) if cat.strip()]
            if raw_categories
            else []
        )
        primary_category = categories[0] if categories else None

        received_time = getattr(item, "ReceivedTime", None)
        received_iso: str | None = None
        fiscal_year: int | None = None
        if received_time is not None:
            try:
                fiscal_year = int(received_time.year)
                received_iso = datetime(
                    received_time.year,
                    received_time.month,
                    received_time.day,
                    received_time.hour,
                    received_time.minute,
                    received_time.second,
                ).isoformat()
            except Exception:
                received_iso = str(received_time)

        sender_email = cls._extract_sender_email(item)
        sender_name = str(getattr(item, "SenderName", "") or "").strip()
        subject = str(getattr(item, "Subject", "") or "")
        entry_id = str(getattr(item, "EntryID", "") or "")

        return {
            "id": entry_id,
            "subject": subject,
            "remitente_email": sender_email,
            "remitente_nombre": sender_name,
            "carpeta_origen": folder_path,
            "carpeta_nombre": folder_name,
            "bandeja_cuenta": store_name,
            "categoria": primary_category,
            "categorias": categories,
            "ano_fiscal": fiscal_year,
            "received_at": received_iso,
        }

    def get_filter_summary(self) -> dict[str, Any]:
        """Return available folders, categories, and years across all synced messages."""
        folder_counts: Counter[str] = Counter()
        category_counts: Counter[str] = Counter()
        year_counts: Counter[int] = Counter()

        for msg in self.messages:
            folder = msg.get("carpeta_origen")
            if isinstance(folder, str) and folder:
                folder_counts[folder] += 1

            cats = msg.get("categorias") or (
                [msg["categoria"]] if msg.get("categoria") else []
            )
            for cat in cats:
                if isinstance(cat, str) and cat:
                    category_counts[cat] += 1

            year = msg.get("ano_fiscal")
            if isinstance(year, int):
                year_counts[year] += 1

        return {
            "total_messages": len(self.messages),
            "carpetas": [
                {"carpeta": name, "count": count}
                for name, count in sorted(folder_counts.items())
            ],
            "categorias": [
                {"categoria": name, "count": count}
                for name, count in sorted(category_counts.items())
            ],
            "anos": [
                {"ano": year, "count": count}
                for year, count in sorted(year_counts.items(), reverse=True)
            ],
        }

    def list_messages(
        self,
        *,
        categoria: str | None = None,
        ano: int | None = None,
        folder: str | None = None,
        remitente: str | None = None,
        page: int = 1,
        page_size: int = 25,
    ) -> dict[str, Any]:
        filtered = list(self.messages)

        if categoria:
            cat_cf = categoria.strip().casefold()
            filtered = [
                item
                for item in filtered
                if (item.get("categoria") or "").casefold() == cat_cf
                or any(
                    str(c).casefold() == cat_cf
                    for c in (item.get("categorias") or [])
                )
            ]
        if ano is not None:
            filtered = [item for item in filtered if item.get("ano_fiscal") == ano]
        if folder:
            folder_cf = folder.strip().casefold()
            filtered = [
                item
                for item in filtered
                if (item.get("carpeta_origen") or "").casefold() == folder_cf
                or (item.get("carpeta_nombre") or "").casefold() == folder_cf
                or (item.get("carpeta_origen") or "").casefold().startswith(f"{folder_cf}/")
            ]
        if remitente:
            rem_cf = remitente.strip().casefold()
            filtered = [
                item
                for item in filtered
                if rem_cf in (item.get("remitente_email") or "").casefold()
                or rem_cf in (item.get("remitente_nombre") or "").casefold()
            ]

        page_number = max(1, int(page))
        page_size_number = max(1, min(int(page_size), 10000))
        start = (page_number - 1) * page_size_number
        end = start + page_size_number
        items = filtered[start:end]

        return {
            "page": page_number,
            "page_size": page_size_number,
            "count": len(filtered),
            "items": items,
        }

    def create_event_for_message(
        self,
        access_token: str,
        *,
        message_id: str,
        start_date: str | date,
        end_date: str | date,
        time_zone: str,
        subject: str,
        body: str | None = None,
        postergado_por_proyecto: str = "No",
    ) -> dict[str, Any]:
        if not message_id:
            raise OutlookGraphError("message_id_required")

        start = self._parse_date(start_date)
        end = self._parse_date(end_date)
        if end < start:
            raise OutlookGraphError("end_date_before_start_date")

        try:
            tzinfo = ZoneInfo(time_zone)
        except Exception as exc:  # pragma: no cover - zone validation path
            raise OutlookGraphError("invalid_time_zone") from exc

        iso_year, iso_week, _ = start.isocalendar()
        semana_iso = f"{iso_year}-W{iso_week:02d}"

        # Enforce maximum of 4 planned requests per ISO week (excluding the same message if being updated)
        existing_in_week = [
            mid
            for mid, plan in self.planned_by_message.items()
            if plan.get("semana_iso") == semana_iso and mid != message_id
        ]
        if len(existing_in_week) >= self.MAX_REQUESTS_PER_WEEK:
            raise OutlookGraphError(
                f"La Semana {iso_week} ya alcanzó el límite de {self.MAX_REQUESTS_PER_WEEK} solicitudes planificadas.",
                status_code=409,
            )

        event_key = f"{message_id}:{start.isoformat()}:{end.isoformat()}"
        existing = self.planned_events.get(event_key)
        if existing is not None:
            result = {
                "message_id": message_id,
                "event_id": existing["event_id"],
                "status": "existing",
                "notification_status": "pending",
                "time_zone": time_zone,
                "start_date": start.isoformat(),
                "end_date": end.isoformat(),
                "semana_iso": semana_iso,
                "semana_numero": iso_week,
                "postergado_por_proyecto": postergado_por_proyecto or "No",
            }
            self._record_message_plan(
                message_id,
                result,
                subject=subject,
                postergado_por_proyecto=postergado_por_proyecto,
            )
            return result

        event_subject = subject.strip() or f"Planificación para {message_id}"
        html_body = body or "<p>Se programó la solicitud para revisión.</p>"

        if access_token == "local-outlook-mapi":
            event_id = self._create_local_calendar_event(
                subject=event_subject,
                body=html_body,
                start=start,
                end=end,
            )
            result = {
                "message_id": message_id,
                "event_id": event_id,
                "status": "created",
                "notification_status": "pending",
                "time_zone": time_zone,
                "start_date": start.isoformat(),
                "end_date": end.isoformat(),
                "semana_iso": semana_iso,
                "semana_numero": iso_week,
                "postergado_por_proyecto": postergado_por_proyecto or "No",
            }
            self.planned_events[event_key] = {"event_id": event_id, "start": start, "end": end}
            self._record_message_plan(
                message_id,
                result,
                subject=subject,
                postergado_por_proyecto=postergado_por_proyecto,
            )
            return result

        start_dt = datetime.combine(start, time(9, 0), tzinfo=tzinfo)
        end_dt = datetime.combine(end, time(17, 0), tzinfo=tzinfo)

        payload = {
            "subject": event_subject,
            "body": {"contentType": "HTML", "content": html_body},
            "start": {"dateTime": start_dt.isoformat(), "timeZone": time_zone},
            "end": {"dateTime": end_dt.isoformat(), "timeZone": time_zone},
        }

        try:
            response = httpx.post(
                "https://graph.microsoft.com/v1.0/me/events",
                headers={
                    "Authorization": f"Bearer {access_token}",
                    "Content-Type": "application/json",
                },
                json=payload,
                timeout=30.0,
            )
            response.raise_for_status()
        except httpx.HTTPStatusError as exc:
            if exc.response is not None and exc.response.status_code == 409:
                raise OutlookGraphError("calendar_conflict_detected", status_code=409) from exc
            raise OutlookGraphError(
                "graph_event_creation_failed",
                status_code=exc.response.status_code if exc.response else None,
            ) from exc
        except httpx.RequestError as exc:
            raise OutlookGraphError("graph_unavailable") from exc

        event = response.json()
        event_id = event.get("id")
        if not isinstance(event_id, str) or not event_id:
            raise OutlookGraphError("event_id_missing")

        result = {
            "message_id": message_id,
            "event_id": event_id,
            "status": "created",
            "notification_status": "pending",
            "time_zone": time_zone,
            "start_date": start.isoformat(),
            "end_date": end.isoformat(),
            "semana_iso": semana_iso,
            "semana_numero": iso_week,
            "postergado_por_proyecto": postergado_por_proyecto or "No",
        }
        self.planned_events[event_key] = {"event_id": event_id, "start": start, "end": end}
        self._record_message_plan(
            message_id,
            result,
            subject=subject,
            postergado_por_proyecto=postergado_por_proyecto,
        )
        return result

    def _record_message_plan(
        self,
        message_id: str,
        plan_result: dict[str, Any],
        *,
        subject: str,
        postergado_por_proyecto: str = "No",
    ) -> None:
        plan_info = {
            "message_id": message_id,
            "subject": subject,
            "semana_planificada_inicio": plan_result["start_date"],
            "semana_planificada_fin": plan_result["end_date"],
            "semana_iso": plan_result["semana_iso"],
            "semana_numero": plan_result["semana_numero"],
            "outlook_event_id": plan_result["event_id"],
            "estado_planificacion": "Planificado",
            "postergado_por_proyecto": (postergado_por_proyecto or "No").strip() or "No",
        }
        self.planned_by_message[message_id] = plan_info
        for msg in self.messages:
            if msg.get("id") == message_id:
                msg.update(plan_info)

    def _apply_planned_metadata(self, messages: list[dict[str, Any]]) -> None:
        for msg in messages:
            msg_id = str(msg.get("id") or "")
            if msg_id and msg_id in self.planned_by_message:
                msg.update(self.planned_by_message[msg_id])

    def list_planned_messages(self) -> list[dict[str, Any]]:
        planned: list[dict[str, Any]] = []
        seen: set[str] = set()
        for msg in self.messages:
            msg_id = str(msg.get("id") or "")
            if msg.get("estado_planificacion") == "Planificado" and msg_id not in seen:
                seen.add(msg_id)
                planned.append(msg)
        for msg_id, plan in self.planned_by_message.items():
            if msg_id not in seen:
                planned.append(plan)
        planned.sort(key=lambda item: str(item.get("semana_planificada_inicio") or ""))
        return planned

    def list_sent_batches(self) -> list[dict[str, Any]]:
        return list(reversed(self.sent_batches))

    @staticmethod
    def build_planning_excel_bytes(rows: list[dict[str, Any]]) -> bytes:
        """
        Generate a genuine .xlsx workbook matching the Helpdesk layout:
        - A1:B3 merged -> 'LOGO HELPDESK' (bordered, centered, bold)
        - Row 4 empty
        - Row 5 headers (A5:E5): SEMANA | POSTERGADO POR PROYECTO | USUARIO | CORREO | ASUNTO
        - Rows 6+: Planned items
        """
        def xml_escape(val: Any) -> str:
            return html.escape(str(val if val is not None else ""), quote=True)

        sheet_rows_xml: list[str] = []

        # Rows 1-3: Merged A1:B3 box for LOGO HELPDESK (style s="1" = bold centered with border)
        sheet_rows_xml.append(
            '<row r="1" ht="22" customHeight="1">'
            '<c r="A1" t="inlineStr" s="1"><is><t>LOGO HELPDESK</t></is></c>'
            '<c r="B1" s="1"/>'
            "</row>"
        )
        sheet_rows_xml.append(
            '<row r="2" ht="22" customHeight="1">'
            '<c r="A2" s="1"/><c r="B2" s="1"/>'
            "</row>"
        )
        sheet_rows_xml.append(
            '<row r="3" ht="22" customHeight="1">'
            '<c r="A3" s="1"/><c r="B3" s="1"/>'
            "</row>"
        )
        # Row 4: Spacer row
        sheet_rows_xml.append('<row r="4" ht="16" customHeight="1"/>')

        # Row 5: Table Headers (style s="2" = bold header with border)
        headers = [
            ("A", "SEMANA"),
            ("B", "POSTERGADO POR PROYECTO"),
            ("C", "USUARIO"),
            ("D", "CORREO"),
            ("E", "ASUNTO"),
        ]
        header_cells = "".join(
            f'<c r="{col}5" t="inlineStr" s="2"><is><t>{xml_escape(text)}</t></is></c>'
            for col, text in headers
        )
        sheet_rows_xml.append(f'<row r="5" ht="22" customHeight="1">{header_cells}</row>')

        # Data rows starting at row 6 (style s="3" = regular cell with border)
        for idx, item in enumerate(rows, start=6):
            semana_val = xml_escape(item.get("semana", ""))
            postergado_val = xml_escape(item.get("postergado_por_proyecto", "No"))
            usuario_val = xml_escape(item.get("usuario", ""))
            correo_val = xml_escape(item.get("correo", ""))
            asunto_val = xml_escape(item.get("asunto", ""))
            row_xml = (
                f'<row r="{idx}" ht="20" customHeight="1">'
                f'<c r="A{idx}" t="inlineStr" s="3"><is><t>{semana_val}</t></is></c>'
                f'<c r="B{idx}" t="inlineStr" s="3"><is><t>{postergado_val}</t></is></c>'
                f'<c r="C{idx}" t="inlineStr" s="3"><is><t>{usuario_val}</t></is></c>'
                f'<c r="D{idx}" t="inlineStr" s="3"><is><t>{correo_val}</t></is></c>'
                f'<c r="E{idx}" t="inlineStr" s="3"><is><t>{asunto_val}</t></is></c>'
                "</row>"
            )
            sheet_rows_xml.append(row_xml)

        sheet_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<cols>"
            '<col min="1" max="1" width="24" customWidth="1"/>'
            '<col min="2" max="2" width="28" customWidth="1"/>'
            '<col min="3" max="3" width="24" customWidth="1"/>'
            '<col min="4" max="4" width="32" customWidth="1"/>'
            '<col min="5" max="5" width="48" customWidth="1"/>'
            "</cols>"
            f"<sheetData>{''.join(sheet_rows_xml)}</sheetData>"
            '<mergeCells count="1"><mergeCell ref="A1:B3"/></mergeCells>'
            "</worksheet>"
        )

        styles_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="3">'
            '<font><sz val="11"/><name val="Calibri"/></font>'
            '<font><b/><sz val="12"/><name val="Calibri"/></font>'
            '<font><b/><sz val="11"/><name val="Calibri"/></font>'
            "</fonts>"
            '<fills count="3">'
            '<fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill>'
            '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F4F0"/><bgColor indexed="64"/></patternFill></fill>'
            "</fills>"
            '<borders count="2">'
            "<border><left/><right/><top/><bottom/><diagonal/></border>"
            '<border>'
            '<left style="thin"><color auto="1"/></left>'
            '<right style="thin"><color auto="1"/></right>'
            '<top style="thin"><color auto="1"/></top>'
            '<bottom style="thin"><color auto="1"/></bottom>'
            "<diagonal/>"
            "</border>"
            "</borders>"
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="4">'
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">'
            '<alignment horizontal="center" vertical="center"/>'
            "</xf>"
            '<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
            '<alignment horizontal="left" vertical="center"/>'
            "</xf>"
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1">'
            '<alignment horizontal="left" vertical="center" wrapText="1"/>'
            "</xf>"
            "</cellXfs>"
            "</styleSheet>"
        )

        workbook_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Planificacion Helpdesk" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>"
        )

        workbook_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>"
        )

        root_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        )

        content_types_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>"
        )

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", content_types_xml)
            zf.writestr("_rels/.rels", root_rels_xml)
            zf.writestr("xl/workbook.xml", workbook_xml)
            zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
            zf.writestr("xl/styles.xml", styles_xml)
            zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        return buf.getvalue()

    @classmethod
    def build_planning_email_html(
        cls,
        *,
        intro_text: str,
        rows: list[dict[str, Any]],
    ) -> str:
        """Build HTML body containing the intro text and the Helpdesk Excel table."""
        def esc(val: Any) -> str:
            return html.escape(str(val if val is not None else ""))

        body_rows_html = "".join(
            f"<tr>"
            f'<td style="border:1px solid #222;padding:7px 10px;font-family:Calibri,Arial,sans-serif;font-size:13px;">{esc(r.get("semana", ""))}</td>'
            f'<td style="border:1px solid #222;padding:7px 10px;font-family:Calibri,Arial,sans-serif;font-size:13px;">{esc(r.get("postergado_por_proyecto", "No"))}</td>'
            f'<td style="border:1px solid #222;padding:7px 10px;font-family:Calibri,Arial,sans-serif;font-size:13px;">{esc(r.get("usuario", ""))}</td>'
            f'<td style="border:1px solid #222;padding:7px 10px;font-family:Calibri,Arial,sans-serif;font-size:13px;">{esc(r.get("correo", ""))}</td>'
            f'<td style="border:1px solid #222;padding:7px 10px;font-family:Calibri,Arial,sans-serif;font-size:13px;">{esc(r.get("asunto", ""))}</td>'
            f"</tr>"
            for r in rows
        )

        paragraphs = "".join(
            f'<p style="font-family:Calibri,Arial,sans-serif;font-size:14px;color:#222;margin:0 0 10px;">{esc(line)}</p>'
            for line in (intro_text or "").splitlines()
            if line.strip()
        )

        return (
            f"<div>{paragraphs}"
            '<table style="border-collapse:collapse;margin-top:12px;min-width:680px;font-family:Calibri,Arial,sans-serif;">'
            "<tr>"
            '<td colspan="2" rowspan="3" style="border:1px solid #222;padding:18px 24px;text-align:center;font-weight:700;font-size:15px;background:#ffffff;">'
            "LOGO HELPDESK"
            "</td>"
            '<td colspan="3" style="border:none;"></td>'
            "</tr>"
            '<tr><td colspan="3" style="border:none;"></td></tr>'
            '<tr><td colspan="3" style="border:none;"></td></tr>'
            '<tr><td colspan="5" style="height:14px;border:none;"></td></tr>'
            '<tr style="background:#f2f4f0;font-weight:700;">'
            '<th style="border:1px solid #222;padding:8px 10px;text-align:left;font-size:13px;">SEMANA</th>'
            '<th style="border:1px solid #222;padding:8px 10px;text-align:left;font-size:13px;">POSTERGADO POR PROYECTO</th>'
            '<th style="border:1px solid #222;padding:8px 10px;text-align:left;font-size:13px;">USUARIO</th>'
            '<th style="border:1px solid #222;padding:8px 10px;text-align:left;font-size:13px;">CORREO</th>'
            '<th style="border:1px solid #222;padding:8px 10px;text-align:left;font-size:13px;">ASUNTO</th>'
            "</tr>"
            f"{body_rows_html}"
            "</table>"
            "</div>"
        )

    def send_planning_excel_email(
        self,
        *,
        recipients: list[str],
        cc_recipients: list[str] | None = None,
        subject: str,
        intro_text: str,
        rows: list[dict[str, Any]],
        filename: str = "Planificacion_Helpdesk.xlsx",
        mode: str = "send",
    ) -> dict[str, Any]:
        """
        Generate the Helpdesk Excel (.xlsx) file, attach it to an Outlook email
        with the formatted HTML table, and either open the draft in Outlook ('display')
        or send it immediately ('send') after user approval.
        """
        clean_recipients = [r.strip() for r in recipients if r and "@" in r]
        if not clean_recipients:
            raise OutlookGraphError("Debes indicar al menos un destinatario válido.", status_code=400)
        if not rows:
            raise OutlookGraphError("No hay solicitudes en la tabla para enviar.", status_code=400)

        xlsx_bytes = self.build_planning_excel_bytes(rows)
        html_body = self.build_planning_email_html(intro_text=intro_text, rows=rows)

        safe_filename = re.sub(r"[^A-Za-z0-9_.-]", "_", filename or "Planificacion_Helpdesk.xlsx")
        if not safe_filename.endswith(".xlsx"):
            safe_filename += ".xlsx"

        temp_dir = tempfile.mkdtemp(prefix="umbral_excel_")
        temp_path = os.path.join(temp_dir, safe_filename)
        with open(temp_path, "wb") as fh:
            fh.write(xlsx_bytes)

        try:
            import pythoncom  # type: ignore[import-not-found,import-untyped]
            import win32com.client  # type: ignore[import-not-found,import-untyped]
        except ImportError as exc:
            raise OutlookGraphError("pywin32_not_installed", status_code=503) from exc

        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)  # 0 == olMailItem
                mail.To = "; ".join(clean_recipients)
                if cc_recipients:
                    clean_cc = [c.strip() for c in cc_recipients if c and "@" in c]
                    if clean_cc:
                        mail.CC = "; ".join(clean_cc)
                mail.Subject = subject.strip() or "Planificación Semanal de Solicitudes - Helpdesk"
                mail.HTMLBody = html_body
                mail.Attachments.Add(temp_path)

                if mode == "display":
                    mail.Display(False)
                    status_label = "draft_opened"
                else:
                    mail.Send()
                    status_label = "sent"
            except Exception as exc:
                raise OutlookGraphError(f"Error al procesar el correo en Outlook: {exc}", status_code=503) from exc
        finally:
            pythoncom.CoUninitialize()

        batch_record = {
            "batch_id": f"plan-{uuid.uuid4().hex[:8]}",
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "subject": subject,
            "recipients": clean_recipients,
            "cc_recipients": cc_recipients or [],
            "rows_count": len(rows),
            "filename": safe_filename,
            "status": status_label,
            "rows": rows,
        }
        self.sent_batches.append(batch_record)
        return batch_record

    @staticmethod
    def _create_local_calendar_event(
        *,
        subject: str,
        body: str,
        start: date,
        end: date,
    ) -> str:
        try:
            import pythoncom  # type: ignore[import-not-found,import-untyped]
            import pywintypes  # type: ignore[import-not-found,import-untyped]
            import win32com.client  # type: ignore[import-not-found,import-untyped]
        except ImportError as exc:
            raise OutlookGraphError("pywin32_not_installed", status_code=503) from exc

        pythoncom.CoInitialize()
        try:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                try:
                    appointment = outlook.CreateItem(1)  # 1 == olAppointmentItem
                except Exception:
                    namespace = outlook.GetNamespace("MAPI")
                    calendar_folder = namespace.GetDefaultFolder(9)
                    appointment = calendar_folder.Items.Add(1)

                appointment.Subject = f"[Planificación] {subject}"
                start_dt = datetime.combine(start, time(9, 0))
                end_dt = datetime.combine(end, time(17, 0))
                try:
                    appointment.Start = pywintypes.Time(start_dt)
                    appointment.End = pywintypes.Time(end_dt)
                except Exception:
                    appointment.Start = start_dt
                    appointment.End = end_dt
                appointment.Body = re.sub(r"<[^>]+>", "", body)
                appointment.ReminderSet = True
                appointment.ReminderMinutesBeforeStart = 30
                appointment.Save()
                entry_id = str(getattr(appointment, "EntryID", "") or "")
                return entry_id or f"local-evt-{start.isoformat()}"
            except Exception as exc:
                raise OutlookGraphError("local_calendar_event_creation_failed", status_code=503) from exc
        finally:
            pythoncom.CoUninitialize()

    @staticmethod
    def _parse_date(value: str | date) -> date:
        if isinstance(value, date):
            return value
        try:
            return date.fromisoformat(str(value))
        except ValueError as exc:
            raise OutlookGraphError("invalid_date") from exc
