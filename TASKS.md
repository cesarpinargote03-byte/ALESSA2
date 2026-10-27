# Tareas e hitos

Estado actualizado de los hitos del proyecto **Umbral Pro (Outlook-Zimbra)**.

## Hito 1: Autenticación, seguridad e infraestructura backend

- [x] **1.1 — Preparar FastAPI y dependencias:** configuración validada con Pydantic Settings, cifrado `Fernet` y soporte dual `OUTLOOK_MODE=local` / `graph`.
- [x] **1.2 — Implementar autenticación local MAPI y Entra ID:** login directo con Outlook Desktop (`POST /api/v1/auth/login-local`) sin requerir permisos de administrador en Entra ID, además del flujo OAuth 2.0 con MSAL.
- [x] **1.3 — Seguridad y protección de secretos:** todos los endpoints de negocio exigen sesión autenticada; `.gitignore` configurado en raíz, `backend/` y `outlook-zimbra-frontend/` para bloquear `.env`, llaves y cachés.

## Hito 2: Integración Outlook, Planificación Semanal y Formato Excel Helpdesk

- [x] **2.1 — Sincronizar mensajes Outlook sin tope fijo:** lectura MAPI de alta velocidad (`Folder.GetTable().GetArray()`) recorriendo todas las carpetas y subcarpetas, extrayendo dirección SMTP (`PR_SENDER_SMTP_ADDRESS`), categorías múltiples y año fiscal.
- [x] **2.2 — Planificar por semana en calendario Outlook (máx. 4 por semana):** selección por semana ISO (`YYYY-Www`), control de cupo máximo de 4 solicitudes por semana, campo `postergado_por_proyecto` y creación de citas con `pywintypes.Time`.
- [x] **2.3 — Generación y envío aprobado de Excel Helpdesk (`.xlsx`):** generador `.xlsx` OpenXML nativo con logo oficial de Helpdesk proporcional en `A1:B3`, vista previa editable en pantalla, descarga `.xlsx`, apertura de borrador en Outlook y envío directo con registro de lotes.
- [x] **2.4 — Envío SMTP Zimbra autenticado:** servicio de lotes con validación de destinatarios y protección de endpoints mediante sesión activa.

## Hito 3: Interfaz Ionic/Angular (Umbral Pro — Swiss Sapphire)

- [x] **3.1 — Sistema de diseño Swiss Sapphire en 3 capas + Modo Noche:** Sidebar (`#EAF0F7`), Lienzo (`#F4F7FB`), Superficies (`#FFFFFF`), tipografía de alto contraste, números tabulares (`tabular-nums`) y conmutador de Modo Noche (`body.dark`).
- [x] **3.2 — Bandeja responsiva con Triaje Visual Semántico:** scroll propio en tabla, clases HTML/CSS descentralizadas por columna y detección automática en 0.5s de correos de caída/error (`.email-row-failure` + icono `!`) y servicios restablecidos (`.email-row-success` + icono `✓`).
- [x] **3.3 — Modal de Vista Previa Excel y Aprobación de Envío:** previsualización idéntica al archivo Excel Helpdesk con edición en línea de celdas antes de aprobar el envío o descargar el `.xlsx`.
- [x] **3.4 — Historial de envíos realizados:** consulta de lotes de planificación enviados o abiertos en Outlook desde la pestaña Reenvíos.
