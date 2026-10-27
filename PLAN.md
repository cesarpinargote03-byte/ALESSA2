# Plan técnico

## Arquitectura implementada

- **Frontend (`outlook-zimbra-frontend`):** Aplicación Ionic/Angular 20 con el sistema de diseño **Umbral Pro — Swiss Sapphire** (superficies en 3 capas + Modo Noche), triaje visual semántico de alertas, bandeja paginada con scroll independiente, planificación por semana ISO (máximo 4 solicitudes por semana) y modal de vista previa y aprobación del archivo **Excel Helpdesk (`.xlsx`)**.
- **Backend (`backend`):** Servicio FastAPI con soporte dual de integración Outlook:
  - **Modo Local MAPI (`OUTLOOK_MODE=local`):** Lectura de alta velocidad de todas las cuentas, carpetas y subcarpetas de Outlook Desktop mediante `Folder.GetTable().GetArray()`, creación de eventos en calendario con `pywintypes.Time` y generación/envío de reportes en Excel `.xlsx` con logo incrustado.
  - **Modo Microsoft Graph (`OUTLOOK_MODE=graph`):** OAuth 2.0 mediante MSAL con caché persistente cifrada con `Fernet`.
- **Seguridad de sesiones y secretos:** Cookie de sesión firmada (`HttpOnly`), dependencia `require_graph_access_token` en todos los endpoints de negocio y exclusión estricta de `.env` y cachés en `.gitignore`.

## Modelo de datos de Solicitud y Planificación

| Campo | Tipo | Descripción |
|---|---|---|
| `id` | VARCHAR | `EntryID` de MAPI o identificador de Microsoft Graph |
| `subject` | VARCHAR | Asunto del mensaje |
| `remitente_email` | VARCHAR | Dirección SMTP del remitente (`PR_SENDER_SMTP_ADDRESS`) |
| `remitente_nombre` | VARCHAR | Nombre para mostrar del remitente |
| `categoria` / `categorias` | VARCHAR / LIST | Categorías de Outlook asociadas al correo |
| `ano_fiscal` | INTEGER, nullable | Año de recepción del correo |
| `carpeta_origen` | VARCHAR | Ruta completa de la carpeta/subcarpeta en Outlook |
| `semana_iso` | VARCHAR, nullable | Semana ISO asignada (ej. `2026-W44`) |
| `semana_numero` | INTEGER, nullable | Número de semana del año (`1..53`) |
| `semana_planificada_inicio` | DATE, nullable | Lunes de la semana planificada |
| `semana_planificada_fin` | DATE, nullable | Viernes de la semana planificada |
| `postergado_por_proyecto` | VARCHAR | Valor para la columna B del Excel Helpdesk (`No`, `Sí` o nombre del proyecto) |
| `outlook_event_id` | VARCHAR, nullable | Identificador de la cita creada en el calendario de Outlook |
| `estado_planificacion` | VARCHAR | `Planificado` |

## Contratos de API vigentes

- `POST /api/v1/auth/login-local`: autentica contra la instancia local de Outlook Desktop en Windows y detecta el correo activo.
- `GET /api/v1/auth/status` y `GET /api/v1/auth/me`: estado de autenticación y correo de la sesión.
- `POST /api/v1/outlook/sync`: sincroniza todos los correos de todas las carpetas y subcarpetas sin tope fijo.
- `GET /api/v1/outlook/filtros`: devuelve el resumen de carpetas, categorías y años fiscales.
- `GET /api/v1/correos`: devuelve los correos sincronizados con filtros combinables.
- `GET /api/v1/planificacion`: devuelve las solicitudes planificadas ordenadas por semana.
- `POST /api/v1/planificacion/asignar`: valida el límite máximo de **4 solicitudes por semana**, crea la cita en el calendario de Outlook y guarda la planificación.
- `POST /api/v1/planificacion/excel`: genera y devuelve el archivo `.xlsx` con el logo proporcional de Helpdesk en `A1:B3` (`<xdr:oneCellAnchor>`) y la tabla `SEMANA | POSTERGADO POR PROYECTO | USUARIO | CORREO | ASUNTO`.
- `POST /api/v1/planificacion/enviar-excel`: genera el `.xlsx`, lo adjunta junto con la tabla HTML y el logo en línea (`cid:logo_helpdesk_umbral`) y abre el borrador en Outlook (`mode="display"`) o envía el correo tras la aprobación (`mode="send"`).
- `GET /api/v1/planificacion/envios`: devuelve el historial de lotes de planificación enviados.
- `POST /api/v1/zimbra/reenvio-masivo` y `GET /api/v1/zimbra/reenvio-masivo/{batch_id}`: endpoints de reenvío SMTP protegidos por sesión autenticada.
