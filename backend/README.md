# Backend (FastAPI + Outlook MAPI / Graph + Excel Helpdesk + Zimbra)

Backend de **Umbral Pro** construido con FastAPI. Soporta dos modos de operación con Outlook:
1. **`OUTLOOK_MODE=local` (Recomendado en Windows corporativo):** Se conecta directamente a la aplicación **Outlook Desktop** abierta en Windows mediante COM/MAPI (`pywin32`), sin requerir permisos de administrador en Microsoft Entra ID.
2. **`OUTLOOK_MODE=graph`:** Utiliza OAuth 2.0 con Microsoft Graph y caché persistente cifrada con `Fernet`.

## Requisitos

- Windows con Microsoft Outlook Desktop configurado (para `OUTLOOK_MODE=local`).
- Python 3.11 o superior.

## Configuración segura (`.env`)

1. Copia el archivo de ejemplo (nunca subas `.env` a Git):

```powershell
cd backend
Copy-Item .env.example .env
```

2. Genera llaves criptográficas únicas para tu entorno local y pégalas en `backend/.env`:

```powershell
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

3. Verifica en `backend/.env`:
   - `OUTLOOK_MODE=local`
   - `TOKEN_CACHE_ENCRYPTION_KEY=<tu_llave_fernet_generada>`
   - `SESSION_SECRET_KEY=<tu_llave_secreta_generada>`
   - `MICROSOFT_ADMIN_EMAIL=<tu_correo_configurado_en_outlook>` (opcional en modo local; se autodetecta desde la sesión activa de Outlook).

> **Seguridad:** El archivo `backend/.env` y la carpeta `backend/data/` están excluidos en `.gitignore`. Nunca agregues contraseñas ni llaves reales en `.env.example`.

## Ejecución local

Desde la carpeta `backend/` con el entorno virtual activo:

```powershell
uvicorn app.main:app --reload
```

### Endpoints principales

- **Salud:**
  - `GET /health`
- **Autenticación:**
  - `POST /api/v1/auth/login-local` — Conecta con la sesión activa de Outlook Desktop (MAPI) y devuelve el correo detectado.
  - `GET /api/v1/auth/login` — Inicio de sesión OAuth (cuando `OUTLOOK_MODE=graph`).
  - `GET /api/v1/auth/status` — Verifica si la sesión está activa.
  - `GET /api/v1/auth/me` — Devuelve el correo de la cuenta conectada.
  - `POST /api/v1/auth/logout` — Cierra la sesión.
- **Bandeja y Filtros de Outlook:**
  - `POST /api/v1/outlook/sync` — Sincroniza todas las carpetas y subcarpetas de todos los buzones de Outlook a alta velocidad (`GetTable().GetArray()`).
  - `GET /api/v1/outlook/filtros` — Resumen de carpetas, categorías y años fiscales disponibles.
  - `GET /api/v1/correos` — Lista de correos con filtros combinables (`categoria`, `ano`, `folder`, `remitente`, `page`, `page_size`) sin tope fijo.
- **Planificación Semanal (Cupo máx. 4 por semana) y Excel Helpdesk:**
  - `GET /api/v1/planificacion` — Lista las solicitudes planificadas.
  - `POST /api/v1/planificacion/asignar` — Valida el cupo máximo de 4 solicitudes en la semana ISO seleccionada, agenda el evento en el calendario de Outlook local y guarda el campo `postergado_por_proyecto`.
  - `POST /api/v1/planificacion/excel` — Genera y descarga el archivo `.xlsx` con el logo proporcional de Helpdesk en `A1:B3` y las columnas `SEMANA | POSTERGADO POR PROYECTO | USUARIO | CORREO | ASUNTO`.
  - `POST /api/v1/planificacion/enviar-excel` — Adjunta el archivo `.xlsx` e inserta la tabla HTML con logo en un correo de Outlook para abrir el borrador (`mode="display"`) o enviarlo directamente tras la aprobación (`mode="send"`).
  - `GET /api/v1/planificacion/envios` — Historial de lotes de planificación enviados.
- **Reenvío Zimbra (Protegido por sesión):**
  - `POST /api/v1/zimbra/reenvio-masivo`
  - `GET /api/v1/zimbra/reenvio-masivo/{batch_id}`

## Pruebas automatizadas

```powershell
python -m pytest -p no:cacheprovider
```
