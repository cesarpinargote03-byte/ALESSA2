# Backend

Backend inicial de FastAPI. Implementa el inicio de sesión de Microsoft Entra ID mediante MSAL y una caché persistente cifrada para tokens. Aún no implementa la sincronización de mensajes, la planificación ni el envío SMTP descritos en el [plan técnico](../PLAN.md).

## Requisitos

- Python 3.11 o superior.
- Una aplicación web registrada para **cuentas personales de Microsoft**, con URI de redirección `http://localhost:8000/api/v1/auth/callback`.
- Consentimiento para los permisos delegados configurados en `MICROSOFT_GRAPH_SCOPES` (por defecto, `User.Read Mail.Read Calendars.ReadWrite`).

## Preparación local (Windows PowerShell)

```powershell
cd backend
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[dev]"
Copy-Item .env.example .env
```

Genera valores aleatorios diferentes para `TOKEN_CACHE_ENCRYPTION_KEY` y `SESSION_SECRET_KEY`:

```powershell
python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())"
python -c "import secrets; print(secrets.token_urlsafe(48))"
```

Completa en `.env` el client ID, client secret y correo de la cuenta Microsoft personal permitida. En el registro de aplicación selecciona **Personal Microsoft accounts only** (solo cuentas personales de Microsoft). El backend usa el endpoint `consumers` y valida el correo permitido; no requiere configurar un tenant ID. El client ID identifica la aplicación y no es una contraseña; protege el client secret, las claves de sesión y cifrado, y la caché de tokens. No guardes `.env`, la caché ni secretos en el repositorio. La caché por defecto se escribe en `backend/data/` y se cifra antes de persistir.

## Ejecución

Desde `backend/`, con el entorno virtual activado:

```powershell
uvicorn app.main:app --reload
```

- Salud: `GET http://localhost:8000/health`
- Inicio de sesión: abre `http://localhost:8000/api/v1/auth/login` en el navegador.
- Estado de sesión: `GET http://localhost:8000/api/v1/auth/status`
- Usuario administrador: `GET http://localhost:8000/api/v1/auth/me`
- Cierre de sesión: `POST http://localhost:8000/api/v1/auth/logout`

El callback redirige a `FRONTEND_REDIRECT_URL`. El estado OAuth se valida mediante MSAL y la sesión se transporta en una cookie firmada `HttpOnly`, sin guardar tokens en ella. El cookie `Secure` se activa con `COOKIE_SECURE=true`; debe ser obligatorio en despliegues HTTPS. Se permite HTTP solo para el desarrollo local.

El endpoint de salud no demuestra que Microsoft esté configurado ni que Graph sea accesible. Los endpoints de autenticación fallarán explícitamente si falta configuración.

## Pruebas

Desde la raíz del proyecto, instala el backend en el entorno virtual activo y ejecuta pytest:

```powershell
python -m pip install -e ".\backend[dev]"
python -m pytest -p no:cacheprovider
```

También puedes ejecutar las pruebas desde `backend\` usando el entorno creado allí:

```powershell
cd backend
python -m pytest
```

Las pruebas de autenticación sustituyen MSAL y no contactan Microsoft ni requieren credenciales reales.
