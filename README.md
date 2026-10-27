# Plataforma de Gestión Outlook-Zimbra (Umbral Pro)

Plataforma centralizada para sincronizar todas las bandejas y subcarpetas de Outlook, clasificar alertas con triaje visual semántico, planificar solicitudes por semana (con cupo máximo de 4 solicitudes semanales), generar y previsualizar el formato oficial en **Excel (`.xlsx`) de Helpdesk** y enviarlo por correo con aprobación previa del administrador.

## Funcionalidades principales implementadas

- **Sincronización completa de Outlook (`OUTLOOK_MODE=local` y `graph`):**
  - **Modo Local MAPI (`pywin32` / `win32com.client`):** Conecta directamente con la aplicación de escritorio de Outlook en Windows sin necesidad de registrar una app en Microsoft Entra ID. Recorre todas las cuentas, carpetas y subcarpetas mediante `Folder.GetTable().GetArray()` a alta velocidad, extrayendo remitente SMTP (`PR_SENDER_SMTP_ADDRESS`), múltiples categorías, carpeta de origen y año fiscal sin límite fijo en el contador.
  - **Modo Microsoft Graph (`MSAL`):** Soporte OAuth 2.0 con caché de tokens cifrada mediante `Fernet`.
- **Sistema de Diseño "Umbral Pro — Swiss Sapphire":**
  - Arquitectura de superficies en 3 capas (`#EAF0F7` Sidebar, `#F4F7FB` Lienzo y `#FFFFFF` Tarjetas/Tabla) más **Modo Noche Operativo (`body.dark`)**.
  - **Triaje Visual Semántico en 0.5s:** Identificación automática de correos de caída/error (`.email-row-failure` con borde rojo coral e icono `!`) y servicios restablecidos (`.email-row-success` con borde esmeralda e icono `✓`).
  - **HTML/CSS descentralizado por columna y campo:** Clases independientes (`.col-head-*`, `.col-cell-*`, `.sender-box`, `.subject-box`, etc.) y alineación numérica con `tabular-nums`.
- **Planificación Semanal (Máx. 4 por semana) y Formato Excel Helpdesk:**
  - Asignación por semana ISO (`YYYY-Www`, bloque laboral de lunes a viernes) con validación de cupo máximo de **4 solicitudes por semana** y registro del campo **Postergado por proyecto**.
  - Creación automática del evento en el Calendario de Outlook local (`pywintypes.Time` compatible con configuración regional en español).
  - **Vista Previa Interactiva y Aprobación de Envío:** Simulación en pantalla idéntica a la hoja de Excel con el logo proporcional de **Helpdesk (`A1:B3`)** y columnas `SEMANA | POSTERGADO POR PROYECTO | USUARIO | CORREO | ASUNTO`, permitiendo editar celdas antes de descargar el `.xlsx`, abrir el borrador en Outlook o aprobar el envío directo con adjunto.

## Documentación del proyecto

- [SDD: índice de especificación](SDD.md)
- [Features y criterios de aceptación](features/)
- [Plan técnico y contratos de API](PLAN.md)
- [Tareas e hitos](TASKS.md)
- [Backend: configuración y ejecución local](backend/README.md)
- [Frontend: configuración y ejecución local](outlook-zimbra-frontend/README.md)

## Pila tecnológica

- **IDE:** Visual Studio Code.
- **Backend:** Python 3.11+, FastAPI, `pywin32` (MAPI local), MSAL, Pydantic Settings y Cryptography (`Fernet`).
- **Frontend:** Ionic Framework con Angular 20 (Standalone Components).
- **Generación de Reportes:** Motor OpenXML `.xlsx` nativo en Python (con incrustación proporcional de logo PNG en `xl/drawings/drawing1.xml`).
- **Integraciones de Correo:** Outlook Desktop MAPI / Microsoft Graph y SMTP autenticado con TLS para Zimbra.

## Principios de seguridad

- **Cero credenciales en Git:** Ningún secreto, contraseña, llave de sesión, caché de tokens ni correo corporativo se almacena en el código fuente ni en el repositorio.
- **Protección mediante `.gitignore`:** Los archivos `.env`, `*.bin`, `*.key`, `*.pem`, `data/`, `.venv/` y `node_modules/` están bloqueados en `.gitignore` (raíz, `backend/` y `outlook-zimbra-frontend/`).
- **Sesiones firmadas y endpoints protegidos:** Todos los endpoints de sincronización, lectura de correos, planificación, exportación Excel y reenvío Zimbra exigen sesión autenticada (`HttpOnly` cookie firmada con `SESSION_SECRET_KEY`).
- **Plantilla segura:** Usa siempre `backend/.env.example` como base local para crear `backend/.env` sin subirlo jamás al control de versiones.

## Ejecución rápida en desarrollo

1. **Backend (`http://localhost:8000`):** Consulta la [guía del backend](backend/README.md).
2. **Frontend (`http://localhost:8100`):** Consulta la [guía del frontend](outlook-zimbra-frontend/README.md).
