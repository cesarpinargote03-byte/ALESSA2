# Plataforma de Gestión Outlook-Zimbra

Plataforma para centralizar correos de Outlook, planificar solicitudes en el calendario y reenviar mensajes seleccionados mediante el servidor institucional Zimbra. El sistema está destinado a un administrador único.

## Objetivos

- Consultar y filtrar solicitudes de Outlook desde un dashboard.
- Asignar solicitudes a semanas de planificación y reflejar la planificación en Outlook.
- Reenviar mensajes mediante Zimbra con resultados auditables por lote.
- Mostrar métricas útiles sin ocultar errores ni estados incompletos.

## Documentación del proyecto

- [SDD: índice de especificación](SDD.md)
- [Features y criterios de aceptación](features/)
- [Plan técnico](PLAN.md)
- [Tareas e hitos](TASKS.md)
- [Backend: configuración y ejecución local](backend/README.md)
- [Frontend: configuración y ejecución local](outlook-zimbra-frontend/README.md)

## Pila tecnológica propuesta

- **IDE:** Visual Studio Code.
- **Asistente de IA:** Gemini (Google).
- **Backend:** Python y FastAPI.
- **Frontend:** Ionic Framework con Angular.
- **Base de datos:** PostgreSQL; Supabase Free Tier es una opción para desarrollo o bajo volumen.
- **Integraciones:** Microsoft Graph para Outlook y SMTP autenticado con TLS para Zimbra.
- **Gráficos:** Chart.js o NGX-Charts para Angular.

La meta es priorizar soluciones gratuitas y de código abierto, pero no se garantiza costo cero, disponibilidad ni permanencia de los planes gratuitos. Antes de desplegar se deben revisar cuotas, pausas, límites, retención y términos vigentes de cada proveedor.

## Principios de seguridad

- No se incluirán secretos, tokens o contraseñas en el código, el repositorio ni los logs.
- `.env` se reserva para configuración y secretos locales de desarrollo y debe excluirse del control de versiones. En despliegues se preferirá un gestor de secretos del entorno.
- MSAL gestionará los tokens de Microsoft y su renovación. La caché persistente se protegerá con permisos restringidos y almacenamiento seguro adecuado al entorno. No se guardarán refresh tokens directamente en `.env` ni se implementará manualmente su rotación.
- La renovación transparente no garantiza que nunca se requiera iniciar sesión de nuevo; una revocación, políticas de acceso condicional, cambios de contraseña o consentimiento pueden requerir autenticación interactiva.
- Se solicitarán los permisos mínimos necesarios de Microsoft Entra ID. Las comunicaciones remotas usarán TLS.
- Los logs no incluirán contraseñas, tokens ni contenido sensible de los mensajes.
- El backend se ejecutará localmente durante el desarrollo. Para pruebas externas se podrá usar Ngrok con autenticación obligatoria y token restringido; no se abrirán puertos del router. Ngrok es una herramienta de desarrollo, no un despliegue de producción ni una garantía de seguridad.

## Alcance inicial

El desarrollo se realizará por tareas pequeñas y verificables, empezando por la autenticación segura del backend. Las funciones de Outlook y Zimbra requieren credenciales, permisos y políticas aprobados por sus administradores correspondientes.

## Ejecutar el frontend

Con Node.js instalado, inicia primero el backend según su [guía local](backend/README.md). Después, desde `outlook-zimbra-frontend\` ejecuta `npm install` y `npm start`; la interfaz estará disponible en `http://localhost:8100`. Sin sesión Microsoft verás una vista de muestra claramente identificada. Para usar Outlook, configura el consentimiento de `User.Read`, `Mail.Read` y `Calendars.ReadWrite` y conecta la cuenta desde la aplicación.
