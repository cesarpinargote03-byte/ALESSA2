# Feature 1: Bandeja de Outlook y Triaje Visual Semántico

## Objetivo

Centralizar en el dashboard todos los correos entrantes de todas las carpetas y subcarpetas de Outlook, permitiendo clasificarlos visualmente en medio segundo y filtrarlos sin límite artificial en el contador.

## Historia de usuario

Como administrador único, quiero visualizar todos los correos de mis bandejas y subcarpetas de Outlook, distinguir de inmediato las caídas de servicio frente a las recuperaciones y filtrar por categoría, año fiscal, carpeta y texto de búsqueda.

## Criterios de aceptación implementados

- **Lectura completa de carpetas y subcarpetas:** En modo local MAPI (`OUTLOOK_MODE=local`), el backend recorre recursivamente todas las carpetas de correo usando `Folder.GetTable().GetArray()` y extrae la dirección SMTP del remitente desde `PR_SENDER_SMTP_ADDRESS` sin llamadas lentas a Active Directory.
- **Contador real sin tope fijo:** El contador refleja el total real de correos sincronizados y se incrementa conforme ingresan nuevos mensajes.
- **Triaje Visual Semántico en 0.5 segundos:**
  - Los correos de caída o fallo (`Failure`, `Error`, `Caída`, `Cierre inesperado`, `Down`) muestran borde izquierdo rojo coral (`#F43F5E`) e icono de alerta `!`.
  - Los correos de servicio restablecido (`Success`, `Restablecido`, `Recovered`, `Resuelto`) muestran borde izquierdo verde esmeralda (`#10B981`) e icono `✓`.
- **HTML/CSS descentralizado y responsivo:** Cada columna (`Remitente`, `Asunto`, `Categoría`, `Carpeta`, `Recibido`, `Acciones`) utiliza clases independientes en `home.page.html` y `home.page.scss`, con scroll vertical propio en la tabla y barra de paginación fija.
