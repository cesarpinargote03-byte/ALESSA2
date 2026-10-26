# Feature 2: Planificación semanal

## Objetivo

Asignar una solicitud a una semana, reflejar la planificación en el calendario de Outlook y enviar una confirmación HTML al solicitante.

## Historia de usuario

Como administrador único, quiero asignar una solicitud a una semana de planificación, registrar el periodo en el calendario de Outlook y notificar al solicitante para coordinar el trabajo.

## Criterios de aceptación

- Se puede previsualizar el mensaje HTML antes de enviarlo.
- Si falta una variable requerida o no es válida (por ejemplo, nombre o fecha), el envío queda deshabilitado y se identifica el dato que debe corregirse.
- La zona horaria institucional y la regla para calcular inicio y fin de semana están configuradas explícitamente; no se infieren de la zona horaria del navegador.
- Los errores de Microsoft Graph y los conflictos de calendario se muestran al administrador y nunca se reportan como éxito.
- Se evita crear eventos duplicados cuando se reintenta una operación.
- El sistema registra por separado el resultado de crear el evento y el de enviar la notificación.
- Si se crea el evento, pero falla el envío de la confirmación, se permite reintentar solo el paso fallido sin duplicar el evento.
- El estado de la solicitud solo avanza después de confirmar cada operación correspondiente.

## Flujo esperado

1. Seleccionar solicitud y semana.
2. Validar fechas, zona horaria y datos requeridos por la plantilla.
3. Mostrar la previsualización y solicitar confirmación del administrador.
4. Crear o verificar el evento en Outlook.
5. Enviar la notificación HTML.
6. Persistir y mostrar el resultado de cada paso.

