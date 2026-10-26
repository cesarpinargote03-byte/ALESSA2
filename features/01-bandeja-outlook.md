# Feature 1: Bandeja de Outlook

## Objetivo

Centralizar en el dashboard los correos entrantes que representan solicitudes y permitir encontrar mensajes con filtros combinables.

## Historia de usuario

Como administrador único, quiero visualizar los correos entrantes de Outlook y filtrarlos por categorías, años, remitentes y carpetas para localizar y gestionar solicitudes con rapidez.

## Criterios de aceptación

- Los mensajes se muestran de forma paginada.
- Categoría, año, remitente y carpeta se pueden combinar sin recargar la página completa.
- La interfaz muestra estados de carga y error.
- Si no hay resultados, se muestra “No se encontraron requerimientos para este filtro” sin romper el diseño de la cuadrícula.
- La integración sigue la paginación de Microsoft Graph y respeta sus límites; no asume que todos los mensajes caben en una respuesta.
- Cuando sea compatible con los recursos y permisos configurados, la sincronización incremental evita volver a descargar datos sin cambios.
- Las respuestas de limitación (`429`) respetan `Retry-After`.

## Consideraciones

- La sincronización modifica el estado del servidor y se expone como operación `POST`, no como `GET`.
- Las solicitudes de sincronización repetidas deben ser seguras cuando sea posible.
- Los filtros disponibles se basan en metadatos obtenidos y almacenados por el backend.

