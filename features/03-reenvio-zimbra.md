# Feature 3: Reenvío mediante Zimbra

## Objetivo

Reenviar mensajes seleccionados o categorizados desde Outlook mediante el servidor SMTP institucional de Zimbra.

## Historia de usuario

Como administrador único, quiero reenviar correos seleccionados o categorizados mediante Zimbra para distribuir solicitudes a sus destinatarios.

## Criterios de aceptación

- El backend obtiene el contenido del mensaje desde Outlook y lo retransmite mediante SMTP autenticado con TLS.
- Cada lote tiene un reporte persistido con el resultado por mensaje y destinatario, incluyendo errores no sensibles que ayuden a diagnosticar o reintentar.
- Se aplican límites configurables de tamaño de lote y concurrencia.
- Se aplica una pausa configurable entre envíos (por ejemplo, 1 a 2 segundos). La pausa reduce la velocidad, pero no garantiza evitar límites o bloqueos de Zimbra.
- Solo se reintentan fallos transitorios, con espera progresiva cuando corresponda; los fallos definitivos no se reintentan automáticamente.
- Los mensajes completados no se reenvían automáticamente. Se utiliza una clave o registro de idempotencia.
- Credenciales, datos de autenticación y contenido sensible no aparecen en reportes ni logs.
- El dashboard permite consultar el estado del lote y el resultado por mensaje/destinatario.

