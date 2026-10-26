# Plan técnico

## Arquitectura propuesta

- **Frontend:** Ionic/Angular consume la API de FastAPI y muestra la bandeja, la planificación, los reportes de lotes y las métricas.
- **Backend:** FastAPI valida y autoriza las operaciones, integra Microsoft Graph y SMTP Zimbra, y persiste solicitudes y resultados.
- **Persistencia:** PostgreSQL. Supabase Free Tier es una opción condicionada a sus cuotas y disponibilidad.
- **Autenticación Microsoft:** MSAL con permisos mínimos y caché persistente de tokens protegida. La autenticación interactiva inicial y la renovación siguen las políticas de Microsoft Entra ID.
- **Operaciones externas:** Microsoft Graph por HTTPS/TLS; Zimbra por SMTP con TLS.

## Modelo de datos inicial

### Solicitud

| Campo | Tipo | Descripción |
|---|---|---|
| `id` | UUID, PK | Identificador interno |
| `outlook_message_id` | VARCHAR, UNIQUE | Identificador del mensaje en Outlook |
| `categoria` | VARCHAR, nullable | Categoría de Outlook |
| `ano_fiscal` | INTEGER, nullable | Año fiscal asociado |
| `remitente_email` | VARCHAR | Dirección del remitente |
| `carpeta_origen` | VARCHAR | Carpeta de Outlook |
| `semana_planificada_inicio` | DATE, nullable | Inicio de la semana planificada |
| `semana_planificada_fin` | DATE, nullable | Fin de la semana planificada |
| `outlook_event_id` | VARCHAR, nullable | Identificador del evento de Outlook |
| `estado_planificacion` | ENUM | `Pendiente`, `Previsualizado`, `Planificado`, `Enviado`, `Error` |
| `created_at` / `updated_at` | TIMESTAMPTZ | Auditoría de creación y actualización |

El resultado de crear el evento y el de enviar la notificación deben poder representarse por separado, de modo que los fallos parciales puedan recuperarse sin duplicar operaciones exitosas.

### Lotes Zimbra

Persistir un registro por lote y resultados por mensaje/destinatario: estado, fecha de intento, resultado, código de error no sensible e indicador de reintento. La estructura final se concreta al implementar el feature.

## Contratos de API iniciales

- `POST /api/v1/outlook/sync`: inicia o solicita sincronización de metadatos de Microsoft Graph. Soporta paginación, tratamiento de límites y sincronización incremental si aplica.
- `GET /api/v1/correos?categoria={cat}&ano={ano}&folder={folder}&remitente={email}&page={n}&page_size={size}`: devuelve mensajes almacenados/cacheados con filtros combinables y paginación.
- `POST /api/v1/planificacion/asignar`: recibe `message_id`, fechas de semana y zona horaria. Registra el evento en Outlook y conserva el resultado de cada paso.
- `POST /api/v1/zimbra/reenvio-masivo`: recibe una lista acotada de IDs de Outlook y destinatarios. Responde con un identificador de lote.
- `GET /api/v1/zimbra/reenvio-masivo/{batch_id}`: devuelve estado y resultados persistidos del lote.

Todos los endpoints validan entradas, permisos y límites antes de iniciar operaciones externas. La API deberá autenticar y autorizar al administrador según el entorno de despliegue.

## Consideraciones de fiabilidad

- Respetar la paginación y `Retry-After` ante respuestas `429` de Graph.
- Aplicar idempotencia o detección de operaciones previas para evitar duplicados.
- Registrar fallos parciales y permitir reintentos acotados solo de pasos fallidos.
- Configurar zona horaria institucional y regla de semana de forma explícita.
- Limitar tamaño de lotes, concurrencia, timeout y reintentos de SMTP.
- No incluir secretos ni contenido sensible en logs o reportes.

