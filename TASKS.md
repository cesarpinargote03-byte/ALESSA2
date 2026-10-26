# Tareas e hitos

Las estimaciones son iniciales; se revisarán al conocer el entorno de Microsoft Entra ID, las políticas institucionales, el servidor Zimbra y los límites del proveedor de base de datos.

## Hito 1: Autenticación e infraestructura backend

- [x] **1.1 — Preparar FastAPI y dependencias** (2 h): crear el backend e incorporar MSAL, cliente HTTP, driver PostgreSQL y configuración segura.
- [x] **1.2 — Implementar autenticación de Entra ID** (4 h): OAuth 2.0, permisos mínimos y caché persistente protegida de MSAL. No guardar refresh tokens en `.env`.
- [x] **1.3 — Configurar entorno y pruebas** (2 h): separar configuración por entorno, excluir secretos del control de versiones y probar autenticación/caché. Las pruebas usan MSAL simulado; la validación contra el tenant real queda pendiente de credenciales y consentimiento institucional.

## Hito 2: Integración Outlook y Zimbra

- [ ] **2.1 — Sincronizar mensajes Outlook** (3 h): paginación, carpetas y categorías; tratar límites y errores de Graph.
- [ ] **2.2 — Implementar envío SMTP Zimbra** (3 h): TLS, límites, pausa configurable, reintentos controlados y reporte persistido.
- [ ] **2.3 — Planificar en calendario Outlook** (3 h): zona horaria, prevención de duplicados y recuperación de fallos parciales.

## Hito 3: Interfaz Ionic/Angular

- [x] **3.1 — Crear estructura de aplicación** (2 h): Dashboard, Solicitudes y Configuración. Estructura base de Ionic/Angular lista; módulos de Reenvíos y métricas aún tienen integración pendiente.
- [x] **3.2 — Implementar bandeja y filtros** (4 h): filtros combinables por categoría, año, carpeta y remitente; paginación y estados de carga/error/vacío. La vista de muestra usa datos explícitamente identificados; la carga real requiere autenticación Microsoft.
- [ ] **3.3 — Crear previsualización de planificación** (3 h): modal de confirmación HTML y validación de variables antes del envío.
- [ ] **3.4 — Incorporar métricas** (3 h): gráficos de solicitudes y volumen de correos.
- [ ] **3.5 — Mostrar reportes de reenvío** (3 h): resultados por lote, mensaje y destinatario.

## Orden y dependencias

1. **1.1 → 1.2 → 1.3:** la autenticación y su configuración requieren primero la estructura y dependencias del backend.
2. **2.1 y 2.2** requieren que la configuración y los secretos estén disponibles (1.3); Outlook además requiere autenticación lista (1.2).
3. **2.3** requiere autenticación lista (1.2) y persistencia/configuración del backend (1.3).
4. **3.2 y 3.3** dependen de los contratos y flujos de las funciones correspondientes (2.1 y 2.3).
5. **3.5** depende de la API de lotes y reportes (2.2).
6. **3.1** puede iniciar en paralelo con el backend; **3.4** depende de disponer de métricas/API definidas en el plan.

## Pruebas y validación

- Usar `pytest` para validar entradas, payloads de Microsoft Graph, paginación, caché/renovación de MSAL y manejo de errores.
- Simular SMTP con mocks para probar TLS, pausas, límites, reintentos e informes sin contactar el servidor institucional.
- Añadir pruebas de integración con servicios de prueba cuando estén disponibles, sin credenciales reales en repositorio ni fixtures.
- Antes de producción, revisar permisos de Entra ID, políticas institucionales, almacenamiento de secretos, cuotas de proveedores y exposición de endpoints.
