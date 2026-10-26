# Frontend

Interfaz web de Umbral construida con Ionic y Angular. Incluye bandeja con búsqueda, filtros combinables y paginación, detalle de solicitud, planificación semanal y una vista de muestra para explorar el diseño sin conectar una cuenta.

## Requisitos

- Node.js compatible con Angular 20 (Node.js 20.19+, 22.12+ o 24+).
- Backend en ejecución en `http://localhost:8000` para iniciar sesión y cargar datos reales.

## Desarrollo local

Desde esta carpeta:

```powershell
npm install
npm start
```

Abre `http://localhost:8100`. El servidor local de Angular usa el puerto que el backend permite en CORS por defecto. Para ver datos reales, inicia sesión con Microsoft desde el panel y concede los permisos Graph configurados en `MICROSOFT_GRAPH_SCOPES`. Sin sesión, la aplicación muestra datos ficticios señalados como **vista de muestra**; las acciones que modifican Outlook no se simulan.

## Validación

```powershell
npm run build
npm run lint
```

## Integración

La URL de la API se define en `src/environments/environment.ts`; la compilación de producción usa `src/environments/environment.prod.ts`. La API debe habilitar cookies con credenciales y servir la aplicación bajo el mismo sitio cuando se use producción.
