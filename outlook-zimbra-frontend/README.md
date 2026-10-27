# Frontend (Umbral Pro — Swiss Sapphire)

Interfaz web de **Umbral Pro** construida con **Ionic y Angular 20**. Implementa el sistema de diseño **Swiss Sapphire** en 3 capas con **Modo Noche Operativo**, triaje visual semántico de alertas, bandeja con scroll independiente y estilos descentralizados por columna, planificación por semana (máx. 4 solicitudes por semana) y vista previa interactiva del formato **Excel Helpdesk** antes de aprobar su envío.

## Características de la interfaz

- **Arquitectura de Superficies en 3 Capas (`Swiss Sapphire`):**
  - **Sidebar (`#EAF0F7`):** Navegación con indicador activo elevado en blanco puro y borde zafiro (`3px`).
  - **Lienzo (`#F4F7FB`):** Fondo azul hielo polar de baja fatiga visual.
  - **Tarjetas KPI y Tabla (`#FFFFFF`):** Bordes de precisión (`#DCE5F0`), sombras multicapa y cifras alineadas con `font-variant-numeric: tabular-nums`.
  - **Modo Noche (`body.dark`):** Botón rápido (`☾ / ☀`) en la barra superior.
- **Triaje Visual Semántico en 0.5 segundos:**
  - Filas de caída/error (`.email-row-failure`) con borde izquierdo rojo coral (`#F43F5E`) e icono de alerta `!`.
  - Filas de servicio restablecido (`.email-row-success`) con borde izquierdo verde esmeralda (`#10B981`) e icono `✓`.
- **Clases HTML/SCSS Descentralizadas:**
  - Cada columna y elemento de celda tiene su propia clase independiente en `src/app/home/home.page.scss` (`.col-head-*`, `.col-cell-*`, `.sender-box`, `.sender-name`, `.sender-email`, `.subject-box`, `.subject-title`, `.subject-subtitle`, `.folder-name`, `.received-date`).
- **Planificación Semanal y Vista Previa Excel Helpdesk:**
  - Selector de semanas del año (`YYYY-Www`) con indicador de ocupación (`[X/4 cupos]`) y bloqueo automático al alcanzar 4 solicitudes en una misma semana.
  - Campo **Postergado por proyecto** editable al planificar y en la hoja previa.
  - Modal **Vista Previa y Aprobación de Envío** que muestra el logo oficial (`src/assets/logo-helpdesk.png`) en `A1:B3` y las columnas `SEMANA | POSTERGADO POR PROYECTO | USUARIO | CORREO | ASUNTO` con edición en vivo antes de descargar el `.xlsx`, abrir el borrador en Outlook o aprobar el envío.

## Desarrollo local

Con el backend activo en `http://localhost:8000`:

```powershell
cd outlook-zimbra-frontend
npm install
npm start
```

Abre `http://localhost:8100` en el navegador y haz clic en **Conectar cuenta** o **Sincronizar bandeja** para enlazar con tu sesión local de Outlook.

## Personalización del Logo del Excel

El logo del reporte Excel se encuentra en:
`src/assets/logo-helpdesk.png`

Tanto la vista previa en pantalla como el generador `.xlsx` del backend leen automáticamente ese archivo y conservan su proporción original sin deformarlo.
