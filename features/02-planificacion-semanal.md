# Feature 2: Planificación Semanal (Máx. 4 por Semana) y Formato Excel Helpdesk

## Objetivo

Asignar solicitudes a semanas de trabajo (con un cupo máximo de 4 solicitudes por semana), agendarlas en el calendario de Outlook y generar el reporte oficial en **Excel (`.xlsx`) de Helpdesk** con vista previa interactiva y aprobación del administrador antes de enviarlo por correo.

## Historia de usuario

Como administrador único, quiero planificar hasta 4 solicitudes por semana, registrar si alguna fue postergada por proyecto, visualizar en pantalla cómo quedará el cuadro de Excel con el logo de Helpdesk y aprobar su envío o descarga.

## Criterios de aceptación implementados

- **Selección por semana ISO (`YYYY-Www`):** El modal de planificación permite elegir entre las próximas 20 semanas o cualquier semana del calendario, calculando automáticamente el bloque laboral de lunes (09:00) a viernes (18:00).
- **Límite de 4 solicitudes por semana:** Cada semana muestra su ocupación actual (`[X/4 cupos]`). Al alcanzar `4/4`, el sistema sugiere automáticamente la siguiente semana libre e impide sobrecargar una semana llena tanto en frontend como en backend (`HTTP 409`).
- **Creación de cita en Outlook local:** Se registra en el calendario de Outlook mediante `pywintypes.Time` para garantizar compatibilidad con configuraciones regionales de Windows en español.
- **Formato Excel Helpdesk (`.xlsx`) con Logo Proporcional:**
  - Celdas `A1:B3` combinadas con el logo oficial (`src/assets/logo-helpdesk.png`) anclado mediante `<xdr:oneCellAnchor>` respetando su relación de aspecto exacta sin deformarse.
  - Fila 4 de separación en blanco.
  - Fila 5 con encabezados: `SEMANA | POSTERGADO POR PROYECTO | USUARIO | CORREO | ASUNTO`.
- **Vista Previa Interactiva y Aprobación:**
  - Antes de enviar, el administrador visualiza la hoja de Excel en pantalla, puede editar cualquier celda o destinatario y elegir entre:
    1. **Solo Descargar Excel (`.xlsx`)**
    2. **Abrir Borrador en Outlook (`mode="display"`)**
    3. **Aprobar y Enviar Ahora (`mode="send"`)**
