# Sistema de Gestión de Órdenes óptica

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?logo=google&logoColor=white)](https://developers.google.com/apps-script)
[![JavaScript ES6+](https://img.shields.io/badge/JavaScript-ES6%2B-F7DF1E?logo=javascript&logoColor=black)](https://developer.mozilla.org/docs/Web/JavaScript)
[![HTML5](https://img.shields.io/badge/HTML5-E34F26?logo=html5&logoColor=white)](https://developer.mozilla.org/docs/Web/Guide/HTML/HTML5)
[![CSS3](https://img.shields.io/badge/CSS3-1572B6?logo=css3&logoColor=white)](https://developer.mozilla.org/docs/Web/CSS)
[![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?logo=googlesheets&logoColor=white)](https://www.google.com/sheets/about/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Aplicación web interna para registrar y controlar órdenes ópticas desde la solicitud inicial hasta la entrega al paciente. El proyecto utiliza Google Apps Script como backend, Google Sheets como almacenamiento y una interfaz HTML/CSS/JavaScript servida mediante HtmlService.

## Funcionalidades

- Gestión de órdenes ópticas por etapas.
- Registro de pacientes, referencias y valores económicos.
- Control de facturación, abonos y saldos.
- Seguimiento del envío y recepción en laboratorio.
- Registro de entrega y responsable de la operación.
- Dashboard con órdenes del día y pendientes por fase.
- Historial de observaciones y trazabilidad de acciones.
- Generación de mensajes de WhatsApp con normalización de teléfonos.
- Anulación de órdenes con motivo y registro de auditoría.
- Bloqueo de operaciones concurrentes mediante `LockService`.
- Renderizado seguro de datos externos mediante APIs del DOM.

## 📸 Vista Previa del Sistema

> Sustituye las rutas de ejemplo por capturas o GIFs del entorno de demostración antes de publicar el repositorio.

![Dashboard de órdenes](https://drive.google.com/file/d/1ZKAjumDanb9U6K303pBujsjCLrk68088/view?usp=drive_link)

![Flujo de facturación y abonos](docs/billing-flow-preview.png)

![GIF del ciclo completo de una orden](docs/order-lifecycle-demo.gif)

## 🔄 Ciclo de Vida de una Orden Óptica

1. **Registro inicial:** se capturan los datos del paciente, referencias ópticas, valores y responsable del pago.
2. **Facturación y abono:** se administra el estado de cobro, los abonos acumulados y el saldo pendiente.
3. **Envío a laboratorio:** se registran el número de orden, el laboratorio responsable y la fecha de envío.
4. **Recepción y notificación:** se registra la llegada del producto y se prepara la notificación al paciente mediante WhatsApp.
5. **Entrega al paciente:** se registra la fecha, el estado de retiro y el responsable de la entrega.

## Buenas Prácticas y Optimizaciones

- **Control de concurrencia:** `LockService` evita que operaciones simultáneas sobrescriban filas o estados.
- **Lecturas en bloque:** las consultas de pendientes y métricas utilizan `getValues()` sobre rangos completos para reducir llamadas a `SpreadsheetApp`.
- **Prevención de XSS:** los datos provenientes de Sheets se incorporan al DOM mediante `textContent` y creación explícita de nodos.
- **Normalización de teléfonos:** los números se limpian y se normalizan antes de generar enlaces de WhatsApp, evitando duplicar el prefijo internacional.
- **Separación de responsabilidades:** el backend divide el procesamiento por fase y el frontend centraliza sus listeners mediante `addEventListener`.
- **Trazabilidad:** las operaciones de negocio conservan observaciones, responsables y marcas de auditoría.

## Arquitectura

La aplicación utiliza una arquitectura serverless y monolítica, adecuada para despliegues pequeños y medianos en Google Workspace.

```text
formulariosoptica/
├── Código.gs          # Backend, reglas de negocio y acceso a Sheets
├── Index.html         # Estructura de la interfaz y formularios
├── JavaScript.html    # Estado del frontend, eventos y comunicación con Apps Script
├── Styles.html        # Estilos CSS de la aplicación
└── README.md          # Documentación del proyecto
```

Flujo de comunicación:

```text
Navegador
   │
   ├── Index.html + Styles.html + JavaScript.html
   │
   └── google.script.run
          │
          ▼
Google Apps Script: Código.gs
          │
          ▼
Google Sheets
   ├── Hoja de seguimiento de órdenes
   └── Hoja de asesores
```

## Tecnologías

- Google Apps Script.
- HtmlService.
- Google Sheets API mediante `SpreadsheetApp` como persistencia.
- HTML5.
- CSS3.
- JavaScript ES6+.
- APIs DOM del navegador.
- `google.script.run` para comunicación cliente-servidor.
- `LockService` para control de concurrencia.
- URL pública de WhatsApp para composición de mensajes.

## Estructura de datos

La hoja principal se configura mediante el objeto `CONFIG` de `Código.gs` y utiliza las columnas A:AC para almacenar:

- Datos del paciente y de la orden.
- Valores de montura, lente y accesorios.
- Estado de facturación.
- Datos de laboratorio.
- Fechas de recepción y entrega.
- Historial de observaciones.
- Total, abonos y saldo pendiente.

Las filas de encabezado terminan en la fila 4. Los registros comienzan en la fila 5.

La hoja de asesores debe contener, como mínimo:

- Usuario.
- Contraseña.
- Nombre visible del asesor.

Para un entorno público, se recomienda reemplazar la autenticación basada en texto plano por identidad de Google Workspace o un mecanismo de autenticación administrado.

## Despliegue en Google Apps Script

1. Cree un proyecto nuevo en [Google Apps Script](https://script.google.com/).
2. Añada los archivos `Código.gs`, `Index.html`, `JavaScript.html` y `Styles.html`.
3. Cree o vincule una hoja de cálculo para la aplicación.
4. Configure las hojas con los nombres definidos en `CONFIG.HOJAS`.
5. Verifique que la hoja de seguimiento tenga sus encabezados en las primeras cuatro filas.
6. Configure los datos de asesores desde la fila 5 de la hoja correspondiente.
7. Revise los valores de `DATOS_MENSAJE` en `JavaScript.html` y sustitúyalos por la información institucional autorizada para el despliegue.
8. En Apps Script, seleccione **Implementar > Nueva implementación**.
9. Seleccione **Aplicación web** como tipo de implementación.
10. Configure quién ejecuta la aplicación y quién puede acceder según la política de su organización.
11. Autorice el acceso a la hoja de cálculo cuando Apps Script lo solicite.
12. Abra la URL de la aplicación y pruebe el login, registro, facturación, envío, recepción y entrega con datos no sensibles.

## Configuración pública

Este repositorio no debe contener:

- IDs reales de hojas o proyectos.
- Contraseñas o tokens.
- Números de teléfono de pacientes.
- Direcciones privadas o datos clínicos.
- Exportaciones reales de Google Sheets.

Los datos institucionales del mensaje de WhatsApp están representados por placeholders en el código. Sustitúyalos únicamente en el entorno de despliegue autorizado.

## Consideraciones de seguridad

- El backend valida los estados de facturación y protege las escrituras concurrentes.
- El frontend utiliza `textContent` y creación explícita de nodos para datos procedentes de Sheets.
- La autenticación actual depende de la hoja de asesores; para producción se recomienda un proveedor de identidad.
- Las funciones del backend deben desplegarse con permisos mínimos necesarios.
- Debe revisarse la política de `ALLOWALL` si la aplicación no necesita ser embebida en iframes.

## Estado del proyecto

El código está preparado para revisión pública como proyecto de portafolio. Antes de publicarlo, configure los valores institucionales del despliegue, revise los permisos de la hoja y ejecute pruebas funcionales con datos ficticios.

## Licencia

Este proyecto se distribuye bajo la licencia MIT. Consulta el archivo [LICENSE](LICENSE) para conocer los términos completos de uso, modificación y distribución.
