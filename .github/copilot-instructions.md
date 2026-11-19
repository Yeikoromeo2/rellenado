# Instrucciones para GitHub Copilot — repositorio "rellenado"

Proyecto comercial de difusión por WhatsApp y rellenado de un Excel compartido (Office 365). Todo el texto visible para el usuario va en **español**.

## Estructura
- `index.html`, `app.js`, `manifest.xml`: complemento de Excel "Rellenado CRM" (Office.js + SheetJS), servido desde GitHub Pages.
- `difusion/index.html`: herramienta de difusión en un solo archivo (HTML + CSS + JS, localStorage, SheetJS/ExcelJS por CDN).
- `m365/`: versión Microsoft 365. Office Scripts (`m365/scripts/*.ts`), `CRM_Motor.xlsx`, guías de listas de SharePoint, flujos de Power Automate y el agente de Copilot Studio.

## Reglas de negocio que NUNCA se deben romper
1. **Excel BD-BBPE:** solo se escriben celdas **vacías**; nunca se sobrescribe.
2. **Columnas que se pueden escribir:** solo FECHA CONTACTO, RESPONSABLE, ESTADO CONTACTO, PRÓXIMA ACCIÓN, OBSERVACIÓN, NOMBRE DEL CONTACTO, CORREO y TELÉFONO.
3. **Columnas que no se tocan:** Customer ID, Código de Cliente, Razón Social y Número / Celular.
4. **Candado final:** antes de escribir, cancelar toda la operación si alguna escritura cae fuera de las reglas 1–3.
5. **Encabezados y códigos:** los encabezados se detectan en las 10 primeras filas, sin importar mayúsculas, tildes ni espacios. Los códigos se comparan normalizados (trim, sin espacios, mayúsculas, "123.0" = "123").
6. **Textos que empiezan con "=":** se escriben como texto (con apóstrofo), nunca como fórmula.
7. **Mensajes a clientes:**
   - Se dirigen a la **razón social**, los firma el **ejecutivo** del lote y siempre piden nombre del encargado de compras, correo y teléfono directo.
   - **Nunca** incluyen el código de cliente ni el teléfono; son de uso interno.
   - Dos contactos seguidos no reciben la misma plantilla ni la misma combinación.
8. **Protección de envíos:**
   - Lista No contactar.
   - Consentimiento (opcional).
   - Máximo 1 mensaje por número cada N días (7 por defecto).
   - Máximo 30 envíos en cualquier periodo de 24 h (ventana móvil, no por día calendario), con intervalos irregulares de 7–45 min.
   - Jornada 09:00–18:00, de lunes a viernes.
   - Opción de salida "respondan BAJA".
9. **Envío:** lo hace una persona (abre el chat y presiona Enter). No agregar automatizaciones de WhatsApp no oficiales (Baileys, whatsapp-web.js, WAHA, etc.); el envío automático solo con la API oficial de WhatsApp Business (Meta).

## Office Scripts (`m365/scripts`)
- **TypeScript estricto:** sin `any` explícito ni implícito, sin `import`, y una sola función `main(workbook: ExcelScript.Workbook, ...)` por archivo.
- **Archivos autocontenidos:** cada archivo se pega tal cual en Excel para la web; las utilidades compartidas se duplican en cada archivo.
- **Rendimiento:** leer rangos completos una vez y escribir agrupando celdas contiguas. Evitar llamadas a la API dentro de bucles grandes.
- **Para Power Automate:** devolver un objeto con `ok`, `mensaje` y, si se escribe en el Excel, `cambios` (para deshacer).

## Estilo
- Nombres de variables y funciones en español.
- Mensajes de error claros, que digan qué hacer.
- Para cambiar reglas de negocio, actualizar a la vez el complemento (`app.js`), la herramienta de difusión (`difusion/index.html`) y los Office Scripts, para que se comporten igual.
