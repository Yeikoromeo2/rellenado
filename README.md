# Rellenado CRM · Complemento de Excel

Panel para rellenar el libro compartido **NUEVOS REGISTROS 2026 Actualizado** (hoja **BD-BBPE**) a partir de códigos de cliente (`Customer ID`).
Es 100 % estático: no usa servidores ni bases de datos. Todo se lee y se escribe en la memoria del navegador.

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Panel lateral (350 px) |
| `app.js` | Lógica de cruce y escritura (Office.js + SheetJS) |
| `manifest.xml` | Registro del complemento en Excel Online |
| `assets/` | Íconos |
| `difusion/` | Herramienta de difusión por WhatsApp (un solo HTML) · <https://yeikoromeo2.github.io/rellenado/difusion/> |
| `m365/` | Versión Microsoft 365: Office Scripts, listas de SharePoint, Power Automate y agente de Copilot. Ver [`m365/README.md`](m365/README.md) |

## 1. Publicar en GitHub Pages (una sola vez)

1. En GitHub, abre **Settings → Pages**.
2. En **Source** elige *Deploy from a branch*. En **Branch** elige `main` y la carpeta `/ (root)`. Guarda.
3. Espera 1 o 2 minutos y comprueba que abre <https://yeikoromeo2.github.io/rellenado/>.

> GitHub Pages gratis requiere que el repositorio sea **público**. El código no contiene datos de clientes.

## 2. Instalar el complemento en Excel Online

1. Descarga `manifest.xml` desde este repositorio.
2. Abre el libro en Excel para la web y ve a **Inicio → Complementos → Más complementos**.
3. Entra en **Mis complementos → Cargar mi complemento** y elige `manifest.xml`.
4. En la pestaña **Inicio** aparece el botón **Rellenado CRM**. Ábrelo.

Cada persona que vaya a usar el panel debe cargar el manifiesto. Si tu organización bloquea la carga manual, un administrador puede desplegarlo para todos desde **Centro de administración de Microsoft 365 → Configuración → Aplicaciones integradas**.

## 3. Uso

**Acción 1: fecha y responsable**
1. Pega los códigos (uno por línea). Se limpian solos los espacios, las líneas vacías y los duplicados.
2. La fecha es la de hoy por defecto. Se escribe como fecha real de Excel con formato `mm/dd/yyyy`.
3. Marca el responsable y pulsa **Ejecutar acción 1**. Se rellenan `FECHA CONTACTO` y `RESPONSABLE` en las filas con ese `Customer ID`, **solo si la celda está vacía**.

**Acción 2: cargar el archivo CRM**
1. Elige el `.xlsx` exportado desde la herramienta de difusión. Su columna `Código de Cliente` se cruza con `Customer ID`.
2. Solo se trasladan estas 6 columnas: `ESTADO CONTACTO`, `PRÓXIMA ACCIÓN`, `OBSERVACIÓN`, `NOMBRE DEL CONTACTO`, `CORREO` y `TELÉFONO`.
3. Pulsa **Rellenar información en Excel**. Solo se escriben las celdas que en la hoja están **vacías**.

**Varias hojas a la vez (BD-BBPE + Seguimiento semanal)**
- En ⚙ Configuración → *Hojas a revisar y rellenar*. Vienen marcadas BD-BBPE y cualquier hoja cuyo nombre contenga "seguimiento".
- Cada código se busca en **todas** las hojas marcadas y se rellena donde exista, con las mismas reglas.
- El resumen muestra qué se hizo en cada hoja y avisa si un código está repetido dentro de una hoja.
- **Códigos que no están en ninguna hoja:** aparecen listados con el botón **➕ Rellenar y agregar al final**.
  - Eliges la hoja y se escriben en la **primera fila vacía del final** (Customer ID + los datos de la acción).
  - Justo antes de escribir se vuelve a revisar todas las hojas marcadas, para **no duplicar** un código.
  - En la Acción 2 puedes incluir Razón Social y Número del CRM, **solo en esas filas nuevas**.
  - Se puede deshacer con el botón inferior.

**🎲 Tomar libres en dos filtros (flujo recomendado)**
1. **Complemento → pestaña ① Generar → Generar lista (sin registrar):** propone contactos libres según la meta por hoja (15 + 15). No escribe nada. Copia la lista.
2. **Difusión → 🔎 Validar lista del complemento:** pega la lista y te dice cuáles **sí están** en la base. **📋 Copiar y agregar a la lista del día** los copia y los suma a la *Lista del día* (x/30).
3. **Complemento → pestaña ② Registrar:** pega lo copiado, revisa fecha y responsable, y confirma en el cuadro del panel. Registra solo esos (fecha y responsable) y muestra cuántos faltan.
4. **Complemento → ① Generar diferencia:** propone solo lo que falta, sin repetir los ya propuestos. Repite 2–4 hasta 30.
5. **Difusión → ✅ Generar lote** desde la Lista del día. Todo lo demás sigue igual.
- *Tomar y registrar directo* hace todo en un paso, sin validar. *Reiniciar ronda* empieza de cero, sin tocar el Excel.

**📱 Celulares desde NUEVOS REGISTROS 2026 (v1.7)**
- Solo se rellena en **BD-BBPE** (columna `Customer ID`) y **Seguimiento semanal** (`ID DE TRAX` o `ID TRAX`).
- La hoja **NUEVOS REGISTROS 2026** solo se **lee**: su columna `CELULARES` se cruza por código, como un BUSCARV. Se elige en la tarjeta de hojas → *📱 Celulares desde*.
- Un celular es real si empieza en `9` o en `51`. Solo se quita el `51` y deben quedar **exactamente 9 dígitos**; si no, se descarta. Si la celda trae varios números, se usa el primero válido.
- Con *Solo con celular real* marcado, la lista propuesta trae solo esos contactos.
- **⬇ Excel** descarga el reporte con 3 columnas: `Código de Cliente`, `Razón Social`, `Celular`.
- En la difusión: **🔎 Validar lista del complemento → 📄 Cargar Excel del complemento**. Agrega a la base los que no tenga y completa los celulares que falten.

**Validar antes de rellenar (Acción 1)**
- **🔍 Validar** revisa los códigos en las hojas marcadas **sin escribir nada** y muestra un reporte: *"Se encontraron 15 de 17: 8 ya registrados, 7 nuevos y 2 no existen"*.
- Por cada código indica hoja, fila, fecha y responsable que ya tiene. Se puede filtrar, copiar los ya registrados o los nuevos, y descargar el reporte en Excel.
- Los "ya registrados" se pegan en la herramienta de difusión (Cola → 🧹 Quitar) para sacarlos del lote y completar la diferencia con contactos nuevos.

**Reglas de protección (las dos acciones)**
- `Customer ID` solo se usa para buscar la fila. Nunca se escribe en ella.
- `Razón Social` y `Número de Contacto` nunca se tocan.
- Nunca se sobrescribe una celda que ya tenga datos.
- Ninguna otra columna se modifica. Si algo intentara escribir fuera de estas reglas, la operación se cancela completa.

**Deshacer:** el botón inferior restaura los valores anteriores de la última acción. Excel no permite deshacer con Ctrl+Z los cambios hechos por un complemento.

## 4. Modo local (sin Excel Online)

Si abres <https://yeikoromeo2.github.io/rellenado/> directamente en Chrome o Edge:

1. Carga el Excel maestro desde tu PC.
2. Ejecuta las acciones.
3. Pulsa **Descargar Excel**.

El archivo no sale de tu navegador. La descarga conserva los datos y los formatos de número, pero puede perder colores, estilos y formatos condicionales (limitación de SheetJS). Para el libro compartido usa el complemento.

## Configuración (⚙)

Puedes cambiar desde el panel la hoja destino, los nombres de las columnas ID, de fecha y de responsable, y el formato de fecha. Los encabezados se buscan solos en las 10 primeras filas, sin importar mayúsculas, tildes ni espacios.

## Actualizar

Edita los archivos y súbelos a `main`. Si Excel muestra la versión anterior, cambia `app.js?v=1.0.0` en `index.html` para que el navegador no use la copia guardada.
