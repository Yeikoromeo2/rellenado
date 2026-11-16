# Rellenado CRM · Complemento de Excel

Panel para rellenar el libro compartido **NUEVOS REGISTROS 2026 Actualizado** (hoja **BD-BBPE**) a partir de códigos de cliente (`Customer ID`).
Es 100 % estático: no usa servidores ni bases de datos. Todo se lee y se escribe en la memoria del navegador.

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Panel lateral (350 px) |
| `app.js` | Lógica de cruce y escritura (Office.js + SheetJS) |
| `manifest.xml` | Registro del complemento en Excel Online |
| `assets/` | Íconos |

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
3. Marca el responsable y pulsa **Ejecutar acción 1**. Se rellenan `FECHA CONTACTO` y `RESPONSABLE` en todas las filas con ese `Customer ID`.

**Acción 2: cargar el archivo CRM**
1. Elige el `.xlsx` exportado desde la herramienta de difusión. Su columna `Código de Cliente` se cruza con `Customer ID`.
2. Revisa qué columnas se van a trasladar. Solo aparecen marcadas las que existen en la hoja con el mismo encabezado. `Razón Social` viene desmarcada por defecto.
3. Pulsa **Rellenar información en Excel**. Las celdas vacías del archivo CRM **no borran** lo que ya hay en la hoja.

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
