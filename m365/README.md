# CRM Difusión en Microsoft 365 (sin servidor, sin instalaciones)

Esta carpeta lleva a Microsoft 365 la herramienta de difusión y el complemento de rellenado, con un **agente de Copilot** para manejarlo todo desde Teams.

```
Listas de SharePoint (base de datos)  ← Clientes · Gestiones · Lotes · No contactar · Registro
          ▲                 ▲
Power Automate (en la nube, arranca solo)
  F1 Generar lote ──► Office Script GenerarLote (CRM_Motor.xlsx): mensajes únicos + horas programadas
  F2 Turnos (09:00–18:00) ──► Tarjeta en Teams: [Abrir en WhatsApp] [Ya lo envié] [Pidió BAJA]
  F7 18:10 ──► Office Scripts RellenarAccion1 / RellenarAccion2 en el Excel BD-BBPE
          ▲
Agente de Copilot "CRM Difusión" (Teams / Microsoft 365 Copilot)
  "Genera el lote de hoy para Harold" · "¿Cómo vamos?" · "Marca BAJA al C00123" · "Pasa la gestión al Excel"
```

## Contenido
| Ruta | Qué es |
|---|---|
| `scripts/GenerarLote.ts` | Motor de lotes y mensajes: Spintax, bloques, firma del ejecutivo, solicitud de datos, cadencia 09:00–18:00, tope diario, frecuencia y No contactar |
| `scripts/RellenarAccion1.ts` | Rellena FECHA CONTACTO y RESPONSABLE (mismas reglas del complemento) |
| `scripts/RellenarAccion2.ts` | Rellena las 6 columnas de gestión (mismas reglas del complemento) |
| `scripts/DeshacerRelleno.ts` | Deshace el último relleno sin pisar cambios hechos después |
| `CRM_Motor.xlsx` | Configuración editable: Plantillas, Bloques, Ejecutivos, Ajustes, Respaldos y hoja de Prueba |
| `listas/LISTAS.md` | Columnas exactas de las 5 listas de SharePoint, con plantillas CSV |
| `flujos/FLUJOS.md` | Los 8 flujos de Power Automate paso a paso, con expresiones listas para copiar |
| `tarjetas/turno.json` | Tarjeta de Teams que llega en cada turno de envío |
| `agente/AGENTE.md` | Instrucciones y herramientas del agente de Copilot Studio |

## Reglas que se mantienen (iguales a la versión actual)
- **Excel BD-BBPE:**
  - Solo se rellenan **celdas vacías**: nunca se sobrescribe.
  - **Customer ID, Razón Social y Número nunca se tocan.**
  - Columnas que se rellenan: FECHA CONTACTO (`mm/dd/aaaa`), RESPONSABLE, ESTADO CONTACTO, PRÓXIMA ACCIÓN, OBSERVACIÓN, NOMBRE DEL CONTACTO, CORREO y TELÉFONO.
  - Si falta una columna, el script avisa y no crea nada.
  - **Tu libro no cambia:** no se agregan hojas ni columnas. El historial para deshacer se guarda en SharePoint.
- **Mensajes:**
  - Se dirigen a la **razón social** y los firma el **ejecutivo** del lote.
  - Siempre **piden nombre del encargado de compras, correo y teléfono directo**.
  - **Nunca incluyen códigos de cliente ni teléfonos.**
  - Llevan la opción de salida "respondan BAJA".
- **Protección:**
  - No contactar.
  - Consentimiento (opcional).
  - Máximo 1 mensaje por número cada 7 días.
  - Tope de 40 envíos diarios.
  - Jornada 09:00–18:00, de lunes a viernes.
- **Envío:** lo hace una persona. La tarjeta de Teams abre el chat con el mensaje listo y el ejecutivo presiona Enter.

## Instalación (orden recomendado)

### Paso 0 · Confirmar con tu administrador
- Que puedas usar **Office Scripts** en Excel para la web. En el Centro de administración: Configuración → Configuración de la organización → Office Scripts.
- Que puedas crear flujos de **Power Automate** con conectores estándar.
- Que puedas crear y publicar agentes en **Copilot Studio** y en el canal de Microsoft 365 Copilot.

### Paso 1 · Sitio y listas
1. Crea un sitio de SharePoint para el equipo, por ejemplo *CRM Comercial*.
2. Crea las 5 listas siguiendo `listas/LISTAS.md`.
3. Carga tus clientes en `CRM_Clientes`.

### Paso 2 · Archivos y Office Scripts
1. Sube `CRM_Motor.xlsx` a la biblioteca *Documentos* del sitio.
2. En la hoja **Ejecutivos**, completa el **correo de Teams** de cada ejecutivo. A ese correo le llegan sus tarjetas.
3. Abre `CRM_Motor.xlsx` en Excel para la web y ve a **Automatizar → Nuevo script**. Borra el contenido, pega `scripts/GenerarLote.ts` y guárdalo con el nombre **GenerarLote**.
4. **Prueba el motor:** **Automatizar → GenerarLote → Ejecutar**, sin parámetros. Se llena la hoja **Resultado** con 12 mensajes de ejemplo.
5. Abre el libro **NUEVOS REGISTROS 2026 Actualizado** y crea del mismo modo **RellenarAccion1**, **RellenarAccion2** y **DeshacerRelleno**.
6. **Compartir los scripts:** en cada script ve a **…→ Compartir con otros usuarios de este libro**, para que el equipo y los flujos los puedan usar.

> Puedes ejecutar **RellenarAccion1** a mano desde Excel. Pide los códigos, el responsable y la fecha (AAAA-MM-DD), igual que el complemento.

### Paso 3 · Flujos
Crea F1 a F8 siguiendo `flujos/FLUJOS.md`. Empieza por F1 y F2, que son el corazón del envío.

### Paso 4 · Agente de Copilot
Sigue `agente/AGENTE.md`: crea el agente, pega las instrucciones, agrega los flujos como herramientas y publícalo en Teams y Microsoft 365 Copilot.

### Paso 5 · Prueba completa
1. En Teams, escribe al agente: **"Genera el lote de hoy para [tu nombre], con 3 contactos"** y confirma.
2. Revisa en `CRM_Gestiones` que aparezcan 3 registros con mensaje y hora programada.
3. Espera la tarjeta de turno en Teams. Pulsa **Abrir en WhatsApp**, envía el mensaje y pulsa **Ya lo envié**.
4. Pídele al agente: **"Tipifica [código] como Cotizado – seguimiento"**.
5. Pídele: **"Pasa la gestión de hoy al Excel"** y revisa BD-BBPE. Solo deben haberse llenado celdas vacías.
6. Pídele: **"Deshaz el último relleno"** y comprueba que esas celdas vuelven a quedar vacías.

## Límites a tener en cuenta
- **Office Scripts desde Power Automate:** cada ejecución tiene límite de tiempo y hay un máximo de ejecuciones diarias por usuario. Para lotes de 30–40 contactos por ejecutivo alcanza de sobra.
- **El envío sigue siendo manual (Enter en WhatsApp).** Para enviar automáticamente con la API oficial de WhatsApp (Meta) se necesita cuenta de Meta Business verificada y conectores **premium** de Power Automate.
- **Licencias:** algunas funciones de Copilot Studio dependen del licenciamiento de tu empresa. Revísalo con tu administrador.
- **Herramienta actual:** la herramienta HTML (`../difusion/`) y el complemento de Excel siguen funcionando y usan las mismas reglas. Pueden convivir mientras migras.
