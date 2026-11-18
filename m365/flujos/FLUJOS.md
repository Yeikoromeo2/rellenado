# Flujos de Power Automate

Son 8 flujos. Todos usan solo conectores **estándar** de Microsoft 365: SharePoint, Excel Online (Business), Microsoft Teams y Copilot Studio.

**Antes de empezar**
- **Zona horaria:** en todas las expresiones se usa `SA Pacific Standard Time`, que es Lima (UTC-5).
- **Nombres de acciones:** cambia el nombre de cada acción **exactamente** como se indica, por ejemplo `Contacto` o `MensajeSeguro`. Las expresiones dependen de esos nombres. En las expresiones, los espacios se convierten en `_`.
- **Archivos:** los flujos usan dos archivos guardados en una biblioteca de SharePoint (o en OneDrive):
  - `CRM_Motor.xlsx`: el motor de mensajes y cadencia.
  - El libro **NUEVOS REGISTROS 2026 Actualizado**, que contiene la hoja **BD-BBPE**.
- **Scripts:** crea los 4 Office Scripts **antes** que los flujos (ver `../README.md`, paso 2).

| # | Nombre del flujo | Disparador | Lo usa |
|---|---|---|---|
| F1 | CRM - Generar lote | Copilot (agente) o manual | Agente / ejecutivo |
| F2 | CRM - Turnos de envío | Programado cada 5 min, 09:00–18:00 lun–vie | Automático |
| F3 | CRM - Marcar BAJA | Copilot (agente) | Agente |
| F4 | CRM - Tipificar | Copilot (agente) | Agente |
| F5 | CRM - Estado de campaña | Copilot (agente) | Agente |
| F6 | CRM - Buscar cliente | Copilot (agente) | Agente |
| F7 | CRM - Rellenar Excel BD-BBPE | Programado 18:10 lun–vie + Copilot | Automático / agente |
| F8 | CRM - Deshacer relleno | Copilot (agente) | Agente |

> **Flujos que llama Copilot (F1, F3–F8):** créalos desde **Copilot Studio → tu agente → Herramientas → Agregar herramienta → Nuevo flujo de agente**. Así ya traen el disparador *"Cuando un agente llama al flujo"* y la acción *"Responder al agente"*.
> F7 además se programa: duplícalo y cambia el disparador por *Periodicidad*, o crea uno aparte que lo invoque.

---

## F1 · CRM - Generar lote
**Entradas** (en el disparador *Cuando un agente llama al flujo*):
- `Ejecutivo` (texto): Cinthya Sandoval, Fiorella Viena, Harold Sovero o Lorena Suncion.
- `Tamano` (número, opcional; 30 por defecto).

**Acciones:**
1. **Obtener elementos** (SharePoint) → lista `CRM_Clientes`. Renombra la acción como `Clientes`.
   *Configuración → Paginación*: activada, umbral **20000**.
2. **Seleccionar** → renómbrala `SelClientes`. De: `body('Clientes')?['value']`. Asignación:

   | Clave | Valor (expresión) |
   |---|---|
   | codigo | `item()?['Title']` |
   | razon | `item()?['RazonSocial']` |
   | telefono | `item()?['Telefono']` |
   | consentimiento | `coalesce(item()?['Consentimiento']?['Value'], '')` |
   | ciudad | `coalesce(item()?['Ciudad'], '')` |
3. **Obtener elementos** → `CRM_Gestiones`, paginación activada. Renómbrala `Gestiones`.
4. **Seleccionar** `SelAsignados`. De: `body('Gestiones')?['value']`. Cambia el modo a **texto** (icono a la derecha) y pon `item()?['Title']`.
5. **Filtrar matriz** `Enviados`. De: `body('Gestiones')?['value']`. Condición: `item()?['EstadoEnvio']?['Value']` **es igual a** `Enviado`.
6. **Seleccionar** `SelEnvios`. De: `body('Enviados')`:

   | Clave | Valor |
   |---|---|
   | telefono | `item()?['Telefono']` |
   | fecha | `formatDateTime(convertFromUtc(item()?['FechaEnvio'],'SA Pacific Standard Time'),'yyyy-MM-dd')` |
7. **Filtrar matriz** `EnviadosHoy`. De: `body('SelEnvios')`. Condición: `item()?['fecha']` es igual a `formatDateTime(convertFromUtc(utcNow(),'SA Pacific Standard Time'),'yyyy-MM-dd')`.
8. **Obtener elementos** → `CRM_NoContactar`. Renómbrala `NoContactar`.
9. **Seleccionar** `SelBajasTel` (modo texto) con `item()?['Title']`. Luego **Seleccionar** `SelBajasCod` (modo texto) con `coalesce(item()?['Codigo'],'')`.
10. **Obtener elementos** → `CRM_Lotes`. *Ordenar por*: `Numero desc`. *Recuento superior*: `1`. Renómbrala `UltimoLote`.
11. **Redactar** `Entrada`:
```json
{
  "ahoraLocal": "@{formatDateTime(convertFromUtc(utcNow(),'SA Pacific Standard Time'),'yyyy-MM-dd HH:mm')}",
  "ejecutivo": "@{triggerBody()?['text']}",
  "tamano": @{coalesce(triggerBody()?['number'], 30)},
  "loteNumero": @{add(coalesce(first(body('UltimoLote')?['value'])?['Numero'], 0), 1)},
  "clientes": @{body('SelClientes')},
  "asignados": @{body('SelAsignados')},
  "noContactar": @{union(body('SelBajasTel'), body('SelBajasCod'))},
  "ultimosEnvios": @{body('SelEnvios')},
  "enviadosHoy": @{length(body('EnviadosHoy'))}
}
```
   > `triggerBody()?['text']` y `['number']` son los nombres internos de las entradas. Si tu disparador los llama distinto (`text_1`, etc.), usa el contenido dinámico *Ejecutivo* y *Tamano*.
12. **Ejecutar script** (Excel Online Business) → renómbrala `Motor`:
    - Ubicación y biblioteca donde está `CRM_Motor.xlsx`.
    - Script: **GenerarLote**.
    - entradaJson: `@{string(outputs('Entrada'))}`.
13. **Condición**: `body('Motor')?['result']?['ok']` es igual a `true`.
    - **Sí:**
      - a. **Crear elemento** en `CRM_Lotes`:

        | Campo | Valor |
        |---|---|
        | Title | `body('Motor')?['result']?['loteNombre']` |
        | Numero | `outputs('Entrada')?['loteNumero']` |
        | Fecha | `utcNow()` |
        | Ejecutivo | `body('Motor')?['result']?['ejecutivo']` |
        | Contactos | `body('Motor')?['result']?['resumen']?['contactos']` |
        | Cierre | `body('Motor')?['result']?['resumen']?['cierre']` |
        | Resumen | `body('Motor')?['result']?['mensaje']` |
      - b. **Aplicar a cada uno**: `body('Motor')?['result']?['items']` → **Crear elemento** en `CRM_Gestiones`:

        | Campo | Valor |
        |---|---|
        | Title | `items('Aplicar_a_cada_uno')?['codigo']` |
        | Lote | `body('Motor')?['result']?['loteNombre']` |
        | Ejecutivo | `items('Aplicar_a_cada_uno')?['ejecutivo']` |
        | CorreoEjecutivo | `items('Aplicar_a_cada_uno')?['correoEjecutivo']` |
        | RazonSocial | `items('Aplicar_a_cada_uno')?['razonSocial']` |
        | Telefono | `items('Aplicar_a_cada_uno')?['telefono']` |
        | Mensaje | `items('Aplicar_a_cada_uno')?['mensaje']` |
        | PesoBytes | `items('Aplicar_a_cada_uno')?['bytes']` |
        | Caracteres | `items('Aplicar_a_cada_uno')?['caracteres']` |
        | HoraProgramada | `items('Aplicar_a_cada_uno')?['horaProgramada']` |
        | EstadoEnvio Value | `items('Aplicar_a_cada_uno')?['estado']` |
        | EnlaceWhatsApp | `items('Aplicar_a_cada_uno')?['enlaceWhatsApp']` |
      - c. **Responder al agente**, salida `Resultado` (texto): `body('Motor')?['result']?['mensaje']`.
    - **No:** **Responder al agente** con `body('Motor')?['result']?['mensaje']`.

---

## F2 · CRM - Turnos de envío (tarjeta en Teams)
**Disparador:** *Periodicidad*.
- Intervalo 5, frecuencia Minuto.
- Zona horaria (UTC-05:00) Bogotá, Lima, Quito.
- En "Estas horas": 9 a 17.
- Días: lunes a viernes.

**Acciones:**
1. **Obtener elementos** `CRM_Gestiones` → renombra `Turno`:
   - Consulta de filtro: `(EstadoEnvio eq 'Pendiente' or EstadoEnvio eq 'Reprogramado') and HoraProgramada le '@{utcNow()}'`
   - Ordenar por: `HoraProgramada asc` · Recuento superior: `1`
2. **Condición**: `length(body('Turno')?['value'])` es mayor que `0`. Todo lo siguiente va en **Sí**:
3. **Redactar** `Contacto`: `first(body('Turno')?['value'])`.
4. **Obtener elementos** `CRM_Gestiones` → renombra `TarjetaAbierta`. Filtro: `EstadoEnvio eq 'Notificado' and CorreoEjecutivo eq '@{outputs('Contacto')?['CorreoEjecutivo']}'` · Recuento superior 1.
   **Condición**: si `length(body('TarjetaAbierta')?['value'])` es mayor que 0, pon **Finalizar** (correcto). Así cada ejecutivo tiene una sola tarjeta a la vez.
5. **Obtener elementos** `CRM_NoContactar` → renombra `EsBaja`. Filtro: `Title eq '@{outputs('Contacto')?['Telefono']}'`.
   Si hay resultados: **Actualizar elemento** (Id = `outputs('Contacto')?['ID']`) con EstadoEnvio = `No contactar` y luego **Finalizar**.
6. **Actualizar elemento** (Id del contacto): EstadoEnvio = `Notificado`. Esto evita que otra ejecución tome el mismo contacto.
7. Tres **Redactar** para escapar el texto que va dentro de la tarjeta:
   - `MensajeSeguro`: `replace(replace(replace(replace(outputs('Contacto')?['Mensaje'],'\','\\'),'"','\"'),decodeUriComponent('%0D'),''),decodeUriComponent('%0A'),'\n')`
   - `RazonSegura`: `replace(replace(outputs('Contacto')?['RazonSocial'],'\','\\'),'"','\"')`
   - `HoraTexto`: `formatDateTime(convertFromUtc(outputs('Contacto')?['HoraProgramada'],'SA Pacific Standard Time'),'HH:mm')`
8. **Publicar tarjeta adaptable y esperar una respuesta** (Microsoft Teams) → renombra `Tarjeta`:
   - Publicar como: *Bot de flujo* · Publicar en: *Chat con el bot de flujo*.
   - Destinatario: `outputs('Contacto')?['CorreoEjecutivo']`.
   - Tarjeta adaptable: pega el contenido de `../tarjetas/turno.json` tal cual. Sus expresiones `@{outputs('...')}` se resuelven solas.
   - Mensaje de actualización: `Registrado ✔`.
   - *Configuración → Tiempo de espera*: `PT8H`. Si se agota, el contacto vuelve a Pendiente (paso 10).
9. **Modificador** sobre `body('Tarjeta')?['submitActionId']`:
   - `enviado` → **Actualizar elemento**: EstadoEnvio = `Enviado`, FechaEnvio = `utcNow()`, FechaGestion = `utcNow()`.
   - `despues` → **Actualizar elemento**: EstadoEnvio = `Pendiente`, HoraProgramada = `addMinutes(utcNow(),30)`.
   - `omitir` → **Actualizar elemento**: EstadoEnvio = `Omitido`, FechaGestion = `utcNow()`.
   - `baja` → **Actualizar elemento**: EstadoEnvio = `No contactar`. Luego **Crear elemento** en `CRM_NoContactar`: Title = `outputs('Contacto')?['Telefono']`, Codigo = `outputs('Contacto')?['Title']`, Motivo = `Pidió BAJA (tarjeta)`, Fecha = `utcNow()`.
10. Para la tarjeta sin respuesta: agrega una rama paralela después de `Tarjeta` con *Configurar ejecución después de → se agotó el tiempo de espera / con errores*. En ella, **Actualizar elemento** con EstadoEnvio = `Pendiente`.

> En el paso 1 no se valida la hora de cierre porque el motor ya programó todo dentro de la jornada. La periodicidad (9 a 17 h) evita que lleguen tarjetas fuera de horario.

---

## F3 · CRM - Marcar BAJA
**Entradas:** `CodigoOTelefono` (texto), `Motivo` (texto, opcional).
1. **Obtener elementos** `CRM_Gestiones`:
   - Filtro: `Title eq '@{triggerBody()?['text']}' or Telefono eq '@{triggerBody()?['text']}'`
   - Recuento superior: 1 · ordenar por `Created desc`.
2. Si no hay resultados → **Responder al agente**: "No encontré ese código o teléfono."
3. **Crear elemento** `CRM_NoContactar`:
   - Title = `first(body('Obtener_elementos')?['value'])?['Telefono']`
   - Codigo = `first(...)?['Title']`
   - Motivo = `coalesce(triggerBody()?['text_1'],'Pidió BAJA')`
   - Fecha = `utcNow()`
4. **Obtener elementos** `CRM_Gestiones`. Filtro: `Telefono eq '@{first(body('Obtener_elementos')?['value'])?['Telefono']}' and (EstadoEnvio eq 'Pendiente' or EstadoEnvio eq 'Reprogramado')`.
   Luego **Aplicar a cada uno** → **Actualizar elemento** con EstadoEnvio = `No contactar`.
5. **Responder al agente**: `Marcado como No contactar. Se retiraron @{length(body('Obtener_elementos_2')?['value'])} envío(s) pendiente(s).`

## F4 · CRM - Tipificar
**Entradas** (todas texto, opcionales salvo el código): `Codigo`, `EstadoContacto`, `ProximaAccion`, `Observacion`, `NombreContacto`, `Correo`, `TelefonoDirecto`.
1. **Obtener elementos** `CRM_Gestiones`. Filtro: `Title eq '@{triggerBody()?['text']}'` · ordenar `Created desc` · recuento 1.
2. **Redactar** `Actual`: `first(body('Obtener_elementos')?['value'])`.
3. **Actualizar elemento** (Id = `outputs('Actual')?['ID']`). Cada campo conserva su valor si la entrada viene vacía:
   - EstadoContacto Value = `if(empty(triggerBody()?['text_1']), outputs('Actual')?['EstadoContacto']?['Value'], triggerBody()?['text_1'])`
   - Usa el mismo patrón para ProximaAccion, Observacion, NombreContacto, Correo y TelefonoDirecto.
   - FechaGestion = `utcNow()`
4. **Responder al agente**: "Gestión de @{triggerBody()?['text']} actualizada."

## F5 · CRM - Estado de campaña
**Entrada:** `Fecha` (texto AAAA-MM-DD, opcional; hoy por defecto).
1. **Obtener elementos** `CRM_Gestiones`. Filtro: `HoraProgramada ge '@{convertToUtc(formatDateTime(convertFromUtc(utcNow(),'SA Pacific Standard Time'),'yyyy-MM-dd'),'SA Pacific Standard Time')}'` (inicio del día en Lima), con paginación.
2. Usa un **Filtrar matriz** por cada estado: Enviado, Pendiente, Reprogramado, No contactar y Omitido. También cuenta las gestiones con `EstadoContacto` no vacío.
3. **Responder al agente** con un texto, por ejemplo:
   `Hoy: @{length(body('Enviado'))} enviados, @{length(body('Pendiente'))} pendientes, @{length(body('Reprogramado'))} reprogramados, @{length(body('Bajas'))} bajas, @{length(body('Tipificados'))} tipificados.`

## F6 · CRM - Buscar cliente
**Entrada:** `Texto` (código, razón social o teléfono).
1. **Obtener elementos** `CRM_Gestiones` · recuento 5 · ordenar `Created desc`. Filtro:
   `substringof('@{triggerBody()?['text']}',Title) or substringof('@{triggerBody()?['text']}',RazonSocial) or substringof('@{triggerBody()?['text']}',Telefono)`
2. **Seleccionar** (modo texto):
   `concat(item()?['Title'],' · ',item()?['RazonSocial'],' · ',item()?['EstadoEnvio']?['Value'],' · ',coalesce(item()?['EstadoContacto']?['Value'],'sin tipificar'),' · ',coalesce(item()?['NombreContacto'],''),' ',coalesce(item()?['Correo'],''))`
3. **Responder al agente**: `join(body('Seleccionar'), decodeUriComponent('%0A'))`.

---

## F7 · CRM - Rellenar Excel BD-BBPE
**Disparador:** *Periodicidad* diaria a las **18:10**, de lunes a viernes, zona Lima. Puedes hacer una copia de este flujo como flujo de agente para pedirlo desde Copilot.

1. **Inicializar variable** `Cambios` (Matriz) = `[]` y **Inicializar variable** `Resultado` (Cadena) = vacío.
2. **Redactar** `HoyLima`: `formatDateTime(convertFromUtc(utcNow(),'SA Pacific Standard Time'),'yyyy-MM-dd')`.
3. **Obtener elementos** `CRM_Gestiones` → renombra `EnviadosHoy`. Filtro: `EstadoEnvio eq 'Enviado' and FechaEnvio ge '@{convertToUtc(outputs('HoyLima'),'SA Pacific Standard Time')}'`, con paginación.
4. **Seleccionar** `Ejecutivos` (modo texto): `item()?['Ejecutivo']`. Después, **Redactar** `EjecutivosUnicos`: `union(body('Ejecutivos'), body('Ejecutivos'))`.
5. **Aplicar a cada uno** sobre `outputs('EjecutivosUnicos')`:
   - **Filtrar matriz** `DelEjecutivo`. De: `body('EnviadosHoy')?['value']`. Condición: `item()?['Ejecutivo']` es igual a `items('Aplicar_a_cada_uno')`.
   - **Seleccionar** `Codigos` (modo texto): `item()?['Title']`.
   - **Ejecutar script** → archivo: libro **NUEVOS REGISTROS 2026 Actualizado** · script: **RellenarAccion1** → renombra `Accion1`:
     - codigos: `join(body('Codigos'), decodeUriComponent('%0A'))`
     - responsable: `items('Aplicar_a_cada_uno')`
     - fecha: `outputs('HoyLima')`
   - **Establecer variable** `Cambios` = `union(variables('Cambios'), body('Accion1')?['result']?['cambios'])`. Las variables no se pueden autorreferenciar: usa primero un **Redactar** con esa expresión y luego asigna su salida a la variable.
   - **Anexar a variable de cadena** `Resultado`: `@{body('Accion1')?['result']?['mensaje']}` seguido de un salto de línea.
6. **Obtener elementos** `CRM_Gestiones` → renombra `Gestionados`. Filtro: `FechaGestion ge '@{addDays(utcNow(),-14)}'`, con paginación.
7. **Seleccionar** `FilasCRM`. De: `body('Gestionados')?['value']`. Asignación (claves **exactas**):

   | Clave | Valor |
   |---|---|
   | Código de Cliente | `item()?['Title']` |
   | ESTADO CONTACTO | `coalesce(item()?['EstadoContacto']?['Value'],'')` |
   | PRÓXIMA ACCIÓN | `coalesce(item()?['ProximaAccion']?['Value'],'')` |
   | OBSERVACIÓN | `coalesce(item()?['Observacion'],'')` |
   | NOMBRE DEL CONTACTO | `coalesce(item()?['NombreContacto'],'')` |
   | CORREO | `coalesce(item()?['Correo'],'')` |
   | TELÉFONO | `coalesce(item()?['TelefonoDirecto'],'')` |
8. **Ejecutar script** → mismo libro · script **RellenarAccion2** → renombra `Accion2`. filasJson: `string(body('FilasCRM'))`.
9. **Crear elemento** en `CRM_RegistroRelleno`:

   | Campo | Valor |
   |---|---|
   | Title | `Relleno @{formatDateTime(convertFromUtc(utcNow(),'SA Pacific Standard Time'),'dd/MM/yyyy HH:mm')}` |
   | Fecha | `utcNow()` |
   | Resultado | `@{variables('Resultado')}@{body('Accion2')?['result']?['mensaje']}` |
   | CambiosJson | `string(union(variables('Cambios'), body('Accion2')?['result']?['cambios']))` |
10. **Publicar mensaje en un chat o canal** (Teams) con el resumen, o **Responder al agente** si lo llamó Copilot.

> Repetir F7 es seguro: los scripts **solo llenan celdas vacías**, así que lo que ya está escrito no se vuelve a tocar.

## F8 · CRM - Deshacer relleno
1. **Obtener elementos** `CRM_RegistroRelleno`. Filtro: `Deshecho eq 0` · ordenar `Created desc` · recuento 1.
2. **Ejecutar script** → libro BD-BBPE · script **DeshacerRelleno** · cambiosJson: `first(body('Obtener_elementos')?['value'])?['CambiosJson']`.
3. **Actualizar elemento** del registro: Deshecho = Sí.
4. **Responder al agente**: `body('Ejecutar_script')?['result']?['mensaje']`.
