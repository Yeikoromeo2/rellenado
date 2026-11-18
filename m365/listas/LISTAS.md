# Listas de SharePoint (base de datos del CRM)

Crea un **sitio de SharePoint** para el equipo (por ejemplo, *CRM Comercial*) y dentro de él estas 5 listas.

> **Muy importante:** crea cada columna con el **nombre exacto de la tabla** (sin tildes ni espacios). Ese es su *nombre interno* y los flujos lo usan.
> Después puedes cambiar el nombre que se muestra (por ejemplo, `RazonSocial` → "Razón Social"); el nombre interno no cambia.
>
> La columna **Title** viene creada en toda lista. Solo cámbiale el nombre visible como se indica.

## 1. CRM_Clientes — universo de clientes
| Columna | Tipo | Notas |
|---|---|---|
| Title | (existente) | Mostrar como **Código de Cliente** |
| RazonSocial | Una línea de texto | |
| Telefono | Una línea de texto | Con o sin 51; el motor normaliza (+51) |
| Consentimiento | Opción: `Sí`, `No` | Opcional |
| Ciudad | Una línea de texto | Opcional, se usa como `{{ciudad}}` |

**Carga inicial:** en SharePoint → *Nuevo* → *Lista* → **Desde Excel**, elige tu base con las columnas `Código de Cliente`, `Razón Social`, `Número de Contacto`. O usa la plantilla `CRM_Clientes.csv` de esta carpeta: llénala en Excel, dale formato de tabla y créala desde Excel.

## 2. CRM_Lotes — un registro por lote
| Columna | Tipo |
|---|---|
| Title | (existente) — "Lote #N - dd/mm/aaaa" |
| Numero | Número (sin decimales) |
| Fecha | Fecha |
| Ejecutivo | Una línea de texto |
| Contactos | Número |
| Cierre | Una línea de texto |
| Resumen | Varias líneas de texto (texto sin formato) |

## 3. CRM_Gestiones — un registro por contacto de cada lote
| Columna | Tipo | Valores |
|---|---|---|
| Title | (existente) | Mostrar como **Código de Cliente** (uso interno) |
| Lote | Una línea de texto | |
| Ejecutivo | Una línea de texto | |
| CorreoEjecutivo | Una línea de texto | Correo de Teams del ejecutivo |
| RazonSocial | Una línea de texto | |
| Telefono | Una línea de texto | |
| Mensaje | Varias líneas de texto (**texto sin formato**) | Mensaje único generado |
| PesoBytes | Número | |
| Caracteres | Número | |
| HoraProgramada | Fecha y hora (incluir hora) | |
| EstadoEnvio | Opción | `Pendiente`, `Reprogramado`, `Notificado`, `Enviado`, `Omitido`, `No contactar` |
| FechaEnvio | Fecha y hora | |
| EstadoContacto | Opción | `Cliente potencial – carterizar`, `Cotizado – seguimiento`, `Interesado – sin cotización`, `No responde – 1.er intento`, `Primera compra efectuada` |
| ProximaAccion | Opción | `Activar recurrencia`, `Carterizar`, `Cerrar gestión`, `Contactar`, `Cotizar`, `Recontactar`, `Seguimiento cierre` |
| Observacion | Varias líneas de texto (texto sin formato) | |
| NombreContacto | Una línea de texto | Dato que pide el mensaje |
| Correo | Una línea de texto | Dato que pide el mensaje |
| TelefonoDirecto | Una línea de texto | Dato que pide el mensaje |
| EnlaceWhatsApp | Varias líneas de texto (texto sin formato) | El enlace puede superar 255 caracteres |
| FechaGestion | Fecha y hora | Última vez que alguien la actualizó |

> Copia los valores de las opciones **exactamente** como están arriba (con la raya "–"). Son los mismos que usa el Excel BD-BBPE.

## 4. CRM_NoContactar — lista de exclusión (BAJA)
| Columna | Tipo |
|---|---|
| Title | (existente) — **Teléfono** normalizado (ej. 51987654321) |
| Codigo | Una línea de texto |
| Motivo | Una línea de texto |
| Fecha | Fecha y hora |

## 5. CRM_RegistroRelleno — historial para "deshacer" el Excel
| Columna | Tipo |
|---|---|
| Title | (existente) — "Relleno dd/mm/aaaa hh:mm" |
| Fecha | Fecha y hora |
| Resultado | Varias líneas de texto (texto sin formato) |
| CambiosJson | Varias líneas de texto (texto sin formato) |
| Deshecho | Sí/No (predeterminado: No) |

## Índices recomendados
En **CRM_Gestiones** → Configuración de la lista → *Columnas indizadas*: agrega `EstadoEnvio`, `HoraProgramada`, `FechaEnvio` y `Title`. Con eso los filtros siguen funcionando aunque la lista pase de 5 000 elementos.
