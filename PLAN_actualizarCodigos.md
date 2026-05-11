# Plan — `actualizarCodigos.js`

## Objetivo
Tomar `clientes-zona.xlsx` y actualizarlo con los datos de `clientes-activos.xls`.
El resultado es un Excel completo con:
- Códigos, teléfonos y zonas actualizados donde haya match
- Filas ambiguas marcadas en **rojo**
- Clientes nuevos (solo en activos) agregados al final en **celeste**

---

## Inputs

| Archivo | Rol |
|---|---|
| `clientes-activos_archivos/sheet001.htm` | Fuente de verdad |
| `clientes-zona.xlsx` | Archivo a actualizar |

## Output

`clientes-zona-actualizado.xls`

---

## Estructura de columnas

### `clientes-activos` (sheet001.htm, parseado como HTML)
| Índice | Columna |
|---|---|
| 0 | Código ← actualizar en zona |
| 1 | Razón Social ← comparar |
| 2 | Teléfono ← actualizar en zona |
| 4 | Domicilio ← desempate fuzzy |
| 7 | Código Zona ← actualizar en zona |
| 8 | Zona ← actualizar en zona |

### `clientes-zona` (xlsx, row 0 = header, datos desde row 1)
| Índice | Columna |
|---|---|
| 0 | CODIGO ← reemplazar |
| 1 | RAZON_SOCIAL ← comparar |
| 2 | Teléfono ← reemplazar |
| 4 | Domicilio ← desempate fuzzy |
| 7 | Código Zona ← reemplazar |
| 8 | Zona ← reemplazar |

---

## Algoritmo — 5 partes

---

### Parte 1 — Leer ambos archivos

- `clientes-activos`: leer `clientes-activos_archivos/sheet001.htm` con `fs.readFileSync` encoding `latin1`, parsear con la función `parseTable` existente (la misma de `compare.js`)
- `clientes-zona`: leer con `xlsx.readFile`, convertir a array de arrays con `sheet_to_json({ header: 1 })`, separar header (row 0) de datos (row 1 en adelante)

---

### Parte 2 — Construir el Map de lookup de activos

Estructura:
```
Map<razonSocialNormalizada, Array<{ codigo, telefono, domicilio, codigoZona, zona }>>
```

- Iterar cada fila de `clientes-activos`
- Normalizar `Razón Social` (lowercase, sin tildes, sin espacios dobles)
- Si la clave no existe en el Map → crearla con un array vacío
- Pushear el objeto con los datos relevantes

**¿Por qué un array por clave?**
Porque puede haber múltiples clientes con la misma razón social. Eso es exactamente lo que dispara la comparación fuzzy de domicilio.

---

### Parte 3 — Iterar zona y resolver cada fila

Para cada fila de `clientes-zona`:

```
Normalizar RAZON_SOCIAL
Buscar en el Map
  │
  ├─ 0 matches → dejar fila sin tocar, color normal
  │
  ├─ 1 match exacto → actualizar:
  │     row[0] = codigo
  │     row[2] = telefono
  │     row[7] = codigoZona
  │     row[8] = zona
  │     color normal
  │
  └─ 2+ matches → comparar Domicilio con fuzzy (Dice coefficient, umbral 85%)
        │
        ├─ 1 candidato supera 85% → actualizar, color normal
        └─ empate / ninguno supera 85% / 2+ superan 85% → NO actualizar, marcar en ROJO
```

---

### Parte 4 — Agregar clientes nuevos desde activos

- Construir un `Set` con las razones sociales normalizadas de **todas** las filas de `clientes-zona`
- Iterar `clientes-activos`
- Si la razón social normalizada **no está** en el Set → es un cliente nuevo
- Agregarlo al output con fondo **celeste**

---

### Parte 5 — Generar el output

- Formato HTML-as-XLS (mismo approach que `compare.js`)
- Header: columnas de `clientes-zona`
- Filas normales → sin color
- Filas ambiguas → `style="background-color:#FF0000"`
- Filas nuevas → `style="background-color:#ADD8E6"`
- Escribir `clientes-zona-actualizado.xls` con `fs.writeFileSync` encoding `latin1`

**Estadísticas en consola:**
```
Filas en clientes-zona:        X
Actualizadas con match exacto: X
Actualizadas con fuzzy:        X
Ambiguas (marcadas en rojo):   X
Sin match:                     X
Clientes nuevos agregados:     X
Archivo generado: clientes-zona-actualizado.xls
```

---

## Función fuzzy — Dice Coefficient

Sin librerías externas. Compara bigramas (pares de caracteres consecutivos) entre dos strings.

```
similaridad = (2 * intersección de bigramas) / (total bigramas A + total bigramas B)
```

Umbral: **0.85 (85%)**

---

## Colores de referencia

| Situación | Color | Hex |
|---|---|---|
| Match normal | Sin color | — |
| Ambigua (rojo) | Rojo | `#FF0000` |
| Cliente nuevo (celeste) | Celeste | `#ADD8E6` |

---

## Cómo arrancar
Crear el archivo `actualizarCodigos.js` en la carpeta del proyecto y construirlo parte por parte siguiendo este plan. Pedirle al asistente: _"arrancamos con la parte 1"_.
