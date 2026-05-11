# Retomar — `actualizarCodigos.js`

## ⚠️ ANTES DE ARRANCAR — Bug conocido
La línea 21-22 de `actualizarCodigos.js` tiene un error de corrupción al pegar. Buscá esto:

```js
.replace(/&amp;/g, '&').repla¿
} ce(/<[^>]+>/g, '').trim();
```

Y reemplazalo por esto:

```js
.replace(/&amp;/g, '&')
.replace(/<[^>]+>/g, '').trim();
```

**Sin esto el script no va a correr.**

---

## Objetivo
Tomar `clientes-zona.xlsx` y actualizarlo con los datos de `clientes-activos.xls`.
Generar `clientes-zona-actualizado.xls` con:
- Códigos, teléfonos y zonas actualizados donde haya match
- Filas ambiguas (doble/triple match sin desempate) en **rojo** `#FF0000`
- Clientes nuevos (solo en activos, no en zona) agregados al final en **celeste** `#ADD8E6`

---

## Setup en PC nueva
```bash
git clone https://github.com/LevanRo13/peque-as-automat-DP.git
cd peque-as-automat-DP
npm install
```

Además necesitás copiar manualmente a la carpeta:
- `clientes-activos.xls` + su carpeta `clientes-activos_archivos/` (no están en el repo por `.gitignore`)
- `clientes-zona.xlsx`

---

## Archivos del proyecto
| Archivo | Rol |
|---|---|
| `clientes-activos_archivos/sheet001.htm` | Fuente de verdad (HTML multi-archivo del ERP) |
| `clientes-zona.xlsx` | Archivo a actualizar |
| `actualizarCodigos.js` | Script en construcción |
| `PLAN_actualizarCodigos.md` | Plan detallado completo |

---

## Estructura de columnas

### `clientes-activos` (sheet001.htm)
| Índice | Columna |
|---|---|
| 0 | Código ← actualizar en zona |
| 1 | Razón Social ← comparar |
| 2 | Teléfono ← actualizar en zona |
| 4 | Domicilio ← desempate fuzzy |
| 7 | Código Zona ← actualizar en zona |
| 8 | Zona ← actualizar en zona |

### `clientes-zona` (xlsx, header en row 0, datos desde row 1)
| Índice | Columna |
|---|---|
| 0 | CODIGO ← reemplazar |
| 1 | RAZON_SOCIAL ← comparar |
| 2 | Teléfono ← reemplazar |
| 4 | Domicilio ← desempate fuzzy |
| 7 | Código Zona ← reemplazar |
| 8 | Zona ← reemplazar |

---

## Estado de las partes

- ✅ Parte 1 — Lectura de archivos
- ✅ Parte 2 — Construcción del Map de lookup
- ✅ Parte 3 — Comparación, fuzzy y clasificación por color
- 🔲 Parte 4 — Agregar clientes nuevos desde activos (celeste)
- 🔲 Parte 5 — Generar el Excel de salida

---

## Partes pendientes

### Parte 4 — Agregar clientes nuevos desde activos
- Construir un `Set` con las razones sociales normalizadas de **todas** las filas de `clientes-zona`
- Iterar `clientes-activos`
- Si la razón social normalizada **no está** en el Set → es un cliente nuevo
- Agregarlo a `outputRows` con `color: '#ADD8E6'`

### Parte 5 — Generar el output
- Formato HTML-as-XLS (mismo approach que `compare.js`)
- Header: columnas de `clientes-zona`
- Filas normales → sin color
- Filas ambiguas → `style="background-color:#FF0000"`
- Filas nuevas → `style="background-color:#ADD8E6"`
- Escribir `clientes-zona-actualizado.xls` encoding `latin1`
- Loggear en consola:
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

## Colores
| Situación | Color | Hex |
|---|---|---|
| Match normal | Sin color | — |
| Ambigua | Rojo | `#FF0000` |
| Cliente nuevo | Celeste | `#ADD8E6` |

---

## Función fuzzy — Dice Coefficient
```
similaridad = (2 * intersección de bigramas) / (total bigramas A + total bigramas B)
```
Umbral: **0.85**

---

## Cómo retomar
1. Corregí el bug de la línea 21-22 (ver arriba ⚠️)
2. Pedile al asistente: _"continuemos con la parte 4 de actualizarCodigos"_
