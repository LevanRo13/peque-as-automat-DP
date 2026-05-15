# Retomar — `actualizarCodigos.js`

## Estado actual
✅ Script completo y funcional. No hay partes pendientes de implementación.

---

## Setup en PC nueva
```bash
git clone https://github.com/LevanRo13/peque-as-automat-DP.git
cd peque-as-automat-DP
npm install
```

Copiar manualmente a la carpeta (no están en el repo por `.gitignore`):
- `clientes-activos.xls` + carpeta `clientes-activos_archivos/`
- `nuevosClientes-Zona.xlsx`

Ejecutar:
```bash
node actualizarCodigos.js
```

---

## Objetivo
Cruzar `nuevosClientes-Zona.xlsx` contra `clientes-activos` para actualizar códigos y datos.
Genera `clientes-zona-actualizado.xls` con:
- Filas con match → estructura combinada zona + activos
- Filas sin match o ambiguas → fila intacta de zona en **rojo** `#FF0000`

---

## Inputs
| Archivo | Rol |
|---|---|
| `clientes-activos_archivos/sheet001.htm` | Fuente de verdad — HTML multi-archivo del ERP |
| `nuevosClientes-Zona.xlsx` | Archivo base a actualizar |

## Output
`clientes-zona-actualizado.xls`

---

## Estructura de columnas

### `clientes-activos` (sheet001.htm, parseado como HTML)
| Índice | Columna |
|---|---|
| 0 | Código |
| 1 | Razón Social ← comparar |
| 2 | Teléfono |
| 3 | Condición IVA |
| 4 | Domicilio ← desempate fuzzy |
| 5 | Departamento |
| 6 | Provincia |
| 7 | Código Zona |
| 8 | Zona |

### `nuevosClientes-Zona.xlsx` (header row 0, datos desde row 1)
| Índice | Columna |
|---|---|
| 0 | WKT |
| 1 | nombre |
| 2 | CODIGO |
| 3 | RAZON_SOCIAL ← comparar |
| 4 | Teléfono |
| 5 | Email |
| 6 | Condición IVA |
| 7 | Domicilio ← desempate fuzzy |
| 8 | Departamento |
| 9 | Provincia |
| 10 | Zona |

### Output — estructura de fila con match
| Col | Fuente |
|---|---|
| A - WKT | nuevosClientes-Zona (índice 0) |
| B - nombre | nuevosClientes-Zona (índice 1) |
| C - CODIGO | clientes-activos (índice 0) |
| D - RAZON_SOCIAL | clientes-activos (índice 1) |
| E - Teléfono | clientes-activos (índice 2) |
| F - Condición IVA | clientes-activos (índice 3) |
| G - Domicilio | clientes-activos (índice 4) |
| H - Departamento | clientes-activos (índice 5) |
| I - Provincia | clientes-activos (índice 6) |
| J - Código Zona | clientes-activos (índice 7) |
| K - Zona | clientes-activos (índice 8) |

---

## Lógica de matching

```
Para cada fila de nuevosClientes-Zona:
  │
  ├─ Normalizar RAZON_SOCIAL (índice 3)
  ├─ Buscar en activosMap
  │
  ├─ 0 matches → fila intacta de zona, ROJO
  │
  ├─ 1 match → actualizar con datos de activos, sin color
  │
  └─ 2+ matches → fuzzy Dice en Domicilio (umbral 85%)
        ├─ 1 supera umbral → actualizar, sin color
        └─ empate / ninguno / 2+ superan → fila intacta de zona, ROJO
```

---

## Gotcha importante — `\r\n` en clientes-activos
El HTML del ERP tiene saltos de línea dentro de las celdas (`\r\n  `).
`normalizeStr` los limpia con `.replace(/\r\n\s*/g, ' ')` ANTES del toLowerCase.
Sin esto el match falla porque las razones sociales quedan con espacios en el medio.

---

## Fuzzy — Dice Coefficient con Multiset
```
similaridad = (2 * intersección de bigramas) / (total bigramas A + total bigramas B)
```
- Usa `Map` en lugar de `Set` para contar frecuencia de bigramas (multiset real)
- Umbral: **0.85**

---

## Colores
| Situación | Color | Hex |
|---|---|---|
| Match | Sin color | — |
| Sin match / ambiguo | Rojo | `#FF0000` |
