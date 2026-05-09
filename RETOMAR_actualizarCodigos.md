# Retomar — Feature `actualizarCodigos.js`

## Objetivo
Tomar `clientes-zona.xlsx` y actualizarlo con los datos de `clientes-activos.xls`.
Generar `clientes-zona-actualizado.xls` con:
- Códigos, teléfonos y zonas actualizados donde haya match
- Filas ambiguas (doble/triple match sin desempate) en **rojo** `#FF0000`
- Clientes nuevos (solo en activos, no en zona) agregados al final en **celeste** `#ADD8E6`

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

## Estado actual del script

```javascript
const fs   = require('fs');
const xlsx = require('xlsx');

// --- UTILIDADES ---

function parseTable(html) {
    const rows = [];
    const trs = html.split(/<tr[^>]*>/i).slice(1);
    for (let tr of trs) {
        tr = tr.split(/<\/tr>/i)[0];
        const tds = tr.split(/<(?:td|th)[^>]*>/i).slice(1);
        const cols = [];
        for (let td of tds) {
            let val = td.split(/<\/(?:td|th)>/i)[0];
            val = val
                .replace(/&nbsp;/g, ' ')
                .replace(/&oacute;/g, 'ó').replace(/&iacute;/g, 'í')
                .replace(/&aacute;/g, 'á').replace(/&eacute;/g, 'é')
                .replace(/&uacute;/g, 'ú').replace(/&ntilde;/g, 'ñ')
                .replace(/&Ntilde;/g, 'Ñ').replace(/&#39;/g, "'")
                .replace(/&amp;/g, '&').replace(/<[^>]+>/g, '').trim();
            cols.push(val);
        }
        if (cols.length > 0) rows.push(cols);
    }
    return rows;
}

function normalizeStr(str) {
    if (!str) return '';
    return str
        .toString()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim()
        .replace(/\s+/g, ' ');
}

// --- PARTE 1: LECTURA DE ARCHIVOS ---

const activosHtml   = fs.readFileSync('clientes-activos_archivos/sheet001.htm', 'latin1');
const activosRows   = parseTable(activosHtml);
const activosHeader = activosRows[0];
const activosData   = activosRows.slice(1);

const wbZona     = xlsx.readFile('clientes-zona.xlsx');
const wsZona     = wbZona.Sheets[wbZona.SheetNames[0]];
const zonaRows   = xlsx.utils.sheet_to_json(wsZona, { header: 1 });
const zonaHeader = zonaRows[0];
const zonaData   = zonaRows.slice(1);

// --- PARTE 2: CONSTRUIR MAP DE LOOKUP ---

const activosMap = new Map();

activosData.forEach(row => {
    const razonSocial = normalizeStr(row[1]);
    if (!razonSocial) return;

    const candidato = {
        codigo:     row[0],
        telefono:   row[2],
        domicilio:  normalizeStr(row[4]),
        codigoZona: row[7],
        zona:       row[8],
    };

    if (!activosMap.has(razonSocial)) {
        activosMap.set(razonSocial, []);
    }

    activosMap.get(razonSocial).push(candidato);
});
```

---

## Partes pendientes

### Parte 3 — Iterar zona y resolver cada fila
```
Para cada fila de clientes-zona:
  │
  ├─ Normalizar RAZON_SOCIAL
  ├─ Buscar en activosMap
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
        └─ empate / ninguno / 2+ superan 85% → NO actualizar, marcar ROJO
```

### Parte 4 — Agregar clientes nuevos desde activos
- Construir `Set` con razones sociales normalizadas de `clientes-zona`
- Iterar `clientes-activos`
- Los que no estén en el Set → agregar al output con fondo **celeste**

### Parte 5 — Generar el output
- Formato HTML-as-XLS
- Header: columnas de `clientes-zona`
- Escribir `clientes-zona-actualizado.xls` encoding `latin1`
- Loggear estadísticas en consola

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
1. Abrí `actualizarCodigos.js` y verificá que tenga el código del estado actual de arriba
2. Pedile al asistente: _"continuemos con la parte 3 de actualizarCodigos"_
