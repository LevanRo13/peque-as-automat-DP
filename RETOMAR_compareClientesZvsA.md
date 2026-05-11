# Retomar — Feature `compareClientesZvsA.js`

## Contexto
Estamos construyendo un nuevo script de forma didáctica, parte por parte.
El objetivo es comparar los clientes de `clientes-activos` contra `clientes-zona` y generar un Excel resultante con un indicador `"Nuevo"` para los que no aparecen en Zona Vendedores.

---

## Nombres de archivos actualizados
| Rol | Nombre de archivo |
|---|---|
| Clientes activos del sistema | `clientes-activos.xls` |
| Clientes cargados en Zona Vendedores | `clientes-zona.xlsx` |
| Archivo de salida | `ClientesParaPoint.xls` |

> ⚠️ El script debe reflejar estos nombres. La Parte 1 todavía tiene los nombres viejos — hay que actualizarlos.

---

## Estado actual del script

```javascript
const xlsx = require('xlsx');

// --- PARTE 1: LECTURA DE ARCHIVOS ---

const wbActivos = xlsx.readFile('clientes-activos.xls');  // nombre actualizado
const wsActivos = wbActivos.Sheets[wbActivos.SheetNames[0]];
const activosRows = xlsx.utils.sheet_to_json(wsActivos, { header: 1 });

const activosHeader = activosRows[0];    // fila 0 = títulos de columna
const activosData   = activosRows.slice(1); // fila 1 en adelante = datos reales

const wbZona = xlsx.readFile('clientes-zona.xlsx');  // nombre actualizado
const wsZona = wbZona.Sheets[wbZona.SheetNames[0]];
const zonaRows = xlsx.utils.sheet_to_json(wsZona, { header: 1 });

// Este archivo tiene 2 filas de header (row 0 y row 1)
// Los datos reales arrancan en row 2
const zonaData = zonaRows.slice(2);

// --- PARTE 2: CONSTRUCCIÓN DEL SET DE LOOKUP ---

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

const zonaRazonesSociales = new Set();

zonaData.forEach(row => {
    const razonSocial = normalizeStr(row[6]);
    if (razonSocial) {
        zonaRazonesSociales.add(razonSocial);
    }
});
```

---

## Partes pendientes

### Parte 3 — Comparación y clasificación
- Iterar cada fila de `activosData`
- Tomar `Razón Social` en índice **3**
- Normalizarla y buscar en `zonaRazonesSociales`
- Si está → columna indicadora vacía
- Si no está → columna indicadora con valor `"Nuevo"`
- Guardar todas las filas en `outputRows`

### Parte 4 — Generar el Excel de salida
- Mismo formato HTML-as-XLS que usa el `compare.js` existente
- Headers de `clientes-activos` + columna nueva `"Indicador"` al final
- Escribir archivo `ClientesParaPoint.xls`
- Loggear en consola:
  ```
  Clientes Activos total: X
  Encontrados en Zona Vendedores: X
  Nuevos (no están en Zona Vendedores): X
  Archivo generado: ClientesParaPoint.xls
  ```

---

## Estructura de columnas clave

### `clientes-activos.xls`
| Índice | Columna |
|---|---|
| 0 | Fecha Alta |
| 3 | **Razón Social** ← usada para comparar |
| 5 | Denominación |
| 9 | Domicilio |
| 17 | Vendedor |

### `clientes-zona.xlsx`
| Índice | Columna |
|---|---|
| 6 | **RAZON_SOCIAL** ← usada para comparar |
| 10 | Domicilio |
- Row 0 y Row 1 son headers — datos reales desde Row 2

---

## Cómo retomar
1. Abrí este proyecto en tu editor
2. Actualizá los nombres de archivo en la Parte 1 del script (ya están corregidos arriba)
3. Pegá el estado actual del script en `compareClientesZvsA.js`
4. Continuá con la **Parte 3** pidiendo al asistente: _"continuemos con la parte 3"_
