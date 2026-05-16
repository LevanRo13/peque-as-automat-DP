# Relleno automático de WKT — Guía de uso

## ¿Qué hace esto?

Toma las coordenadas (WKT) y nombres de `Points lista.xlsx` y las copia automáticamente
a `clientes-zona-actualizado.xlsx` para los clientes que no tienen ubicación cargada.

Usa **fuzzy matching** por nombre para encontrar correspondencias aproximadas,
ya que el mismo cliente puede estar escrito distinto en cada archivo
(ej: `"ADRIANA PIERRO"` en clientes vs `"Adriana piero"` en Points).

---

## Archivos

| Archivo | Rol |
|---------|-----|
| `Points lista.xlsx` | **Fuente** — contiene WKT (col A) y nombre (col B) |
| `clientes-zona-actualizado.xlsx` | **Destino** — se lee pero NO se modifica |
| `clientes-zona-resultado.xlsx` | **Output** — copia del destino con WKT rellenados y colores |
| `fill_wkt.py` | Script principal |

---

## Requisitos

Python 3.x con estas librerías:

```bash
pip install openpyxl rapidfuzz
```

---

## Cómo ejecutar

```bash
python fill_wkt.py
```

El script genera `clientes-zona-resultado.xlsx` en la misma carpeta.
**No modifica** los archivos originales.

---

## Cómo interpretar el resultado

El output tiene colores en las columnas **A (WKT)** y **B (nombre)**,
y una columna extra **L (CANDIDATOS_FUZZY)** con notas del match.

| Color | Criterio | Qué hacer |
|-------|----------|-----------|
| **Verde** | Match ≥ 90% — muy confiable | Nada, está listo |
| **Amarillo** | Match 70–89%, un solo candidato | Revisar rápido que el nombre tenga sentido |
| **Naranja** | Match 70–89%, varios candidatos | Ver col L, elegir el candidato correcto y borrar los otros |
| **Rojo** | Sin match < 70% | Cargar WKT a mano o dejar vacío |
| *(sin color)* | Ya tenía WKT cargado | No se tocó |

> **Tip para los naranjas:** la col L muestra todos los candidatos con su score.
> Ejemplo: `CONFLICTO | Adriana piero (85%) | adriana iturriel (73%)`
> Elegís el correcto, actualizás col A y B, y borrás la nota de col L.

---

## Umbrales configurables

Al principio de `fill_wkt.py` podés ajustar estos valores:

```python
THRESHOLD_AUTO  = 90   # >= este score -> auto-fill verde
THRESHOLD_FUZZY = 70   # >= este score -> revisión amarillo/naranja
TOP_N           = 3    # cuántos candidatos mostrar en col L para los conflictos
```

Si los resultados tienen muchos falsos positivos, subí `THRESHOLD_AUTO` a 92–95.
Si hay pocos matches y querés más candidatos para revisar, bajá `THRESHOLD_FUZZY` a 65.

---

## Limitaciones conocidas

- **Domicilios tipo "B° 1 de mayo Casa 2"** no se pueden geocodificar, por eso el approach es por nombre.
- Clientes con nombre muy genérico (`"CINTIA"`, `"DANIEL"`, `"ALE"`) raramente matchean bien — quedan en rojo.
- Si `Points lista.xlsx` no tiene el punto de un cliente, ese cliente siempre va a quedar en rojo sin importar el umbral.
- El script usa `token_sort_ratio` de rapidfuzz, que es tolerante al orden de las palabras (`"Carmen Aguirre"` == `"Aguirre Carmen"`).

---

## Si necesitás actualizar Points lista en el futuro

1. Reemplazá `Points lista.xlsx` con la versión nueva (misma estructura: col A = WKT, col B = nombre).
2. Actualizá `clientes-zona-actualizado.xlsx` con el padrón vigente.
3. Corré `python fill_wkt.py` de nuevo.
4. El output se sobreescribe en `clientes-zona-resultado.xlsx`.
