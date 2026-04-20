# Guía de Uso - Limpieza de Clientes No Compradores

Este directorio contiene un pequeño automatismo para limpiar los reportes de clientes generados por el ERP de la empresa.

## 🎯 **¿Cuál es el problema?**
El ERP, al generar el listado de *"Clientes que no compran"*, arroja un documento que incluye a **TODO el histórico de clientes inactivos**, incluso aquellos dados de baja y que ya ni figuran en el padrón actual de "Clientes Activos". 

## 🛠️ **¿Qué hace el script?**
El archivo `compare.js` realiza una **intersección** (cruce dinámico) entre dos reportes:
1. `Reporte_clientes_activos.xls`
2. `reporte_facturacion_clientes_no_compran_...xls`

El programa procesa las tablas de ambos archivos, toma cada "cliente que no compra" y verifica si el mismo **EXISTE** actualmente en la lista de activos. Si existe, lo preserva en un nuevo archivo; si no existe, lo elimina definitivamente del reporte final. 
Como "bonus track", el comportamiento capta el nombre del **Vendedor** de la lista de activos y lo anexa a la nueva planilla para facilitar el seguimiento de las reactivaciones.

---

## ⚡ **Guía paso a paso para ejecutarlo**

### 1. Preparar los reportes
Asegúrate de que dentro de esta misma carpeta ("Limpieza no compradores") estén presentes ambos archivos extraídos del ERP, con sus nombres originales:
- `Reporte_clientes_activos.xls`
- `reporte_facturacion_clientes_no_compran_2026_04_20_12_31_49.xls` *(Si extraes un reporte nuevo, ten en cuenta que el nombre del archivo cambia por la fecha. Deberás o bien renombrar tu archivo nuevo con este nombre, o actualizar la línea 30 del archivo `compare.js` con el nombre correcto de tu nuevo archivo).*

### 2. Ejecución
El script funciona sobre **Node.js**. 
Simplemente abre tu terminal o consola (PowerShell o CMD) apuntando a esta carpeta y corre:
```bash
node compare.js
```

### 3. Resultados
Una vez ejecutado, deberías ver estadísticas útiles en la consola:
```text
Clientes que No Compran original: [Cantidad inicial ERP]
Clientes borrados (No Compran pero NO son Activos): [Purgados]
Clientes RESULTANTES: [Cantidad final]
```

Finalizado el proceso, se generará **automáticamente** un nuevo archivo limpio en la misma carpeta:
👉 **`Clientes_No_Compran_Con_Vendedor.xls`**

Esta planilla de salida es 100% compatible con tus herramientas ofimáticas y conserva toda la estructura y formato de la tabla original generada por el software.
