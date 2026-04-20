const fs = require('fs');

function parseTable(html) {
    const rows = [];
    const trs = html.split(/<tr[^>]*>/i).slice(1);
    for (let tr of trs) {
        tr = tr.split(/<\/tr>/i)[0];
        const tds = tr.split(/<(?:td|th)[^>]*>/i).slice(1);
        const cols = [];
        for (let td of tds) {
            let val = td.split(/<\/(?:td|th)>/i)[0];
            val = val.replace(/&nbsp;/g, ' ')
                     .replace(/&oacute;/g, 'ó')
                     .replace(/&iacute;/g, 'í')
                     .replace(/&aacute;/g, 'á')
                     .replace(/&eacute;/g, 'é')
                     .replace(/&uacute;/g, 'ú')
                     .replace(/&Oacute;/g, 'Ó')
                     .replace(/&Iacute;/g, 'Í')
                     .replace(/&Aacute;/g, 'Á')
                     .replace(/&Eacute;/g, 'É')
                     .replace(/&Uacute;/g, 'Ú')
                     .replace(/&ntilde;/g, 'ñ')
                     .replace(/&Ntilde;/g, 'Ñ')
                     .replace(/&#39;/g, "'")
                     .replace(/&amp;/g, '&')
                     .replace(/<[^>]+>/g, '') // remove inner tags if any
                     .trim();
            cols.push(val);
        }
        if (cols.length > 0) rows.push(cols);
    }
    return rows;
}

function normalizeStr(str) {
    if (!str) return '';
    return str.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().replace(/\s+/g, ' ');
}

try {
    const activosHtml = fs.readFileSync('Reporte_clientes_activos.xls', 'latin1');
    const noCompranHtml = fs.readFileSync('reporte_facturacion_clientes_no_compran_2026_04_20_12_31_49.xls', 'latin1');

    const activosRows = parseTable(activosHtml);
    const noCompranRows = parseTable(noCompranHtml);

    const activosData = activosRows.slice(1);
    const noCompranData = noCompranRows.slice(1);

    const activosCodigos = new Map();
    const activosNombres = new Map();
    
    activosData.forEach(row => {
        const codigo = row[1];
        const razonSocial = normalizeStr(row[3]);
        const nombreFantasia = normalizeStr(row[4]);
        const vendedor = row[17]; // Columna 17 es "Vendedor"
        
        if (codigo) activosCodigos.set(codigo, vendedor);
        if (razonSocial) activosNombres.set(razonSocial, vendedor);
        if (nombreFantasia) activosNombres.set(nombreFantasia, vendedor);
    });

    const outputRows = [];
    const header = noCompranRows[0].slice(); // Copia del arreglo
    header.push("Vendedor"); // Agregamos título de la columna nueva
    
    noCompranData.forEach(row => {
        const codigo = row[0];
        const cliente = normalizeStr(row[2]);
        
        let found = false;
        let vendedorAsociado = "";

        if (codigo && activosCodigos.has(codigo)) {
            found = true;
            vendedorAsociado = activosCodigos.get(codigo);
        } else if (cliente && activosNombres.has(cliente)) {
            found = true;
            vendedorAsociado = activosNombres.get(cliente);
        }
        
        if (found) {
            const newRow = row.slice();
            newRow.push(vendedorAsociado || "");
            outputRows.push(newRow);
        }
    });

    console.log(`Clientes que No Compran original: ${noCompranData.length}`);
    console.log(`Clientes borrados (No Compran pero NO son Activos / ya inactivos totales): ${noCompranData.length - outputRows.length}`);
    console.log(`Clientes RESULTANTES (Activos que No Compran con Vendedor asignado): ${outputRows.length}`);
    
    let outHtml = `<html><style> body {margin: 0; padding: 0;} td { mso-number-format:'@';} .resumen{background-color:#CCCCCC;font-weight:bold} .nd {mso-number-format:'#,##0.00';text-align:right;} .ne {mso-number-format:'#,##0';text-align:right;} .n {mso-number-format:'#';text-align:right;} .sp {mso-number-format:'#';}</style><body style="font-family:SansSerif;">\n`;
    outHtml += `<table style="padding: 0; font-size:8pt;" border="1">\n`;
    outHtml += `<tr>` + header.map(h => `<th style="background-color:#CCCCCC;font-weight:bold" align="center">${h}</th>`).join('') + `</tr>\n`;
    outputRows.forEach(row => {
        outHtml += `<tr>` + row.map(c => `<td>${c}</td>`).join('') + `</tr>\n`;
    });
    outHtml += `</table></body></html>`;
    
    fs.writeFileSync('Clientes_No_Compran_Con_Vendedor.xls', outHtml, 'latin1');
    console.log(`Archivo generado exitosamente: Clientes_No_Compran_Con_Vendedor.xls`);

} catch (err) {
    console.error(err);
}
