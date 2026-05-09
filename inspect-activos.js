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

const content = fs.readFileSync('clientes-activos_archivos/sheet001.htm', 'latin1');
const rows = parseTable(content);
console.log('Headers:', JSON.stringify(rows[0]));
console.log('Row 1:', JSON.stringify(rows[1]));
console.log('Row 2:', JSON.stringify(rows[2]));
console.log('Total rows:', rows.length);
