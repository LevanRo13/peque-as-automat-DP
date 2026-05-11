const xlsx = require('xlsx');
const fs = require('fs');

// --- PARTE 1: LECTURA DE ARCHIVOS ---

const wbActivos = xlsx.readFile('clientes-activos.xls');
const wsActivos = wbActivos.Sheets[wbActivos.SheetNames[0]];
const activosRows = xlsx.utils.sheet_to_json(wsActivos, { header: 1 });

const activosHeader = activosRows[0];
const activosData = activosRows.slice(1);

const wbZona = xlsx.readFile('clientes-zona.xlsx');
const wsZona = wbZona.Sheets[wbZona.SheetNames[0]];
const zonaRows = xlsx.utils.sheet_to_json(wsZona, { header: 1 });

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

// --- PARTE 3: COMPARACIÓN Y CLASIFICACIÓN ---

const outputRows = [];

activosData.forEach(row => {
	const razonSocial = normalizeStr(row[3]);
	const encontrado = zonaRazonesSociales.has(razonSocial);

	if (!encontrado) {
		const newRow = row.slice();
		outputRows.push(newRow);
	}
});

// --- PARTE 4: GENERAR EL EXCEL DE SALIDA ---

const outputHeader = activosHeader.slice();

const totalActivos = outputRows.length;
const encontrados = outputRows.filter(row => row[row.length - 1] === "").length;
const nuevos = outputRows.filter(row => row[row.length - 1] === "Nuevo").length;

let outHtml = `<html><style>body {margin: 0; padding: 0;} td { mso-number-format:'@';}</style><body style="font-family:SansSerif;">\n`;
outHtml += `<table style="padding: 0; font-size:8pt;" border="1">\n`;

outHtml += `<tr>` + outputHeader.map(h => `<th style="background-color:#CCCCCC;font-weight:bold" align="center">${h ?? ''}</th>`).join('') + `</tr>\n`;

outputRows.forEach(row => {
	outHtml += `<tr>` + row.map(c => `<td>${c ?? ''}</td>`).join('') + `</tr>\n`;
});

outHtml += `</table></body></html>`;

fs.writeFileSync('ClientesParaPoint.xls', outHtml, 'latin1');

console.log(`Clientes Activos total:               ${activosData.length}`);
console.log(`Encontrados en Zona Vendedores:       ${activosData.length - outputRows.length}`);
console.log(`Nuevos (no están en Zona Vendedores): ${outputRows.length}`);
console.log(`Archivo generado: ClientesParaPoint.xls`);