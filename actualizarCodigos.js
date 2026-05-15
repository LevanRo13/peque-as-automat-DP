const fs = require('fs');
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
				.replace(/&amp;/g, '&')
				.replace(/<[^>]+>/g, '').trim();
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
		.replace(/\r\n\s*/g, ' ')
		.toLowerCase()
		.normalize("NFD")
		.replace(/[\u0300-\u036f]/g, "")
		.trim()
		.replace(/\s+/g, ' ');
}

function diceCoefficient(a, b) {
	if (!a || !b) return 0;
	if (a === b) return 1;

	const bigrams = str => {
		const result = new Map();
		for (let i = 0; i < str.length - 1; i++) {
			const bg = str[i] + str[i + 1];
			result.set(bg, (result.get(bg) || 0) + 1);
		}
		return result;
	};

	const bigramsA = bigrams(a);
	const bigramsB = bigrams(b);

	let intersection = 0;
	bigramsA.forEach((count, bg) => {
		if (bigramsB.has(bg)) {
			intersection += Math.min(count, bigramsB.get(bg));
		}
	});

	let sizeA = 0;
	let sizeB = 0;
	bigramsA.forEach(count => sizeA += count);
	bigramsB.forEach(count => sizeB += count);
	return (2 * intersection) / (sizeA + sizeB);
}

const FUZZY_THRESHOLD = 0.85;

// --- PARTE 1: LECTURA DE ARCHIVOS ---

const wbActivos = xlsx.readFile('clientes-activos.xls');
const wsActivos = wbActivos.Sheets[wbActivos.SheetNames[0]];
const activosRows = xlsx.utils.sheet_to_json(wsActivos, { header: 1 });
const activosData = activosRows.slice(1);

const wbZona = xlsx.readFile('nuevosClientes-Zona.xlsx');
const wsZona = wbZona.Sheets[wbZona.SheetNames[0]];
const zonaRows = xlsx.utils.sheet_to_json(wsZona, { header: 1 });
const zonaHeader = zonaRows[0];
const zonaData = zonaRows.slice(1);

// --- PARTE 2: CONSTRUIR MAP DE LOOKUP DE ACTIVOS ---
// clave: razonSocial normalizada
// valor: array de candidatos { raw, domicilio }

const activosMap = new Map();

activosData.forEach(row => {
	const razonSocial = normalizeStr(row[3]); // índice 3 = Razón Social
	if (!razonSocial) return;

	const candidato = {
		codigo: row[1],                        // índice 1 = Código (columna B)
		domicilio: normalizeStr(row[9]),       // índice 9 = Domicilio
	};

	if (!activosMap.has(razonSocial)) {
		activosMap.set(razonSocial, []);
	}
	activosMap.get(razonSocial).push(candidato);
});

// --- PARTE 3: COMPARACIÓN Y CONSTRUCCIÓN DEL OUTPUT ---

const outputHeader = ['WKT', 'nombre', 'CODIGO', 'RAZON_SOCIAL', 'Teléfono', 'Email', 'Condición IVA', 'Domicilio', 'Departamento', 'Provincia', 'Zona'];

const outputRows = [];
let countExacto = 0;
let countFuzzy = 0;
let countRojo = 0;
let countSinMatch = 0;

zonaData.forEach(zonaRow => {
	const razonSocial = normalizeStr(zonaRow[3]); // índice 3 = RAZON_SOCIAL
	const candidatos = activosMap.get(razonSocial);

	// Sin match — columnas A y B vacías, color rojo
	if (!candidatos || candidatos.length === 0) {
		countSinMatch++;
		const newRow = Array.from({ length: outputHeader.length }, (_, i) => i < 2 ? '' : (zonaRow[i] ?? ''));
		outputRows.push({ row: newRow, color: '#FF0000' });
		return;
	}

	// Resolver el match ganador
	let ganador = null;

	if (candidatos.length === 1) {
		ganador = candidatos[0];
		countExacto++;
	} else {
		// Múltiples — desempate por fuzzy en domicilio
		const domicilioZona = normalizeStr(zonaRow[7]); // índice 7 = Domicilio
		const superanUmbral = candidatos.filter(c =>
			diceCoefficient(domicilioZona, c.domicilio) >= FUZZY_THRESHOLD
		);

		if (superanUmbral.length === 1) {
			ganador = superanUmbral[0];
			countFuzzy++;
		} else {
			// Empate o ninguno supera — rojo
			countRojo++;
			outputRows.push({ row: Array.from({ length: outputHeader.length }, (_, i) => zonaRow[i] ?? ''), color: '#FF0000' });
			return;
		}
	}

	// Sobreescribir CODIGO (índice 2) con el código nuevo de activos
	const newRow = Array.from({ length: outputHeader.length }, (_, i) => zonaRow[i] ?? '');
	newRow[2] = ganador.codigo;
	outputRows.push({ row: newRow, color: null });
});

// --- PARTE 4: GENERAR EL EXCEL DE SALIDA ---

let outHtml = `<html><style>body {margin: 0; padding: 0;} td { mso-number-format:'@';}</style><body style="font-family:SansSerif;">\n`;
outHtml += `<table style="padding: 0; font-size:8pt;" border="1">\n`;

outHtml += `<tr>` + outputHeader.map(h => `<th style="background-color:#CCCCCC;font-weight:bold" align="center">${h}</th>`).join('') + `</tr>\n`;

outputRows.forEach(({ row, color }) => {
	const trStyle = color ? ` style="background-color:${color}"` : '';
	outHtml += `<tr${trStyle}>` + row.map(c => `<td>${c ?? ''}</td>`).join('') + `</tr>\n`;
});

outHtml += `</table></body></html>`;

fs.writeFileSync('clientes-zona-actualizado.xls', outHtml, 'latin1');

console.log(`Filas en nuevosClientes-Zona:   ${zonaData.length}`);
console.log(`Match exacto:                   ${countExacto}`);
console.log(`Match por fuzzy:                ${countFuzzy}`);
console.log(`Sin match (rojo):               ${countSinMatch}`);
console.log(`Ambiguos (rojo):                ${countRojo}`);
console.log(`Archivo generado: clientes-zona-actualizado.xls`);
