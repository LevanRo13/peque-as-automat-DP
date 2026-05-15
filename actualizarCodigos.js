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

const activosHtml = fs.readFileSync('clientes-activos_archivos/sheet001.htm', 'latin1');
const activosRows = parseTable(activosHtml);
const activosData = activosRows.slice(1);

const wbZona     = xlsx.readFile('nuevosClientes-Zona.xlsx');
const wsZona     = wbZona.Sheets[wbZona.SheetNames[0]];
const zonaRows   = xlsx.utils.sheet_to_json(wsZona, { header: 1 });
const zonaHeader = zonaRows[0];
const zonaData   = zonaRows.slice(1);

// --- PARTE 2: CONSTRUIR MAP DE LOOKUP DE ACTIVOS ---
// clave: razonSocial normalizada
// valor: array de candidatos { raw, domicilio }

const activosMap = new Map();

activosData.forEach(row => {
	const razonSocial = normalizeStr(row[1]);
	if (!razonSocial) return;

	const candidato = {
		raw:       row,
		domicilio: normalizeStr(row[4]),
	};

	if (!activosMap.has(razonSocial)) {
		activosMap.set(razonSocial, []);
	}
	activosMap.get(razonSocial).push(candidato);
});

// --- PARTE 3: COMPARACIÓN Y CONSTRUCCIÓN DEL OUTPUT ---

const outputRows = [];
let countExacto  = 0;
let countFuzzy   = 0;
let countRojo    = 0;
let countSinMatch = 0;

zonaData.forEach(zonaRow => {
	const razonSocial = normalizeStr(zonaRow[3]); // índice 3 = RAZON_SOCIAL
	const candidatos  = activosMap.get(razonSocial);

	// Sin match — fila intacta de zona, color rojo
	if (!candidatos || candidatos.length === 0) {
		countSinMatch++;
		outputRows.push({ row: zonaRow.map(c => c ?? ''), color: '#FF0000' });
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
			outputRows.push({ row: zonaRow.map(c => c ?? ''), color: '#FF0000' });
			return;
		}
	}

	// Armar fila de salida: [WKT, nombre] de zona + resto de activos
	const activoRow = ganador.raw;
	const newRow = [
		zonaRow[0]  ?? '', // A - WKT           ← de zona
		zonaRow[1]  ?? '', // B - nombre         ← de zona
		activoRow[0] ?? '', // C - CODIGO        ← de activos
		activoRow[1] ?? '', // D - RAZON_SOCIAL  ← de activos
		activoRow[2] ?? '', // E - Teléfono
		activoRow[3] ?? '', // F - Condición IVA
		activoRow[4] ?? '', // G - Domicilio
		activoRow[5] ?? '', // H - Departamento
		activoRow[6] ?? '', // I - Provincia
		activoRow[7] ?? '', // J - Código Zona
		activoRow[8] ?? '', // K - Zona
	];

	outputRows.push({ row: newRow, color: null });
});

// --- PARTE 4: GENERAR EL EXCEL DE SALIDA ---

const outputHeader = ['WKT', 'nombre', 'CODIGO', 'RAZON_SOCIAL', 'Teléfono', 'Condición IVA', 'Domicilio', 'Departamento', 'Provincia', 'Código Zona', 'Zona'];

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
