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
		.toLowerCase()
		.normalize("NFD")
		.replace(/[\u0300-\u036f]/g, "")
		.trim()
		.replace(/\s+/g, ' ');
}

// Lectura
const activosHtml = fs.readFileSync('clientes-activos_archivos/sheet001.htm', 'latin1');
const activosRows = parseTable(activosHtml);
const activosHeader = activosRows[0];
const activosData = activosRows.slice(1);

const wbZona = xlsx.readFile('clientes-zona.xlsx');
const wsZona = wbZona.Sheets[wbZona.SheetNames[0]];
const zonaRows = xlsx.utils.sheet_to_json(wsZona, { header: 1 });
const zonaHeader = zonaRows[0];
const zonaData = zonaRows.slice(1);

// Map
const activosMap = new Map();

activosData.forEach(row => {
	const razonSocial = normalizeStr(row[1]);
	if (!razonSocial) return;

	const candidato = {
		codigo: row[0],
		telefono: row[2],
		domicilio: normalizeStr(row[4]),
		codigoZona: row[7],
		zona: row[8],
	};
	if (!activosMap.has(razonSocial)) {
		activosMap.set(razonSocial, []);
	}
	activosMap.get(razonSocial).push(candidato);
});

// --- FUNCIÓN FUZZY: DICE COEFFICIENT ---

function diceCoefficient(a, b) {
	if (!a || !b) return 0;
	if (a === b) return 1;

	const bigrams = str => {
		const result = new Set();
		for (let i = 0; i < str.length - 1; i++) {
			result.add(str[i] + str[i + 1]);
		}
		return result;
	};

	const bigramsA = bigrams(a);
	const bigramsB = bigrams(b);

	let intersection = 0;
	bigramsA.forEach(bg => {
		if (bigramsB.has(bg)) intersection++;
	});

	return (2 * intersection) / (bigramsA.size + bigramsB.size);
}

const FUZZY_THRESHOLD = 0.85;

// --- PARTE 3: COMPARACIÓN Y ACTUALIZACIÓN ---

const outputRows = [];
let countExacto = 0;
let countFuzzy = 0;
let countRojo = 0;
let countSinMatch = 0;

zonaData.forEach(row => {
	const razonSocial = normalizeStr(row[1]);
	const candidatos = activosMap.get(razonSocial);
	const newRow = row.map(c => c ?? '');

	// 0 matches — sin tocar
	if (!candidatos || candidatos.length === 0) {
		countSinMatch++;
		outputRows.push({ row: newRow, color: null });
		return;
	}

	// 1 match exacto — actualizar directo
	if (candidatos.length === 1) {
		const match = candidatos[0];
		newRow[0] = match.codigo;
		newRow[2] = match.telefono;
		newRow[7] = match.codigoZona;
		newRow[8] = match.zona;
		countExacto++;
		outputRows.push({ row: newRow, color: null });
		return;
	}

	// 2+ matches — desempate por fuzzy en domicilio
	const domicilioZona = normalizeStr(row[4]);
	const superanUmbral = candidatos.filter(c =>
		diceCoefficient(domicilioZona, c.domicilio) >= FUZZY_THRESHOLD
	);

	if (superanUmbral.length === 1) {
		const match = superanUmbral[0];
		newRow[0] = match.codigo;
		newRow[2] = match.telefono;
		newRow[7] = match.codigoZona;
		newRow[8] = match.zona;
		countFuzzy++;
		outputRows.push({ row: newRow, color: null });
		return;
	}

	// Empate, ninguno supera, o 2+ superan — rojo
	countRojo++;
	outputRows.push({ row: newRow, color: '#FF0000' });
});

// --- PARTE 4: AGREGAR CLIENTES NUEVOS (celeste) ---

const razonesEnZona = new Set(
	zonaData.map(row => normalizeStr(row[1]))
);

let countNuevos = 0;

activosData.forEach(row => {
	const razonSocial = normalizeStr(row[1]);
	if (!razonSocial) return;
	if (razonesEnZona.has(razonSocial)) return;

	// Construir fila con el mismo ancho que zonaHeader
	const newRow = new Array(zonaHeader.length).fill('');
	newRow[0] = row[0] ?? '';  // CODIGO
	newRow[1] = row[1] ?? '';  // RAZON_SOCIAL
	newRow[2] = row[2] ?? '';  // Teléfono
	newRow[4] = row[4] ?? '';  // Domicilio
	newRow[7] = row[7] ?? '';  // Código Zona
	newRow[8] = row[8] ?? '';  // Zona

	outputRows.push({ row: newRow, color: '#ADD8E6' });
	countNuevos++;
});

// --- PARTE 5: GENERAR OUTPUT HTML-as-XLS ---

function cellStyle(color) {
	return color ? ` style="background-color:${color}"` : '';
}

function escapeHtml(val) {
	return String(val ?? '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;');
}

let html = '<html><head><meta charset="windows-1252"></head><body><table>';

// Header
html += '<tr>';
for (const col of zonaHeader) {
	html += `<th>${escapeHtml(col)}</th>`;
}
html += '</tr>';

// Filas
for (const { row, color } of outputRows) {
	html += `<tr${cellStyle(color)}>`;
	for (const cell of row) {
		html += `<td>${escapeHtml(cell)}</td>`;
	}
	html += '</tr>';
}

html += '</table></body></html>';

fs.writeFileSync('clientes-zona-actualizado.xls', html, 'latin1');

// Log final
console.log(`Filas en clientes-zona:        ${zonaData.length}`);
console.log(`Actualizadas con match exacto: ${countExacto}`);
console.log(`Actualizadas con fuzzy:        ${countFuzzy}`);
console.log(`Ambiguas (marcadas en rojo):   ${countRojo}`);
console.log(`Sin match:                     ${countSinMatch}`);
console.log(`Clientes nuevos agregados:     ${countNuevos}`);
console.log(`Archivo generado: clientes-zona-actualizado.xls`);