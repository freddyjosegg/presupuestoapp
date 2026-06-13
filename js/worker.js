console.log('[Worker] Script cargado en el hilo del worker.');
try {
    importScripts('xlsx.full.min.js');
    console.log('[Worker] xlsx.full.min.js cargado correctamente.');
} catch (err) {
    console.error('[Worker ERROR] Falló importScripts:', err);
}

self.onmessage = function (e) {
    console.log('[Worker] Mensaje recibido en onmessage.');
    const { arrayBuffer } = e.data;
    try {
        console.log('[Worker] Tipo de arrayBuffer recibido:', typeof arrayBuffer, arrayBuffer ? arrayBuffer.byteLength : 'null');
        const data = new Uint8Array(arrayBuffer);
        console.log('[Worker] Uint8Array creado con éxito. Tamaño:', data.length);
        
        console.log('[Worker] Llamando a XLSX.read...');
        // Optimize XLSX read options to load only values, skipping styles, formulas and formatted texts
        const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: false,
            cellNF: false,
            cellHTML: false,
            cellFormula: false,
            cellText: false
        });
        console.log('[Worker] XLSX.read finalizado con éxito.');
        
        console.log('[Worker] Obteniendo hojas de Excel...');
        const sheet_cf = workbook.Sheets['CONSULTA FLETE USD'];
        const sheet_dir = workbook.Sheets['TARIFA ENTREGA A DIRECCION USD'];
        const sheet_age = workbook.Sheets['TARIFA CONSIGNADO AGENCIA USD'];
        
        if (!sheet_cf || !sheet_dir || !sheet_age) {
            throw new Error("El archivo de Excel no contiene las hojas requeridas ('CONSULTA FLETE USD', 'TARIFA ENTREGA A DIRECCION USD', 'TARIFA CONSIGNADO AGENCIA USD').");
        }
        
        // Optimize cell ranges to limit columns to A-P (16 columns), preventing SheetJS from parsing 16k+ columns
        [sheet_dir, sheet_age].forEach(sheet => {
            if (sheet && sheet['!ref']) {
                const refParts = sheet['!ref'].split(':');
                if (refParts.length === 2) {
                    const endRow = refParts[1].replace(/^[A-Z]+/i, '');
                    sheet['!ref'] = `A1:P${endRow}`;
                }
            }
        });
        
        console.log('[Worker] Hojas validadas y rangos recortados con éxito. Parseando CONSULTA FLETE USD...');
        // 1. Parse CONSULTA FLETE USD
        const cfRows = XLSX.utils.sheet_to_json(sheet_cf, { header: 1 });
        console.log('[Worker] cfRows parseado a JSON. Cantidad de filas:', cfRows.length);
        
        // Parse routes
        console.log('[Worker] Parseando rutas...');
        const routes = {};
        for (let i = 1; i < cfRows.length; i++) {
            const row = cfRows[i];
            if (!row || !row[1] || !row[3]) continue;
            const originCode = String(row[1]).trim().toUpperCase();
            const destCode = String(row[3]).trim().toUpperCase();
            const escala = row[5] ? String(row[5]).trim() : '';
            const kms = parseFloat(row[6]) || 0;
            routes[`${originCode}_${destCode}`] = { escala, kms };
        }
        console.log('[Worker] Rutas parseadas con éxito:', Object.keys(routes).length);
        
        // Parse origins
        console.log('[Worker] Parseando orígenes...');
        const originsMap = new Map();
        for (let i = 2; i < cfRows.length; i++) {
            const row = cfRows[i];
            if (!row) continue;
            const code = row[8] ? String(row[8]).trim().toUpperCase() : null;
            const name = row[9] ? String(row[9]).trim() : null;
            const reparte = row[10] ? String(row[10]).trim() : '';
            
            if (code && code !== 'O' && code !== 'ORIGEN') {
                originsMap.set(code, { code, name: name || code, reparte });
            }
        }
        const origins = Array.from(originsMap.values()).sort((a, b) => a.name.localeCompare(b.name));
        console.log('[Worker] Orígenes parseados con éxito:', origins.length);
        
        // Parse destinations
        console.log('[Worker] Parseando destinos...');
        const destMap = new Map();
        for (let i = 2; i < cfRows.length; i++) {
            const row = cfRows[i];
            if (!row) continue;
            const code = row[12] ? String(row[12]).trim().toUpperCase() : null;
            const name = row[13] ? String(row[13]).trim() : null;
            const reparte = row[14] ? String(row[14]).trim() : '';
            
            if (code && code !== 'O' && code !== 'DESTINO') {
                destMap.set(code, { code, name: name || code, reparte });
            }
        }
        const destinations = Array.from(destMap.values()).sort((a, b) => a.name.localeCompare(b.name));
        console.log('[Worker] Destinos parseados con éxito:', destinations.length);
        
        // 2. Parse tariffs
        console.log('[Worker] Parseando tarifas...');
        const tariffs = {
            direccion: parseTariffRows(sheet_dir),
            agencia: parseTariffRows(sheet_age)
        };
        console.log('[Worker] Tarifas parseadas con éxito. Dirección:', tariffs.direccion.length, 'Agencia:', tariffs.agencia.length);
        
        // 3. Parse constants
        console.log('[Worker] Parseando constantes...');
        const constants = {
            tdgBase: getCellValue(sheet_dir, 'O23', 0.6),
            gcdMin: getCellValue(sheet_dir, 'O26', 0.8),
            carBase: getCellValue(sheet_dir, 'O29', 1.2),
            carfBase: getCellValue(sheet_dir, 'O32', 1.8),
            seguroMin: getCellValue(sheet_dir, 'O36', 1.76),
            seguroRate: getCellValue(sheet_dir, 'O37', 3.2),
            containerMin: getCellValue(sheet_dir, 'O40', 0.75)
        };
        console.log('[Worker] Constantes parseadas con éxito:', JSON.stringify(constants));
        
        console.log('[Worker] Posteo de mensaje de vuelta al hilo principal...');
        self.postMessage({
            success: true,
            data: {
                origins,
                destinations,
                routes,
                tariffs,
                constants
            }
        });
        console.log('[Worker] Mensaje enviado.');
    } catch (error) {
        self.postMessage({
            success: false,
            error: error.message || String(error)
        });
    }
};

// Helper: Parse tariff rows
function parseTariffRows(sheet) {
    console.log('[Worker - parseTariffRows] Llamando a sheet_to_json...');
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    console.log('[Worker - parseTariffRows] sheet_to_json finalizado. Cantidad de filas:', rows.length);
    
    const tariffsList = [];
    for (let i = 7; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        
        const label = row[0] ? String(row[0]).trim() : '';
        const limitVal = parseFloat(row[1]);
        
        if (!label) continue;
        
        let limit = limitVal;
        if (isNaN(limitVal)) {
            if (label.includes('0,500') || label.includes('0.500')) {
                limit = 0.5;
            } else {
                continue;
            }
        }
        
        tariffsList.push({
            label,
            limit,
            Y01: parseFloat(row[2]) || 0,
            Y02: parseFloat(row[3]) || 0,
            Y03: parseFloat(row[4]) || 0,
            Y04: parseFloat(row[5]) || 0,
            Y05: parseFloat(row[6]) || 0
        });
    }
    console.log('[Worker - parseTariffRows] Bucle finalizado. Cantidad de tarifas parseadas:', tariffsList.length);
    console.log('[Worker - parseTariffRows] Ordenando tarifas...');
    const sorted = tariffsList.sort((a, b) => a.limit - b.limit);
    console.log('[Worker - parseTariffRows] Tarifas ordenadas.');
    return sorted;
}

// Helper: Get cell value in sheet
function getCellValue(sheet, address, defaultValue) {
    const cell = sheet[address];
    if (!cell) return defaultValue;
    const val = parseFloat(cell.v);
    return isNaN(val) ? defaultValue : val;
}
