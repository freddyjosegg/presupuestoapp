importScripts('xlsx.full.min.js');

self.onmessage = function (e) {
    const { arrayBuffer } = e.data;
    try {
        const data = new Uint8Array(arrayBuffer);
        
        // Optimize XLSX read options to load only values, skipping styles, formulas and formatted texts
        const workbook = XLSX.read(data, {
            type: 'array',
            cellDates: false,
            cellNF: false,
            cellHTML: false,
            cellFormula: false,
            cellText: false
        });
        
        const sheet_cf = workbook.Sheets['CONSULTA FLETE USD'];
        const sheet_dir = workbook.Sheets['TARIFA ENTREGA A DIRECCION USD'];
        const sheet_age = workbook.Sheets['TARIFA CONSIGNADO AGENCIA USD'];
        
        if (!sheet_cf || !sheet_dir || !sheet_age) {
            throw new Error("El archivo de Excel no contiene las hojas requeridas ('CONSULTA FLETE USD', 'TARIFA ENTREGA A DIRECCION USD', 'TARIFA CONSIGNADO AGENCIA USD').");
        }
        
        // 1. Parse CONSULTA FLETE USD
        const cfRows = XLSX.utils.sheet_to_json(sheet_cf, { header: 1 });
        
        // Parse routes
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
        
        // Parse origins
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
        
        // Parse destinations
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
        
        // 2. Parse tariffs
        const tariffs = {
            direccion: parseTariffRows(sheet_dir),
            agencia: parseTariffRows(sheet_age)
        };
        
        // 3. Parse constants
        const constants = {
            tdgBase: getCellValue(sheet_dir, 'O23', 0.6),
            gcdMin: getCellValue(sheet_dir, 'O26', 0.8),
            carBase: getCellValue(sheet_dir, 'O29', 1.2),
            carfBase: getCellValue(sheet_dir, 'O32', 1.8),
            seguroMin: getCellValue(sheet_dir, 'O36', 1.76),
            seguroRate: getCellValue(sheet_dir, 'O37', 3.2),
            containerMin: getCellValue(sheet_dir, 'O40', 0.75)
        };
        
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
    } catch (error) {
        self.postMessage({
            success: false,
            error: error.message || String(error)
        });
    }
};

// Helper: Parse tariff rows
function parseTariffRows(sheet) {
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
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
    return tariffsList.sort((a, b) => a.limit - b.limit);
}

// Helper: Get cell value in sheet
function getCellValue(sheet, address, defaultValue) {
    const cell = sheet[address];
    if (!cell) return defaultValue;
    const val = parseFloat(cell.v);
    return isNaN(val) ? defaultValue : val;
}
