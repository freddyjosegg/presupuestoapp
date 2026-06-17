const assert = require('assert');
const path = require('path');
const fs = require('fs');
const XLSX = require('../js/xlsx.full.min.js');

// === IMPORTAR FUNCIONES MATEMÁTICAS A PROBAR ===
function roundHalfUpNum(val, decimals = 2) {
    if (val === undefined || val === null || isNaN(val)) return 0;
    const factor = Math.pow(10, decimals);
    return Math.round((val + 1e-9) * factor) / factor;
}

function truncDec(val, decimals = 3) {
    if (val === undefined || val === null || isNaN(val)) return 0;
    const factor = Math.pow(10, decimals);
    return Math.floor((val + 1e-9) * factor) / factor;
}

function roundUpDec(val, decimals = 3) {
    if (val === undefined || val === null || isNaN(val)) return 0;
    const factor = Math.pow(10, decimals);
    return Math.ceil((val - 1e-9) * factor) / factor;
}

function runScenario(tariffs, constants, bultos, tariffType, isCOD, weight, escala, declaredVal, discPercent, count, exRate, useSeguro, useCar, useCarf, useContenedor) {
    const list = tariffs[tariffType];
    
    // 1. Base Freight Lookup
    let baseFreight = 0;
    if (list && list.length > 0) {
        const matchedRow = list.find(r => r.limit >= weight);
        if (matchedRow) {
            baseFreight = matchedRow[escala] || 0;
        } else {
            baseFreight = list[list.length - 1][escala] || 0;
        }
    }
    
    // 2. Discount
    const hasDiscount = discPercent > 0;
    const discountVal = hasDiscount ? baseFreight * (discPercent / 100) : 0;
    const netFreight = roundUpDec(baseFreight - discountVal, 3);
    
    // 3. TDG
    const applyTdg = weight > 160;
    const tdgVal = applyTdg ? constants.tdgBase : 0;
    
    // 4. GCD
    let gcdVal = 0;
    if (isCOD) {
        gcdVal = Math.max(netFreight * 0.1, constants.gcdMin);
    }
    
    // 5. CAR / CARF
    const carVal = useCar ? constants.carBase : 0;
    const carfVal = useCarf ? constants.carfBase : 0;
    
    // 6. Seguro
    let seguroVal = 0;
    if (useSeguro && declaredVal > 0) {
        if (declaredVal <= 55) {
            seguroVal = constants.seguroMin;
        } else {
            seguroVal = declaredVal * (constants.seguroRate / 100);
        }
        seguroVal = roundUpDec(seguroVal, 3);
    }
    
    // 7. Contenedor
    const containerVal = useContenedor ? (baseFreight * 0.1) : 0;
    
    // 8. Subtotal
    const subtotal = netFreight + tdgVal + gcdVal + carVal + carfVal + seguroVal + containerVal;
    
    // 9. IVA
    const ivaVal = truncDec(subtotal * 0.16, 3);
    
    // 10. Franqueo Postal
    let franqueoVal = 0;
    if (bultos && bultos.length > 0) {
        const sumBultosWeight = bultos.reduce((a, b) => a + b, 0) || 1;
        let sumFranqueo = 0;
        bultos.forEach(bw => {
            if (bw <= 30) {
                const propBaseFreight = (bw / sumBultosWeight) * baseFreight;
                sumFranqueo += propBaseFreight * 0.1;
            }
        });
        franqueoVal = roundUpDec(sumFranqueo, 3);
    }
    
    // 11. IGTF
    const rawIgtfBase = subtotal + ivaVal + franqueoVal;
    const igtfBase = roundHalfUpNum(rawIgtfBase, 2);
    const igtfVal = igtfBase * 0.03;
    
    // 12. Totales
    const totalUSD = igtfBase + igtfVal;
    
    const fleteBs = roundUpDec(netFreight * exRate, 3);
    const compBs = roundUpDec((tdgVal + gcdVal + carVal + carfVal + seguroVal + containerVal) * exRate, 3);
    const taxBs = roundUpDec((ivaVal + franqueoVal) * exRate, 3);
    
    const totalBsRaw = roundUpDec(fleteBs + compBs + taxBs, 3);
    const totalBs = roundHalfUpNum(totalBsRaw, 2);
    
    return {
        baseFreight,
        discountVal,
        netFreight,
        tdgVal,
        gcdVal,
        carVal,
        carfVal,
        seguroVal,
        containerVal,
        subtotal,
        ivaVal,
        franqueoVal,
        igtfVal,
        totalUSDSinIgtf: igtfBase,
        totalUSD,
        totalBs
    };
}

// === CARGAR BASE DE DATOS LOCAL ===
function loadLocalDatabase() {
    const dbDir = path.resolve(__dirname, '../DB');
    const files = fs.readdirSync(dbDir);
    const excelFile = files.find(file => {
        const ext = path.extname(file).toLowerCase();
        return (ext === '.xlsx' || ext === '.xlsm') && !file.startsWith('~$');
    });

    if (!excelFile) {
        throw new Error('No se encontró el archivo Excel de tarifas.');
    }

    const filePath = path.join(dbDir, excelFile);
    const fileBuffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

    const sheet_dir = workbook.Sheets['TARIFA ENTREGA A DIRECCION USD'];
    const sheet_age = workbook.Sheets['TARIFA CONSIGNADO AGENCIA USD'];

    const getCellValue = (sheet, address, defaultValue) => {
        const cell = sheet[address];
        if (!cell) return defaultValue;
        const val = parseFloat(cell.v);
        return isNaN(val) ? defaultValue : val;
    };

    const constants = {
        tdgBase: getCellValue(sheet_dir, 'O23', 0.6),
        gcdMin: getCellValue(sheet_dir, 'O26', 0.8),
        carBase: getCellValue(sheet_dir, 'O29', 1.2),
        carfBase: getCellValue(sheet_dir, 'O32', 1.8),
        seguroMin: getCellValue(sheet_dir, 'O36', 1.76),
        seguroRate: getCellValue(sheet_dir, 'O37', 3.2),
        containerMin: getCellValue(sheet_dir, 'O40', 0.75)
    };

    const parseTariffRows = (sheet) => {
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
    };

    const tariffs = {
        direccion: parseTariffRows(sheet_dir),
        agencia: parseTariffRows(sheet_age)
    };

    return { tariffs, constants };
}

// === PRUEBAS UNITARIAS DE CASOS MATEMÁTICOS ===
function testMath() {
    console.log('Iniciando Pruebas Unitarias de Fórmulas Matemáticas...');
    const { tariffs, constants } = loadLocalDatabase();

    // --- Caso 1: Caso Base con 40kg, Origen Dirección, Escala Y03, Valor Declarado $55, Tasa de cambio 36.50 ---
    console.log('Ejecutando Caso 1 (40kg, Y03, con Seguro Mínimo)...');
    const bultos1 = [40];
    const res1 = runScenario(tariffs, constants, bultos1, 'direccion', false, 40, 'Y03', 55, 0, 1, 36.50, true, false, false, false);
    // Valores de tarifa esperados para Y03 en rango de 40kg:
    // flete base para 40kg en Y03 es $98.65
    assert.ok(Math.abs(res1.baseFreight - 98.65) < 0.01, 'Flete base incorrecto');
    assert.ok(Math.abs(res1.netFreight - 98.65) < 0.01, 'Flete neto con 0% descuento incorrecto');
    // Subtotal: flete + seguro mínimo ($1.76) = 98.65 + 1.76 = 100.41
    assert.ok(Math.abs(res1.subtotal - 100.41) < 0.01, 'Subtotal incorrecto');
    // IVA: TRUNC(100.41 * 0.16, 3) = 16.065
    assert.ok(Math.abs(res1.ivaVal - 16.065) < 0.01, 'IVA incorrecto');
    // Franqueo: Bulto es 40 > 30, por lo que aplica 0 de franqueo (solo aplica a bultos <= 30)
    assert.ok(res1.franqueoVal === 0, 'Franqueo Ipostel incorrecto para bulto > 30kg');
    // IGTF Base: round(100.41 + 16.065 + 0, 2) = round(116.475, 2) = 116.48
    // IGTF: 116.48 * 0.03 = 3.4944 (y total con IGTF = 116.48 + 3.4944 = 119.9744 -> visual es 119.97)
    assert.ok(Math.abs(res1.totalUSDSinIgtf - 116.48) < 0.01, 'Total USD sin IGTF incorrecto');
    assert.ok(Math.abs(res1.totalUSD - (116.48 * 1.03)) < 0.01, 'Total USD con IGTF incorrecto');
    // Total Bs (Fórmula de Excel de redondeo por componente):
    // fleteBs = roundup(98.65 * 36.5, 3) = roundup(3600.725, 3) = 3600.725
    // compBs = roundup(1.76 * 36.5, 3) = roundup(64.24, 3) = 64.240
    // taxBs = roundup(16.065 * 36.5, 3) = roundup(586.3725, 3) = 586.373
    // totalBsRaw = roundup(3600.725 + 64.240 + 586.373, 3) = roundup(4251.338, 3) = 4251.338
    // totalBs = roundHalfUp(4251.338, 2) = 4251.34
    assert.ok(Math.abs(res1.totalBs - 4251.34) < 0.1, 'Total Bolívares incorrecto para el Caso 1');

    // --- Caso 2: Peso > 160kg (Aplica TDG), con CAR, CARF, Contenedor y Descuento del 10% ---
    console.log('Ejecutando Caso 2 (200kg, Y01, con TDG, Contenedor, CAR/CARF y 10% Descuento)...');
    const bultos2 = [200];
    const res2 = runScenario(tariffs, constants, bultos2, 'direccion', false, 200, 'Y01', 0, 10, 1, 38.20, false, true, true, true);
    
    // Si peso es 200kg, el flete base para Y01 es $159.56
    // Descuento 10% = 15.956
    // netFreight = roundup(159.56 - 15.956, 3) = roundup(143.604, 3) = 143.604
    assert.ok(Math.abs(res2.baseFreight - 159.56) < 0.01, 'Flete base Y01 incorrecto');
    assert.ok(Math.abs(res2.netFreight - 143.60) < 0.01, 'Flete neto con 10% descuento incorrecto');
    
    // TDG = 0.6 (peso > 160)
    assert.ok(Math.abs(res2.tdgVal - 0.6) < 0.01, 'Trámite de guía incorrecto');
    // Contenedor: 10% del Flete Base = 15.956
    assert.ok(Math.abs(res2.containerVal - 15.956) < 0.01, 'Cargo por contenedor incorrecto');
    // CAR (1.2) y CARF (1.8)
    assert.ok(Math.abs(res2.carVal - 1.2) < 0.01, 'CAR incorrecto');
    assert.ok(Math.abs(res2.carfVal - 1.8) < 0.01, 'CARF incorrecto');

    // --- Caso 3: Bulto pequeño (<= 30kg) donde IPOSTEL (10%) sí aplica ---
    console.log('Ejecutando Caso 3 (15kg, Y02, con IPOSTEL)...');
    const bultos3 = [15];
    const res3 = runScenario(tariffs, constants, bultos3, 'direccion', false, 15, 'Y02', 0, 0, 1, 36.50, false, false, false, false);
    
    // Flete base de 15kg en Y02 = $45.60 (referencia aproximada)
    // Franqueo: 10% de flete base ($45.60 * 0.1) = $4.56 ya que bulto es <= 30
    assert.ok(Math.abs(res3.franqueoVal - roundUpDec(res3.baseFreight * 0.1, 3)) < 0.01, 'IPOSTEL no se aplicó o se calculó mal para bulto <= 30kg');

    console.log('✅ TODAS LAS PRUEBAS UNITARIAS MATEMÁTICAS PASARON EXITOSAMENTE.');
}

testMath();
