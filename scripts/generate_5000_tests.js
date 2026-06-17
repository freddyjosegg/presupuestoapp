const fs = require('fs');
const path = require('path');
// Cargar SheetJS localmente (es compatible con CommonJS en Node.js)
const XLSX = require('../js/xlsx.full.min.js');

// === FUNCIONES MATEMÁTICAS DE REDONDEO DE LA APP (IDÉNTICAS A MAIN.JS) ===
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

// === MOTOR DE CÁLCULO (IDÉNTICO A RUNSCENARIO EN MAIN.JS) ===
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
    
    // 3. TDG (Tramite de guia): if weight > 160kg, apply tdgBase
    const applyTdg = weight > 160;
    const tdgVal = applyTdg ? constants.tdgBase : 0;
    
    // 4. GCD (Cobro a destino)
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
    
    // 8. Subtotal (Base Imponible)
    const subtotal = netFreight + tdgVal + gcdVal + carVal + carfVal + seguroVal + containerVal;
    
    // 9. IVA (16%)
    const ivaVal = truncDec(subtotal * 0.16, 3);
    
    // 10. Franqueo Postal (IPOSTEL 10%)
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
    
    // 11. IGTF (3%)
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

// === LECTURA Y PROCESAMIENTO DE EXCEL BASE ===
function loadExcelDatabase() {
    const dbDir = path.resolve(__dirname, '../DB');
    const files = fs.readdirSync(dbDir);
    const excelFile = files.find(file => {
        const ext = path.extname(file).toLowerCase();
        return (ext === '.xlsx' || ext === '.xlsm') && !file.startsWith('~$');
    });

    if (!excelFile) {
        throw new Error('No se encontró el archivo Excel de base de datos en la carpeta DB.');
    }

    const filePath = path.join(dbDir, excelFile);
    console.log(`Cargando archivo de tarifas: ${excelFile}...`);
    const fileBuffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

    const sheet_cf = workbook.Sheets['CONSULTA FLETE USD'];
    const sheet_dir = workbook.Sheets['TARIFA ENTREGA A DIRECCION USD'];
    const sheet_age = workbook.Sheets['TARIFA CONSIGNADO AGENCIA USD'];

    // Parser de constantes
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

    // Parser de rutas
    const cfRows = XLSX.utils.sheet_to_json(sheet_cf, { header: 1 });
    const routes = [];
    for (let i = 1; i < cfRows.length; i++) {
        const row = cfRows[i];
        if (!row || !row[1] || !row[3]) continue;
        const originCode = String(row[1]).trim().toUpperCase();
        const originName = String(row[2]).trim();
        const destCode = String(row[3]).trim().toUpperCase();
        const destName = String(row[4]).trim();
        const escala = row[5] ? String(row[5]).trim() : '';
        const kms = parseFloat(row[6]) || 0;
        
        if (escala && escala !== '#N/A') {
            routes.push({ originCode, originName, destCode, destName, escala, kms });
        }
    }

    // Parser de tarifas
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

    return { routes, tariffs, constants };
}

// === SIMULACIÓN DE 5000 PRESUPUESTOS ===
function generateTests() {
    const { routes, tariffs, constants } = loadExcelDatabase();
    console.log(`Base de datos leída con éxito: ${routes.length} rutas válidas disponibles.`);
    
    console.log('Generando 5000 presupuestos con datos aleatorios...');
    const resultsData = [];

    for (let id = 1; id <= 5000; id++) {
        // Seleccionar ruta aleatoria
        const route = routes[Math.floor(Math.random() * routes.length)];
        
        // Parámetros aleatorios
        const count = Math.floor(Math.random() * 5) + 1; // 1 a 5 bultos
        const bultos = [];
        for (let j = 0; j < count; j++) {
            // Peso de bultos: 70% bultos normales (0.1 - 30kg), 30% bultos grandes (30 - 150kg)
            const isLarge = Math.random() < 0.3;
            const weight = isLarge ? (Math.random() * 120 + 30) : (Math.random() * 29.9 + 0.1);
            bultos.push(parseFloat(weight.toFixed(3)));
        }
        const physWeight = parseFloat(bultos.reduce((a, b) => a + b, 0).toFixed(3));

        // Dimensiones aleatorias (35% de probabilidad)
        const hasDimensions = Math.random() < 0.35;
        const largo = hasDimensions ? Math.floor(Math.random() * 90) + 10 : 0;
        const ancho = hasDimensions ? Math.floor(Math.random() * 90) + 10 : 0;
        const alto = hasDimensions ? Math.floor(Math.random() * 90) + 10 : 0;
        const volWeight = hasDimensions ? parseFloat(((largo * ancho * alto) / 5000).toFixed(3)) : 0;
        
        const finalWeight = Math.max(physWeight, volWeight);

        const declaredVal = parseFloat((Math.random() * 600).toFixed(2)); // $0.00 a $600.00
        const exRate = parseFloat((Math.random() * 15 + 35).toFixed(2)); // Tasa de cambio Bs. 35.00 a 50.00

        // Opciones adicionales aleatorias
        const useSeguro = Math.random() < 0.7; // 70% seguro activado
        const useCar = Math.random() < 0.25;  // 25% CAR activado
        const useCarf = Math.random() < 0.15; // 15% CARF activado
        const useContenedor = Math.random() < 0.2; // 20% Contenedor

        // Descuentos (0%, 5%, 10%, 15%, 20%)
        const discountOptions = [0, 0, 0, 5, 10, 15, 20];
        const discPercent = discountOptions[Math.floor(Math.random() * discountOptions.length)];

        // Ejecutar los 4 escenarios para la fila actual
        const resOrigDir = runScenario(tariffs, constants, bultos, 'direccion', false, finalWeight, route.escala, declaredVal, discPercent, count, exRate, useSeguro, useCar, useCarf, useContenedor);
        const resOrigAge = runScenario(tariffs, constants, bultos, 'agencia', false, finalWeight, route.escala, declaredVal, discPercent, count, exRate, useSeguro, useCar, useCarf, useContenedor);
        const resDestDir = runScenario(tariffs, constants, bultos, 'direccion', true, finalWeight, route.escala, declaredVal, discPercent, count, exRate, useSeguro, useCar, useCarf, useContenedor);
        const resDestAge = runScenario(tariffs, constants, bultos, 'agencia', true, finalWeight, route.escala, declaredVal, discPercent, count, exRate, useSeguro, useCar, useCarf, useContenedor);

        resultsData.push({
            'ID Test': id,
            'Código Origen': route.originCode,
            'Nombre Origen': route.originName,
            'Código Destino': route.destCode,
            'Nombre Destino': route.destName,
            'Escala': route.escala,
            'Distancia (km)': route.kms,
            'Cantidad Bultos': count,
            'Pesos Bultos (kg)': bultos.join(', '),
            'Peso Físico Total (kg)': physWeight,
            'Largo (cm)': largo || '',
            'Ancho (cm)': ancho || '',
            'Alto (cm)': alto || '',
            'Peso Volumétrico (kg)': volWeight || '',
            'Peso de Cálculo (kg)': finalWeight,
            'Valor Declarado ($)': declaredVal,
            'Tasa de Cambio (Bs.)': exRate,
            'Seguro Activado': useSeguro ? 'Sí' : 'No',
            'CAR Activado': useCar ? 'Sí' : 'No',
            'CARF Activado': useCarf ? 'Sí' : 'No',
            'Contenedor Activado': useContenedor ? 'Sí' : 'No',
            'Descuento (%)': discPercent ? `${discPercent}%` : '0%',
            
            // Resultados - Origen Dirección
            'Orig-Dir Base ($)': resOrigDir.baseFreight,
            'Orig-Dir IVA ($)': resOrigDir.ivaVal,
            'Orig-Dir Franqueo ($)': resOrigDir.franqueoVal,
            'Orig-Dir IGTF ($)': resOrigDir.igtfVal,
            'Orig-Dir Total ($)': resOrigDir.totalUSDSinIgtf,
            'Orig-Dir Total con IGTF ($)': resOrigDir.totalUSD,
            'Orig-Dir Total (Bs.)': resOrigDir.totalBs,
            
            // Resultados - Origen Agencia
            'Orig-Age Total con IGTF ($)': resOrigAge.totalUSD,
            'Orig-Age Total (Bs.)': resOrigAge.totalBs,

            // Resultados - Destino Dirección (COD)
            'Dest-Dir Total con IGTF ($)': resDestDir.totalUSD,
            'Dest-Dir Total (Bs.)': resDestDir.totalBs,

            // Resultados - Destino Agencia (COD)
            'Dest-Age Total con IGTF ($)': resDestAge.totalUSD,
            'Dest-Age Total (Bs.)': resDestAge.totalBs
        });

        if (id % 1000 === 0) {
            console.log(`Generados ${id} presupuestos...`);
        }
    }

    console.log('Creando hoja de cálculo de resultados...');
    const worksheet = XLSX.utils.json_to_sheet(resultsData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Presupuestos Generados');

    const outputName = 'presupuestos_5000_test.xlsx';
    const outputPath = path.resolve(__dirname, '..', outputName);
    
    console.log(`Escribiendo archivo de resultados en: ${outputName}...`);
    const wopts = { bookType: 'xlsx', bookSST: false, type: 'buffer' };
    const wbout = XLSX.write(workbook, wopts);
    fs.writeFileSync(outputPath, wbout);
    console.log('✅ Archivo Excel generado correctamente.');
}

generateTests();
