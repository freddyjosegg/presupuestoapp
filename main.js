// Configuration and Constants
const DEFAULT_EXCEL_PATH = 'TARIFA ACTUAL PAQUETERIA USD 210526.xlsm';
const STORAGE_KEY = 'presupuestoapp_db';
const BCV_API_URL = 'https://ve.dolarapi.com/v1/dolares/oficial';

// State of the application
let appState = {
    origins: [],        // { code, name, reparte }
    destinations: [],   // { code, name, reparte }
    routes: {},         // key: 'ORIG_DEST' -> { escala, kms }
    tariffs: {
        direccion: [],  // array of { limit, Y01, Y02, Y03, Y04, Y05 }
        agencia: []     // array of { limit, Y01, Y02, Y03, Y04, Y05 }
    },
    constants: {
        tdgBase: 0.6,
        gcdMin: 0.8,
        carBase: 1.2,
        carfBase: 1.8,
        seguroMin: 1.76,
        seguroRate: 3.2, // in percent (3.2%)
        containerMin: 0.75
    },
    selectedOption: 'origDir', // 'origDir', 'origAge', 'destDir', 'destAge'
    bultos: [], // array of individual weights: [w1, w2, ...]
    excelName: '',
    excelSize: 0,
    excelDate: ''
};

let userEditedRate = false;

// UI Elements
const dropzone = document.getElementById('dropzone');
const excelFileInput = document.getElementById('excelFileInput');
const excelStatus = document.getElementById('excelStatus');
const originInput = document.getElementById('originInput');
const originList = document.getElementById('originList');
const destinationInput = document.getElementById('destinationInput');
const destinationList = document.getElementById('destinationList');
const routeInfoCard = document.getElementById('routeInfoCard');
const valEscala = document.getElementById('valEscala');
const valKms = document.getElementById('valKms');
const valAtiende = document.getElementById('valAtiende');
const weightInput = document.getElementById('weightInput');
const bultosCount = document.getElementById('bultosCount');
const dimLargo = document.getElementById('dimLargo');
const dimAncho = document.getElementById('dimAncho');
const dimAlto = document.getElementById('dimAlto');
const valCubaje = document.getElementById('valCubaje');
const valPesoFinal = document.getElementById('valPesoFinal');
const declaredValue = document.getElementById('declaredValue');
const exchangeRate = document.getElementById('exchangeRate');
const chkSeguro = document.getElementById('chkSeguro');
const chkCar = document.getElementById('chkCar');
const chkCarf = document.getElementById('chkCarf');
const chkContenedor = document.getElementById('chkContenedor');
const discountPercent = document.getElementById('discountPercent');
const toggleBultosBtn = document.getElementById('toggleBultosBtn');
const bultosCard = document.querySelector('.bultos-card');
const bultosContent = document.getElementById('bultosContent');
const bultosInputsContainer = document.getElementById('bultosInputsContainer');
const btnSyncBultos = document.getElementById('btnSyncBultos');
const btnPrint = document.getElementById('btnPrint');
const btnCopy = document.getElementById('btnCopy');
const btnFetchRate = document.getElementById('btnFetchRate');
const toast = document.getElementById('toast');

// Matrix Card Elements
const matrixCards = {
    origDir: document.getElementById('cardOrigDir'),
    origAge: document.getElementById('cardOrigAge'),
    destDir: document.getElementById('cardDestDir'),
    destAge: document.getElementById('cardDestAge')
};

// Initialize Application
window.addEventListener('DOMContentLoaded', () => {
    initApp();
});

function initApp() {
    setupEventListeners();
    
    // Fetch BCV Rate at startup
    fetchBCVRate(true);
    
    // 1. Try loading from LocalStorage
    const cachedData = localStorage.getItem(STORAGE_KEY);
    if (cachedData) {
        try {
            const parsed = JSON.parse(cachedData);
            appState = { ...appState, ...parsed };
            updateStatus('ready', `Cargado de caché local`, `${appState.excelName || 'Excel guardado'}`);
            populateDropdowns();
            calculateAll();
            // Fetch file to check for updates silently
            fetchDefaultExcel(true);
        } catch (e) {
            console.error('Error loading cached database, fetching new one...', e);
            fetchDefaultExcel();
        }
    } else {
        fetchDefaultExcel();
    }
}

// Fetch BCV Exchange Rate from ve.dolarapi.com
async function fetchBCVRate(silent = false) {
    if (btnFetchRate) {
        btnFetchRate.classList.add('loading');
    }
    
    try {
        // Force fresh load using cache-buster query param
        const response = await fetch(`${BCV_API_URL}?t=${Date.now()}`);
        if (!response.ok) throw new Error('Network response was not ok');
        
        const data = await response.json();
        const rate = parseFloat(data.promedio);
        
        if (!isNaN(rate) && rate > 0) {
            // Overwrite only if user hasn't typed a custom rate, or they clicked refresh (not silent)
            if (!userEditedRate || !silent) {
                exchangeRate.value = rate.toFixed(2);
                calculateAll();
                if (!silent) {
                    showToast(`Tasa BCV actualizada: Bs. ${rate.toFixed(2)}`);
                }
            }
        }
    } catch (error) {
        console.error('Error fetching BCV rate:', error);
        if (!silent) {
            showToast('Error al conectar con la API de tasa de cambio', 'error');
        }
    } finally {
        if (btnFetchRate) {
            btnFetchRate.classList.remove('loading');
        }
    }
}

// Fetch default Excel from server
async function fetchDefaultExcel(silent = false) {
    if (!silent) {
        updateStatus('loading', 'Cargando base de datos...', 'Descargando archivo Excel base...');
    }
    
    try {
        const response = await fetch(DEFAULT_EXCEL_PATH);
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        
        appState.excelName = DEFAULT_EXCEL_PATH;
        appState.excelSize = blob.size;
        appState.excelDate = new Date().toLocaleDateString('es-VE');
        
        parseExcelData(arrayBuffer);
        
        // Save to cache
        saveCache();
        
        updateStatus('ready', 'Excel conectado', `${appState.excelName} (${(appState.excelSize/1024/1024).toFixed(2)} MB)`);
        populateDropdowns();
        calculateAll();
        showToast('Base de datos Excel cargada correctamente');
    } catch (error) {
        console.error('Error loading default Excel:', error);
        if (!silent) {
            updateStatus('error', 'Error al cargar Excel', 'Arrastra el archivo Excel manualmente');
            showToast('No se pudo cargar el Excel por defecto. Por favor arrástralo manualmente.', 'error');
        }
    }
}

// Save parsed data to LocalStorage
function saveCache() {
    const dataToCache = {
        origins: appState.origins,
        destinations: appState.destinations,
        routes: appState.routes,
        tariffs: appState.tariffs,
        constants: appState.constants,
        excelName: appState.excelName,
        excelSize: appState.excelSize,
        excelDate: appState.excelDate
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToCache));
}

// Update Excel connection status indicator
function updateStatus(type, title, desc) {
    const indicator = excelStatus.querySelector('.status-indicator');
    const titleEl = excelStatus.querySelector('.status-title');
    const descEl = excelStatus.querySelector('.status-desc');
    
    indicator.className = `status-indicator ${type}`;
    titleEl.textContent = title;
    descEl.textContent = desc;
}

// Parse excel workbook array buffer
function parseExcelData(arrayBuffer) {
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: true, cellHTML: false });
    
    // Parse sheets
    parseConsultaFleteSheet(workbook.Sheets['CONSULTA FLETE USD']);
    parseTariffSheet(workbook.Sheets['TARIFA ENTREGA A DIRECCION USD'], 'direccion');
    parseTariffSheet(workbook.Sheets['TARIFA CONSIGNADO AGENCIA USD'], 'agencia');
}

// Parse sheet: CONSULTA FLETE USD
function parseConsultaFleteSheet(sheet) {
    if (!sheet) return;
    
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // 1. Parse route map (Columns A to G)
    // Row 0 is headers: Columna1, Columna2, Columna3, Columna4, Columna5, ESCALA, KMS
    const routes = {};
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || !row[1] || !row[3]) continue;
        const originCode = String(row[1]).trim().toUpperCase();
        const destCode = String(row[3]).trim().toUpperCase();
        const escala = row[5] ? String(row[5]).trim() : '';
        const kms = parseFloat(row[6]) || 0;
        
        routes[`${originCode}_${destCode}`] = { escala, kms };
    }
    appState.routes = routes;
    
    // 2. Parse origins (Columns I to K, starting at row 3 (index 2))
    const originsMap = new Map();
    for (let i = 2; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        const code = row[8] ? String(row[8]).trim().toUpperCase() : null;
        const name = row[9] ? String(row[9]).trim() : null;
        const reparte = row[10] ? String(row[10]).trim() : '';
        
        if (code && code !== 'O' && code !== 'ORIGEN') {
            originsMap.set(code, { code, name: name || code, reparte });
        }
    }
    appState.origins = Array.from(originsMap.values()).sort((a, b) => a.name.localeCompare(b.name));
    
    // 3. Parse destinations (Columns M to O, starting at row 3 (index 2))
    const destMap = new Map();
    for (let i = 2; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        const code = row[12] ? String(row[12]).trim().toUpperCase() : null;
        const name = row[13] ? String(row[13]).trim() : null;
        const reparte = row[14] ? String(row[14]).trim() : '';
        
        if (code && code !== 'O' && code !== 'DESTINO') {
            destMap.set(code, { code, name: name || code, reparte });
        }
    }
    appState.destinations = Array.from(destMap.values()).sort((a, b) => a.name.localeCompare(b.name));
}

// Parse tariff sheets (Direccion & Agencia)
function parseTariffSheet(sheet, type) {
    if (!sheet) return;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // Read tariffs starting from row 8 (index 7)
    const tariffsList = [];
    for (let i = 7; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;
        
        const label = row[0] ? String(row[0]).trim() : '';
        const limitVal = parseFloat(row[1]);
        
        if (!label) continue;
        
        // Special check for first row which has "Hasta  0,500" and limit might be blank
        let limit = limitVal;
        if (isNaN(limitVal)) {
            if (label.includes('0,500') || label.includes('0.500')) {
                limit = 0.5;
            } else {
                continue; // Skip invalid rows
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
    
    // Sort ascending by limit
    tariffsList.sort((a, b) => a.limit - b.limit);
    appState.tariffs[type] = tariffsList;
    
    // If it's the direction sheet, let's also read the constants
    if (type === 'direccion') {
        appState.constants.tdgBase = getCellValue(sheet, 'O23', 0.6);
        appState.constants.gcdMin = getCellValue(sheet, 'O26', 0.8);
        appState.constants.carBase = getCellValue(sheet, 'O29', 1.2);
        appState.constants.carfBase = getCellValue(sheet, 'O32', 1.8);
        appState.constants.seguroMin = getCellValue(sheet, 'O36', 1.76);
        
        // Seguro rate is typically 3.2 (meaning 3.2%).
        const rawSeguroRate = getCellValue(sheet, 'O37', 3.2);
        appState.constants.seguroRate = rawSeguroRate;
        
        appState.constants.containerMin = getCellValue(sheet, 'O40', 0.75);
    }
}

// Helper: Get cell value in sheet
function getCellValue(sheet, address, defaultValue) {
    const cell = sheet[address];
    if (!cell) return defaultValue;
    const val = parseFloat(cell.v);
    return isNaN(val) ? defaultValue : val;
}

// Populate search lists for Dropdowns
function populateDropdowns() {
    renderDropdownItems(appState.origins, originList, originInput);
    renderDropdownItems(appState.destinations, destinationList, destinationInput);
}

function renderDropdownItems(items, listEl, inputEl) {
    listEl.innerHTML = '';
    
    if (items.length === 0) {
        listEl.innerHTML = '<div class="dropdown-item no-results">No hay datos</div>';
        return;
    }
    
    items.forEach(item => {
        const div = document.createElement('div');
        div.className = 'dropdown-item';
        div.dataset.value = item.code;
        div.textContent = `${item.name} (${item.code})`;
        div.addEventListener('click', () => {
            inputEl.value = `${item.name} (${item.code})`;
            inputEl.dataset.value = item.code;
            listEl.classList.remove('show');
            inputEl.parentElement.classList.remove('open');
            
            // Trigger route analysis & calculation
            onRouteChanged();
        });
        listEl.appendChild(div);
    });
}

// Filter dropdown list based on search term
function filterDropdown(inputEl, listEl, items) {
    const filter = inputEl.value.toLowerCase();
    const childs = listEl.getElementsByClassName('dropdown-item');
    let hasResults = false;
    
    for (let i = 0; i < childs.length; i++) {
        const item = childs[i];
        if (item.classList.contains('no-results')) continue;
        
        const txtValue = item.textContent || item.innerText;
        if (txtValue.toLowerCase().indexOf(filter) > -1) {
            item.style.display = '';
            hasResults = true;
        } else {
            item.style.display = 'none';
        }
    }
    
    // Show "no results" if none match
    let noResultsEl = listEl.querySelector('.no-results');
    if (!hasResults) {
        if (!noResultsEl) {
            noResultsEl = document.createElement('div');
            noResultsEl.className = 'dropdown-item no-results';
            noResultsEl.textContent = 'Sin resultados';
            listEl.appendChild(noResultsEl);
        }
        noResultsEl.style.display = '';
    } else if (noResultsEl) {
        noResultsEl.style.display = 'none';
    }
}

// Handle Route Change (Origin/Destination)
function onRouteChanged() {
    const originCode = originInput.dataset.value;
    const destCode = destinationInput.dataset.value;
    
    if (!originCode || !destCode) {
        routeInfoCard.classList.add('hidden');
        return;
    }
    
    const routeKey = `${originCode}_${destCode}`;
    const route = appState.routes[routeKey];
    
    if (route) {
        valEscala.textContent = route.escala;
        valKms.textContent = `${route.kms} km`;
        
        // Find destination coverage
        const dest = appState.destinations.find(d => d.code === destCode);
        const coverageText = dest && dest.reparte ? dest.reparte : `ATIENDE AG. ${destCode}`;
        valAtiende.textContent = coverageText;
        
        routeInfoCard.classList.remove('hidden');
        calculateAll();
    } else {
        valEscala.textContent = '#N/A';
        valKms.textContent = '-';
        valAtiende.textContent = 'Ruta no encontrada';
        routeInfoCard.classList.remove('hidden');
        // Clear matrix prices
        clearMatrixPrices();
    }
}

// Event Listeners Setup
function setupEventListeners() {
    // Dropzone Drag & Drop
    dropzone.addEventListener('click', () => excelFileInput.click());
    
    excelFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleExcelFile(file);
    });
    
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    
    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });
    
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleExcelFile(file);
    });
    
    // Dropdowns toggling & filtering
    setupDropdownUI(originInput, originList, () => appState.origins);
    setupDropdownUI(destinationInput, destinationList, () => appState.destinations);
    
    // Calculation triggers
    const triggerInputs = [
        weightInput, bultosCount, dimLargo, dimAncho, dimAlto,
        declaredValue, exchangeRate, chkSeguro, chkCar, chkCarf,
        chkContenedor, discountPercent
    ];
    triggerInputs.forEach(input => {
        input.addEventListener('input', () => {
            if (input === weightInput || input === bultosCount) {
                syncBultosList();
            }
            if (input === exchangeRate) {
                userEditedRate = true;
            }
            calculateAll();
        });
    });
    
    // Matrix selection clicks
    Object.keys(matrixCards).forEach(opt => {
        matrixCards[opt].addEventListener('click', () => {
            Object.values(matrixCards).forEach(c => c.classList.remove('active'));
            matrixCards[opt].classList.add('active');
            appState.selectedOption = opt;
            renderDetailedReceipt();
        });
    });
    
    // Accordion toggle for bultos
    toggleBultosBtn.addEventListener('click', () => {
        bultosCard.classList.toggle('open');
        bultosContent.classList.toggle('hidden');
    });
    
    // Sync bultos button
    btnSyncBultos.addEventListener('click', () => {
        const totalW = parseFloat(weightInput.value) || 0;
        const count = parseInt(bultosCount.value) || 1;
        const avgW = totalW / count;
        
        const inputs = bultosInputsContainer.querySelectorAll('input');
        inputs.forEach(inp => {
            inp.value = avgW.toFixed(3);
        });
        
        updateBultosFromInputs();
        calculateAll();
        showToast('Pesos de bultos ecualizados correctamente');
    });
    
    // Print budget
    btnPrint.addEventListener('click', () => {
        window.print();
    });
    
    // Copy summary
    btnCopy.addEventListener('click', copySummaryToClipboard);
    
    // Fetch rate button
    if (btnFetchRate) {
        btnFetchRate.addEventListener('click', () => fetchBCVRate(false));
    }
}

// Handle local Dropdown focus and open
function setupDropdownUI(inputEl, listEl, getItemsFn) {
    const parent = inputEl.parentElement;
    const arrow = parent.querySelector('.dropdown-arrow-btn');
    
    const toggle = (e) => {
        e.stopPropagation();
        const isOpen = listEl.classList.contains('show');
        
        // Close all dropdowns first
        document.querySelectorAll('.dropdown-list').forEach(l => l.classList.remove('show'));
        document.querySelectorAll('.searchable-select').forEach(s => s.classList.remove('open'));
        
        if (!isOpen) {
            listEl.classList.add('show');
            parent.classList.add('open');
            inputEl.focus();
        }
    };
    
    inputEl.addEventListener('click', toggle);
    arrow.addEventListener('click', toggle);
    
    inputEl.addEventListener('input', () => {
        listEl.classList.add('show');
        parent.classList.add('open');
        filterDropdown(inputEl, listEl, getItemsFn());
    });
    
    // Close on click outside
    document.addEventListener('click', (e) => {
        if (!parent.contains(e.target)) {
            listEl.classList.remove('show');
            parent.classList.remove('open');
        }
    });
}

// Handle Uploaded File
function handleExcelFile(file) {
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xlsm')) {
        showToast('Formato de archivo inválido. Utilice .xlsx o .xlsm', 'error');
        return;
    }
    
    updateStatus('loading', 'Cargando Excel subido...', file.name);
    
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const arrayBuffer = e.target.result;
            appState.excelName = file.name;
            appState.excelSize = file.size;
            appState.excelDate = new Date().toLocaleDateString('es-VE');
            
            parseExcelData(arrayBuffer);
            saveCache();
            
            updateStatus('ready', 'Excel Conectado (Manual)', `${file.name} (${(file.size/1024/1024).toFixed(2)} MB)`);
            populateDropdowns();
            
            // If the route was already selected, update route info
            onRouteChanged();
            
            showToast('Tarifas actualizadas correctamente desde el Excel subido');
        } catch (err) {
            console.error('Error parsing uploaded Excel:', err);
            updateStatus('error', 'Error al procesar Excel', 'El archivo no tiene el formato esperado');
            showToast('Error al parsear el archivo Excel. Verifica que sea la plantilla correcta.', 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Sync the detailed bultos list inputs with the bultos count
function syncBultosList() {
    const count = Math.max(1, parseInt(bultosCount.value) || 1);
    const totalW = Math.max(0, parseFloat(weightInput.value) || 0);
    const avgW = totalW / count;
    
    // Save existing inputs values if possible
    const currentValues = [];
    const inputs = bultosInputsContainer.querySelectorAll('input');
    inputs.forEach(inp => currentValues.push(parseFloat(inp.value) || 0));
    
    bultosInputsContainer.innerHTML = '';
    
    for (let i = 0; i < count; i++) {
        const item = document.createElement('div');
        item.className = 'bulto-input-item';
        
        const span = document.createElement('span');
        span.textContent = `Bulto ${i + 1}`;
        
        const input = document.createElement('input');
        input.type = 'number';
        input.min = '0';
        input.step = '0.001';
        
        // Restore value if available, else assign average
        let val = avgW;
        if (i < currentValues.length && currentValues.length === count) {
            val = currentValues[i];
        }
        input.value = val.toFixed(3);
        
        // Listen to changes in individual inputs
        input.addEventListener('input', () => {
            updateBultosFromInputs();
            updateTotalWeightFromBultos();
            calculateAll();
        });
        
        item.appendChild(span);
        item.appendChild(input);
        bultosInputsContainer.appendChild(item);
    }
    
    updateBultosFromInputs();
}

// Read weights from bultos inputs and populate appState
function updateBultosFromInputs() {
    const inputs = bultosInputsContainer.querySelectorAll('input');
    const weights = [];
    inputs.forEach(inp => {
        weights.push(Math.max(0, parseFloat(inp.value) || 0));
    });
    appState.bultos = weights;
}

// Sum bultos and set physical weight
function updateTotalWeightFromBultos() {
    const sum = appState.bultos.reduce((a, b) => a + b, 0);
    weightInput.value = sum.toFixed(3);
}

// Global calculations for all 4 scenarios
function calculateAll() {
    const originCode = originInput.dataset.value;
    const destCode = destinationInput.dataset.value;
    
    if (!originCode || !destCode) {
        clearMatrixPrices();
        return;
    }
    
    const routeKey = `${originCode}_${destCode}`;
    const route = appState.routes[routeKey];
    if (!route || route.escala === '#N/A' || !route.escala) {
        clearMatrixPrices();
        return;
    }
    
    // Core parameters
    const physWeight = parseFloat(weightInput.value) || 0;
    const bCount = Math.max(1, parseInt(bultosCount.value) || 1);
    const exRate = parseFloat(exchangeRate.value) || 36.50;
    const decVal = parseFloat(declaredValue.value) || 0;
    const discPercent = parseFloat(discountPercent.value) || 0;
    
    // Volumetric weight
    const largo = parseFloat(dimLargo.value) || 0;
    const ancho = parseFloat(dimAncho.value) || 0;
    const alto = parseFloat(dimAlto.value) || 0;
    const volWeight = (largo * ancho * alto) / 5000;
    
    valCubaje.textContent = volWeight.toFixed(2);
    
    // Max weight
    const finalWeight = Math.max(physWeight, volWeight);
    valPesoFinal.textContent = finalWeight.toFixed(3);
    
    // Ensure bultos list is populated
    if (appState.bultos.length !== bCount) {
        syncBultosList();
    }
    
    // Calculate for each of the 4 combinations
    const results = {
        origDir: runScenario('direccion', false, finalWeight, route.escala, decVal, discPercent, bCount, exRate),
        origAge: runScenario('agencia', false, finalWeight, route.escala, decVal, discPercent, bCount, exRate),
        destDir: runScenario('direccion', true, finalWeight, route.escala, decVal, discPercent, bCount, exRate),
        destAge: runScenario('agencia', true, finalWeight, route.escala, decVal, discPercent, bCount, exRate)
    };
    
    // Store calculations in state for receipt rendering
    appState.calculations = results;
    
    // Render matrix card summaries
    Object.keys(results).forEach(opt => {
        const res = results[opt];
        const usdEl = document.getElementById(`matrixUSD_${opt}`);
        const bsEl = document.getElementById(`matrixBs_${opt}`);
        
        usdEl.textContent = `$${res.totalUSD.toFixed(2)}`;
        bsEl.textContent = `Bs. ${res.totalBs.toFixed(2)}`;
    });
    
    // Render active receipt details
    renderDetailedReceipt();
}

// Single scenario calculation formula engine (exactly mirroring Excel)
function runScenario(tariffType, isCOD, weight, escala, declaredVal, discPercent, count, exRate) {
    const list = appState.tariffs[tariffType];
    
    // 1. Base Freight Lookup
    let baseFreight = 0;
    if (list && list.length > 0) {
        // Find row where limit >= weight
        const matchedRow = list.find(r => r.limit >= weight);
        if (matchedRow) {
            baseFreight = matchedRow[escala] || 0;
        } else {
            // Cap at the maximum weight class
            baseFreight = list[list.length - 1][escala] || 0;
        }
    }
    
    // 2. Discount
    const hasDiscount = discPercent > 0;
    const discountVal = hasDiscount ? baseFreight * (discPercent / 100) : 0;
    const netFreight = Math.ceil((baseFreight - discountVal) * 1000) / 1000; // ROUNDUP 3 decs
    
    // 3. TDG (Tramite de guia): if weight > 160kg, apply tdgBase
    const applyTdg = weight > 160;
    const tdgVal = applyTdg ? appState.constants.tdgBase : 0;
    
    // 4. GCD (Cobro a destino): if COD, apply 10% of net flete, min gcdMin
    let gcdVal = 0;
    if (isCOD) {
        gcdVal = Math.max(netFreight * 0.1, appState.constants.gcdMin);
    }
    
    // 5. CAR / CARF
    const carVal = chkCar.checked ? appState.constants.carBase : 0;
    const carfVal = chkCarf.checked ? appState.constants.carfBase : 0;
    
    // 6. Seguro (Insurance)
    let seguroVal = 0;
    if (chkSeguro.checked && declaredVal > 0) {
        if (declaredVal <= 55) {
            seguroVal = appState.constants.seguroMin; // 1.76
        } else {
            seguroVal = declaredVal * (appState.constants.seguroRate / 100); // 3.2%
        }
        seguroVal = Math.ceil(seguroVal * 1000) / 1000; // ROUNDUP 3 decs
    }
    
    // 7. Contenedor: 10% of Base Freight
    const containerVal = chkContenedor.checked ? (baseFreight * 0.1) : 0;
    
    // 8. Subtotal (Base Imponible)
    const subtotal = netFreight + tdgVal + gcdVal + carVal + carfVal + seguroVal + containerVal;
    
    // 9. IVA (16%)
    // Excel: TRUNC(subtotal * 0.16, 3)
    const ivaVal = Math.floor((subtotal * 0.16) * 1000) / 1000;
    
    // 10. Franqueo Postal (IPOSTEL 10%)
    // Only applies to individual bultos <= 30 kg
    let franqueoVal = 0;
    if (appState.bultos && appState.bultos.length > 0) {
        // Total weight sum for proportional weight sharing
        const sumBultosWeight = appState.bultos.reduce((a, b) => a + b, 0) || 1;
        
        let sumFranqueo = 0;
        appState.bultos.forEach(bw => {
            if (bw <= 30) {
                // Proportional Base Freight portion for this package
                const propBaseFreight = (bw / sumBultosWeight) * baseFreight;
                // 10% of that portion
                sumFranqueo += propBaseFreight * 0.1;
            }
        });
        franqueoVal = Math.ceil(sumFranqueo * 1000) / 1000; // ROUNDUP 3 decs
    }
    
    // 11. IGTF (3%)
    // Base is: Subtotal + IVA + Franqueo
    const igtfBase = subtotal + ivaVal + franqueoVal;
    const igtfVal = igtfBase * 0.03;
    
    // 12. Total USD
    const totalUSD = igtfBase + igtfVal;
    
    // 13. Total Bs
    // El IGTF solo aplica para pagos en divisas, por lo que el total en bolívares lo excluye.
    const totalBs = igtfBase * exRate;
    
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
        totalUSD,
        totalBs,
        exRate,
        weight,
        escala
    };
}

// Clear prices from the matrix and receipt
function clearMatrixPrices() {
    routeInfoCard.classList.add('hidden');
    Object.keys(matrixCards).forEach(opt => {
        document.getElementById(`matrixUSD_${opt}`).textContent = '$0.00';
        document.getElementById(`matrixBs_${opt}`).textContent = 'Bs. 0.00';
    });
    
    // Clear receipt
    document.getElementById('itemBaseUSD').textContent = '$0.00';
    document.getElementById('itemSubtotalUSD').textContent = '$0.00';
    document.getElementById('itemIvaUSD').textContent = '$0.00';
    document.getElementById('itemFranqueoUSD').textContent = '$0.00';
    document.getElementById('itemIgtfUSD').textContent = '$0.00';
    document.getElementById('itemTotalUSD').textContent = '$0.00';
    document.getElementById('itemTotalBs').textContent = 'Bs. 0.00';
    document.getElementById('receiptRoute').textContent = 'Ruta: -';
}

// Render detailed invoice breakdown for selected option
function renderDetailedReceipt() {
    const opt = appState.selectedOption;
    if (!appState.calculations || !appState.calculations[opt]) return;
    
    const res = appState.calculations[opt];
    
    // 1. Meta Titles & Badges
    const badge = document.getElementById('receiptBadge');
    const title = document.getElementById('receiptTitle');
    const route = document.getElementById('receiptRoute');
    const date = document.getElementById('receiptDate');
    
    date.textContent = `Fecha: ${appState.excelDate || new Date().toLocaleDateString('es-VE')}`;
    
    const originName = originInput.value.split('(')[0].trim();
    const destName = destinationInput.value.split('(')[0].trim();
    route.textContent = `Ruta: ${originName} → ${destName} (${res.escala})`;
    
    if (opt.startsWith('orig')) {
        badge.textContent = 'Pago en Origen';
        badge.className = 'receipt-badge tag-origen';
    } else {
        badge.textContent = 'Cobro en Destino';
        badge.className = 'receipt-badge tag-destino';
    }
    
    if (opt.endsWith('Dir')) {
        title.textContent = 'Presupuesto: Entrega a Dirección';
    } else {
        title.textContent = 'Presupuesto: Retiro en Agencia';
    }
    
    // 2. Line Items
    document.getElementById('itemBaseUSD').textContent = `$${res.baseFreight.toFixed(2)}`;
    
    // Discount
    const discRow = document.getElementById('discountRow');
    const netRow = document.getElementById('fleteNetoRow');
    if (res.discountVal > 0) {
        discRow.classList.remove('hidden');
        netRow.classList.remove('hidden');
        document.getElementById('itemDiscountPercent').textContent = discountPercent.value;
        document.getElementById('itemDiscountUSD').textContent = `-$${res.discountVal.toFixed(2)}`;
        document.getElementById('itemNetUSD').textContent = `$${res.netFreight.toFixed(2)}`;
    } else {
        discRow.classList.add('hidden');
        netRow.classList.add('hidden');
    }
    
    // Row toggles for secondary charges
    toggleReceiptRow('itemTdgRow', 'itemTdgUSD', res.tdgVal);
    toggleReceiptRow('itemGcdRow', 'itemGcdUSD', res.gcdVal);
    toggleReceiptRow('itemCarRow', 'itemCarUSD', res.carVal);
    toggleReceiptRow('itemCarfRow', 'itemCarfUSD', res.carfVal);
    toggleReceiptRow('itemSeguroRow', 'itemSeguroUSD', res.seguroVal);
    toggleReceiptRow('itemContenedorRow', 'itemContenedorUSD', res.containerVal);
    
    // 3. Taxes & Totals
    document.getElementById('itemSubtotalUSD').textContent = `$${res.subtotal.toFixed(2)}`;
    document.getElementById('itemIvaUSD').textContent = `$${res.ivaVal.toFixed(2)}`;
    document.getElementById('itemFranqueoUSD').textContent = `$${res.franqueoVal.toFixed(2)}`;
    document.getElementById('itemIgtfUSD').textContent = `$${res.igtfVal.toFixed(2)}`;
    
    document.getElementById('itemTotalUSD').textContent = `$${res.totalUSD.toFixed(2)}`;
    document.getElementById('itemTotalBs').textContent = `Bs. ${res.totalBs.toFixed(2)}`;
}

// Toggle row visibility in receipt
function toggleReceiptRow(rowId, valId, value) {
    const row = document.getElementById(rowId);
    const valEl = document.getElementById(valId);
    
    if (value > 0) {
        row.classList.remove('hidden');
        valEl.textContent = `$${value.toFixed(2)}`;
    } else {
        row.classList.add('hidden');
    }
}

// Copy Text Summary to Clipboard
function copySummaryToClipboard() {
    const opt = appState.selectedOption;
    if (!appState.calculations || !appState.calculations[opt]) return;
    
    const res = appState.calculations[opt];
    const originName = originInput.value;
    const destName = destinationInput.value;
    
    const scenarioName = opt === 'origDir' ? 'Pago en Origen - Entrega a Dirección' :
                        opt === 'origAge' ? 'Pago en Origen - Retiro en Agencia' :
                        opt === 'destDir' ? 'Cobro en Destino - Entrega a Dirección' :
                                            'Cobro en Destino - Retiro en Agencia';
                                            
    let text = `=== PRESUPUESTO DE FLETE ===\n`;
    text += `Escenario: ${scenarioName}\n`;
    text += `Ruta: ${originName} → ${destName}\n`;
    text += `Escala: ${res.escala} | Peso Cálculo: ${res.weight.toFixed(3)} kg\n`;
    text += `-----------------------------------\n`;
    text += `Flete Base: $${res.baseFreight.toFixed(2)}\n`;
    if (res.discountVal > 0) {
        text += `Descuento (${discountPercent.value}%): -$${res.discountVal.toFixed(2)}\n`;
        text += `Flete Neto: $${res.netFreight.toFixed(2)}\n`;
    }
    if (res.tdgVal > 0) text += `Trámite Guía (TDG): $${res.tdgVal.toFixed(2)}\n`;
    if (res.gcdVal > 0) text += `Cobro a Destino (GCD): $${res.gcdVal.toFixed(2)}\n`;
    if (res.carVal > 0) text += `Acuse de Recibo (CAR): $${res.carVal.toFixed(2)}\n`;
    if (res.carfVal > 0) text += `Acuse con Factura (CAR F): $${res.carfVal.toFixed(2)}\n`;
    if (res.seguroVal > 0) text += `Seguro de Mercancía: $${res.seguroVal.toFixed(2)}\n`;
    if (res.containerVal > 0) text += `Contenedor (10%): $${res.containerVal.toFixed(2)}\n`;
    text += `Subtotal Base: $${res.subtotal.toFixed(2)}\n`;
    text += `IVA (16%): $${res.ivaVal.toFixed(2)}\n`;
    text += `Franqueo Postal (10%): $${res.franqueoVal.toFixed(2)}\n`;
    text += `IGTF (3%): $${res.igtfVal.toFixed(2)}\n`;
    text += `-----------------------------------\n`;
    text += `TOTAL ($): $${res.totalUSD.toFixed(2)}\n`;
    text += `TOTAL (Bs.): Bs. ${res.totalBs.toFixed(2)}\n`;
    text += `Tasa de cambio: Bs. ${res.exRate.toFixed(2)}/$\n`;
    text += `Generado por PresupuestoApp`;
    
    navigator.clipboard.writeText(text).then(() => {
        showToast('Resumen copiado al portapapeles');
    }).catch(err => {
        console.error('Failed to copy summary:', err);
        showToast('Error al copiar el resumen', 'error');
    });
}

// Helper: Show Toast Notification
function showToast(message, type = 'success') {
    toast.textContent = message;
    toast.className = `toast ${type === 'error' ? 'error' : 'success'}`;
    toast.classList.remove('hidden');
    
    setTimeout(() => {
        toast.classList.add('hidden');
    }, 3500);
}
