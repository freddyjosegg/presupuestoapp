// Configuration and Constants
const DEFAULT_EXCEL_PATH = 'TARIFA ACTUAL PAQUETERIA USD 210526.xlsm';
const STORAGE_KEY = 'presupuestoapp_db';
const BCV_API_URL = 'https://ve.dolarapi.com/v1/dolares/oficial';

// Utility: Format Bs. with point as thousands separator and comma for decimals
function roundHalfUp(val, decimals = 2) {
    if (val === undefined || val === null || isNaN(val)) {
        return (0).toFixed(decimals);
    }
    const factor = Math.pow(10, decimals);
    return (Math.round((val + 1e-9) * factor) / factor).toFixed(decimals);
}

function roundHalfUpNum(val, decimals = 2) {
    if (val === undefined || val === null || isNaN(val)) {
        return 0;
    }
    const factor = Math.pow(10, decimals);
    return Math.round((val + 1e-9) * factor) / factor;
}

function truncDec(val, decimals = 3) {
    if (val === undefined || val === null || isNaN(val)) {
        return 0;
    }
    const factor = Math.pow(10, decimals);
    return Math.floor((val + 1e-9) * factor) / factor;
}

function roundUpDec(val, decimals = 3) {
    if (val === undefined || val === null || isNaN(val)) {
        return 0;
    }
    const factor = Math.pow(10, decimals);
    return Math.ceil((val - 1e-9) * factor) / factor;
}


// Utility: Format Bs. with point as thousands separator and comma for decimals
function formatBs(value) {
    if (value === undefined || value === null || isNaN(value)) {
        return '0,00';
    }
    const parts = roundHalfUp(value, 2).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return parts.join(',');
}

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
// IndexedDB Helper Functions
const DB_NAME = 'PresupuestoAppDB';
const DB_VERSION = 1;
const STORE_NAME = 'excel_store';

function openDB() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = (e) => {
            const db = e.target.result;
            if (!db.objectStoreNames.contains(STORE_NAME)) {
                db.createObjectStore(STORE_NAME);
            }
        };
        request.onsuccess = (e) => resolve(e.target.result);
        request.onerror = (e) => reject(e.target.error);
    });
}

async function getDBItem(key) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE_NAME, 'readonly');
        const store = transaction.objectStore(STORE_NAME);
        const request = store.get(key);
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function setDBItem(key, value) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE_NAME, 'readwrite');
        const store = transaction.objectStore(STORE_NAME);
        const request = store.put(value, key);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error);
    });
}

// Register Service Worker for PWA
function registerServiceWorker() {
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', () => {
            navigator.serviceWorker.register('sw.js')
                .then(reg => console.log('Service Worker registrado con éxito:', reg.scope))
                .catch(err => console.error('Error al registrar Service Worker:', err));
        });
    }
}

// PWA custom install prompt banner logic
let deferredPrompt;
function setupPwaInstall() {
    const installBtn = document.getElementById('pwaInstallBtn');
    if (!installBtn) return;
    
    window.addEventListener('beforeinstallprompt', (e) => {
        e.preventDefault();
        deferredPrompt = e;
        installBtn.classList.remove('hidden');
        
        installBtn.addEventListener('click', () => {
            installBtn.classList.add('hidden');
            deferredPrompt.prompt();
            deferredPrompt.userChoice.then((choiceResult) => {
                if (choiceResult.outcome === 'accepted') {
                    console.log('El usuario aceptó la instalación de PresupuestoApp');
                } else {
                    console.log('El usuario rechazó la instalación de PresupuestoApp');
                }
                deferredPrompt = null;
            });
        });
    });
    
    window.addEventListener('appinstalled', (evt) => {
        console.log('PresupuestoApp instalada con éxito en el sistema.');
        installBtn.classList.add('hidden');
        showToast('¡PresupuestoApp instalada con éxito!');
    });
}

// Initialize Application
window.addEventListener('DOMContentLoaded', () => {
    initApp();
});

async function initApp() {
    setupEventListeners();
    setupPwaInstall();
    registerServiceWorker();
    
    // Fetch BCV Rate at startup
    fetchBCVRate(true);
    
    // 1. Try loading from IndexedDB
    try {
        const cachedState = await getDBItem('app_state');
        if (cachedState) {
            appState = { ...appState, ...cachedState };
            updateStatus('ready', `Cargado de caché local`, `${appState.excelName || 'Excel guardado'}`);
            populateDropdowns();
            calculateAll();
            // Fetch file headers to check for updates silently
            checkForExcelUpdates();
        } else {
            // First run: fetch full excel
            fetchDefaultExcel();
        }
    } catch (e) {
        console.error('Error loading from IndexedDB, fetching new one...', e);
        fetchDefaultExcel();
    }
}

// Check if Excel has changed on the server using a HEAD request (Content-Length and Last-Modified)
async function checkForExcelUpdates() {
    try {
        const response = await fetch(DEFAULT_EXCEL_PATH, { method: 'HEAD' });
        if (!response.ok) return;
        
        const serverSize = parseInt(response.headers.get('Content-Length')) || 0;
        const serverLastModified = response.headers.get('Last-Modified') || '';
        
        const cachedMeta = await getDBItem('excel_metadata');
        
        // If file differs from local version, download and reprocess
        if (!cachedMeta || cachedMeta.size !== serverSize || cachedMeta.lastModified !== serverLastModified) {
            console.log('Detectado nuevo archivo Excel en el servidor. Actualizando...');
            fetchDefaultExcel(true); // silent fetch in background
        } else {
            console.log('El archivo Excel local está al día con el servidor.');
        }
    } catch (error) {
        console.warn('No se pudo verificar actualizaciones del Excel (modo offline/red caída).', error);
    }
}

// Fetch BCV Exchange Rate from ve.dolarapi.com
async function fetchBCVRate(silent = false) {
    if (btnFetchRate) {
        btnFetchRate.classList.add('loading');
    }
    
    try {
        const response = await fetch(`${BCV_API_URL}?t=${Date.now()}`);
        if (!response.ok) throw new Error('Network response was not ok');
        
        const data = await response.json();
        const rate = parseFloat(data.promedio);
        
        if (!isNaN(rate) && rate > 0) {
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
        
        const serverSize = parseInt(response.headers.get('Content-Length')) || 0;
        const serverLastModified = response.headers.get('Last-Modified') || '';
        
        const blob = await response.blob();
        const arrayBuffer = await blob.arrayBuffer();
        
        appState.excelName = DEFAULT_EXCEL_PATH;
        appState.excelSize = blob.size;
        appState.excelDate = new Date().toLocaleDateString('es-VE');
        
        const parsedData = await processExcelWithWorker(arrayBuffer);
        loadParsedData(parsedData);
        
        // Save to IndexedDB database
        await setDBItem('app_state', {
            origins: appState.origins,
            destinations: appState.destinations,
            routes: appState.routes,
            tariffs: appState.tariffs,
            constants: appState.constants,
            excelName: appState.excelName,
            excelSize: appState.excelSize,
            excelDate: appState.excelDate
        });
        
        await setDBItem('excel_metadata', {
            name: DEFAULT_EXCEL_PATH,
            size: serverSize || blob.size,
            lastModified: serverLastModified || new Date().toUTCString(),
            date: appState.excelDate
        });
        
        updateStatus('ready', 'Excel conectado', `${appState.excelName} (${(appState.excelSize/1024/1024).toFixed(2)} MB)`);
        populateDropdowns();
        calculateAll();
        if (!silent) {
            showToast('Base de datos Excel cargada correctamente');
        }
    } catch (error) {
        console.error('Error loading default Excel:', error);
        if (!silent) {
            updateStatus('error', 'Error al cargar Excel', 'Arrastra el archivo Excel manualmente');
            showToast('No se pudo cargar el Excel por defecto. Por favor arrástralo manualmente.', 'error');
        }
    }
}

// Update Excel connection status indicator
function updateStatus(type, title, desc) {
    if (!excelStatus) return;
    const indicator = excelStatus.querySelector('.status-indicator');
    const titleEl = excelStatus.querySelector('.status-title');
    const descEl = excelStatus.querySelector('.status-desc');
    
    if (indicator) indicator.className = `status-indicator ${type}`;
    if (titleEl) titleEl.textContent = title;
    if (descEl) descEl.textContent = desc;
}

// Process excel workbook using Web Worker in a background thread to prevent UI freezing
function processExcelWithWorker(arrayBuffer) {
    return new Promise((resolve, reject) => {
        const worker = new Worker('worker.js');
        worker.onmessage = function (e) {
            const { success, data, error } = e.data;
            worker.terminate();
            if (success) {
                resolve(data);
            } else {
                reject(new Error(error));
            }
        };
        worker.onerror = function (err) {
            worker.terminate();
            reject(err);
        };
        // Use transferable object to avoid copying memory buffer
        worker.postMessage({ arrayBuffer }, [arrayBuffer]);
    });
}

// Load parsed data from worker into application state
function loadParsedData(data) {
    appState.origins = data.origins;
    appState.destinations = data.destinations;
    appState.routes = data.routes;
    appState.tariffs = data.tariffs;
    appState.constants = data.constants;
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
    // Excel Status Card acts as Drag & Drop and Click selector
    excelStatus.addEventListener('click', () => excelFileInput.click());
    
    excelFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleExcelFile(file);
    });
    
    excelStatus.addEventListener('dragover', (e) => {
        e.preventDefault();
        excelStatus.classList.add('dragover');
    });
    
    excelStatus.addEventListener('dragleave', () => {
        excelStatus.classList.remove('dragover');
    });
    
    excelStatus.addEventListener('drop', (e) => {
        e.preventDefault();
        excelStatus.classList.remove('dragover');
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
    
    weightInput.addEventListener('blur', () => {
        const val = parseFloat(weightInput.value) || 0;
        weightInput.value = val.toFixed(3);
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
    reader.onload = async (e) => {
        try {
            const arrayBuffer = e.target.result;
            appState.excelName = file.name;
            appState.excelSize = file.size;
            appState.excelDate = new Date().toLocaleDateString('es-VE');
            
            // Use background Web Worker to parse
            const parsedData = await processExcelWithWorker(arrayBuffer);
            loadParsedData(parsedData);
            // Save to IndexedDB
            await setDBItem('app_state', {
                origins: appState.origins,
                destinations: appState.destinations,
                routes: appState.routes,
                tariffs: appState.tariffs,
                constants: appState.constants,
                excelName: appState.excelName,
                excelSize: appState.excelSize,
                excelDate: appState.excelDate
            });
            
            await setDBItem('excel_metadata', {
                name: file.name,
                size: file.size,
                lastModified: new Date(file.lastModified).toUTCString(),
                date: appState.excelDate
            });
            
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
        input.value = avgW.toFixed(3);
        
        // Listen to changes in individual inputs
        input.addEventListener('input', () => {
            updateBultosFromInputs();
            updateTotalWeightFromBultos();
            calculateAll();
        });
        
        input.addEventListener('blur', () => {
            const v = parseFloat(input.value) || 0;
            input.value = v.toFixed(3);
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
    
    valCubaje.textContent = volWeight.toFixed(3);
    
    // Max weight
    const finalWeight = Math.max(physWeight, volWeight);
    valPesoFinal.textContent = finalWeight.toFixed(3);
    
    // Si no hay peso ingresado (peso final es 0 o menor), limpiar los precios pero mantener la tarjeta de ruta visible
    if (finalWeight <= 0) {
        clearMatrixPrices(false);
        return;
    }
    
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
        
        usdEl.innerHTML = `$${roundHalfUp(res.totalUSDSinIgtf, 2)}<span class="igtf-subtext">($${roundHalfUp(res.totalUSD, 2)} con IGTF)</span>`;
        bsEl.textContent = `Bs. ${formatBs(res.totalBs)}`;
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
    const netFreight = roundUpDec(baseFreight - discountVal, 3); // ROUNDUP 3 decs
    
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
        seguroVal = roundUpDec(seguroVal, 3); // ROUNDUP 3 decs
    }
    
    // 7. Contenedor: 10% of Base Freight
    const containerVal = chkContenedor.checked ? (baseFreight * 0.1) : 0;
    
    // 8. Subtotal (Base Imponible)
    const subtotal = netFreight + tdgVal + gcdVal + carVal + carfVal + seguroVal + containerVal;
    
    // 9. IVA (16%)
    // Excel: TRUNC(subtotal * 0.16, 3)
    const ivaVal = truncDec(subtotal * 0.16, 3);
    
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
        franqueoVal = roundUpDec(sumFranqueo, 3); // ROUNDUP 3 decs
    }
    
    // 11. IGTF (3%)
    // Base is: Subtotal + IVA + Franqueo
    // Redondeamos la base en USD a 2 decimales para que coincida con el total en USD visualizado por el usuario
    const rawIgtfBase = subtotal + ivaVal + franqueoVal;
    const igtfBase = roundHalfUpNum(rawIgtfBase, 2);
    
    const igtfVal = igtfBase * 0.03;
    
    // 12. Total USD
    const totalUSD = igtfBase + igtfVal;
    
    // 13. Total Bs
    // El IGTF solo aplica para pagos en divisas, por lo que el total en bolívares lo excluye.
    // Usamos el igtfBase ya redondeado (que representa el total en USD sin IGTF)
    const totalBs = truncDec(igtfBase * exRate, 2);
    
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
        totalBs,
        exRate,
        weight,
        escala
    };
}

// Clear prices from the matrix and receipt
function clearMatrixPrices(hideRouteCard = true) {
    if (hideRouteCard) {
        routeInfoCard.classList.add('hidden');
    }
    Object.keys(matrixCards).forEach(opt => {
        document.getElementById(`matrixUSD_${opt}`).innerHTML = `$0.00<span class="igtf-subtext">($0.00 con IGTF)</span>`;
        document.getElementById(`matrixBs_${opt}`).textContent = 'Bs. 0,00';
    });
    
    // Clear receipt
    document.getElementById('itemBaseUSD').textContent = '$0.00';
    document.getElementById('itemSubtotalUSD').textContent = '$0.00';
    document.getElementById('itemIvaUSD').textContent = '$0.00';
    document.getElementById('itemFranqueoUSD').textContent = '$0.00';
    document.getElementById('itemFranqueoRow').classList.remove('hidden');
    document.getElementById('itemIgtfUSD').textContent = '$0.00';
    document.getElementById('itemTotalUSD').innerHTML = `$0.00<span class="igtf-subtext">($0.00 con IGTF)</span>`;
    document.getElementById('itemTotalBs').textContent = 'Bs. 0,00';
    document.getElementById('receiptRoute').textContent = '-';
    document.getElementById('receiptTitle').textContent = '-';
    document.getElementById('receiptBadge').textContent = '-';
    document.getElementById('printNoteDate').textContent = '-';
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
    route.textContent = `${originName} → ${destName} (${res.escala})`;
    
    if (opt.startsWith('orig')) {
        badge.innerHTML = '<span class="text-origen">Pago en Origen</span>';
    } else {
        badge.innerHTML = '<span class="text-destino">Cobro en Destino</span>';
    }
    
    if (opt.endsWith('Dir')) {
        title.textContent = 'Entrega a Dirección';
    } else {
        title.textContent = 'Retiro en Agencia';
    }
    
    // 2. Line Items
    document.getElementById('itemBaseUSD').textContent = `$${roundHalfUp(res.baseFreight, 2)}`;
    
    // Discount
    const discRow = document.getElementById('discountRow');
    const netRow = document.getElementById('fleteNetoRow');
    if (res.discountVal > 0) {
        discRow.classList.remove('hidden');
        netRow.classList.remove('hidden');
        document.getElementById('itemDiscountPercent').textContent = discountPercent.value;
        document.getElementById('itemDiscountUSD').textContent = `-$${roundHalfUp(res.discountVal, 2)}`;
        document.getElementById('itemNetUSD').textContent = `$${roundHalfUp(res.netFreight, 2)}`;
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
    document.getElementById('itemSubtotalUSD').textContent = `$${roundHalfUp(res.subtotal, 2)}`;
    document.getElementById('itemIvaUSD').textContent = `$${roundHalfUp(res.ivaVal, 2)}`;
    toggleReceiptRow('itemFranqueoRow', 'itemFranqueoUSD', res.franqueoVal);
    document.getElementById('itemIgtfUSD').textContent = `$${roundHalfUp(res.igtfVal, 2)}`;
    
    document.getElementById('itemTotalUSD').innerHTML = `$${roundHalfUp(res.totalUSDSinIgtf, 2)}<span class="igtf-subtext">($${roundHalfUp(res.totalUSD, 2)} con IGTF)</span>`;
    document.getElementById('itemTotalBs').textContent = `Bs. ${formatBs(res.totalBs)}`;
    
    // Set legal note date
    const printDate = appState.excelDate || new Date().toLocaleDateString('es-VE');
    document.getElementById('printNoteDate').textContent = printDate;
}

// Toggle row visibility in receipt
function toggleReceiptRow(rowId, valId, value) {
    const row = document.getElementById(rowId);
    const valEl = document.getElementById(valId);
    
    if (value > 0) {
        row.classList.remove('hidden');
        valEl.textContent = `$${roundHalfUp(value, 2)}`;
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
    text += `Flete Base: $${roundHalfUp(res.baseFreight, 2)}\n`;
    if (res.discountVal > 0) {
        text += `Descuento (${discountPercent.value}%): -$${roundHalfUp(res.discountVal, 2)}\n`;
        text += `Flete Neto: $${roundHalfUp(res.netFreight, 2)}\n`;
    }
    if (res.tdgVal > 0) text += `Trámite Guía (TDG): $${roundHalfUp(res.tdgVal, 2)}\n`;
    if (res.gcdVal > 0) text += `Cobro a Destino (GCD): $${roundHalfUp(res.gcdVal, 2)}\n`;
    if (res.carVal > 0) text += `Acuse de Recibo (CAR): $${roundHalfUp(res.carVal, 2)}\n`;
    if (res.carfVal > 0) text += `Acuse con Factura (CAR F): $${roundHalfUp(res.carfVal, 2)}\n`;
    if (res.seguroVal > 0) text += `Seguro de Mercancía: $${roundHalfUp(res.seguroVal, 2)}\n`;
    if (res.containerVal > 0) text += `Contenedor (10%): $${roundHalfUp(res.containerVal, 2)}\n`;
    text += `Subtotal Base: $${roundHalfUp(res.subtotal, 2)}\n`;
    text += `IVA (16%): $${roundHalfUp(res.ivaVal, 2)}\n`;
    text += `Franqueo Postal (10%): $${roundHalfUp(res.franqueoVal, 2)}\n`;
    text += `IGTF (3%): $${roundHalfUp(res.igtfVal, 2)}\n`;
    text += `-----------------------------------\n`;
    text += `TOTAL ($): $${roundHalfUp(res.totalUSDSinIgtf, 2)} ($${roundHalfUp(res.totalUSD, 2)} con IGTF)\n`;
    text += `TOTAL (Bs.): Bs. ${formatBs(res.totalBs)}\n`;
    text += `Tasa de cambio: Bs. ${formatBs(res.exRate)}/$\n`;
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
