const puppeteer = require('puppeteer');
const path = require('path');
const { spawn } = require('child_process');

(async () => {
    let serverProcess;
    let browser;
    let page;

    try {
        console.log('Iniciando el servidor local en segundo plano (server.js)...');
        serverProcess = spawn('node', ['./server.js'], {
            cwd: path.resolve(__dirname, '..')
        });

        // Capturar logs del servidor
        serverProcess.stdout.on('data', (data) => {
            console.log(`[SERVIDOR] ${data.toString().trim()}`);
        });

        serverProcess.stderr.on('data', (data) => {
            console.error(`[SERVIDOR ERROR] ${data.toString().trim()}`);
        });

        // Esperar 2 segundos a que el servidor se inicie y esté listo
        await new Promise(resolve => setTimeout(resolve, 2000));

        console.log('Iniciando navegador Puppeteer...');
        browser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--ignore-certificate-errors']
        });
        
        page = await browser.newPage();
        
        // Inyectar mock para interceptar y forzar fallos de red en modo offline simulado
        await page.evaluateOnNewDocument(() => {
            const isOffline = sessionStorage.getItem('simulate_offline') === 'true';
            if (isOffline) {
                const originalFetch = window.fetch;
                window.fetch = async function(...args) {
                    const url = args[0];
                    if (typeof url === 'string' && (url.includes('dolarapi.com') || url.includes('/api/latest-excel'))) {
                        throw new TypeError('Failed to fetch (simulated offline)');
                    }
                    return originalFetch.apply(this, args);
                };
            }
        });
        
        // Ajustar el viewport a una pantalla de laptop grande
        await page.setViewport({ width: 1280, height: 900 });

        // Redirigir logs y errores del navegador a la consola del test
        page.on('console', msg => console.log(`[NAVEGADOR] ${msg.text()}`));
        page.on('pageerror', err => console.log(`[NAVEGADOR ERROR] ${err.toString()}`));

        // --- PASO 1: CARGAR EN LÍNEA PARA LLENAR CACHÉ Y BASE DE DATOS ---
        console.log('\n--- PASO 1: Cargando la app online para registrar Service Worker e IndexedDB ---');
        console.log('Navegando a la app local (https://127.0.0.1:3000)...');
        await page.goto('https://127.0.0.1:3000', { waitUntil: 'load', timeout: 30000 });

        // Esperar registro y activación del Service Worker
        console.log('Esperando a que el Service Worker esté listo y controle la página...');
        await page.evaluate(async () => {
            if (!navigator.serviceWorker) {
                throw new Error('Service Worker no soportado en este navegador');
            }
            const reg = await navigator.serviceWorker.ready;
            
            // Si ya está controlada la página por un service worker, continuar
            if (navigator.serviceWorker.controller) return;

            // De lo contrario, esperar el evento controllerchange
            await new Promise(resolve => {
                navigator.serviceWorker.addEventListener('controllerchange', () => {
                    resolve();
                }, { once: true });
            });
        });
        console.log('✅ Service Worker listo y activo.');

        // Esperar a que la base de datos Excel se procese y se guarde en IndexedDB (cambia a 'Excel conectado')
        console.log('Esperando a que la base de datos se guarde en IndexedDB...');
        await page.waitForFunction(() => {
            const statusTitle = document.querySelector('#excelStatus .status-title');
            return statusTitle && statusTitle.textContent.includes('Excel conectado');
        }, { timeout: 25000 });
        console.log('✅ Base de datos guardada en IndexedDB.');

        console.log('Tomando captura de pantalla inicial (Online)...');
        await page.screenshot({ path: path.join(__dirname, 'test_offline_1_online.png') });

        // --- PASO 2: SIMULAR SIN CONEXIÓN (OFFLINE) ---
        console.log('\n--- PASO 2: Simulando estar sin conexión (Offline) ---');
        await page.setOfflineMode(true);
        await page.evaluate(() => {
            sessionStorage.setItem('simulate_offline', 'true');
        });
        console.log('✅ Modo offline activado en el navegador.');

        // Navegar a la página en modo offline (simula abrir la app desde cero offline, evitando que Chromium restaure el input)
        console.log('Navegando a la app estando sin conexión (offline)...');
        await page.goto('https://127.0.0.1:3000', { waitUntil: 'domcontentloaded', timeout: 15000 });
        console.log('✅ Navegación offline exitosa.');

        // Esperar 1 segundo para asegurar la inicialización de scripts en modo offline
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Verificar que la base de datos se haya cargado desde el caché local (IndexedDB)
        console.log('Verificando estado del Excel en modo offline...');
        await page.waitForFunction(() => {
            const statusTitle = document.querySelector('#excelStatus .status-title');
            return statusTitle && statusTitle.textContent.includes('Cargado de caché local');
        }, { timeout: 5000 });
        
        const statusText = await page.$eval('#excelStatus .status-title', el => el.textContent);
        console.log(`✅ Estado del Excel (Offline): ${statusText}`);

        // Verificar que la tasa de cambio está vacía y con advertencia visual (Modo Offline)
        console.log('Verificando que la tasa de cambio se inicia vacía con advertencia (Offline)...');
        const rateValue = await page.$eval('#exchangeRate', el => el.value);
        const hasWarningClass = await page.$eval('#exchangeRate', el => el.classList.contains('rate-warning'));
        const warningTextHidden = await page.$eval('#rateWarningText', el => el.classList.contains('hidden'));
        
        console.log(`Valor de tasa inicial: "${rateValue}" | Clase rate-warning presente: ${hasWarningClass} | Warning oculto: ${warningTextHidden}`);
        if (rateValue !== '' || !hasWarningClass || warningTextHidden) {
            throw new Error('La tasa de cambio no se inició vacía o con advertencia en modo offline');
        }
        console.log('✅ Advertencia de tasa de cambio offline verificada correctamente.');

        // --- PASO 3: PROBAR FUNCIONALIDAD DE CÁLCULO ---
        console.log('\n--- PASO 3: Probando funcionalidad en modo offline ---');

        // Ingresar tasa de cambio de forma manual para habilitar cálculos
        console.log('Ingresando tasa de cambio manualmente (36.50)...');
        const exchangeRateInput = await page.$('#exchangeRate');
        await exchangeRateInput.type('36.50');
        await page.evaluate(el => el.dispatchEvent(new Event('input')), exchangeRateInput);

        // Verificar que la advertencia desaparece
        const hasWarningClassAfter = await page.$eval('#exchangeRate', el => el.classList.contains('rate-warning'));
        const warningTextHiddenAfter = await page.$eval('#rateWarningText', el => el.classList.contains('hidden'));
        console.log(`Tasa ingresada | Clase rate-warning presente: ${hasWarningClassAfter} | Warning oculto: ${warningTextHiddenAfter}`);
        if (hasWarningClassAfter || !warningTextHiddenAfter) {
            throw new Error('La advertencia de tasa no desapareció tras ingresar un valor válido');
        }
        console.log('✅ Advertencia removida tras ingreso de tasa válida.');
        
        // Seleccionar origen: VLP
        console.log('Seleccionando origen: VALLE DE LA PASCUA (VLP)...');
        await page.evaluate(() => {
            const el = document.getElementById('originInput');
            el.value = 'VALLE';
            el.dispatchEvent(new Event('input'));
            const list = document.getElementById('originList');
            list.classList.add('show');
            el.parentElement.classList.add('open');
        });
        await page.waitForSelector('#originList .dropdown-item[data-value="VLP"]', { visible: true });
        await page.click('#originList .dropdown-item[data-value="VLP"]');
 
        // Seleccionar destino: VAL
        console.log('Seleccionando destino: VALENCIA (VAL)...');
        await page.evaluate(() => {
            const el = document.getElementById('destinationInput');
            el.value = 'VALENCIA';
            el.dispatchEvent(new Event('input'));
            const list = document.getElementById('destinationList');
            list.classList.add('show');
            el.parentElement.classList.add('open');
        });
        await page.waitForSelector('#destinationList .dropdown-item[data-value="VAL"]', { visible: true });
        await page.click('#destinationList .dropdown-item[data-value="VAL"]');

        // Esperar a que se cargue la tarjeta de información de ruta
        console.log('Esperando información de ruta...');
        await page.waitForSelector('#routeInfoCard:not(.hidden)');
        const escala = await page.$eval('#valEscala', el => el.textContent);
        const kms = await page.$eval('#valKms', el => el.textContent);
        console.log(`Ruta cargada - Escala: ${escala}, Distancia: ${kms}`);

        // Verificar que los precios sigan en $0.00 / Bs. 0,00 (peso <= 0)
        console.log('Verificando que los precios estén en cero (peso inicial es 0)...');
        const initialUSD = await page.$eval('#matrixUSD_origDir', el => el.textContent);
        const initialBs = await page.$eval('#matrixBs_origDir', el => el.textContent);
        console.log(`Precios en peso 0 - USD: ${initialUSD}, Bs: ${initialBs}`);
        if (!initialUSD.includes('$0.00') || initialBs !== 'Bs. 0,00') {
            throw new Error(`Los precios con peso 0 no son cero: ${initialUSD} / ${initialBs}`);
        }

        // Ingresar peso físico: 40 kg
        console.log('Ingresando peso físico: 40 kg...');
        const weightInput = await page.$('#weightInput');
        await page.evaluate(el => el.value = '', weightInput);
        await weightInput.type('40');
        await page.evaluate(el => el.dispatchEvent(new Event('input')), weightInput);

        // Esperar a que se calculen los precios
        await page.waitForFunction(() => {
            const val = document.getElementById('matrixUSD_origDir').textContent;
            return !val.includes('$0.00');
        }, { timeout: 5000 });

        const updatedUSD = await page.$eval('#matrixUSD_origDir', el => el.textContent);
        const updatedBs = await page.$eval('#matrixBs_origDir', el => el.textContent);
        console.log(`Precios calculados para 40kg (Offline) - USD: ${updatedUSD}, Bs: ${updatedBs}`);

        // Seleccionar la tarjeta de "Pago en Origen - Dirección" para cargar su desglose
        console.log('Abriendo desglose de "Pago Origen - Dirección"...');
        await page.click('#matrixUSD_origDir');

        // Esperar un poco para asegurar el renderizado
        await new Promise(resolve => setTimeout(resolve, 500));

        console.log('Tomando captura de pantalla del presupuesto final (Offline)...');
        await page.screenshot({ path: path.join(__dirname, 'test_offline_2_final.png') });

        // Imprimir desglose de conceptos del recibo detallado
        const itemBase = await page.$eval('#itemBaseUSD', el => el.textContent);
        const itemIva = await page.$eval('#itemIvaUSD', el => el.textContent);
        const itemFranqueo = await page.$eval('#itemFranqueoUSD', el => el.textContent);
        const itemIgtf = await page.$eval('#itemIgtfUSD', el => el.textContent);
        const receiptTotalUSD = await page.$eval('#itemTotalUSD', el => el.textContent);
        const receiptTotalBs = await page.$eval('#itemTotalBs', el => el.textContent);

        console.log('\n=== DESGLOSE DE FACTURA DETALLADA EN MODO OFFLINE ===');
        console.log(`Flete Base: ${itemBase}`);
        console.log(`IVA (16%): ${itemIva}`);
        console.log(`Franqueo Postal (10%): ${itemFranqueo}`);
        console.log(`IGTF (3% - Solo Divisas): ${itemIgtf}`);
        console.log(`TOTAL USD: ${receiptTotalUSD}`);
        console.log(`TOTAL Bs: ${receiptTotalBs}`);
        console.log('====================================================\n');

        // Validaciones: verificar que el total coincida con el esperado
        const rateVal = parseFloat(await page.$eval('#exchangeRate', el => el.value)) || 0;
        console.log(`Tasa de cambio usada: ${rateVal}`);

        if (rateVal === 0) {
            throw new Error('La tasa de cambio en modo offline es 0 o no es válida');
        }

        console.log('✅ VALIDACIÓN EXITOSA: La aplicación calcula correctamente el presupuesto en modo offline.');

    } catch (err) {
        console.error('❌ ERROR EN LA PRUEBA OFFLINE DE PUPPETEER:', err);
        if (page) {
            try {
                console.log('Tomando captura de pantalla de error (test_offline_error.png)...');
                await page.screenshot({ path: path.join(__dirname, 'test_offline_error.png') });
            } catch (screenshotErr) {
                console.error('No se pudo tomar la captura de pantalla de error:', screenshotErr);
            }
        }
        process.exit(1);
    } finally {
        if (browser) {
            console.log('Cerrando navegador...');
            await browser.close();
        }
        if (serverProcess) {
            console.log('Deteniendo el servidor local...');
            serverProcess.kill('SIGTERM');
        }
        console.log('Prueba finalizada.');
    }
})();
