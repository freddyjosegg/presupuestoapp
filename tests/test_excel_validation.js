const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

(async () => {
    let serverProcess;
    let browser;
    let page;
    const tempDir = path.resolve(__dirname, 'temp_excel');
    const badExcelPath = path.join(tempDir, 'corrupt_tariff.xlsx');

    try {
        console.log('Creando archivo de Excel corrupto para la prueba...');
        if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir);
        }
        // Escribir contenido basura (no es un zip/xlsx válido)
        fs.writeFileSync(badExcelPath, 'Contenido basura para simular un archivo Excel corrupto');

        console.log('Iniciando el servidor local en segundo plano (server.js)...');
        serverProcess = spawn('node', ['./server.js'], {
            cwd: path.resolve(__dirname, '..')
        });

        serverProcess.stdout.on('data', (data) => {
            console.log(`[SERVIDOR] ${data.toString().trim()}`);
        });

        serverProcess.stderr.on('data', (data) => {
            console.error(`[SERVIDOR ERROR] ${data.toString().trim()}`);
        });

        // Esperar a que el servidor esté listo
        await new Promise(resolve => setTimeout(resolve, 2000));

        console.log('Iniciando navegador Puppeteer...');
        browser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--ignore-certificate-errors']
        });
        
        page = await browser.newPage();
        await page.setViewport({ width: 1280, height: 900 });

        page.on('console', msg => console.log(`[NAVEGADOR] ${msg.text()}`));
        page.on('pageerror', err => console.log(`[NAVEGADOR ERROR] ${err.toString()}`));

        // --- PASO 1: CARGAR LA APP ONLINE ---
        console.log('\n--- PASO 1: Cargando la app local ---');
        await page.goto('https://127.0.0.1:3000', { waitUntil: 'load', timeout: 30000 });

        // Esperar a que la base de datos por defecto cargue primero
        await page.waitForFunction(() => {
            const statusTitle = document.querySelector('#excelStatus .status-title');
            return statusTitle && (statusTitle.textContent.includes('Excel conectado') || statusTitle.textContent.includes('Cargado de caché local'));
        }, { timeout: 15000 });
        console.log('✅ Base de datos por defecto cargada.');

        // --- PASO 2: SUBIR EXCEL CORRUPTO ---
        console.log('\n--- PASO 2: Subiendo archivo Excel corrupto/malo ---');
        const fileInput = await page.$('#excelFileInput');
        if (!fileInput) {
            throw new Error('No se encontró el input de archivos #excelFileInput.');
        }

        // Subir el archivo corrupto
        await fileInput.uploadFile(badExcelPath);
        console.log('✅ Archivo corrupto subido al input.');

        // --- PASO 3: VERIFICAR MANEJO DE ERRORES ---
        console.log('\n--- PASO 3: Verificando que el error se maneje con elegancia ---');

        // 1. Verificar el toast de error
        console.log('Esperando mensaje de error en el Toast...');
        await page.waitForFunction(() => {
            const toastEl = document.getElementById('toast');
            return toastEl && !toastEl.classList.contains('hidden') && toastEl.classList.contains('error');
        }, { timeout: 8000 });

        const toastText = await page.$eval('#toast', el => el.textContent);
        console.log(`✅ Mensaje en Toast detectado: "${toastText}"`);
        if (!toastText.includes('Error al parsear')) {
            throw new Error('El mensaje de error del toast no es el esperado.');
        }

        // 2. Verificar el indicador de estado del Excel
        console.log('Verificando estado de error en la tarjeta de estado del Excel...');
        const statusIndicatorClass = await page.$eval('#excelStatus .status-indicator', el => el.className);
        const statusTitle = await page.$eval('#excelStatus .status-title', el => el.textContent);
        const statusDesc = await page.$eval('#excelStatus .status-desc', el => el.textContent);

        console.log(`Clase del indicador: "${statusIndicatorClass}"`);
        console.log(`Título del estado: "${statusTitle}"`);
        console.log(`Descripción del estado: "${statusDesc}"`);

        if (!statusIndicatorClass.includes('error') || !statusTitle.includes('Error al procesar Excel')) {
            throw new Error('La tarjeta de estado no refleja un estado de error correcto.');
        }
        console.log('✅ Estado de error visualizado correctamente en el header.');

        console.log('✅ VALIDACIÓN EXITOSA: La app maneja archivos corruptos sin crashear y notifica al usuario.');

    } catch (err) {
        console.error('❌ ERROR EN LA PRUEBA DE VALIDACIÓN DE EXCEL:', err);
        process.exit(1);
    } finally {
        // Limpieza de archivos temporales
        try {
            if (fs.existsSync(badExcelPath)) {
                fs.unlinkSync(badExcelPath);
            }
            if (fs.existsSync(tempDir)) {
                fs.rmdirSync(tempDir);
            }
        } catch (cleanupErr) {
            console.error('Error al limpiar archivos temporales de prueba:', cleanupErr);
        }

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
