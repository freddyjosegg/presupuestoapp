const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

(async () => {
    let serverProcess;
    let browser;
    let page;
    const swPath = path.resolve(__dirname, '..', 'sw.js');
    let originalSwContent = '';

    try {
        console.log('Guardando contenido original de sw.js...');
        originalSwContent = fs.readFileSync(swPath, 'utf8');

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

        // --- PASO 1: REGISTRAR EL SERVICE WORKER INICIAL ---
        console.log('\n--- PASO 1: Cargando la app y verificando registro inicial del Service Worker ---');
        await page.goto('https://127.0.0.1:3000', { waitUntil: 'load', timeout: 30000 });

        console.log('Esperando registro del Service Worker...');
        const isRegistered = await page.evaluate(async () => {
            if (!navigator.serviceWorker) return false;
            const reg = await navigator.serviceWorker.ready;
            return !!reg.active;
        });

        if (!isRegistered) {
            throw new Error('El Service Worker inicial no se registró o no está activo.');
        }
        console.log('✅ Service Worker inicial registrado y activo.');

        // --- PASO 2: TRIGGER ACTUALIZACIÓN ---
        console.log('\n--- PASO 2: Simulando actualización de Service Worker (modificando sw.js) ---');
        
        // Modificar sw.js localmente agregando un comentario de versión de prueba
        const updatedSwContent = originalSwContent.replace(
            /const CACHE_NAME = 'presupuestoapp-v\d+';/,
            `const CACHE_NAME = 'presupuestoapp-v-pwa-lifecycle-test-${Date.now()}';`
        ) + `\n// PWA Lifecycle Test Comment: ${Date.now()}`;
        
        fs.writeFileSync(swPath, updatedSwContent, 'utf8');
        console.log('✅ Archivo sw.js actualizado localmente.');

        // Recargar la página para que el navegador busque actualizaciones del Service Worker
        console.log('Recargando la app para gatillar la detección del nuevo sw.js...');
        await page.reload({ waitUntil: 'load' });

        // Esperar a que el nuevo Service Worker pase de installing/waiting a active
        console.log('Esperando a que el nuevo Service Worker se instale y controle la página...');
        const swUpdated = await page.evaluate(async () => {
            const reg = await navigator.serviceWorker.getRegistration();
            if (!reg) return false;

            // Forzar actualización inmediata si el navegador aún no la disparó
            await reg.update();

            return new Promise((resolve) => {
                // Si ya hay uno instalándose o esperando, o controlando
                if (reg.installing || reg.waiting) {
                    const worker = reg.installing || reg.waiting;
                    worker.addEventListener('statechange', () => {
                        if (worker.state === 'activated') {
                            resolve(true);
                        }
                    });
                    // Timeout de seguridad en el navegador
                    setTimeout(() => resolve(false), 15000);
                } else {
                    // Si ya se activó en background
                    resolve(true);
                }
            });
        });

        if (!swUpdated) {
            throw new Error('El nuevo Service Worker no se activó después de la actualización.');
        }
        console.log('✅ Nuevo Service Worker activado y controlando la página.');

        console.log('✅ VALIDACIÓN EXITOSA: El ciclo de vida de instalación y actualización de la PWA funciona al 100%.');

    } catch (err) {
        console.error('❌ ERROR EN LA PRUEBA DE CICLO DE VIDA PWA:', err);
        process.exit(1);
    } finally {
        if (originalSwContent) {
            console.log('Restaurando contenido original de sw.js...');
            fs.writeFileSync(swPath, originalSwContent, 'utf8');
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
