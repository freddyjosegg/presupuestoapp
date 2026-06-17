const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const { PNG } = require('pngjs');
const pixelmatch = require('pixelmatch').default;

(async () => {
    let serverProcess;
    let browser;
    let page;
    const baselineDir = path.resolve(__dirname, 'baseline');
    const tempDir = path.resolve(__dirname, 'temp');

    try {
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
        
        page.on('console', msg => console.log(`[NAVEGADOR] ${msg.text()}`));
        page.on('pageerror', err => console.log(`[NAVEGADOR ERROR] ${err.toString()}`));

        // Crear directorios si no existen
        if (!fs.existsSync(baselineDir)) {
            fs.mkdirSync(baselineDir, { recursive: true });
        }
        if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir, { recursive: true });
        }

        // --- FUNCION DE PRUEBA DE REGRESIÓN VISUAL POR VISTA ---
        const runVisualTest = async (viewport, name) => {
            console.log(`\n--- Probando resolución: ${name} (${viewport.width}x${viewport.height}) ---`);
            await page.setViewport(viewport);

            console.log('Navegando a la app local...');
            await page.goto('https://127.0.0.1:3000', { waitUntil: 'load', timeout: 30000 });

            // Esperar a que cargue el Excel por defecto
            await page.waitForFunction(() => {
                const statusTitle = document.querySelector('#excelStatus .status-title');
                return statusTitle && (statusTitle.textContent.includes('Excel conectado') || statusTitle.textContent.includes('Cargado de caché local'));
            }, { timeout: 15000 });

            // Rellenar un caso estándar para renderizar toda la UI completa
            console.log('Rellenando campos de prueba...');
            
            // Seleccionar origen: VLP
            await page.evaluate(() => {
                const el = document.getElementById('originInput');
                el.value = 'VALLE';
                el.dispatchEvent(new Event('input'));
                const list = document.getElementById('originList');
                list.classList.add('show');
                el.parentElement.classList.add('open');
            });
            await page.waitForSelector('#originList .dropdown-item[data-value="VLP"]', { visible: true });
            await page.evaluate(() => {
                document.querySelector('#originList .dropdown-item[data-value="VLP"]').click();
            });

            // Seleccionar destino: VAL
            await page.evaluate(() => {
                const el = document.getElementById('destinationInput');
                el.value = 'VALENCIA';
                el.dispatchEvent(new Event('input'));
                const list = document.getElementById('destinationList');
                list.classList.add('show');
                el.parentElement.classList.add('open');
            });
            await page.waitForSelector('#destinationList .dropdown-item[data-value="VAL"]', { visible: true });
            await page.evaluate(() => {
                document.querySelector('#destinationList .dropdown-item[data-value="VAL"]').click();
            });

            // Peso físico: 40 kg
            const weightInput = await page.$('#weightInput');
            await page.evaluate(el => el.value = '', weightInput);
            await weightInput.type('40');
            await page.evaluate(el => el.dispatchEvent(new Event('input')), weightInput);

            // Activar algunos servicios adicionales
            await page.evaluate(() => {
                document.getElementById('chkCar').click();
                document.getElementById('chkContenedor').click();
            });

            // Esperar cálculo
            await page.waitForFunction(() => {
                const val = document.getElementById('matrixUSD_origDir').textContent;
                return !val.includes('$0.00');
            }, { timeout: 5000 });

            // Hacer click en la primera tarjeta para mostrar desglose de factura
            await page.evaluate(() => {
                document.getElementById('matrixUSD_origDir').click();
            });
            await new Promise(resolve => setTimeout(resolve, 800)); // Esperar renderizado y transiciones suavemente

            // Definir rutas de imágenes
            const baselinePath = path.join(baselineDir, `${name}.png`);
            const currentPath = path.join(tempDir, `${name}.png`);
            const diffPath = path.join(tempDir, `${name}_diff.png`);

            // Capturar pantalla actual
            console.log(`Tomando captura de pantalla actual para ${name}...`);
            await page.screenshot({ path: currentPath, fullPage: true });

            // Si no hay baseline, guardamos la actual como baseline
            if (!fs.existsSync(baselinePath)) {
                console.log(`[AVISO] No se encontró baseline para ${name}. Guardando captura actual como referencia.`);
                fs.copyFileSync(currentPath, baselinePath);
                return;
            }

            // Comparar pixel a pixel con pixelmatch
            console.log('Comparando captura con la imagen patrón (baseline)...');
            const imgBaseline = PNG.sync.read(fs.readFileSync(baselinePath));
            const imgCurrent = PNG.sync.read(fs.readFileSync(currentPath));

            const { width, height } = imgBaseline;
            const diff = new PNG({ width, height });

            // Ejecutar comparación
            const numDiffPixels = pixelmatch(
                imgBaseline.data,
                imgCurrent.data,
                diff.data,
                width,
                height,
                { threshold: 0.15 } // Tolerancia de suavizado de subpíxeles
            );

            console.log(`Resultado de comparación para ${name}: ${numDiffPixels} píxeles diferentes.`);

            // Guardar imagen de diferencias
            fs.writeFileSync(diffPath, PNG.sync.write(diff));

            // Si los píxeles diferentes superan un umbral de tolerancia por renderizado de fuentes (ej. 300 píxeles)
            const tolerance = 500; // Tolerancia razonable para cambios menores de suavizado de tipografías
            if (numDiffPixels > tolerance) {
                throw new Error(`¡Regresión visual detectada en la vista ${name}! Diferencia de ${numDiffPixels} píxeles. Imagen de diferencias guardada en: ${diffPath}`);
            }

            console.log(`✅ Regresión visual pasada con éxito para la vista ${name}.`);
        };

        // Ejecutar test para Desktop
        await runVisualTest({ width: 1280, height: 1000 }, 'desktop');

        // Ejecutar test para Mobile
        await runVisualTest({ width: 375, height: 812 }, 'mobile');

        console.log('\n✅ VALIDACIÓN EXITOSA: La regresión visual ha pasado correctamente en Desktop y Mobile.');

    } catch (err) {
        console.error('❌ ERROR EN LA PRUEBA DE REGRESIÓN VISUAL:', err.message || err);
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
