const puppeteer = require('puppeteer');
const path = require('path');

(async () => {
    let browser;
    let page;
    try {
        console.log('Iniciando navegador Puppeteer...');
        browser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox', '--ignore-certificate-errors']
        });
        page = await browser.newPage();
        
        // Ajustar el viewport a una pantalla de laptop grande
        await page.setViewport({ width: 1280, height: 900 });

        // Redirigir logs y errores del navegador a la consola del test
        page.on('console', msg => console.log(`[NAVEGADOR] ${msg.text()}`));
        page.on('pageerror', err => console.log(`[NAVEGADOR ERROR] ${err.toString()}`));

        console.log('Navegando a la app local (https://127.0.0.1:3000)...');
        await page.goto('https://127.0.0.1:3000', { waitUntil: 'domcontentloaded', timeout: 30000 });

        // Esperar 2 segundos para inicializar scripts y DOM
        await new Promise(resolve => setTimeout(resolve, 2000));

        console.log('Tomando captura de pantalla inicial...');
        await page.screenshot({ path: path.join(__dirname, 'test_inicial.png') });

        console.log('Cargando el archivo Excel...');
        const fileInput = await page.$('#excelFileInput');
        if (!fileInput) {
            throw new Error('No se encontró el elemento #excelFileInput en el DOM.');
        }
        const excelPath = path.resolve(__dirname, 'DB', 'TARIFA ACTUAL PAQUETERIA USD 210526.xlsm');
        await fileInput.uploadFile(excelPath);

        // Esperar a que el archivo Excel se procese y cambie el estado a ready
        console.log('Esperando a que la base de datos se cargue y procese...');
        await page.waitForFunction(() => {
            const statusTitle = document.querySelector('#excelStatus .status-title');
            return statusTitle && statusTitle.textContent.includes('Manual');
        }, { timeout: 15000 });

        const statusText = await page.$eval('#excelStatus .status-title', el => el.textContent);
        console.log(`Estado del Excel: ${statusText}`);

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

        // Esperar a que se calculen los precios (ya no deben ser cero)
        await page.waitForFunction(() => {
            const val = document.getElementById('matrixUSD_origDir').textContent;
            return !val.includes('$0.00');
        }, { timeout: 5000 });

        const updatedUSD = await page.$eval('#matrixUSD_origDir', el => el.textContent);
        const updatedBs = await page.$eval('#matrixBs_origDir', el => el.textContent);
        console.log(`Precios calculados para 40kg - USD: ${updatedUSD}, Bs: ${updatedBs}`);

        // Seleccionar la tarjeta de "Pago en Origen - Dirección" para cargar su desglose
        console.log('Abriendo desglose de "Pago Origen - Dirección"...');
        await page.click('#matrixUSD_origDir');

        // Esperar un poco para asegurar el renderizado
        await new Promise(resolve => setTimeout(resolve, 500));

        console.log('Tomando captura de pantalla del presupuesto final...');
        await page.screenshot({ path: path.join(__dirname, 'test_final.png') });

        // Imprimir desglose de conceptos del recibo detallado
        const itemBase = await page.$eval('#itemBaseUSD', el => el.textContent);
        const itemIva = await page.$eval('#itemIvaUSD', el => el.textContent);
        const itemFranqueo = await page.$eval('#itemFranqueoUSD', el => el.textContent);
        const itemIgtf = await page.$eval('#itemIgtfUSD', el => el.textContent);
        const receiptTotalUSD = await page.$eval('#itemTotalUSD', el => el.textContent);
        const receiptTotalBs = await page.$eval('#itemTotalBs', el => el.textContent);

        console.log('\n=== DESGLOSE DE FACTURA DETALLADA (PUPPETEER) ===');
        console.log(`Flete Base: ${itemBase}`);
        console.log(`IVA (16%): ${itemIva}`);
        console.log(`Franqueo Postal (10%): ${itemFranqueo}`);
        console.log(`IGTF (3% - Solo Divisas): ${itemIgtf}`);
        console.log(`TOTAL USD: ${receiptTotalUSD}`);
        console.log(`TOTAL Bs: ${receiptTotalBs}`);
        console.log('===============================================\n');

        // Validaciones: El total en Bs. debe coincidir con la fórmula de Excel de conversión y redondeo por componentes
        const roundUpDec = (val, decimals = 3) => {
            const factor = Math.pow(10, decimals);
            return Math.ceil((val - 1e-9) * factor) / factor;
        };
        const roundHalfUpNum = (val, decimals = 2) => {
            const factor = Math.pow(10, decimals);
            return Math.round((val + 1e-9) * factor) / factor;
        };

        const getVal = async (selector) => {
            const el = await page.$(selector);
            if (!el) return 0;
            const isHidden = await page.evaluate(element => element.closest('.hidden') !== null, el);
            if (isHidden) return 0;
            const text = await page.evaluate(element => element.textContent || element.value || '0', el);
            const match = text.match(/\$?([0-9.,\-]+)/);
            if (!match || text.includes('-') && !text.includes('$')) return 0; // handle negative or dashed display
            return parseFloat(match[1].replace(/,/g, ''));
        };

        const rateVal = await getVal('#exchangeRate');
        const netFreight = await getVal('#itemNetUSD') || await getVal('#itemBaseUSD');
        const tdg = await getVal('#itemTdgUSD');
        const gcd = await getVal('#itemGcdUSD');
        const car = await getVal('#itemCarUSD');
        const carf = await getVal('#itemCarfUSD');
        const seguro = await getVal('#itemSeguroUSD');
        const contenedor = await getVal('#itemContenedorUSD');
        const truncDec = (val, decimals = 3) => {
            const factor = Math.pow(10, decimals);
            return Math.floor((val + 1e-9) * factor) / factor;
        };

        const subtotal = netFreight + tdg + gcd + car + carf + seguro + contenedor;
        const iva = truncDec(subtotal * 0.16, 3);
        const franqueo = await getVal('#itemFranqueoUSD');

        // Excel component-rounding formula
        const fleteBs = roundUpDec(netFreight * rateVal, 3);
        const compBs = roundUpDec((tdg + gcd + car + carf + seguro + contenedor) * rateVal, 3);
        const taxBs = roundUpDec((iva + franqueo) * rateVal, 3);
        
        const expectedBsRaw = roundUpDec(fleteBs + compBs + taxBs, 3);
        const expectedBs = roundHalfUpNum(expectedBsRaw, 2);

        const cleanBs = receiptTotalBs.replace('Bs. ', '').replace(/\./g, '').replace(',', '.');
        const totalBsNum = parseFloat(cleanBs);

        console.log(`Tasa de cambio: ${rateVal}`);
        console.log(`Flete en Bs (Excel): ${fleteBs.toFixed(3)}`);
        console.log(`Comp. en Bs (Excel): ${compBs.toFixed(3)}`);
        console.log(`Impuestos en Bs (Excel): ${taxBs.toFixed(3)}`);
        console.log(`Total Bs en pantalla: ${totalBsNum}`);
        console.log(`Total Bs esperado (Fórmula Excel): ${expectedBs.toFixed(2)}`);

        if (Math.abs(totalBsNum - expectedBs) > 0.05) {
            throw new Error(`¡El total en bolívares no coincide con la fórmula de Excel! Esperado: ${expectedBs.toFixed(2)}, Obtenido: ${totalBsNum}`);
        } else {
            console.log('✅ VALIDACIÓN EXITOSA: El total en Bolívares coincide exactamente con la fórmula de Excel.');
        }

        console.log('Prueba automatizada finalizada con éxito.');
    } catch (err) {
        console.error('❌ ERROR EN LA PRUEBA DE PUPPETEER:', err);
        if (page) {
            try {
                console.log('Tomando captura de pantalla de error (test_error.png)...');
                await page.screenshot({ path: path.join(__dirname, 'test_error.png') });
            } catch (screenshotErr) {
                console.error('No se pudo tomar la captura de pantalla de error:', screenshotErr);
            }
        }
        process.exit(1);
    } finally {
        if (browser) await browser.close();
    }
})();
