const https = require('https');
const fs = require('fs');
const path = require('path');

const PORT = 3000;

// HTTPS self-signed certificate options
const options = {
    key: fs.readFileSync(path.join(__dirname, 'key.pem')),
    cert: fs.readFileSync(path.join(__dirname, 'cert.pem'))
};

const MIME_TYPES = {
    '.html': 'text/html',
    '.css': 'text/css',
    '.js': 'text/javascript',
    '.json': 'application/json',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.gif': 'image/gif',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12'
};

const server = https.createServer(options, (req, res) => {
    console.log(`${req.method} ${req.url}`);

    // Parse URL path and decode URI components (supporting spaces in filenames)
    let cleanUrl = decodeURIComponent(req.url.split('?')[0]);

    // Endpoint to get the latest Excel file metadata
    if (cleanUrl === '/api/latest-excel') {
        const dbDir = path.resolve(__dirname, './DB');
        fs.readdir(dbDir, (err, files) => {
            if (err) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'No se pudo leer la carpeta DB', details: err.message }));
                return;
            }

            // Filter for Excel files, excluding temporary ones
            const excelFiles = files.filter(file => {
                const ext = path.extname(file).toLowerCase();
                return (ext === '.xlsx' || ext === '.xlsm') && !file.startsWith('~$');
            });

            if (excelFiles.length === 0) {
                res.writeHead(404, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: 'No se encontró ningún archivo Excel en la carpeta DB' }));
                return;
            }

            let fileStats = [];
            let processed = 0;

            excelFiles.forEach(file => {
                const filePath = path.join(dbDir, file);
                fs.stat(filePath, (statErr, stats) => {
                    if (!statErr) {
                        fileStats.push({
                            name: file,
                            url: `/DB/${encodeURIComponent(file)}`,
                            lastModified: stats.mtime,
                            size: stats.size
                        });
                    }

                    processed++;
                    if (processed === excelFiles.length) {
                        if (fileStats.length === 0) {
                            res.writeHead(404, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'No se pudieron leer estadísticas de los archivos' }));
                            return;
                        }

                        // Sort descending by mtime (newest first)
                        fileStats.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));

                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify(fileStats[0]));
                    }
                });
            });
        });
        return;
    }

    let filePath = cleanUrl === '/' ? './index.html' : '.' + cleanUrl;
    filePath = path.resolve(__dirname, filePath);

    // Security check: ensure filePath is within workspace
    if (!filePath.startsWith(__dirname)) {
        res.statusCode = 403;
        res.end('Acceso denegado');
        return;
    }

    const extname = String(path.extname(filePath)).toLowerCase();
    const contentType = MIME_TYPES[extname] || 'application/octet-stream';

    fs.readFile(filePath, (error, content) => {
        if (error) {
            if (error.code === 'ENOENT') {
                res.writeHead(404, { 'Content-Type': 'text/html' });
                res.end('<h1>404 Archivo no encontrado</h1>', 'utf-8');
            } else {
                res.writeHead(500);
                res.end(`Error de servidor: ${error.code} ..\n`);
            }
        } else {
            res.writeHead(200, { 'Content-Type': contentType });
            res.end(content, 'utf-8');
        }
    });
});

server.listen(PORT, () => {
    console.log(`\n==================================================`);
    console.log(`PresupuestoApp (HTTPS) está corriendo en:`);
    console.log(`👉 https://localhost:${PORT}`);
    console.log(`==================================================\n`);
});
