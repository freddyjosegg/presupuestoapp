const fs = require('fs');
const path = require('path');

module.exports = (req, res) => {
    // Enable CORS
    res.setHeader('Access-Control-Allow-Credentials', true);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
    res.setHeader(
        'Access-Control-Allow-Headers',
        'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
    );

    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
    }

    try {
        const dbDir = path.join(process.cwd(), 'DB');
        
        if (!fs.existsSync(dbDir)) {
            return res.status(404).json({ error: 'No se encontró la carpeta DB' });
        }

        fs.readdir(dbDir, (err, files) => {
            if (err) {
                return res.status(500).json({ error: 'No se pudo leer la carpeta DB', details: err.message });
            }

            // Filter for Excel files, excluding temporary ones
            const excelFiles = files.filter(file => {
                const ext = path.extname(file).toLowerCase();
                return (ext === '.xlsx' || ext === '.xlsm') && !file.startsWith('~$');
            });

            if (excelFiles.length === 0) {
                return res.status(404).json({ error: 'No se encontró ningún archivo Excel en la carpeta DB' });
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
                            return res.status(404).json({ error: 'No se pudieron leer estadísticas de los archivos' });
                        }

                        // Sort descending by mtime (newest first)
                        fileStats.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));

                        return res.status(200).json(fileStats[0]);
                    }
                });
            });
        });
    } catch (error) {
        return res.status(500).json({ error: 'Error interno de la serverless function', details: error.message });
    }
};
