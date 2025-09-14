// --- server.js - COD COMPLET FINAL (cu toate finisajele) ---

const express = require('express');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');
const { Pool } = require('pg');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

const pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

pool.query(`
    CREATE TABLE IF NOT EXISTS reguli (
        id SERIAL PRIMARY KEY,
        furnizor TEXT NOT NULL,
        criterii TEXT NOT NULL,
        tip_material TEXT NOT NULL,
        descriere_raport TEXT NOT NULL
    )
`).then(() => console.log('Baza de date PostgreSQL și tabelul "reguli" sunt pregătite.'))
  .catch(err => console.error('Eroare la crearea tabelului:', err));

const app = express();
app.use(express.json());
const PORT = process.env.PORT || 3000;
const upload = multer({ dest: 'uploads/' });
app.use(express.static('public'));

// --- API PENTRU ADMINISTRAREA REGULILOR ---
app.get('/api/reguli', async (req, res) => {
    try {
        const result = await pool.query('SELECT * FROM reguli ORDER BY tip_material ASC');
        res.json(result.rows);
    } catch (err) { res.status(500).json({ error: err.message }); }
});
app.post('/api/reguli', async (req, res) => {
    try {
        const { furnizor, criterii, tip_material, descriere_raport } = req.body;
        const sql = 'INSERT INTO reguli (furnizor, criterii, tip_material, descriere_raport) VALUES ($1, $2, $3, $4) RETURNING id';
        const result = await pool.query(sql, [furnizor, criterii, tip_material, descriere_raport]);
        res.json({ id: result.rows[0].id });
    } catch (err) { res.status(400).json({ error: err.message }); }
});
app.delete('/api/reguli/:id', async (req, res) => {
    try {
        await pool.query('DELETE FROM reguli WHERE id = $1', [req.params.id]);
        res.json({ message: "Șters" });
    } catch (err) { res.status(400).json({ error: err.message }); }
});
app.put('/api/reguli/:id', async (req, res) => {
    try {
        const { furnizor, criterii, tip_material, descriere_raport } = req.body;
        const sql = 'UPDATE reguli SET furnizor = $1, criterii = $2, tip_material = $3, descriere_raport = $4 WHERE id = $5';
        await pool.query(sql, [furnizor, criterii, tip_material, descriere_raport, req.params.id]);
        res.json({ message: "Succes" });
    } catch (err) { res.status(400).json({ error: err.message }); }
});

// --- RUTE PENTRU IMPORT / EXPORT REGULI ---
app.post('/api/migrate', upload.single('sablon_import'), async (req, res) => {
    const sablonFile = req.file;
    if (!sablonFile) { return res.status(400).send('Niciun fișier selectat.'); }
    const client = await pool.connect();
    try {
        console.log('Începem importul...');
        const workbook = xlsx.readFile(sablonFile.path);
        const sheetName = workbook.SheetNames[0];
        const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
        await client.query('BEGIN');
        await client.query('DELETE FROM reguli');
        const sql = `INSERT INTO reguli (furnizor, tip_material, descriere_raport, criterii) VALUES ($1, $2, $3, $4)`;
        let count = 0;
        for (const row of sheetData) {
            const tipMaterial = row['Tip Material'];
            const descriereRaport = row['Cod Culoare / Descriere'];
            const criterii = row['Material Corespondent (Criterii)'];
            const furnizor = row['Furnizor'] || 'ORICE';
            if (tipMaterial && descriereRaport && criterii) {
                await client.query(sql, [
                    String(furnizor).trim(),
                    String(tipMaterial).trim(),
                    String(descriereRaport).trim(),
                    String(criterii).trim()
                ]);
                count++;
            }
        }
        await client.query('COMMIT');
        console.log(`${count} de reguli au fost importate cu succes!`);
        res.send(`${count} de reguli au fost importate cu succes!`);
    } catch (error) {
        await client.query('ROLLBACK');
        console.error('Importul a eșuat:', error);
        res.status(500).send(`Importul a eșuat: ${error.message}`);
    } finally {
        client.release();
        if (sablonFile) fs.unlinkSync(sablonFile.path);
    }
});
app.get('/api/export-rules', async (req, res) => {
    try {
        console.log("Se generează fișierul de export...");
        const result = await pool.query('SELECT * FROM reguli ORDER BY tip_material ASC');
        const reguli = result.rows;
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Reguli Stoc');
        worksheet.columns = [
            { header: 'Tip Material', key: 'tip_material', width: 30 },
            { header: 'Cod Culoare / Descriere', key: 'descriere_raport', width: 35 },
            { header: 'Material Corespondent (Criterii)', key: 'criterii', width: 50 },
            { header: 'Furnizor', key: 'furnizor', width: 30 }
        ];
        worksheet.getRow(1).font = { bold: true };
        worksheet.addRows(reguli);
        const buffer = await workbook.xlsx.writeBuffer();
        const numeFisier = `reguli_backup_${new Date().toLocaleDateString('ro-RO').replace(/\./g, '-')}.xlsx`;
        res.setHeader('Content-Disposition', `attachment; filename="${numeFisier}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
        console.log("Exportul a fost generat.");
    } catch (error) {
        console.error("Eroare la exportul regulilor:", error);
        res.status(500).send("Eroare la exportul regulilor.");
    }
});

// --- API PENTRU A EXTRAGE FURNIZORII ---
app.post('/api/get-suppliers', upload.single('stoc'), (req, res) => {
    const stocFile = req.file;
    try {
        const workbook = xlsx.readFile(stocFile.path);
        let sheetName = workbook.SheetNames.find(name => name.includes('800'));
        if (!sheetName) { sheetName = workbook.SheetNames[0]; }
        if (!sheetName) { throw new Error("Fișierul Excel este gol sau corupt."); }
        const sheet = workbook.Sheets[sheetName];
        const stocData = xlsx.utils.sheet_to_json(sheet);
        const suppliers = new Set();
        for (const rand of stocData) {
            const supplierName = rand['Name 1'] || rand['Nume fz'];
            if (supplierName) { suppliers.add(supplierName.toString().trim()); }
        }
        res.json(Array.from(suppliers).sort());
    } catch (error) {
        console.error('Eroare la extragerea furnizorilor:', error);
        res.status(500).send(error.message);
    } finally {
        if (stocFile) fs.unlinkSync(stocFile.path);
    }
});

// --- LOGICA PRINCIPALĂ A APLICAȚIEI ---
async function runProcessing(stocFilePath, selectedSuppliers) {
    const client = await pool.connect();
    try {
        const reguliResult = await client.query('SELECT * FROM reguli');
        const reguli = reguliResult.rows;
        const workbook = xlsx.readFile(stocFilePath);
        let sheetName = workbook.SheetNames.find(name => name.includes('800'));
        if (!sheetName) {
            console.log("Avertisment: Nu s-a găsit un sheet cu '800' în nume. Se folosește primul sheet disponibil.");
            sheetName = workbook.SheetNames[0];
        }
        if (!sheetName) { throw new Error("Fișierul Excel este gol sau corupt."); }
        const sheet = workbook.Sheets[sheetName];
        const stocData = xlsx.utils.sheet_to_json(sheet);
        const filteredStocData = stocData.filter(rand => { const supplierName = rand['Name 1'] || rand['Nume fz']; return selectedSuppliers.includes(supplierName); });
        const cantitatiFinale = {};
        for (const rand of filteredStocData) {
            const descriere = rand['Material Description'] || '';
            const furnizor = rand['Name 1'] || rand['Nume fz'] || '';
            const cantitate = parseFloat(rand['Unrestricted'] || rand['Stoc (to)']) || 0;
            if (!descriere || cantitate <= 0) continue;
            const descriereCurataLower = descriere.toLowerCase().trim().replace(/-d$/, '');
            let regulaPotrivita = null;
            for (const regula of reguli) {
                const furnizorMatch = (regula.furnizor.toUpperCase() === 'ORICE' || regula.furnizor.toLowerCase() === furnizor.toLowerCase());
                let criteriiMatch = false;
                if (regula.criterii.includes('/')) {
                    const orCodes = regula.criterii.split('/').map(c => c.trim().toLowerCase());
                    if (orCodes.some(code => descriereCurataLower === code)) {
                        criteriiMatch = true;
                    }
                } else {
                    const andKeywords = regula.criterii.split(',').map(c => c.trim().toLowerCase()).filter(c => c);
                    if (andKeywords.length > 0 && andKeywords.every(keyword => descriereCurataLower.includes(keyword))) {
                        criteriiMatch = true;
                    }
                }
                if (furnizorMatch && criteriiMatch) {
                    regulaPotrivita = regula;
                    break;
                }
            }
            if (regulaPotrivita) {
                const cheieUnica = `${regulaPotrivita.tip_material}|${regulaPotrivita.descriere_raport}`;
                cantitatiFinale[cheieUnica] = (cantitatiFinale[cheieUnica] || 0) + cantitate;
            }
        }
        return cantitatiFinale;
    } finally {
        client.release();
    }
}

// --- RUTELE PRINCIPALE ---
const handleUploads = upload.fields([ { name: 'stoc', maxCount: 1 } ]);
app.post('/process', handleUploads, async (req, res) => {
    const stocFile = req.files.stoc[0];
    try {
        const selectedSuppliers = JSON.parse(req.body.suppliers);
        const rezultateStoc = await runProcessing(stocFile.path, selectedSuppliers);
        const rezultateFormatate = Object.keys(rezultateStoc).map(cheie => {
            const [tipMaterial, descriereRaport] = cheie.split('|');
            return { tipMaterial, descriereRaport, cantitate: rezultateStoc[cheie] };
        }).filter(item => item.cantitate >= 1);
        res.json(rezultateFormatate);
    } catch (error) {
        console.error('Eroare la /process:', error);
        res.status(500).send(error.message);
    } finally {
        if (stocFile) fs.unlinkSync(stocFile.path);
    }
});
app.post('/download', handleUploads, async (req, res) => {
    const stocFile = req.files.stoc[0];
    try {
        const selectedSuppliers = JSON.parse(req.body.suppliers);
        const rezultateStoc = await runProcessing(stocFile.path, selectedSuppliers);
        const excelBuffer = await generateExcelReport(rezultateStoc);
        const numeFisier = `Stoc_Materie_Prima_${new Date().toLocaleDateString('ro-RO').replace(/\./g, '-')}.xlsx`;
        res.setHeader('Content-Disposition', `attachment; filename="${numeFisier}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(excelBuffer);
    } catch (error) {
        console.error('Eroare la /download:', error);
        res.status(500).send(error.message);
    } finally {
        if (stocFile) fs.unlinkSync(stocFile.path);
    }
});
app.post('/download-pdf', handleUploads, async (req, res) => {
    const stocFile = req.files.stoc[0];
    try {
        const selectedSuppliers = JSON.parse(req.body.suppliers);
        const rezultateStoc = await runProcessing(stocFile.path, selectedSuppliers);
        const pdfBuffer = await generatePdfReport(rezultateStoc);
        const numeFisier = `Stoc_materie_prima-uz_extern-${new Date().toLocaleDateString('ro-RO').replace(/\./g, '-')}.pdf`;
        res.setHeader('Content-Disposition', `attachment; filename="${numeFisier}"`);
        res.setHeader('Content-Type', 'application/pdf');
        res.send(pdfBuffer);
    } catch (error) {
        console.error('Eroare la /download-pdf:', error);
        res.status(500).send(error.message);
    } finally {
        if (stocFile) fs.unlinkSync(stocFile.path);
    }
});

app.listen(PORT, () => { console.log(`Serverul FINAL a pornit la http://localhost:${PORT}`); });

// --- FUNCȚIILE DE GENERARE A RAPOARTELOR ---
async function generateExcelReport(rezultateStoc) {
    const workbook = new ExcelJS.Workbook();
    
    // Pas 1: Pregătirea datelor (cu formatare ZN)
    const dateTabel = Object.keys(rezultateStoc)
        .map(cheie => {
            const [tipMaterial, codCuloare] = cheie.split('|');
            const cantitate = rezultateStoc[cheie];
            let formattedCod = codCuloare;
            if (tipMaterial.trim() === 'ZN' && !isNaN(parseFloat(codCuloare))) {
                formattedCod = parseFloat(codCuloare).toFixed(2);
            }
            return {
                tip: tipMaterial.trim(),
                cod: formattedCod,
                cantitate: parseFloat(cantitate.toFixed(3)),
                status: cantitate >= 10 ? 'Stoc Suficient' : 'Stoc Redus'
            };
        })
        .filter(row => row.cantitate >= 1)
        .sort((a, b) => a.tip.localeCompare(b.tip));
    
    // Stiluri comune
    const legendValue = { richText: [{ font: { bold: true, color: { argb: 'FF000000' } }, text: '* ≥10 tone: Stoc Suficient ⚫\n' }, { font: { bold: true, color: { argb: 'FFFF0000' } }, text: '* 1-10 tone: Stoc Redus 🔴\n' }, { font: { bold: true, color: { argb: 'FFFF0000' } }, text: '* <1 tonă: Nu se afișează în acest tabel ❌' }] };
    const defaultFont11 = { name: 'Calibri', size: 11 };
    const redBoldFont11 = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFF0000' } };
    const greenFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } };
    const headerFont14 = { name: 'Calibri', size: 14, bold: true };
    const centerAlignment = { vertical: 'middle', horizontal: 'center' };
    const borderStyle = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

    // --- Sheet 1: Stoc Detaliat ---
    const worksheet1 = workbook.addWorksheet('Stoc Detaliat');
    worksheet1.columns = [{ header: 'Tip Material', key: 'tip', width: 30 }, { header: 'Cod Culoare / Descriere', key: 'cod', width: 35 }, { header: 'Cantitate Totală (tone)', key: 'cantitate', width: 25 }, { header: 'Status', key: 'status', width: 20 }];
    worksheet1.addRows(dateTabel);
    worksheet1.autoFilter = 'A1:D1';
    const headerRow1 = worksheet1.getRow(1);
    headerRow1.font = headerFont14;
    headerRow1.eachCell((cell) => { cell.fill = greenFill; cell.alignment = centerAlignment; cell.border = borderStyle; });
    worksheet1.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
            const statusVal = row.getCell('status').value;
            row.eachCell({ includeEmpty: true }, cell => {
                cell.font = statusVal === 'Stoc Redus' ? redBoldFont11 : defaultFont11;
                cell.border = borderStyle;
            });
        }
    });
    const legendRowIndex1 = worksheet1.lastRow.number + 2;
    worksheet1.mergeCells(`A${legendRowIndex1}:D${legendRowIndex1}`);
    const legendCell1 = worksheet1.getCell(`A${legendRowIndex1}`);
    legendCell1.value = legendValue; legendCell1.alignment = { wrapText: true, vertical: 'top' }; worksheet1.getRow(legendRowIndex1).height = 55;

    // --- Sheet 2: Stoc Materie Prima - Uz Extern ---
    const worksheet2 = workbook.addWorksheet('Stoc Materie Prima - Uz Extern');
    worksheet2.columns = [{ header: 'Tip Material', key: 'tip', width: 30 }, { header: 'Cod Culoare / Descriere', key: 'cod', width: 35 }, { header: 'Status', key: 'status', width: 20 }];
    worksheet2.addRows(dateTabel);
    worksheet2.autoFilter = 'A1:C1';
    const headerRow2 = worksheet2.getRow(1);
    headerRow2.font = headerFont14;
    headerRow2.eachCell((cell) => { cell.fill = greenFill; cell.alignment = centerAlignment; cell.border = borderStyle; });
    worksheet2.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {
            const statusVal = row.getCell('status').value;
            row.eachCell({ includeEmpty: true }, cell => {
                cell.font = statusVal === 'Stoc Redus' ? redBoldFont11 : defaultFont11;
                cell.border = borderStyle;
            });
        }
    });
    const legendRowIndex2 = worksheet2.lastRow.number + 2;
    worksheet2.mergeCells(`A${legendRowIndex2}:C${legendRowIndex2}`);
    const legendCell2 = worksheet2.getCell(`A${legendRowIndex2}`);
    legendCell2.value = legendValue; legendCell2.alignment = { wrapText: true, vertical: 'top' }; worksheet2.getRow(legendRowIndex2).height = 55;

    // --- Sheet 3: Stoc - UZ Extern- simplificat (RECONSTRUIT CORECT) ---
    const worksheet3 = workbook.addWorksheet('Stoc - UZ Extern- simplificat');
    const headerSimplificat = ['SUPREM', 'NEOMAT', 'MAT 0.50', 'MAT 0.45', 'MAT 0.40', 'LUCIOS 0.50', 'LUCIOS 0.45', 'LUCIOS 0.40', 'LUCIOS 0.35', 'LUCIOS 0.30', 'ZN', '> 0.50', 'IMITATIE LEMN'];
    const dataToHeaderMap = { 'MAT 0.5': 'MAT 0.50', 'MAT 0.4': 'MAT 0.40', '> 0.5': '> 0.50', 'LUCIOS 0.5': 'LUCIOS 0.50', 'LUCIOS 0.4': 'LUCIOS 0.40', 'LUCIOS 0.3': 'LUCIOS 0.30' };
    const groupMappings = { 'Imitatie Lemn SP': 'IMITATIE LEMN', 'Imitatie Lemn DP': 'IMITATIE LEMN' };
    const woodMapping = {
        'WOOD DARK WAL SP': 'LEMN NUC INCHIS Simplu Prevopsit',
        'WOOD GOLDEN OAK SP': 'LEMN STEJAR AURIU Simplu Prevopsit',
        'WOOD GOLDEN OAK DP': 'LEMN STEJAR AURIU Dublu Prevopsit'
    };

    worksheet3.columns = headerSimplificat.map(h => ({ header: h, key: h, width: 21.22 }));
    worksheet3.autoFilter = { from: 'A1', to: { row: 1, column: headerSimplificat.length } };
    
    const headerRow3 = worksheet3.getRow(1);
    headerRow3.font = { name: 'Calibri', size: 18, bold: true };
    headerRow3.eachCell(cell => { cell.fill = greenFill; cell.alignment = centerAlignment; cell.border = borderStyle; });
    headerRow3.height = 30;

    const groupedData = {};
    headerSimplificat.forEach(h => groupedData[h] = []);
    
    dateTabel.forEach(row => {
        const originalTip = row.tip; // Deja curățat cu trim()
        const group = groupMappings[originalTip] || dataToHeaderMap[originalTip] || originalTip;
        if (groupedData[group]) {
            groupedData[group].push({ cod: row.cod, status: row.status });
        }
    });

    // Sortăm listele de culori A-Z în interiorul fiecărui grup
    for (const key in groupedData) {
        groupedData[key].sort((a, b) => String(a.cod).localeCompare(String(b.cod)));
    }

    const maxRows = Math.max(0, ...Object.values(groupedData).map(arr => arr.length));
    const redFont18 = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FFFF0000' } };
    const blackFont18 = { name: 'Calibri', size: 18, bold: true, color: { argb: 'FF000000' } };
    
    // Creăm rânduri "dense" pentru a avea borduri complete
    for (let i = 0; i < maxRows; i++) {
        const rowData = {};
        for (const header of headerSimplificat) {
            const item = groupedData[header]?.[i];
            if (item) {
                if (header === 'IMITATIE LEMN' && woodMapping[item.cod]) {
                    rowData[header] = woodMapping[item.cod];
                } else {
                    rowData[header] = item.cod;
                }
            } else {
                rowData[header] = ''; // Adaugă celulă goală pentru a crea grila completă
            }
        }
        const addedRow = worksheet3.addRow(rowData);
        
        // Aplicăm stilurile pe TOATE celulele din rând (inclusiv cele goale)
        addedRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const headerName = worksheet3.getColumn(colNumber).header;
            const item = groupedData[headerName]?.[i];
            
            cell.alignment = centerAlignment;
            cell.border = borderStyle;

            if (item) {
                cell.font = item.status === 'Stoc Redus' ? redFont18 : blackFont18;
            } else {
                cell.font = blackFont18; // Aplicăm stilul de bază și celulelor goale
                cell.value = ''; // Asigură-te că sunt goale, nu null
            }
        });
    }

    // Adăugăm legenda la final
    const legendRowIndex3 = (worksheet3.lastRow ? worksheet3.lastRow.number : 1) + 2;
    worksheet3.mergeCells(legendRowIndex3, 1, legendRowIndex3, headerSimplificat.length);
    const legendCell3 = worksheet3.getCell(legendRowIndex3, 1);
    legendCell3.value = legendValue;
    legendCell3.alignment = { wrapText: true, vertical: 'top' };
    worksheet3.getRow(legendRowIndex3).height = 55;

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}
async function generatePdfReport(rezultateStoc) { const dateTabel = Object.keys(rezultateStoc).map(cheie => { const [tipMaterial, codCuloare] = cheie.split('|'); const cantitate = rezultateStoc[cheie]; return { tip: tipMaterial, cod: codCuloare, status: cantitate >= 10 ? 'Stoc Suficient' : 'Stoc Redus', cantitate: cantitate }; }).filter(row => row.cantitate >= 1).sort((a, b) => a.tip.localeCompare(b.tip)); let htmlRows = ''; dateTabel.forEach(row => { const color = row.status === 'Stoc Redus' ? 'red' : 'black'; const fontWeight = 'bold'; htmlRows += `<tr style="color: ${color}; font-weight: ${fontWeight};"><td>${row.tip}</td><td>${row.cod}</td><td>${row.status}</td></tr>`; }); const legendaHtml = `<div class="legenda"><p><strong>* ≥10 tone: Stoc Suficient ⚫</strong></p><p style="color: red;"><strong>* 1-10 tone: Stoc Redus 🔴</strong></p><p style="color: red;"><strong>* <1 tonă: Nu se afișează în acest tabel ❌</strong></p></div>`; const dataCurenta = new Date().toLocaleDateString('ro-RO'); const htmlContent = `<html><head><style>body { font-family: Calibri, sans-serif; } table { width: 100%; border-collapse: collapse; page-break-inside: auto; } tr { page-break-inside: avoid; page-break-after: auto; } thead { display: table-header-group; } th, td { border: 1px solid #cccccc; padding: 8px; text-align: left; } th { background-color: #C6E0B4; font-size: 14px; font-weight: bold; } h1 { font-size: 20px; text-align: center; margin-bottom: 20px; } .legenda { margin-top: 30px; border-top: 1px solid #ccc; padding-top: 15px; page-break-inside: avoid; } .legenda p { margin: 2px 0; }</style></head><body><h1>Stoc Materie Prima - Uz Extern [${dataCurenta}]</h1><table><thead><tr><th>Tip Material</th><th>Cod Culoare / Descriere</th><th>Status</th></tr></thead><tbody>${htmlRows}</tbody></table>${legendaHtml}</body></html>`; const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] }); const page = await browser.newPage(); await page.setContent(htmlContent, { waitUntil: 'networkidle0' }); const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true, landscape: true, margin: { top: '25px', right: '25px', bottom: '25px', left: '25px' } }); await browser.close(); return pdfBuffer; }