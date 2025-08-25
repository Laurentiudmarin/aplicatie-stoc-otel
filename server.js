// --- server.js - COD COMPLET FINAL (cu Data și Legendă în PDF) ---

const express = require('express');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');
const sqlite3 = require('sqlite3').verbose();
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

const app = express();
app.use(express.json());
const PORT = 3000;
const upload = multer({ dest: 'uploads/' });
const db = new sqlite3.Database('./database.sqlite');

// ... (Tot codul de la început până la funcțiile de generare Excel/PDF rămâne neschimbat) ...
db.serialize(() => { db.run(`CREATE TABLE IF NOT EXISTS reguli (id INTEGER PRIMARY KEY AUTOINCREMENT, furnizor TEXT NOT NULL, criterii TEXT NOT NULL, tip_material TEXT NOT NULL, descriere_raport TEXT NOT NULL)`); });
app.use(express.static('public'));
app.get('/api/reguli', (req, res) => { db.all("SELECT * FROM reguli ORDER BY tip_material ASC", [], (err, rows) => { if (err) { res.status(500).json({ error: err.message }); return; } res.json(rows); }); });
app.post('/api/reguli', (req, res) => { const { furnizor, criterii, tip_material, descriere_raport } = req.body; const sql = `INSERT INTO reguli (furnizor, criterii, tip_material, descriere_raport) VALUES (?, ?, ?, ?)`; db.run(sql, [furnizor, criterii, tip_material, descriere_raport], function(err) { if (err) { res.status(400).json({ error: err.message }); return; } res.json({ id: this.lastID }); }); });
app.delete('/api/reguli/:id', (req, res) => { db.run('DELETE FROM reguli WHERE id = ?', req.params.id, function(err) { if (err) { res.status(400).json({ error: err.message }); return; } res.json({ changes: this.changes }); }); });
app.put('/api/reguli/:id', (req, res) => { const { furnizor, criterii, tip_material, descriere_raport } = req.body; const sql = `UPDATE reguli SET furnizor = ?, criterii = ?, tip_material = ?, descriere_raport = ? WHERE id = ?`; const params = [furnizor, criterii, tip_material, descriere_raport, req.params.id]; db.run(sql, params, function(err) { if (err) { res.status(400).json({ error: err.message }); return; } res.json({ message: "Succes", changes: this.changes }); }); });
app.post('/api/get-suppliers', upload.single('stoc'), (req, res) => { const stocFile = req.file; try { const workbook = xlsx.readFile(stocFile.path); const sheetName = workbook.SheetNames.find(name => name.includes('800')); if (!sheetName) { throw new Error("Nu am găsit niciun sheet care să conțină '800' în nume."); } const sheet = workbook.Sheets[sheetName]; const stocData = xlsx.utils.sheet_to_json(sheet); const suppliers = new Set(); for (const rand of stocData) { const supplierName = rand['Name 1'] || rand['Nume fz']; if (supplierName) { suppliers.add(supplierName.toString().trim()); } } res.json(Array.from(suppliers).sort()); } catch (error) { console.error('Eroare la extragerea furnizorilor:', error); res.status(500).send(error.message); } finally { if (stocFile) fs.unlinkSync(stocFile.path); } });
async function runProcessing(stocFilePath, selectedSuppliers) { return new Promise((resolve, reject) => { db.all("SELECT * FROM reguli", [], (err, reguli) => { if (err) return reject(err); const workbook = xlsx.readFile(stocFilePath); const sheetName = workbook.SheetNames.find(name => name.includes('800')); if (!sheetName) { return reject(new Error("Nu am găsit niciun sheet care să conțină '800' în nume.")); } const sheet = workbook.Sheets[sheetName]; const stocData = xlsx.utils.sheet_to_json(sheet); const filteredStocData = stocData.filter(rand => { const supplierName = rand['Name 1'] || rand['Nume fz']; return selectedSuppliers.includes(supplierName); }); const cantitatiFinale = {}; for (const rand of filteredStocData) { const descriere = rand['Material Description'] || ''; const furnizor = rand['Name 1'] || rand['Nume fz'] || ''; const cantitate = parseFloat(rand['Unrestricted'] || rand['Stoc (to)']) || 0; if (!descriere || cantitate <= 0) continue; let regulaPotrivita = null; for (const regula of reguli) { const furnizorMatch = (regula.furnizor.toUpperCase() === 'ORICE' || regula.furnizor.toLowerCase() === furnizor.toLowerCase()); const criterii = regula.criterii.split(',').map(c => c.trim().toLowerCase()).filter(c => c); const descriereLower = descriere.toLowerCase(); const criteriiMatch = criterii.length > 0 && criterii.every(c => descriereLower.includes(c)); if (furnizorMatch && criteriiMatch) { regulaPotrivita = regula; break; } } if (regulaPotrivita) { const cheieUnica = `${regulaPotrivita.tip_material}|${regulaPotrivita.descriere_raport}`; cantitatiFinale[cheieUnica] = (cantitatiFinale[cheieUnica] || 0) + cantitate; } } resolve(cantitatiFinale); }); }); }
app.post('/process', upload.single('stoc'), async (req, res) => { const stocFile = req.file; try { const selectedSuppliers = JSON.parse(req.body.suppliers); const rezultateStoc = await runProcessing(stocFile.path, selectedSuppliers); const rezultateFormatate = Object.keys(rezultateStoc).map(cheie => { const [tipMaterial, descriereRaport] = cheie.split('|'); return { tipMaterial, descriereRaport, cantitate: rezultateStoc[cheie] }; }).filter(item => item.cantitate >= 1); res.json(rezultateFormatate); } catch (error) { console.error('Eroare la /process:', error); res.status(500).send(error.message); } finally { if (stocFile) fs.unlinkSync(stocFile.path); } });
app.post('/download', upload.single('stoc'), async (req, res) => { const stocFile = req.file; try { const selectedSuppliers = JSON.parse(req.body.suppliers); const rezultateStoc = await runProcessing(stocFile.path, selectedSuppliers); const excelBuffer = await generateExcelReport(rezultateStoc); const numeFisier = `Stoc_Materie_Prima_${new Date().toLocaleDateString('ro-RO').replace(/\./g, '-')}.xlsx`; res.setHeader('Content-Disposition', `attachment; filename="${numeFisier}"`); res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); res.send(excelBuffer); } catch (error) { console.error('Eroare la /download:', error); res.status(500).send(error.message); } finally { if (stocFile) fs.unlinkSync(stocFile.path); } });
app.post('/download-pdf', upload.single('stoc'), async (req, res) => { const stocFile = req.file; try { const selectedSuppliers = JSON.parse(req.body.suppliers); const rezultateStoc = await runProcessing(stocFile.path, selectedSuppliers); const pdfBuffer = await generatePdfReport(rezultateStoc); const numeFisier = `Stoc_materie_prima-uz_extern-${new Date().toLocaleDateString('ro-RO').replace(/\./g, '-')}.pdf`; res.setHeader('Content-Disposition', `attachment; filename="${numeFisier}"`); res.setHeader('Content-Type', 'application/pdf'); res.send(pdfBuffer); } catch (error) { console.error('Eroare la /download-pdf:', error); res.status(500).send(error.message); } finally { if (stocFile) fs.unlinkSync(stocFile.path); } });
app.listen(PORT, () => { console.log(`Serverul FINAL a pornit la http://localhost:${PORT}`); });

// --- FUNCȚIILE DE GENERARE EXCEL (neschimbate) ---
async function generateExcelReport(rezultateStoc) {
    const workbook = new ExcelJS.Workbook();
    const dateTabel = Object.keys(rezultateStoc).map(cheie => { const [tipMaterial, codCuloare] = cheie.split('|'); const cantitate = rezultateStoc[cheie]; return { tip: tipMaterial, cod: codCuloare, cantitate: parseFloat(cantitate.toFixed(3)), status: cantitate >= 10 ? 'Stoc Suficient' : 'Stoc Redus' }; }).filter(row => row.cantitate >= 1).sort((a, b) => a.tip.localeCompare(b.tip));
    const worksheet1 = workbook.addWorksheet('Stoc Detaliat');
    worksheet1.columns = [{ header: 'Tip Material', key: 'tip', width: 30 }, { header: 'Cod Culoare / Descriere', key: 'cod', width: 35 }, { header: 'Cantitate Totală (tone)', key: 'cantitate', width: 25 }, { header: 'Status', key: 'status', width: 20 }];
    worksheet1.addRows(dateTabel);
    worksheet1.autoFilter = 'A1:D1';
    const headerRow1 = worksheet1.getRow(1);
    headerRow1.font = { name: 'Calibri', size: 14, bold: true };
    headerRow1.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }; cell.alignment = { vertical: 'middle', horizontal: 'center' }; });
    const redBoldFont = { name: 'Calibri', bold: true, color: { argb: 'FFFF0000' } };
    worksheet1.eachRow((row, rowNumber) => { if (rowNumber > 1) { if (row.getCell('status').value === 'Stoc Redus') { row.getCell('cod').font = redBoldFont; row.getCell('cantitate').font = redBoldFont; row.getCell('status').font = redBoldFont; } } });
    const legendRowIndex1 = worksheet1.lastRow.number + 2;
    worksheet1.mergeCells(`A${legendRowIndex1}:D${legendRowIndex1}`);
    const legendCell1 = worksheet1.getCell(`A${legendRowIndex1}`);
    legendCell1.value = { richText: [{ font: { bold: true, color: { argb: 'FF000000' } }, text: '* ≥10 tone: Stoc Suficient ⚫\n' }, { font: { bold: true, color: { argb: 'FFFF0000' } }, text: '* 1-10 tone: Stoc Redus 🔴\n' }, { font: { bold: true, color: { argb: 'FFFF0000' } }, text: '* <1 tonă: Nu se afișează în acest tabel ❌' }] };
    legendCell1.alignment = { wrapText: true, vertical: 'top' };
    worksheet1.getRow(legendRowIndex1).height = 55;
    const worksheet2 = workbook.addWorksheet('Stoc Materie Prima - Uz Extern');
    worksheet2.columns = [{ header: 'Tip Material', key: 'tip', width: 30 }, { header: 'Cod Culoare / Descriere', key: 'cod', width: 35 }, { header: 'Status', key: 'status', width: 20 }];
    worksheet2.addRows(dateTabel);
    worksheet2.autoFilter = 'A1:C1';
    const headerRow2 = worksheet2.getRow(1);
    headerRow2.font = { name: 'Calibri', size: 14, bold: true };
    headerRow2.eachCell((cell) => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6E0B4' } }; cell.alignment = { vertical: 'middle', horizontal: 'center' }; });
    worksheet2.eachRow((row, rowNumber) => { if (rowNumber > 1) { if (row.getCell('status').value === 'Stoc Redus') { row.getCell('cod').font = redBoldFont; row.getCell('status').font = redBoldFont; } } });
    const legendRowIndex2 = worksheet2.lastRow.number + 2;
    worksheet2.mergeCells(`A${legendRowIndex2}:C${legendRowIndex2}`);
    const legendCell2 = worksheet2.getCell(`A${legendRowIndex2}`);
    legendCell2.value = legendCell1.value;
    legendCell2.alignment = { wrapText: true, vertical: 'top' };
    worksheet2.getRow(legendRowIndex2).height = 55;
    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}


// --- FUNCȚIA DE GENERARE PDF (MODIFICATĂ) ---
async function generatePdfReport(rezultateStoc) {
    const dateTabel = Object.keys(rezultateStoc)
        .map(cheie => {
            const [tipMaterial, codCuloare] = cheie.split('|');
            const cantitate = rezultateStoc[cheie];
            return {
                tip: tipMaterial,
                cod: codCuloare,
                status: cantitate >= 10 ? 'Stoc Suficient' : 'Stoc Redus',
                cantitate: cantitate
            };
        })
        .filter(row => row.cantitate >= 1)
        .sort((a, b) => a.tip.localeCompare(b.tip));

    let htmlRows = '';
    dateTabel.forEach(row => {
        const color = row.status === 'Stoc Redus' ? 'red' : 'black';
        const fontWeight = row.status === 'Stoc Redus' ? 'bold' : 'normal';
        htmlRows += `<tr style="color: ${color}; font-weight: ${fontWeight};"><td>${row.tip}</td><td>${row.cod}</td><td>${row.status}</td></tr>`;
    });
    
    // NOU: Am creat un element HTML pentru legendă
    const legendaHtml = `
        <div class="legenda">
            <p><strong>* ≥10 tone: Stoc Suficient ⚫</strong></p>
            <p style="color: red;"><strong>* 1-10 tone: Stoc Redus 🔴</strong></p>
            <p style="color: red;"><strong>* <1 tonă: Nu se afișează în acest tabel ❌</strong></p>
        </div>
    `;

    // NOU: Am creat data curentă
    const dataCurenta = new Date().toLocaleDateString('ro-RO');

    const htmlContent = `
        <html>
            <head>
                <style>
                    body { font-family: Calibri, sans-serif; }
                    table { width: 100%; border-collapse: collapse; page-break-inside: auto; }
                    tr { page-break-inside: avoid; page-break-after: auto; }
                    thead { display: table-header-group; }
                    th, td { border: 1px solid #cccccc; padding: 8px; text-align: left; }
                    th { background-color: #C6E0B4; font-size: 14px; font-weight: bold; }
                    h1 { font-size: 20px; text-align: center; margin-bottom: 20px; }
                    .legenda { margin-top: 30px; border-top: 1px solid #ccc; padding-top: 15px; page-break-inside: avoid; }
                    .legenda p { margin: 2px 0; }
                </style>
            </head>
            <body>
                <h1>Stoc Materie Prima - Uz Extern [${dataCurenta}]</h1>
                <table>
                    <thead><tr><th>Tip Material</th><th>Cod Culoare / Descriere</th><th>Status</th></tr></thead>
                    <tbody>${htmlRows}</tbody>
                </table>
                ${legendaHtml}
            </body>
        </html>`;

    const browser = await puppeteer.launch({ args: ['--no-sandbox'] });
    const page = await browser.newPage();
    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
    const pdfBuffer = await page.pdf({ format: 'A4', printBackground: true, landscape: true, margin: { top: '25px', right: '25px', bottom: '25px', left: '25px' } });
    await browser.close();
    return pdfBuffer;
}