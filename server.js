require('dotenv').config();

const express = require('express');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');
const { runAutomation } = require('./automation');

const app = express();

app.use(express.static(__dirname));
app.use(express.json());

const upload = multer({ dest: 'uploads/' });

/* ---------------- PDF MERGER ---------------- */
app.post('/merge', upload.array('files'), async (req, res) => {
    try {
        const mergedPdf = await PDFDocument.create();

        for (const file of req.files) {
            const filePath = file.path;
            const ext = path.extname(file.originalname).toLowerCase();
            const bytes = fs.readFileSync(filePath);

            if (ext === '.pdf') {
                const pdf = await PDFDocument.load(bytes);
                const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                pages.forEach(p => mergedPdf.addPage(p));
            } else if (ext === '.jpg' || ext === '.jpeg') {
                const image = await mergedPdf.embedJpg(bytes);
                const page = mergedPdf.addPage([image.width, image.height]);
                page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
            } else if (ext === '.png') {
                const image = await mergedPdf.embedPng(bytes);
                const page = mergedPdf.addPage([image.width, image.height]);
                page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
            }

            fs.unlinkSync(filePath);
        }

        const mergedBytes = await mergedPdf.save();

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=merged.pdf');
        res.send(Buffer.from(mergedBytes));

    } catch (err) {
        console.error("MERGE ERROR:", err);
        res.status(500).send('Error merging files');
    }
});

/* ---------------- PLAYWRIGHT AUTOMATION ---------------- */
app.post('/generate-pdf', async (req, res) => {
    try {
        const { systemId } = req.body;

        const pdfBytes = await runAutomation(systemId);

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=generated.pdf');
        res.send(Buffer.from(pdfBytes));

    } catch (err) {
        console.error("AUTOMATION ERROR:", err);
        res.status(500).send(err.message);
    }
});

/* ---------------- START SERVER ---------------- */
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});