require('dotenv').config();

const express = require('express');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');
const { runPart1DataEntry, DEFAULT_SYSTEM_IDS } = require('./part1-automation');

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
                page.drawImage(image, {
                    x: 0, y: 0,
                    width: image.width,
                    height: image.height
                });

            } else if (ext === '.png') {
                const image = await mergedPdf.embedPng(bytes);
                const page = mergedPdf.addPage([image.width, image.height]);
                page.drawImage(image, {
                    x: 0, y: 0,
                    width: image.width,
                    height: image.height
                });

            } else {
                throw new Error(`Unsupported file: ${file.originalname}`);
            }

            fs.unlinkSync(filePath);
        }

        const mergedBytes = await mergedPdf.save();

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=merged.pdf');
        res.send(Buffer.from(mergedBytes));

    } catch (err) {
        console.error("MERGE ERROR:", err);
        res.status(500).send(err.message || 'Error merging files');
    }
});

/* ---------------- PLAYWRIGHT VIA NGROK ---------------- */
app.post('/generate-pdf', async (req, res) => {
    try {
        const { systemId } = req.body;

        if (!systemId) {
            return res.status(400).send('systemId is required');
        }

        console.log("Forwarding to local automation:", systemId);

        const response = await fetch('https://expertly-uncritical-annmarie.ngrok-free.dev/run-local-automation', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ systemId })
        });

        if (!response.ok) {
            const text = await response.text();
            throw new Error(text || 'Local automation failed');
        }

        const arrayBuffer = await response.arrayBuffer();

        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=generated.pdf');
        res.send(Buffer.from(arrayBuffer));

    } catch (err) {
        console.error("REMOTE ERROR:", err);
        res.status(500).send(err.message || 'Failed to connect to local automation');
    }
});

/* ---------------- START SERVER ---------------- */
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});

/* ---------------- PART1 DATA ENTRY ---------------- */
app.post('/run-part1-data-entry', async (req, res) => {
    try {
        const { systemIds, headless } = req.body || {};

        const summary = await runPart1DataEntry({
            systemIds: Array.isArray(systemIds) && systemIds.length > 0 ? systemIds : DEFAULT_SYSTEM_IDS,
            headless: Boolean(headless)
        });

        res.json(summary);
    } catch (err) {
        console.error('PART1 ERROR:', err);
        res.status(500).json({
            error: err.message || 'Failed to run Part1 data entry automation'
        });
    }
});
