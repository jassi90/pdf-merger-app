const express = require('express');
const multer = require('multer');
const { PDFDocument } = require('pdf-lib');
const fs = require('fs');
const path = require('path');

const app = express();

// Serve static files (index.html)
app.use(express.static(__dirname));

// Configure multer to save uploaded files in "uploads" folder
const upload = multer({ dest: 'uploads/' });

// Home route — serves index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Merge route — handles uploaded files and combines them
app.post('/merge', upload.array('files'), async (req, res) => {
    try {
        const mergedPdf = await PDFDocument.create();

        for (const file of req.files) {
            const filePath = file.path;
            const ext = path.extname(file.originalname).toLowerCase();
            const bytes = fs.readFileSync(filePath);

            if (ext === '.pdf') {
                // Merge PDF pages
                const pdf = await PDFDocument.load(bytes);
                const pages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                pages.forEach(p => mergedPdf.addPage(p));
            } else if (ext === '.jpg' || ext === '.jpeg') {
                // Embed JPG image
                const image = await mergedPdf.embedJpg(bytes);
                const page = mergedPdf.addPage([image.width, image.height]);
                page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
            } else if (ext === '.png') {
                // Embed PNG image
                const image = await mergedPdf.embedPng(bytes);
                const page = mergedPdf.addPage([image.width, image.height]);
                page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
            } else {
                console.log(`Skipped unsupported file: ${file.originalname}`);
            }

            // Delete uploaded file after processing
            fs.unlinkSync(filePath);
        }

        // Save merged PDF to buffer
        const mergedBytes = await mergedPdf.save();

        // Send merged PDF to browser for download
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', 'attachment; filename=merged.pdf');
        res.send(Buffer.from(mergedBytes));
    } catch (err) {
        console.error(err);
        res.status(500).send('Error merging files');
    }
});

// Start server
app.listen(3000, () => {
    console.log('Server running at http://localhost:3000');
});