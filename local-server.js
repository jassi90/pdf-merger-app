require('dotenv').config();

const express = require('express');
const { runAutomation } = require('./automation');

const app = express();
app.use(express.json({ limit: '1mb' }));

app.post('/run-local-automation', async (req, res) => {
  try {
    const { systemId } = req.body;

    if (!systemId) {
      return res.status(400).send('systemId is required');
    }

    console.log('Received systemId:', systemId);

    const pdfBytes = await runAutomation(systemId);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=generated.pdf');
    res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error('LOCAL AUTOMATION ERROR:', err);
    res.status(500).send(err.message || 'Local automation failed');
  }
});

const PORT = 4000;
app.listen(PORT, () => {
  console.log(`Local automation running at http://localhost:${PORT}`);
});