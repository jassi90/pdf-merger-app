require('dotenv').config();

const express = require('express');
const { runAutomation: runDemographicAutomation } = require('./automation');
const { runAutomation: runPart1Automation } = require('./part1-automation');
const { runAutomation: runPart2Automation } = require('./part2-automation');

const app = express();
app.use(express.json({ limit: '1mb' }));

function logLocalUsage(routeName, systemId, req) {
  const timestamp = new Date().toISOString();
  const clientIp =
    req.headers['x-forwarded-for'] ||
    req.socket?.remoteAddress ||
    'unknown';

  console.log(
    `-----------------------------------------------------
    [LOCAL USAGE] ${timestamp} 
    route= ${routeName} 
    systemId= ${systemId ?? 'n/a'} 
    ip= ${clientIp}
    -------------------------------------------------------`
  
  );
}

app.post('/run-local-automation', async (req, res) => {
  try {
    const { systemId } = req.body;

    if (!systemId) {
      return res.status(400).send('systemId is required');
    }

    logLocalUsage('/run-local-automation', systemId, req);

    const pdfBytes = await runDemographicAutomation(systemId);

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=generated.pdf');
    res.send(Buffer.from(pdfBytes));
  } catch (err) {
    console.error('LOCAL AUTOMATION ERROR:', err);
    res.status(500).send(err.message || 'Local automation failed');
  }
});

app.post('/run-local-part2', async (req, res) => {
  try {
    const { systemId } = req.body || {};
    const parsedSystemId = Number(systemId);

    if (!Number.isFinite(parsedSystemId)) {
      return res.status(400).json({ error: 'systemId is required for Part2' });
    }

    logLocalUsage('/run-local-part2', parsedSystemId, req);

    const result = await runPart2Automation(parsedSystemId);
    res.json(result);
  } catch (err) {
    console.error('LOCAL PART2 ERROR:', err);
    res.status(500).json({
      error: err.message || 'Local Part2 automation failed'
    });
  }
});

app.post('/run-local-part1', async (req, res) => {
  try {
    const { systemId } = req.body || {};
    const parsedSystemId = Number(systemId);

    if (!Number.isFinite(parsedSystemId)) {
      return res.status(400).json({ error: 'systemId is required for Part1' });
    }

    logLocalUsage('/run-local-part1', parsedSystemId, req);

    const result = await runPart1Automation(parsedSystemId);
    res.json(result);
  } catch (err) {
    console.error('LOCAL PART1 ERROR:', err);
    res.status(500).json({
      error: err.message || 'Local Part1 automation failed'
    });
  }
});

const PORT = 4000;
app.listen(PORT, () => {
  console.log(`Local automation running at http://localhost:${PORT}`);
});
