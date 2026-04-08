const { chromium } = require('playwright');
const ExcelJs = require('exceljs');
const { PDFDocument } = require('pdf-lib');
const pdfParse = require('pdf-parse');
const fs = require('fs/promises');
const path = require('path');
const { spawn } = require('child_process');

function formatDate(dateString) {
  const date = new Date(dateString);
  if (isNaN(date.getTime())) return dateString;
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const year = date.getFullYear();
  return `${month}-${day}-${year}`;
}

async function loginIllinois(page) {
  await page.goto('https://portal.illinoisabp.com/');
  await page.getByLabel('Username').fill(process.env.ILLINOIS_USERNAME);
  await page.getByLabel('Password').fill(process.env.ILLINOIS_PASSWORD);
  await page.getByRole('button', { name: 'Sign in' }).first().click();
  await page.waitForTimeout(2000);
}

async function loginCSG(page) {
  await page.goto('https://portal2.carbonsolutionsgroup.com/admin/login');
  await page.fill('input[type="email"]', process.env.EMAIL);
  await page.fill('input[type="password"]', process.env.PASSWORD);
  await page.getByRole('button', { name: 'Login' }).click();
  await page.waitForTimeout(2000);
}

async function loadInstallerInfo(installerInfoPath) {
  const workbook = new ExcelJs.Workbook();
  await workbook.xlsx.readFile(installerInfoPath);
  return workbook.getWorksheet('Sheet1');
}

function findInstallerRow(worksheet, docketValue) {
  let output = { row: -1, column: -1 };

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      if (cell.value === docketValue) {
        output = { row: rowNumber, column: colNumber };
      }
    });
  });

  if (output.row === -1) {
    throw new Error(`Installer "${docketValue}" not found in Excel file.`);
  }

  return {
    I1: worksheet.getCell(output.row, 1).value,
    I2: worksheet.getCell(output.row, 2).value,
    I3: worksheet.getCell(output.row, 3).value,
    I4: worksheet.getCell(output.row, 4).value,
    I5: worksheet.getCell(output.row, 5).value,
    I6: worksheet.getCell(output.row, 6).value,
    I7: worksheet.getCell(output.row, 7).value,
    I8: worksheet.getCell(output.row, 8).value
  };
}

function determineNetMetering(worksheet, utility) {
  let nm = 'Yes';
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (cell.value === utility) nm = 'No';
    });
  });
  return nm;
}

async function getChecklistData(page, systemId, worksheet) {
  await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/edit?step=2.2`);
  await page.getByRole('button', { name: 'Save', exact: true }).click();
  await page.waitForTimeout(2000);

  await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist?onlyIncompleTasks=false&onlyMyTasks=false&showAll=false`);

  const ABPID = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
  const AC = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText();
  const systemName = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
  const Utility = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(11) > td:nth-child(2)').innerText();
  const installerPart1 = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(12) > td:nth-child(2)').innerText();

  const interconnectionDate = await page.getByRole('row', { name: 'Interconnection Approval Date' }).getByRole('cell').nth(1).innerText();
  const projectOnlineDate = await page.getByRole('row', { name: 'Project Online Date' }).getByRole('cell').nth(1).innerText();
  const dateOfProject = await page.getByRole('row', { name: "Date of Project's Certificate" }).getByRole('cell').nth(1).innerText();
  const completationDate = await page.getByRole('row', { name: 'Construction Completion Date' }).getByRole('cell').nth(1).innerText();

  let NON = await page.getByRole('row', { name: 'PJM Gats or MRETs Unit ID' }).getByRole('cell').nth(1).innerText();
  if (!NON || NON === 'MISSING!') {
    NON = 'NON123456';
  }

  const I9 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
  const I10 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(12) > td:nth-child(2)').innerText();

  const installer = findInstallerRow(worksheet, I10);
  const NM = determineNetMetering(worksheet, Utility);

  let I3 = installer.I3;
  if (I3 == null) {
    I3 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText();
  }

  const WDate = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(13) > td:nth-child(2)').innerText();
  const I11 = formatDate(WDate);

  const demographics = {
    D1: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(16) > td:nth-child(2)').innerText(),
    D2: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(17) > td:nth-child(2)').innerText(),
    D3: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(18) > td:nth-child(2)').innerText(),
    D4: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(19) > td:nth-child(2)').innerText(),
    D5: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(20) > td:nth-child(2)').innerText(),
    D6: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(21) > td:nth-child(2)').innerText(),
    D7: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(22) > td:nth-child(2)').innerText(),
    D8: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(23) > td:nth-child(2)').innerText(),
    D9: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(26) > td:nth-child(2)').innerText(),
    D10: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(27) > td:nth-child(2)').innerText(),
    D11: await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(28) > td:nth-child(2)').innerText()
  };

  const tableSelector = '#part2-section3 > fieldset > table > tbody > tr:nth-child(31) > td:nth-child(2) > table';
  const rows = await page.$$(tableSelector + ' tr');
  const laborData = [];

  for (let i = 0; i < rows.length; i++) {
    const cells = await rows[i].$$('td');
    if (cells.length >= 2) {
      laborData.push({
        zip: (await cells[0].innerText()).trim(),
        hours: (await cells[1].innerText()).trim()
      });
    }
  }

  const inverterTable = '#part2-section4 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2) > table';
  const inverterRows = await page.$$(inverterTable + ' tr');
  const inverterData = [];

  for (let i = 0; i < inverterRows.length; i++) {
    const cells = await inverterRows[i].$$('td');
    if (cells.length >= 4) {
      inverterData.push({
        model: (await cells[1].innerText()).trim(),
        size: (await cells[2].innerText()).trim(),
        quantity: (await cells[3].innerText()).trim()
      });
    }
  }

  const inverterModels = inverterData
    .slice(1)
    .map((x) => `${x.model} = ${x.size} x ${x.quantity}`)
    .join(', ');

  const J1 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(33) > td:nth-child(2)').innerText();
  const J2 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(34) > td:nth-child(2)').innerText();
  const J3 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(35) > td:nth-child(2)').innerText();
  const J4 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(37) > td:nth-child(2)').innerText();

  const E1 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
  const E2 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
  const E3 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2) > table > tbody > tr.bg-white > td:nth-child(1)').first().innerText();
  const E5 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(7) > td:nth-child(2)').innerText();
  const E6 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
  const E7 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
  const Amount = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(11) > td:nth-child(2)').innerText();
  const E8 = Amount.replace(/\$|,/g, '');

  return {
    systemId,
    ABPID,
    AC: Number(AC),
    Utility,
    systemName,
    installerPart1,
    interconnectionDate,
    projectOnlineDate,
    dateOfProject,
    completationDate,
    NON,
    I9,
    I10,
    I11,
    installer: {
      I1: String(installer.I1 ?? ''),
      I2: String(installer.I2 ?? ''),
      I3: String(I3 ?? ''),
      I4: String(installer.I4 ?? ''),
      I5: String(installer.I5 ?? ''),
      I6: String(installer.I6 ?? ''),
      I7: String(installer.I7 ?? ''),
      I8: String(installer.I8 ?? '')
    },
    NM,
    demographics,
    laborData,
    inverterModels,
    J1,
    J2,
    J3,
    J4,
    E1,
    E2,
    E3,
    E5,
    E6,
    E7,
    E8
  };
}

async function revisitIfNeeded(page) {
  if (await page.getByRole('button', { name: 'Save and Continue' }).isVisible()) {
    const combo = page.getByRole('combobox').first();
    if (await combo.isVisible()) {
      await combo.selectOption('No');
    }
  }

  if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
    await page.getByRole('button', { name: 'Revisit' }).click();
    await page.waitForTimeout(2000);
    await page.getByRole('button', { name: 'OK' }).last().click();
    await page.waitForTimeout(2000);
  }
}

async function openIllinoisApplication(page, ABPID) {
  await page.goto('https://portal.illinoisabp.com/');
  await page.getByRole('button', { name: 'View Project Applications' }).click();
  await page.getByRole('button', { name: 'Clear Filters' }).click();
  await page.waitForTimeout(5000);

  await page.getByRole('columnheader', { name: 'sort Project Application ID' }).getByLabel('filter', { exact: true }).click();
  await page.getByRole('columnheader', { name: 'sort Project Application ID' }).getByLabel('filter', { exact: true }).fill(ABPID);

  await page.waitForTimeout(5000);
  await page.locator('.mx-name-actionButton13').click();
  await page.getByRole('button', { name: 'Section 1 - Project Details' }).click();
  await page.waitForTimeout(4000);
}

async function fillIllinoisForms(page, data) {
  await revisitIfNeeded(page);
  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(2000);
  await page.getByRole('button', { name: 'OK' }).last().click();

  await revisitIfNeeded(page);

  await page.getByLabel('Interconnection Approval Date').fill(data.interconnectionDate);
  await page.getByLabel('Project Online Date *').fill(data.projectOnlineDate);
  await page.getByLabel('Date of Project’s Certificate').fill(data.dateOfProject);
  await page.getByLabel('Date on which Construction').fill(data.completationDate);
  await page.getByLabel('REC Tracking System *').selectOption('GATS');
  await page.getByLabel('PJM GATS or M-RETS Unit ID *').fill(data.NON);
  await page.getByLabel('Name on REC Tracking System').fill('Carbon Solutions SREC LLC');
  await page.getByRole('combobox').nth(2).selectOption(data.NM);
  await page.getByRole('button', { name: 'Save and Continue' }).click();

  await revisitIfNeeded(page);

  await page.getByLabel('Legal Business Name *').fill(data.installer.I1);
  await page.getByLabel('Street *').fill(data.installer.I2);
  await page.getByLabel('Apartment or Suite').fill(data.installer.I3);
  await page.getByLabel('City *').fill(data.installer.I4);
  await page.getByLabel('State *').fill(data.installer.I5);
  await page.getByLabel('Zip Code *').fill(data.installer.I6);
  await page.getByLabel('Phone *').fill(data.installer.I7);
  await page.getByLabel('Email *').fill(data.installer.I8);
  await page.getByLabel('Name of the Qualified Person').fill(data.I9);
  await page.getByLabel('ICC Docket Number for the').fill(data.I10);
  await page.getByRole('combobox').nth(1).selectOption('No');
  await page.getByPlaceholder('mm-dd-yyyy').fill(data.I11);

  const d = data.demographics;

  if (d.D1 !== '0') await page.getByLabel('White').fill(d.D1);
  if (d.D2 !== '0') await page.getByLabel('Black or African American').fill(d.D2);
  if (d.D3 !== '0') await page.getByLabel('American Indian or Alaskan').fill(d.D3);
  if (d.D4 !== '0') await page.getByLabel('Asian').fill(d.D4);
  if (d.D5 !== '0') await page.getByLabel('Hawaiian or Other Pacific').fill(d.D5);
  if (d.D6 !== '0') await page.getByLabel('More than one Race').fill(d.D6);
  if (d.D7 !== '0') await page.getByLabel('Some Other Race').fill(d.D7);
  if (d.D8 !== '0') await page.getByLabel('Employee Declines to Identify', { exact: true }).fill(d.D8);

  await page.getByLabel('Hispanic or Latino', { exact: true }).fill(d.D9);
  await page.getByLabel('Not Hispanic or Latino').fill(d.D10);
  await page.getByLabel('Employee Declines to Identify Ethnicity').fill(d.D11);

  if (d.D9 === '0') {
    await page.getByLabel('Hispanic or Latino', { exact: true }).click();
    await page.getByLabel('Hispanic or Latino', { exact: true }).press('Control+A');
    await page.getByLabel('Hispanic or Latino', { exact: true }).press('Backspace');
  }

  if (d.D10 === '0') {
    await page.getByLabel('Not Hispanic or Latino').click();
    await page.getByLabel('Not Hispanic or Latino').press('Control+A');
    await page.getByLabel('Not Hispanic or Latino').press('Backspace');
  }

  if (d.D11 === '0') {
    await page.getByLabel('Employee Declines to Identify Ethnicity').click();
    await page.getByLabel('Employee Declines to Identify Ethnicity').press('Control+A');
    await page.getByLabel('Employee Declines to Identify Ethnicity').press('Backspace');
  }

  const demographicsConcat =
    d.D1 + d.D2 + d.D3 + d.D4 + d.D5 + d.D6 + d.D7 + d.D8 + d.D9 + d.D10 + d.D11;

  if (demographicsConcat === '00000000000') {
    await page.getByLabel('Employee Declines to Identify', { exact: true }).fill('1');
    await page.getByLabel('Employee Declines to Identify Ethnicity').fill('1');
  }

  if (data.laborData.length > 1 && Number(data.laborData[1].zip) > 0) {
    await page.getByLabel('Yes').nth(2).check();

    for (let i = 1; i < data.laborData.length; i++) {
      await page.getByRole('button', { name: 'Add Zip Code' }).click();
      await page.getByLabel('Zip', { exact: true }).fill(data.laborData[i].zip);
      await page.getByLabel('Hours', { exact: true }).fill(data.laborData[i].hours);
      await page.getByRole('button', { name: 'Save', exact: true }).click();
    }
  }

  await page.getByLabel('Solar Training Pipeline').fill(data.J1);
  await page.getByLabel('Craft Apprenticeship Program').fill(data.J2);
  await page.getByLabel('Multi-Cultural Job Training').fill(data.J3);
  await page
    .locator('div')
    .filter({ hasText: /^Total number of graduates of Job Training programs who worked on the project$/ })
    .getByRole('textbox')
    .fill(data.J4);

  await page.getByRole('button', { name: 'Save and Continue' }).click();

  await revisitIfNeeded(page);

  await page.getByLabel('Module Manufacturer / Make*').fill(data.E1);
  await page.getByLabel('Module Model*').fill(data.E2);
  await page.getByLabel('Inverter Manufacturer / Make*').fill(data.E3);
  await page.getByLabel('Inverter Model*').fill(data.inverterModels);

  if (data.E3.includes('Enphase')) {
    await page
      .locator('div')
      .filter({ hasText: /^ANSI C\.12\+\/- 5%$/ })
      .getByRole('combobox')
      .selectOption(data.AC > 10 ? 'ANSI' : '____5_');

    await page.getByLabel('Meter Manufacturer / Make*').fill(data.E3);
    await page.getByLabel('Meter Model*').fill('Envoy');
  } else if (data.AC > 10) {
    await page
      .locator('div')
      .filter({ hasText: /^ANSI C\.12\+\/- 5%$/ })
      .getByRole('combobox')
      .selectOption('ANSI');

    await page.getByLabel('Meter Manufacturer / Make*').fill(data.E5);
    await page.getByLabel('Meter Model*').fill(data.E6);
  } else if (data.E3 === 'APSystems') {
    await page
      .locator('div')
      .filter({ hasText: /^ANSI C\.12\+\/- 5%$/ })
      .getByRole('combobox')
      .selectOption('____5_');

    await page.getByLabel('Meter Manufacturer / Make*').fill(data.E3);
    await page.getByLabel('Meter Model*').fill('ECU');
  } else if (data.E3 === 'Hoymiles') {
    await page
      .locator('div')
      .filter({ hasText: /^ANSI C\.12\+\/- 5%$/ })
      .getByRole('combobox')
      .selectOption('____5_');

    await page.getByLabel('Meter Manufacturer / Make*').fill(data.E3);
    await page.getByLabel('Meter Model*').fill('DTU');
  } else {
    await page
      .locator('div')
      .filter({ hasText: /^ANSI C\.12\+\/- 5%$/ })
      .getByRole('combobox')
      .selectOption('____5_');

    await page.getByLabel('Meter Manufacturer / Make*').fill(data.E3);
    await page.getByLabel('Meter Model*').fill(data.inverterModels);
  }

  await page.getByLabel('Inverter Details').fill('');
  await page.getByRole('combobox').nth(2).selectOption(data.E7);
  await page.getByLabel('Total Project Cost ($)*').fill(data.E8);
  await page.getByRole('button', { name: 'Save and Continue' }).click();
}

async function runPythonExtraction(pythonScriptPath, filteredFilePath, systemId) {
  return new Promise((resolve, reject) => {
    const pythonProcess = spawn('python', [pythonScriptPath, filteredFilePath, String(systemId)], {
      shell: true
    });

    pythonProcess.on('exit', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`Python script exited with code ${code}`));
    });

    pythonProcess.on('error', reject);
  });
}

async function processAndUploadScheduleA(page, systemId, systemName, options) {
  const { downloadFolder, pythonScriptPath } = options;

  await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist?onlyIncompleTasks=false&onlyMyTasks=false&showAll=false`);
  await page.getByRole('link', { name: 'Post Install File Uploads' }).click();
  await page.waitForTimeout(5000);

  const totalFiles = await page.locator('.grid.grid-cols-3.gap-4.items-start.mb-1').nth(0).locator('.flex-auto svg').count();
  const totalShA = await page.locator('.grid.grid-cols-3.gap-4.items-start.mb-1').nth(1).locator('.flex-auto svg').count();

  if (!(totalFiles > 0 && totalShA === 0)) {
    return { uploaded: false, reason: 'No eligible files found' };
  }

  const downloadedFiles = [];

  for (let j = 0; j < totalFiles; j++) {
    try {
      const downloadPromise = page.waitForEvent('download', { timeout: 10000 });
      await page.locator('.grid.grid-cols-3.gap-4.items-start.mb-1').nth(0).locator('.flex-auto svg').nth(j).click();

      const download = await downloadPromise;
      const filePath = path.resolve(downloadFolder, `${systemId}_SchA_${j}.pdf`);
      await download.saveAs(filePath);
      downloadedFiles.push(filePath);
    } catch (error) {
      console.warn(`Skipping file ${j + 1}: ${error.message}`);
    }
  }

  const mergedPdf = await PDFDocument.create();
  const outputFilePath = path.resolve(downloadFolder, `${systemId}_SchA.pdf`);

  for (const file of downloadedFiles) {
    try {
      const stats = await fs.stat(file);
      if (stats.size === 0) continue;

      const pdfBytes = await fs.readFile(file);
      const pdf = await PDFDocument.load(pdfBytes);
      const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
      copiedPages.forEach((p) => mergedPdf.addPage(p));
      await fs.unlink(file);
    } catch (err) {
      console.warn(`Failed processing ${file}: ${err.message}`);
    }
  }

  const mergedPdfBytes = await mergedPdf.save();
  await fs.writeFile(outputFilePath, mergedPdfBytes);

  const filteredPdf = await PDFDocument.create();
  const originalPdf = await PDFDocument.load(mergedPdfBytes);
  const textToCheck = ['Generator Owner’s Consent'];

  for (let i = 0; i < originalPdf.getPageCount(); i++) {
    const singlePagePdf = await PDFDocument.create();
    const [singlePage] = await singlePagePdf.copyPages(originalPdf, [i]);
    singlePagePdf.addPage(singlePage);

    const pageBytes = await singlePagePdf.save();
    const pageData = await pdfParse(pageBytes);

    if (textToCheck.some((text) => pageData.text.includes(text))) {
      const [copiedPage] = await filteredPdf.copyPages(originalPdf, [i]);
      filteredPdf.addPage(copiedPage);
    }
  }

  const filteredPdfBytes = await filteredPdf.save();
  const filteredFilePath = path.resolve(downloadFolder, `${systemId}SchA_api.pdf`);
  await fs.writeFile(filteredFilePath, filteredPdfBytes);

  await runPythonExtraction(pythonScriptPath, filteredFilePath, systemId);

  const extractedTextFile = path.resolve(downloadFolder, `${systemId}extracted_text.txt`);
  const extractedText = await fs.readFile(extractedTextFile, 'utf-8');
  const pdfText = extractedText.toLowerCase().replace(/\s+/g, '');
  const matchesSystem = pdfText.includes(systemName.toLowerCase().replace(/\s+/g, ''));

  if (matchesSystem) {
    await page.getByLabel('Tracking System Approval -').setInputFiles(filteredFilePath);
    await page.waitForTimeout(10000);
    await page.getByRole('button', { name: 'Save', exact: true }).click();
    await page.waitForTimeout(15000);
  }

  await Promise.allSettled([
    fs.unlink(filteredFilePath),
    fs.unlink(outputFilePath)
  ]);

  return {
    uploaded: matchesSystem,
    reason: matchesSystem ? 'Uploaded successfully' : 'Verification failed'
  };
}

async function processSystem(page, systemId, worksheet, options) {
  const data = await getChecklistData(page, systemId, worksheet);
  await openIllinoisApplication(page, data.ABPID);
  await fillIllinoisForms(page, data);

  const scheduleAResult = await processAndUploadScheduleA(page, systemId, data.systemName, options);

  return {
    systemId,
    ABPID: data.ABPID,
    status: 'success',
    scheduleAResult
  };
}

async function runIllinoisAbpAutomation(systemId, options = {}) {
  const {
    headless = false,
    installerInfoPath = process.env.INSTALLER_INFO_PATH,
    downloadFolder = process.env.DOWNLOAD_FOLDER,
    pythonScriptPath = process.env.PYTHON_SCRIPT_PATH
  } = options;

  if (!Number.isFinite(Number(systemId))) {
    throw new Error('A valid systemId is required');
  }

  if (!installerInfoPath) {
    throw new Error('INSTALLER_INFO_PATH is missing');
  }

  if (!downloadFolder) {
    throw new Error('DOWNLOAD_FOLDER is missing');
  }

  if (!pythonScriptPath) {
    throw new Error('PYTHON_SCRIPT_PATH is missing');
  }

  const browser = await chromium.launch({
    headless,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();

  try {
    const worksheet = await loadInstallerInfo(installerInfoPath);

    await loginIllinois(page);
    await loginCSG(page);

    const result = await processSystem(page, Number(systemId), worksheet, {
      downloadFolder,
      pythonScriptPath
    });

    await browser.close();
    return result;
  } catch (err) {
    await browser.close();
    throw err;
  }
}

module.exports = { runIllinoisAbpAutomation };