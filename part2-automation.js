require('dotenv').config();

const { chromium } = require('playwright');
const ExcelJs = require('exceljs');
const { PDFDocument } = require('pdf-lib');
const pdfParse = require('pdf-parse');
const fs = require('fs/promises');
const path = require('path');
const { spawn } = require('child_process');

async function runAutomation(systemId) {
    const browser = await chromium.launch({
        headless: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();

    try {
        function formatDate(dateString) {
            const date = new Date(dateString);
            if (isNaN(date.getTime())) {
                return dateString;
            }
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const day = date.getDate().toString().padStart(2, '0');
            const year = date.getFullYear();
            return `${month}-${day}-${year}`;
        }

        await page.goto('https://portal.illinoisabp.com/');
        await page.getByLabel('Username').fill(process.env.ABP_EMAIL);
        await page.getByLabel('Password').fill(process.env.ABP_PW);
        await page.getByRole('button', { name: 'Sign in' }).first().click();

        await page.waitForTimeout(5000);
        await page.goto('https://portal2.carbonsolutionsgroup.com/admin/login');
        await page.fill('input[type="email"]', process.env.EMAIL);
        await page.fill('input[type="password"]', process.env.PASSWORD);
        await page.getByRole('button', { name: 'Login' }).click();

        await page.waitForTimeout(2000);
        await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/edit?step=2.2`);
        await page.getByRole('button', { name: 'Save', exact: true }).click();
        await page.waitForTimeout(2000);
        await page.waitForLoadState('load');
        await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist?onlyIncompleTasks=false&onlyMyTasks=false&showAll=false`);

        let ABPID = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
        const interconnectionDate = await page.getByRole('row', { name: 'Interconnection Approval Date' }).getByRole('cell').nth(1).innerText();
        const projectOnlineDate = await page.getByRole('row', { name: 'Project Online Date' }).getByRole('cell').nth(1).innerText();
        const dateOfProject = await page.getByRole('row', { name: "Date of Project's Certificate" }).getByRole('cell').nth(1).innerText();
        const completationDate = await page.getByRole('row', { name: 'Construction Completion Date' }).getByRole('cell').nth(1).innerText();
        const Utility = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(11) > td:nth-child(2)').innerText();

        let NON = await page.getByRole('row', { name: 'PJM Gats or MRETs Unit ID' }).getByRole('cell').nth(1).innerText();
        if (NON === 'MISSING!' || NON === '') {
            NON = 'NON123456';
        }

        const I9 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
        const I10 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(12) > td:nth-child(2)').innerText();

        let NM = 'Yes';
        let output = { row: -1, column: -1 };

        const workbook = new ExcelJs.Workbook();
        await workbook.xlsx.readFile(process.env.INSTALLER_INFO_PATH);
        const worksheet = workbook.getWorksheet('Sheet1');

        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value === I10) {
                    output.row = rowNumber;
                    output.column = colNumber;
                }
            });
        });

        if (output.row === -1) {
            throw new Error(`Installer not found in Excel for docket: ${I10}`);
        }

        const I1 = worksheet.getCell(output.row, 1).value;
        const I2 = worksheet.getCell(output.row, 2).value;
        let I3 = worksheet.getCell(output.row, 3).value;
        const I4 = worksheet.getCell(output.row, 4).value;
        const I5 = worksheet.getCell(output.row, 5).value;
        const I6 = worksheet.getCell(output.row, 6).value;
        const I7 = worksheet.getCell(output.row, 7).value;
        const I8 = worksheet.getCell(output.row, 8).value;

        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (cell.value === Utility) {
                    NM = 'No';
                }
            });
        });

        if (I3 === null) {
            I3 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText();
        }

        const WDate = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(13) > td:nth-child(2)').innerText();
        const I11 = formatDate(WDate);

        let D1 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(16) > td:nth-child(2)').innerText();
        let D2 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(17) > td:nth-child(2)').innerText();
        let D3 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(18) > td:nth-child(2)').innerText();
        let D4 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(19) > td:nth-child(2)').innerText();
        let D5 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(20) > td:nth-child(2)').innerText();
        let D6 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(21) > td:nth-child(2)').innerText();
        let D7 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(22) > td:nth-child(2)').innerText();
        let D8 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(23) > td:nth-child(2)').innerText();
        let D9 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(26) > td:nth-child(2)').innerText();
        let D10 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(27) > td:nth-child(2)').innerText();
        let D11 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(28) > td:nth-child(2)').innerText();

        const tableSelector = '#part2-section3 > fieldset > table > tbody > tr:nth-child(31) > td:nth-child(2) > table';
        const rows = await page.$$(tableSelector + ' tr');

        const data = {};

        for (let i = 0; i < rows.length; i++) {
            const cells = await rows[i].$$('td');
            if (cells.length >= 2) {
                const zValue = await cells[0].innerText();
                const hValue = await cells[1].innerText();
                data[`z${i + 1}`] = zValue.trim();
                data[`h${i + 1}`] = hValue.trim();
            }
        }

        const numberOfRows = rows.length;

        const inverterTable = '#part2-section4 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2) > table';
        const inverterRows = await page.$$(inverterTable + ' tr');
        const inverterData = {};
        const inverterRowsCount = inverterRows.length;

        for (let i = 0; i < inverterRows.length; i++) {
            const cells = await inverterRows[i].$$('td');
            if (cells.length >= 4) {
                const model = await cells[1].innerText();
                const size = await cells[2].innerText();
                const quantity = await cells[3].innerText();
                inverterData[`m${i + 1}`] = model.trim();
                inverterData[`s${i + 1}`] = size.trim();
                inverterData[`q${i + 1}`] = quantity.trim();
            }
        }

        let inverterModels = '';
        for (let i = 2; i <= inverterRowsCount; i++) {
            const mValue = inverterData[`m${i}`];
            const sValue = inverterData[`s${i}`];
            const qValue = inverterData[`q${i}`];
            const invertermodel = mValue + ' = ' + sValue + ' x ' + qValue;
            inverterModels += (inverterModels ? ', ' : '') + invertermodel;
        }

        const J1 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(33) > td:nth-child(2)').innerText();
        const J2 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(34) > td:nth-child(2)').innerText();
        const J3 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(35) > td:nth-child(2)').innerText();
        const J4 = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(37) > td:nth-child(2)').innerText();

        const AC = Number(await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText());

        const E1 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        const E2 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
        const E3 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2)').first().innerText();
        const E5 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(7) > td:nth-child(2)').innerText();
        const E6 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
        const E7 = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
        const Amount = await page.locator('#part2-section4 > fieldset > table > tbody > tr:nth-child(11) > td:nth-child(2)').innerText();
        const E8 = Amount.replace(/\$|,/g, '');

        await page.waitForTimeout(2000);
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

        if (await page.getByRole('button', { name: 'Save and Continue' }).isVisible()) {
            await page.getByRole('combobox').first().selectOption('No');
        }

        if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
            await page.getByRole('button', { name: 'Revisit' }).click();
            await page.waitForTimeout(2000);
            await page.getByRole('button', { name: 'OK' }).last().click();
            await page.waitForTimeout(2000);
        }

        await page.getByRole('button', { name: 'Save and Continue' }).click();
        await page.waitForTimeout(2000);
        await page.getByRole('button', { name: 'OK' }).last().click();

        await page.waitForTimeout(2000);

        if (await page.getByRole('button', { name: 'Save and Continue' }).isVisible()) {
            await page.getByRole('combobox').first().selectOption('No');
        }

        if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
            await page.getByRole('button', { name: 'Revisit' }).click();
            await page.waitForTimeout(2000);
            await page.getByRole('button', { name: 'OK' }).last().click();
            await page.waitForTimeout(2000);
        }

        await page.getByLabel('Interconnection Approval Date').fill(interconnectionDate);
        await page.getByLabel('Project Online Date *').fill(projectOnlineDate);
        await page.getByLabel('Date of Project’s Certificate').fill(dateOfProject);
        await page.getByLabel('Date on which Construction').fill(completationDate);
        await page.getByLabel('REC Tracking System *').selectOption('GATS');
        await page.getByLabel('PJM GATS or M-RETS Unit ID *').fill(NON);
        await page.getByLabel('Name on REC Tracking System').fill('Carbon Solutions SREC LLC');
        await page.getByRole('combobox').nth(2).selectOption(NM);
        await page.getByRole('button', { name: 'Save and Continue' }).click();

        await page.waitForTimeout(2000);

        if (await page.getByRole('button', { name: 'Save and Continue' }).isVisible()) {
            await page.getByRole('combobox').first().selectOption('No');
        }

        if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
            await page.getByRole('button', { name: 'Revisit' }).click();
            await page.waitForTimeout(2000);
            await page.getByRole('button', { name: 'OK' }).last().click();
            await page.waitForTimeout(2000);

            if (numberOfRows > 2) {
                for (let i = 2; i <= numberOfRows; i++) {
                    await page.locator('#mxui_widget_VerticalScrollContainer_0 > div.mx-scrollcontainer-middle.region-content > div > div.mx-placeholder > div > div > div > div:nth-child(3) > div > div.mx-name-layoutGrid4.mx-layoutgrid.mx-layoutgrid-fluid > div > div.col-lg.col-md.col > div > div > div > div.mx-name-container5 > div.mx-listview.mx-name-listView1 > ul > li.mx-name-index-0 > div > div > div > div > div:nth-child(3) > div > div > div:nth-child(2) > button').click();
                    await page.getByRole('button', { name: 'OK' }).last().click();
                    await page.waitForTimeout(2000);
                }
            }
        }

        await page.getByLabel('Legal Business Name *').fill(I1?.toString() || '');
        await page.getByLabel('Street *').fill(I2?.toString() || '');
        await page.getByLabel('Apartment or Suite').fill(I3?.toString() || '');
        await page.getByLabel('City *').fill(I4?.toString() || '');
        await page.getByLabel('State *').fill(I5?.toString() || '');
        await page.getByLabel('Zip Code *').fill(I6?.toString() || '');
        await page.getByLabel('Phone *').fill(I7?.toString() || '');
        await page.getByLabel('Email *').fill(I8?.toString() || '');
        await page.getByLabel('Name of the Qualified Person').fill(I9);
        await page.getByLabel('ICC Docket Number for the').fill(I10);
        await page.getByRole('combobox').nth(1).selectOption('No');
        await page.getByPlaceholder('mm-dd-yyyy').fill(I11);

        await page.waitForTimeout(2000);
        if (D1 !== '0') await page.getByLabel('White').fill(D1);
        await page.waitForTimeout(2000);
        if (D2 !== '0') await page.getByLabel('Black or African American').fill(D2);
        await page.waitForTimeout(2000);
        if (D3 !== '0') await page.getByLabel('American Indian or Alaskan').fill(D3);
        await page.waitForTimeout(2000);
        if (D4 !== '0') await page.getByLabel('Asian').fill(D4);
        await page.waitForTimeout(2000);
        if (D5 !== '0') await page.getByLabel('Hawaiian or Other Pacific').fill(D5);
        await page.waitForTimeout(2000);
        if (D6 !== '0') await page.getByLabel('More than one Race').fill(D6);
        await page.waitForTimeout(2000);
        if (D7 !== '0') await page.getByLabel('Some Other Race').fill(D7);
        await page.waitForTimeout(2000);
        if (D8 !== '0') await page.getByLabel('Employee Declines to Identify', { exact: true }).fill(D8);

        await page.getByLabel('Hispanic or Latino', { exact: true }).fill(D9);
        await page.getByLabel('Not Hispanic or Latino').fill(D10);
        await page.getByLabel('Employee Declines to Identify Ethnicity').fill(D11);

        await page.waitForTimeout(2000);

        if (D9 === '0') {
            await page.getByLabel('Hispanic or Latino', { exact: true }).click();
            await page.getByLabel('Hispanic or Latino', { exact: true }).press('Control+A');
            await page.getByLabel('Hispanic or Latino', { exact: true }).press('Backspace');
        }

        await page.waitForTimeout(2000);

        if (D10 === '0') {
            await page.getByLabel('Not Hispanic or Latino').click();
            await page.getByLabel('Not Hispanic or Latino').press('Control+A');
            await page.getByLabel('Not Hispanic or Latino').press('Backspace');
        }

        await page.waitForTimeout(2000);

        if (D11 === '0') {
            await page.getByLabel('Employee Declines to Identify Ethnicity').click();
            await page.getByLabel('Employee Declines to Identify Ethnicity').press('Control+A');
            await page.getByLabel('Employee Declines to Identify Ethnicity').press('Backspace');
        }

        await page.waitForTimeout(2000);

        let D = D1 + D2 + D3 + D4 + D5 + D6 + D7 + D8 + D9 + D10 + D11;

        if (D === '00000000000') {
            await page.getByLabel('Employee Declines to Identify', { exact: true }).fill('1');
            await page.getByLabel('Employee Declines to Identify Ethnicity').fill('1');
        }

        if (data['z2'] > 0) {
            await page.getByLabel('Yes').nth(2).check();

            for (let i = 2; i <= numberOfRows; i++) {
                const zValue = data[`z${i}`];
                const hValue = data[`h${i}`];

                await page.waitForTimeout(3000);
                await page.getByRole('button', { name: 'Add Zip Code' }).click();
                await page.getByLabel('Zip', { exact: true }).fill(zValue);
                await page.getByLabel('Hours', { exact: true }).fill(hValue);
                await page.getByRole('button', { name: 'Save', exact: true }).click();
            }
        }

        await page.getByLabel('Solar Training Pipeline').fill(J1);
        await page.getByLabel('Craft Apprenticeship Program').fill(J2);
        await page.getByLabel('Multi-Cultural Job Training').fill(J3);
        await page.locator('div').filter({ hasText: /^Total number of graduates of Job Training programs who worked on the project$/ }).getByRole('textbox').fill(J4);
        await page.getByRole('button', { name: 'Save and Continue' }).click();

        await page.waitForTimeout(2000);

        if (await page.getByRole('button', { name: 'Save and Continue' }).isVisible()) {
            await page.getByRole('combobox').first().selectOption('No');
        }

        if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
            await page.getByRole('button', { name: 'Revisit' }).click();
            await page.waitForTimeout(2000);
            await page.getByRole('button', { name: 'OK' }).last().click();
            await page.waitForTimeout(2000);
        }

        await page.getByLabel('Module Manufacturer / Make*').fill(E1);
        await page.getByLabel('Module Model*').fill(E2);
        await page.getByLabel('Inverter Manufacturer / Make*').fill(E3);
        await page.getByLabel('Inverter Model*').fill(inverterModels);

        if (AC > 10) {
            await page.locator('div').filter({ hasText: /^ANSI C\.12\+\/- 5%$/ }).getByRole('combobox').selectOption('ANSI');
            await page.getByLabel('Meter Manufacturer / Make*').fill(E5);
            await page.getByLabel('Meter Model*').fill(E6);
        } else if (E3 === 'APSystems') {
            await page.locator('div').filter({ hasText: /^ANSI C\.12\+\/- 5%$/ }).getByRole('combobox').selectOption('____5_');
            await page.getByLabel('Meter Manufacturer / Make*').fill(E3);
            await page.getByLabel('Meter Model*').fill('ECU');
        } else if (E3 === 'Hoymiles') {
            await page.locator('div').filter({ hasText: /^ANSI C\.12\+\/- 5%$/ }).getByRole('combobox').selectOption('____5_');
            await page.getByLabel('Meter Manufacturer / Make*').fill(E3);
            await page.getByLabel('Meter Model*').fill('DTU');
        } else {
            await page.locator('div').filter({ hasText: /^ANSI C\.12\+\/- 5%$/ }).getByRole('combobox').selectOption('____5_');
            await page.getByLabel('Meter Manufacturer / Make*').fill(E3);
            await page.getByLabel('Meter Model*').fill(inverterModels);
        }

        if (E3.includes('Enphase') && AC > 10) {
            await page.locator('div').filter({ hasText: /^ANSI C\.12\+\/- 5%$/ }).getByRole('combobox').selectOption('ANSI');
            await page.getByLabel('Meter Manufacturer / Make*').fill(E3);
            await page.getByLabel('Meter Model*').fill('Envoy');
        }

        if (E3.includes('Enphase') && AC <= 10) {
            await page.locator('div').filter({ hasText: /^ANSI C\.12\+\/- 5%$/ }).getByRole('combobox').selectOption('____5_');
            await page.getByLabel('Meter Manufacturer / Make*').fill(E3);
            await page.getByLabel('Meter Model*').fill('Envoy');
        }

        await page.getByLabel('Inverter Details').fill('');
        await page.getByRole('combobox').nth(2).selectOption(E7);
        await page.getByLabel('Total Project Cost ($)*').fill(E8);
        await page.getByRole('button', { name: 'Save and Continue' }).click();

        const downloadFolder = process.env.DOWNLOAD_FOLDER;
        const pythonScriptPath = process.env.PYTHON_SCRIPT_PATH;

        if (!downloadFolder) {
            throw new Error('DOWNLOAD_FOLDER is missing in environment variables');
        }

        if (!pythonScriptPath) {
            throw new Error('PYTHON_SCRIPT_PATH is missing in environment variables');
        }

        let uploadResult = {
            uploaded: false,
            reason: 'No files processed'
        };

        try {
            await page.waitForTimeout(2000);
            await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist?onlyIncompleTasks=false&onlyMyTasks=false&showAll=false`);

            const systemName = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
            await page.getByRole('link', { name: 'Post Install File Uploads' }).click();
            await page.waitForTimeout(5000);

            const totalFiles = await page.locator('.grid.grid-cols-3.gap-4.items-start.mb-1').nth(0).locator('.flex-auto svg').count();
            const totalShA = await page.locator('.grid.grid-cols-3.gap-4.items-start.mb-1').nth(1).locator('.flex-auto svg').count();

            const downloadedFiles = [];

            if (totalFiles > 0 && totalShA === 0) {
                for (let j = 0; j < totalFiles; j++) {
                    try {
                        const downloadPromise = page.waitForEvent('download', { timeout: 10000 });
                        await page.locator('.grid.grid-cols-3.gap-4.items-start.mb-1').nth(0).locator('.flex-auto svg').nth(j).click();

                        const download = await downloadPromise;
                        const filePath = path.resolve(downloadFolder, `${systemId}_SchA_${j}.pdf`);
                        await download.saveAs(filePath);
                        downloadedFiles.push(filePath);
                    } catch (error) {
                        console.error(`Skipping file ${j + 1} due to download failure:`, error.message);
                    }
                }

                const mergedPdf = await PDFDocument.create();
                const outputFilePath = path.resolve(downloadFolder, `${systemId}_SchA.pdf`);

                for (const file of downloadedFiles) {
                    try {
                        const fileStats = await fs.stat(file);
                        if (fileStats.size === 0) {
                            continue;
                        }

                        const pdfBytes = await fs.readFile(file);
                        const pdf = await PDFDocument.load(pdfBytes);
                        const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
                        copiedPages.forEach(page => mergedPdf.addPage(page));

                        await fs.unlink(file);
                    } catch (err) {
                        console.error(`Error processing file ${file}:`, err);
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

                    if (textToCheck.some(text => pageData.text.includes(text))) {
                        const [copiedPage] = await filteredPdf.copyPages(originalPdf, [i]);
                        filteredPdf.addPage(copiedPage);
                    }
                }

                const filteredPdfBytes = await filteredPdf.save();
                const filteredFilePath = path.resolve(downloadFolder, `${systemId}SchA_api.pdf`);
                await fs.writeFile(filteredFilePath, filteredPdfBytes);

                const pythonProcess = spawn('python', [pythonScriptPath, filteredFilePath, String(systemId)], { shell: true });

                await new Promise((resolve, reject) => {
                    pythonProcess.on('exit', async (code) => {
                        if (code === 0) {
                            try {
                                const extractedTextFile = path.resolve(downloadFolder, `${systemId}extracted_text.txt`);
                                const extractedText = await fs.readFile(extractedTextFile, 'utf-8');
                                const pdfText = extractedText.toLowerCase().replace(/\s+/g, '');

                                if (pdfText.includes(systemName.toLowerCase().replace(/\s+/g, ''))) {
                                    await page.getByLabel('Tracking System Approval -').setInputFiles(filteredFilePath);
                                    await page.waitForTimeout(10000);
                                    await page.getByRole('button', { name: 'Save', exact: true }).click();
                                    await page.waitForTimeout(15000);

                                    uploadResult = {
                                        uploaded: true,
                                        reason: 'Uploaded successfully'
                                    };
                                } else {
                                    uploadResult = {
                                        uploaded: false,
                                        reason: 'Verification failed'
                                    };
                                }

                                await fs.unlink(filteredFilePath).catch(() => {});
                                await fs.unlink(outputFilePath).catch(() => {});
                            } catch (err) {
                                reject(err);
                                return;
                            }
                        } else {
                            reject(new Error(`Python script exited with code ${code}`));
                            return;
                        }

                        resolve();
                    });

                    pythonProcess.on('error', reject);
                });
            } else {
                uploadResult = {
                    uploaded: false,
                    reason: 'No files found or Schedule A already exists'
                };
            }
        } catch (err) {
            uploadResult = {
                uploaded: false,
                reason: err.message
            };
        }

        await browser.close();

        return {
            success: true,
            systemId,
            ABPID,
            scheduleAResult: uploadResult
        };
    } catch (err) {
        await browser.close();
        throw err;
    }
}

module.exports = { runAutomation };