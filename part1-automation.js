const fs = require('fs');
const path = require('path');
const ExcelJs = require('exceljs');
const { chromium } = require('playwright');

const DEFAULT_SYSTEM_IDS = [];


const PORTAL_URL = 'https://portal.illinoisabp.com/';
const ADMIN_LOGIN_URL = 'https://portal2.carbonsolutionsgroup.com/admin/login';

function resolveCredentials() {
  const abpUsername = process.env.ABP_Email;
  const abpPassword = process.env.ABP_PW;
  const portal2Email = process.env.CSG_ADMIN_EMAIL || process.env.EMAIL;
  const portal2Password = process.env.CSG_ADMIN_PASSWORD || process.env.PASSWORD;

  if (!abpUsername || !abpPassword) {
    throw new Error(
      'Missing ABP credentials. Set ABP_USERNAME (or ABP_EMAIL/ABP_USER/SREC_USERNAME) and ABP_PASSWORD (or ABP_PASS/SREC_PASSWORD).'
    );
  }

  if (!portal2Email || !portal2Password) {
    throw new Error('Missing Portal2 credentials. Set CSG_ADMIN_EMAIL/CSG_ADMIN_PASSWORD (or EMAIL/PASSWORD) in .env');
  }

  return { abpUsername, abpPassword, portal2Email, portal2Password };
}

function parseSystemIds(inputIds) {
  if (Array.isArray(inputIds) && inputIds.length > 0) {
    return inputIds
      .map((value) => Number(value))
      .filter((value) => Number.isFinite(value));
  }

  return [...DEFAULT_SYSTEM_IDS];
}

async function readCellText(page, selector) {
  const locator = page.locator(selector);
  const count = await locator.count();
  if (count === 0) {
    return '';
  }
  return (await locator.innerText()).trim();
}

async function clickIfVisible(page, roleName, options = {}) {
  const locator = page.getByRole('button', { name: roleName, ...options });
  if (await locator.isVisible().catch(() => false)) {
    await locator.click();
    return true;
  }
  return false;
}

async function safeGoto(page, url, waitMs = 1500) {
  await page.goto(url, { waitUntil: 'domcontentloaded' });
  await page.waitForTimeout(waitMs);
}

function parseAddress(addressRaw) {
  const lines = String(addressRaw || '').split('\n').map((line) => line.trim()).filter(Boolean);
  const street = (lines[0] || '').replace(/,/g, '').trim();

  const cityStateZipChunk = lines[1] || '';
  const cityParts = cityStateZipChunk.split(',');
  const city = (cityParts[0] || '').trim();

  const stateZip = (cityParts[1] || '').trim().split(/\s+/);
  const state = (stateZip[0] || '').trim();
  const zipCode = (stateZip[1] || '').trim();

  return { street, city, state, zipCode };
}

function buildInterconnectionDate() {
  const today = new Date();
  const interconnectionDate = new Date();
  interconnectionDate.setMonth(today.getMonth() + 6);

  const month = String(interconnectionDate.getMonth() + 1).padStart(2, '0');
  const day = String(interconnectionDate.getDate()).padStart(2, '0');
  const year = interconnectionDate.getFullYear();

  return `${month}-${day}-${year}`;
}

async function loginAbp(page, creds) {
  await safeGoto(page, PORTAL_URL);
  await page.getByLabel('Username').fill(creds.abpUsername);
  await page.getByLabel('Password').fill(creds.abpPassword);
  await page.getByRole('button', { name: 'Sign in' }).first().click();
  await page.waitForTimeout(2000);
}

async function loginPortal2(page, creds) {
  await safeGoto(page, ADMIN_LOGIN_URL);
  await page.getByLabel('Email').fill(creds.portal2Email);
  await page.getByLabel('Password').fill(creds.portal2Password);
  await page.getByRole('button', { name: 'Login' }).click();
  await page.waitForTimeout(2000);
}

async function extractRowsData(page) {
  const tableSelector = '#part1-section4 > fieldset > table > tbody > tr:nth-child(6) > td:nth-child(2) > table';
  const rows = await page.$$(tableSelector + ' tr');
  const data = {};

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const cells = await rows[rowIndex].$$('td');
    if (cells.length < 6) {
      continue;
    }

    const sValue = (await cells[0].innerText()).trim();
    const qValue = (await cells[1].innerText()).trim();
    const tValue = (await cells[2].innerText()).trim();
    const tiValue = (await cells[3].innerText()).trim();
    const oValue = (await cells[4].innerText()).trim();
    const bValue = (await cells[5].innerText()).trim();

    data[`s${rowIndex + 1}`] = sValue;
    data[`q${rowIndex + 1}`] = qValue;
    data[`t${rowIndex + 1}`] = tValue;
    data[`ti${rowIndex + 1}`] = String(Math.round(Number(tiValue || '0')));
    data[`o${rowIndex + 1}`] = String(Math.round(Number(oValue || '0')));
    data[`b${rowIndex + 1}`] = bValue;
  }

  return { rowsCount: rows.length, data };
}

async function extractSystemData(page, systemId) {
  await safeGoto(page, `https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist#status-info`, 2000);

  const disclosureID = await readCellText(page, '#status-info > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)');
  const address = await readCellText(page, '#customer-info > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)');
  const latitude = await readCellText(page, '#part1-section1 > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)');
  const longitude = await readCellText(page, '#part1-section1 > fieldset > table > tbody > tr:nth-child(9) > td:nth-child(2)');
  const parcelNumber = await readCellText(page, '#part1-section1 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)');

  const projectType = await readCellText(page, '#part1-section2 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)');
  const financing = await readCellText(page, '#part1-section2 > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)');
  const projectCategory = await readCellText(page, '#part1-section2 > fieldset > table > tbody > tr:nth-child(11) > td:nth-child(2)');

  const ownerName = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)');
  const ownerPhone = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(9) > td:nth-child(2)');
  const ownerEmail = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)');

  const installerName = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(12) > td:nth-child(2)');
  const installerStreet = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(13) > td:nth-child(2)');
  const installerSuite = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(14) > td:nth-child(2)');
  const installerCity = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(15) > td:nth-child(2)');
  const installerState = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(16) > td:nth-child(2)');
  const installerZip = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(17) > td:nth-child(2)');
  const installerPhone = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(18) > td:nth-child(2)');
  const installerEmail = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(19) > td:nth-child(2)');
  let graduates = await readCellText(page, '#part1-section3 > fieldset > table > tbody > tr:nth-child(20) > td:nth-child(2)');
  if (graduates === 'MISSING!') {
    graduates = '0';
  }

  const ac = await readCellText(page, '#part1-section4 > fieldset > table > tbody > tr:nth-child(1) > td:nth-child(2)');
  const inverterEfficiency = await readCellText(page, '#part1-section4 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)');
  let groundCoverRatio = await readCellText(page, '#part1-section4 > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)');
  if (groundCoverRatio) {
    groundCoverRatio = '0.4';
  }

  const minimumShadingCriteria = await readCellText(page, '#part1-section4 > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)');
  const batteryBackup = await readCellText(page, '#part1-section4 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2)');
  const { rowsCount, data } = await extractRowsData(page);

  const customCapacityFactor = await readCellText(page, '#part1-section5 > fieldset > table > tbody > tr:nth-child(1) > td:nth-child(2)');
  const explanation = await readCellText(page, '#part1-section5 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)');
  const utilityName = await readCellText(page, '#part1-section6 > fieldset > table > tbody > tr:nth-child(1) > td:nth-child(2)');
  const projectdateRaw = await readCellText(page, '#part1-section6 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)');
  const projectdate = Number(projectdateRaw || '0');

  return {
    disclosureID,
    address,
    latitude,
    longitude,
    parcelNumber,
    projectType,
    financing,
    projectCategory,
    ownerName,
    ownerPhone,
    ownerEmail,
    installerName,
    installerStreet,
    installerSuite,
    installerCity,
    installerState,
    installerZip,
    installerPhone,
    installerEmail,
    graduates,
    ac,
    inverterEfficiency,
    groundCoverRatio,
    minimumShadingCriteria,
    batteryBackup,
    rowsCount,
    data,
    customCapacityFactor,
    explanation,
    utilityName,
    projectdate
  };
}

async function loadMappings(systemData) {
  const workbook = new ExcelJs.Workbook();
  const defaultPath = path.resolve(__dirname, 'InstallerInfo.xlsx');
  const installerInfoPath = process.env.INSTALLER_INFO_PATH || defaultPath;

  if (!fs.existsSync(installerInfoPath)) {
    throw new Error(`Installer info workbook not found at: ${installerInfoPath}`);
  }

  await workbook.xlsx.readFile(installerInfoPath);
  const worksheet = workbook.getWorksheet('UtilityNames');
  if (!worksheet) {
    throw new Error('Worksheet "UtilityNames" not found in installer workbook.');
  }

  let utilityRow = -1;
  let projectTypeRow = -1;
  let financingRow = -1;

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      if (cell.value === systemData.utilityName) {
        utilityRow = rowNumber;
      }
      if (cell.value === systemData.projectType) {
        projectTypeRow = rowNumber;
      }
      if (cell.value === systemData.financing) {
        financingRow = rowNumber;
      }
    });
  });

  if (utilityRow < 0 || projectTypeRow < 0 || financingRow < 0) {
    throw new Error('Could not map utility/projectType/financing values from InstallerInfo workbook.');
  }

  return {
    ABPUtility: worksheet.getCell(utilityRow, 2).value,
    ABPUtility2: worksheet.getCell(utilityRow, 3).value,
    projectType1: worksheet.getCell(projectTypeRow, 2).value,
    financing1: worksheet.getCell(financingRow, 2).value
  };
}

async function openApplicationFromDisclosure(page, disclosureID) {
  await safeGoto(page, PORTAL_URL);
  await page.getByRole('button', { name: 'View Project Applications' }).click();
  await page.getByRole('button', { name: 'New DG Project Application' }).click();
  await page.getByLabel('FormID').fill(disclosureID);
  await page.getByRole('button', { name: 'Search', exact: true }).click();
  await page.waitForTimeout(3500);

  const disclosureIDCount = await page.getByRole('cell', { name: disclosureID }).count();
  if (disclosureIDCount === 0) {
    return { found: false };
  }

  await page.getByRole('cell', { name: disclosureID }).first().click();
  await page.getByRole('button', { name: 'Start Application' }).click();
  await page.waitForTimeout(1500);

  if (await clickIfVisible(page, 'OK')) {
    await page.waitForTimeout(1000);
  }

  if (await page.getByLabel('close').isVisible().catch(() => false)) {
    await page.getByLabel('close').click();
  }

  return { found: true };
}

async function getApplicationId(page, disclosureID) {
  await page.getByRole('button', { name: 'Disclosure Form Search' }).click();
  await page.getByLabel('Form ID').fill(disclosureID);
  await page.getByRole('button', { name: 'Search', exact: true }).click();
  await page.waitForTimeout(2000);

  const applicationText = await page.locator('div.mx-datagrid-data-wrapper').nth(1).innerText();
  const applicationID = applicationText.replace(/[^0-9.-]/g, '').toString();

  if (await page.getByRole('button', { name: 'Close page' }).isVisible().catch(() => false)) {
    await page.getByRole('button', { name: 'Close page' }).click();
  }

  return applicationID;
}

async function fillSection1(page, systemData, mapping) {
  await page.getByRole('button', { name: 'Section 1 - Project Location' }).click();
  await page.waitForTimeout(1500);

  if (await clickIfVisible(page, 'Revisit')) {
    await page.waitForTimeout(1500);
    await page.getByRole('button', { name: 'OK' }).last().click();
    await page.waitForTimeout(1500);
  }

  await page.getByLabel('Latitude').fill(systemData.latitude);
  await page.getByLabel('Longitude').fill(systemData.longitude);
  await page.getByLabel('Parcel Number').fill(systemData.parcelNumber);
  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(1500);

  if (await clickIfVisible(page, 'Revisit')) {
    await page.waitForTimeout(1500);
    await page.getByRole('button', { name: 'OK' }).last().click();
    await page.waitForTimeout(1500);
  }

  await page.getByLabel('Project Type *').selectOption(String(mapping.projectType1));
  await page.getByLabel('Financing Structure *').selectOption(String(mapping.financing1));

  await page.locator('text="No"').nth(0).click();
  await page.locator('text="No"').nth(1).click();
  await page.locator('.form-control').nth(2).selectOption('No');
  await page.locator('.form-control').nth(3).selectOption(systemData.projectCategory);
  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(4000);
}

async function fillSection2(page, systemData) {
  if (await clickIfVisible(page, 'Revisit')) {
    await page.waitForTimeout(1500);
    await page.getByRole('button', { name: 'OK' }).last().click();
    await page.waitForTimeout(1500);
  }

  const { street, city, state, zipCode } = parseAddress(systemData.address);

  await page.getByLabel('Name of Owner or Point of').fill(systemData.ownerName);
  await page.getByLabel('Street *').first().fill(street);
  await page.getByLabel('City *').first().fill(city);
  await page.getByLabel('State *').first().fill(state);
  await page.getByLabel('Zip Code *').first().fill(zipCode);
  await page.getByLabel('Phone *').first().fill(systemData.ownerPhone);
  await page.getByLabel('Email', { exact: true }).fill(systemData.ownerEmail);

  await page.locator('.radio').nth(0).click();
  await page.getByLabel('Legal Business Name *').fill(systemData.installerName);
  await page.getByLabel('Street *').nth(1).fill(systemData.installerStreet);
  await page.getByLabel('Apartment or Suite').nth(1).fill(systemData.installerSuite);
  await page.getByLabel('City *').nth(1).fill(systemData.installerCity);
  await page.getByLabel('State *').nth(1).fill(systemData.installerState);
  await page.getByLabel('Zip Code *').nth(1).fill(systemData.installerZip);
  await page.getByLabel('Phone *').nth(1).fill(systemData.installerPhone);
  await page.getByLabel('Email *').fill(systemData.installerEmail);

  await page.waitForTimeout(1000);
  await page.getByText('Yes').nth(1).click();
  await page.getByLabel('Number of Graduates of Job').fill(systemData.graduates);
  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(1500);
}

async function fillSection3(page, systemData) {
  if (await clickIfVisible(page, 'Revisit')) {
    await page.waitForTimeout(1500);
    await page.getByRole('button', { name: 'Continue' }).click();
    await page.waitForTimeout(1500);
  }

  const deleteSelector = '.btn.mx-button.mx-name-actionButton5.spacing-outer-left.btn-sm.btn-default';
  const totalDelete = await page.locator(deleteSelector).count();
  for (let idx = 0; idx < totalDelete; idx++) {
    await page.locator(deleteSelector).first().click();
    await page.waitForTimeout(1000);
  }

  await page.locator('.form-control').nth(1).fill(systemData.ac);
  await page.getByLabel('Inverter Efficiency (%) *').fill(systemData.inverterEfficiency);
  await page.getByLabel('Ground Cover Ratio *').fill(systemData.groundCoverRatio);
  await page.waitForTimeout(1000);
  await page.getByLabel(systemData.minimumShadingCriteria).check();
  await page.getByRole('combobox').selectOption(systemData.batteryBackup);

  for (let idx = 2; idx <= systemData.rowsCount; idx++) {
    await page.getByRole('button', { name: 'Add' }).click();

    const sValue = systemData.data[`s${idx}`] || '';
    const qValue = systemData.data[`q${idx}`] || '';
    const tValue = systemData.data[`t${idx}`] || '';
    const tiValue = systemData.data[`ti${idx}`] || '';
    const oValue = systemData.data[`o${idx}`] || '';
    const bValue = systemData.data[`b${idx}`] || 'No';

    await page.waitForTimeout(1500);
    await page.getByPlaceholder('to 1000').click();
    await page.keyboard.type(sValue, { delay: 60 });

    await page.getByPlaceholder('Greater than').click();
    await page.keyboard.type(qValue, { delay: 60 });

    await page.getByPlaceholder('to 90').click();
    await page.keyboard.type(tiValue, { delay: 60 });

    await page.getByPlaceholder('to 359 inclusive').click();
    await page.keyboard.type(oValue, { delay: 60 });

    if (tValue === 'Fixed - Roof Mount') {
      await page.getByLabel('Roof').check();
      await page.getByLabel('Tracking Type *').selectOption('Fixed_Mount___Roof_Mounted');
    } else {
      await page.getByLabel('Ground', { exact: true }).check();
      await page.getByLabel('Tracking Type *').selectOption('Fixed_Mount');
    }

    await page.getByLabel('Are you using Bifacial Panels').getByLabel(bValue).check();
    if (bValue === 'Yes') {
      await page.getByPlaceholder('0 to 1', { exact: true }).fill('0.4');
    }

    await page.getByRole('button', { name: 'Save', exact: true }).click();
  }

  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(3000);
}

async function fillSection5(page, systemData) {
  await page.goBack();
  await page.getByRole('button', { name: 'Section 5 - REC Estimate' }).click();
  await page.waitForTimeout(1500);

  if (await clickIfVisible(page, 'Revisit')) {
    await page.waitForTimeout(1500);
    await page.getByRole('button', { name: 'OK' }).last().click();
    await page.waitForTimeout(1500);
  }

  await page.getByLabel('PVWatts').check();
  await page.getByRole('button', { name: 'Calculate' }).click();
  await page.waitForTimeout(3000);
  await page.getByRole('button', { name: 'OK' }).last().click();

  await page.getByLabel('Custom Capacity Factor').check();
  await page.getByPlaceholder('Example: for 50.5%, enter').fill(systemData.customCapacityFactor);
  await page.getByLabel('Explanation of Custom').fill(systemData.explanation);
  await page.getByRole('button', { name: 'Calculate' }).click();
  await page.waitForTimeout(3000);
  await page.getByRole('button', { name: 'OK' }).last().click();

  const recErrorChecks = await page.locator('input[type="checkbox"]').count();
  if (recErrorChecks > 0) {
    await page.locator('input[type="checkbox"]').first().check();
    await page.waitForTimeout(1000);
  }

  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(1200);
  await page.getByRole('button', { name: 'OK' }).click();
}

async function fillSection6(page, systemData, mapping) {
  if (await clickIfVisible(page, 'Revisit')) {
    await page.waitForTimeout(1500);
    await page.getByRole('button', { name: 'OK' }).last().click();
    await page.waitForTimeout(1500);
  }

  await page.locator('.form-control').nth(0).selectOption(String(mapping.ABPUtility));
  await page.waitForTimeout(2000);

  if (systemData.projectdate === 0) {
    await page.locator('text="No"').first().click();
  } else {
    await page.locator('text="Yes"').first().click();
  }

  const count = await page.locator('.form-control').count();
  await page.locator('.form-control').nth(1).click();
  if (count === 7 && mapping.ABPUtility2) {
    await page.waitForTimeout(1000);
    await page.locator('.form-control').nth(1).selectOption(String(mapping.ABPUtility2));
    await page.waitForTimeout(1000);
    await page.locator('.form-control').nth(2).click();
  }

  await page.keyboard.press('Control+A');
  await page.keyboard.press('Backspace');
  await page.keyboard.type(buildInterconnectionDate(), { delay: 50 });

  await page.getByRole('button', { name: 'Save and Continue' }).click();
  await page.waitForTimeout(2500);
}

async function runSingleSystem(page, systemId) {
  const systemData = await extractSystemData(page, systemId);
  if (!systemData.disclosureID) {
    return { systemId, status: 'failed', reason: 'Disclosure ID not found' };
  }

  const mapping = await loadMappings(systemData);

  const openResult = await openApplicationFromDisclosure(page, systemData.disclosureID);
  if (!openResult.found) {
    return { systemId, status: 'skipped', reason: `${systemData.disclosureID} not found` };
  }

  const applicationID = await getApplicationId(page, systemData.disclosureID);
  await page.getByRole('columnheader', { name: 'sort Project Application ID ?' }).getByRole('spinbutton').fill(applicationID);
  await page.waitForTimeout(1500);

  const submittedExists = (await page.locator('text=Submitted').count()) > 0;
  if (submittedExists) {
    return { systemId, status: 'skipped', reason: 'Already submitted', applicationID };
  }

  await page.locator('#mxui_widget_VerticalScrollContainer_0 > div.mx-scrollcontainer-middle.region-content > div > div.mx-placeholder > div > div > div > div.mx-dataview.mx-name-dataView2.form-horizontal > div > div > div > div.mx-name-dataGrid22.widget-datagrid.widget-datagrid-selectable-rows.widget-datagrid-selection-method-click > div.widget-datagrid-content.sticky-table-container > div.widget-datagrid-grid.table > div > div:nth-child(2) > div:nth-child(6) > div > div > div > div.col-lg.col-md.col > div > button').click();

  await page.waitForTimeout(1500);

  await fillSection1(page, systemData, mapping);
  await fillSection2(page, systemData);
  await fillSection3(page, systemData);
  await fillSection5(page, systemData);
  await fillSection6(page, systemData, mapping);

  return { systemId, status: 'completed', applicationID };
}

async function runPart1DataEntry(options = {}) {
  const { headless = false, systemIds } = options;
  const creds = resolveCredentials();
  const ids = parseSystemIds(systemIds);

  if (ids.length === 0) {
    throw new Error('No valid system IDs supplied.');
  }

  const browser = await chromium.launch({
    headless,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const page = await browser.newPage();
  const results = [];

  try {
    await loginAbp(page, creds);
    await loginPortal2(page, creds);

    for (const systemId of ids) {
      try {
        const result = await runSingleSystem(page, systemId);
        results.push(result);
      } catch (error) {
        results.push({
          systemId,
          status: 'failed',
          reason: error.message
        });
      }
    }

    return {
      total: ids.length,
      completed: results.filter((r) => r.status === 'completed').length,
      skipped: results.filter((r) => r.status === 'skipped').length,
      failed: results.filter((r) => r.status === 'failed').length,
      results
    };
  } finally {
    await browser.close();
  }
}

module.exports = { runPart1DataEntry, DEFAULT_SYSTEM_IDS };
