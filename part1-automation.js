require('dotenv').config();

const { chromium } = require('playwright');
const ExcelJs = require('exceljs');

async function runAutomation(systemId) {
    const browser = await chromium.launch({
        headless: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });

    const page = await browser.newPage();
    let uploadResult = {
        uploaded: false,
        reason: 'Part1 automation did not complete'
    };

    try {

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
        await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist#status-info`);
        await page.waitForTimeout(2000);
        const disclosureID = await page.locator('#status-info > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        const address = await page.locator('#customer-info > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText();
        const Latitude = await page.locator('#part1-section1 > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
        const Longitude = await page.locator('#part1-section1 > fieldset > table > tbody > tr:nth-child(9) > td:nth-child(2)').innerText();
        const ParcelNumber = await page.locator('#part1-section1 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
        const projectType = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        const financing = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
        const parcel = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText();
        const parcelYes = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2)').innerText();
        const expansion = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(6) > td:nth-child(2)').innerText();
        const expansionYes = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(7) > td:nth-child(2)').innerText();
        let publicSchool = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
        const publicSchoolQ1 = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(9) > td:nth-child(2)').innerText();
        const publicSchoolQ2 = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
        const projectCategory = await page.locator('#part1-section2 > fieldset > table > tbody > tr:nth-child(11) > td:nth-child(2)').innerText();
        const ownerName = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        const ownerPhone = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(9) > td:nth-child(2)').innerText();
        const ownerEmail = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(10) > td:nth-child(2)').innerText();
        const InstallerName = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(12) > td:nth-child(2)').innerText();
        const InstallerStreet = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(13) > td:nth-child(2)').innerText();
        const InstallerSuite = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(14) > td:nth-child(2)').innerText();
        const InstallerCity = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(15) > td:nth-child(2)').innerText();
        const InstallerState = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(16) > td:nth-child(2)').innerText();
        const InstallerZip = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(17) > td:nth-child(2)').innerText();
        const InstallerPhone = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(18) > td:nth-child(2)').innerText();
        const InstallerEmail = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(19) > td:nth-child(2)').innerText();
        let graduates = await page.locator('#part1-section3 > fieldset > table > tbody > tr:nth-child(20) > td:nth-child(2)').innerText();

        if (graduates === 'MISSING!') {
            graduates = '0';
        }

        const AC = await page.locator('#part1-section4 > fieldset > table > tbody > tr:nth-child(1) > td:nth-child(2)').innerText();
        const InverterEfficiency = await page.locator('#part1-section4 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        let GroundCoverRatio = await page.locator('#part1-section4 > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
        const MinimumShadingCriteria = await page.locator('#part1-section4 > fieldset > table > tbody > tr:nth-child(4) > td:nth-child(2)').innerText();
        const BatteryBackup = await page.locator('#part1-section4 > fieldset > table > tbody > tr:nth-child(5) > td:nth-child(2)').innerText();
        const tableSelector = '#part1-section4 > fieldset > table > tbody > tr:nth-child(6) > td:nth-child(2) > table';
        const rows = await page.$$(tableSelector + ' tr');


        const data = {};

        for (let i = 0; i < rows.length; i++) {
            const cells = await rows[i].$$('td');
            if (cells.length >= 2) {
                const sValue = await cells[0].innerText();
                const qValue = await cells[1].innerText();
                const tValue = await cells[2].innerText();
                const tiValue = await cells[3].innerText();
                const oValue = await cells[4].innerText();
                const bValue = await cells[5].innerText();

                data[`s${i + 1}`] = sValue.trim(); // trim() added to remove extra spaces
                data[`q${i + 1}`] = qValue.trim(); // trim() added to remove extra spaces
                data[`t${i + 1}`] = tValue.trim();
                data[`ti${i + 1}`] = (Math.round(Number(tiValue.trim()))).toString();
                data[`o${i + 1}`] = (Math.round(Number(oValue.trim()))).toString();
                data[`b${i + 1}`] = bValue.trim();
            }
        }

        const numberOfRows = rows.length; // Get the number of rows dynamically


        const CustomCapacityFactor = await page.locator('#part1-section5 > fieldset > table > tbody > tr:nth-child(1) > td:nth-child(2)').innerText();
        const Explanation = await page.locator('#part1-section5 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        let UtilityName = await page.locator('#part1-section6 > fieldset > table > tbody > tr:nth-child(1) > td:nth-child(2)').innerText();
        let projectdate = await page.locator('#part1-section6 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        projectdate = Number(projectdate);

        const today = new Date();
        const interconnectionDate = new Date();
        interconnectionDate.setMonth(today.getMonth() + 6);
        // Formatting as MM-DD-YYYY
        const month = String(interconnectionDate.getMonth() + 1).padStart(2, '0'); // Months are 0-based
        const day = String(interconnectionDate.getDate()).padStart(2, '0');
        const year = interconnectionDate.getFullYear();

        const formattedDate = `${month}-${day}-${year}`;


        const lines = address.split("\n");
        const street = lines[0].replace(/,/g, '').trim();

        const cityStateZip = lines[1].trim().split(",");
        const city = cityStateZip[0].trim();

        const stateZip = cityStateZip[1].trim().split(" ");
        const state = stateZip[0].trim();
        const zipCode = stateZip[1].trim();



        let utilityOutput = { row: -1, column: -1 };
        let projectTypeOutput = { row: -1, column: -1 };
        let financingOutput = { row: -1, column: -1 };
        const workbook = new ExcelJs.Workbook();
        await workbook.xlsx.readFile(process.env.INSTALLER_INFO_PATH);
        const worksheet = workbook.getWorksheet('UtilityNames');
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value === UtilityName) {
                    utilityOutput.row = rowNumber;
                    utilityOutput.column = colNumber;
                }
                if (cell.value === projectType) {
                    projectTypeOutput.row = rowNumber;
                    projectTypeOutput.column = colNumber;
                }
                if (cell.value === financing) {
                    financingOutput.row = rowNumber;
                    financingOutput.column = colNumber;
                }
            })
        })

        if (utilityOutput.row === -1) {
            throw new Error(`Utility "${UtilityName}" not found in UtilityNames sheet`);
        }
        if (projectTypeOutput.row === -1) {
            throw new Error(`Project Type "${projectType}" not found in UtilityNames sheet`);
        }
        if (financingOutput.row === -1) {
            throw new Error(`Financing "${financing}" not found in UtilityNames sheet`);
        }

        const ABPUtility = worksheet.getCell(utilityOutput.row, 2).value;
        const ABPUtility2 = worksheet.getCell(utilityOutput.row, 3).value;
        const projectType1 = worksheet.getCell(projectTypeOutput.row, 2).value;
        const financing1 = worksheet.getCell(financingOutput.row, 2).value;





        await page.goto('https://portal.illinoisabp.com/');
        await page.waitForTimeout(2000);
        await page.getByRole('button', { name: 'View Project Applications' }).click();
        await page.getByRole('button', { name: 'New DG Project Application' }).click();
        await page.getByLabel('FormID').click();
        await page.getByLabel('FormID').fill(disclosureID);
        await page.getByRole('button', { name: 'Search', exact: true }).click();
        await page.waitForTimeout(4000);
        const disclosureIDCount = await page.getByRole('cell', { name: disclosureID }).count();
        if (disclosureIDCount > 0) {
            await page.getByRole('cell', { name: disclosureID }).click();
            await page.getByRole('button', { name: 'Start Application' }).click();
            await page.waitForTimeout(2000);
            await page.goto('https://portal.illinoisabp.com/');
            await page.waitForTimeout(2000);
            await page.getByRole('button', { name: 'View Project Applications' }).click();
            await page.getByRole('button', { name: 'New DG Project Application' }).click();
            await page.getByLabel('FormID').click();
            await page.getByLabel('FormID').fill(disclosureID);
            await page.getByRole('button', { name: 'Search', exact: true }).click();
            await page.waitForTimeout(2000);
            await page.getByRole('cell', { name: disclosureID }).click();
            await page.getByRole('button', { name: 'Start Application' }).click();
            await page.getByRole('button', { name: 'OK' }).click();
            await page.getByLabel('close').click();
            await page.getByRole('button', { name: 'Disclosure Form Search' }).click();
            await page.getByLabel('Form ID').click();
            await page.getByLabel('Form ID').fill(disclosureID);
            await page.getByRole('button', { name: 'Search', exact: true }).click();
            await page.waitForTimeout(2000);
            const applicationText = await page.locator('div.mx-datagrid-data-wrapper').nth(1).innerText();
            const applicationID = (applicationText.replace(/[^0-9.-]/g, '')).toString();
            console.log('Application ID', applicationID);
            await page.getByRole('button', { name: 'Close page' }).click();
            await page.getByRole('columnheader', { name: 'sort Project Application ID ' }).getByRole('spinbutton').fill(applicationID);
            await page.waitForTimeout(2000);

            const textToCheck = 'Submitted';
            const textLocator = page.locator(`text=${textToCheck}`);

            // Check if the text exists
            const textExists = await textLocator.count() > 0;

            if (textExists) {
                uploadResult = {
                    uploaded: false,
                    reason: `Application ${applicationID} is already submitted`
                };
            } else {
                await page.locator('#mxui_widget_VerticalScrollContainer_0 > div.mx-scrollcontainer-middle.region-content > div > div.mx-placeholder > div > div > div > div.mx-dataview.mx-name-dataView2.form-horizontal > div > div > div > div.mx-name-dataGrid22.widget-datagrid.widget-datagrid-selectable-rows.widget-datagrid-selection-method-click > div.widget-datagrid-content.sticky-table-container > div.widget-datagrid-grid.table > div > div:nth-child(2) > div:nth-child(6) > div > div > div > div.col-lg.col-md.col > div > button').click();

                await page.waitForTimeout(2000);
                await page.getByRole('button', { name: 'Section 1 - Project Location' }).click();
                await page.waitForTimeout(2000);
                if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
                    await page.getByRole('button', { name: 'Revisit' }).click();
                    await page.waitForTimeout(2000);
                    await page.getByRole('button', { name: 'OK' }).last().click();
                    await page.waitForTimeout(2000);
                }
                await page.getByLabel('Latitude').fill(Latitude);
                await page.getByLabel('Longitude').fill(Longitude);
                await page.getByLabel('Parcel Number').fill(ParcelNumber);
                await page.getByRole('button', { name: 'Save and Continue' }).click();
                await page.waitForTimeout(2000);
                if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
                    await page.getByRole('button', { name: 'Revisit' }).click();
                    await page.waitForTimeout(2000);
                    await page.getByRole('button', { name: 'OK' }).last().click();
                    await page.waitForTimeout(2000);
                }
                await page.getByLabel('Project Type *').selectOption(projectType1);
                await page.getByLabel('Financing Structure *').selectOption(financing1);
                let n = 2;
                await page.locator(`text="No"`).nth(0).click();
                await page.locator(`text="No"`).nth(1).click();
                /*await page.locator(`text=${parcel}`).nth(0).click();
                if(parcel === 'Yes'){
                    await page.getByLabel('Co-located Pair Application').fill(parcelYes);
                    n = n + 1;
                }
                await page.waitForTimeout(5000);
                await page.locator(`text=${parcel}`).nth(1).click();
                if(expansion === 'Yes'){
                    await page.getByLabel('Application ID of the').fill(expansionYes);
                    n = n + 1;
                }
                await page.waitForTimeout(5000);*/
                await page.locator('.form-control').nth(n).selectOption('No');
                // await page.getByRole('combobox').nth(2).selectOption(publicSchool);
                await page.locator('.form-control').nth(n + 1).selectOption(projectCategory);
                await page.getByRole('button', { name: 'Save and Continue' }).click();
                await page.waitForTimeout(10000);
                if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
                    await page.getByRole('button', { name: 'Revisit' }).click();
                    await page.waitForTimeout(2000);
                    await page.getByRole('button', { name: 'OK' }).last().click();
                    await page.waitForTimeout(2000);
                }

                await page.getByLabel('Name of Owner or Point of').fill(ownerName);
                await page.getByLabel('Street *').first().fill(street);
                await page.getByLabel('City *').first().fill(city);
                await page.getByLabel('State *').first().fill(state);
                await page.getByLabel('Zip Code *').first().fill(zipCode);
                await page.getByLabel('Phone *').first().fill(ownerPhone);
                await page.getByLabel('Email', { exact: true }).fill(ownerEmail);
                await page.locator('.radio').nth(0).click();
                await page.getByLabel('Legal Business Name *').fill(InstallerName);
                await page.getByLabel('Street *').nth(1).fill(InstallerStreet);
                await page.getByLabel('Apartment or Suite').nth(1).fill(InstallerSuite);
                await page.getByLabel('City *').nth(1).fill(InstallerCity);
                await page.getByLabel('State *').nth(1).fill(InstallerState);
                await page.getByLabel('Zip Code *').nth(1).fill(InstallerZip);
                await page.getByLabel('Phone *').nth(1).fill(InstallerPhone);
                await page.getByLabel('Email *').fill(InstallerEmail);
                await page.waitForTimeout(2000);
                await page.getByText('Yes').nth(1).click();

                await page.getByLabel('Number of Graduates of Job').fill(graduates);
                await page.getByRole('button', { name: 'Save and Continue' }).click();

                await page.waitForTimeout(2000);
                if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
                    await page.getByRole('button', { name: 'Revisit' }).click();
                    await page.waitForTimeout(2000);
                    await page.getByRole('button', { name: 'Continue' }).click();

                    await page.waitForTimeout(2000);
                }
                const totaldelete = await page.locator('.btn.mx-button.mx-name-actionButton5.spacing-outer-left.btn-sm.btn-default').count();
                for (let i = 0; i < totaldelete; i++) {
                    await page.locator('.btn.mx-button.mx-name-actionButton5.spacing-outer-left.btn-sm.btn-default').nth(0).click();
                    await page.waitForTimeout(2000);
                }
                await page.locator('.form-control').nth(1).fill(AC);
                await page.getByLabel('Inverter Efficiency (%) *').fill(InverterEfficiency);
                if (!GroundCoverRatio || GroundCoverRatio === 'MISSING!') {
                    GroundCoverRatio = '0.4';
                }
                await page.getByLabel('Ground Cover Ratio *').fill(GroundCoverRatio);
                await page.waitForTimeout(10000);
                await page.getByLabel(MinimumShadingCriteria).check();
                await page.getByRole('combobox').selectOption(BatteryBackup);


                for (let i = 2; i <= numberOfRows; i++) {
                    await page.getByRole('button', { name: 'Add' }).click();
                    const sValue = await data[`s${i}`];
                    const qValue = await data[`q${i}`];
                    const tValue = await data[`t${i}`];
                    const tiValue = await data[`ti${i}`];
                    const oValue = await data[`o${i}`];
                    const bValue = await data[`b${i}`];

                    await page.waitForTimeout(3000);
                    await page.getByPlaceholder('to 1000').click();
                    await page.keyboard.type(sValue, { delay: 500 });
                    await page.getByPlaceholder('Greater than').click();
                    await page.keyboard.type(qValue, { delay: 500 });
                    await page.getByPlaceholder('to 90').click();
                    await page.keyboard.type(tiValue, { delay: 500 });
                    await page.getByPlaceholder('to 359 inclusive').click();
                    await page.keyboard.type(oValue, { delay: 500 });
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
                await page.waitForTimeout(4000);
                await page.goBack();
                await page.getByRole('button', { name: 'Section 5 - REC Estimate' }).click();
                await page.waitForTimeout(2000);
                if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
                    await page.getByRole('button', { name: 'Revisit' }).click();
                    await page.waitForTimeout(2000);
                    await page.getByRole('button', { name: 'OK' }).last().click();
                    await page.waitForTimeout(2000);
                }
                await page.getByLabel('PVWatts').check();
                await page.getByRole('button', { name: 'Calculate' }).click();
                await page.waitForTimeout(5000);
                await page.getByRole('button', { name: 'OK' }).last().click();
                await page.waitForTimeout(1000);
                await page.getByLabel('Custom Capacity Factor').check();
                await page.getByPlaceholder('Example: for 50.5%, enter').fill(CustomCapacityFactor);
                await page.getByLabel('Explanation of Custom').fill(Explanation);
                await page.getByRole('button', { name: 'Calculate' }).click();
                await page.waitForTimeout(5000);
                await page.getByRole('button', { name: 'OK' }).last().click();

                const recerror = await page.locator('input[type="checkbox"]').count();
                console.log(recerror);
                if (recerror > 0) {
                    await page.locator('input[type="checkbox"]').check();
                    await page.waitForTimeout(2000);
                }
                await page.waitForTimeout(1000);
                await page.getByRole('button', { name: 'Save and Continue' }).click();

                await page.waitForTimeout(2000);
                await page.getByRole('button', { name: 'OK' }).click();
                if (await page.getByRole('button', { name: 'Revisit' }).isVisible()) {
                    await page.getByRole('button', { name: 'Revisit' }).click();
                    await page.waitForTimeout(2000);
                    await page.getByRole('button', { name: 'OK' }).last().click();
                    await page.waitForTimeout(2000);
                }


                await page.locator('.form-control').nth(0).selectOption(ABPUtility);
                await page.waitForTimeout(5000);

                if (projectdate === 0) {
                    await page.locator('text="No"').click();

                } else {
                    await page.locator('text="Yes"').click();
                }

                const count = await page.locator('.form-control').count();
                await page.locator('.form-control').nth(1).click();
                if (count === 7) {
                    await page.waitForTimeout(3000);
                    await page.locator('.form-control').nth(1).selectOption(ABPUtility2);
                    await page.waitForTimeout(3000);
                    await page.locator('.form-control').nth(2).click();
                }


                await page.keyboard.press('Control+A'); // Select all text
                await page.keyboard.press('Backspace'); // Clear it
                await page.keyboard.type(formattedDate, { delay: 500 });
                await page.getByRole('button', { name: 'Save and Continue' }).click();
                await page.waitForTimeout(5000);

                uploadResult = {
                    uploaded: true,
                    reason: `Part1 completed for application ${applicationID}`
                };

                console.log(`Part1 completed for system ${systemId}`);
            }
        } else {
            uploadResult = {
                uploaded: false,
                reason: `Disclosure ID ${disclosureID} not found for system ${systemId}`
            };
        }
    await browser.close();

        return {
            success: true,
            systemId,
            disclosureID,
            uploadResult
        };
    } catch (err) {
        await browser.close();
        throw err;
    }
}

module.exports = { runAutomation };
