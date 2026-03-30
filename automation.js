const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');

async function runAutomation(systemId) {
    const browser = await chromium.launch({ headless: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    const page = await browser.newPage();

    try {
        // LOGIN
        await page.goto('https://portal2.carbonsolutionsgroup.com/admin/login');
        await page.waitForTimeout(1000);
        await page.fill('input[type="email"]', process.env.EMAIL);
        await page.fill('input[type="password"]', process.env.PASSWORD);
        await page.click('button:has-text("Login")');
        await page.waitForTimeout(1000);

        // NAVIGATE
        await page.goto(`https://portal2.carbonsolutionsgroup.com/admin/solar_panel_system/${systemId}/checklist?showAll=false#monitoring-info`);
        await page.waitForTimeout(1000);
        const systemName = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(8) > td:nth-child(2)').innerText();
        const address = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(6) > td:nth-child(2)').innerText();
        const ABPID = await page.locator('#tracking-system-info > fieldset > table > tbody > tr:nth-child(3) > td:nth-child(2)').innerText();
        const Installer = await page.locator('#part2-section3 > fieldset > table > tbody > tr:nth-child(2) > td:nth-child(2)').innerText();
        await page.waitForTimeout(1000);
        // LOAD PDF TEMPLATE
        const pdfPath = path.join(__dirname, 'templates', 'demographicwaiver.pdf');
        const pdfBytes = fs.readFileSync(pdfPath);

        const pdfDoc = await PDFDocument.load(pdfBytes);
        const form = pdfDoc.getForm();

        form.getTextField('AV').setText('Carbon Solutions Group');
        form.getTextField('AV_ID').setText('9');
        form.getTextField('CustomerName').setText(systemName);
        form.getTextField('CustomerAddress').setText(address);
        form.getTextField('ProjectApplicationID').setText(ABPID);
        form.getTextField('PreviousAV').setText('N/A');
        form.getTextField('PreviousDesignee').setText(Installer);
        
        form.getTextField('Details').setText(
            Installer === "Solar City STL"
                ? "Removed due to issues"
                : `${Installer} went out of business.`
        );

        form.getTextField('SignerName').setText('Jaspreet Kaur');

        const today = new Date();
        form.getTextField('Date').setText(
            `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`
        );

        // ADD SIGNATURE
        const signaturePath = path.join(__dirname, 'templates', 'signature.png');
        const signatureBytes = fs.readFileSync(signaturePath);
        const signatureImage = await pdfDoc.embedPng(signatureBytes);

        const page1 = pdfDoc.getPages()[0];
        page1.drawImage(signatureImage, {
            x: 70,
            y: 40,
            width: 150,
            height: 50
        });

        const finalPdf = await pdfDoc.save();

        await browser.close();

        return finalPdf;

    } catch (err) {
        await browser.close();
        throw err;
    }
}

module.exports = { runAutomation };