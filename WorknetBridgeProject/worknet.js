const { chromium } = require('playwright');
const readExcel = require('./Excel');
const path = require('path');
  
console.log('script started');
const data = readExcel(path.join(__dirname,'./Bridge_Completion.xlsx'));

(async () => {
  try {
    const browser = await chromium.launch({
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage'
      ]
    });
    console.log('Browser opened.');

    // Isolated browser context (no saved cookies, no account)
    const context = await browser.newContext();
    const page = await context.newPage();
    page.setDefaultTimeout(120_000);
    page.setDefaultNavigationTimeout(120_000);
    console.log('new context and page opened successfully');
    

    // 1. LOGIN PAGE
    try {
      await page.goto('https://apps.illinoisworknet.com/SiteAdministration/Groups/Default');

      await page.getByRole('textbox',{ name: "User name" }).fill(data.Login[0].Username);
      await page.getByRole('textbox',{ name: "Password" }).fill(data.Login[0].Password);
      await page.getByRole('button',{ name: "Sign-In" }).click();
      
      // Wait until login completes
      await page.getByRole('link',{ name: "CEJA/FEJA Programs" }).waitFor({ state: 'visible' });
      console.log('login success');

    } catch (error) {
      console.error('Login failed:', error);
      await browser.close();
      process.exit(1);
    }    

    try { //"Customer Search Page"
      await page.goto('https://apps.illinoisworknet.com/SiteAdministration/ClimateWorks/Admin/Index/');
      await page.locator('#CustomerTable').waitFor({ state: 'visible', timeout: 120000 });
    } catch (error) {
        console.error("Search Page won't load", error);
        await browser.close();
        process.exit(1);
    }
    
    console.log('Search page fully loaded.');

    for (const row of data.Participants) {
      let pages = await context.pages();
      let driver = await pages[pages.length - 1];
      let popup;

      try { //Search for Participant
        await driver.getByRole('textbox',{ name: "Name" }).fill('');
        await driver.getByRole('textbox',{ name: "Name" }).type(`${row.FirstName} ${row.LastName}`);
        await Promise.all([
          driver.waitForLoadState('networkidle'),
          driver.getByRole('button', { name: 'Search' }).click()
        ]);
        const ResultsCount = await driver.locator('#CustomerTable tbody tr').count();
        if (ResultsCount < 1) {
          console.error(`Participant ${row.FirstName} ${row.LastName} not found in Worknet. Is it a typo, or was the participant not registered correctly?`);
        }
        else {
          const CorrectParticipant =  driver.locator('#CustomerTable tbody tr').filter({ hasText: row.WorknetID });
          await CorrectParticipant.getByRole('link', { name: row.LastName }).scrollIntoViewIfNeeded();
          [popup] = await Promise.all([
            driver.waitForEvent('popup'),
            CorrectParticipant.getByRole('link', { name: row.LastName }).click()
          ]);  
          popup.setDefaultTimeout(120_000);
          popup.setDefaultNavigationTimeout(120_000);        
          await popup.waitForLoadState('domcontentloaded');
          console.log('Case file opened');
        }
      } catch (error) {
        console.error("Problem finding or opening the participant's case record. ", error);
        throw error;
      };  
      
      let popup2;

      try { //Edit Participant Service Record to Enter Bridge Completion Details
        const PanelLink = await popup.getByRole('link', { name: 'Training/Services' });
        await PanelLink.scrollIntoViewIfNeeded();
        [popup2] = await Promise.all([
          popup.waitForEvent('popup'),
          PanelLink.click()
        ]);
        popup2.setDefaultTimeout(120_000);
        popup2.setDefaultNavigationTimeout(120_000);
        await popup2.waitForLoadState('domcontentloaded');
        await popup2.locator('#btnAddStepModal').waitFor({ state: 'visible' });

        await Promise.all([
          popup2.waitForURL('**ISTEP/Plan/EditCustomerService**').toBeVisible(),
          popup2.locator('tr', { hasText: "Bridge Training" }).locator('a.edit-service').click()
        ]); 

        await popup2.locator('div.container-fluid').waitFor({ state: 'visible' });
      } catch (error) {
        console.error("Failed navigating to Bridge Training service record. ", error);
        throw error;
      }
      
      try { //Bridge Completion Goal Tab
        await popup2.locator('input#CustomerService_StartDate.form-control.datepicker.hasDatepicker').fill(/*Template Bridge Start Date Field*/);
        await popup2.locator('input#CustomerService_DueDate.form-control.datepicker.hasDatepicker').fill(/*Template Bridge Planned Completion Date Field*/);
        await popup2.locator('#CustomerService_Status').selectOption({ label: 'Successful Completion' });
        await popup2.locator('input#CustomerService_CompletionDate.form-control.datepicker.hasDatepicker').waitFor({ state: 'visible', timeout: 30000 });
        await popup2.locator('input#CustomerService_CompletionDate.form-control.datepicker.hasDatepicker').fill(/*Template Actual Bridge Completion Date Field*/);
        await popup2.locator('input#CustomerService_WeeklyHours.form-control').fill('32.5');
      } catch (error) {
        console.error("Did not properly fill goal details. ", error);
        throw error;
      } 

      try { //Service Provider Tab
        await Promise.all([
          popup2.locator('p').filter({ hasText: 'WH - Dawson Technical' }).locator('#CustomerService_SelectedProvider').waitFor({ state: 'visible' }),
          popup2.getByRole('link',{ name: "Service Provider" }).click()
        ]);
        await popup2.locator('p').filter({ hasText: 'WH - Dawson Technical' }).locator('#CustomerService_SelectedProvider').check();
        //const providerCheckbox = await popup2.getByLabel('WH - Dawson Technical Institute - 3901 South State Street (Career Center) Chicago IL 60609');
      } catch (error) {
        console.error("Failed to select Service Provider. ", error);
        throw error;
      }  
      
      try {//Post-Assessments Tab
        await popup2.getByRole('link', { name: "Post-Assessments" }).click();
        await popup2.getByRole('button', { name: "Add Post-Assessment" }).click();
        await popup2.waitForSelector('text=Add/Edit Post-Assessment');
        await popup2.getByRole('textbox',{ name: "Name" }).fill("Bridge");
        await popup2.locator('input#Score.form-control').fill("100");
        await popup2.getByRole('textbox',{ name: "Date" }).fill("12/19/2025");
        await popup2.getByRole('button', { name: "Save" }).click();
        await popup2.waitForLoadState('networkidle');
      } catch (error) {
        console.error("Failed to enter Post-Assessment. ", error);
        throw error;
      }

      try {//Earned Credentials Tab
        await popup2.getByRole('link',{ name: "Earned Credentials" }).click();
        await popup2.getByLabel('First Aid/CPR Certification').scrollIntoViewIfNeeded();
        await Promise.all([
          popup2.waitForSelector('text=Add/Edit Credential'),
          popup2.getByLabel('First Aid/CPR Certification').click()
        ]);        
        await popup2.getByRole('combobox', { name: 'CredentialSource' }).selectOption('3');
        await popup2.getByRole('textbox',{ name: "DateAttained" }).fill(/*CPR Cert Date*/);
        await popup2.getByRole('textbox',{ name: "Institution" }).fill(`${row.Institution}`);
        await popup2.getByRole('button', { name: "Save" }).click();
        await popup2.waitForLoadState('domcontentloaded');
      } catch (error) {
        console.error("Failed to add First-Aid/CPR Cert. ", error);
        throw error;
      }

      try {//Update Customer Service (Button Is not visible on earned credentials tab)
        await popup2.getByRole('link',{ name: "Status (Default)" }).click();
        await popup2.getByRole('button', { name: "Update Customer Service" }).scrollIntoViewIfNeeded();
        //Case Note Required (Box pop-up)
        await popup2.getByRole('button', { name: "Update Customer Service" }).click();
        await popup2.getByRole('textbox',{ name: "ContactDate" }).fill(/*Bridge Completion Date*/);
        await popup2.getByRole('textbox', { name: "Subject" }).fill("Bridge Training Complete");
        await popup2.getByPlaceholder('Comment').fill("Participant successfully completed Bridge Training.");
        await popup2.getByLabel('Save as case note without sending a message/email').check();
        await popup2.getByRole('button', { name: "Add Case Note" }).click();
      } catch (error) {
        console.error("Failed customer service update / case note submission. ", error);
        throw error;
      }

      console.log(`${row.FirstName} ${row.LastName} Bridge Completion entered.`);

      //Close all the unnecessary browser tabs     
      for (const p = 0; p < pages.length; p++) {
        if (pages[p] !== popup2) {
          await pages[p].close();
        }
      }
      await popup2.bringToFront();
      
      try {//Return to Search Page
        await Promise.all([
          popup2.locator('#CustomerTable').waitFor({ state: 'visible' }),
          popup2.goto('https://apps.illinoisworknet.com/SiteAdministration/ClimateWorks/Admin/Index/')
        ]);
      } catch (error) {
        console.error("Search Page won't load", error);
        await browser.close();
        process.exit(1);
      }
    
    console.log('Search page fully loaded.');
    }
  } finally {  
    await browser.close();
  }  
})();
