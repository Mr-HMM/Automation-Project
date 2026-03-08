import { test, expect } from '@playwright/test';
import { readExcel, writeExcel  } from './Excel';

console.log('script started');
const data = readExcel(/*local path to excel template*/);

test('lisc', async ({ page }) => {
  test.setTimeout(1800000);
    
  try {
    await page.goto(/*URL to grant funder salesforce platform*/);
    await page.waitForLoadState("domcontentloaded");
  } catch (error) {
    console.error("Failed loading first page. ", error);
    throw error;
  }

  try { //login
    await page.getByRole('textbox', { name: 'Username' }).fill(data.Login[0].Username);
    await page.getByRole('textbox', { name: 'Password' }).fill(data.Login[0].Password);
    await Promise.all([
      page.locator('.themeHeaderBottomRow').waitFor({ state: 'visible' }),
      page.getByRole('button', { name: 'Log in' }).click()
    ]);      
  } catch (error) {
      console.error("Login failed. ", error);
      throw error;
  };
  console.log("Logged in.");

  try { //Navigate to records for the correct training cohort
    await Promise.all([
      page.getByRole('button', { name: 'Select a List View: Groups &' }).waitFor({ state: 'visible' }),
      page.getByRole('menuitem', { name: 'Groups and Classes' }).click()
    ]);
    await page.getByRole('button', { name: 'Select a List View: Groups &' }).click();
    await page.locator('span').filter({ hasText: 'All Groups & Classes' }).first().click();
    await page.waitForLoadState("domcontentloaded");
    await page.getByRole('searchbox', { name: 'Search this list...' }).click();
    await page.getByRole('searchbox', { name: 'Search this list...' }).fill('current');
    await Promise.all([
      page.waitForLoadState("domcontentloaded"),
      page.getByRole('searchbox', { name: 'Search this list...' }).press('Enter')
    ]);  
    await Promise.all([
      page.getByRole('tab', { name: 'Assignments and Services' }).waitFor({ state: 'visible' }),
      page.getByRole('link', { name: /*Insert name of cohort or at least part of it*/ }).click()
    ]);
    await page.getByRole('tab', { name: 'Assignments and Services' }).click();
    await page.getByRole('link', { name: 'View All Group / Class' }).scrollIntoViewIfNeeded();
    await Promise.all([
      page.waitForLoadState("domcontentloaded"),
      page.getByRole('link', { name: 'View All Group / Class' }).click()
    ]);
    await expect(page.getByRole('heading', { name: 'Group / Class Assignments' })).toBeVisible();
  } catch (error) {
      console.error("Problem pulling up the cohort roster. ", error);
      throw error;
  };
  console.log("Cohort file opened.");
  
  
  for (let x = 0; x < data.Trainees.length; x++) {//Loop through each trainee for individual service record entries
    let Trainee = data.Trainees[x];

    if (!Trainee['Completed?']) {
      const href = await page.getByRole('link', { name: `${Trainee.Name} FOC ` }).getAttribute('href');

      const [page2] = await Promise.all([
        page.waitForEvent('popup'),
        page.evaluate(url => window.open(url, '_blank'), href)
      ]);
      
      let page3;
      [page3] = await Promise.all([
        page2.waitForEvent('popup'),        
        page2.getByRole('button', { name: 'Record Service' }).click()
      ]);
      await page3.waitForLoadState('domcontentloaded');

      for (let y = 0; y < data.Sessions.length; y++) {//Loop through each individual workshop service entry for the trainee 
        let Session = data.Sessions[y];
        await expect(page3.getByText('Select an ECM Service Entry')).toBeVisible({timeout: 30000});

        try {
          switch (Session.WorkshopType) {
            case 'Financial Workshop':
              await page3.getByLabel('Service Entry Type').selectOption('a171Q00000XbwUkQAJ');
              await Promise.all([
                page3.waitForLoadState("domcontentloaded"),
                page3.getByRole('button', { name: 'Next' }).click()
              ]);
              await expect(page3.getByRole('heading', { name: 'Financial Coaching -' })).toBeVisible();
              await page3.getByRole('textbox', { name: 'Date' }).click();
              await page3.getByRole('textbox', { name: 'Date' }).fill('');
              await page3.getByRole('textbox', { name: 'Date' }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date' }).fill(Session.Date);
              await page3.getByLabel('*Start Time').selectOption({ value: Session.Time});
              await page3.getByRole('textbox', { name: 'Staff Person' }).click();
              await page3.getByRole('textbox', { name: 'Staff Person' }).fill(/*The author's name*/);
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption("In person");
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(Session.Duration);
              await page3.getByLabel('*Digital Skills Training/').selectOption('No');
              await page3.getByLabel('*Worked on Goals').selectOption('No');

              switch (Session.Topic) {
                case 'Financial Literacy':
                  await page3.locator('[id="img_j_id0:theForm:thePageBlock:j_id970:j_id971:client"]').click();
                  await page3.getByRole('cell', { name: 'Apartment Rental', exact: true }).getByRole('listbox').selectOption('0');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.locator('[id="img_j_id0:theForm:thePageBlock:j_id1200:j_id1201:client"]').click();
                  await page3.getByRole('cell', { name: 'Stored Value Card (Prepaid Card)', exact: true }).getByRole('listbox').selectOption('1');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.locator('textarea[name="j_id0:theForm:thePageBlock:j_id1200:j_id1201:client:clientList:0:j_id1223:1:j_id1225"]').click();
                  await page3.locator('textarea[name="j_id0:theForm:thePageBlock:j_id1200:j_id1201:client:clientList:0:j_id1223:1:j_id1225"]').fill('Consumer habits; money mindset; earning methods');
                  break;        
                  
                case 'Budgeting/Money Management':
                  await page3.getByRole('button', { name: 'Show Section -' }).nth(3).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('15');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.locator('[id="img_j_id0:theForm:thePageBlock:j_id970:j_id971:client"]').click();
                  await page3.getByRole('cell', { name: 'Apartment Rental', exact: true }).getByRole('listbox').selectOption('0');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Financial Management for Homeowners', exact: true }).getByRole('listbox').selectOption('1');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.locator('[id="img_j_id0:theForm:thePageBlock:j_id1545:j_id1546:client"]').click();
                  await page3.getByRole('cell', { name: 'Certificate of Deposit', exact: true }).getByRole('listbox').selectOption('4');
                  await page3.getByRole('link', { name: 'Add' }).nth(2).click();
                  await page3.locator('textarea[name="j_id0:theForm:thePageBlock:j_id1545:j_id1546:client:clientList:0:j_id1568:1:j_id1570"]').click();
                  await page3.locator('textarea[name="j_id0:theForm:thePageBlock:j_id1545:j_id1546:client:clientList:0:j_id1568:1:j_id1570"]').fill('Budgeting; Spending habits; bank accounts');
                  break;

                case 'Credit/Credit Scores':
                  await page3.getByRole('button', { name: 'Show Section -' }).first().click();
                  await page3.getByRole('cell', { name: 'ChexSystems Error', exact: true }).getByRole('listbox').selectOption('2');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.getByRole('cell', { name: 'ChexSystems Error', exact: true }).getByRole('listbox').selectOption('5');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.getByRole('cell', { name: 'ChexSystems Error', exact: true }).getByRole('listbox').selectOption('8');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.getByRole('button', { name: 'Show Section -' }).nth(2).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('2');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('3');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('4');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('7');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('12');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('14');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('20');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('21');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  break;
              
                case 'Banking':
                  await page3.getByRole('button', { name: 'Show Section -' }).nth(3).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('12');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('4');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.getByRole('cell', { name: 'Auto-Title Loan', exact: true }).getByRole('listbox').selectOption('14');
                  await page3.getByRole('link', { name: 'Add' }).click();
                  await page3.locator('[id="img_j_id0:theForm:thePageBlock:j_id970:j_id971:client"]').click();
                  await page3.getByRole('cell', { name: 'Apartment Rental', exact: true }).getByRole('listbox').selectOption('0');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.getByRole('cell', { name: 'Financial Management for Homeowners', exact: true }).getByRole('listbox').selectOption('1');
                  await page3.getByRole('link', { name: 'Add' }).nth(1).click();
                  await page3.locator('[id="img_j_id0:theForm:thePageBlock:j_id1545:j_id1546:client"]').click();
                  await page3.getByRole('cell', { name: 'Certificate of Deposit', exact: true }).getByRole('listbox').selectOption('3');
                  await page3.getByRole('link', { name: 'Add' }).nth(2).click();
                  break;
              }
              if (!data.Sessions[y+1]) {
                await Promise.all([
                  page3.waitForLoadState("networkidle"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id2463:bottom:theButtons:j_id2464:saveButton"]').click()
                ]);
                await expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${Trainee.Name} FOC` })).toBeVisible();
                await page3.close();
              }    
              else { 
                await Promise.all([
                  page3.waitForLoadState("domcontentloaded"), 
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id2463:bottom:theButtons:j_id2464:saveAndNewButton"]').click()
                ]);
              }
              break;

            case 'Employment/Education Workshop':
              await page3.getByLabel('Service Entry Type').selectOption('a1736000001cu6GAAQ');
              await page3.getByRole('button', { name: 'Next' }).click();
              await page3.waitForLoadState("domcontentloaded");
              await expect(page3.getByRole('heading', { name: 'Employment Counseling -' })).toBeVisible();
              await page3.getByRole('textbox', { name: 'Date / Time' }).click();
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill('');
              await page3.getByRole('textbox', { name: 'Date / Time' }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill(`${Session.Date} ${Session.Time}`);
              await page3.getByLabel('*Start Time').selectOption({ value: Session.Time });
              await page3.getByRole('textbox', { name: 'Date', exact: true }).click();
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill('');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill(Session.Date);
              if (Session.Facilitator) {
                await page3.getByRole('textbox', { name: 'Staff Person' }).click();
                await page3.getByRole('textbox', { name: 'Staff Person' }).fill(Session.Facilitator);
              }
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption("In person");
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${Session.Duration}`);
              await page3.getByLabel('*Digital Skills Training/').selectOption('No');
              await page3.getByRole('button', { name: 'Show Section -' }).first().click();
              await page3.locator('select[name="j_id0:theForm:thePageBlock:j_id165:j_id166:client:clientList:0:j_id188:1:j_id190"]').selectOption('Initiated/continued search');
              if (!data.Sessions[y+1]) {
                await Promise.all([
                  page3.waitForLoadState("networkidle"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id1083:bottom:theButtons:j_id1084:saveButton"]').click()
                ]);
                await page3.waitForLoadState("networkidle",{ timeout: 60000 });
                await page3.close();
              }
              else await Promise.all([ 
                page3.waitForLoadState("domcontentloaded"),
                page3.locator('input[name="j_id0:theForm:thePageBlock:j_id1083:bottom:theButtons:j_id1084:saveAndNewButton"]').click()
              ]);
              break;

            case 'Income Supports Workshop':
              await page3.getByLabel('Service Entry Type').selectOption('a1736000001cuPIAAY');
              await page3.getByRole('button', { name: 'Next' }).click();
              await page3.waitForLoadState("domcontentloaded");
              await expect(page3.getByRole('heading', { name: 'Income Supports Counseling -' })).toBeVisible();
              await page3.getByLabel('*Start Time').selectOption({ value: Session.Time});
              await page3.getByRole('textbox', { name: 'Date', exact: true }).click();
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill('');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill(Session.Date);           
              await page3.getByRole('textbox', { name: 'Duration (Hours)' }).click();
              await page3.getByRole('textbox', { name: 'Duration (Hours)' }).fill('');
              await page3.getByRole('textbox', { name: 'Duration (Hours)' }).fill('.33');
              await page3.getByRole('textbox', { name: 'Date / Time' }).click();
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill('');
              await page3.getByRole('textbox', { name: 'Date / Time' }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill(`${Session.Date} ${Session.Time}`);
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption("In person");
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${Session.Duration}`);
              await page3.getByLabel('*Digital Skills Training/').selectOption('No');
              if (!data.Sessions[y+1]) {
                await Promise.all([
                  page3.waitForLoadState("networkidle"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id1930:bottom:theButtons:j_id1931:saveButton"]').click()
                ]);
                await expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${Trainee.Name} FOC` })).toBeVisible();
                await page3.close();
              }
              else await Promise.all([
                page3.waitForLoadState("domcontentloaded"),
                page3.locator('input[name="j_id0:theForm:thePageBlock:j_id1930:bottom:theButtons:j_id1931:saveAndNewButton"]').click()
              ]);
              break;

            case 'Digital Literacy Workshop':
              await page3.getByLabel('Service Entry Type').selectOption('a171Q00000WXJIrQAP');
              await page3.getByRole('button', { name: 'Next' }).click();
              await page3.waitForLoadState("domcontentloaded");
              await expect(page3.getByRole('heading', { name: 'Digital Navigation -' })).toBeVisible();
              await page3.getByRole('textbox', { name: 'Date' }).fill(Session.Date);
              await page3.getByRole('textbox', { name: 'Date' }).press('Escape');
              await page3.getByLabel('*Start Time').selectOption({ value: Session.Time});
              if (Session.Facilitator) {
                await page3.getByRole('textbox', { name: 'Staff Person' }).click();
                await page3.getByRole('textbox', { name: 'Staff Person' }).fill(Session.Facilitator);
              }
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption("In person");
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${Session.Duration}`);
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByRole('button', { name: 'Show Section -' }).first().click();
              await page3.getByRole('cell', { name: 'Digital Skills Training', exact: true }).getByRole('listbox').selectOption('0');
              await page3.getByRole('link', { name: 'Add' }).click();
              await page3.getByRole('cell', { name: 'Getting a computer (desktop, laptop or tablet)', exact: true }).getByRole('listbox').selectOption('1');
              await page3.getByRole('link', { name: 'Add' }).click();
              await page3.getByRole('button', { name: 'Show Section -' }).nth(1).click();
              await page3.locator('select[name="j_id0:theForm:thePageBlock:j_id395:j_id396:client:clientList:0:j_id418:0:j_id420"]').selectOption('Gifted to user');
              await page3.locator('select[name="j_id0:theForm:thePageBlock:j_id395:j_id396:client:clientList:0:j_id418:1:j_id420"]').selectOption('Laptop');
              await page3.locator('select[name="j_id0:theForm:thePageBlock:j_id395:j_id396:client:clientList:0:j_id418:2:j_id420"]').selectOption('Google Chromebook');
              if (!data.Sessions[y+1]) {
                await Promise.all([
                  page3.waitForLoadState("networkidle"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id968:bottom:theButtons:j_id969:saveButton"]').click()
                ]);
                await expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${Trainee.Name} FOC` })).toBeVisible();
                await page3.close();
                Trainee['Completed?'] = 'Y';
              }
              else await Promise.all([
                page3.waitForLoadState("domcontentloaded"),
                page3.locator('input[name="j_id0:theForm:thePageBlock:j_id968:bottom:theButtons:j_id969:saveAndNewButton"]').click()
              ]);
              break;
          }
        } catch (error) {
          console.error(`Failed to enter ${Session.Date} service for ${Trainee.Name}`, error);
        }
      }
      
      console.log(`${Trainee.Name} service record entries completed?`);
      await page2.close();
    }
  }  
  writeExcel('C:/Users/Harold/Documents/CurrentWater_sessions.xlsx', data);
});