import { test, expect } from '@playwright/test';
import { readExcel, writeExcel  } from './Excel';

console.log('script started');
const data = readExcel('C:/Users/Harold/Documents/ceja_winter_sessions.xlsx');

test('lisc', async ({ page }) => {
  test.setTimeout(600000);
    
  try {
    await page.goto('https://liscecdev.my.site.com/FOCCommunity/s/');
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
    await page.getByRole('searchbox', { name: 'Search this list...' }).fill('ceja');
    await Promise.all([
      page.waitForLoadState("domcontentloaded"),
      page.getByRole('searchbox', { name: 'Search this list...' }).press('Enter')
    ]);  
    await Promise.all([
      page.getByRole('tab', { name: 'Assignments and Services' }).waitFor({ state: 'visible' }),
      page.getByRole('link', { name: 'CEJA Winter' }).click()
    ]);
  } catch (error) {
      console.error("Problem navigating to the correct training group. ", error);
      throw error;
  };
  console.log("Cohort file opened.");
  
  let page1;

  try {
    await expect(page.locator('records-highlights2 lightning-formatted-text').filter({ hasText: 'CEJA Winter 2025' })).toBeVisible();
    [page1] = await Promise.all([
      page.waitForEvent('popup'),
      page.getByRole('button', { name: 'Schedule Session' }).click()  
    ]);
  } catch (error) {
      console.error("Training group not loading right. Can't schedule sessions. ", error);
      throw error;
  };  

  for (let i = 0; i < data.Sessions.length; i++) { //Loop session data entry for each session listed in excel
    if (!data.Sessions[i]['Entered?']) {
      try {
        await expect(page1.getByRole('heading', { name: 'Workshop/Class -' })).toBeVisible();
        await page1.getByLabel('*Start Time').selectOption({ value: data.Sessions[i].Time });
        await page1.getByRole('textbox', { name: 'Date', exact: true }).click();
        await page1.getByRole('textbox', { name: 'Date', exact: true }).fill('');
        await page1.getByRole('textbox', { name: 'Date', exact: true }).press('Escape');
        await page1.getByRole('textbox', { name: 'Date', exact: true }).fill(data.Sessions[i].Date);
        await page1.getByRole('textbox', { name: 'Date / Time' }).click();
        await page1.getByRole('textbox', { name: 'Date / Time' }).fill(`${data.Sessions[i].Date} ${data.Sessions[i].Time}`);
        await page1.getByRole('textbox', { name: 'Date / Time' }).press('Escape');
        await page1.getByLabel('Status').selectOption('Recorded Attendance');
        await page1.getByRole('textbox', { name: 'Duration (Minutes)' }).click();
        await page1.getByRole('textbox', { name: 'Duration (Minutes)' }).fill(`${data.Sessions[i].Duration}`);
        await page1.getByLabel('Type of Workshop/Class', { exact: true }).scrollIntoViewIfNeeded();
        await page1.getByLabel('Type of Workshop/Class', { exact: true }).selectOption(data.Sessions[i].WorkshopType);
        if (data.Sessions[i].Topic) {
          await page1.getByLabel('Topic/Course', { exact: true }).scrollIntoViewIfNeeded();
          await page1.getByLabel('Topic/Course', { exact: true }).selectOption(data.Sessions[i].Topic);
        }
        await page1.getByLabel('*Curriculum Developer/Owner').scrollIntoViewIfNeeded();
        await page1.getByLabel('*Curriculum Developer/Owner').selectOption(data.Sessions[i].CurriculumDev);
        await page1.getByLabel('Training Method').scrollIntoViewIfNeeded();
        if (data.Sessions[i].WorkshopType === 'Digital Literacy Workshop') {
          await page1.getByRole('checkbox', { name: 'Curriculum contains digital' }).check();
        }
        await page1.getByLabel('Training Method').selectOption(data.Sessions[i].TrainingMethod);
        if (data.Sessions[i].Other_workshop_type) {
          await page1.getByRole('textbox', { name: 'If other type of workshop/' }).click();
          await page1.getByRole('textbox', { name: 'If other type of workshop/' }).fill(data.Sessions[i].OtherWorkshopType);
        }  
        if ((data.Sessions[i].CurriculumDev === 'Other') || data.Sessions[i].OtherCurriculumDev) {
          await page1.getByRole('textbox', { name: 'If other curriculum developer' }).click();
          await page1.getByRole('textbox', { name: 'If other curriculum developer' }).fill(data.Sessions[i].OtherCurriculumDev);
        }
        if (data.Sessions[i].Facilitator) {
          await page1.getByRole('textbox', { name: 'Workshop Facilitator' }).click();
          await page1.getByRole('textbox', { name: 'Workshop Facilitator' }).fill(data.Sessions[i].Facilitator);
        }
        await Promise.all([
          page1.waitForLoadState("domcontentloaded"),
          page1.locator('select[name="j_id0:theForm:thePageBlock:j_id207:j_id208:mdsPB:j_id225:1:j_id226"]').selectOption('Attended')
        ]);
        if (!(data.Sessions[i+1])) {
          await Promise.all([
            page1.waitForLoadState("domcontentloaded"),
            page1.locator('input[name="j_id0:theForm:thePageBlock:j_id1170:bottom:theButtons:j_id1171:saveButton"]').click()
          ]);  
          await expect(page1.locator('records-highlights2 lightning-formatted-text').filter({ hasText: 'CEJA Winter 2025' })).toBeVisible();
          data.Sessions[i].Complete = 'Y';
          await page1.close();
        }
        else {
          await Promise.all([
            page1.waitForLoadState("domcontentloaded"),
            page1.locator('input[name="j_id0:theForm:thePageBlock:j_id1170:bottom:theButtons:j_id1171:saveAndNewButton"]').click()
          ]);
          await expect(page1.getByLabel('Group Service Entry Type')).toBeVisible();
          data.Sessions[i].Complete = 'Y';
          await page1.getByLabel('Group Service Entry Type').selectOption('a1B36000001lgTfEAI');
          await page1.getByRole('button', { name: 'Next' }).click();
        }

      } catch (error) {
          console.error("Failed to add session record.", error);
          writeExcel('./ceja_winter_sessions.xlsx', data);
          throw error;
      }
    }
  }

  writeExcel('C:/Users/Harold/Documents/ceja_winter_sessions.xlsx', data);
  console.log("Sessions uploaded.");

  await page.getByRole('tab', { name: 'Assignments and Services' }).click();
  await page.getByRole('link', { name: 'View All Group / Class' }).scrollIntoViewIfNeeded();
  await page.getByRole('link', { name: 'View All Group / Class' }).click();
  
  for (let x = 0; x < data.Trainees.length; x++) {//Loop through each trainee for individual service record entries
    if (!data.Trainees[x]['Completed?']) {
      const [page2] = await Promise.all([
        page.waitForEvent('popup'),
        page.getByRole('link', { name: `${data.Trainees[x].Name} FOC ` }).click({ modifiers: ['Control'] })
      ]);
      let page3;
      [page3] = await Promise.all([
        page2.waitForEvent('popup'),        
        page2.getByRole('button', { name: 'Record Service' }).click()
      ]);
      await page3.waitForLoadState('domcontentloaded');

      for (let y = 0; y < data.Sessions.length; y++) {//Loop through each individual workshop service entry for the trainee 
        await expect(page3.getByText('Select an ECM Service Entry')).toBeVisible();

        try {
          switch (data.Sessions[y].WorkshopType) {
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
              await page3.getByRole('textbox', { name: 'Date' }).fill(data.Sessions[y].Date);
              await page3.getByLabel('*Start Time').selectOption({ value: data.Sessions[y].Time});
              await page3.getByRole('textbox', { name: 'Staff Person' }).click();
              await page3.getByRole('textbox', { name: 'Staff Person' }).fill('Harold Merrell');
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption(data.Sessions[y].TrainingMethod);
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${`${data.Sessions[y].Duration}`}`);
              await page3.getByLabel('*Digital Skills Training/').selectOption('No');
              await page3.getByLabel('*Worked on Goals').selectOption('No');

              switch (data.Sessions[y].Topic) {
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
                  page3.waitForLoadState("domcontentloaded"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id2463:bottom:theButtons:j_id2464:saveButton"]').click()
                ]);
                await expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${data.Trainees[x].Name} FOC` })).toBeVisible();
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
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill(`${data.Sessions[y].Date} ${data.Sessions[y].Time}`);
              await page3.getByLabel('*Start Time').selectOption({ value: data.Sessions[y].Time });
              await page3.getByRole('textbox', { name: 'Date', exact: true }).click();
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill('');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill(data.Sessions[y].Date);
              if (data.Sessions[y].Facilitator) {
                await page3.getByRole('textbox', { name: 'Staff Person' }).click();
                await page3.getByRole('textbox', { name: 'Staff Person' }).fill(data.Sessions[y].Facilitator);
              }
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption(data.Sessions[y].TrainingMethod);
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${data.Sessions[y].Duration}`);
              await page3.getByLabel('*Digital Skills Training/').selectOption('No');
              await page3.getByRole('button', { name: 'Show Section -' }).first().click();
              await page3.locator('select[name="j_id0:theForm:thePageBlock:j_id165:j_id166:client:clientList:0:j_id188:1:j_id190"]').selectOption('Initiated/continued search');
              if (!data.Sessions[y+1]) {
                await Promise.all([
                  page3.waitForLoadState("domcontentloaded"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id1083:bottom:theButtons:j_id1084:saveButton"]').click()
                ]);
                expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${data.Trainees[x].Name} FOC` })).toBeVisible();
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
              await page3.getByLabel('*Start Time').selectOption({ value: data.Sessions[y].Time});
              await page3.getByRole('textbox', { name: 'Date', exact: true }).click();
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill('');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date', exact: true }).fill(data.Sessions[y].Date);           
              await page3.getByRole('textbox', { name: 'Duration (Hours)' }).click();
              await page3.getByRole('textbox', { name: 'Duration (Hours)' }).fill('');
              await page3.getByRole('textbox', { name: 'Duration (Hours)' }).fill('.33');
              await page3.getByRole('textbox', { name: 'Date / Time' }).click();
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill('');
              await page3.getByRole('textbox', { name: 'Date / Time' }).press('Escape');
              await page3.getByRole('textbox', { name: 'Date / Time' }).fill(`${data.Sessions[y].Date} ${data.Sessions[y].Time}`);
              await page3.getByLabel('*Contact with').selectOption('Client');
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption(data.Sessions[y].TrainingMethod);
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${data.Sessions[y].Duration}`);
              await page3.getByLabel('*Digital Skills Training/').selectOption('No');
              if (!data.Sessions[y+1]) {
                await Promise.all([
                  page3.waitForLoadState("domcontentloaded"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id1930:bottom:theButtons:j_id1931:saveButton"]').click()
                ]);
                await expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${data.Trainees[x].Name} FOC` })).toBeVisible();
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
              await page3.getByRole('textbox', { name: 'Date' }).fill(data.Sessions[y].Date);
              await page3.getByRole('textbox', { name: 'Date' }).press('Escape');
              await page3.getByLabel('*Start Time').selectOption({ value: data.Sessions[y].Time});
              if (data.Sessions[y].Facilitator) {
                await page3.getByRole('textbox', { name: 'Staff Person' }).click();
                await page3.getByRole('textbox', { name: 'Staff Person' }).fill(data.Sessions[y].Facilitator);
              }
              await page3.getByLabel('*Reach person you attempted').selectOption('Yes');
              await page3.getByLabel('*Contact/Location Method').selectOption(data.Sessions[y].TrainingMethod);
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).click();
              await page3.getByRole('textbox', { name: '* Duration (Minutes)' }).fill(`${data.Sessions[y].Duration}`);
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
                  page3.waitForLoadState("domcontentloaded"),
                  page3.locator('input[name="j_id0:theForm:thePageBlock:j_id968:bottom:theButtons:j_id969:saveButton"]').click()
                ]);
                await expect(page3.locator('records-highlights2 lightning-formatted-text').filter({ hasText: `${data.Trainees[x].Name} FOC` })).toBeVisible();
                await page3.close();
              }
              else await Promise.all([
                page3.waitForLoadState("domcontentloaded"),
                page3.locator('input[name="j_id0:theForm:thePageBlock:j_id968:bottom:theButtons:j_id969:saveAndNewButton"]').click()
              ]);
              break;
          }
        } catch (error) {
          console.error(`Failed to enter ${data.Sessions[y].Date} service for ${data.Trainees[x].Name}`, error);
        }
      }
      data.Trainees[x]['Completed?'] = 'Y';
      console.log(`${data.Trainees[x].Name} service record entries completed?`);
      await page2.close();
    }
  }  
  writeExcel('C:/Users/Harold/Documents/ceja_winter_sessions.xlsx', data);
});