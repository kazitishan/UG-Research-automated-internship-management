require('dotenv').config();

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const excelLink = 'https://stevens0-my.sharepoint.com/:x:/g/personal/amansisi_stevens_edu/IQAPeIiCWRGhT5m2yKSlo5mfASQIQJKu7A1rriAmTScU5xU?e=qdky4a';
const portalLink = 'https://stevens0.sharepoint.com/sites/UndergraduateResearch/SitePages/Summer-Internships.aspx';
const email = process.env.EMAIL;
const password = process.env.PASSWORD;

// Parse due date from text like "due Feb. 2, 2026" or "due Feb 2, 2026"
function parseDueDate(text) {
  const match = text.match(/due\s+([A-Za-z]+\.?)\s*(\d{1,2}),?\s*(\d{4})/i);
  if (!match) return null;
  const [, monthStr, day, year] = match;
  const monthMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  const month = monthMap[monthStr.replace('.', '').toLowerCase()];
  if (month === undefined) return null;
  return new Date(parseInt(year), month, parseInt(day));
}

// Sanitize text for use as filename
function sanitizeFilename(text) {
    return text.replace(/[/\\:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 200);
  }
  
  // Get file extension from URL or Content-Type
  function getImageExtension(url, contentType) {
    const urlMatch = url.match(/\.(jpg|jpeg|png|gif|webp)(?:\?|$)/i);
    if (urlMatch) return urlMatch[1].toLowerCase();
    if (contentType?.includes('png')) return 'png';
    if (contentType?.includes('gif')) return 'gif';
    if (contentType?.includes('webp')) return 'webp';
    return 'jpg';
}

// Check if a date is in the past (before today)
function isPastDue(date) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  date.setHours(0, 0, 0, 0);
  return date < today;
}

(async () => {
  const browser = await chromium.launch({ headless: false });
  const page = await browser.newPage();
  
  try {
    await page.goto(portalLink);
    
    // Step 1: Entering Email
    await page.waitForSelector('#i0116');
    await page.fill('#i0116', email);
    await page.waitForTimeout(1 * 1000);
    await page.click('#idSIButton9');
    
    // Step 2: Entering Password
    await page.waitForSelector('#input28');
    await page.fill('#input28', password);
    await page.waitForTimeout(1 * 1000);
    await page.click('input[type="submit"][value="Verify"]');
    
    // Step 3: Click Okta Verify button (Push Notification)
    await page.click('a[aria-label*="push notification to the Okta Verify app"]');
    
    // Step 4: Wait for "Summer Research Internships" text to show up to know that we have loaded into the page
    await page.waitForSelector('text=Summer Research Internships', { timeout: 60000 });
    
    // Step 5: Zoom out so that the rest of the content on the page can load
    await page.evaluate(() => {
        document.body.style.zoom = '0.05';
    });
    
    // Step 6: Extract all internship cards
    await page.waitForSelector('[data-automation-id="grid-layout"][aria-label*="External Research Internships"]', { timeout: 60 * 1000 });
    const internships = await page.$$eval(
      '[data-automation-id="grid-layout"][aria-label*="External Research Internships"] [role="listitem"] a[href]',
      (links) =>
        links.map((link) => {
          const titleEl = link.querySelector('[data-automation-id="quick-links-item-title"]');
          const imgEl = link.querySelector('img');
          return {
            Text: titleEl ? titleEl.textContent.trim() : '',
            Link: link.href,
            Img: imgEl ? imgEl.src : '',
          };
        })
    );
    await page.evaluate(() => { // set zoom back to normal
        document.body.style.zoom = '1.00';
    });
    
    // Step 7: Filter to inactive internships (due date is past)
    const inactive = [];
    const today = new Date();
    for (const internship of internships) {
      const dueDate = parseDueDate(internship.Text);
      if (dueDate && isPastDue(dueDate)) {
        inactive.push(internship);
      }
    }
    
    console.log('Inactive internships (past due date):');
    console.log(JSON.stringify(inactive, null, 2));

    // Step 8: Download images to images folder
    const imagesDir = path.join(__dirname, 'images');
    fs.mkdirSync(imagesDir, { recursive: true });

    const request = page.context().request;
    const seenNames = new Set();

    for (const internship of inactive) {
      if (!internship.Img) continue;

      let baseName = sanitizeFilename(internship.Text);
      if (seenNames.has(baseName)) {
        let i = 1;
        while (seenNames.has(`${baseName} (${i})`)) i++;
        baseName = `${baseName} (${i})`;
      }
      seenNames.add(baseName);

      try {
        const response = await request.get(internship.Img);
        const contentType = response.headers()['content-type'];
        const ext = getImageExtension(internship.Img, contentType);
        const filePath = path.join(imagesDir, `${baseName}.${ext}`);

        await response.body().then((buffer) => fs.writeFileSync(filePath, buffer));
        console.log(`Downloaded: ${baseName}.${ext}`);
      } catch (err) {
        console.error(`Failed to download ${internship.Text}:`, err.message);
      }
    }

    // Step 9: Press the edit button
    await page.waitForSelector('[data-automation-id="pageCommandBarEditButton"]', { timeout: 10000 });
    await page.click('[data-automation-id="pageCommandBarEditButton"]');
    await page.waitForTimeout(1000);

    // Step 10: Expand the inactive/closed internships section
    await page.click('button[aria-label*="inactive/closed"]');

    // Step 11: Click the Quick links hover label, then Add links button
    await page.click('div[data-automation-id="CanvasControl"][id="5597bd71-4df8-46e5-9ed6-4eb29c72e2e8"]');
    await page.waitForSelector('[data-automation-id="quickLinksTopActionsAddLinks"]', { timeout: 10000 });
    await page.click('[data-automation-id="quickLinksTopActionsAddLinks"]');
    
    // Step 12: Click on "From a link" button
    await page.waitForSelector('button[title="From a link"]', { timeout: 10000 });
    await page.click('button[title="From a link"]');
    await page.waitForTimeout(1000);

    // Step 13: Wait for and fill in the URL textarea with the first inactive internship link
    await page.waitForSelector('textarea[data-automation-id="filePickerLinkTextField"]', { timeout: 10000 });
    await page.fill('textarea[data-automation-id="filePickerLinkTextField"]', inactive[0].Link);

    // Step 14: Click the "Add" button
    await page.waitForTimeout(1000);
    await page.getByRole('button').nth(-2).click();
    await page.waitForTimeout(2000);

    // Step 15: Replace the default text in the text input with the internship text
    await page.waitForSelector('input[type="text"][id^="field-"]', { timeout: 10000 });
    
    // Clear the existing value and fill with the internship text
    const textInput = page.locator('input[type="text"][id^="field-"]').first();
    await textInput.fill('');
    await textInput.fill(inactive[0].Text);
    await page.waitForTimeout(1000);

    // Step 16: Click the "Open in new tab" checkbox
    await page.waitForSelector('input[data-automation-id="openInNewTabCpanetoggle"]', { timeout: 10000 });
    const checkbox = page.locator('input[data-automation-id="openInNewTabCpanetoggle"]');
    
    // Check if it's already checked, if not, check it
    const isChecked = await checkbox.isChecked();
    if (!isChecked) {
      await checkbox.click();
    }
    
    await page.waitForTimeout(1000);

    // Wait 10000 seconds before ending
    await page.waitForTimeout(10000 * 1000);
    
  } catch (error) {
    console.error('Error:', error.message);
    await page.waitForTimeout(30000);
  }
  
  await browser.close();
})();