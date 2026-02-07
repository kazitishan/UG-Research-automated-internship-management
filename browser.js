require('dotenv').config();

const { chromium } = require('playwright');

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
    
    // Wait 10 seconds before ending
    await page.waitForTimeout(10 * 1000);
    
  } catch (error) {
    console.error('Error:', error.message);
    await page.waitForTimeout(30000);
  }
  
  await browser.close();
})();