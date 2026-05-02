// scrape.js — Salondata Weekly Stats scraper
// Logs into reports.salondata.com, downloads Employee_Stats.xlsx and
// Employee_Return_Stats.xlsx via the portal's Excel export button.
//
// Env: SALONDATA_USERNAME, SALONDATA_PASSWORD
// Output: Employee_Stats.xlsx, Employee_Return_Stats.xlsx (in cwd)

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'https://reports.salondata.com/static/reports/index.html';
const USERNAME = process.env.SALONDATA_USERNAME;
const PASSWORD = process.env.SALONDATA_PASSWORD;
if (!USERNAME || !PASSWORD) {
  console.error('Missing SALONDATA_USERNAME or SALONDATA_PASSWORD env vars.');
  process.exit(1);
}

const STORES = '3750,3800,3826,4216';

function fmt(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

// Most recent Friday on or before today (week-ending date).
function mostRecentFriday() {
  const d = new Date();
  const back = (d.getDay() - 5 + 7) % 7; // Friday = 5
  d.setDate(d.getDate() - back);
  return d;
}

const FRIDAY = process.env.WEEK_END ? new Date(process.env.WEEK_END) : mostRecentFriday();
const SATURDAY = new Date(FRIDAY); SATURDAY.setDate(FRIDAY.getDate() - 6); // prior Saturday
const FRI = fmt(FRIDAY);
const SAT = fmt(SATURDAY);

// Hash routes pulled from the portal URLs.
// - wecustomerreturn:    start = end = Friday
// - performance:         start = prior Saturday, end = Friday
const REPORTS = [
  {
    label: 'Employee Stats (performance)',
    hash: `#performance:store=${STORES}&start=${SAT}&end=${FRI}`,
    output: 'Employee_Stats.csv',
  },
  {
    label: 'Employee Customer Return',
    hash: `#wecustomerreturn:store=${STORES}&start=${FRI}&end=${FRI}`,
    output: 'Employee_Return_Stats.csv',
  },
];

const SCRIPT_DIR = __dirname;
const SCREENSHOT_DIR = path.join(SCRIPT_DIR, 'screenshots');
if (!fs.existsSync(SCREENSHOT_DIR)) fs.mkdirSync(SCREENSHOT_DIR);

async function login(page) {
  console.log('Opening Salondata...');
  await page.goto(BASE_URL, { waitUntil: 'networkidle', timeout: 30000 });
  await page.waitForTimeout(2500);

  const emailField = await page.$('input[type="email"]');
  if (!emailField) {
    console.log('  Already logged in (no email field).');
    return;
  }

  console.log('  Logging in...');
  await page.fill('input[type="email"]', USERNAME);
  await page.fill('input[type="password"]', PASSWORD);

  const submitted = await page.evaluate(() => {
    const btn = document.querySelector(
      'button[type="submit"], input[type="submit"], button.login, button.signin'
    );
    if (btn) { btn.click(); return true; }
    const allButtons = Array.from(document.querySelectorAll('button'));
    const textBtn = allButtons.find((b) => /log|sign/i.test(b.innerText));
    if (textBtn) { textBtn.click(); return true; }
    return false;
  });
  if (!submitted) await page.keyboard.press('Enter');
  await page.waitForTimeout(4000);
  console.log('  Logged in.');
}

async function downloadReport(page, report) {
  console.log(`\n[${report.label}] week ending ${FRI}`);
  console.log(`  navigating to ${report.hash}`);
  await page.goto(BASE_URL + report.hash, { waitUntil: 'networkidle', timeout: 30000 });
  await page.waitForTimeout(4000);

  const slug = report.label.replace(/\s+/g, '_').toLowerCase();
  await page.screenshot({
    path: path.join(SCREENSHOT_DIR, `${slug}_loaded.png`),
    fullPage: true,
  });

  // Salondata reports expose Print / Download PDF / Download CSV buttons.
  // We want the CSV; convert to xlsx in a follow-up step.
  const exportSelectors = [
    'button:has-text("Download CSV")',
    'a:has-text("Download CSV")',
    'button:has-text("CSV")',
    'a:has-text("CSV")',
    '[aria-label*="CSV" i]',
  ];

  let exportLocator = null;
  for (const sel of exportSelectors) {
    const loc = page.locator(sel).first();
    if (await loc.count()) { exportLocator = loc; break; }
  }
  if (!exportLocator) {
    await page.screenshot({
      path: path.join(SCREENSHOT_DIR, `${slug}_no_export_button.png`),
      fullPage: true,
    });
    throw new Error(
      `[${report.label}] could not find an Excel export button. ` +
      `Inspect screenshots/${slug}_no_export_button.png and update exportSelectors in scrape.js.`
    );
  }

  console.log('  Clicking export button, waiting for download...');
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 60000 }),
    exportLocator.click(),
  ]);

  const outPath = path.join(SCRIPT_DIR, report.output);
  await download.saveAs(outPath);
  const size = fs.statSync(outPath).size;
  console.log(`  Saved ${report.output} (${size.toLocaleString()} bytes).`);
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
    acceptDownloads: true,
  });
  const page = await context.newPage();

  try {
    await login(page);
    for (const report of REPORTS) {
      await downloadReport(page, report);
    }
  } catch (err) {
    console.error('\nScraper error:', err.message);
    await page.screenshot({
      path: path.join(SCREENSHOT_DIR, 'fatal_error.png'),
      fullPage: true,
    }).catch(() => {});
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
