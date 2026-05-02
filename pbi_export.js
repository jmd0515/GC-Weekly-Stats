// pbi_export.js — Local Playwright Power BI export (PARKED — work in progress)
//
// Status: scaffolded but not yet functional. The persistent Playwright profile
// approach didn't survive a restart in initial testing (Microsoft Entra likely
// requires interactive auth too often). For now, All_Salons.xlsx is updated
// manually and committed to the repo as the source of truth. Re-enable this
// once a viable Power BI auth strategy is in place.
//
// Runs on the user's PC (Windows Task Scheduler) using a dedicated Playwright
// profile (one-time interactive login, then headless).
//
// Flow:
//   1. Open the Customer Experience report in Power BI
//   2. Set Start/End date filters to the most recent Friday
//   3. Hover the table visual, click "⋯" → "Export data"
//   4. Save the downloaded xlsx as All_Salons.xlsx in the repo folder
//
// Power BI is canvas-rendered and report-specific, so this script is a STARTING
// POINT. Run with HEADED=true for the first runs, watch the screenshots in
// screenshots/, and tune the selectors / wait points to match the real page.

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const REPORT_URL = 'https://app.powerbi.com/Redirect?action=OpenReport&appId=b1897d2f-ff01-4fd3-be45-85d7b2eea180&reportObjectId=3865e403-e2f6-4f9f-91dd-27f64d0b80ae&ctid=99c1c3f9-432d-45d9-aca1-6becd3bff20d&reportPage=ReportSectione1813b8d90c461002000&pbi_source=appShareLink';

const SCRIPT_DIR = __dirname;
const SCREENSHOT_DIR = path.join(SCRIPT_DIR, 'screenshots');
if (!fs.existsSync(SCREENSHOT_DIR)) fs.mkdirSync(SCREENSHOT_DIR);

// Dedicated Playwright profile. First run: launch headed, sign in to Microsoft
// once, close the script (Ctrl+C). Auth cookies persist in this folder; future
// runs reuse them headlessly until Entra refresh token expires (~30–90 days).
//
// Sharing the user's main Edge profile doesn't work for enterprise Power BI
// because the Microsoft auth lives in Windows SSO (WAM), not Edge cookies.
const PROFILE_DIR = process.env.PBI_PROFILE
  || path.join(SCRIPT_DIR, '.pbi-profile');

const HEADLESS = process.env.HEADLESS === 'true';

function fmtDate(d) {
  // Power BI US locale typically expects M/D/YYYY in date inputs
  return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

function mostRecentFriday() {
  const d = new Date();
  const back = (d.getDay() - 5 + 7) % 7;
  d.setDate(d.getDate() - back);
  return d;
}

async function snap(page, name) {
  const p = path.join(SCREENSHOT_DIR, `pbi_${name}.png`);
  await page.screenshot({ path: p, fullPage: true }).catch(() => {});
  console.log(`  📸 ${path.basename(p)}`);
}

async function main() {
  const friday = mostRecentFriday();
  const fridayStr = fmtDate(friday);
  console.log(`Power BI export — week ending ${fridayStr}`);
  console.log(`Profile dir: ${PROFILE_DIR}`);

  const firstRun = !fs.existsSync(PROFILE_DIR);
  if (firstRun) {
    console.log('First run detected — launching headed browser for one-time login.');
    console.log('Sign in to Microsoft / Power BI, wait for the report to load, then press Ctrl+C.');
  }

  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: HEADLESS && !firstRun,
    viewport: { width: 1600, height: 1000 },
    acceptDownloads: true,
  });
  const page = context.pages()[0] || (await context.newPage());

  try {
    console.log('Opening Power BI report...');
    await page.goto(REPORT_URL, { waitUntil: 'load', timeout: 60000 });

    if (firstRun) {
      console.log('\n>>> Sign in now in the browser window. Once the report fully loads, press Ctrl+C in this terminal.');
      console.log('>>> The script will exit; auth cookies are saved. Re-run normally on the next invocation.');
      await new Promise(() => {}); // hang here so user can authenticate
    }

    // Power BI loads slowly; wait for the canvas to settle.
    await page.waitForTimeout(15000);
    await snap(page, '1_loaded');

    // ── Date filters ────────────────────────────────────────────────────────
    // Power BI date slicers expose two text inputs (Start / End) that accept
    // typed dates. Selector varies per report; the generic selector below
    // targets visible date inputs, which matches most "between" slicers.
    console.log(`Setting start/end to ${fridayStr}...`);
    const dateInputs = page.locator('input[aria-label*="Start" i], input[aria-label*="End" i], input[type="text"][role="textbox"]');
    const inputCount = await dateInputs.count();
    console.log(`  Found ${inputCount} candidate date input(s).`);

    // Best-effort: fill the first two visible date inputs we can find.
    // If this misses the right inputs, watch screenshots to find better selectors.
    let filled = 0;
    for (let i = 0; i < inputCount && filled < 2; i++) {
      const inp = dateInputs.nth(i);
      if (!(await inp.isVisible().catch(() => false))) continue;
      await inp.click({ clickCount: 3 }).catch(() => {}); // select existing text
      await inp.fill(fridayStr).catch(() => {});
      await inp.press('Enter').catch(() => {});
      filled++;
    }
    await page.waitForTimeout(5000);
    await snap(page, '2_dates_set');

    // ── Hover the table, click "⋯", click "Export data" ─────────────────────
    // Power BI tables use role="grid" on the rendered visual; hovering exposes
    // the More-options ("⋯") button with aria-label "More options".
    console.log('Looking for table visual...');
    const visual = page.locator('visual-container').first();
    await visual.hover({ timeout: 10000 });
    await page.waitForTimeout(1000);
    await snap(page, '3_visual_hover');

    const moreOptions = visual.locator('button[aria-label*="More options" i], button[title*="More options" i]').first();
    await moreOptions.click({ timeout: 10000 });
    await page.waitForTimeout(1500);
    await snap(page, '4_more_menu_open');

    const exportItem = page.locator('[role="menuitem"]:has-text("Export data"), button:has-text("Export data")').first();
    await exportItem.click({ timeout: 10000 });
    await page.waitForTimeout(2500);
    await snap(page, '5_export_dialog');

    // ── Export dialog: choose Summarized data + Excel, then click Export ────
    // The dialog typically has radio buttons for data type. Default is fine
    // for "current visual" exports; we just confirm by clicking the Export button.
    const exportBtn = page.locator('button:has-text("Export"):not(:has-text("Export data"))').last();
    const downloadPromise = page.waitForEvent('download', { timeout: 120000 });
    await exportBtn.click({ timeout: 10000 });
    await snap(page, '6_export_clicked');

    const download = await downloadPromise;
    const dst = path.join(SCRIPT_DIR, 'All_Salons.xlsx');
    await download.saveAs(dst);
    const size = fs.statSync(dst).size;
    console.log(`✅ Saved All_Salons.xlsx (${size.toLocaleString()} bytes)`);
  } catch (err) {
    console.error('\n❌ Error:', err.message);
    await snap(page, 'fatal').catch(() => {});
    process.exitCode = 1;
  } finally {
    await context.close();
  }
}

main();
