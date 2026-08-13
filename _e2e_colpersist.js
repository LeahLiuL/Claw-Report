const { chromium } = require('playwright');
const path = require('path');

const HTML = path.resolve(__dirname, 'cul_daily_movement.html');
const PWD = 'CUL1234';

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage();
  const jsErrors = [];
  page.on('pageerror', e => jsErrors.push(String(e.message || e)));
  page.on('console', m => { if (m.type() === 'error') jsErrors.push(m.text()); });

  async function ensureLogin() {
    // Idempotent: fill + click Sign In, then wait for overlay hidden + data ready.
    await page.fill('#loginPwd', PWD);
    await page.click('button:has-text("Sign In")');
    await page.waitForFunction(() =>
      document.getElementById('loginOverlay') &&
      document.getElementById('loginOverlay').classList.contains('hidden') &&
      window.visibleCols && typeof TODAY_DATA !== "undefined", null, { timeout: 8000 });
  }

  // 1) Initial load + login
  await page.goto('file://' + HTML);
  await ensureLogin();

  // 2) Hide Remark + Code on Full Schedule (view '2'), then persist
  const saved = await page.evaluate(() => {
    window.visibleCols['2'].delete('remark');
    window.visibleCols['2'].delete('code');
    window.saveColVisibility();
    return localStorage.getItem('cul_movement_cols');
  });
  console.log('localStorage after save:', saved);

  // 3) Reload (simulate "refresh next time")
  await page.reload();
  await page.waitForFunction(() => window.visibleCols && typeof TODAY_DATA !== "undefined", null, { timeout: 8000 });

  // 4) Check persisted state survived reload
  const state = await page.evaluate(() => {
    const v = window.visibleCols['2'];
    return { hasRemark: v.has('remark'), hasCode: v.has('code'), hasVessel: v.has('vessel') };
  });
  console.log('visibleCols[2] after reload:', JSON.stringify(state));

  // 5) Render Full Schedule and read actual header columns
  await page.click('.tab-btn[data-tab="fullScheduleView"]').catch(() => {});
  await page.waitForTimeout(300);
  const headerCols = await page.evaluate(() => {
    const ths = document.querySelectorAll('#fullTable thead th');
    return Array.from(ths).map(t => t.textContent.trim());
  });
  console.log('Full Schedule header cols after reload:', JSON.stringify(headerCols));

  const headerHasRemark = headerCols.includes('Remark');
  const headerHasCode = headerCols.includes('Code');

  console.log('JS errors:', jsErrors.length, jsErrors.slice(0, 5));

  const ok = state.hasRemark === false
    && state.hasCode === false
    && state.hasVessel === true
    && headerHasRemark === false
    && headerHasCode === false
    && jsErrors.length === 0;
  await browser.close();
  console.log('OVERALL:', ok ? 'PASS' : 'FAIL');
  process.exit(ok ? 0 : 1);
})();
