const { chromium } = require('C:/Users/leahliu/.workbuddy/binaries/node/workspace/node_modules/playwright');
(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();
  const url = 'file:///' + require('path').resolve('cul_daily_movement.html').replace(/\\/g, '/');
  const jsErrors = [];
  page.on('pageerror', e => jsErrors.push('PAGEERR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') jsErrors.push('CONSOLE: ' + m.text()); });

  await page.goto(url, { waitUntil: 'networkidle' });
  await page.waitForTimeout(1000);
  const dbg0 = await page.evaluate(() => ({ overlay: document.getElementById('loginOverlay') ? document.getElementById('loginOverlay').className : null, hasMaint: typeof MAINT_DATA !== 'undefined' }));
  if (!dbg0.overlay || dbg0.overlay.indexOf('hidden') < 0) {
    await page.fill('#loginPwd', 'CUL1234');
    await page.click('button:has-text("Sign In")');
    await page.waitForTimeout(1000);
  }
  await page.click('.tab-btn[data-tab="maintView"]').catch(e => console.log('tab click err', e.message));
  await page.waitForTimeout(600);

  const cul = await page.evaluate(() => ({
    op: document.getElementById('maintOpFilter').value,
    total: document.getElementById('maintChipTotal').textContent.trim(),
    plog: document.getElementById('maintChipPortLog').textContent.trim(),
    vs: document.getElementById('maintChipVSched').textContent.trim(),
    svcRows: document.querySelectorAll('#maintServiceTbody tr').length,
    portRows: document.querySelectorAll('#maintPortTbody tr').length,
    monthRows: document.querySelectorAll('#maintMonthTbody tr').length,
    bars: document.querySelectorAll('#maintMonthBars > div').length,
    svcHead: Array.from(document.querySelectorAll('#maintServiceThead th')).map(t=>t.textContent.trim()),
    firstSvc: (document.querySelector('#maintServiceTbody tr')?Array.from(document.querySelector('#maintServiceTbody tr').querySelectorAll('td')).map(td=>td.textContent.trim()):[]),
    firstPort: (document.querySelector('#maintPortTbody tr')?Array.from(document.querySelector('#maintPortTbody tr').querySelectorAll('td')).map(td=>td.textContent.trim()):[]),
  }));
  console.log('CUL-default:', JSON.stringify(cul, null, 2));

  await page.evaluate(() => {
    const s = document.getElementById('maintOpFilter');
    s.value = 'ALL';
    s.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await page.waitForTimeout(400);
  const all = await page.evaluate(() => ({
    op: document.getElementById('maintOpFilter').value,
    total: document.getElementById('maintChipTotal').textContent.trim(),
    plog: document.getElementById('maintChipPortLog').textContent.trim(),
    vs: document.getElementById('maintChipVSched').textContent.trim(),
  }));
  console.log('ALL:', JSON.stringify(all));

  await page.evaluate(() => {
    const f = document.getElementById('maintFrom'); const t = document.getElementById('maintTo');
    f.value = '2026-03-01'; t.value = '2026-03-31';
    f.dispatchEvent(new Event('change', { bubbles: true }));
    t.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await page.waitForTimeout(300);
  const mar = await page.evaluate(() => ({ total: document.getElementById('maintChipTotal').textContent.trim(), monthRows: document.querySelectorAll('#maintMonthTbody tr').length }));
  console.log('MAR-2026:', JSON.stringify(mar));

  console.log('JS errors:', jsErrors.length, jsErrors.slice(0,5));
  const ok = cul.op==='CUL' && cul.total==='1,636' && cul.plog.indexOf('72.6%')>=0 && cul.vs.indexOf('86.5%')>=0
    && cul.svcRows>0 && cul.portRows>0 && cul.monthRows>0 && cul.bars>0
    && cul.svcHead.join('|')==='Service|Calls|Port Log Maint|Port Log Rate|Vessel Sched Maint|Vessel Sched Rate'
    && all.op==='ALL' && all.total==='9,283' && all.plog.indexOf('15.8%')>=0 && all.vs.indexOf('54.1%')>=0
    && mar.total!=='9,283' && mar.monthRows>=1
    && jsErrors.length===0;
  await browser.close();
  console.log('OVERALL:', ok ? 'PASS' : 'FAIL');
  process.exit(ok ? 0 : 1);
})();
