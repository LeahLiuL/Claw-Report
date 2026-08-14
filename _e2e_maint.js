const { chromium } = require('C:/Users/leahliu/.workbuddy/binaries/node/workspace/node_modules/playwright');
const path = require('path');
(async () => {
  const browser = await chromium.launch();
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();
  const url = 'file:///' + path.resolve('cul_daily_movement.html').replace(/\\/g, '/');
  const jsErrors = [];
  page.on('pageerror', e => jsErrors.push('PAGEERR: ' + e.message));
  page.on('console', m => { if (m.type() === 'error') jsErrors.push('CONSOLE: ' + m.text()); });

  await page.goto(url, { waitUntil: 'networkidle' });
  await page.waitForTimeout(1000);
  const dbg0 = await page.evaluate(() => ({ overlay: document.getElementById('loginOverlay')?.className, hasMaint: typeof MAINT_DATA !== 'undefined' }));
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
    serviceTableGone: document.getElementById('maintServiceTable') === null,
    portHead: Array.from(document.querySelectorAll('#maintPortThead th')).map(t=>t.textContent.trim()),
    portRows: document.querySelectorAll('#maintPortTbody > tr:not(.detail-wrap)').length,
    monthRows: document.querySelectorAll('#maintMonthTbody tr').length,
    bars: document.querySelectorAll('#maintMonthBars > div').length,
    firstRow: (() => { const tr = document.querySelector('#maintPortTbody > tr:not(.detail-wrap)'); return tr ? Array.from(tr.querySelectorAll('td')).map(td=>td.textContent.trim()) : []; })(),
  }));
  console.log('CUL-default:', JSON.stringify(cul, null, 2));

  // Expand first port row -> detail should appear with unmaintained records
  const expand = await page.evaluate(() => {
    const tr = document.querySelector('#maintPortTbody > tr:not(.detail-wrap)');
    if(!tr) return { clicked:false };
    tr.click();
    const detail = tr.nextElementSibling;
    const open = detail && detail.classList.contains('detail-wrap') && detail.style.display !== 'none';
    const detailRows = detail ? detail.querySelector('tbody').querySelectorAll('tr').length : 0;
    return { clicked:true, open, detailRows,
      redCells: detail ? detail.querySelectorAll('td[style*="C0392B"]').length : 0 };
  });
  console.log('EXPAND:', JSON.stringify(expand));

  // switch operator to ALL
  await page.evaluate(() => {
    const s = document.getElementById('maintOpFilter');
    s.value = 'ALL'; s.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await page.waitForTimeout(400);
  const all = await page.evaluate(() => ({
    op: document.getElementById('maintOpFilter').value,
    total: document.getElementById('maintChipTotal').textContent.trim(),
    plog: document.getElementById('maintChipPortLog').textContent.trim(),
    vs: document.getElementById('maintChipVSched').textContent.trim(),
  }));
  console.log('ALL:', JSON.stringify(all));

  console.log('JS errors:', jsErrors.length, jsErrors.slice(0,5));
  const ok = cul.op==='CUL' && cul.total==='1,636' && cul.plog.indexOf('72.6%')>=0 && cul.vs.indexOf('86.5%')>=0
    && cul.serviceTableGone === true
    && cul.portHead.join('|')==='|Port|Calls|Port Log Maint|Port Log Rate|Vessel Sched Maint|Vessel Sched Rate|Unmaintained'
    && cul.portRows>0 && cul.monthRows>0 && cul.bars>0
    && expand.clicked && expand.open && expand.detailRows>0 && expand.redCells>0
    && all.op==='ALL' && all.total==='9,283' && all.plog.indexOf('15.8%')>=0 && all.vs.indexOf('54.1%')>=0
    && jsErrors.length===0;
  await browser.close();
  console.log('OVERALL:', ok ? 'PASS' : 'FAIL');
  process.exit(ok ? 0 : 1);
})();
