/* Container Explorer — client-side filter, dedup, and Excel export.
   Works as a static GitHub Pages site: fetches data.json + meta.json. */

const REQUIRED_COLUMNS = [
  'VVD', 'LANE', 'CONT NO.', 'FULL/EMPTY', 'TYPE/SIZE', 'POL', 'POD',
  'CONT_Weight', 'AWK', 'DG', 'RF', 'BB', 'SLOT_OPR', 'CONT_OPR',
  'REVENUE MONTH', 'TARGET_PORT'
];

const PAGE_SIZE = 50;
const $ = (id) => document.getElementById(id);

let ALL_DATA = [];
let filtered = [];
let currentPage = 1;

/* ---------- Data loading ---------- */
async function decompressGzip(buf) {
  const ds = new DecompressionStream('gzip');
  const stream = new Response(buf).body.pipeThrough(ds);
  const text = await new Response(stream).text();
  return JSON.parse(text);
}

function applyMeta(meta) {
  if (!meta) {
    $('metaUpdated').textContent = '';
    return;
  }
  if (meta.generatedAt) {
    const d = new Date(meta.generatedAt);
    $('metaUpdated').textContent = `更新于 ${d.toLocaleString('zh-CN')}`;
  }
  if (meta.rowCount) $('metaRows').textContent = `共 ${Number(meta.rowCount).toLocaleString()} 条记录`;
}

async function loadData() {
  try {
    const metaRes = await fetch('meta.json', { cache: 'no-store' }).catch(() => null);
    let meta = null;
    if (metaRes && metaRes.ok) meta = await metaRes.json();

    const compressed = !!(meta && meta.compressed);
    const url = compressed ? 'data.json.gz' : 'data.json';
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`${url} HTTP ${res.status}`);

    ALL_DATA = compressed ? await decompressGzip(await res.arrayBuffer()) : await res.json();
    if (!Array.isArray(ALL_DATA)) throw new Error('data 格式不正确');

    $('metaRows').textContent = `共 ${ALL_DATA.length.toLocaleString()} 条记录`;
    applyMeta(meta);
  } catch (err) {
    console.error(err);
    $('metaRows').textContent = '数据加载失败';
    $('tableBody').innerHTML =
      `<tr class="empty-row"><td colspan="16">无法加载数据：${err.message}<br>请通过本地服务器或 GitHub Pages 访问（file:// 直接打开会被浏览器拦截）。</td></tr>`;
  }
}

/* ---------- Helpers ---------- */
function parseContainerInput(raw) {
  return new Set(
    (raw || '')
      .toUpperCase()
      .split(/[\s,;]+/)
      .map((s) => s.trim())
      .filter(Boolean)
  );
}

// Robust REVENUE MONTH comparison → number (handles 2026-07, 2026/07, Jul-2026, etc.)
function revenueValue(v) {
  if (v == null) return -Infinity;
  if (v instanceof Date) return v.getTime();
  const s = String(v).trim();
  const parsed = Date.parse(s);
  if (!Number.isNaN(parsed) && /\D/.test(s) && /\d{4}/.test(s)) return parsed; // has letters/month name
  // yyyy-mm or yyyymm
  const m = s.match(/(\d{4})[-/]?(\d{1,2})/);
  if (m) return Number(m[1]) * 100 + Number(m[2]);
  return Number(s) || 0;
}

function compareRevenue(a, b) {
  return revenueValue(a) - revenueValue(b);
}

/* ---------- Dedup ---------- */
function dedup(rows) {
  const groups = new Map();
  for (const r of rows) {
    const key = String(r['CONT NO.'] ?? '').trim().toUpperCase();
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  }
  const out = [];
  for (const group of groups.values()) {
    // 1) Prefer rows where TARGET_PORT === POL
    const matched = group.filter(
      (r) => String(r['TARGET_PORT'] ?? '').trim() === String(r['POL'] ?? '').trim()
    );
    const candidates = matched.length ? matched : group;
    // 2) Among candidates, take the largest REVENUE MONTH
    let best = candidates[0];
    for (const r of candidates) {
      if (compareRevenue(r['REVENUE MONTH'], best['REVENUE MONTH']) > 0) best = r;
    }
    out.push(best);
  }
  return out;
}

/* ---------- Filter ---------- */
function applyFilters() {
  const vvd = $('vvd').value.trim().toUpperCase();
  const lane = $('lane').value.trim().toUpperCase();
  const contSet = parseContainerInput($('containers').value);
  const dedupOn = $('dedup').checked;

  let rows = ALL_DATA;
  if (vvd) rows = rows.filter((r) => String(r['VVD'] ?? '').toUpperCase().includes(vvd));
  if (lane) rows = rows.filter((r) => String(r['LANE'] ?? '').toUpperCase().includes(lane));
  if (contSet.size) {
    rows = rows.filter((r) => contSet.has(String(r['CONT NO.'] ?? '').trim().toUpperCase()));
  }
  if (dedupOn) rows = dedup(rows);

  filtered = rows;
  currentPage = 1;
  render();
}

/* ---------- Render ---------- */
const YES_FLAGS = new Set(['Y', 'YES', 'TRUE', '1', 'X']);
function fmtCell(col, val) {
  if (val == null || val === '') return '';
  const s = String(val);
  if (['AWK', 'DG', 'RF', 'BB'].includes(col)) {
    return YES_FLAGS.has(s.trim().toUpperCase())
      ? '<span class="flag-y">Y</span>'
      : '<span class="flag-n">N</span>';
  }
  return s.replace(/[<>&]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]));
}

function render() {
  const total = filtered.length;
  const pages = Math.max(1, Math.ceil(total / PAGE_SIZE));
  currentPage = Math.min(currentPage, pages);
  const start = (currentPage - 1) * PAGE_SIZE;
  const slice = filtered.slice(start, start + PAGE_SIZE);

  $('resultCount').textContent = `共 ${total.toLocaleString()} 条${$('dedup').checked ? '（已去重）' : ''}`;
  $('exportBtn').disabled = total === 0;

  const body = $('tableBody');
  if (total === 0) {
    body.innerHTML = '<tr class="empty-row"><td colspan="16">无匹配结果，请调整筛选条件</td></tr>';
  } else {
    body.innerHTML = slice
      .map(
        (r) =>
          '<tr>' +
          REQUIRED_COLUMNS.map((c) => `<td>${fmtCell(c, r[c])}</td>`).join('') +
          '</tr>'
      )
      .join('');
  }

  const pag = $('pagination');
  if (total > PAGE_SIZE) {
    pag.hidden = false;
    $('pageInfo').textContent = `第 ${currentPage} / ${pages} 页`;
    $('prevPage').disabled = currentPage === 1;
    $('nextPage').disabled = currentPage === pages;
  } else {
    pag.hidden = true;
  }
}

/* ---------- Export Excel (styled) ---------- */
function exportExcel() {
  if (!filtered.length || typeof XLSX === 'undefined') {
    alert('没有可导出的数据，或 Excel 组件未加载。');
    return;
  }
  const aoa = [REQUIRED_COLUMNS, ...filtered.map((r) => REQUIRED_COLUMNS.map((c) => r[c] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // Column widths
  ws['!cols'] = REQUIRED_COLUMNS.map((c) => ({
    wch: Math.max(10, Math.min(22, c.length + 4)),
  }));

  // Header style: bold + brand fill + white text
  const headerFill = { fgColor: { rgb: '4F46E5' } };
  const lastCol = String.fromCharCode(64 + REQUIRED_COLUMNS.length);
  for (let c = 1; c <= REQUIRED_COLUMNS.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c: c - 1 });
    ws[addr].s = {
      font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 11 },
      fill: headerFill,
      alignment: { vertical: 'center', horizontal: 'left' },
      border: { bottom: { style: 'thin', color: { rgb: 'CBD5E1' } } },
    };
  }

  // Freeze header row + autofilter
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  ws['!autofilter'] = { ref: `A1:${lastCol}${aoa.length}` };

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Containers');
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  XLSX.writeFile(wb, `containers_${stamp}.xlsx`);
}

/* ---------- Events ---------- */
function bindEvents() {
  $('filterForm').addEventListener('submit', (e) => {
    e.preventDefault();
    applyFilters();
  });
  $('filterForm').addEventListener('reset', () => {
    setTimeout(() => {
      ALL_DATA.length && renderEmpty();
      filtered = [];
      currentPage = 1;
      render();
    }, 0);
  });
  $('containers').addEventListener('input', () => {
    $('contCount').textContent = `已识别 ${parseContainerInput($('containers').value).size} 个箱号`;
  });
  $('exportBtn').addEventListener('click', exportExcel);
  $('prevPage').addEventListener('click', () => { currentPage--; render(); });
  $('nextPage').addEventListener('click', () => { currentPage++; render(); });
}

function renderEmpty() {
  $('tableBody').innerHTML = '<tr class="empty-row"><td colspan="16">请输入筛选条件后点击「查询」</td></tr>';
  $('resultCount').textContent = '共 0 条';
  $('exportBtn').disabled = true;
  $('pagination').hidden = true;
}

/* ---------- Init ---------- */
bindEvents();
loadData();
