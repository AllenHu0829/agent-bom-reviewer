/* ──────────────────────────────────────────────────────────
   BOM 审核专家 – app.js
   ────────────────────────────────────────────────────────── */

const COLUMN_ALIASES = {
  partNumber:  ['料号','物料编码','part number','partnumber','物料号','p/n','pn','编码','内部料号','item number'],
  description: ['物料描述','描述','description','desc','品名','名称','物料名称'],
  pkg:         ['封装','package','footprint','pkg','封装规格'],
  value:       ['value','参数值','参数','规格','值','参数/规格'],
  qty:         ['用量','数量','quantity','qty','需求数量'],
  ref:         ['位号','reference','ref','reference designator','designator','位置','ref des','位号标识'],
  mfr:         ['厂商','manufacturer','制造商','厂家','品牌','vendor','供应商'],
  mpn:         ['厂商料号','mpn','manufacturer part number','厂家型号','型号','制造商料号','mfr part']
};

let bomRows = [];
let bomColumns = [];
let columnMap = {};
let banList = [];
let auditResults = [];
let highlightedRow = -1;

/* ═══ DOM refs ═══ */
const $ = id => document.getElementById(id);
const bomUpload     = $('bomUpload');
const bomFileInput  = $('bomFileInput');
const banUpload     = $('banUpload');
const banFileInput  = $('banFileInput');
const fileInfo      = $('fileInfo');
const tableSection  = $('tableSection');
const bomHead       = $('bomHead');
const bomBody       = $('bomBody');
const tableTitle    = $('tableTitle');
const tableCount    = $('tableCount');
const summaryBar    = $('summaryBar');
const resultsPanel  = $('resultsPanel');
const resultsPlaceholder = $('resultsPlaceholder');
const resultsGroups = $('resultsGroups');
const exportReportBtn = $('exportReportBtn');

/* ═══ File Upload ═══ */
function setupUploadBox(box, input, handler) {
  box.addEventListener('click', () => input.click());
  input.addEventListener('change', e => { if (e.target.files[0]) handler(e.target.files[0]); });
  box.addEventListener('dragover', e => { e.preventDefault(); box.classList.add('drag-over'); });
  box.addEventListener('dragleave', () => box.classList.remove('drag-over'));
  box.addEventListener('drop', e => {
    e.preventDefault();
    box.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handler(e.dataTransfer.files[0]);
  });
}

setupUploadBox(bomUpload, bomFileInput, f => parseFile(f, 'bom'));
setupUploadBox(banUpload, banFileInput, f => parseFile(f, 'ban'));

function parseFile(file, type) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'csv') {
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: r => onFileParsed(r.data, r.meta.fields, type, file.name)
    });
  } else {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      const fields = json.length ? Object.keys(json[0]) : [];
      onFileParsed(json, fields, type, file.name);
    };
    reader.readAsArrayBuffer(file);
  }
}

function onFileParsed(data, fields, type, fileName) {
  if (type === 'bom') {
    bomColumns = fields;
    bomRows = data.map((r, i) => ({ _idx: i + 1, ...r }));
    columnMap = mapColumns(fields);
    renderFileTag('bom', fileName);
    renderTable();
    runAudit();
  } else {
    banList = extractBanList(data, fields);
    renderFileTag('ban', fileName);
    if (bomRows.length) runAudit();
  }
}

function extractBanList(data, fields) {
  const list = new Set();
  const lowerFields = fields.map(f => f.toLowerCase().trim());
  data.forEach(row => {
    fields.forEach((f, i) => {
      const v = String(row[f] || '').trim();
      if (v) list.add(v.toUpperCase());
    });
  });
  return list;
}

function renderFileTag(type, name) {
  let tag = fileInfo.querySelector(`[data-type="${type}"]`);
  if (!tag) {
    tag = document.createElement('span');
    tag.className = 'file-tag';
    tag.dataset.type = type;
    fileInfo.appendChild(tag);
  }
  const icon = type === 'bom' ? '📄' : '🚫';
  tag.innerHTML = `${icon} ${name} <span class="remove-file" data-rm="${type}">✕</span>`;
  tag.querySelector('.remove-file').addEventListener('click', e => {
    e.stopPropagation();
    if (type === 'bom') { bomRows = []; bomColumns = []; columnMap = {}; clearTable(); clearResults(); }
    else { banList = []; }
    tag.remove();
    if (bomRows.length) runAudit();
  });
}

/* ═══ Column Mapping ═══ */
function mapColumns(fields) {
  const map = {};
  for (const [key, aliases] of Object.entries(COLUMN_ALIASES)) {
    for (const f of fields) {
      const fl = f.toLowerCase().trim();
      if (aliases.includes(fl)) { map[key] = f; break; }
    }
  }
  return map;
}

function col(row, key) {
  return columnMap[key] ? String(row[columnMap[key]] ?? '').trim() : '';
}

/* ═══ Table Render ═══ */
function renderTable() {
  bomHead.innerHTML = '';
  bomBody.innerHTML = '';
  if (!bomColumns.length) return;

  const thRow = document.createElement('tr');
  const thIdx = document.createElement('th');
  thIdx.textContent = '#';
  thRow.appendChild(thIdx);
  bomColumns.forEach(c => {
    const th = document.createElement('th');
    th.textContent = c;
    thRow.appendChild(th);
  });
  bomHead.appendChild(thRow);

  bomRows.forEach((row, i) => {
    const tr = document.createElement('tr');
    tr.dataset.rowIdx = i;
    const tdIdx = document.createElement('td');
    tdIdx.textContent = row._idx;
    tr.appendChild(tdIdx);
    bomColumns.forEach(c => {
      const td = document.createElement('td');
      td.textContent = row[c] ?? '';
      tr.appendChild(td);
    });
    bomBody.appendChild(tr);
  });

  tableCount.textContent = `(${bomRows.length} 行)`;
}

function clearTable() {
  bomHead.innerHTML = '';
  bomBody.innerHTML = '';
  tableCount.textContent = '';
}

function highlightTableRow(rowIdx) {
  bomBody.querySelectorAll('tr').forEach(tr => tr.classList.remove('row-highlight'));
  if (rowIdx < 0) return;
  const tr = bomBody.querySelector(`tr[data-row-idx="${rowIdx}"]`);
  if (tr) {
    tr.classList.add('row-highlight');
    tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }
}

/* ═══ Audit Engine ═══ */
function runAudit() {
  auditResults = [];
  if (!bomRows.length) { clearResults(); return; }

  checkBannedParts();
  checkQuantity();
  checkDuplicatePN();
  checkDuplicateMPN();
  checkMultiPNSameSpec();

  applyRowStyles();
  renderResults();
  updateSummary();
}

function addResult(type, severity, desc, rows) {
  auditResults.push({ type, severity, desc, rows });
}

/* ── Rule 1: Banned Parts ── */
function checkBannedParts() {
  if (!banList.size) return;
  bomRows.forEach((row, i) => {
    const pn = col(row, 'partNumber').toUpperCase();
    const mpn = col(row, 'mpn').toUpperCase();
    if (pn && banList.has(pn)) {
      addResult('banned', 'error', `禁用物料：料号 "${col(row,'partNumber')}"`, [i]);
    }
    if (mpn && banList.has(mpn)) {
      addResult('banned', 'error', `禁用物料：厂商料号 "${col(row,'mpn')}"`, [i]);
    }
  });
}

/* ── Rule 2: Quantity Check ── */
function parseRefs(refStr) {
  if (!refStr) return [];
  return refStr.split(/[,\s/;、]+/).map(s => s.trim()).filter(Boolean);
}

function checkQuantity() {
  if (!columnMap.qty || !columnMap.ref) return;
  bomRows.forEach((row, i) => {
    const qtyVal = parseInt(col(row, 'qty'), 10);
    const refs = parseRefs(col(row, 'ref'));
    if (isNaN(qtyVal) || refs.length === 0) return;
    if (refs.length !== qtyVal) {
      addResult('qty', 'error',
        `数量不匹配：用量=${qtyVal}，位号数=${refs.length}（${col(row,'ref')}）`,
        [i]);
    }
  });
}

/* ── Rule 3: Duplicate Part Number ── */
function checkDuplicatePN() {
  if (!columnMap.partNumber) return;
  const map = {};
  bomRows.forEach((row, i) => {
    const pn = col(row, 'partNumber');
    if (!pn) return;
    (map[pn] = map[pn] || []).push(i);
  });
  for (const [pn, idxs] of Object.entries(map)) {
    if (idxs.length > 1) {
      const rowNums = idxs.map(i => bomRows[i]._idx).join(', ');
      addResult('dup-pn', 'warning', `重复料号 "${pn}"（行 ${rowNums}）`, idxs);
    }
  }
}

/* ── Rule 4: Duplicate MPN ── */
function checkDuplicateMPN() {
  if (!columnMap.mpn) return;
  const map = {};
  bomRows.forEach((row, i) => {
    const mpn = col(row, 'mpn');
    if (!mpn) return;
    (map[mpn] = map[mpn] || []).push(i);
  });
  for (const [mpn, idxs] of Object.entries(map)) {
    if (idxs.length > 1) {
      const rowNums = idxs.map(i => bomRows[i]._idx).join(', ');
      addResult('dup-mpn', 'warning', `重复厂商料号 "${mpn}"（行 ${rowNums}）`, idxs);
    }
  }
}

/* ── Rule 5: Same Value+Pkg, Multiple PNs ── */
function checkMultiPNSameSpec() {
  if (!columnMap.value || !columnMap.pkg || !columnMap.partNumber) return;
  const groups = {};
  bomRows.forEach((row, i) => {
    const v = col(row, 'value').toUpperCase();
    const p = col(row, 'pkg').toUpperCase();
    if (!v || !p) return;
    const key = `${v}||${p}`;
    if (!groups[key]) groups[key] = {};
    const pn = col(row, 'partNumber');
    if (!pn) return;
    (groups[key][pn] = groups[key][pn] || []).push(i);
  });
  for (const [spec, pnMap] of Object.entries(groups)) {
    const pns = Object.keys(pnMap);
    if (pns.length > 1) {
      const allIdxs = Object.values(pnMap).flat();
      const [v, p] = spec.split('||');
      const pnList = pns.join(', ');
      addResult('multi-pn', 'info',
        `同规格多料号：Value=${v} 封装=${p}，料号：${pnList}`,
        allIdxs);
    }
  }
}

/* ═══ Row Styles ═══ */
function applyRowStyles() {
  const errorRows = new Set();
  const warnRows = new Set();
  auditResults.forEach(r => {
    r.rows.forEach(i => {
      if (r.severity === 'error') errorRows.add(i);
      else if (r.severity === 'warning') warnRows.add(i);
    });
  });
  bomBody.querySelectorAll('tr').forEach(tr => {
    const idx = parseInt(tr.dataset.rowIdx, 10);
    tr.classList.remove('row-error', 'row-warning');
    if (errorRows.has(idx)) tr.classList.add('row-error');
    else if (warnRows.has(idx)) tr.classList.add('row-warning');
  });
}

/* ═══ Results Render ═══ */
function clearResults() {
  resultsGroups.innerHTML = '';
  resultsPlaceholder.style.display = 'flex';
  updateSummary(true);
}

function renderResults() {
  resultsGroups.innerHTML = '';
  if (!auditResults.length) {
    resultsPlaceholder.style.display = 'flex';
    resultsPlaceholder.querySelector('.placeholder-text').textContent = '✅ BOM 审核通过，未发现问题';
    resultsPlaceholder.querySelector('.placeholder-icon').textContent = '🎉';
    return;
  }
  resultsPlaceholder.style.display = 'none';

  const groups = { error: [], warning: [], info: [] };
  auditResults.forEach(r => groups[r.severity].push(r));

  const labels = { error: '错误', warning: '警告', info: '提示' };
  for (const sev of ['error', 'warning', 'info']) {
    if (!groups[sev].length) continue;
    const div = document.createElement('div');
    div.className = `result-group group-${sev}`;
    div.innerHTML = `
      <div class="result-group-header">
        <span>${labels[sev]}</span>
        <span class="badge">${groups[sev].length}</span>
      </div>
      <div class="result-list"></div>`;
    const list = div.querySelector('.result-list');
    groups[sev].forEach(r => {
      const item = document.createElement('div');
      item.className = 'result-item';
      const typeLabel = {
        banned: '禁用物料', qty: '数量不匹配',
        'dup-pn': '重复料号', 'dup-mpn': '重复厂商料号', 'multi-pn': '同规格多料号'
      }[r.type] || r.type;
      item.innerHTML = `
        <span class="ri-type ri-type-${r.type}">${typeLabel}</span>
        <span class="ri-desc">${r.desc}</span>
        <span class="ri-rows">行 ${r.rows.map(i => bomRows[i]._idx).join(',')}</span>`;
      item.addEventListener('click', () => highlightTableRow(r.rows[0]));
      list.appendChild(item);
    });

    const header = div.querySelector('.result-group-header');
    const listEl = div.querySelector('.result-list');
    header.addEventListener('click', () => {
      listEl.style.display = listEl.style.display === 'none' ? 'block' : 'none';
    });

    resultsGroups.appendChild(div);
  }
}

function updateSummary(empty) {
  $('sumTotal').textContent = empty ? 0 : bomRows.length;
  $('sumError').textContent = empty ? 0 : auditResults.filter(r => r.severity === 'error').length;
  $('sumWarn').textContent = empty ? 0 : auditResults.filter(r => r.severity === 'warning').length;
  $('sumInfo').textContent = empty ? 0 : auditResults.filter(r => r.severity === 'info').length;
}

/* ═══ Export Report ═══ */
exportReportBtn.addEventListener('click', () => {
  if (!auditResults.length && !bomRows.length) return;

  const reportRows = auditResults.map(r => ({
    '严重程度': { error: '错误', warning: '警告', info: '提示' }[r.severity],
    '类型': { banned:'禁用物料', qty:'数量不匹配', 'dup-pn':'重复料号', 'dup-mpn':'重复厂商料号', 'multi-pn':'同规格多料号' }[r.type],
    '描述': r.desc,
    '涉及行号': r.rows.map(i => bomRows[i]._idx).join(', ')
  }));

  if (!reportRows.length) {
    reportRows.push({ '严重程度': '-', '类型': '-', '描述': '审核通过，未发现问题', '涉及行号': '-' });
  }

  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.json_to_sheet(reportRows);
  XLSX.utils.book_append_sheet(wb, ws1, '审核结果');

  if (bomRows.length) {
    const bomData = bomRows.map(r => {
      const o = {};
      bomColumns.forEach(c => o[c] = r[c]);
      return o;
    });
    const ws2 = XLSX.utils.json_to_sheet(bomData);
    XLSX.utils.book_append_sheet(wb, ws2, 'BOM数据');
  }

  XLSX.writeFile(wb, 'BOM审核报告.xlsx');
});
