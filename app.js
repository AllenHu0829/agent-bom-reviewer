/* ──────────────────────────────────────────────────────────
   BOM 审核专家 – app.js v3
   流程：上传 → 自动列映射（显示状态摘要，可展开修改） → 点击分析
   ────────────────────────────────────────────────────────── */

const BOM_FIELDS = [
  { key: 'partNumber',  label: '料号',       required: true,  aliases: ['料号','物料编码','part number','partnumber','物料号','p/n','pn','编码','内部料号','item number','item no','item no.'] },
  { key: 'mpn',         label: '厂商料号',   required: true,  aliases: ['厂商料号','mpn','manufacturer part number','厂家型号','型号','制造商料号','mfr part','mfr p/n','vendor p/n'] },
  { key: 'qty',         label: '基本用量',   required: true,  aliases: ['基本用量','用量','数量','quantity','qty','需求数量','usage','基本用量(pcs)','基本用量（pcs）'] },
  { key: 'ref',         label: '位号',       required: true,  aliases: ['位号','reference','ref','reference designator','designator','位置','ref des','位号标识','ref.des','ref des.'] },
  { key: 'description', label: '物料描述',   required: false, aliases: ['物料描述','描述','description','desc','品名','名称','物料名称','规格描述'] },
  { key: 'unit',        label: '计量单位',   required: false, aliases: ['计量单位','单位','unit','uom','计量'] },
  { key: 'lossRate',    label: '子件损耗率', required: false, aliases: ['子件损耗率','损耗率','loss rate','损耗','scrap rate','loss','损耗率(%)','损耗率（%）'] },
];

const FIELD_OPTIONS = BOM_FIELDS.map(f => ({ key: f.key, label: f.label + (f.required ? ' *' : ''), required: f.required }));

let bomRows = [];
let bomColumns = [];
let columnMap = {};
let banList = new Set();
let auditResults = [];

/* ═══ DOM refs ═══ */
const $ = id => document.getElementById(id);
const bomUpload        = $('bomUpload');
const bomFileInput     = $('bomFileInput');
const banUpload        = $('banUpload');
const banFileInput     = $('banFileInput');
const fileInfo         = $('fileInfo');
const mappingSection   = $('mappingSection');
const mappingStatusBar = $('mappingStatusBar');
const mappingStatusText = $('mappingStatusText');
const mappingToggle    = $('mappingToggle');
const mappingDetail    = $('mappingDetail');
const mappingList      = $('mappingList');
const btnAnalyze       = $('btnAnalyze');
const bomHead          = $('bomHead');
const bomBody          = $('bomBody');
const tableCount       = $('tableCount');
const resultsPlaceholder = $('resultsPlaceholder');
const resultsGroups    = $('resultsGroups');
const exportReportBtn  = $('exportReportBtn');

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
    columnMap = autoMapColumns(fields);
    renderFileTag('bom', fileName);
    renderTable();
    showMappingUI();
    clearResults();
  } else {
    banList = extractBanList(data, fields);
    renderFileTag('ban', fileName);
  }
}

function extractBanList(data, fields) {
  const list = new Set();
  data.forEach(row => {
    fields.forEach(f => {
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
    if (type === 'bom') {
      bomRows = []; bomColumns = []; columnMap = {};
      clearTable(); clearResults(); hideMappingUI();
    } else {
      banList = new Set();
    }
    tag.remove();
  });
}

/* ═══ Auto Column Mapping ═══ */
function autoMapColumns(fields) {
  const map = {};
  const used = new Set();

  for (const field of BOM_FIELDS) {
    for (const f of fields) {
      if (used.has(f)) continue;
      const fl = f.toLowerCase().trim();
      if (field.aliases.includes(fl)) {
        map[field.key] = f;
        used.add(f);
        break;
      }
    }
  }

  for (const field of BOM_FIELDS) {
    if (map[field.key]) continue;
    for (const f of fields) {
      if (used.has(f)) continue;
      const fl = f.toLowerCase().trim();
      const matched = field.aliases.some(a => fl.includes(a) || a.includes(fl));
      if (matched) {
        map[field.key] = f;
        used.add(f);
        break;
      }
    }
  }

  return map;
}

function col(row, key) {
  return columnMap[key] ? String(row[columnMap[key]] ?? '').trim() : '';
}

/* ═══ Mapping UI ═══ */
function showMappingUI() {
  mappingSection.style.display = 'block';
  mappingDetail.style.display = 'none';
  renderMappingList();
  updateMappingStatus();
}

function hideMappingUI() {
  mappingSection.style.display = 'none';
  mappingList.innerHTML = '';
}

mappingToggle.addEventListener('click', () => {
  const visible = mappingDetail.style.display !== 'none';
  mappingDetail.style.display = visible ? 'none' : 'block';
  mappingToggle.textContent = visible ? '修改映射 ▾' : '收起 ▴';
});

function updateMappingStatus() {
  const requiredFields = BOM_FIELDS.filter(f => f.required);
  const mappedRequired = requiredFields.filter(f => columnMap[f.key]);
  const totalMapped = BOM_FIELDS.filter(f => columnMap[f.key]).length;
  const allRequiredOk = mappedRequired.length === requiredFields.length;

  if (allRequiredOk) {
    mappingStatusText.textContent = `已自动识别 ${totalMapped}/${BOM_FIELDS.length} 列（${requiredFields.length} 项必填全部匹配）`;
    mappingStatusText.className = 'mapping-status-text status-ok';
  } else {
    const missing = requiredFields.filter(f => !columnMap[f.key]).map(f => f.label);
    mappingStatusText.textContent = `已识别 ${totalMapped}/${BOM_FIELDS.length} 列，缺少必填项：${missing.join('、')}`;
    mappingStatusText.className = 'mapping-status-text status-warn';
    mappingDetail.style.display = 'block';
    mappingToggle.textContent = '收起 ▴';
  }
}

function renderMappingList() {
  mappingList.innerHTML = '';

  bomColumns.forEach(colName => {
    const row = document.createElement('div');
    row.className = 'mapping-row';

    const sampleVal = getSampleValue(colName);
    const nameEl = document.createElement('div');
    nameEl.className = 'mapping-col-name';
    nameEl.innerHTML = colName + (sampleVal ? `<span class="sample">例: ${sampleVal}</span>` : '');

    const arrow = document.createElement('div');
    arrow.className = 'arrow';
    arrow.textContent = '→';

    const select = document.createElement('select');
    select.className = 'mapping-select';
    select.dataset.col = colName;

    const optNone = document.createElement('option');
    optNone.value = '';
    optNone.textContent = '— 未映射 —';
    select.appendChild(optNone);

    FIELD_OPTIONS.forEach(fo => {
      const opt = document.createElement('option');
      opt.value = fo.key;
      opt.textContent = fo.label;
      select.appendChild(opt);
    });

    const currentKey = Object.entries(columnMap).find(([k, v]) => v === colName);
    if (currentKey) select.value = currentKey[0];

    updateSelectStyle(select);
    select.addEventListener('change', () => onMappingChange(select));

    row.appendChild(nameEl);
    row.appendChild(arrow);
    row.appendChild(select);
    mappingList.appendChild(row);
  });
}

function getSampleValue(colName) {
  for (const row of bomRows.slice(0, 3)) {
    const v = String(row[colName] || '').trim();
    if (v) return v.length > 20 ? v.slice(0, 20) + '…' : v;
  }
  return '';
}

function onMappingChange(select) {
  const selectedKey = select.value;
  const colName = select.dataset.col;

  if (selectedKey) {
    const prevCol = columnMap[selectedKey];
    if (prevCol && prevCol !== colName) {
      columnMap[selectedKey] = colName;
      const prevSelect = mappingList.querySelector(`select[data-col="${CSS.escape(prevCol)}"]`);
      if (prevSelect) { prevSelect.value = ''; updateSelectStyle(prevSelect); }
    } else {
      columnMap[selectedKey] = colName;
    }

    const oldKey = Object.entries(columnMap).find(([k, v]) => v === colName && k !== selectedKey);
    if (oldKey) delete columnMap[oldKey[0]];
  } else {
    const oldKey = Object.entries(columnMap).find(([k, v]) => v === colName);
    if (oldKey) delete columnMap[oldKey[0]];
  }

  mappingList.querySelectorAll('select').forEach(s => updateSelectStyle(s));
  updateMappingStatus();
}

function updateSelectStyle(select) {
  select.classList.remove('mapped-required', 'mapped-optional', 'unmapped', 'missing-required');
  const key = select.value;
  if (!key) {
    select.classList.add('unmapped');
  } else {
    const field = BOM_FIELDS.find(f => f.key === key);
    select.classList.add(field?.required ? 'mapped-required' : 'mapped-optional');
  }
}

/* ═══ Analyze Button ═══ */
btnAnalyze.addEventListener('click', () => {
  const missing = BOM_FIELDS.filter(f => f.required && !columnMap[f.key]);
  if (missing.length) {
    const errEl = document.querySelector('.mapping-actions .mapping-error');
    if (errEl) errEl.remove();
    const err = document.createElement('span');
    err.className = 'mapping-error';
    err.textContent = `缺少必填映射：${missing.map(f => f.label).join('、')}，请展开修改映射`;
    btnAnalyze.parentElement.insertBefore(err, btnAnalyze);
    mappingDetail.style.display = 'block';
    mappingToggle.textContent = '收起 ▴';
    return;
  }

  const errEl = document.querySelector('.mapping-actions .mapping-error');
  if (errEl) errEl.remove();

  runAudit();
});

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
  bomRows.forEach((row, i) => {
    const qtyRaw = col(row, 'qty');
    const qtyVal = parseInt(qtyRaw, 10);
    const refs = parseRefs(col(row, 'ref'));
    if (isNaN(qtyVal) || refs.length === 0) return;
    if (refs.length !== qtyVal) {
      addResult('qty', 'error',
        `数量不匹配：基本用量=${qtyVal}，位号数=${refs.length}（${col(row,'ref')}）`,
        [i]);
    }
  });
}

/* ── Rule 3: Duplicate Part Number ── */
function checkDuplicatePN() {
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
  resultsPlaceholder.querySelector('.placeholder-icon').textContent = '📋';
  resultsPlaceholder.querySelector('.placeholder-text').textContent = '确认列映射后，点击"开始 BOM 分析"';
  updateSummary(true);
  bomBody.querySelectorAll('tr').forEach(tr => tr.classList.remove('row-error', 'row-warning'));
}

function renderResults() {
  resultsGroups.innerHTML = '';
  if (!auditResults.length) {
    resultsPlaceholder.style.display = 'flex';
    resultsPlaceholder.querySelector('.placeholder-text').textContent = '审核通过，未发现问题';
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
        'dup-pn': '重复料号', 'dup-mpn': '重复厂商料号'
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
    '类型': { banned:'禁用物料', qty:'数量不匹配', 'dup-pn':'重复料号', 'dup-mpn':'重复厂商料号' }[r.type],
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
