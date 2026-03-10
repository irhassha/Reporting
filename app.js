const excelInput = document.getElementById('excelInput');
const copyTableBtn = document.getElementById('copyTableBtn');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const tableBody = document.getElementById('tableBody');
const remarkBlockTemplate = document.getElementById('remarkBlockTemplate');

const modalBackdrop = document.getElementById('modalBackdrop');
const closeModalBtn = document.getElementById('closeModalBtn');
const applyLogBtn = document.getElementById('applyLogBtn');
const logTextarea = document.getElementById('logTextarea');

const imageInput = document.getElementById('imageInput');
const processImageBtn = document.getElementById('processImageBtn');
const downloadImageBtn = document.getElementById('downloadImageBtn');
const berthCanvas = document.getElementById('berthCanvas');
const ctx = berthCanvas.getContext('2d');
const filterOpacityInput = document.getElementById('filterOpacity');
const filterOpacityValue = document.getElementById('filterOpacityValue');

const metricSelect = document.getElementById('metricSelect');
const groupBySelect = document.getElementById('groupBySelect');
const serviceFilterList = document.getElementById('serviceFilterList');
const performanceChartCanvas = document.getElementById('performanceChart');
const averageChartCanvas = document.getElementById('averageChart');

const resultRows = [];
const analyticsRows = [];
let activeRemarkRowIndex = null;
let loadedImage = null;
let performanceChart = null;
let averageChart = null;

const DISPLAY_COLUMNS = [
  'Vessel Name', 'Service', 'ATB', 'ATD', 'Total (Boxes)', 'Total Teus', 'Crane Intencity (CI)', 'Vessel Rate (VR)', 'ATB to Fst lift', 'Lst lift to ATD'
];

const MASTER_WINDOW_SCHEDULE = {
  'JPI-A': { day: 0, time: '23:00' },
  IN1: { day: 0, time: '13:00' },
  SEAGULL: { day: 0, time: '18:00' },
  KCI: { day: 1, time: '22:00' },
  CMI: { day: 2, time: '00:01' },
  CMI2: { day: 2, time: '00:01' },
  IA8: { day: 3, time: '13:00' },
  CIM: { day: 3, time: '12:00' },
  IA15: { day: 4, time: '09:00' },
  JKF: { day: 4, time: '18:00' },
  JTH: { day: 5, time: '13:00' },
  SE8: { day: 5, time: '13:00' },
  'JPI-B': { day: 6, time: '08:00' },
  KIS: { day: 6, time: '16:00' },
  CIT: { day: 0, time: '01:00' },
  IA1: { day: 6, time: '23:00' }
};

const MONTH_ORDER = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const columnAliasMap = {
  'Vessel Name': ['Vessel Name', 'Vessel', 'VesselName'],
  Service: ['Service', 'Service Name'],
  ATB: ['ATB', 'Actual Time Berthing'],
  ATD: ['ATD', 'Actual Time Departure', 'Departure', 'ATD/Departure'],
  'Total (Boxes)': ['Total (Boxes)', 'Total Boxes', 'Boxes'],
  'Total Teus': ['Total Teus', 'Teus', 'Total TEUS'],
  'Crane Intencity (CI)': ['Crane Intencity (CI)', 'CI', 'Crane Intensity'],
  'Vessel Rate (VR)': ['Vessel Rate (VR)', 'VR', 'Vessel Rate'],
  'ATB to Fst lift': ['ATB to Fst lift', 'ATB to First Lift'],
  'Lst lift to ATD': ['Lst lift to ATD', 'Last Lift to ATD'],
  Operator: ['Operator', 'OPR', 'Operator Name']
};

excelInput.addEventListener('change', handleExcelUpload);
copyTableBtn.addEventListener('click', copyTableToExcel);
exportExcelBtn.addEventListener('click', exportFilteredData);
closeModalBtn.addEventListener('click', closeModal);
modalBackdrop.addEventListener('click', (event) => {
  if (event.target === modalBackdrop) closeModal();
});
applyLogBtn.addEventListener('click', applyLogParse);
imageInput.addEventListener('change', handleImageUpload);
processImageBtn.addEventListener('click', processBerthImage);
downloadImageBtn.addEventListener('click', downloadCanvasImage);
filterOpacityInput.addEventListener('input', () => {
  filterOpacityValue.textContent = `${Math.round(Number(filterOpacityInput.value) * 100)}%`;
});
metricSelect?.addEventListener('change', renderAnalyticsCharts);
groupBySelect?.addEventListener('change', renderAnalyticsCharts);

function normalizeHeader(header) {
  return String(header || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function findValueByAlias(row, aliasList) {
  const normalizedMap = new Map(Object.keys(row || {}).map((key) => [normalizeHeader(key), row[key]]));
  for (const alias of aliasList) {
    const v = normalizedMap.get(normalizeHeader(alias));
    if (v !== undefined && v !== null) return v;
  }
  return '';
}

function toDisplayText(value) {
  if (value === undefined || value === null) return '';
  if (value instanceof Date) return value.toLocaleString();
  return String(value);
}

function formatVR(rawValue) {
  const clean = String(rawValue ?? '').trim().replace(/,/g, '.').replace(/[^\d.-]/g, '');
  const num = Number(clean);
  return Number.isFinite(num) ? num.toFixed(1) : '';
}

function createRowFromSource(sourceRow) {
  const row = {};
  DISPLAY_COLUMNS.forEach((col) => {
    const raw = findValueByAlias(sourceRow, columnAliasMap[col] || [col]);
    row[col] = col === 'Vessel Rate (VR)' ? formatVR(raw) : toDisplayText(raw);
  });
  row.Operator = toDisplayText(findValueByAlias(sourceRow, columnAliasMap.Operator || ['Operator']));
  row['Window Status'] = toDisplayText(sourceRow['Window Status'] || sourceRow.Window || sourceRow['VS Window'] || '');
  row['BOA Status'] = toDisplayText(sourceRow['BOA Status'] || sourceRow.BOA || '');
  row['MM-YY'] = toDisplayText(sourceRow['MM-YY'] || sourceRow.MMYY || sourceRow.Month || '');
  row.remarkBlocks = [];
  return row;
}

function operatorAllowed(sourceRow) {
  const op = toDisplayText(findValueByAlias(sourceRow, columnAliasMap.Operator || ['Operator'])).trim().toUpperCase();
  return op === 'MCC' || op === 'MAE';
}

async function handleExcelUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

  const enrichedRows = enrichOperationalMetrics(rawRows);

  resultRows.length = 0;
  analyticsRows.length = 0;
  enrichedRows.filter(operatorAllowed).forEach((src) => {
    resultRows.push(createRowFromSource(src));
    analyticsRows.push(src);
  });

  renderTable();
  renderServiceFilters();
  renderAnalyticsCharts();
  const ready = resultRows.length > 0;
  copyTableBtn.disabled = !ready;
  exportExcelBtn.disabled = !ready;
}

function renderTable() {
  tableBody.innerHTML = '';
  if (!resultRows.length) {
    tableBody.innerHTML = '<tr><td class="px-4 py-6 text-center text-slate-500" colspan="11">No rows match Operator MCC/MAE.</td></tr>';
    return;
  }

  resultRows.forEach((rowData, rowIndex) => {
    const tr = document.createElement('tr');
    tr.className = 'border-b border-slate-200 dark:border-slate-800 align-top';

    DISPLAY_COLUMNS.forEach((col) => {
      const td = document.createElement('td');
      td.className = 'px-2 py-2 text-xs break-words';
      td.textContent = rowData[col] || '';
      tr.appendChild(td);
    });

    const remarkTd = document.createElement('td');
    remarkTd.className = 'px-2 py-2 align-top min-w-[26rem]';
    remarkTd.appendChild(buildRemarkCell(rowIndex));
    tr.appendChild(remarkTd);
    tableBody.appendChild(tr);
  });
}

function isDynamicRemark(value) {
  return ['QC Trouble Spreader', 'DOOG Units', 'LOOG Units', 'D/L OOG Units'].includes(value);
}

function blockComplete(block) {
  if (!block?.value) return false;
  return isDynamicRemark(block.value) ? /\d+/.test(String(block.extra || '')) : true;
}

function formatPresetBlock(block) {
  if (!block?.value) return '';
  if (block.value === 'QC Trouble Spreader') {
    const lane = String(block.extra || '').replace(/\D/g, '').slice(0, 3);
    return lane ? `QC ${lane} Trouble Spreader` : 'QC Trouble Spreader';
  }
  if (['DOOG Units', 'LOOG Units', 'D/L OOG Units'].includes(block.value)) {
    const num = String(block.extra || '').replace(/\D/g, '');
    const prefix = block.value.replace(' Units', '');
    return num ? `${prefix} ${num} Units` : block.value;
  }
  return block.value;
}

function buildRemarkCell(rowIndex) {
  const row = resultRows[rowIndex];
  const container = document.createElement('div');
  container.className = 'space-y-2';

  const list = document.createElement('div');
  list.className = 'space-y-2';

  row.remarkBlocks.forEach((block, blockIndex) => {
    if (block.kind === 'log') {
      const wrap = document.createElement('div');
      wrap.className = 'rounded-xl border border-primary/30 bg-primary/10 p-2 text-xs flex items-start justify-between gap-2';
      const text = document.createElement('div');
      text.className = 'whitespace-pre-wrap';
      text.textContent = block.text;
      const remove = document.createElement('button');
      remove.className = 'h-7 w-7 rounded bg-slate-200 dark:bg-slate-700';
      remove.textContent = '✕';
      remove.addEventListener('click', () => {
        row.remarkBlocks.splice(blockIndex, 1);
        renderTable();
      });
      wrap.append(text, remove);
      list.appendChild(wrap);
      return;
    }

    const node = remarkBlockTemplate.content.cloneNode(true);
    const select = node.querySelector('.remark-select');
    const extraInput = node.querySelector('.remark-extra');
    const chip = node.querySelector('.remark-chip');
    const editBtn = node.querySelector('.edit-remark-btn');
    const resetBtn = node.querySelector('.reset-remark-btn');
    const removeBtn = node.querySelector('.remove-remark-btn');

    const dynamicMap = {
      'QC Trouble Spreader': 'QC [801-808] Trouble Spreader',
      'DOOG Units': 'DOOG [number] Units',
      'LOOG Units': 'LOOG [number] Units',
      'D/L OOG Units': 'D/L OOG [number] Units'
    };

    select.value = block.value || '';
    extraInput.value = block.extra || '';

    const updateMode = () => {
      const done = blockComplete(block) && !block.editing;
      if (done) {
        select.classList.add('hidden');
        extraInput.classList.add('hidden');
        chip.classList.remove('hidden');
        chip.textContent = formatPresetBlock(block);
        editBtn.classList.remove('hidden');
        resetBtn.classList.remove('hidden');
      } else {
        select.classList.remove('hidden');
        if (dynamicMap[block.value]) {
          extraInput.classList.remove('hidden');
          extraInput.placeholder = dynamicMap[block.value];
        } else {
          extraInput.classList.add('hidden');
        }
        chip.classList.add('hidden');
        editBtn.classList.add('hidden');
        resetBtn.classList.add('hidden');
      }
    };

    select.addEventListener('change', () => {
      block.value = select.value;
      if (!isDynamicRemark(block.value)) block.extra = '';
      block.editing = isDynamicRemark(block.value);
      updateMode();
    });

    extraInput.addEventListener('input', () => {
      block.extra = extraInput.value;
      if (blockComplete(block)) block.editing = false;
      updateMode();
    });

    editBtn.addEventListener('click', () => {
      block.editing = true;
      updateMode();
    });

    resetBtn.addEventListener('click', () => {
      block.value = '';
      block.extra = '';
      block.editing = true;
      updateMode();
    });

    removeBtn.addEventListener('click', () => {
      row.remarkBlocks.splice(blockIndex, 1);
      renderTable();
    });

    updateMode();
    list.appendChild(node);
  });

  const actions = document.createElement('div');
  actions.className = 'flex gap-2';

  const addRemarkBtn = document.createElement('button');
  addRemarkBtn.className = 'h-8 px-3 rounded bg-slate-100 dark:bg-slate-700 text-xs';
  addRemarkBtn.type = 'button';
  addRemarkBtn.textContent = '(+) Add Remark';
  addRemarkBtn.addEventListener('click', () => {
    row.remarkBlocks.push({ kind: 'preset', value: '', extra: '', editing: true });
    renderTable();
  });

  const pasteLogBtn = document.createElement('button');
  pasteLogBtn.className = 'h-8 px-3 rounded bg-slate-100 dark:bg-slate-700 text-xs';
  pasteLogBtn.type = 'button';
  pasteLogBtn.textContent = '📋 Paste Log';
  pasteLogBtn.addEventListener('click', () => {
    activeRemarkRowIndex = rowIndex;
    logTextarea.value = '';
    modalBackdrop.classList.remove('hidden');
  });

  actions.append(addRemarkBtn, pasteLogBtn);
  container.append(list, actions);
  return container;
}

function closeModal() {
  modalBackdrop.classList.add('hidden');
}

function applyLogParse() {
  if (activeRemarkRowIndex === null) return;
  const parsed = parseOperationalLog(logTextarea.value || '');
  resultRows[activeRemarkRowIndex].remarkBlocks.push({ kind: 'log', text: parsed });
  closeModal();
  renderTable();
}

function extractTime(text, regexA, regexB) {
  const a = text.match(regexA);
  if (a?.[1]) return toHHMM(a[1]);
  const b = text.match(regexB);
  if (b?.[1]) return toHHMM(b[1]);
  return '';
}

function toHHMM(raw) {
  const [h, m] = String(raw).split(':').map(Number);
  if (Number.isNaN(h) || Number.isNaN(m)) return '';
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function parseOperationalLog(text) {
  const gangWayTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Gang\s*Way\s*Down/i, /Gang\s*Way\s*Down[^\n]*(\d{1,2}:\d{2})/i);
  const quarantineTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Quarantine\s*Clearance/i, /Quarantine\s*Clearance[^\n]*(\d{1,2}:\d{2})/i);
  const commenceTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Commence(?:\s+Loading|\s+Discharge|\b)/i, /Commence(?:\s+Loading|\s+Discharge|\b)[^\n]*(\d{1,2}:\d{2})/i);

  const tkbmTime = quarantineTime ? addRandomMinutes(quarantineTime, 3, 7) : 'N/A';
  const breaks = detectMealBreaks(gangWayTime, commenceTime);
  const breakText = breaks.length ? `${breaks.join(', ')}, ` : '';

  return `Gang way ready ${gangWayTime || 'N/A'}, ${breakText}Quarantine ${quarantineTime || 'N/A'}, TKBM ${tkbmTime}, Start Ops ${commenceTime || 'N/A'}`;
}

function parseMinutes(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function addRandomMinutes(hhmm, min, max) {
  const base = parseMinutes(hhmm);
  const add = Math.floor(Math.random() * (max - min + 1)) + min;
  const t = (base + add) % (24 * 60);
  return `${String(Math.floor(t / 60)).padStart(2, '0')}:${String(t % 60).padStart(2, '0')}`;
}

function detectMealBreaks(gangWayTime, commenceTime) {
  if (!gangWayTime || !commenceTime) return [];
  const start = parseMinutes(gangWayTime);
  let end = parseMinutes(commenceTime);
  if (end < start) end += 24 * 60;
  const windows = [
    { from: 690, to: 780, label: 'Meal Break S1' },
    { from: 1050, to: 1140, label: 'Meal Break S2' },
    { from: 210, to: 300, label: 'Meal Break S3' }
  ];
  return windows.filter((w) => {
    const a = w.from < start ? w.from + 1440 : w.from;
    const b = w.to < start ? w.to + 1440 : w.to;
    return Math.max(start, a) <= Math.min(end, b);
  }).map((w) => w.label);
}

function getJoinedRemarks(row) {
  return row.remarkBlocks.map((block) => (block.kind === 'log' ? block.text : formatPresetBlock(block))).filter(Boolean).join('\n');
}

function escapeTSVCell(value) {
  const text = String(value ?? '').replace(/\r/g, '');
  if (!/[\t\n"]/.test(text)) return text;
  return `"${text.replace(/"/g, '""')}"`;
}

async function copyTableToExcel() {
  if (!resultRows.length) return;
  const rows = [[...DISPLAY_COLUMNS, 'Remark'].map(escapeTSVCell).join('\t')];
  resultRows.forEach((row) => {
    const vals = DISPLAY_COLUMNS.map((c) => escapeTSVCell(row[c] || ''));
    vals.push(escapeTSVCell(getJoinedRemarks(row)));
    rows.push(vals.join('\t'));
  });
  await navigator.clipboard.writeText(rows.join('\n'));
  copyTableBtn.textContent = '✅ Copied';
  setTimeout(() => (copyTableBtn.textContent = '📋 Copy Table'), 1200);
}

function exportFilteredData() {
  const exportData = resultRows.map((row) => ({
    ...Object.fromEntries(DISPLAY_COLUMNS.map((c) => [c, row[c] || ''])),
    Operator: row.Operator || '',
    'MM-YY': row['MM-YY'] || '',
    'Window Status': row['Window Status'] || '',
    'BOA Status': row['BOA Status'] || '',
    Remark: getJoinedRemarks(row)
  }));
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');
  XLSX.writeFile(wb, 'maersk_weekly_filtered.xlsx');
}


function toDateValue(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) return raw;
  if (typeof raw === 'number') {
    const parsed = XLSX.SSF.parse_date_code(raw);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S || 0);
  }
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? null : date;
}

function parseTimeToHM(timeString) {
  const [h, m] = String(timeString || '').split(':').map(Number);
  return { h: Number.isFinite(h) ? h : 0, m: Number.isFinite(m) ? m : 0 };
}

function getClosestWindowDateTime(ataDate, schedule) {
  if (!ataDate || !schedule) return null;
  const { h, m } = parseTimeToHM(schedule.time);
  const baseSunday = new Date(ataDate);
  baseSunday.setHours(0, 0, 0, 0);
  baseSunday.setDate(baseSunday.getDate() - baseSunday.getDay());
  const candidate = new Date(baseSunday);
  candidate.setDate(baseSunday.getDate() + schedule.day);
  candidate.setHours(h, m, 0, 0);

  const prev = new Date(candidate);
  prev.setDate(candidate.getDate() - 7);
  const next = new Date(candidate);
  next.setDate(candidate.getDate() + 7);

  const options = [prev, candidate, next];
  return options.reduce((best, current) => (Math.abs(current - ataDate) < Math.abs(best - ataDate) ? current : best), options[0]);
}

function normalizeWindowStatus(value) {
  const text = String(value || '').trim().toLowerCase();
  if (!text) return '';
  if (text.includes('out')) return 'Out of Window';
  if (text.includes('on')) return 'On Window';
  return '';
}

function parseMonthLabel(rawValue, fallbackDate) {
  const rawText = String(rawValue || '').trim();
  if (/^[A-Za-z]{3}-\d{2}$/.test(rawText)) return rawText;

  const fromDate = toDateValue(rawValue) || fallbackDate;
  if (!fromDate) return '';
  return `${MONTH_ORDER[fromDate.getMonth()]}-${String(fromDate.getFullYear()).slice(-2)}`;
}

function enrichOperationalMetrics(rows) {
  return rows.map((row) => {
    const service = toDisplayText(findValueByAlias(row, columnAliasMap.Service || ['Service'])).trim().toUpperCase();
    const ata = toDateValue(findValueByAlias(row, ['ATA', 'Actual Time Arrival', 'Actual Time of Arrival']));
    const etb = toDateValue(findValueByAlias(row, ['ETB', 'Estimated Time Berthing', 'Est Time Berthing']));

    let windowStatus = normalizeWindowStatus(findValueByAlias(row, ['VS Window']));
    if (!windowStatus && ata && MASTER_WINDOW_SCHEDULE[service]) {
      const target = getClosestWindowDateTime(ata, MASTER_WINDOW_SCHEDULE[service]);
      const gapHours = Math.abs((ata - target) / 36e5);
      windowStatus = gapHours > 4 ? 'Out of Window' : 'On Window';
    }

    let boaStatus = '';
    if (ata && etb) {
      const btr = new Date(ata.getTime() + 2 * 60 * 60 * 1000);
      boaStatus = btr < etb ? 'BOA' : 'NOT BOA';
    }

    const monthLabel = parseMonthLabel(findValueByAlias(row, ['MM-YY', 'MM/YY', 'Month']), ata || etb);

    return {
      ...row,
      'Window Status': windowStatus || 'Unknown',
      'BOA Status': boaStatus || 'Unknown',
      'MM-YY': monthLabel,
      _service: service,
      _operator: toDisplayText(findValueByAlias(row, columnAliasMap.Operator || ['Operator'])).trim().toUpperCase(),
      _vr: Number.parseFloat(formatVR(findValueByAlias(row, columnAliasMap['Vessel Rate (VR)'] || ['Vessel Rate (VR)']))) || 0,
      _ci: Number.parseFloat(String(findValueByAlias(row, columnAliasMap['Crane Intencity (CI)'] || ['CI'])).replace(/,/g, '.')) || 0
    };
  });
}

function renderServiceFilters() {
  if (!serviceFilterList) return;
  const services = [...new Set(analyticsRows.map((row) => row._service).filter(Boolean))].sort();
  serviceFilterList.innerHTML = '';
  if (!services.length) {
    serviceFilterList.innerHTML = '<p class="text-xs text-slate-400">Upload Excel data to load services.</p>';
    return;
  }

  services.forEach((service) => {
    const id = `svc-${service.replace(/[^a-z0-9]/gi, '-')}`;
    const label = document.createElement('label');
    label.className = 'flex items-center gap-2';
    label.innerHTML = `<input id="${id}" data-service="${service}" type="checkbox" class="rounded border-slate-600 bg-slate-900" checked /><span>${service}</span>`;
    serviceFilterList.appendChild(label);
  });

  serviceFilterList.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
    checkbox.addEventListener('change', renderAnalyticsCharts);
  });
}

function selectedServicesSet() {
  const set = new Set();
  serviceFilterList?.querySelectorAll('input[type="checkbox"][data-service]').forEach((el) => {
    if (el.checked) set.add(el.dataset.service);
  });
  return set;
}

function monthSortKey(label) {
  const [mon, yr] = String(label || '').split('-');
  const monthIndex = MONTH_ORDER.indexOf(mon);
  const year = Number.parseInt(yr, 10);
  return (Number.isFinite(year) ? year : 0) * 12 + (monthIndex >= 0 ? monthIndex : 0);
}

function generateWindowPerformanceChart(excelData, options = {}) {
  const metric = options.metric || metricSelect?.value || 'window';
  const groupBy = options.groupBy || groupBySelect?.value || 'month';
  const selectedServices = selectedServicesSet();
  const scoped = excelData.filter((row) => !selectedServices.size || selectedServices.has(row._service));

  const groupKey = (row) => {
    if (groupBy === 'service') return row._service || 'Unknown';
    if (groupBy === 'operator') return row._operator || 'Unknown';
    return row['MM-YY'] || 'Unknown';
  };

  const grouped = new Map();
  scoped.forEach((row) => {
    const key = groupKey(row);
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(row);
  });

  let labels = [...grouped.keys()];
  if (groupBy === 'month') labels = labels.sort((a, b) => monthSortKey(a) - monthSortKey(b));
  else labels = labels.sort();

  const payload = labels.map((label) => {
    const rows = grouped.get(label) || [];
    const total = rows.length || 1;
    const onWindow = rows.filter((r) => r['Window Status'] === 'On Window').length;
    const outWindow = rows.filter((r) => r['Window Status'] === 'Out of Window').length;
    const boa = rows.filter((r) => r['BOA Status'] === 'BOA').length;
    const notBoa = rows.filter((r) => r['BOA Status'] === 'NOT BOA').length;
    const avgVR = rows.reduce((a, r) => a + (r._vr || 0), 0) / total;
    const avgCI = rows.reduce((a, r) => a + (r._ci || 0), 0) / total;

    return {
      label,
      windowOnPct: (onWindow / total) * 100,
      windowOutPct: (outWindow / total) * 100,
      boaPct: (boa / total) * 100,
      notBoaPct: (notBoa / total) * 100,
      avgVR,
      avgCI
    };
  });

  return {
    metric,
    labels: payload.map((p) => p.label),
    onWindow: payload.map((p) => p.windowOnPct),
    outWindow: payload.map((p) => p.windowOutPct),
    boa: payload.map((p) => p.boaPct),
    notBoa: payload.map((p) => p.notBoaPct),
    avgVR: payload.map((p) => p.avgVR),
    avgCI: payload.map((p) => p.avgCI)
  };
}

function chartThemeOptions(isPercentMetric) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { labels: { color: '#e2e8f0' } },
      datalabels: {
        color: '#f8fafc',
        anchor: 'end',
        align: 'top',
        formatter: (v) => (isPercentMetric ? `${v.toFixed(1)}%` : Number(v).toFixed(2))
      }
    },
    scales: {
      x: { ticks: { color: '#e2e8f0' }, grid: { color: 'rgba(226,232,240,0.2)' } },
      y: {
        beginAtZero: true,
        ticks: {
          color: '#e2e8f0',
          callback: (v) => (isPercentMetric ? `${v}%` : v)
        },
        grid: { color: 'rgba(226,232,240,0.2)' }
      }
    }
  };
}

function renderAnalyticsCharts() {
  if (!performanceChartCanvas || !averageChartCanvas || !analyticsRows.length || typeof Chart === 'undefined') return;
  if (window.ChartDataLabels) Chart.register(window.ChartDataLabels);

  const computed = generateWindowPerformanceChart(analyticsRows);
  const isWindow = computed.metric === 'window';
  const isBoa = computed.metric === 'boa';
  const isPercent = isWindow || isBoa;

  let datasets;
  let chartType = 'bar';
  if (isWindow) {
    datasets = [
      { label: 'On Window', data: computed.onWindow, backgroundColor: '#105b8a' },
      { label: 'Out of Window', data: computed.outWindow, backgroundColor: '#e57347' }
    ];
  } else if (isBoa) {
    datasets = [
      { label: 'BOA', data: computed.boa, backgroundColor: '#105b8a' },
      { label: 'NOT BOA', data: computed.notBoa, backgroundColor: '#e57347' }
    ];
  } else if (computed.metric === 'vr') {
    datasets = [{ label: 'Average Vessel Rate', data: computed.avgVR, backgroundColor: '#105b8a' }];
  } else {
    chartType = 'line';
    datasets = [{ label: 'Average CI', data: computed.avgCI, borderColor: '#e57347', backgroundColor: 'rgba(229,115,71,0.25)', tension: 0.3, fill: true }];
  }

  if (performanceChart) performanceChart.destroy();
  performanceChart = new Chart(performanceChartCanvas, {
    type: chartType,
    data: { labels: computed.labels, datasets },
    options: chartThemeOptions(isPercent),
    plugins: [window.ChartDataLabels].filter(Boolean)
  });

  const monthlyData = generateWindowPerformanceChart(analyticsRows, { metric: computed.metric, groupBy: 'month' });
  const averageValue = (() => {
    if (isWindow) return monthlyData.onWindow.reduce((a, b) => a + b, 0) / Math.max(monthlyData.onWindow.length, 1);
    if (isBoa) return monthlyData.boa.reduce((a, b) => a + b, 0) / Math.max(monthlyData.boa.length, 1);
    if (computed.metric === 'vr') return monthlyData.avgVR.reduce((a, b) => a + b, 0) / Math.max(monthlyData.avgVR.length, 1);
    return monthlyData.avgCI.reduce((a, b) => a + b, 0) / Math.max(monthlyData.avgCI.length, 1);
  })();

  if (averageChart) averageChart.destroy();
  averageChart = new Chart(averageChartCanvas, {
    type: 'bar',
    data: {
      labels: ['AVERAGE JAN-OCT'],
      datasets: [{
        label: isWindow ? 'On Window %' : isBoa ? 'BOA %' : computed.metric === 'vr' ? 'Avg VR' : 'Avg CI',
        data: [averageValue],
        backgroundColor: '#105b8a'
      }]
    },
    options: chartThemeOptions(isPercent),
    plugins: [window.ChartDataLabels].filter(Boolean)
  });
}

function handleImageUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    const img = new Image();
    img.onload = () => {
      loadedImage = img;
      berthCanvas.width = img.width;
      berthCanvas.height = img.height;
      ctx.drawImage(img, 0, 0);
      processImageBtn.disabled = false;
      downloadImageBtn.disabled = true;
    };
    img.src = reader.result;
  };
  reader.readAsDataURL(file);
}

function processBerthImage() {
  if (!loadedImage) return;
  ctx.drawImage(loadedImage, 0, 0);
  const w = berthCanvas.width;
  const h = berthCanvas.height;
  const leftKeep = Math.max(Math.round(w * 0.12), 110);
  const opacity = Math.min(Math.max(Number(filterOpacityInput.value) || 0.8, 0), 1);

  const source = ctx.getImageData(0, 0, w, h);
  const src = source.data;

  const target = { r: 145, g: 230, b: 245 };
  const tolerance = 72;
  const mask = new Uint8Array(w * h);

  for (let y = 0; y < h; y++) {
    for (let x = leftKeep; x < w; x++) {
      const idx = y * w + x;
      const i = idx * 4;
      const r = src[i];
      const g = src[i + 1];
      const b = src[i + 2];
      const a = src[i + 3];
      if (!a) continue;
      const dist = Math.hypot(r - target.r, g - target.g, b - target.b);
      if (dist <= tolerance) mask[idx] = 1;
    }
  }

  const output = ctx.createImageData(w, h);
  const out = output.data;

  for (let y = 0; y < h; y++) {
    for (let x = 0; x < leftKeep; x++) {
      const idx = y * w + x;
      const i = idx * 4;
      out[i] = src[i];
      out[i + 1] = src[i + 1];
      out[i + 2] = src[i + 2];
      out[i + 3] = src[i + 3];
    }
  }

  for (let y = 0; y < h; y++) {
    for (let x = leftKeep; x < w; x++) {
      const idx = y * w + x;
      const i = idx * 4;
      out[i] = src[i];
      out[i + 1] = src[i + 1];
      out[i + 2] = src[i + 2];
      out[i + 3] = src[i + 3];
    }
  }

  const visited = new Uint8Array(w * h);
  const n4 = [1, -1, w, -w];
  for (let y = 0; y < h; y++) {
    for (let x = leftKeep; x < w; x++) {
      const seed = y * w + x;
      if (!mask[seed] || visited[seed]) continue;

      const q = [seed];
      visited[seed] = 1;
      let head = 0;
      let minX = x;
      let maxX = x;
      let minY = y;
      let maxY = y;
      let count = 0;

      while (head < q.length) {
        const cur = q[head++];
        const cx = cur % w;
        const cy = Math.floor(cur / w);
        count++;
        if (cx < minX) minX = cx;
        if (cx > maxX) maxX = cx;
        if (cy < minY) minY = cy;
        if (cy > maxY) maxY = cy;

        for (const d of n4) {
          const n = cur + d;
          if (n < 0 || n >= w * h || visited[n] || !mask[n]) continue;
          const nx = n % w;
          const ny = Math.floor(n / w);
          if (nx < leftKeep || ny < 0 || ny >= h) continue;
          visited[n] = 1;
          q.push(n);
        }
      }

      const boxW = maxX - minX + 1;
      const boxH = maxY - minY + 1;
      const area = boxW * boxH;
      const density = count / Math.max(area, 1);
      if (count < 80 || boxW < 10 || boxH < 10 || density < 0.09 || boxW > w * 0.45 || boxH > h * 0.45) continue;

      const pad = 2;
      const sx = Math.max(leftKeep, minX - pad);
      const ex = Math.min(w - 1, maxX + pad);
      const sy = Math.max(0, minY - pad);
      const ey = Math.min(h - 1, maxY + pad);

      for (let yy = sy; yy <= ey; yy++) {
        for (let xx = sx; xx <= ex; xx++) {
          const idx = yy * w + xx;
          const i = idx * 4;
          out[i] = src[i];
          out[i + 1] = src[i + 1];
          out[i + 2] = src[i + 2];
          out[i + 3] = src[i + 3];
        }
      }
    }
  }

  for (let y = 0; y < h; y++) {
    for (let x = leftKeep; x < w; x++) {
      const idx = y * w + x;
      if (mask[idx]) continue;
      const i = idx * 4;
      const r = out[i];
      const g = out[i + 1];
      const b = out[i + 2];
      out[i] = Math.round(r * (1 - opacity) + 204 * opacity);
      out[i + 1] = Math.round(g * (1 - opacity) + 204 * opacity);
      out[i + 2] = Math.round(b * (1 - opacity) + 204 * opacity);
      out[i + 3] = 255;
    }
  }

  ctx.putImageData(output, 0, 0);
  downloadImageBtn.disabled = false;
}
function downloadCanvasImage() {
  const link = document.createElement('a');
  link.download = 'berth_window_filtered.png';
  link.href = berthCanvas.toDataURL('image/png');
  link.click();
}
