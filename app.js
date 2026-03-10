const excelInput = document.getElementById('excelInput');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const tableBody = document.getElementById('tableBody');
const remarkCellTemplate = document.getElementById('remarkCellTemplate');

const modalBackdrop = document.getElementById('modalBackdrop');
const closeModalBtn = document.getElementById('closeModalBtn');
const applyLogBtn = document.getElementById('applyLogBtn');
const logTextarea = document.getElementById('logTextarea');

const imageInput = document.getElementById('imageInput');
const processImageBtn = document.getElementById('processImageBtn');
const downloadImageBtn = document.getElementById('downloadImageBtn');
const berthCanvas = document.getElementById('berthCanvas');
const ctx = berthCanvas.getContext('2d');

const resultRows = [];
let activeRemarkOutput = null;
let loadedImage = null;

const DISPLAY_COLUMNS = [
  'Vessel Name',
  'Service',
  'ATB',
  'ATD',
  'Total (Boxes)',
  'Total Teus',
  'Crane Intencity (CI)',
  'Vessel Rate (VR)',
  'ATB to Fst lift',
  'Lst lift to ATD'
];

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
exportExcelBtn.addEventListener('click', exportFilteredData);
closeModalBtn.addEventListener('click', closeModal);
modalBackdrop.addEventListener('click', (event) => {
  if (event.target === modalBackdrop) closeModal();
});
applyLogBtn.addEventListener('click', applyLogParse);

imageInput.addEventListener('change', handleImageUpload);
processImageBtn.addEventListener('click', processBerthImage);
downloadImageBtn.addEventListener('click', downloadCanvasImage);

function normalizeHeader(header) {
  return String(header || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function findValueByAlias(row, aliasList) {
  const normalizedMap = new Map(
    Object.keys(row || {}).map((key) => [normalizeHeader(key), row[key]])
  );
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

function createRowFromSource(sourceRow) {
  const row = {};
  DISPLAY_COLUMNS.forEach((col) => {
    row[col] = toDisplayText(findValueByAlias(sourceRow, columnAliasMap[col] || [col]));
  });
  row.Remark = '';
  return row;
}

function operatorAllowed(sourceRow) {
  const op = toDisplayText(findValueByAlias(sourceRow, columnAliasMap.Operator || ['Operator']))
    .trim()
    .toUpperCase();
  return op === 'MCC' || op === 'MAE';
}

async function handleExcelUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: 'array', cellDates: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: '' });

  resultRows.length = 0;
  rawRows.filter(operatorAllowed).forEach((src) => resultRows.push(createRowFromSource(src)));

  renderTable();
  exportExcelBtn.disabled = resultRows.length === 0;
}

function renderTable() {
  tableBody.innerHTML = '';
  if (!resultRows.length) {
    tableBody.innerHTML = '<tr><td colspan="11" class="empty">No rows match Operator MCC/MAE.</td></tr>';
    return;
  }

  resultRows.forEach((rowData, rowIndex) => {
    const tr = document.createElement('tr');
    DISPLAY_COLUMNS.forEach((col) => {
      const td = document.createElement('td');
      td.textContent = rowData[col] || '';
      tr.appendChild(td);
    });

    const remarkTd = document.createElement('td');
    const cell = remarkCellTemplate.content.cloneNode(true);
    const select = cell.querySelector('.remark-select');
    const extraInput = cell.querySelector('.remark-extra');
    const pasteBtn = cell.querySelector('.paste-log-btn');
    const output = cell.querySelector('.remark-output');

    select.addEventListener('change', () => {
      const dynamicMap = {
        'QC Trouble Spreader': 'QC [801-808] Trouble Spreader',
        'DOOG Units': 'DOOG [number] Units',
        'LOOG Units': 'LOOG [number] Units',
        'D/L OOG Units': 'D/L OOG [number] Units'
      };

      if (dynamicMap[select.value]) {
        extraInput.classList.remove('hidden');
        extraInput.placeholder = dynamicMap[select.value];
      } else {
        extraInput.classList.add('hidden');
        extraInput.value = '';
      }
      updateRemarkOutput(rowIndex, output, select.value, extraInput.value);
    });

    extraInput.addEventListener('input', () => {
      updateRemarkOutput(rowIndex, output, select.value, extraInput.value);
    });

    pasteBtn.addEventListener('click', () => {
      activeRemarkOutput = { rowIndex, output };
      logTextarea.value = '';
      modalBackdrop.classList.remove('hidden');
      modalBackdrop.setAttribute('aria-hidden', 'false');
    });

    output.textContent = rowData.Remark;
    remarkTd.appendChild(cell);
    tr.appendChild(remarkTd);
    tableBody.appendChild(tr);
  });
}

function updateRemarkOutput(rowIndex, outputEl, selectedValue, extraValue) {
  if (!selectedValue) {
    resultRows[rowIndex].Remark = '';
    outputEl.textContent = '';
    return;
  }

  let content = selectedValue;
  if (selectedValue === 'QC Trouble Spreader') {
    const lane = String(extraValue || '').replace(/\D/g, '').slice(0, 3);
    content = lane ? `QC ${lane} Trouble Spreader` : 'QC Trouble Spreader';
  } else if (['DOOG Units', 'LOOG Units', 'D/L OOG Units'].includes(selectedValue)) {
    const num = String(extraValue || '').replace(/\D/g, '');
    const prefix = selectedValue.replace(' Units', '');
    content = num ? `${prefix} ${num} Units` : selectedValue;
  }

  const existingLog = extractLogRemark(resultRows[rowIndex].Remark);
  resultRows[rowIndex].Remark = [content, existingLog].filter(Boolean).join(' | ');
  outputEl.textContent = resultRows[rowIndex].Remark;
}

function extractLogRemark(remarkText) {
  if (!remarkText) return '';
  const seg = remarkText
    .split(' | ')
    .find((part) => part.toLowerCase().includes('gang way ready'));
  return seg || '';
}

function closeModal() {
  modalBackdrop.classList.add('hidden');
  modalBackdrop.setAttribute('aria-hidden', 'true');
}

function applyLogParse() {
  if (!activeRemarkOutput) return;
  const logText = logTextarea.value || '';
  const parsed = parseOperationalLog(logText);

  const existing = resultRows[activeRemarkOutput.rowIndex].Remark;
  const baseRemark = existing
    .split(' | ')
    .filter((part) => part && !part.toLowerCase().includes('gang way ready'));
  baseRemark.push(parsed);

  const merged = baseRemark.filter(Boolean).join(' | ');
  resultRows[activeRemarkOutput.rowIndex].Remark = merged;
  activeRemarkOutput.output.textContent = merged;
  closeModal();
}

function parseOperationalLog(text) {
  const gangWayTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Gang\s*Way\s*Down/i);
  const quarantineTime = extractTime(
    text,
    /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Quarantine\s*Clearance/i
  );
  const commenceTime = extractTime(
    text,
    /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Commence(?:\s+Loading|\s+Discharge|\b)/i
  );

  const gw = gangWayTime || 'N/A';
  const qt = quarantineTime || 'N/A';
  const cm = commenceTime || 'N/A';

  const tkbmTime = quarantineTime ? addRandomMinutes(quarantineTime, 3, 7) : 'N/A';
  const breaks = detectMealBreaks(gangWayTime, commenceTime);
  const breakText = breaks.length ? `${breaks.join(', ')}, ` : '';

  return `Gang way ready ${gw}, ${breakText}Quarantine ${qt}, TKBM ${tkbmTime}, Start Ops ${cm}`;
}

function extractTime(text, regex) {
  const match = text.match(regex);
  if (match?.[1]) return toHHMM(match[1]);

  const fallback = text.match(
    new RegExp(`${regex.source.split('\\\d{1,2}:\\d{2}')[1]}[^\\n]*(\\d{1,2}:\\d{2})`, 'i')
  );
  if (fallback?.[1]) return toHHMM(fallback[1]);
  return '';
}

function toHHMM(raw) {
  const [h, m] = String(raw).split(':').map((v) => Number(v));
  if (Number.isNaN(h) || Number.isNaN(m)) return '';
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function parseMinutes(hhmm) {
  const [h, m] = hhmm.split(':').map(Number);
  return h * 60 + m;
}

function addRandomMinutes(hhmm, min, max) {
  const base = parseMinutes(hhmm);
  const randomIncrement = Math.floor(Math.random() * (max - min + 1)) + min;
  const total = (base + randomIncrement) % (24 * 60);
  return `${String(Math.floor(total / 60)).padStart(2, '0')}:${String(total % 60).padStart(2, '0')}`;
}

function detectMealBreaks(gangWayTime, commenceTime) {
  if (!gangWayTime || !commenceTime) return [];
  const start = parseMinutes(gangWayTime);
  let end = parseMinutes(commenceTime);
  if (end < start) end += 24 * 60;

  const windows = [
    { from: 11 * 60 + 30, to: 13 * 60, label: 'Meal Break S1' },
    { from: 17 * 60 + 30, to: 19 * 60, label: 'Meal Break S2' },
    { from: 3 * 60 + 30, to: 5 * 60, label: 'Meal Break S3' }
  ];

  return windows
    .filter((w) => {
      const a = w.from < start ? w.from + 24 * 60 : w.from;
      const b = w.to < start ? w.to + 24 * 60 : w.to;
      return Math.max(start, a) <= Math.min(end, b);
    })
    .map((w) => w.label);
}

function exportFilteredData() {
  const exportData = resultRows.map((row) => ({
    ...Object.fromEntries(DISPLAY_COLUMNS.map((c) => [c, row[c] || ''])),
    Remark: row.Remark || ''
  }));

  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Filtered Data');
  XLSX.writeFile(wb, 'maersk_weekly_filtered.xlsx');
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
  const imageData = ctx.getImageData(0, 0, berthCanvas.width, berthCanvas.height);
  const data = imageData.data;

  const target = { r: 145, g: 230, b: 245 };
  const tolerance = 80;
  const leftKeepWidth = Math.max(Math.round(berthCanvas.width * 0.12), 110);

  for (let y = 0; y < berthCanvas.height; y++) {
    for (let x = leftKeepWidth; x < berthCanvas.width; x++) {
      const i = (y * berthCanvas.width + x) * 4;
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      const a = data[i + 3];
      if (a === 0) continue;

      const dist = Math.sqrt((r - target.r) ** 2 + (g - target.g) ** 2 + (b - target.b) ** 2);
      if (dist > tolerance && isLikelyColored(r, g, b)) {
        const gray = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
        data[i] = Math.min(255, gray + 30);
        data[i + 1] = Math.min(255, gray + 30);
        data[i + 2] = Math.min(255, gray + 30);
        data[i + 3] = 160;
      }
    }
  }

  ctx.putImageData(imageData, 0, 0);
  downloadImageBtn.disabled = false;
}

function isLikelyColored(r, g, b) {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  return max - min > 20;
}

function downloadCanvasImage() {
  const link = document.createElement('a');
  link.download = 'berth_window_filtered.png';
  link.href = berthCanvas.toDataURL('image/png');
  link.click();
}
