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

const resultRows = [];
let activeRemarkRowIndex = null;
let loadedImage = null;

const DISPLAY_COLUMNS = [
  'Vessel Name', 'Service', 'ATB', 'ATD', 'Total (Boxes)', 'Total Teus', 'Crane Intencity (CI)', 'Vessel Rate (VR)', 'ATB to Fst lift', 'Lst lift to ATD'
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

  resultRows.length = 0;
  rawRows.filter(operatorAllowed).forEach((src) => resultRows.push(createRowFromSource(src)));

  renderTable();
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
    remarkTd.className = 'px-2 py-2 align-top';
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
    Remark: getJoinedRemarks(row)
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
  const w = berthCanvas.width;
  const h = berthCanvas.height;
  const leftKeep = Math.max(Math.round(w * 0.12), 110);

  const image = ctx.getImageData(0, 0, w, h);
  const data = image.data;
  const target = { r: 145, g: 230, b: 245 };
  const tolerance = 78;
  const targetMask = new Uint8Array(w * h);
  const colorMask = new Uint8Array(w * h);

  for (let y = 0; y < h; y++) {
    for (let x = leftKeep; x < w; x++) {
      const idx = y * w + x;
      const i = idx * 4;
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      const a = data[i + 3];
      if (!a) continue;

      const dist = Math.hypot(r - target.r, g - target.g, b - target.b);
      const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      const max = Math.max(r, g, b);
      const min = Math.min(r, g, b);
      const sat = max === 0 ? 0 : (max - min) / max;

      if (dist <= tolerance) {
        targetMask[idx] = 1;
      }

      const isColoredBox = sat > 0.2 && lum > 30 && lum < 245 && dist > tolerance;
      const isWhiteBoxWithDarkBorder = lum > 170 && lum < 248;
      if (isColoredBox || isWhiteBoxWithDarkBorder) {
        colorMask[idx] = 1;
      }
    }
  }

  const visited = new Uint8Array(w * h);
  const dirs = [
    [1, 0],
    [-1, 0],
    [0, 1],
    [0, -1]
  ];

  for (let y = 0; y < h; y++) {
    for (let x = leftKeep; x < w; x++) {
      const seed = y * w + x;
      if (!colorMask[seed] || visited[seed] || targetMask[seed]) continue;

      const queue = [seed];
      visited[seed] = 1;
      let head = 0;
      let minX = x;
      let maxX = x;
      let minY = y;
      let maxY = y;
      let count = 0;
      let targetTouch = 0;

      while (head < queue.length) {
        const cur = queue[head++];
        const cx = cur % w;
        const cy = Math.floor(cur / w);
        count++;
        if (cx < minX) minX = cx;
        if (cx > maxX) maxX = cx;
        if (cy < minY) minY = cy;
        if (cy > maxY) maxY = cy;

        for (const [dx, dy] of dirs) {
          const nx = cx + dx;
          const ny = cy + dy;
          if (nx < leftKeep || nx >= w || ny < 0 || ny >= h) continue;
          const n = ny * w + nx;
          if (targetMask[n]) {
            targetTouch++;
            continue;
          }
          if (!colorMask[n] || visited[n]) continue;
          visited[n] = 1;
          queue.push(n);
        }
      }

      const boxW = maxX - minX + 1;
      const boxH = maxY - minY + 1;
      const area = boxW * boxH;
      const denseEnough = count / Math.max(area, 1) > 0.12;
      const validSize = boxW >= 8 && boxH >= 8 && count >= 40;
      if (!validSize || !denseEnough || targetTouch > count * 0.12) continue;

      const pad = 1;
      const sx = Math.max(leftKeep, minX - pad);
      const ex = Math.min(w - 1, maxX + pad);
      const sy = Math.max(0, minY - pad);
      const ey = Math.min(h - 1, maxY + pad);

      for (let yy = sy; yy <= ey; yy++) {
        for (let xx = sx; xx <= ex; xx++) {
          const idx = (yy * w + xx) * 4;
          data[idx] = 204;
          data[idx + 1] = 204;
          data[idx + 2] = 204;
          data[idx + 3] = 255;
        }
      }
    }
  }

  ctx.putImageData(image, 0, 0);
  downloadImageBtn.disabled = false;
}
function downloadCanvasImage() {
  const link = document.createElement('a');
  link.download = 'berth_window_filtered.png';
  link.href = berthCanvas.toDataURL('image/png');
  link.click();
}
