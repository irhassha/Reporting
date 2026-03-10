const excelInput = document.getElementById('excelInput');
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

function createRowFromSource(sourceRow) {
  const row = {};
  DISPLAY_COLUMNS.forEach((col) => {
    row[col] = toDisplayText(findValueByAlias(sourceRow, columnAliasMap[col] || [col]));
  });
  row.remarkBlocks = [];
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
    remarkTd.appendChild(buildRemarkCell(rowIndex));
    tr.appendChild(remarkTd);
    tableBody.appendChild(tr);
  });
}

function buildRemarkCell(rowIndex) {
  const row = resultRows[rowIndex];
  const wrapper = document.createElement('div');
  wrapper.className = 'remark-cell';

  const list = document.createElement('div');
  list.className = 'remark-list';

  row.remarkBlocks.forEach((block, blockIndex) => {
    if (block.kind === 'log') {
      const logBlock = document.createElement('div');
      logBlock.className = 'remark-block glass-subpanel';
      const output = document.createElement('div');
      output.className = 'remark-output';
      output.textContent = block.text;
      const removeBtn = document.createElement('button');
      removeBtn.className = 'mini-btn';
      removeBtn.type = 'button';
      removeBtn.textContent = '✕';
      removeBtn.title = 'Remove remark';
      removeBtn.addEventListener('click', () => {
        row.remarkBlocks.splice(blockIndex, 1);
        renderTable();
      });
      logBlock.appendChild(removeBtn);
      logBlock.appendChild(output);
      output.style.gridColumn = '1 / -1';
      list.appendChild(logBlock);
      return;
    }

    const blockNode = remarkBlockTemplate.content.cloneNode(true);
    const select = blockNode.querySelector('.remark-select');
    const extraInput = blockNode.querySelector('.remark-extra');
    const removeBtn = blockNode.querySelector('.remove-remark-btn');
    const output = blockNode.querySelector('.remark-output');

    select.value = block.value || '';
    output.textContent = formatPresetBlock(block);

    const dynamicMap = {
      'QC Trouble Spreader': 'QC [801-808] Trouble Spreader',
      'DOOG Units': 'DOOG [number] Units',
      'LOOG Units': 'LOOG [number] Units',
      'D/L OOG Units': 'D/L OOG [number] Units'
    };

    if (dynamicMap[select.value]) {
      extraInput.classList.remove('hidden');
      extraInput.placeholder = dynamicMap[select.value];
      extraInput.value = block.extra || '';
    }

    select.addEventListener('change', () => {
      block.value = select.value;
      if (dynamicMap[select.value]) {
        extraInput.classList.remove('hidden');
        extraInput.placeholder = dynamicMap[select.value];
      } else {
        extraInput.classList.add('hidden');
        extraInput.value = '';
        block.extra = '';
      }
      output.textContent = formatPresetBlock(block);
    });

    extraInput.addEventListener('input', () => {
      block.extra = extraInput.value;
      output.textContent = formatPresetBlock(block);
    });

    removeBtn.addEventListener('click', () => {
      row.remarkBlocks.splice(blockIndex, 1);
      renderTable();
    });

    list.appendChild(blockNode);
  });

  const actions = document.createElement('div');
  actions.className = 'remark-actions';

  const addRemarkBtn = document.createElement('button');
  addRemarkBtn.type = 'button';
  addRemarkBtn.textContent = '(+) Add Remark';
  addRemarkBtn.className = 'mini-btn';
  addRemarkBtn.addEventListener('click', () => {
    row.remarkBlocks.push({ kind: 'preset', value: '', extra: '' });
    renderTable();
  });

  const pasteLogBtn = document.createElement('button');
  pasteLogBtn.type = 'button';
  pasteLogBtn.className = 'icon-btn';
  pasteLogBtn.title = 'Paste log';
  pasteLogBtn.textContent = '📋';
  pasteLogBtn.addEventListener('click', () => {
    activeRemarkRowIndex = rowIndex;
    logTextarea.value = '';
    modalBackdrop.classList.remove('hidden');
    modalBackdrop.setAttribute('aria-hidden', 'false');
  });

  actions.appendChild(addRemarkBtn);
  actions.appendChild(pasteLogBtn);

  wrapper.appendChild(list);
  wrapper.appendChild(actions);
  return wrapper;
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

function closeModal() {
  modalBackdrop.classList.add('hidden');
  modalBackdrop.setAttribute('aria-hidden', 'true');
}

function applyLogParse() {
  if (activeRemarkRowIndex === null) return;
  const parsed = parseOperationalLog(logTextarea.value || '');
  resultRows[activeRemarkRowIndex].remarkBlocks.push({ kind: 'log', text: parsed });
  closeModal();
  renderTable();
}

function parseOperationalLog(text) {
  const gangWayTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Gang\s*Way\s*Down/i, /Gang\s*Way\s*Down[^\n]*(\d{1,2}:\d{2})/i);
  const quarantineTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Quarantine\s*Clearance/i, /Quarantine\s*Clearance[^\n]*(\d{1,2}:\d{2})/i);
  const commenceTime = extractTime(text, /(?:^|\n|\b)(\d{1,2}:\d{2})[^\n]*Commence(?:\s+Loading|\s+Discharge|\b)/i, /Commence(?:\s+Loading|\s+Discharge|\b)[^\n]*(\d{1,2}:\d{2})/i);

  const gw = gangWayTime || 'N/A';
  const qt = quarantineTime || 'N/A';
  const cm = commenceTime || 'N/A';
  const tkbmTime = quarantineTime ? addRandomMinutes(quarantineTime, 3, 7) : 'N/A';
  const breaks = detectMealBreaks(gangWayTime, commenceTime);
  const breakText = breaks.length ? `${breaks.join(', ')}, ` : '';

  return `Gang way ready ${gw}, ${breakText}Quarantine ${qt}, TKBM ${tkbmTime}, Start Ops ${cm}`;
}

function extractTime(text, primaryRegex, fallbackRegex) {
  const primary = text.match(primaryRegex);
  if (primary?.[1]) return toHHMM(primary[1]);
  const fallback = text.match(fallbackRegex);
  if (fallback?.[1]) return toHHMM(fallback[1]);
  return '';
}

function toHHMM(raw) {
  const [h, m] = String(raw).split(':').map(Number);
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
  const exportData = resultRows.map((row) => {
    const joinedRemarks = row.remarkBlocks
      .map((block) => (block.kind === 'log' ? block.text : formatPresetBlock(block)))
      .filter(Boolean)
      .join('\n');

    return {
      ...Object.fromEntries(DISPLAY_COLUMNS.map((c) => [c, row[c] || ''])),
      Remark: joinedRemarks
    };
  });

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

  const width = berthCanvas.width;
  const height = berthCanvas.height;
  const leftKeepWidth = Math.max(Math.round(width * 0.12), 110);

  const imageData = ctx.getImageData(0, 0, width, height);
  const data = imageData.data;
  const target = { r: 145, g: 230, b: 245 };
  const tolerance = 78;
  const mask = new Uint8Array(width * height);

  for (let y = 0; y < height; y++) {
    for (let x = leftKeepWidth; x < width; x++) {
      const idx = y * width + x;
      const i = idx * 4;
      const r = data[i];
      const g = data[i + 1];
      const b = data[i + 2];
      const a = data[i + 3];
      if (a === 0) continue;

      const dist = Math.sqrt((r - target.r) ** 2 + (g - target.g) ** 2 + (b - target.b) ** 2);
      if (dist > tolerance && isLikelyColored(r, g, b)) mask[idx] = 1;
    }
  }

  const visited = new Uint8Array(width * height);
  const components = [];

  for (let y = 0; y < height; y++) {
    for (let x = leftKeepWidth; x < width; x++) {
      const seed = y * width + x;
      if (!mask[seed] || visited[seed]) continue;

      const queue = [seed];
      visited[seed] = 1;
      let head = 0;
      let minX = x;
      let maxX = x;
      let minY = y;
      let maxY = y;
      let count = 0;

      while (head < queue.length) {
        const current = queue[head++];
        const cx = current % width;
        const cy = Math.floor(current / width);
        count += 1;

        if (cx < minX) minX = cx;
        if (cx > maxX) maxX = cx;
        if (cy < minY) minY = cy;
        if (cy > maxY) maxY = cy;

        const neighbors = [current - 1, current + 1, current - width, current + width];
        neighbors.forEach((n) => {
          if (n < 0 || n >= width * height) return;
          const nx = n % width;
          const ny = Math.floor(n / width);
          if (nx < leftKeepWidth || ny < 0 || ny >= height) return;
          if (!mask[n] || visited[n]) return;
          visited[n] = 1;
          queue.push(n);
        });
      }

      const boxW = maxX - minX + 1;
      const boxH = maxY - minY + 1;
      if (count >= 60 && boxW >= 8 && boxH >= 8) {
        components.push({ x: minX, y: minY, w: boxW, h: boxH });
      }
    }
  }

  ctx.fillStyle = '#cccccc';
  components.forEach((box) => {
    const safeX = Math.max(box.x, leftKeepWidth);
    const safeW = box.w - (safeX - box.x);
    if (safeW > 0) ctx.fillRect(safeX, box.y, safeW, box.h);
  });

  downloadImageBtn.disabled = false;
}

function isLikelyColored(r, g, b) {
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  return max - min > 22;
}

function downloadCanvasImage() {
  const link = document.createElement('a');
  link.download = 'berth_window_filtered.png';
  link.href = berthCanvas.toDataURL('image/png');
  link.click();
}
