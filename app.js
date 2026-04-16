const state = {
  curveWorkbook: null,
  stockWorkbook: null,
  curveFileName: '',
  stockFileName: '',
  curveSheet: '',
  stockSheet: '',
  rows: [],
  filteredRows: []
};

const freqOrder = ['A+', 'A', 'B', 'C', 'D'];

const palette = {
  'A+': { bg: 'rgba(224,106,106,0.16)', border: 'rgba(224,106,106,0.36)', text: '#ffe0e0', fill: 'FCE1E1', row: 'F9ECEC' },
  'A':  { bg: 'rgba(255,157,77,0.16)', border: 'rgba(255,157,77,0.34)', text: '#ffe7d4', fill: 'FCE8D7', row: 'FBF1E8' },
  'B':  { bg: 'rgba(245,196,83,0.16)', border: 'rgba(245,196,83,0.34)', text: '#fff3c7', fill: 'FCF1D2', row: 'FBF7E7' },
  'C':  { bg: 'rgba(111,157,255,0.16)', border: 'rgba(111,157,255,0.34)', text: '#e0ebff', fill: 'DFE9FD', row: 'EEF4FD' },
  'D':  { bg: 'rgba(151,166,188,0.18)', border: 'rgba(151,166,188,0.28)', text: '#edf2fa', fill: 'E6EBF3', row: 'F4F7FB' },
  default: { bg: 'rgba(255,255,255,0.05)', border: 'rgba(255,255,255,0.10)', text: '#edf2fa', fill: 'EEF2F7', row: 'F8FAFC' }
};

const els = {
  launchScreen: document.getElementById('launchScreen'),
  appShell: document.getElementById('appShell'),
  curveInput: document.getElementById('curveInput'),
  stockInput: document.getElementById('stockInput'),
  curveCard: document.getElementById('curveCard'),
  stockCard: document.getElementById('stockCard'),
  curveTag: document.getElementById('curveTag'),
  stockTag: document.getElementById('stockTag'),
  curveFileNote: document.getElementById('curveFileNote'),
  stockFileNote: document.getElementById('stockFileNote'),
  openPanelBtn: document.getElementById('openPanelBtn'),
  curveChipName: document.getElementById('curveChipName'),
  stockChipName: document.getElementById('stockChipName'),
  replaceCurveBtn: document.getElementById('replaceCurveBtn'),
  replaceStockBtn: document.getElementById('replaceStockBtn'),
  backToStartBtn: document.getElementById('backToStartBtn'),
  curveSheetSelect: document.getElementById('curveSheetSelect'),
  stockSheetSelect: document.getElementById('stockSheetSelect'),
  metricSelect: document.getElementById('metricSelect'),
  thresholdA: document.getElementById('thresholdA'),
  thresholdB: document.getElementById('thresholdB'),
  processBtn: document.getElementById('processBtn'),
  statusLine: document.getElementById('statusLine'),
  processInfo: document.getElementById('processInfo'),
  kpiMatched: document.getElementById('kpiMatched'),
  kpiMatchedSub: document.getElementById('kpiMatchedSub'),
  kpiPriority: document.getElementById('kpiPriority'),
  kpiPrioritySub: document.getElementById('kpiPrioritySub'),
  kpiCritical: document.getElementById('kpiCritical'),
  kpiCriticalSub: document.getElementById('kpiCriticalSub'),
  kpiZero: document.getElementById('kpiZero'),
  kpiZeroSub: document.getElementById('kpiZeroSub'),
  distributionGrid: document.getElementById('distributionGrid'),
  distributionSub: document.getElementById('distributionSub'),
  searchInput: document.getElementById('searchInput'),
  frequencyFilter: document.getElementById('frequencyFilter'),
  syncFilter: document.getElementById('syncFilter'),
  resultCount: document.getElementById('resultCount'),
  tableWrap: document.getElementById('tableWrap'),
  exportXlsxBtn: document.getElementById('exportXlsxBtn'),
  exportPdfBtn: document.getElementById('exportPdfBtn'),
  exportCsvBtn: document.getElementById('exportCsvBtn')
};

function normalizeHeader(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeSearch(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function parseNumber(value) {
  if (value == null || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  let text = String(value).trim();
  if (!text) return 0;
  text = text.replace(/\s+/g, '');
  text = text.replace(/R\$/gi, '').replace(/%/g, '');

  const hasComma = text.includes(',');
  const hasDot = text.includes('.');
  if (hasComma && hasDot) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (hasComma) {
    text = text.replace(/\./g, '').replace(',', '.');
  }
  const num = Number(text.replace(/[^0-9.-]/g, ''));
  return Number.isFinite(num) ? num : 0;
}

function formatInt(value) {
  return new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 }).format(Number(value || 0));
}

function formatDecimal(value, digits = 2) {
  return new Intl.NumberFormat('pt-BR', {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  }).format(Number(value || 0));
}

function formatCurrency(value) {
  return new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  }).format(Number(value || 0));
}

function formatPercent(value) {
  return `${formatDecimal(value, 2)}%`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function getWorkbookRows(workbook, sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
}

function rowsToObjects(matrix) {
  if (!matrix || !matrix.length) return [];
  const headers = matrix[0].map((cell, index) => String(cell || `col_${index + 1}`).trim() || `col_${index + 1}`);
  return matrix.slice(1).map((row) => Object.fromEntries(
    headers.map((header, index) => [header, row[index] ?? ''])
  ));
}

function detectCurveColumns(headers) {
  const normalized = headers.map((header) => ({ header, norm: normalizeHeader(header) }));
  const pick = (tests) => normalized.find(({ norm }) => tests.some((test) => typeof test === 'string' ? norm.includes(test) : test(norm)))?.header || '';
  return {
    product: pick(['produto', 'descricao']),
    sku: pick([norm => norm === 'codigo', norm => norm === 'codigo sku', norm => norm.includes('sku'), norm => norm.includes('codigo')]),
    quantity: pick([norm => norm.includes('quantidade'), norm => norm === 'qtd']),
    value: pick([norm => norm === 'valor', norm => norm.includes('valor total'), norm => norm.includes('faturamento')])
  };
}

function detectStockColumns(headers) {
  const normalized = headers.map((header) => ({ header, norm: normalizeHeader(header) }));
  const pick = (tests) => normalized.find(({ norm }) => tests.some((test) => typeof test === 'string' ? norm.includes(test) : test(norm)))?.header || '';
  const product = pick(['produto', 'descricao']);
  const sku = pick([norm => norm.includes('codigo sku'), norm => norm === 'sku', norm => norm.includes('sku'), norm => norm.includes('codigo')]);
  const total = pick([norm => norm === 'total', norm => norm.includes('estoque total'), norm => norm.includes('saldo total')]);
  const excluded = new Set([normalizeHeader(product), normalizeHeader(sku), 'unidade', 'empresa', 'filial']);
  const numericHeaders = headers.filter((header) => {
    const norm = normalizeHeader(header);
    return norm && !excluded.has(norm);
  });
  return { product, sku, total, numericHeaders };
}

function normalizeSku(code) {
  if (code == null || code === '') return '';
  let sku = String(code).trim().toUpperCase().replace(/\s+/g, '');
  if (!sku) return '';

  if (sku.startsWith('ACBF000-')) {
    return sku.replace(/-K\d{3}[A-Z]*S?$/i, '-K050');
  }

  sku = sku.replace(/(-K\d{3}[A-Z]*)S$/i, '$1');
  const match = sku.match(/^(.*-K)\d{3}([A-Z]*)$/i);
  if (match) {
    sku = `${match[1]}001${match[2] || ''}`.toUpperCase();
  }
  return sku;
}

function isBundleName(name) {
  const text = normalizeSearch(name).trim();
  return /^(kit\b|combo\b|pack\b|\d+\s*un\b)/.test(text);
}

function scoreSource(source, baseSku) {
  const rawSku = String(source.rawSku || '').trim().toUpperCase().replace(/\s+/g, '');
  const name = String(source.name || '').trim();
  let score = 0;

  if (!isBundleName(name)) score += 100;
  if (rawSku === baseSku) score += 80;
  else if (rawSku.startsWith(baseSku)) score += 60;
  else if (normalizeSku(rawSku) === baseSku) score += 40;

  const lengthBonus = Math.max(0, 20 - Math.min(name.length, 60) / 3);
  score += lengthBonus;

  return score;
}

function chooseCanonicalSource(sources, baseSku) {
  if (!sources.length) return { name: baseSku, description: baseSku };
  return [...sources].sort((a, b) => {
    const scoreDiff = scoreSource(b, baseSku) - scoreSource(a, baseSku);
    if (scoreDiff !== 0) return scoreDiff;
    return String(a.name || '').length - String(b.name || '').length;
  })[0];
}

function buildCurveRows(objects) {
  if (!objects.length) return [];
  const headers = Object.keys(objects[0]);
  const columns = detectCurveColumns(headers);
  const map = new Map();

  for (const row of objects) {
    const rawSku = row[columns.sku];
    const baseSku = normalizeSku(rawSku);
    if (!baseSku) continue;

    const product = String(row[columns.product] ?? '').trim() || baseSku;
    const quantity = parseNumber(row[columns.quantity]);
    const value = parseNumber(row[columns.value]);

    const current = map.get(baseSku) || {
      skuBase: baseSku,
      soldQty: 0,
      soldValue: 0,
      sources: []
    };

    current.soldQty += quantity;
    current.soldValue += value;
    current.sources.push({
      rawSku: rawSku,
      name: product,
      description: product,
      sourceType: 'curve'
    });

    map.set(baseSku, current);
  }

  return [...map.values()].map((item) => {
    const canonical = chooseCanonicalSource(item.sources, item.skuBase);
    return {
      ...item,
      product: canonical.name || item.skuBase,
      description: canonical.description || canonical.name || item.skuBase
    };
  });
}

function buildStockRows(objects) {
  if (!objects.length) return [];
  const headers = Object.keys(objects[0]);
  const columns = detectStockColumns(headers);
  const map = new Map();

  for (const row of objects) {
    const rawSku = row[columns.sku];
    const baseSku = normalizeSku(rawSku);
    if (!baseSku) continue;

    const stockTotal = columns.total
      ? parseNumber(row[columns.total])
      : columns.numericHeaders.reduce((sum, header) => sum + parseNumber(row[header]), 0);
    const product = String(row[columns.product] ?? '').trim() || baseSku;

    const current = map.get(baseSku) || {
      skuBase: baseSku,
      stockTotal: 0,
      sources: []
    };

    current.stockTotal += stockTotal;
    current.sources.push({
      rawSku: rawSku,
      name: product,
      description: product,
      sourceType: 'stock'
    });

    map.set(baseSku, current);
  }

  return [...map.values()].map((item) => {
    const canonical = chooseCanonicalSource(item.sources, item.skuBase);
    return {
      ...item,
      stockProduct: canonical.name || item.skuBase,
      stockDescription: canonical.description || canonical.name || item.skuBase
    };
  });
}

function classifyABC(cumulative, thresholdA, thresholdB) {
  if (cumulative <= thresholdA) return 'A';
  if (cumulative <= thresholdB) return 'B';
  return 'C';
}

function addCurveClasses(rows, metricKey, thresholdA, thresholdB) {
  const sorted = [...rows].sort((a, b) => (b[metricKey] || 0) - (a[metricKey] || 0));
  const total = sorted.reduce((sum, row) => sum + (row[metricKey] || 0), 0);
  let cumulative = 0;

  return sorted.map((row) => {
    const metric = row[metricKey] || 0;
    const individualPct = total > 0 ? (metric / total) * 100 : 0;
    cumulative += individualPct;
    return {
      ...row,
      individualPct,
      cumulativePct: cumulative,
      abcClass: classifyABC(cumulative, thresholdA, thresholdB)
    };
  });
}

function getFrequency(row) {
  const salesWeight = row.soldQty <= 0 && row.soldValue <= 0
    ? 0
    : row.abcClass === 'A' ? 5
    : row.abcClass === 'B' ? 3
    : 1.5;

  const stock = row.stockTotal || 0;
  const soldQty = row.soldQty || 0;
  const coverage = soldQty > 0 ? stock / soldQty : null;

  let coverageWeight = 0;
  if (stock <= 0 && salesWeight > 0) coverageWeight = 3.5;
  else if (soldQty <= 0) coverageWeight = stock > 0 ? -1.5 : 0;
  else if (coverage <= 0.5) coverageWeight = 3;
  else if (coverage <= 1) coverageWeight = 2.4;
  else if (coverage <= 2) coverageWeight = 1.4;
  else if (coverage <= 4) coverageWeight = 0.3;
  else if (coverage <= 8) coverageWeight = -0.7;
  else coverageWeight = -1.7;

  let absolutePenalty = 0;
  if (stock >= 500) absolutePenalty = 1.5;
  else if (stock >= 200) absolutePenalty = 1;
  else if (stock >= 80) absolutePenalty = 0.5;

  const score = salesWeight + coverageWeight - absolutePenalty;
  let frequency = 'D';

  if (stock <= 0 && row.abcClass === 'A') frequency = 'A+';
  else if (score >= 7) frequency = 'A+';
  else if (score >= 5) frequency = 'A';
  else if (score >= 3.2) frequency = 'B';
  else if (score >= 1.4) frequency = 'C';

  let syncStatus = 'matched';
  if (!row.hasStockMatch) syncStatus = 'missing';
  else if ((row.stockTotal || 0) <= 0) syncStatus = 'zero';

  return {
    ...row,
    coverage,
    score,
    frequency,
    syncStatus
  };
}

function mergeRows(curveRows, stockRows, metricKey, thresholdA, thresholdB) {
  const curveWithClass = addCurveClasses(curveRows, metricKey, thresholdA, thresholdB);
  const curveMap = new Map(curveWithClass.map((item) => [item.skuBase, item]));
  const stockMap = new Map(stockRows.map((item) => [item.skuBase, item]));
  const baseSkus = [...new Set([...curveMap.keys(), ...stockMap.keys()])];

  const merged = baseSkus.map((skuBase) => {
    const curve = curveMap.get(skuBase);
    const stock = stockMap.get(skuBase);

    const allSources = [
      ...(curve?.sources || []),
      ...(stock?.sources || [])
    ];
    const canonical = chooseCanonicalSource(allSources, skuBase);
    const product = canonical.name || curve?.product || stock?.stockProduct || skuBase;
    const description = canonical.description || product;

    const row = {
      skuBase,
      product,
      description,
      soldQty: curve?.soldQty || 0,
      soldValue: curve?.soldValue || 0,
      individualPct: curve?.individualPct || 0,
      cumulativePct: curve?.cumulativePct || 0,
      abcClass: curve?.abcClass || '—',
      stockTotal: stock?.stockTotal ?? 0,
      hasStockMatch: Boolean(stock)
    };

    return getFrequency(row);
  });

  return merged.sort((a, b) => {
    const orderDiff = freqOrder.indexOf(a.frequency) - freqOrder.indexOf(b.frequency);
    if (orderDiff !== 0) return orderDiff;
    return (b.soldQty || 0) - (a.soldQty || 0);
  });
}

function getFreqPalette(freq) {
  return palette[freq] || palette.default;
}

function renderDistribution(rows) {
  const counts = Object.fromEntries(freqOrder.map((freq) => [freq, 0]));
  rows.forEach((row) => {
    counts[row.frequency] = (counts[row.frequency] || 0) + 1;
  });

  const labels = {
    'A+': 'ruptura',
    'A': 'alto giro',
    'B': 'recorrente',
    'C': 'periódico',
    'D': 'baixa'
  };

  const rules = {
    'A+': 'sem saldo ou muito curto',
    'A': 'contagem mais próxima',
    'B': 'rotina recorrente',
    'C': 'cadência periódica',
    'D': 'baixa prioridade'
  };

  els.distributionGrid.innerHTML = freqOrder.map((freq) => {
    const colors = getFreqPalette(freq);
    return `
      <div class="distribution-item">
        <div class="distribution-top">
          <div class="distribution-label">
            <span class="distribution-dot" style="background:${colors.bg}; border:1px solid ${colors.border};"></span>
            <span>${freq}</span>
          </div>
          <span class="section-subtle">${labels[freq]}</span>
        </div>
        <div class="distribution-count">${formatInt(counts[freq] || 0)}</div>
        <div class="distribution-rule">${rules[freq]}</div>
      </div>
    `;
  }).join('');

  const high = (counts['A+'] || 0) + (counts['A'] || 0);
  els.distributionSub.textContent = `${formatInt(high)} itens em alta frequência`;
}

function renderSummary(rows) {
  const matched = rows.filter((row) => row.hasStockMatch).length;
  const highPriority = rows.filter((row) => row.frequency === 'A+' || row.frequency === 'A').length;
  const critical = rows.filter((row) => row.stockTotal <= 0 || (row.coverage != null && row.coverage <= 1)).length;
  const zero = rows.filter((row) => row.stockTotal <= 0).length;
  const metricLabel = els.metricSelect.value === 'value' ? 'valor vendido' : 'quantidade vendida';

  els.kpiMatched.textContent = formatInt(matched);
  els.kpiMatchedSub.textContent = `${formatInt(rows.length)} SKU base no painel`;
  els.kpiPriority.textContent = formatInt(highPriority);
  els.kpiPrioritySub.textContent = 'Frequência mais alta';
  els.kpiCritical.textContent = formatInt(critical);
  els.kpiCriticalSub.textContent = 'Sem saldo ou cobertura até 1x';
  els.kpiZero.textContent = formatInt(zero);
  els.kpiZeroSub.textContent = 'Estoque igual a zero';
  els.statusLine.textContent = `${metricLabel} + estoque total cruzados por SKU base.`;
  els.processInfo.textContent = `Base ABC por ${metricLabel}. ACBF000 → K050.`;

  renderDistribution(rows);
}

function getActiveRows() {
  const query = normalizeSearch(els.searchInput.value);
  const frequency = els.frequencyFilter.value;
  const sync = els.syncFilter.value;

  return state.rows.filter((row) => {
    const haystack = normalizeSearch(`${row.product} ${row.description} ${row.skuBase}`);
    if (query && !haystack.includes(query)) return false;
    if (frequency && row.frequency !== frequency) return false;
    if (sync && row.syncStatus !== sync) return false;
    return true;
  });
}

function renderTable() {
  const rows = getActiveRows();
  state.filteredRows = rows;
  els.resultCount.textContent = `${formatInt(rows.length)} itens`;

  if (!rows.length) {
    els.tableWrap.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true">
            <circle cx="12" cy="12" r="7.25"></circle>
            <path d="M9.4 9.4 14.6 14.6"></path>
            <path d="M14.6 9.4 9.4 14.6"></path>
          </svg>
        </div>
        <strong>Nenhum item encontrado</strong>
        <span>Ajuste filtros ou reimporte os arquivos.</span>
      </div>
    `;
    return;
  }

  const body = rows.map((row) => {
    const colors = getFreqPalette(row.frequency);
    const syncLabel = row.syncStatus === 'matched'
      ? 'Com estoque'
      : row.syncStatus === 'zero'
        ? 'Saldo zerado'
        : 'Sem estoque';
    const subtitle = row.description && row.description !== row.product ? row.description : '';
    const coverageText = row.coverage == null ? '—' : `${formatDecimal(row.coverage, 2)}x`;

    return `
      <tr>
        <td class="product-cell">
          <div class="product-stack">
            <div class="product-title">${escapeHtml(row.product)}</div>
            ${subtitle ? `<div class="product-sub">${escapeHtml(subtitle)}</div>` : ''}
          </div>
        </td>
        <td><span class="mono">${escapeHtml(row.skuBase)}</span></td>
        <td><span class="pill" style="background:${colors.bg};border-color:${colors.border};color:${colors.text};">${escapeHtml(row.frequency)}</span></td>
        <td><span class="pill pill-neutral">${escapeHtml(row.abcClass)}</span></td>
        <td>${formatInt(row.soldQty)}</td>
        <td>${formatInt(row.stockTotal)}</td>
        <td>${coverageText}</td>
        <td><span class="pill pill-neutral">${syncLabel}</span></td>
      </tr>
    `;
  }).join('');

  els.tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Produto</th>
          <th>SKU base</th>
          <th>Frequência</th>
          <th>Curva</th>
          <th>Qtd. vendida</th>
          <th>Estoque total</th>
          <th>Cobertura</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>${body}</tbody>
    </table>
  `;
}

function updateFileState(kind, fileName) {
  const isCurve = kind === 'curve';
  const card = isCurve ? els.curveCard : els.stockCard;
  const tag = isCurve ? els.curveTag : els.stockTag;
  const note = isCurve ? els.curveFileNote : els.stockFileNote;
  const chip = isCurve ? els.curveChipName : els.stockChipName;

  card.classList.add('loaded');
  tag.textContent = 'pronto';
  tag.classList.add('ok');
  note.textContent = fileName;
  chip.textContent = fileName;
}

function populateSheetSelect(select, workbook, preferredName) {
  const options = workbook.SheetNames.map((sheetName) =>
    `<option value="${escapeHtml(sheetName)}">${escapeHtml(sheetName)}</option>`
  ).join('');
  select.innerHTML = options;
  select.value = preferredName || workbook.SheetNames[0] || '';
  select.disabled = workbook.SheetNames.length <= 1;
}

async function handleImport(kind, file) {
  try {
    if (!file) return;
    
    // Validação para evitar falha silenciosa se a biblioteca não carregar
    if (typeof XLSX === 'undefined') {
      throw new Error("Biblioteca SheetJS (XLSX) não encontrada. Verifique as tags <script> no seu HTML.");
    }

    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });

    if (kind === 'curve') {
      state.curveWorkbook = workbook;
      state.curveFileName = file.name;
      state.curveSheet = workbook.SheetNames[0] || '';
      populateSheetSelect(els.curveSheetSelect, workbook, state.curveSheet);
    } else {
      state.stockWorkbook = workbook;
      state.stockFileName = file.name;
      state.stockSheet = workbook.SheetNames[0] || '';
      populateSheetSelect(els.stockSheetSelect, workbook, state.stockSheet);
    }

    updateFileState(kind, file.name);

    if (state.curveWorkbook && state.stockWorkbook) {
      els.openPanelBtn.disabled = false;
      if (!els.appShell.classList.contains('hidden')) processData();
    }
  } catch (error) {
    console.error(`Erro ao importar a planilha (${kind}):`, error);
    alert(`Erro ao importar arquivo: ${error.message}`);
  }
}

function processData() {
  try {
    if (!state.curveWorkbook || !state.stockWorkbook) return;

    const thresholdA = Math.min(100, Math.max(1, parseNumber(els.thresholdA.value) || 80));
    const thresholdB = Math.min(100, Math.max(thresholdA, parseNumber(els.thresholdB.value) || 95));
    els.thresholdA.value = thresholdA;
    els.thresholdB.value = thresholdB;

    state.curveSheet = els.curveSheetSelect.value || state.curveSheet;
    state.stockSheet = els.stockSheetSelect.value || state.stockSheet;

    const curveObjects = rowsToObjects(getWorkbookRows(state.curveWorkbook, state.curveSheet));
    const stockObjects = rowsToObjects(getWorkbookRows(state.stockWorkbook, state.stockSheet));
    const curveRows = buildCurveRows(curveObjects);
    const stockRows = buildStockRows(stockObjects);

    const metricKey = els.metricSelect.value === 'value' ? 'soldValue' : 'soldQty';
    state.rows = mergeRows(curveRows, stockRows, metricKey, thresholdA, thresholdB);

    renderSummary(state.rows);
    renderTable();

    const enabled = state.rows.length > 0;
    els.exportXlsxBtn.disabled = !enabled;
    els.exportPdfBtn.disabled = !enabled;
    els.exportCsvBtn.disabled = !enabled;
  } catch (error) {
    console.error("Erro ao processar os dados:", error);
    alert("Ocorreu um erro ao processar os dados das planilhas. Verifique o console (F12).");
  }
}

function openPanel() {
  els.launchScreen.classList.add('hidden');
  els.appShell.classList.remove('hidden');
  processData();
}

function backToStart() {
  els.appShell.classList.add('hidden');
  els.launchScreen.classList.remove('hidden');
}

function buildExportName(ext) {
  const now = new Date();
  const pad = (num) => String(num).padStart(2, '0');
  return `frequencia_contagem_${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}.${ext}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function getExportRows() {
  return state.filteredRows?.length ? state.filteredRows : state.rows;
}

function applyExcelRowStyle(row, freq) {
  const colors = getFreqPalette(freq);
  row.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.row } };
    cell.font = { color: { argb: '1F2937' }, name: 'Arial', size: 11 };
    cell.border = {
      top: { style: 'thin', color: { argb: 'E5E7EB' } },
      bottom: { style: 'thin', color: { argb: 'E5E7EB' } }
    };
  });
}

async function exportXlsx() {
  const rows = getExportRows();
  if (!rows.length) return;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ChatGPT';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Frequência de contagem');
  sheet.columns = [
    { header: 'Produto', key: 'product', width: 34 },
    { header: 'SKU BASE', key: 'skuBase', width: 20 },
    { header: 'Frequência', key: 'frequency', width: 12 },
    { header: 'Curva', key: 'abcClass', width: 10 },
    { header: 'Qtd. vendida', key: 'soldQty', width: 14 },
    { header: 'Valor vendido', key: 'soldValue', width: 16 },
    { header: 'Estoque total', key: 'stockTotal', width: 14 },
    { header: 'Cobertura', key: 'coverage', width: 12 },
    { header: 'Descrição', key: 'description', width: 42 },
    { header: '% Individual', key: 'individualPct', width: 14 },
    { header: '% Acumulado', key: 'cumulativePct', width: 14 },
    { header: 'Status', key: 'status', width: 14 }
  ];

  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.autoFilter = 'A1:L1';

  const header = sheet.getRow(1);
  header.font = { bold: true, color: { argb: 'F8FAFC' }, name: 'Arial', size: 11 };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F2937' } };
  header.eachCell((cell) => {
    cell.border = { bottom: { style: 'thin', color: { argb: '374151' } } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });

  rows.forEach((item) => {
    const row = sheet.addRow({
      product: item.product,
      skuBase: item.skuBase,
      frequency: item.frequency,
      abcClass: item.abcClass,
      soldQty: item.soldQty,
      soldValue: item.soldValue,
      stockTotal: item.stockTotal,
      coverage: item.coverage == null ? '' : item.coverage,
      description: item.description,
      individualPct: item.individualPct / 100,
      cumulativePct: item.cumulativePct / 100,
      status: item.syncStatus === 'matched' ? 'Com estoque' : item.syncStatus === 'zero' ? 'Saldo zerado' : 'Sem estoque'
    });

    row.getCell(2).font = { name: 'Consolas', size: 11, color: { argb: '111827' } };
    row.getCell(6).numFmt = 'R$ #,##0.00';
    row.getCell(8).numFmt = '0.00x';
    row.getCell(10).numFmt = '0.00%';
    row.getCell(11).numFmt = '0.00%';

    applyExcelRowStyle(row, item.frequency);

    const freqCell = row.getCell(3);
    const colors = getFreqPalette(item.frequency);
    freqCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors.fill } };
    freqCell.font = { bold: true, color: { argb: '111827' }, name: 'Arial', size: 11 };
  });

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(
    new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
    buildExportName('xlsx')
  );
}

function buildPdfRows(rows) {
  return rows.map((item) => ([
    item.product,
    item.skuBase,
    item.frequency,
    item.description,
    formatInt(item.stockTotal)
  ]));
}

function hexToRgb(hex) {
  const clean = hex.replace('#', '');
  return [0, 2, 4].map((index) => parseInt(clean.slice(index, index + 2), 16));
}

function exportPdf() {
  const rows = getExportRows();
  if (!rows.length) return;

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });

  doc.setFillColor(11, 17, 27);
  doc.rect(0, 0, doc.internal.pageSize.getWidth(), doc.internal.pageSize.getHeight(), 'F');
  doc.setTextColor(237, 242, 250);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(18);
  doc.text('Frequência de contagem', 36, 40);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.setTextColor(151, 166, 188);
  doc.text('Produto · SKU base · frequência · descrição · estoque total', 36, 58);

  doc.autoTable({
    startY: 74,
    head: [['Produto', 'SKU BASE', 'Frequência', 'Descrição', 'Estoque total']],
    body: buildPdfRows(rows),
    theme: 'grid',
    styles: {
      fontSize: 9,
      cellPadding: 7,
      textColor: [28, 37, 51],
      lineColor: [229, 231, 235],
      lineWidth: 0.5
    },
    headStyles: {
      fillColor: [31, 41, 55],
      textColor: [248, 250, 252],
      fontStyle: 'bold'
    },
    didParseCell(data) {
      if (data.section !== 'body') return;
      const freq = rows[data.row.index]?.frequency;
      const colors = getFreqPalette(freq);
      data.cell.styles.fillColor = hexToRgb(`#${colors.row}`);
      if (data.column.index === 2) {
        data.cell.styles.fillColor = hexToRgb(`#${colors.fill}`);
        data.cell.styles.fontStyle = 'bold';
      }
    }
  });

  doc.save(buildExportName('pdf'));
}

function exportCsv() {
  const rows = getExportRows();
  if (!rows.length) return;

  const header = [
    'Produto', 'SKU BASE', 'Frequência', 'Curva', 'Qtd. vendida',
    'Valor vendido', 'Estoque total', 'Cobertura', 'Descrição',
    '% Individual', '% Acumulado', 'Status'
  ];

  const body = rows.map((item) => [
    item.product,
    item.skuBase,
    item.frequency,
    item.abcClass,
    item.soldQty,
    item.soldValue,
    item.stockTotal,
    item.coverage == null ? '' : formatDecimal(item.coverage, 2),
    item.description,
    formatDecimal(item.individualPct, 2),
    formatDecimal(item.cumulativePct, 2),
    item.syncStatus === 'matched' ? 'Com estoque' : item.syncStatus === 'zero' ? 'Saldo zerado' : 'Sem estoque'
  ]);

  const csv = [header, ...body]
    .map((row) => row.map((value) => `"${String(value ?? '').replace(/"/g, '""')}"`).join(';'))
    .join('\n');

  downloadBlob(new Blob([csv], { type: 'text/csv;charset=utf-8;' }), buildExportName('csv'));
}

els.curveCard.addEventListener('click', () => els.curveInput.click());
els.stockCard.addEventListener('click', () => els.stockInput.click());
els.replaceCurveBtn.addEventListener('click', () => els.curveInput.click());
els.replaceStockBtn.addEventListener('click', () => els.stockInput.click());
els.curveInput.addEventListener('change', async (event) => {
  const file = event.target.files?.[0];
  if (file) await handleImport('curve', file);
  event.target.value = '';
});
els.stockInput.addEventListener('change', async (event) => {
  const file = event.target.files?.[0];
  if (file) await handleImport('stock', file);
  event.target.value = '';
});
els.openPanelBtn.addEventListener('click', openPanel);
els.backToStartBtn.addEventListener('click', backToStart);
els.processBtn.addEventListener('click', processData);
els.metricSelect.addEventListener('change', processData);
els.curveSheetSelect.addEventListener('change', processData);
els.stockSheetSelect.addEventListener('change', processData);
els.thresholdA.addEventListener('change', processData);
els.thresholdB.addEventListener('change', processData);
els.searchInput.addEventListener('input', renderTable);
els.frequencyFilter.addEventListener('change', renderTable);
els.syncFilter.addEventListener('change', renderTable);
els.exportXlsxBtn.addEventListener('click', exportXlsx);
els.exportPdfBtn.addEventListener('click', exportPdf);
els.exportCsvBtn.addEventListener('click', exportCsv);
