const state = {
  curveWorkbook: null,
  stockWorkbook: null,
  compositionWorkbooks: [],
  curveFileName: '',
  stockFileName: '',
  compositionFileNames: [],
  curveSheet: '',
  stockSheet: '',
  compositionSheets: [],
  compositionMap: null,
  compositionLoadPromise: null,
  compositionLoadError: null,
  warningsExpanded: false,
  warnings: [],
  rows: [],
  filteredRows: []
};

const COMPOSITION_FILE_NAME = 'composicao.xls';
const manualSkuCompositions = [
  {
    sourceSku: '179P',
    componentSku: 'CMPO00C-K001',
    componentQty: 1,
    componentDescription: 'Ponds Creme Facial C Limpeza Demaquilante 100g Importado'
  },
  {
    sourceSku: '3139P',
    componentSku: 'DDCM000-K001',
    componentQty: 6,
    componentDescription: 'Desodorante Canon Musk Argentino Importado Aerosol 250ml'
  },
  {
    sourceSku: '427P',
    componentSku: 'HBCO40M-K001',
    componentQty: 1,
    componentDescription: 'Ultra Corega Creme Menta Fixador de Dentaduras 40g'
  }
];
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
  lowStockDays: document.getElementById('lowStockDays'),
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
  warningsPanel: document.getElementById('warningsPanel'),
  warningsToggle: document.getElementById('warningsToggle'),
  warningsCount: document.getElementById('warningsCount'),
  warningsList: document.getElementById('warningsList'),
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
  const olistId = pick([norm => norm === 'id olist', norm => norm.includes('id olist'), norm => norm === 'id']);
  const total = pick([norm => norm === 'total', norm => norm.includes('estoque total'), norm => norm.includes('saldo total')]);
  const excluded = new Set([normalizeHeader(product), normalizeHeader(sku), normalizeHeader(olistId), 'unidade', 'empresa', 'filial']);
  const numericHeaders = headers.filter((header) => {
    const norm = normalizeHeader(header);
    return norm && !excluded.has(norm);
  });
  return { product, sku, olistId, total, numericHeaders };
}

function detectCompositionColumns(headers) {
  const normalized = headers.map((header) => ({ header, norm: normalizeHeader(header) }));
  const pick = (exact, tests = []) => {
    const exactMatch = normalized.find(({ norm }) => norm === exact);
    if (exactMatch) return exactMatch.header;
    return normalized.find(({ norm }) => tests.some((test) => typeof test === 'string' ? norm.includes(test) : test(norm)))?.header || '';
  };

  return {
    kitId: pick('id kit fabricado', [norm => norm.includes('id kit'), norm => norm.includes('id fabricado')]),
    kitSku: pick('sku kit fabricado', [norm => norm.includes('sku kit'), norm => norm.includes('sku fabricado')]),
    kitDescription: pick('descricao kit fabricado', [norm => norm.includes('descricao kit'), norm => norm.includes('descricao fabricado')]),
    componentId: pick('id componente', [norm => norm.includes('id componente')]),
    componentSku: pick('sku componente', [norm => norm.includes('sku componente')]),
    componentDescription: pick('descricao componente', [norm => norm.includes('descricao componente')]),
    componentQty: pick('quantidade componente', [norm => norm.includes('quantidade componente'), norm => norm === 'qtd componente'])
  };
}

function normalizeSkuBasic(code) {
  if (code == null || code === '') return '';
  return String(code).trim().toUpperCase().replace(/\s+/g, '');
}

function normalizeSkuKitLookup(code) {
  const sku = normalizeSkuBasic(code);
  if (!sku) return '';
  return sku.replace(/(-K\d{3}[A-Z]*)S$/i, '$1').toUpperCase();
}

function normalizeSkuComponent(code) {
  return normalizeSkuBasic(code);
}

function legacyNormalizeSku(code) {
  let sku = normalizeSkuBasic(code);
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

function normalizeSku(code) {
  return legacyNormalizeSku(code);
}

function createEmptyCompositionMap() {
  return {
    byKitSkuRaw: new Map(),
    byKitSkuNorm: new Map(),
    byKitSkuLegacy: new Map(),
    byKitId: new Map(),
    componentsByKey: new Map(),
    warnings: []
  };
}

function addCompositionLookup(map, component) {
  const rawKey = String(component.kitSkuRaw || '').trim();
  const basicKey = normalizeSkuBasic(component.kitSkuRaw);
  const lookupKey = normalizeSkuKitLookup(component.kitSkuRaw);
  const legacyKey = legacyNormalizeSku(component.kitSkuRaw);
  const keys = [
    [map.byKitSkuRaw, rawKey],
    [map.byKitSkuNorm, basicKey],
    [map.byKitSkuNorm, lookupKey],
    [map.byKitSkuLegacy, legacyKey]
  ];

  keys.forEach(([target, key]) => {
    if (!key) return;
    const list = target.get(key) || [];
    if (!list.includes(component)) list.push(component);
    target.set(key, list);
  });

  if (component.kitId) {
    const list = map.byKitId.get(component.kitId) || [];
    if (!list.includes(component)) list.push(component);
    map.byKitId.set(component.kitId, list);
  }
}

function buildCompositionMap(objects) {
  const map = createEmptyCompositionMap();
  if (!objects.length) return map;

  const headers = Object.keys(objects[0]);
  const columns = detectCompositionColumns(headers);

  for (const row of objects) {
    const kitSkuRaw = String(row[columns.kitSku] ?? '').trim();
    const componentSkuRaw = String(row[columns.componentSku] ?? '').trim();
    const kitId = String(row[columns.kitId] ?? '').trim();
    const componentId = String(row[columns.componentId] ?? '').trim();
    const componentQty = parseNumber(row[columns.componentQty]);

    if (!kitSkuRaw) continue;
    if (!componentSkuRaw) {
      map.warnings.push(`Composição ignorada para kit ${kitSkuRaw}: linha sem SKU de componente.`);
      continue;
    }
    if (!componentQty) {
      map.warnings.push(`Composição de ${kitSkuRaw} -> ${componentSkuRaw} sem quantidade válida.`);
    }

    const component = {
      kitId,
      kitSkuRaw,
      kitSkuNorm: normalizeSkuKitLookup(kitSkuRaw),
      kitDescription: String(row[columns.kitDescription] ?? '').trim(),
      componentId,
      componentSkuRaw,
      componentSkuNorm: normalizeSkuComponent(componentSkuRaw),
      componentDescription: String(row[columns.componentDescription] ?? '').trim(),
      componentQty
    };

    const dedupeKey = [
      normalizeSkuKitLookup(kitSkuRaw) || normalizeSkuBasic(kitSkuRaw),
      componentId || normalizeSkuComponent(componentSkuRaw)
    ].join('::');
    const existing = map.componentsByKey.get(dedupeKey);

    if (existing) {
      if (existing.componentQty === componentQty) continue;
      if (componentQty > 0 && existing.componentQty !== componentQty) {
        map.warnings.push(`Composição duplicada conflitante em ${kitSkuRaw} -> ${componentSkuRaw}: ${existing.componentQty} vs ${componentQty}. Usada a maior quantidade.`);
        // Quando a mesma composição vem conflitante, escolhemos a maior quantidade
        // não zero para evitar subcontagem física dos componentes.
        if (componentQty > existing.componentQty) {
          existing.componentQty = componentQty;
        }
      }
      continue;
    }

    map.componentsByKey.set(dedupeKey, component);
    addCompositionLookup(map, component);
  }

  return map;
}

function mergeCompositionMaps(maps) {
  const merged = createEmptyCompositionMap();
  maps.forEach((map) => {
    map.warnings.forEach((warning) => merged.warnings.push(warning));
    map.componentsByKey.forEach((component, key) => {
      const existing = merged.componentsByKey.get(key);
      if (existing) {
        if (existing.componentQty !== component.componentQty) {
          merged.warnings.push(`Composição duplicada conflitante em ${component.kitSkuRaw} -> ${component.componentSkuRaw}: ${existing.componentQty} vs ${component.componentQty}. Usada a maior quantidade.`);
          if (component.componentQty > existing.componentQty) existing.componentQty = component.componentQty;
        }
        return;
      }
      merged.componentsByKey.set(key, component);
      addCompositionLookup(merged, component);
    });
  });
  return merged;
}

function findKitComponents(rawSku, compositionMap) {
  if (!compositionMap) return [];
  const raw = String(rawSku ?? '').trim();
  const basic = normalizeSkuBasic(rawSku);
  const kitLookup = normalizeSkuKitLookup(rawSku);
  const legacy = legacyNormalizeSku(rawSku);

  return compositionMap.byKitSkuRaw.get(raw)
    || compositionMap.byKitSkuNorm.get(basic)
    || compositionMap.byKitSkuNorm.get(kitLookup)
    || compositionMap.byKitSkuLegacy.get(legacy)
    || [];
}

function findManualComponents(rawSku) {
  const sku = normalizeSkuBasic(rawSku);
  if (!sku) return [];
  return manualSkuCompositions
    .filter((item) => normalizeSkuBasic(item.sourceSku) === sku)
    .map((item) => ({
      kitId: '',
      kitSkuRaw: item.sourceSku,
      kitSkuNorm: normalizeSkuBasic(item.sourceSku),
      kitDescription: item.componentDescription,
      componentId: '',
      componentSkuRaw: item.componentSku,
      componentSkuNorm: normalizeSkuComponent(item.componentSku),
      componentDescription: item.componentDescription,
      componentQty: item.componentQty,
      isManualComposition: true
    }));
}

function looksLikeKitSku(rawSku) {
  return /-K\d{3}[A-Z]*S?$/i.test(normalizeSkuBasic(rawSku));
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

function buildSimpleProductRows(objects, compositionMap) {
  if (!objects.length) return [];
  const headers = Object.keys(objects[0]);
  const columns = detectCurveColumns(headers);
  const map = new Map();
  const warnings = [];

  for (const row of objects) {
    const rawSku = row[columns.sku];
    const product = String(row[columns.product] ?? '').trim() || normalizeSkuBasic(rawSku);
    const quantity = parseNumber(row[columns.quantity]);
    const value = parseNumber(row[columns.value]);
    const manualComponents = findManualComponents(rawSku);
    const kitComponents = manualComponents.length ? manualComponents : findKitComponents(rawSku, compositionMap);

    if (!normalizeSkuBasic(rawSku)) {
      warnings.push(`Produto final com SKU vazio na curva: ${product || 'sem produto'}.`);
      continue;
    }

    if (kitComponents.length) {
      const totalComponentQty = kitComponents.reduce((sum, component) => sum + (component.componentQty || 0), 0);
      kitComponents.forEach((component) => {
        if (!component.componentSkuNorm) {
          warnings.push(`Produto final com SKU vazio ao explodir kit ${component.kitSkuRaw}.`);
          return;
        }
        const valueShare = totalComponentQty > 0 ? (component.componentQty || 0) / totalComponentQty : 0;
        addSimpleProductSale(map, {
          key: component.componentSkuNorm,
          sku: component.componentSkuNorm,
          rawSku: component.componentSkuRaw,
          product: component.componentDescription || component.componentSkuRaw,
          description: component.componentDescription || component.componentSkuRaw,
          olistId: component.componentId,
          soldQty: quantity * (component.componentQty || 0),
          soldValue: value * valueShare,
          sourceType: component.isManualComposition ? 'manual-composition' : 'composition',
          sourceKitSku: component.kitSkuRaw
        });
      });
      continue;
    }

    if (looksLikeKitSku(rawSku)) {
      warnings.push(`SKU ${normalizeSkuBasic(rawSku)} tem padrão de kit/fabricado, mas não foi encontrado na composição.`);
    }

    const baseSku = normalizeSkuComponent(rawSku);
    if (!baseSku) continue;

    addSimpleProductSale(map, {
      key: baseSku,
      sku: baseSku,
      rawSku,
      product,
      description: product,
      olistId: '',
      soldQty: quantity,
      soldValue: value,
      sourceType: 'curve'
    });
  }

  const rows = [...map.values()].map((item) => {
    const canonical = chooseCanonicalSource(item.sources, item.skuBase);
    return {
      ...item,
      product: canonical.name || item.skuBase,
      description: canonical.description || canonical.name || item.skuBase
    };
  });

  rows.warnings = warnings;
  return rows;
}

function addSimpleProductSale(map, sale) {
  const current = map.get(sale.key) || {
    skuBase: sale.sku,
    sku: sale.sku,
    olistId: sale.olistId || '',
    soldQty: 0,
    soldValue: 0,
    sources: []
  };

  current.soldQty += sale.soldQty;
  current.soldValue += sale.soldValue;
  if (!current.olistId && sale.olistId) current.olistId = sale.olistId;
  current.sources.push({
    rawSku: sale.rawSku,
    name: sale.product,
    description: sale.description,
    sourceType: sale.sourceType,
    sourceKitSku: sale.sourceKitSku
  });

  map.set(sale.key, current);
}

function buildStockRows(objects) {
  if (!objects.length) return [];
  const headers = Object.keys(objects[0]);
  const columns = detectStockColumns(headers);
  const map = new Map();

  for (const row of objects) {
    const rawSku = row[columns.sku];
    const baseSku = normalizeSkuComponent(rawSku);
    if (!baseSku) continue;

    const stockTotal = columns.total
      ? parseNumber(row[columns.total])
      : columns.numericHeaders.reduce((sum, header) => sum + parseNumber(row[header]), 0);
    const product = String(row[columns.product] ?? '').trim() || baseSku;

    const current = map.get(baseSku) || {
      skuBase: baseSku,
      sku: baseSku,
      olistId: String(row[columns.olistId] ?? '').trim(),
      stockTotal: 0,
      sources: []
    };

    current.stockTotal += stockTotal;
    if (!current.olistId && columns.olistId) current.olistId = String(row[columns.olistId] ?? '').trim();
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

function getFrequency(row, lowStockDays = 3) {
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
  const lowStockThreshold = soldQty > 0 ? (soldQty / 30) * lowStockDays : 0;
  if (!row.hasStockMatch) syncStatus = 'missing';
  else if ((row.stockTotal || 0) <= 0) syncStatus = 'zero';
  else if (lowStockThreshold > 0 && (row.stockTotal || 0) < lowStockThreshold) syncStatus = 'low';

  return {
    ...row,
    coverage,
    score,
    frequency,
    syncStatus,
    lowStockDays,
    lowStockThreshold
  };
}

function mergeRows(curveRows, stockRows, metricKey, thresholdA, thresholdB, lowStockDays = 3) {
  const curveWithClass = addCurveClasses(curveRows, metricKey, thresholdA, thresholdB);
  const curveMap = new Map(curveWithClass.map((item) => [item.skuBase, item]));
  const stockMap = new Map(stockRows.map((item) => [item.skuBase, item]));
  const baseSkus = [...new Set([...curveMap.keys(), ...stockMap.keys()])];
  const warnings = [];

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
    const rowWarnings = [];
    if (curve && !stock) {
      const warning = `Componente/produto ${skuBase} não encontrado no estoque.`;
      rowWarnings.push(warning);
      warnings.push(warning);
    }

    const row = {
      skuBase,
      sku: skuBase,
      olistId: curve?.olistId || stock?.olistId || '',
      product,
      description,
      soldQty: curve?.soldQty || 0,
      soldValue: curve?.soldValue || 0,
      individualPct: curve?.individualPct || 0,
      cumulativePct: curve?.cumulativePct || 0,
      abcClass: curve?.abcClass || '—',
      stockTotal: stock?.stockTotal ?? 0,
      currentQty: '',
      hasStockMatch: Boolean(stock),
      warnings: rowWarnings
    };

    return getFrequency(row, lowStockDays);
  });

  const sorted = merged.sort((a, b) => {
    const orderDiff = freqOrder.indexOf(a.frequency) - freqOrder.indexOf(b.frequency);
    if (orderDiff !== 0) return orderDiff;
    return (b.soldQty || 0) - (a.soldQty || 0);
  });
  sorted.warnings = warnings;
  return sorted;
}

function getFreqPalette(freq) {
  return palette[freq] || palette.default;
}

function getStatusLabel(status) {
  if (status === 'matched') return 'Com estoque';
  if (status === 'low') return 'Estoque baixo';
  if (status === 'zero') return 'Sem saldo';
  return 'Sem estoque';
}

function getStatusPalette(status) {
  if (status === 'matched') {
    return { bg: 'rgba(55,185,130,0.14)', border: 'rgba(55,185,130,0.38)', text: '#d6f1e6', fill: 'D9EDE3' };
  }
  if (status === 'low') {
    return { bg: 'rgba(245,196,83,0.18)', border: 'rgba(245,196,83,0.42)', text: '#fff3c7', fill: 'FCEFC7' };
  }
  return { bg: 'rgba(224,106,106,0.16)', border: 'rgba(224,106,106,0.40)', text: '#ffe0e0', fill: 'F7D6D6' };
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
  const low = rows.filter((row) => row.syncStatus === 'low').length;
  const metricLabel = els.metricSelect.value === 'value' ? 'valor vendido' : 'quantidade vendida';

  els.kpiMatched.textContent = formatInt(matched);
  els.kpiMatchedSub.textContent = `${formatInt(rows.length)} produtos simples`;
  els.kpiPriority.textContent = formatInt(highPriority);
  els.kpiPrioritySub.textContent = 'Frequência mais alta';
  els.kpiCritical.textContent = formatInt(critical);
  els.kpiCriticalSub.textContent = `${formatInt(low)} com estoque baixo`;
  els.kpiZero.textContent = formatInt(zero);
  els.kpiZeroSub.textContent = 'Estoque igual a zero';
  els.statusLine.textContent = `Resultado por produto simples. Kits convertidos em componentes.`;
  if (els.processInfo) {
    els.processInfo.textContent = `Base ABC por ${metricLabel}. Pronto para contagem física.`;
  }

  renderDistribution(rows);
  renderWarnings();
}

function renderWarnings() {
  if (!els.warningsPanel || !els.warningsList) return;
  const uniqueWarnings = [...new Set(state.warnings || [])].filter(Boolean);
  if (!uniqueWarnings.length) {
    els.warningsPanel.classList.add('hidden');
    els.warningsList.innerHTML = '';
    if (els.warningsCount) els.warningsCount.textContent = '0';
    return;
  }
  els.warningsPanel.classList.remove('hidden');
  els.warningsPanel.classList.toggle('collapsed', !state.warningsExpanded);
  if (els.warningsCount) els.warningsCount.textContent = formatInt(uniqueWarnings.length);
  if (els.warningsToggle) els.warningsToggle.setAttribute('aria-expanded', String(state.warningsExpanded));
  els.warningsList.innerHTML = uniqueWarnings.slice(0, 30)
    .map((warning) => `<li>${escapeHtml(warning)}</li>`)
    .join('');
}

function toggleWarnings() {
  state.warningsExpanded = !state.warningsExpanded;
  renderWarnings();
}

function refreshLaunchAvailability() {
  const ready = Boolean(state.curveWorkbook && state.stockWorkbook);
  els.openPanelBtn.disabled = !ready;
  els.openPanelBtn.removeAttribute?.('title');
}

function getActiveRows() {
  const query = normalizeSearch(els.searchInput.value);
  const frequency = els.frequencyFilter.value;
  const sync = els.syncFilter.value;

  return state.rows.filter((row) => {
    const haystack = normalizeSearch(`${row.product} ${row.description} ${row.sku}`);
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
    const syncLabel = getStatusLabel(row.syncStatus);
    const statusColors = getStatusPalette(row.syncStatus);
    const subtitle = row.description && row.description !== row.product ? row.description : '';

    return `
      <tr>
        <td class="product-cell">
          <div class="product-stack">
            <div class="product-title">${escapeHtml(row.product)}</div>
            ${subtitle ? `<div class="product-sub">${escapeHtml(subtitle)}</div>` : ''}
          </div>
        </td>
        <td><span class="mono">${escapeHtml(row.sku)}</span></td>
        <td>${escapeHtml(row.olistId || '')}</td>
        <td><span class="pill" style="background:${colors.bg};border-color:${colors.border};color:${colors.text};">${escapeHtml(row.frequency)}</span></td>
        <td><span class="pill pill-neutral">${escapeHtml(row.abcClass)}</span></td>
        <td>${formatInt(row.soldQty)}</td>
        <td>${formatInt(row.stockTotal)}</td>
        <td class="current-qty-cell">${escapeHtml(row.currentQty || '')}</td>
        <td><span class="pill" style="background:${statusColors.bg};border-color:${statusColors.border};color:${statusColors.text};">${syncLabel}</span></td>
      </tr>
    `;
  }).join('');

  els.tableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Produto</th>
          <th>SKU</th>
          <th>ID Olist</th>
          <th>Frequência</th>
          <th>Curva</th>
          <th>Qtd. vendida ajustada</th>
          <th>Estoque sistema</th>
          <th>QTD ATUAL</th>
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

function getPreferredCompositionSheet(workbook) {
  return workbook.SheetNames.find((sheetName) => normalizeHeader(sheetName) === 'componentes')
    || workbook.SheetNames[0]
    || '';
}

function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes.buffer;
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

    refreshLaunchAvailability();
    if (state.curveWorkbook && state.stockWorkbook && !els.appShell.classList.contains('hidden')) {
      await processData();
    }
  } catch (error) {
    console.error(`Erro ao importar a planilha (${kind}):`, error);
    alert(`Erro ao importar arquivo: ${error.message}`);
  }
}

async function loadBundledComposition() {
  if (state.compositionWorkbooks.length) return state.compositionWorkbooks[0];
  if (state.compositionLoadPromise) return state.compositionLoadPromise;

  state.compositionLoadPromise = (async () => {
  try {
    if (typeof XLSX === 'undefined') {
      throw new Error("Biblioteca SheetJS (XLSX) não encontrada. Verifique as tags <script> no seu HTML.");
    }

    let buffer = null;
    const bundledComposition = window.BUNDLED_COMPOSITION_XLS;
    if (bundledComposition?.base64) {
      buffer = base64ToArrayBuffer(bundledComposition.base64);
    } else if (window.location.protocol !== 'file:') {
      const response = await fetch(`${COMPOSITION_FILE_NAME}?v=${Date.now()}`, { cache: 'no-store' });
      if (!response.ok) {
        throw new Error(`Não foi possível carregar ${COMPOSITION_FILE_NAME}.`);
      }
      buffer = await response.arrayBuffer();
    } else {
      throw new Error(`Composição embutida não encontrada.`);
    }

    const workbook = XLSX.read(buffer, { type: 'array' });

    state.compositionWorkbooks = [workbook];
    state.compositionFileNames = [bundledComposition?.fileName || COMPOSITION_FILE_NAME];
    state.compositionSheets = [getPreferredCompositionSheet(workbook)];
    state.compositionLoadError = null;

    refreshLaunchAvailability();
    if (state.curveWorkbook && state.stockWorkbook && !els.appShell.classList.contains('hidden')) {
      await processData();
    }
    return workbook;
  } catch (error) {
    state.compositionLoadError = error;
    const warning = `${error.message} Verifique se composicao-data.js foi carregado antes do app.js.`;
    state.warnings = [...(state.warnings || []), warning];
    console.warn(warning, error);
    refreshLaunchAvailability();
    return null;
  } finally {
    state.compositionLoadPromise = null;
  }
  })();

  return state.compositionLoadPromise;
}

function getCompositionObjects() {
  return state.compositionWorkbooks.flatMap((workbook, index) => {
    const sheetName = state.compositionSheets[index] || getPreferredCompositionSheet(workbook);
    return rowsToObjects(getWorkbookRows(workbook, sheetName));
  });
}

async function processData() {
  try {
    if (!state.curveWorkbook || !state.stockWorkbook) return;

    if (!state.compositionWorkbooks.length) {
      await loadBundledComposition();
    }
    if (!state.compositionWorkbooks.length) {
      state.warnings = [
        ...(state.warnings || []),
        `Composição interna não carregada: ${COMPOSITION_FILE_NAME}. O resultado não foi atualizado.`
      ];
      renderWarnings();
      return;
    }

    const thresholdA = Math.min(100, Math.max(1, parseNumber(els.thresholdA.value) || 80));
    const thresholdB = Math.min(100, Math.max(thresholdA, parseNumber(els.thresholdB.value) || 95));
    const lowStockDays = Math.min(90, Math.max(1, parseNumber(els.lowStockDays?.value) || 3));
    els.thresholdA.value = thresholdA;
    els.thresholdB.value = thresholdB;
    if (els.lowStockDays) els.lowStockDays.value = lowStockDays;

    state.curveSheet = els.curveSheetSelect.value || state.curveSheet;
    state.stockSheet = els.stockSheetSelect.value || state.stockSheet;

    const curveObjects = rowsToObjects(getWorkbookRows(state.curveWorkbook, state.curveSheet));
    const stockObjects = rowsToObjects(getWorkbookRows(state.stockWorkbook, state.stockSheet));
    const compositionObjects = getCompositionObjects();
    state.compositionMap = buildCompositionMap(compositionObjects);
    const curveRows = buildSimpleProductRows(curveObjects, state.compositionMap);
    const stockRows = buildStockRows(stockObjects);

    const metricKey = els.metricSelect.value === 'value' ? 'soldValue' : 'soldQty';
    state.rows = mergeRows(curveRows, stockRows, metricKey, thresholdA, thresholdB, lowStockDays);
    state.warnings = [
      ...(state.compositionMap.warnings || []),
      ...(curveRows.warnings || []),
      ...(state.rows.warnings || [])
    ];
    if (!state.compositionMap.componentsByKey.size) {
      state.warnings.push('Nenhum componente gerado a partir dos arquivos de composição.');
    }

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

async function openPanel() {
  els.launchScreen.classList.add('hidden');
  els.appShell.classList.remove('hidden');
  await processData();
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
  row.eachCell((cell) => {
    cell.font = { color: { argb: '1F2937' }, name: 'Arial', size: 11 };
    cell.border = {
      top: { style: 'thin', color: { argb: 'E5E7EB' } },
      bottom: { style: 'thin', color: { argb: 'E5E7EB' } },
      left: { style: 'thin', color: { argb: 'F3F4F6' } },
      right: { style: 'thin', color: { argb: 'F3F4F6' } }
    };
  });
}

function applyExcelPillStyle(cell, fill, bold = true) {
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
  cell.font = { bold, color: { argb: '111827' }, name: 'Arial', size: 11 };
  cell.alignment = { vertical: 'middle', horizontal: 'center' };
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
    { header: 'SKU', key: 'sku', width: 20 },
    { header: 'ID Olist', key: 'olistId', width: 14 },
    { header: 'Frequência', key: 'frequency', width: 12 },
    { header: 'Curva', key: 'abcClass', width: 10 },
    { header: 'Qtd. vendida ajustada', key: 'soldQty', width: 22 },
    { header: 'Estoque sistema', key: 'stockTotal', width: 18 },
    { header: 'QTD ATUAL', key: 'currentQty', width: 18 },
    { header: 'Status', key: 'status', width: 14 }
  ];

  sheet.views = [{ state: 'frozen', ySplit: 1 }];
  sheet.autoFilter = 'A1:I1';

  const header = sheet.getRow(1);
  header.font = { bold: true, color: { argb: '111827' }, name: 'Arial', size: 11 };
  header.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F4F6' } };
  header.eachCell((cell) => {
    cell.border = { bottom: { style: 'thin', color: { argb: '374151' } } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });

  rows.forEach((item) => {
    const row = sheet.addRow({
      product: item.product,
      sku: item.sku,
      olistId: item.olistId,
      frequency: item.frequency,
      abcClass: item.abcClass,
      soldQty: item.soldQty,
      stockTotal: item.stockTotal,
      currentQty: '',
      status: getStatusLabel(item.syncStatus)
    });

    row.getCell(2).font = { name: 'Consolas', size: 11, color: { argb: '111827' } };

    applyExcelRowStyle(row, item.frequency);

    const freqCell = row.getCell(4);
    applyExcelPillStyle(freqCell, getFreqPalette(item.frequency).fill);
    applyExcelPillStyle(row.getCell(9), getStatusPalette(item.syncStatus).fill);
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
    item.sku,
    item.frequency,
    formatInt(item.stockTotal),
    '',
    getStatusLabel(item.syncStatus)
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
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });

  doc.setTextColor(17, 24, 39);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(16);
  doc.text('Contagem física por produto simples', 36, 40);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(10);
  doc.setTextColor(75, 85, 99);
  doc.text('Kits convertidos em componentes. Preencha QTD ATUAL manualmente.', 36, 58);

  doc.autoTable({
    startY: 74,
    head: [['Produto', 'SKU', 'Freq.', 'Estoque', 'QTD ATUAL', 'Status']],
    body: buildPdfRows(rows),
    theme: 'grid',
    styles: {
      fontSize: 7.5,
      cellPadding: 4,
      textColor: [28, 37, 51],
      lineColor: [229, 231, 235],
      lineWidth: 0.5
    },
    headStyles: {
      fillColor: [243, 244, 246],
      textColor: [17, 24, 39],
      fontStyle: 'bold'
    },
    columnStyles: {
      0: { cellWidth: 190 },
      1: { cellWidth: 86 },
      2: { cellWidth: 36, halign: 'center' },
      3: { cellWidth: 52, halign: 'right' },
      4: { cellWidth: 62 },
      5: { cellWidth: 72, halign: 'center' }
    },
    didParseCell(data) {
      if (data.section !== 'body') return;
      const row = rows[data.row.index];
      if (data.column.index === 2) {
        data.cell.styles.fillColor = hexToRgb(`#${getFreqPalette(row.frequency).fill}`);
        data.cell.styles.fontStyle = 'bold';
      }
      if (data.column.index === 5) {
        data.cell.styles.fillColor = hexToRgb(`#${getStatusPalette(row.syncStatus).fill}`);
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
    'Produto', 'SKU', 'ID Olist', 'Frequência', 'Curva',
    'Qtd. vendida ajustada', 'Estoque sistema', 'QTD ATUAL', 'Status'
  ];

  const body = rows.map((item) => [
    item.product,
    item.sku,
    item.olistId,
    item.frequency,
    item.abcClass,
    item.soldQty,
    item.stockTotal,
    '',
    getStatusLabel(item.syncStatus)
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
els.openPanelBtn.addEventListener('click', () => openPanel());
els.backToStartBtn.addEventListener('click', backToStart);
els.processBtn.addEventListener('click', () => processData());
els.metricSelect.addEventListener('change', () => processData());
els.curveSheetSelect.addEventListener('change', () => processData());
els.stockSheetSelect.addEventListener('change', () => processData());
els.thresholdA.addEventListener('change', () => processData());
els.thresholdB.addEventListener('change', () => processData());
els.lowStockDays?.addEventListener('change', () => processData());
els.warningsToggle?.addEventListener('click', toggleWarnings);
els.searchInput.addEventListener('input', renderTable);
els.frequencyFilter.addEventListener('change', renderTable);
els.syncFilter.addEventListener('change', renderTable);
els.exportXlsxBtn.addEventListener('click', exportXlsx);
els.exportPdfBtn.addEventListener('click', exportPdf);
els.exportCsvBtn.addEventListener('click', exportCsv);

loadBundledComposition();
