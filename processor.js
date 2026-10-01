const ExcelJS = require('exceljs');

const NEW_HEADER = 'REVENDA / USO OU CONSUMO / IMOBILIZADO';
const OPTIONS = ['REVENDA', 'USO OU CONSUMO', 'IMOBILIZADO'];
const HEADER_SCAN_ROWS = 50;

const TARGETS = [
  { nome: 'Chave NF', re: /^chave nf$/ },
  { nome: 'Tags (etiquetas)', re: /^tags?( \(etiquetas?\))?$/ },
  { nome: 'Eventos', re: /^eventos$/ },
];

function norm(v) {
  return String(v == null ? '' : v)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function cellText(cell) {
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((t) => t.text).join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    return '';
  }
  return String(v);
}

function hasValue(cell) {
  const v = cell.value;
  if (v == null || v === '') return false;
  if (typeof v === 'object' && v.richText) return v.richText.some((t) => t.text !== '');
  return true;
}

function colLetter(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function colNumber(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

// Localiza a linha de cabeçalho e as colunas a excluir em uma aba.
function findTargets(ws) {
  let best = null;
  const limit = Math.min(ws.rowCount, HEADER_SCAN_ROWS);
  for (let r = 1; r <= limit; r++) {
    const found = {};
    ws.getRow(r).eachCell({ includeEmpty: false }, (cell, c) => {
      const t = norm(cellText(cell));
      for (const tg of TARGETS) if (!found[tg.nome] && tg.re.test(t)) found[tg.nome] = c;
    });
    const n = Object.keys(found).length;
    if (n && (!best || n > best.n)) best = { row: r, found, n };
  }
  return best;
}

const REF_RE = /(?<![A-Za-z0-9_.])((?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\dA-Za-z_(])/g;

function sheetOfRef(prefix, ownerName) {
  if (!prefix) return ownerName;
  return prefix.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'");
}

function formulaOf(cell) {
  const v = cell.value;
  return v && typeof v === 'object' && v.formula ? v.formula : null;
}

function shiftFormulas(wb, ws, removedSorted) {
  const removedSet = new Set(removedSorted);
  const shift = (c) => c - removedSorted.filter((x) => x < c).length;
  let changed = 0;
  wb.eachSheet((sh) => {
    sh.eachRow((row) => {
      row.eachCell((cell) => {
        const v = cell.value;
        if (!v || typeof v !== 'object' || !v.formula) return;
        const nf = v.formula.replace(REF_RE, (all, pre, d1, col, d2, rw) => {
          if (sheetOfRef(pre, sh.name) !== ws.name) return all;
          const n = colNumber(col);
          if (removedSet.has(n)) return '#REF!';
          return `${pre || ''}${d1}${colLetter(shift(n))}${d2}${rw}`;
        });
        if (nf !== v.formula) { cell.value = { ...v, formula: nf }; changed++; }
      });
    });
  });
  return changed;
}

const THIN = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };

// Deixa a aba pronta para uso: cabeçalho destacado, filtro, painel congelado,
// bordas, larguras ajustadas, formatos de valor/data e destaque na coluna de classificação.
function formatarSaida(ws, headerRow, lastRow, newCol, totalRows) {
  // ExcelJS compartilha o objeto de estilo entre células iguais: sempre troca o estilo inteiro.
  const apply = (cell, patch) => {
    const st = JSON.parse(JSON.stringify(cell.style || {}));
    patch(st);
    cell.style = st;
  };
  const widths = {};
  for (let c = 1; c <= newCol; c++) {
    const hc = ws.getCell(headerRow, c);
    const head = norm(cellText(hc));
    widths[c] = cellText(hc).length + 3;

    apply(hc, (st) => {
      st.font = { ...(st.font || {}), bold: true };
      st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
      st.border = BORDER;
      st.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });

    for (let r = headerRow + 1; r <= lastRow; r++) {
      const cell = ws.getCell(r, c);
      const v = cell.value;
      const len = v instanceof Date ? 10 : typeof v === 'number' ? 14 : cellText(cell).length;
      apply(cell, (st) => {
        st.border = BORDER;
        if (typeof v === 'number' && /valor|total/.test(head)) st.numFmt = '#,##0.00';
        if (v instanceof Date && /data/.test(head)) st.numFmt = 'dd/mm/yyyy';
        if (totalRows.has(r)) st.font = { ...(st.font || {}), bold: true };
        if (c === newCol) {
          st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
          st.alignment = { horizontal: 'center' };
        }
      });
      if (len + 2 > widths[c]) widths[c] = len + 2;
    }
  }
  for (let c = 1; c <= newCol; c++) ws.getColumn(c).width = Math.min(Math.max(widths[c], 10), 60);

  const view = (ws.views && ws.views[0]) || {};
  ws.views = [{ ...view, state: 'frozen', xSplit: view.xSplit || 0, ySplit: headerRow, activeCell: 'A1' }];
  ws.autoFilter = { from: { row: headerRow, column: 1 }, to: { row: headerRow, column: newCol } };
}

function processSheet(wb, ws) {
  const hit = findTargets(ws);
  if (!hit) return null;

  const sr = { aba: ws.name, linhaCabecalho: hit.row, removidas: [], naoEncontradas: [] };
  for (const tg of TARGETS) {
    if (hit.found[tg.nome]) sr.removidas.push({ nome: tg.nome, posicao: colLetter(hit.found[tg.nome]) });
    else sr.naoEncontradas.push(tg.nome);
  }
  const removed = Object.values(hit.found).sort((a, b) => a - b);
  const removedSet = new Set(removed);

  const oldLastRow = ws.rowCount;
  const oldColCount = ws.columnCount;
  const remap = (c) => c - removed.filter((x) => x < c).length;

  // linhas de total nas colunas excluídas (ficam fora da lista suspensa)
  const totalRows = new Set();
  for (let r = hit.row + 1; r <= oldLastRow; r++) {
    for (const c of removed) {
      if (/^(sub)?total/.test(norm(cellText(ws.getCell(r, c))))) totalRows.add(r);
    }
  }

  // mesclagens
  const merges = Object.values(ws._merges || {}).map((m) => m.model);
  const newMerges = [];
  for (const m of merges) {
    const top = m.top, bottom = m.bottom, left = m.left, right = m.right;
    const keep = [];
    for (let c = left; c <= right; c++) if (!removedSet.has(c)) keep.push(c);
    const touches = keep.length !== right - left + 1;
    if (!touches) { newMerges.push({ top, bottom, left: remap(left), right: remap(right) }); continue; }
    const master = ws.getCell(top, left);
    const text = hasValue(master) ? master.value : null;
    if (!keep.length) {
      continue;
    }
    newMerges.push({ top, bottom, left: remap(keep[0]), right: remap(keep[keep.length - 1]), text, style: master.style, movedFrom: removedSet.has(left) });
  }
  for (const m of merges) ws.unmergeCells(m.top, m.left, m.bottom, m.right);

  // larguras e estilos de coluna antes de excluir
  const colProps = [];
  for (let c = 1; c <= oldColCount; c++) {
    const col = ws.getColumn(c);
    colProps[c] = { width: col.width, hidden: col.hidden, style: col.style };
  }

  // congelamento de painéis
  const view = ws.views && ws.views[0];
  if (view && view.state === 'frozen' && view.xSplit) {
    view.xSplit = Math.max(0, view.xSplit - removed.filter((c) => c <= view.xSplit).length);
  }
  const af = ws.autoFilter;

  // excluir do fim para o começo
  for (const c of [...removed].reverse()) ws.spliceColumns(c, 1);
  shiftFormulas(wb, ws, removed);

  for (let c = 1; c <= oldColCount; c++) {
    if (removedSet.has(c)) continue;
    const col = ws.getColumn(remap(c));
    if (colProps[c].width) col.width = colProps[c].width;
    col.hidden = colProps[c].hidden;
  }

  // mesclagens ajustadas
  for (const m of newMerges) {
    if (m.text != null && m.movedFrom) {
      const cell = ws.getCell(m.top, m.left);
      cell.value = m.text;
      cell.style = m.style;
    }
    ws.mergeCells(m.top, m.left, m.bottom, m.right);
  }

  // autofiltro
  if (af && typeof af === 'string') {
    const mm = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(af);
    if (mm) {
      const a = colNumber(mm[1]), b = colNumber(mm[3]);
      const keep = []; for (let c = a; c <= b; c++) if (!removedSet.has(c)) keep.push(remap(c));
      ws.autoFilter = keep.length ? `${colLetter(keep[0])}${mm[2]}:${colLetter(keep[keep.length - 1])}${mm[4]}` : null;
    }
  }

  // última coluna preenchida e última linha de dados
  let lastCol = 0, lastRow = hit.row;
  ws.eachRow({ includeEmpty: false }, (row, r) => {
    row.eachCell({ includeEmpty: false }, (cell, c) => {
      if (hasValue(cell)) { if (c > lastCol) lastCol = c; if (r > lastRow) lastRow = r; }
    });
  });
  const newCol = lastCol + 1;

  const refHeader = ws.getCell(hit.row, lastCol);
  const hc = ws.getCell(hit.row, newCol);
  hc.value = NEW_HEADER;
  hc.style = JSON.parse(JSON.stringify(refHeader.style || {}));
  ws.getColumn(newCol).width = Math.max(NEW_HEADER.length + 3, 20);

  // lista suspensa nas linhas de dados (exceto totais)
  let dados = 0;
  let runStart = null;
  const ranges = [];
  for (let r = hit.row + 1; r <= lastRow + 1; r++) {
    const isData = r <= lastRow && !totalRows.has(r) && ws.getRow(r).hasValues;
    if (isData) { dados++; if (runStart == null) runStart = r; }
    else if (runStart != null) { ranges.push([runStart, r - 1]); runStart = null; }
  }
  for (const [a, b] of ranges) {
    for (let r = a; r <= b; r++) {
      ws.getCell(r, newCol).dataValidation = {
        type: 'list',
        allowBlank: true,
        formulae: [`"${OPTIONS.join(',')}"`],
        showErrorMessage: true,
        errorTitle: 'Valor inválido',
        error: `Escolha: ${OPTIONS.join(', ')}`,
      };
    }
  }

  formatarSaida(ws, hit.row, lastRow, newCol, totalRows);

  sr.colunaAdicionada ={ nome: NEW_HEADER, posicao: colLetter(newCol) };
  sr.linhasDeDados = dados;
  sr.linhasTotalIgnoradas = [...totalRows];
  return sr;
}

async function ajustarPlanilha(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const report = { abas: [], abasSemColunas: [] };
  const sheets = [];
  wb.eachSheet((ws) => sheets.push(ws));

  for (const ws of sheets) {
    const r = processSheet(wb, ws);
    if (!r) { report.abasSemColunas.push(ws.name); continue; }
    report.abas.push(r);
  }

  if (!report.abas.length) {
    report.erro = 'Nenhuma aba contém as colunas "Chave NF", "Tags (etiquetas)" ou "Eventos".';
    return { report, buffer: null };
  }

  const out = await wb.xlsx.writeBuffer();
  return { report, buffer: Buffer.from(out) };
}

module.exports = { ajustarPlanilha };
