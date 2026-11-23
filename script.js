/* ══════════════════════════════════════════════════════════════
   DataForge Enterprise Consolidator Suite — script.js
   Modules: FileManager | FileReaderModule | DataCleaner |
            ChronoSorter | ExcelGenerator | Pipeline
   ══════════════════════════════════════════════════════════════ */

// ══════════════════════════════════════════════════════════════
// MODULE: FileManager
// Handles file state, validation, and file-badge UI
// ══════════════════════════════════════════════════════════════
const FM = (() => {
  let _files = [];

  const SUPPORTED = ['csv', 'tsv', 'txt', 'xlsx'];

  const getExt = f => f.name.split('.').pop().toLowerCase();

  const fmtSize = bytes => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
  };

  const getTypeLabel = ext => ({ csv: 'CSV', tsv: 'TSV', txt: 'TXT', xlsx: 'XLSX' }[ext] || '???');

  const add = newFiles => {
    for (const f of newFiles) {
      const ext = getExt(f);
      if (!SUPPORTED.includes(ext)) {
        _appendBadge(f, true);
        continue;
      }
      if (_files.find(x => x.name === f.name && x.size === f.size)) continue;
      _files.push(f);
      _appendBadge(f, false);
    }
    _updateButtons();
  };

  const remove = name => {
    _files = _files.filter(f => f.name !== name);
    document.querySelectorAll('.file-item').forEach(el => {
      if (el.dataset.name === name) el.remove();
    });
    _updateButtons();
  };

  const clear = () => {
    _files = [];
    document.getElementById('fileList').innerHTML = '';
    _updateButtons();
  };

  const get = () => _files;

  const _updateButtons = () => {
    const has = _files.length > 0;
    document.getElementById('processBtn').disabled = !has;
    document.getElementById('clearBtn').disabled = !has;
  };

  const _appendBadge = (f, isError) => {
    const ext = getExt(f);
    const el = document.createElement('div');
    el.className = 'file-item' + (isError ? ' error' : '');
    el.dataset.name = f.name;
    el.innerHTML = `
      <span class="file-icon">${getTypeLabel(ext)}</span>
      <div class="file-meta">
        <div class="file-name">${f.name}</div>
        <div class="file-size">${fmtSize(f.size)}</div>
      </div>
      <span class="file-status ${isError ? 'err' : 'ok'}">${isError ? '✕ ERR' : '✓ LISTO'}</span>
      ${!isError ? `<button class="file-remove" onclick="FM.remove('${f.name}')" title="Quitar">×</button>` : ''}
    `;
    document.getElementById('fileList').appendChild(el);
  };

  return { add, remove, clear, get };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: FileReaderModule
// Multi-format parser: CSV / TSV / TXT (auto-delimiter) + XLSX
// Integrates convertidor-csv multi-format ingestion logic
// ══════════════════════════════════════════════════════════════
const FileReaderModule = (() => {

  // Auto-detect delimiter by frequency in the first 2000 chars
  const detectDelimiter = text => {
    const candidates = [',', '\t', '|', ';'];
    let best = ',', max = 0;
    const sample = text.slice(0, 2000);
    for (const d of candidates) {
      const count = (sample.match(new RegExp('\\' + d, 'g')) || []).length;
      if (count > max) { max = count; best = d; }
    }
    return best;
  };

  // Decode as UTF-8; fall back to Windows-1252 (Excel "CSV" export) if invalid bytes appear
  const _decode = async file => {
    const buf = await file.arrayBuffer();
    const utf8 = new TextDecoder('utf-8').decode(buf);
    const text = utf8.includes('�') ? new TextDecoder('windows-1252').decode(buf) : utf8;
    return text.replace(/^﻿/, '');
  };

  // Numbers: "19.990" / "1.234.567" = CL thousands, "1.234,56" = CL decimal, "12.5" = plain.
  // Currency "$42.000" / "$-384.000" / "-$5.000" → number. Leading-zero codes ("000123") stay text.
  const toValue = v => {
    if (typeof v === 'number') return v;
    let s = String(v ?? '').trim();
    if (s === '') return null;
    const cur = s.match(/^(-?)\s*\$\s*(-?)\s*(\d[\d.,]*)$/);
    if (cur) s = (cur[1] || cur[2]) + cur[3];
    if (/^-?0\d/.test(s)) return cur ? Number(s.replace(/\./g, '')) : s;
    if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) return Number(s.replace(/\./g, ''));
    if (/^-?\d{1,3}(\.\d{3})*,\d+$/.test(s)) return Number(s.replace(/\./g, '').replace(',', '.'));
    if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
    return cur ? String(v).trim() : s;
  };

  const _filled = r => r.filter(c => String(c ?? '').trim() !== '').length;

  // Header row = first row whose filled cells reach >=50% of the widest row.
  // Skips report preambles like "Fecha: ,01/04/2026" / ",,,,,,Gastos".
  const findHeaderRow = rows => {
    const width = Math.max(...rows.map(_filled));
    return Math.max(0, rows.findIndex(r => _filled(r) >= Math.max(2, width * 0.5)));
  };

  // Turn raw 2D rows into { headers, data, title } with preamble skipped and unique headers
  const rowsToDataset = (name, rows) => {
    const h = findHeaderRow(rows);
    // A single-cell line above the header (",,,,,Resumen Multi Estación") is the table title
    const titleRow = rows.slice(0, h).reverse().find(r => _filled(r) === 1);
    const title = titleRow ? String(titleRow.find(c => String(c ?? '').trim())).trim() : null;
    const width = Math.max(...rows.slice(h).map(r => r.length));
    const raw = Array.from({ length: width }, (_, i) => rows[h][i] ?? '');
    // Drop fully-empty trailing columns
    let last = width - 1;
    while (last > 0 && !String(raw[last]).trim() && rows.slice(h + 1).every(r => !String(r[last] ?? '').trim())) last--;
    const headers = DataCleaner.sanitizeHeaders(raw.slice(0, last + 1).map((c, i) => String(c).trim() || `Col_${i + 1}`));
    const data = rows.slice(h + 1)
      .filter(r => _filled(r) > 0)
      .map(r => Object.fromEntries(headers.map((k, i) => [k, toValue(r[i])])));
    return { name, title, headers, data };
  };

  // Split raw rows into blocks separated by blank lines. Blocks that continue the
  // previous table (same row width, similar fill) are re-joined, so a stray blank
  // line inside the data doesn't cut it in two.
  const _splitBlocks = rows => {
    const blocks = [];
    let cur = [];
    for (const r of rows) {
      if (_filled(r)) cur.push(r);
      else if (cur.length) { blocks.push(cur); cur = []; }
    }
    if (cur.length) blocks.push(cur);
    const merged = [];
    const maxFilled = b => Math.max(...b.map(_filled));
    for (const b of blocks) {
      const prev = merged[merged.length - 1];
      const w = prev && prev[prev.length - 1].length;
      if (prev && w > 2 && b.every(r => r.length === w) && maxFilled(b) >= maxFilled(prev) * 0.5) prev.push(...b);
      else merged.push(b);
    }
    return merged;
  };

  // Key/value block: every row has at most 2 filled cells ("Monto Neto(A)", "$1.286.706.755")
  const _isKeyValue = b => b.every(r => _filled(r) <= 2);

  // Whole file → main table (largest block) + extra sheets for report summaries.
  // All key/value summary blocks become one table: one row per block, one column per key.
  const rowsToReport = (name, rows) => {
    const blocks = _splitBlocks(rows);
    const mainIdx = blocks.reduce((best, b, i) => b.length > blocks[best].length ? i : best, 0);
    const main = rowsToDataset(name, blocks[mainIdx]);
    const extras = [];
    const kvRows = [];
    const keyNames = new Map(); // whitespace-insensitive key → first spelling seen
    blocks.forEach((b, i) => {
      if (i === mainIdx) return;
      if (_isKeyValue(b)) {
        const obj = {};
        b.forEach((r, j) => {
          const cells = r.filter(c => String(c ?? '').trim() !== '').map(c => String(c).trim());
          // A lone title ("Resumen") goes in the label column the other blocks use ("Nombre de Sucursal")
          const labelCol = keyNames.size ? keyNames.values().next().value : 'Sección';
          let [k, v] = cells.length === 1 && j === 0 ? [labelCol, cells[0]] : cells;
          if (cells.length === 1 && j > 0) { k = cells[0]; v = ''; }
          const norm = k.replace(/\s+/g, '').toLowerCase();
          if (!keyNames.has(norm)) keyNames.set(norm, k);
          obj[keyNames.get(norm)] = toValue(v);
        });
        kvRows.push(obj);
      } else {
        const t = rowsToDataset(name, b);
        if (t.data.length) extras.push({ ...t, title: t.title || `Tabla ${extras.length + 1}` });
      }
    });
    if (kvRows.length) {
      const headers = [...new Set(kvRows.flatMap(Object.keys))];
      extras.unshift({ name, title: 'Resumen', headers, data: kvRows.map(r => Object.fromEntries(headers.map(h => [h, r[h] ?? null]))) });
    }
    return { ...main, extras };
  };

  // Parse flat text files (CSV, TSV, TXT, pipe-separated, etc.)
  const _parseText = async file => {
    const text = await _decode(file);
    const res = Papa.parse(text, { header: false, skipEmptyLines: false, delimiter: detectDelimiter(text) });
    if (!res.data.some(r => _filled(r))) throw new Error(`El archivo "${file.name}" está vacío.`);
    return rowsToReport(file.name, res.data);
  };

  // Parse XLSX via ExcelJS (convertidor-csv repo integration)
  const _parseXLSX = async file => {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    if (!ws) throw new Error(`El archivo "${file.name}" no contiene hojas de cálculo.`);

    const rows = [];
    const width = ws.columnCount;
    ws.eachRow({ includeEmpty: true }, row => {
      const vals = row.values.slice(1).map(v => {
        if (v instanceof Date) return v.toLocaleDateString('es-CL');
        if (v && typeof v === 'object' && v.result !== undefined) return v.result;
        if (v && typeof v === 'object' && v.richText) return v.richText.map(t => t.text).join('');
        if (v && typeof v === 'object' && v.text !== undefined) return v.text;
        return v ?? '';
      });
      // Pad to sheet width so block detection sees equal row lengths, like a CSV
      rows.push(Array.from({ length: width }, (_, i) => vals[i] ?? ''));
    });
    if (!rows.some(r => _filled(r))) throw new Error(`El archivo "${file.name}" está vacío.`);
    return rowsToReport(file.name, rows);
  };

  // Public: route by extension
  const parse = async file => {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'xlsx') return _parseXLSX(file);
    return _parseText(file);
  };

  return { parse, rowsToDataset, toValue };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: DataCleaner
// Removes trash rows (totals, subtotals, blanks)
// Sanitizes duplicate / empty column headers
// ══════════════════════════════════════════════════════════════
const DataCleaner = (() => {

  const TRASH_PATTERNS = [
    /total\s*producci[oó]n/i,
    /\btotal(es)?\b/i,
    /subtotal/i,
  ];

  // A total row has a short cell that IS a total label (e.g. "Total", "Subtotal:"),
  // not a long comment that merely mentions "total".
  const _isTrash = row => {
    const vals = Object.values(row);
    if (vals.every(v => v === null || v === '')) return true;
    return vals.some(v => typeof v === 'string' && v.length <= 40 &&
      /^\s*(sub)?total/i.test(v) && TRASH_PATTERNS.some(p => p.test(v)));
  };

  // Remove trash rows (totals, blanks, subtotals)
  const clean = (data, filterTotals) => {
    if (!filterTotals) return data;
    return data.filter(row => !_isTrash(row));
  };

  // Rename duplicate or empty column headers so ExcelJS doesn't collapse
  const sanitizeHeaders = headers => {
    const seen = new Map();
    return headers.map(h => {
      let base = h ? String(h).trim() : 'Col';
      if (!base) base = 'Col';
      const count = seen.get(base) || 0;
      seen.set(base, count + 1);
      return count === 0 ? base : `${base}_${count}`;
    });
  };

  return { clean, sanitizeHeaders };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: ChronoSorter
// Universal date/time parser + strict chronological sort
// Supports DD/MM/YYYY, YYYY-MM-DD, AM/PM and 24h formats
// ══════════════════════════════════════════════════════════════
const ChronoSorter = (() => {

  const _parseDate = row => {
    const keys = Object.keys(row);
    const dateKey = keys.find(k => k && /fecha|date/i.test(k));
    const timeKey = keys.find(k => k && /hora|hour|time/i.test(k));

    const d = dateKey ? String(row[dateKey] || '').trim() : '';
    const h = timeKey ? String(row[timeKey] || '').trim() : '';

    if (!d) return new Date(0);

    let dt;
    // Optional time inside the date cell: "01/04/2026 00:00", "2026-04-01T13:45:00", "01/04/2026 04:49 PM"
    const rCL  = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ T](.+))?$/;   // DD/MM/YYYY [time]
    const rISO = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:[ T](.+))?$/;    // YYYY-MM-DD [time]
    let inlineTime = '';

    if (rCL.test(d)) {
      const [, dd, mm, yyyy, t] = d.match(rCL);
      dt = new Date(+yyyy, +mm - 1, +dd); inlineTime = t || '';
    } else if (rISO.test(d)) {
      const [, yyyy, mm, dd, t] = d.match(rISO);
      dt = new Date(+yyyy, +mm - 1, +dd); inlineTime = t || '';
    } else {
      dt = new Date(d);
      if (isNaN(dt.getTime())) return new Date(0);
    }

    // Inject time component (from the hora column, else from the date cell itself)
    const time = h || inlineTime;
    if (time) {
      const parts = time.trim().split(' ');
      const [hrs, mins] = parts[0].split(':');
      let hh = parseInt(hrs) || 0;
      let mm2 = parseInt(mins) || 0;
      const secs = parseInt(parts[0].split(':')[2]) || 0;
      if (parts[1]) {
        const ampm = parts[1].toUpperCase();
        if (ampm === 'PM' && hh < 12) hh += 12;
        if (ampm === 'AM' && hh === 12) hh = 0;
      }
      dt.setHours(hh, mm2, secs);
    }

    return dt;
  };

  // Parse each row's date once (not on every comparison); stable for equal dates
  const sort = data => data.map(r => [_parseDate(r).getTime(), r]).sort((a, b) => a[0] - b[0]).map(x => x[1]);

  return { sort };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: ExcelGenerator
// Corporate formatting engine — reportsformat integration
// Applies themes, dynamic widths, accounting formats, tab colors
// ══════════════════════════════════════════════════════════════
const ExcelGenerator = (() => {

  // reportsformat: Corporate palette constants (ARGB)
  const CORP = {
    header_font : 'FFFFFFFF',
    tab_color   : 'FF2563EB',
    currency_kw : ['monto','recauda','total','tarifa','gasto','importe','valor','precio','$'],
  };

  const _isCurrencyCol = name =>
    CORP.currency_kw.some(kw => name.toLowerCase().includes(kw));

  // Dynamic column width: max of header length vs longest cell value
  const _calcWidth = (safeHeader, rawHeader, rows) => {
    let max = safeHeader.length;
    for (const row of rows) {
      const val = String(row[rawHeader] ?? '');
      if (val.length > max) max = val.length;
    }
    return Math.min(Math.max(max + 2, 10), 45);
  };

  const generate = async (ws, rawHeaders, safeHeaders, dataRows, opts) => {
    const { theme, freeze, accounting } = opts;

    // Excel cells max 32767 chars; longer values corrupt the file
    const tableRows = dataRows.map(r => rawHeaders.map(h => {
      const v = r[h] ?? '';
      return typeof v === 'string' && v.length > 32767 ? v.slice(0, 32767) : v;
    }));

    // Build Excel Table with corporate theme
    ws.addTable({
      name: `Tabla${ws.id}`,
      ref: 'A1',
      headerRow: true,
      style: { theme, showRowStripes: true, showBandedRows: true },
      columns: safeHeaders.map(h => ({ name: h, filterButton: true })),
      rows: tableRows
    });

    // Apply dynamic widths and accounting formats per column
    ws.columns.forEach((col, i) => {
      col.width = _calcWidth(safeHeaders[i], rawHeaders[i], dataRows);
      if (accounting && _isCurrencyCol(safeHeaders[i])) {
        col.numFmt = '"$"#,##0';
      }
    });

    // Corporate header row styling
    const headerRow = ws.getRow(1);
    headerRow.eachCell(cell => {
      cell.font = { bold: true, color: { argb: CORP.header_font }, name: 'Calibri', size: 11 };
    });
    headerRow.height = 22;

    // Uniform data row heights (reportsformat band style)
    for (let i = 2; i <= dataRows.length + 1; i++) {
      ws.getRow(i).height = 16;
    }

    // Freeze header row
    if (freeze) {
      ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1 }];
    }

    // Corporate tab color
    ws.properties.tabColor = { argb: CORP.tab_color };
  };

  return { generate };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: Pipeline
// Async orchestrator — drives all 5 processing stages with
// live progress feedback, logging, and stats display
// ══════════════════════════════════════════════════════════════
const Pipeline = (() => {

  let _running = false;
  const STEP_IDS = ['step-read', 'step-validate', 'step-clean', 'step-format', 'step-export'];

  // ── UI helpers ──────────────────────────────────────────────

  const _setStep = idx => {
    STEP_IDS.forEach((id, i) => {
      const el = document.getElementById(id);
      el.classList.remove('active', 'done');
      if (i < idx) el.classList.add('done');
      if (i === idx) el.classList.add('active');
    });
  };

  const _setProgress = pct => {
    document.getElementById('progressBar').style.width = pct + '%';
  };

  const _log = (msg, type = 'info') => {
    const panel = document.getElementById('logPanel');
    panel.classList.add('visible');
    const ts = new Date().toLocaleTimeString('es', { hour12: false, hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const line = document.createElement('div');
    line.className = `log-line ${type}`;
    line.innerHTML = `<span class="log-ts">${ts}</span><span class="log-msg">${msg}</span>`;
    panel.appendChild(line);
    panel.scrollTop = panel.scrollHeight;
  };

  const _showStats = (files, rows, dropped, cols) => {
    const bar = document.getElementById('statsBar');
    bar.style.display = 'grid';
    document.getElementById('sFiles').textContent = files;
    document.getElementById('sRows').textContent = rows.toLocaleString('es');
    document.getElementById('sDropped').textContent = dropped;
    document.getElementById('sCols').textContent = cols;
  };

  const _showResult = (ok, title, body) => {
    const el = document.getElementById('resultBanner');
    el.style.display = 'block';
    el.className = `result-banner ${ok ? 'success' : 'error'}`;
    document.getElementById('resultTitle').textContent = title;
    document.getElementById('resultBody').textContent = body;
  };

  const _reset = () => {
    STEP_IDS.forEach(id => document.getElementById(id).classList.remove('active', 'done'));
    _setProgress(0);
    document.getElementById('progressWrap').style.display = 'none';
    document.getElementById('logPanel').innerHTML = '';
    document.getElementById('logPanel').classList.remove('visible');
    document.getElementById('resultBanner').style.display = 'none';
    document.getElementById('statsBar').style.display = 'none';
  };

  // Non-blocking tick so the browser can repaint between steps
  const _tick = () => new Promise(r => setTimeout(r, 30));

  // ── Main orchestrator ────────────────────────────────────────

  const run = async () => {
    if (_running) return;
    _running = true;
    _reset();

    const processBtn = document.getElementById('processBtn');
    const clearBtn   = document.getElementById('clearBtn');
    processBtn.disabled = true;
    clearBtn.disabled = true;

    document.getElementById('progressWrap').style.display = 'block';
    _setProgress(5);

    const files = FM.get();
    const opts = {
      chron      : document.getElementById('toggleChron').checked,
      filter     : document.getElementById('toggleFilter').checked,
      accounting : document.getElementById('toggleAcct').checked,
      freeze     : document.getElementById('toggleFreeze').checked,
      theme      : document.getElementById('tableTheme').value,
      outputName : document.getElementById('outputName').value.trim() || 'Consolidado_Maestro'
    };

    try {
      // ── STEP 1: READ & PARSE ────────────────────────────────
      _setStep(0);
      _log(`Iniciando lectura de ${files.length} archivo(s)...`);
      await _tick();

      const datasets = await Promise.all(files.map(f => FileReaderModule.parse(f)));
      _setProgress(25);
      _log(`${datasets.length} archivo(s) parseados correctamente.`, 'ok');

      // ── STEP 2: VALIDATE STRUCTURE ──────────────────────────
      _setStep(1);
      _log('Validando consistencia de estructura de columnas...');
      await _tick();

      // Different structures are merged by column name (union, first-seen order)
      const rawHeaders = [...new Set(datasets.flatMap(d => d.headers))];
      datasets.forEach(d => {
        if (d.headers.length !== rawHeaders.length)
          _log(`"${d.name}" tiene ${d.headers.length} de ${rawHeaders.length} columnas; las faltantes quedan vacías.`, 'info');
      });
      _log(`Estructura validada: ${rawHeaders.length} columnas.`, 'ok');
      _setProgress(40);

      // ── STEP 3: CLEAN & SORT ────────────────────────────────
      _setStep(2);
      _log('Uniendo, limpiando y ordenando datos...');
      await _tick();

      let master = datasets.flatMap(d => d.data);
      const rawCount = master.length;

      if (opts.filter) master = DataCleaner.clean(master, true);
      const dropped = rawCount - master.length;

      if (opts.chron) master = ChronoSorter.sort(master);

      _log(`${master.length.toLocaleString('es')} filas consolidadas. ${dropped} filas filtradas.`, 'ok');
      _setProgress(60);
      _showStats(files.length, master.length, dropped, rawHeaders.length);

      // ── STEP 4: FORMAT ──────────────────────────────────────
      _setStep(3);
      _log('Aplicando formato corporativo al binario Excel...');
      await _tick();

      const safeHeaders = DataCleaner.sanitizeHeaders(rawHeaders);

      const wb = new ExcelJS.Workbook();
      wb.creator = 'DataForge Suite v2.0';
      wb.created = new Date();

      const ws = wb.addWorksheet('Reporte Consolidado', {
        properties: { defaultRowHeight: 16 },
        pageSetup: { orientation: 'landscape', fitToPage: true }
      });

      await ExcelGenerator.generate(ws, rawHeaders, safeHeaders, master, opts);

      // Report summaries found below/above the detail (per-branch totals, etc.) → own sheets
      const usedNames = new Set(['reporte consolidado']);
      for (const d of datasets) {
        for (const ex of d.extras || []) {
          let base = (datasets.length > 1 ? `${ex.title} - ${d.name}` : ex.title)
            .replace(/[\\\/\?\*\[\]:]/g, '_').slice(0, 28) || 'Resumen';
          let sheetName = base, n = 2;
          while (usedNames.has(sheetName.toLowerCase())) sheetName = `${base} ${n++}`;
          usedNames.add(sheetName.toLowerCase());
          const exWs = wb.addWorksheet(sheetName, { properties: { defaultRowHeight: 16 } });
          await ExcelGenerator.generate(exWs, ex.headers, DataCleaner.sanitizeHeaders(ex.headers), ex.data, opts);
          _log(`Hoja "${sheetName}": ${ex.data.length} filas de resumen.`, 'ok');
        }
      }
      _log('Tema, bandas, anchos dinámicos y formatos contables aplicados.', 'ok');
      _setProgress(80);

      // ── STEP 5: EXPORT ──────────────────────────────────────
      _setStep(4);
      _log('Compilando binario .xlsx y preparando descarga...');
      await _tick();

      // Extra pause so UI repaints before the heavy binary write
      await new Promise(r => setTimeout(r, 120));

      const buffer   = await wb.xlsx.writeBuffer();
      const filename = `${opts.outputName}_${Date.now()}.xlsx`;
      saveAs(
        new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        filename
      );

      _setProgress(100);
      STEP_IDS.forEach(id => {
        document.getElementById(id).classList.remove('active');
        document.getElementById(id).classList.add('done');
      });
      _log(`Archivo "${filename}" descargado exitosamente.`, 'ok');
      _showResult(true, '✓ Exportación completada',
        `${master.length.toLocaleString('es')} filas en ${rawHeaders.length} columnas → ${filename}`);

    } catch (err) {
      _log(`ERROR: ${err.message}`, 'err');
      _showResult(false, '✕ Error en el pipeline', err.message);
      console.error('[DataForge]', err);
    } finally {
      _running = false;
      processBtn.disabled = FM.get().length === 0;
      clearBtn.disabled   = FM.get().length === 0;
    }
  };

  return { run };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: JsonFlattener
// Robust JSON → flat rows engine. Handles:
//   · Array of objects          [{…},{…}]
//   · Array of arrays           [[…],[…]]
//   · Nested objects (deep)     {a:{b:{c:1}}}
//   · Mixed arrays              [{a:1, b:{c:2}}]
//   · Root object (single row)  {name:"x", val:1}
//   · Array of primitives       [1,2,3]
//   · Multiple root keys        {users:[…], meta:{…}}
// ══════════════════════════════════════════════════════════════
const JsonFlattener = (() => {

  const _isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);
  const _isObjArray = v => Array.isArray(v) && v.length > 0 && v.every(_isObj);

  // Cartesian merge of row lists (parent fields repeated on each child row)
  const _cross = (rows, sub) => rows.flatMap(r => sub.map(s => ({ ...r, ...s })));

  // Flatten one record into 1+ rows with dot-notation keys.
  // The first nested array of objects per record is exploded into rows (e.g. order → items);
  // any further ones are kept as JSON text so rows don't multiply out of control.
  const _expand = (obj, prefix = '', state = { exploded: false }) => {
    let rows = [{}];
    const set = (k, v) => rows.forEach(r => { r[k] = v; });
    for (const [k, v] of Object.entries(obj)) {
      const key = prefix ? `${prefix}.${k}` : k;
      if (v === null || v === undefined) set(key, '');
      else if (_isObjArray(v) && !state.exploded) {
        state.exploded = true;
        rows = _cross(rows, v.flatMap(item => _expand(item, key, state)));
      } else if (Array.isArray(v)) {
        set(key, v.some(x => x !== null && typeof x === 'object') ? JSON.stringify(v) : v.join(' | '));
      } else if (_isObj(v)) rows = _cross(rows, _expand(v, key, state));
      else set(key, v);
    }
    return rows;
  };

  // Header array + body (array of arrays) of the same width, as siblings in one object
  // e.g. { data_header: { main: [...] }, data_body: [[...]] }
  const _findHeaderBodyPair = obj => {
    const stringArrays = [], bodies = [];
    const walk = o => {
      if (Array.isArray(o)) {
        if (o.length && o.every(x => typeof x === 'string')) stringArrays.push(o);
        else if (o.length && o.every(Array.isArray)) bodies.push(o);
        return;
      }
      if (_isObj(o)) Object.values(o).forEach(walk);
    };
    walk(obj);
    for (const body of bodies) {
      const width = Math.max(...body.map(r => r.length));
      const headerArr = stringArrays.find(h => h.length === width);
      if (headerArr) return { headerArr, bodyArr: body };
    }
    return null;
  };

  // Largest array of objects anywhere in the tree (e.g. { status, data: { results: [...] } })
  const _findRecords = obj => {
    let best = null;
    const walk = o => {
      if (_isObjArray(o)) { if (!best || o.length > best.length) best = o; return; }
      if (_isObj(o)) Object.values(o).forEach(walk);
    };
    walk(obj);
    return best;
  };

  // Array → rows
  const _fromArray = arr => {
    if (arr.length === 0) throw new Error('El array JSON está vacío.');
    if (arr.every(Array.isArray)) {
      const first = arr[0];
      const hasHeader = first.length && first.every(x => typeof x === 'string');
      const width = Math.max(...arr.map(r => r.length));
      const headers = Array.from({ length: width }, (_, i) => hasHeader && first[i] ? String(first[i]) : `Col_${i + 1}`);
      return (hasHeader ? arr.slice(1) : arr).map(r =>
        Object.fromEntries(headers.map((h, i) => [h, _cell(r[i])])));
    }
    return arr.flatMap(item => _isObj(item) ? _expand(item)
      : Array.isArray(item) ? [{ value: JSON.stringify(item) }] : [{ value: item ?? '' }]);
  };

  // Cell inside an array row: objects show their most meaningful label field
  const _cell = v => {
    if (v === null || v === undefined) return '';
    if (typeof v !== 'object') return v;
    if (Array.isArray(v)) return v.join(' | ');
    for (const k of ['service_number', 'name', 'label', 'title', 'text', 'value', 'description', 'code'])
      if (v[k] !== undefined && v[k] !== null) return String(v[k]);
    return Object.values(v).find(x => typeof x === 'string' && x.trim()) ?? JSON.stringify(v);
  };

  // Accepts normal JSON or NDJSON / JSON Lines (one object per line)
  const parse = raw => {
    const text = raw.replace(/^\uFEFF/, '').trim();
    try { return JSON.parse(text); } catch (e) {
      const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
      if (lines.length > 1) {
        try { return lines.map(l => JSON.parse(l.replace(/,$/, ''))); } catch (_) {}
      }
      throw e;
    }
  };

  // Convert any JSON value into { headers, rows } ready for Excel
  const flatten = parsed => {
    let rows;
    if (Array.isArray(parsed)) rows = _fromArray(parsed);
    else if (_isObj(parsed)) {
      const pair = _findHeaderBodyPair(parsed);
      const records = _findRecords(parsed);
      if (pair && (!records || pair.bodyArr.length >= records.length)) {
        rows = pair.bodyArr.map(r => Object.fromEntries(pair.headerArr.map((h, i) => [h, _cell(r[i])])));
      } else if (records) {
        // Root-level scalars (e.g. empresa, periodo) repeated as columns on every row
        const rootScalars = Object.fromEntries(Object.entries(parsed).filter(([, v]) => v === null || typeof v !== 'object'));
        rows = _fromArray(records).map(r => ({ ...rootScalars, ...r }));
      } else rows = _expand(parsed);
    } else throw new Error('El JSON debe ser un objeto o un array.');

    if (rows.length === 0) throw new Error('No se encontraron filas de datos en el JSON.');

    // Unified headers (union of all keys, first-seen order); missing cells → ''
    const headers = [...new Set(rows.flatMap(Object.keys))];
    const normalized = rows.map(r => Object.fromEntries(headers.map(h => [h, r[h] ?? ''])));
    return { headers, rows: normalized };
  };

  // Detect JSON structure type for the UI badge
  const describeStructure = parsed => {
    if (Array.isArray(parsed)) {
      if (parsed.length === 0) return 'Array vacío';
      const t = typeof parsed[0];
      if (t === 'object' && !Array.isArray(parsed[0])) return `Array de objetos [${parsed.length}]`;
      if (Array.isArray(parsed[0])) return `Array de arrays [${parsed.length}]`;
      return `Array de ${t}s [${parsed.length}]`;
    }
    if (typeof parsed === 'object' && parsed !== null) {
      const keys = Object.keys(parsed);
      const arrKeys = keys.filter(k => Array.isArray(parsed[k]));
      if (arrKeys.length) return `Objeto con array "${arrKeys[0]}"`;
      return `Objeto plano {${keys.slice(0,3).join(', ')}${keys.length>3?'…':''}}`;
    }
    return typeof parsed;
  };

  return { parse, flatten, describeStructure };
})();


// ══════════════════════════════════════════════════════════════
// MODULE: JsonTab
// UI controller for the JSON → Excel tab
// ══════════════════════════════════════════════════════════════
const JsonTab = (() => {

  let _parsed  = null;
  let _flat    = null;
  let _debounceTimer = null;

  const _el = id => document.getElementById(id);

  // Live parse & preview on textarea input (debounced 400ms)
  const _onInput = () => {
    clearTimeout(_debounceTimer);
    _debounceTimer = setTimeout(_parse, 400);
  };

  const _parse = () => {
    const raw = _el('jsonInput').value.trim();
    const ta  = _el('jsonInput');

    _el('jmRows').textContent  = '—';
    _el('jmCols').textContent  = '—';
    _el('jmType').textContent  = '—';
    _el('jmType').className    = 'jm-val';
    _el('jsonExportBtn').disabled = true;
    _el('jsonPreviewWrap').classList.remove('visible');

    if (!raw) { ta.className = 'json-textarea'; _parsed = null; _flat = null; return; }

    try {
      _parsed = JsonFlattener.parse(raw);
      _flat   = JsonFlattener.flatten(_parsed);

      ta.className = 'json-textarea valid';
      _el('jmType').textContent = JsonFlattener.describeStructure(_parsed);
      _el('jmType').className   = 'jm-val ok';
      _el('jmRows').textContent = _flat.rows.length.toLocaleString('es');
      _el('jmCols').textContent = _flat.headers.length;
      _el('jsonExportBtn').disabled = false;

      _renderPreview();
    } catch (e) {
      ta.className = 'json-textarea invalid';
      _el('jmType').textContent = e instanceof SyntaxError ? `JSON inválido: ${e.message}` : e.message;
      _el('jmType').className   = 'jm-val err';
      _parsed = null; _flat = null;
    }
  };

  const _renderPreview = () => {
    if (!_flat) return;
    const MAX_PREVIEW_ROWS = 50;
    const { headers, rows } = _flat;
    const previewRows = rows.slice(0, MAX_PREVIEW_ROWS);

    const esc = v => String(v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const thead = `<tr>${headers.map(h => `<th title="${esc(h)}">${esc(h)}</th>`).join('')}</tr>`;
    const tbody = previewRows.map(r =>
      `<tr>${headers.map(h => `<td title="${esc(r[h])}">${r[h] === '' ? '<span style="opacity:.3">—</span>' : esc(r[h])}</td>`).join('')}</tr>`
    ).join('');

    _el('previewTable').innerHTML = `<thead>${thead}</thead><tbody>${tbody}</tbody>`;
    _el('previewBadge').textContent =
      rows.length > MAX_PREVIEW_ROWS
        ? `Vista previa — primeras ${MAX_PREVIEW_ROWS} de ${rows.length.toLocaleString('es')} filas`
        : `${rows.length} fila${rows.length !== 1 ? 's' : ''}`;
    _el('jsonPreviewWrap').classList.add('visible');
  };

  const _export = async () => {
    if (!_flat) return;

    const btn = _el('jsonExportBtn');
    btn.disabled = true;
    btn.textContent = '⏳ GENERANDO…';

    try {
      const { headers, rows } = _flat;
      const safeHeaders = DataCleaner.sanitizeHeaders(headers);

      const theme      = _el('jsonTheme').value;
      const accounting = _el('jsonToggleAcct').checked;
      const freeze     = _el('jsonToggleFreeze').checked;
      const outputName = _el('jsonOutputName').value.trim() || 'JSON_Export';

      const wb = new ExcelJS.Workbook();
      wb.creator = 'DataForge Suite v2.0 — JSON Engine';
      wb.created = new Date();

      // If multi-sheet option is on AND JSON root is an object with multiple arrays,
      // create one sheet per key; otherwise single sheet
      const multiSheet = _el('jsonToggleMulti').checked;
      const isRootObj  = !Array.isArray(_parsed) && typeof _parsed === 'object' && _parsed !== null;
      const arrKeys    = isRootObj ? Object.keys(_parsed).filter(k => Array.isArray(_parsed[k]) && _parsed[k].length > 0) : [];

      if (multiSheet && arrKeys.length > 1) {
        for (const key of arrKeys) {
          try {
            const sub = JsonFlattener.flatten(_parsed[key]);
            const safeH = DataCleaner.sanitizeHeaders(sub.headers);
            const sheetName = key.slice(0, 31).replace(/[\\\/\?\*\[\]:]/g, '_') || 'Hoja';
            const ws = wb.addWorksheet(sheetName, {
              properties: { defaultRowHeight: 16 },
              pageSetup: { orientation: 'landscape', fitToPage: true }
            });
            await ExcelGenerator.generate(ws, sub.headers, safeH, sub.rows,
              { theme, freeze, accounting });
          } catch (_) { /* skip malformed sub-arrays */ }
        }
      } else {
        const ws = wb.addWorksheet('Datos JSON', {
          properties: { defaultRowHeight: 16 },
          pageSetup: { orientation: 'landscape', fitToPage: true }
        });
        await ExcelGenerator.generate(ws, headers, safeHeaders, rows,
          { theme, freeze, accounting });
      }

      const buffer   = await wb.xlsx.writeBuffer();
      const filename = `${outputName}_${Date.now()}.xlsx`;
      saveAs(
        new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        filename
      );

      btn.textContent = '✓ DESCARGADO';
      setTimeout(() => { btn.textContent = '▶ EXPORTAR A EXCEL'; btn.disabled = false; }, 2500);

    } catch (err) {
      console.error('[DataForge JSON]', err);
      alert(`No se pudo exportar: ${err.message}`);
      btn.textContent = '✕ ERROR — REINTENTAR';
      btn.disabled = false;
      setTimeout(() => { btn.textContent = '▶ EXPORTAR A EXCEL'; }, 3000);
    }
  };

  const _clearEditor = () => {
    _el('jsonInput').value = '';
    _el('jsonInput').className = 'json-textarea';
    _el('jsonPreviewWrap').classList.remove('visible');
    _el('jmRows').textContent = '—';
    _el('jmCols').textContent = '—';
    _el('jmType').textContent = '—';
    _el('jmType').className = 'jm-val';
    _el('jsonExportBtn').disabled = true;
    _parsed = null; _flat = null;
  };

  const _loadSample = () => {
    const sample = {
      empresa: "DataForge S.A.",
      periodo: "2025-Q1",
      transacciones: [
        { id: 1, fecha: "2025-01-05", concepto: "Pago proveedor A", monto: 150000, estado: "pagado" },
        { id: 2, fecha: "2025-01-12", concepto: "Venta cliente B",  monto: 320000, estado: "cobrado" },
        { id: 3, fecha: "2025-02-03", concepto: "Servicio mensual", monto: 89000,  estado: "pendiente" },
        { id: 4, fecha: "2025-02-18", concepto: "Reembolso gastos", monto: 12500,  estado: "pagado" },
        { id: 5, fecha: "2025-03-01", concepto: "Contrato anual",   monto: 980000, estado: "cobrado" }
      ]
    };
    _el('jsonInput').value = JSON.stringify(sample, null, 2);
    _parse();
  };

  const init = () => {
    _el('jsonInput').addEventListener('input', _onInput);
    _el('jsonExportBtn').addEventListener('click', _export);
    _el('jsonClearBtn').addEventListener('click', _clearEditor);
    _el('jsonSampleBtn').addEventListener('click', _loadSample);
  };

  return { init };
})();


// ══════════════════════════════════════════════════════════════
// EVENT WIRING — DOMContentLoaded
// ══════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  const dropZone  = document.getElementById('dropZone');
  const fileInput = document.getElementById('fileInput');

  // Click on dropzone opens file picker
  dropZone.addEventListener('click', () => fileInput.click());

  // File picker change
  fileInput.addEventListener('change', e => {
    FM.add(Array.from(e.target.files));
    fileInput.value = ''; // allow re-selecting same file
  });

  // Drag & drop
  dropZone.addEventListener('dragover',  e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', ()  => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    FM.add(Array.from(e.dataTransfer.files));
  });

  // Process button
  document.getElementById('processBtn').addEventListener('click', Pipeline.run);

  // Clear button
  document.getElementById('clearBtn').addEventListener('click', () => {
    FM.clear();
    document.getElementById('resultBanner').style.display  = 'none';
    document.getElementById('statsBar').style.display      = 'none';
    document.getElementById('logPanel').innerHTML          = '';
    document.getElementById('logPanel').classList.remove('visible');
    document.getElementById('progressWrap').style.display  = 'none';
    document.getElementById('progressBar').style.width     = '0%';
    ['step-read','step-validate','step-clean','step-format','step-export'].forEach(id => {
      document.getElementById(id).classList.remove('active', 'done');
    });
  });

  // Expose FM.remove globally for inline onclick in file badges
  window.FM = FM;

  // ── Tab switching ──────────────────────────────────────────
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.tab;
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById(target).classList.add('active');
    });
  });

  // ── Init JSON tab module ───────────────────────────────────
  JsonTab.init();
});