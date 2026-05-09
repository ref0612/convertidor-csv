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

  // Parse flat text files (CSV, TSV, TXT, pipe-separated, etc.)
  const _parseText = file => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      const text = e.target.result;
      const delim = detectDelimiter(text);
      Papa.parse(text, {
        header: true,
        skipEmptyLines: 'greedy',
        dynamicTyping: true,
        delimiter: delim,
        complete: res => resolve({ name: file.name, data: res.data, headers: res.meta.fields }),
        error: err => reject(new Error(`PapaParse [${file.name}]: ${err.message}`))
      });
    };
    reader.onerror = () => reject(new Error(`No se pudo leer el archivo: ${file.name}`));
    reader.readAsText(file, 'UTF-8');
  });

  // Parse XLSX via ExcelJS (convertidor-csv repo integration)
  const _parseXLSX = async file => {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.worksheets[0];
    if (!ws) throw new Error(`El archivo "${file.name}" no contiene hojas de cálculo.`);

    const rows = [];
    let headers = [];

    ws.eachRow((row, i) => {
      const vals = row.values.slice(1).map(v => {
        if (v && typeof v === 'object' && v.result !== undefined) return v.result;
        if (v && typeof v === 'object' && v.text !== undefined) return v.text;
        return v ?? null;
      });
      if (i === 1) {
        headers = vals.map(String);
      } else {
        const obj = {};
        headers.forEach((h, idx) => { obj[h] = vals[idx] ?? null; });
        rows.push(obj);
      }
    });

    return { name: file.name, data: rows, headers };
  };

  // Public: route by extension
  const parse = async file => {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext === 'xlsx') return _parseXLSX(file);
    return _parseText(file);
  };

  return { parse };
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

  const _isTrash = row => {
    if (Object.values(row).every(v => v === null || v === '')) return true;
    const str = JSON.stringify(Object.values(row)).toLowerCase();
    return TRASH_PATTERNS.some(p => p.test(str));
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
    const rCL  = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/;   // DD/MM/YYYY
    const rISO = /^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/;    // YYYY-MM-DD

    if (rCL.test(d)) {
      const [, dd, mm, yyyy] = d.match(rCL);
      dt = new Date(+yyyy, +mm - 1, +dd);
    } else if (rISO.test(d)) {
      const [, yyyy, mm, dd] = d.match(rISO);
      dt = new Date(+yyyy, +mm - 1, +dd);
    } else {
      dt = new Date(d);
      if (isNaN(dt.getTime())) return new Date(0);
    }

    // Inject time component
    if (h) {
      const parts = h.split(' ');
      const [hrs, mins] = parts[0].split(':');
      let hh = parseInt(hrs) || 0;
      let mm2 = parseInt(mins) || 0;
      if (parts[1]) {
        const ampm = parts[1].toUpperCase();
        if (ampm === 'PM' && hh < 12) hh += 12;
        if (ampm === 'AM' && hh === 12) hh = 0;
      }
      dt.setHours(hh, mm2);
    }

    return dt;
  };

  const sort = data => [...data].sort((a, b) => _parseDate(a) - _parseDate(b));

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

    const tableRows = dataRows.map(r => rawHeaders.map(h => r[h] ?? ''));

    // Build Excel Table with corporate theme
    ws.addTable({
      name: 'ConsolidadoMaestro',
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

      const refHeaders = JSON.stringify(datasets[0].headers);
      const mismatch = datasets.find(d => JSON.stringify(d.headers) !== refHeaders);
      if (mismatch) {
        throw new Error(
          `Estructura incompatible en "${mismatch.name}". ` +
          `Todos los archivos deben tener exactamente los mismos encabezados.`
        );
      }
      _log(`Estructura validada: ${datasets[0].headers.length} columnas.`, 'ok');
      _setProgress(40);

      // ── STEP 3: CLEAN & SORT ────────────────────────────────
      _setStep(2);
      _log('Uniendo, limpiando y ordenando datos...');
      await _tick();

      const rawHeaders = datasets[0].headers;
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

  // Flatten a single object recursively into dot-notation keys
  const _flattenObj = (obj, prefix = '', result = {}) => {
    for (const [k, v] of Object.entries(obj)) {
      const key = prefix ? `${prefix}.${k}` : k;
      if (v === null || v === undefined) {
        result[key] = '';
      } else if (Array.isArray(v)) {
        if (v.length === 0) {
          result[key] = '';
        } else if (typeof v[0] === 'object' && v[0] !== null) {
          // Nested array of objects: serialize as JSON string (too complex to expand inline)
          result[key] = JSON.stringify(v);
        } else {
          // Array of primitives: join with pipe
          result[key] = v.join(' | ');
        }
      } else if (typeof v === 'object') {
        _flattenObj(v, key, result);
      } else {
        result[key] = v;
      }
    }
    return result;
  };

  // Find the dominant array in a root object (largest array value)
  const _extractDominantArray = obj => {
    let best = null, bestLen = -1;
    for (const v of Object.values(obj)) {
      if (Array.isArray(v) && v.length > bestLen) { best = v; bestLen = v.length; }
    }
    return best;
  };

  // Serialize a cell value: if it's an object, extract the most meaningful string field
  const _serializeCell = v => {
    if (v === null || v === undefined) return '';
    if (typeof v !== 'object') return v;
    if (Array.isArray(v)) return v.join(' | ');
    // Object cell: prefer common "label" keys, fallback to first string value
    const labelKeys = ['service_number','name','label','title','text','value','description','code'];
    for (const k of labelKeys) {
      if (v[k] !== undefined && v[k] !== null) return String(v[k]);
    }
    // Fallback: first string-valued key
    for (const val of Object.values(v)) {
      if (typeof val === 'string' && val.trim()) return val;
    }
    return JSON.stringify(v);
  };

  // Deep search for a pair: headerArray + dataBody (array of arrays) anywhere in the object tree
  const _findHeaderBodyPair = obj => {
    if (typeof obj !== 'object' || obj === null) return null;

    // Direct pattern: { data_header: { main: [...] }, data_body: [[...]] }
    const allKeys = Object.keys(obj);
    let headerArr = null, bodyArr = null;

    // Search for header array (array of strings)
    const findHeaderArray = o => {
      if (Array.isArray(o) && o.length > 0 && typeof o[0] === 'string') return o;
      if (typeof o === 'object' && o !== null) {
        for (const v of Object.values(o)) {
          const r = findHeaderArray(v);
          if (r) return r;
        }
      }
      return null;
    };

    // Search for data body (array of arrays of non-string values)
    const findBodyArray = o => {
      if (Array.isArray(o) && o.length > 0 && Array.isArray(o[0])) return o;
      if (typeof o === 'object' && o !== null) {
        for (const v of Object.values(o)) {
          const r = findBodyArray(v);
          if (r) return r;
        }
      }
      return null;
    };

    headerArr = findHeaderArray(obj);
    bodyArr   = findBodyArray(obj);

    if (headerArr && bodyArr) return { headerArr, bodyArr };
    return null;
  };

  // Convert any JSON value into { headers, rows } ready for Excel
  const flatten = (parsed) => {
    let rows = [];

    // Case 0: header_array + data_body pattern (e.g. API reports with data_header.main + data_body)
    if (typeof parsed === 'object' && parsed !== null && !Array.isArray(parsed)) {
      const pair = _findHeaderBodyPair(parsed);
      if (pair) {
        const { headerArr, bodyArr } = pair;
        rows = bodyArr.map(row =>
          Object.fromEntries(headerArr.map((h, i) => [h, _serializeCell(row[i])]))
        );
        // fall through to header/normalize logic below
        if (rows.length === 0) throw new Error('No se encontraron filas de datos en el JSON.');
        const headers = headerArr;
        const normalized = rows.map(r => {
          const out = {};
          headers.forEach(h => { out[h] = r[h] !== undefined ? r[h] : ''; });
          return out;
        });
        return { headers, rows: normalized };
      }
    }

    // Case 1: top-level array
    if (Array.isArray(parsed)) {
      if (parsed.length === 0) throw new Error('El array JSON está vacío.');

      if (typeof parsed[0] === 'object' && parsed[0] !== null && !Array.isArray(parsed[0])) {
        // Array of objects → flatten each
        rows = parsed.map(item => _flattenObj(item));
      } else if (Array.isArray(parsed[0])) {
        // Array of arrays → first row = headers
        const [headerRow, ...dataRows] = parsed;
        rows = dataRows.map(r => {
          const obj = {};
          headerRow.forEach((h, i) => { obj[String(h)] = _serializeCell(r[i]); });
          return obj;
        });
      } else {
        // Array of primitives → single column "value"
        rows = parsed.map((v, i) => ({ index: i + 1, value: v }));
      }

    // Case 2: top-level object
    } else if (typeof parsed === 'object' && parsed !== null) {
      const dominant = _extractDominantArray(parsed);
      if (dominant && dominant.length > 0 && typeof dominant[0] === 'object') {
        // Has a dominant nested array → use it, merge scalar root keys as extra columns
        const rootScalars = {};
        for (const [k, v] of Object.entries(parsed)) {
          if (!Array.isArray(v) && typeof v !== 'object') rootScalars[k] = v;
        }
        rows = dominant.map(item => ({ ...rootScalars, ..._flattenObj(item) }));
      } else {
        // Single object → one row
        rows = [_flattenObj(parsed)];
      }

    } else {
      throw new Error('El JSON debe ser un objeto o un array.');
    }

    if (rows.length === 0) throw new Error('No se encontraron filas de datos en el JSON.');

    // Build unified headers (union of all keys, preserving first-seen order)
    const headerSet = new Map();
    for (const row of rows) {
      for (const k of Object.keys(row)) {
        if (!headerSet.has(k)) headerSet.set(k, true);
      }
    }
    const headers = [...headerSet.keys()];

    // Normalize all rows to have every header
    const normalized = rows.map(r => {
      const out = {};
      headers.forEach(h => { out[h] = r[h] !== undefined ? r[h] : ''; });
      return out;
    });

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

  return { flatten, describeStructure };
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
      _parsed = JSON.parse(raw);
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
      _el('jmType').textContent = 'JSON inválido';
      _el('jmType').className   = 'jm-val err';
      _parsed = null; _flat = null;
    }
  };

  const _renderPreview = () => {
    if (!_flat) return;
    const MAX_PREVIEW_ROWS = 50;
    const { headers, rows } = _flat;
    const previewRows = rows.slice(0, MAX_PREVIEW_ROWS);

    const thead = `<tr>${headers.map(h => `<th title="${h}">${h}</th>`).join('')}</tr>`;
    const tbody = previewRows.map(r =>
      `<tr>${headers.map(h => `<td title="${r[h]}">${r[h] === '' ? '<span style="opacity:.3">—</span>' : r[h]}</td>`).join('')}</tr>`
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
            const sheetName = key.slice(0, 31).replace(/[\\\/\?\*\[\]]/g, '_');
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
})();