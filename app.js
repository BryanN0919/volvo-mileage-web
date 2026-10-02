(() => {
  'use strict';

  const REQUIRED_HEADERS = [
    'Started',
    'Start odometer (miles)',
    'Start address',
    'End odometer (miles)',
    'End address'
  ];

  const FORM_SHEET_NAME = 'Mileage Log';
  const FIRST_DATA_ROW = 9;
  const TEMPLATE_LAST_DATA_ROW = 72;
  const TEMPLATE_TOTALS_ROW = 73;
  const TEMPLATE_PRINT_END_ROW = 80;
  const TEMPLATE_LAST_ROW = 87;

  const els = {
    fileInput: document.getElementById('fileInput'),
    chooseButton: document.getElementById('chooseButton'),
    dropZone: document.getElementById('dropZone'),
    fileStatus: document.getElementById('fileStatus'),
    fileName: document.getElementById('fileName'),
    fileDetails: document.getElementById('fileDetails'),
    changeFile: document.getElementById('changeFile'),
    message: document.getElementById('message'),
    resultsSection: document.getElementById('resultsSection'),
    tripCount: document.getElementById('tripCount'),
    excludedCount: document.getElementById('excludedCount'),
    totalMiles: document.getElementById('totalMiles'),
    rateInput: document.getElementById('rateInput'),
    totalPay: document.getElementById('totalPay'),
    detectedMonth: document.getElementById('detectedMonth'),
    tripTableBody: document.getElementById('tripTableBody'),
    masterCheckbox: document.getElementById('masterCheckbox'),
    selectAll: document.getElementById('selectAll'),
    clearAll: document.getElementById('clearAll'),
    generateButton: document.getElementById('generateButton'),
    loadingOverlay: document.getElementById('loadingOverlay'),
    loadingText: document.getElementById('loadingText'),
    stepImport: document.getElementById('stepImport'),
    stepReview: document.getElementById('stepReview'),
    stepGenerate: document.getElementById('stepGenerate')
  };

  let state = {
    file: null,
    trips: [],
    excluded: 0,
    detectedDate: null,
    rate: 0.625
  };

  function setLoading(show, text = 'Processing trip log…') {
    els.loadingText.textContent = text;
    els.loadingOverlay.classList.toggle('hidden', !show);
  }

  function showMessage(text, type = 'error') {
    els.message.textContent = text;
    els.message.className = `message ${type}`;
  }

  function clearMessage() {
    els.message.className = 'message hidden';
    els.message.textContent = '';
  }

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} bytes`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function monthLabel(date) {
    return date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  }

  function shortDate(date) {
    return date.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
  }

  function currency(value) {
    return `$${value.toFixed(3)}`;
  }

  function selectedTrips() {
    return state.trips.filter(t => t.included);
  }

  function escapeHtml(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) {
      throw new Error('An Excel XML file could not be read.');
    }
    return doc;
  }

  function serializeXml(doc) {
    return new XMLSerializer().serializeToString(doc);
  }

  function directChild(parent, tagName) {
    for (const child of parent.children) {
      if (child.localName === tagName) return child;
    }
    return null;
  }

  function createElementLike(doc, parent, localName) {
    const namespace = parent.namespaceURI || doc.documentElement.namespaceURI;
    return doc.createElementNS(namespace, localName);
  }

  function excelCellColumn(cellRef) {
    return String(cellRef || '').replace(/\d+/g, '');
  }

  function columnToNumber(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  function excelSerialToDate(value) {
    const serial = Number(value);
    if (!Number.isFinite(serial)) return null;
    const epoch = Date.UTC(1899, 11, 30);
    return new Date(epoch + serial * 86400000);
  }

  function dateToExcelSerial(date) {
    const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
    return (utc - Date.UTC(1899, 11, 30)) / 86400000;
  }

  function parseStarted(value) {
    if (value instanceof Date) return value;
    if (typeof value === 'number') return excelSerialToDate(value);
    const text = String(value ?? '').trim();
    if (!text) return null;
    const m = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) {
      return new Date(
        Number(m[1]), Number(m[2]) - 1, Number(m[3]),
        Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)
      );
    }
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function cellValue(cell, sharedStrings) {
    if (!cell) return '';
    const type = cell.getAttribute('t') || '';
    if (type === 'inlineStr') {
      const isNode = directChild(cell, 'is');
      if (!isNode) return '';
      return Array.from(isNode.getElementsByTagNameNS('*', 't')).map(n => n.textContent || '').join('');
    }
    const v = directChild(cell, 'v');
    const raw = v ? v.textContent : '';
    if (type === 's') return sharedStrings[Number(raw)] ?? '';
    if (type === 'str') return raw;
    if (type === 'b') return raw === '1';
    if (raw === '') return '';
    const numeric = Number(raw);
    return Number.isFinite(numeric) ? numeric : raw;
  }

  async function readSharedStrings(zip) {
    const file = zip.file('xl/sharedStrings.xml');
    if (!file) return [];
    const doc = parseXml(await file.async('string'));
    return Array.from(doc.getElementsByTagNameNS('*', 'si')).map(si =>
      Array.from(si.getElementsByTagNameNS('*', 't')).map(t => t.textContent || '').join('')
    );
  }

  function normalizeTarget(target) {
    const clean = String(target || '').replace(/^\//, '');
    return clean.startsWith('xl/') ? clean : `xl/${clean}`;
  }

  async function findWorksheetPath(zip, desiredName) {
    const workbookFile = zip.file('xl/workbook.xml');
    const relFile = zip.file('xl/_rels/workbook.xml.rels');
    if (!workbookFile || !relFile) throw new Error('This file is not a standard Excel .xlsx workbook.');

    const wbDoc = parseXml(await workbookFile.async('string'));
    const relDoc = parseXml(await relFile.async('string'));
    const relationshipMap = new Map();
    for (const rel of Array.from(relDoc.getElementsByTagNameNS('*', 'Relationship'))) {
      relationshipMap.set(rel.getAttribute('Id'), rel.getAttribute('Target'));
    }

    const sheets = Array.from(wbDoc.getElementsByTagNameNS('*', 'sheet'));
    const chosen = sheets.find(s => s.getAttribute('name') === desiredName) || sheets[0];
    if (!chosen) throw new Error('No worksheets were found in the Excel workbook.');

    const relationshipId = chosen.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || chosen.getAttribute('r:id');
    const target = relationshipMap.get(relationshipId);
    if (!target) throw new Error('The workbook worksheet relationship could not be resolved.');
    return normalizeTarget(target);
  }

  function rowCellsByHeader(sheetDoc, sharedStrings) {
    const rows = Array.from(sheetDoc.getElementsByTagNameNS('*', 'row'));
    const firstRow = rows.find(r => r.getAttribute('r') === '1') || rows[0];
    if (!firstRow) throw new Error('The Volvo workbook has no header row.');

    const map = new Map();
    for (const cell of Array.from(firstRow.getElementsByTagNameNS('*', 'c'))) {
      const header = String(cellValue(cell, sharedStrings) ?? '').trim();
      if (header) map.set(header, excelCellColumn(cell.getAttribute('r')));
    }
    return map;
  }

  function findCellInRow(row, column) {
    return Array.from(row.getElementsByTagNameNS('*', 'c')).find(c => excelCellColumn(c.getAttribute('r')) === column) || null;
  }

  async function parseVolvoWorkbook(file) {
    const buffer = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const sharedStrings = await readSharedStrings(zip);
    const sheetPath = await findWorksheetPath(zip, 'Sheet1');
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) throw new Error('The Volvo worksheet could not be found.');
    const sheetDoc = parseXml(await sheetFile.async('string'));
    const headers = rowCellsByHeader(sheetDoc, sharedStrings);

    const missing = REQUIRED_HEADERS.filter(h => !headers.has(h));
    if (missing.length) {
      throw new Error(`Missing required Volvo column${missing.length > 1 ? 's' : ''}: ${missing.join(', ')}`);
    }

    const trips = [];
    let excluded = 0;
    const rows = Array.from(sheetDoc.getElementsByTagNameNS('*', 'row')).filter(r => Number(r.getAttribute('r')) >= 2);

    for (const row of rows) {
      const startOdo = Number(cellValue(findCellInRow(row, headers.get('Start odometer (miles)')), sharedStrings));
      const endOdo = Number(cellValue(findCellInRow(row, headers.get('End odometer (miles)')), sharedStrings));
      const miles = endOdo - startOdo;
      if (!Number.isFinite(startOdo) || !Number.isFinite(endOdo) || !(miles > 0.000001)) {
        excluded++;
        continue;
      }

      const startedValue = cellValue(findCellInRow(row, headers.get('Started')), sharedStrings);
      const started = parseStarted(startedValue);
      if (!started || Number.isNaN(started.getTime())) {
        excluded++;
        continue;
      }

      trips.push({
        id: `trip-${row.getAttribute('r')}`,
        sourceRow: Number(row.getAttribute('r')),
        date: new Date(started.getFullYear(), started.getMonth(), started.getDate()),
        started,
        startOdo,
        endOdo,
        startAddress: String(cellValue(findCellInRow(row, headers.get('Start address')), sharedStrings) ?? '').trim(),
        endAddress: String(cellValue(findCellInRow(row, headers.get('End address')), sharedStrings) ?? '').trim(),
        miles,
        included: true
      });
    }

    trips.sort((a, b) => b.started.getTime() - a.started.getTime());
    return { trips, excluded };
  }

  function recalculate() {
    state.rate = Math.max(0, Number(els.rateInput.value) || 0);
    const selected = selectedTrips();
    const miles = selected.reduce((sum, t) => sum + t.miles, 0);
    const pay = miles * state.rate;

    els.tripCount.textContent = String(selected.length);
    els.excludedCount.textContent = `${state.excluded + (state.trips.length - selected.length)} excluded`;
    els.totalMiles.textContent = miles.toFixed(1);
    els.totalPay.textContent = currency(pay);

    state.trips.forEach(t => {
      const payCell = document.querySelector(`[data-pay-id="${t.id}"]`);
      if (payCell) payCell.textContent = currency(t.miles * state.rate);
    });

    const checked = selected.length;
    els.masterCheckbox.checked = state.trips.length > 0 && checked === state.trips.length;
    els.masterCheckbox.indeterminate = checked > 0 && checked < state.trips.length;
    els.generateButton.disabled = checked === 0;

    if (checked > 0 && state.file) {
      clearMessage();
    }
  }

  function renderTrips() {
    els.tripTableBody.innerHTML = '';
    const fragment = document.createDocumentFragment();
    state.trips.forEach(t => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="check-col"><input type="checkbox" data-trip-id="${t.id}" ${t.included ? 'checked' : ''} aria-label="Include trip ${shortDate(t.date)}" /></td>
        <td class="trip-date">${shortDate(t.date)}</td>
        <td class="address" title="${escapeHtml(t.startAddress)}">${escapeHtml(t.startAddress)}</td>
        <td class="address" title="${escapeHtml(t.endAddress)}">${escapeHtml(t.endAddress)}</td>
        <td class="number-col">${Math.round(t.startOdo).toLocaleString()}</td>
        <td class="number-col">${Math.round(t.endOdo).toLocaleString()}</td>
        <td class="number-col"><strong>${t.miles.toFixed(1)}</strong></td>
        <td class="number-col" data-pay-id="${t.id}">${currency(t.miles * state.rate)}</td>`;
      fragment.appendChild(tr);
    });
    els.tripTableBody.appendChild(fragment);

    els.tripTableBody.querySelectorAll('input[type="checkbox"]').forEach(box => {
      box.addEventListener('change', event => {
        const trip = state.trips.find(t => t.id === event.target.dataset.tripId);
        if (trip) trip.included = event.target.checked;
        recalculate();
      });
    });
  }

  function updateSteps(stage) {
    [els.stepImport, els.stepReview, els.stepGenerate].forEach(el => el.classList.remove('active', 'complete'));
    if (stage === 'import') {
      els.stepImport.classList.add('active');
    } else if (stage === 'review') {
      els.stepImport.classList.add('complete');
      els.stepReview.classList.add('active');
    } else {
      els.stepImport.classList.add('complete');
      els.stepReview.classList.add('complete');
      els.stepGenerate.classList.add('active');
    }
  }

  function resetResults() {
    state.file = null;
    state.trips = [];
    state.excluded = 0;
    state.detectedDate = null;
    els.resultsSection.classList.add('hidden');
    els.fileStatus.classList.add('hidden');
    els.dropZone.classList.remove('hidden');
    els.fileInput.value = '';
    updateSteps('import');
  }

  async function handleFile(file) {
    clearMessage();
    if (!file) return;
    if (!/\.xls(x|m)$/i.test(file.name)) {
      showMessage('Please select an Excel .xlsx or .xlsm Volvo trip log.');
      return;
    }

    setLoading(true, 'Reading Volvo trip log…');
    try {
      const { trips, excluded } = await parseVolvoWorkbook(file);
      if (!trips.length) throw new Error('No trips with positive mileage were found.');

      state.file = file;
      state.trips = trips;
      state.excluded = excluded;
      state.detectedDate = trips[0].date;

      els.fileName.textContent = file.name;
      els.fileDetails.textContent = `${formatBytes(file.size)} • ${trips.length} eligible trips found`;
      els.fileStatus.classList.remove('hidden');
      els.dropZone.classList.add('hidden');
      els.resultsSection.classList.remove('hidden');
      els.detectedMonth.textContent = monthLabel(state.detectedDate);
      renderTrips();
      recalculate();
      updateSteps('review');
      showMessage(`Loaded ${trips.length} reimbursable trips for ${monthLabel(state.detectedDate)}. The Excel form will resize automatically.`, 'success');
    } catch (err) {
      console.error(err);
      resetResults();
      showMessage(err.message || 'Unable to read this Excel file.');
    } finally {
      setLoading(false);
    }
  }

  function clearCellContents(cell) {
    if (!cell) return;
    for (const child of Array.from(cell.children)) cell.removeChild(child);
    cell.removeAttribute('t');
  }

  function getOrCreateCell(doc, sheetData, ref) {
    let cell = Array.from(sheetData.getElementsByTagNameNS('*', 'c')).find(c => c.getAttribute('r') === ref);
    if (cell) return cell;

    const rowNumber = Number(ref.match(/\d+/)[0]);
    const col = excelCellColumn(ref);
    let row = Array.from(sheetData.getElementsByTagNameNS('*', 'row')).find(r => Number(r.getAttribute('r')) === rowNumber);
    if (!row) {
      row = createElementLike(doc, sheetData, 'row');
      row.setAttribute('r', String(rowNumber));
      const rows = Array.from(sheetData.getElementsByTagNameNS('*', 'row'));
      const next = rows.find(r => Number(r.getAttribute('r')) > rowNumber);
      if (next) sheetData.insertBefore(row, next); else sheetData.appendChild(row);
    }

    cell = createElementLike(doc, row, 'c');
    cell.setAttribute('r', ref);
    const colNumber = columnToNumber(col);
    const existing = Array.from(row.getElementsByTagNameNS('*', 'c'));
    const nextCell = existing.find(c => columnToNumber(excelCellColumn(c.getAttribute('r'))) > colNumber);
    if (nextCell) row.insertBefore(cell, nextCell); else row.appendChild(cell);
    return cell;
  }

  function setNumber(doc, sheetData, ref, value) {
    const cell = getOrCreateCell(doc, sheetData, ref);
    clearCellContents(cell);
    const v = createElementLike(doc, cell, 'v');
    v.textContent = String(value);
    cell.appendChild(v);
  }

  function setInlineString(doc, sheetData, ref, value) {
    const cell = getOrCreateCell(doc, sheetData, ref);
    clearCellContents(cell);
    cell.setAttribute('t', 'inlineStr');
    const isNode = createElementLike(doc, cell, 'is');
    const t = createElementLike(doc, isNode, 't');
    const text = String(value ?? '');
    if (/^\s|\s$/.test(text)) t.setAttribute('xml:space', 'preserve');
    t.textContent = text;
    isNode.appendChild(t);
    cell.appendChild(isNode);
  }

  function setFormula(doc, sheetData, ref, formula, cachedValue) {
    const cell = getOrCreateCell(doc, sheetData, ref);
    clearCellContents(cell);
    const f = createElementLike(doc, cell, 'f');
    f.textContent = formula;
    cell.appendChild(f);
    const v = createElementLike(doc, cell, 'v');
    v.textContent = String(cachedValue);
    cell.appendChild(v);
  }

  function clearRange(doc, sheetData, columns, startRow, endRow) {
    for (let row = startRow; row <= endRow; row++) {
      for (const col of columns) {
        const cell = getOrCreateCell(doc, sheetData, `${col}${row}`);
        clearCellContents(cell);
      }
    }
  }

  function shiftRowElement(row, newRowNumber) {
    row.setAttribute('r', String(newRowNumber));
    for (const cell of Array.from(row.getElementsByTagNameNS('*', 'c'))) {
      const ref = cell.getAttribute('r');
      if (!ref) continue;
      cell.setAttribute('r', `${excelCellColumn(ref)}${newRowNumber}`);
    }
  }

  function cloneBlankDataRow(templateRow, rowNumber) {
    const row = templateRow.cloneNode(true);
    shiftRowElement(row, rowNumber);
    for (const cell of Array.from(row.getElementsByTagNameNS('*', 'c'))) {
      clearCellContents(cell);
    }
    return row;
  }

  function resizeMileageSection(sheetDoc, sheetData, tripCount) {
    const delta = tripCount - (TEMPLATE_LAST_DATA_ROW - FIRST_DATA_ROW + 1);
    const originalRows = Array.from(sheetData.getElementsByTagNameNS('*', 'row'));
    const templateDataRow = originalRows.find(row => Number(row.getAttribute('r')) === TEMPLATE_LAST_DATA_ROW)
      || originalRows.find(row => Number(row.getAttribute('r')) === FIRST_DATA_ROW);
    if (!templateDataRow) throw new Error('The mileage template is missing its trip-row formatting.');

    const footerRows = originalRows
      .filter(row => Number(row.getAttribute('r')) >= TEMPLATE_TOTALS_ROW)
      .map(row => row.cloneNode(true));

    // Remove the template's fixed trip section and footer. They are rebuilt below.
    for (const row of originalRows) {
      if (Number(row.getAttribute('r')) >= FIRST_DATA_ROW) sheetData.removeChild(row);
    }

    // Create exactly as many formatted trip rows as are needed.
    for (let i = 0; i < tripCount; i++) {
      sheetData.appendChild(cloneBlankDataRow(templateDataRow, FIRST_DATA_ROW + i));
    }

    // Move totals/signature/instructions down or up with the resized trip section.
    for (const row of footerRows) {
      shiftRowElement(row, Number(row.getAttribute('r')) + delta);
      sheetData.appendChild(row);
    }

    const mergeCells = sheetDoc.getElementsByTagNameNS('*', 'mergeCells')[0];
    if (mergeCells) {
      for (const merge of Array.from(mergeCells.getElementsByTagNameNS('*', 'mergeCell'))) {
        const ref = merge.getAttribute('ref') || '';
        const dataMerge = /^B(\d+):C\1$/.exec(ref);
        if (dataMerge) {
          const rowNum = Number(dataMerge[1]);
          if (rowNum >= FIRST_DATA_ROW && rowNum <= TEMPLATE_LAST_DATA_ROW) {
            mergeCells.removeChild(merge);
            continue;
          }
        }

        const rangeMatch = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
        if (rangeMatch) {
          const startRow = Number(rangeMatch[2]);
          const endRow = Number(rangeMatch[4]);
          if (startRow >= TEMPLATE_TOTALS_ROW) {
            merge.setAttribute('ref', `${rangeMatch[1]}${startRow + delta}:${rangeMatch[3]}${endRow + delta}`);
          }
        }
      }

      for (let row = FIRST_DATA_ROW; row < FIRST_DATA_ROW + tripCount; row++) {
        const merge = createElementLike(sheetDoc, mergeCells, 'mergeCell');
        merge.setAttribute('ref', `B${row}:C${row}`);
        mergeCells.appendChild(merge);
      }
      mergeCells.setAttribute('count', String(mergeCells.getElementsByTagNameNS('*', 'mergeCell').length));
    }

    const lastRow = TEMPLATE_LAST_ROW + delta;
    const dimension = sheetDoc.getElementsByTagNameNS('*', 'dimension')[0];
    if (dimension) dimension.setAttribute('ref', `B1:P${lastRow}`);

    return {
      dataLastRow: FIRST_DATA_ROW + tripCount - 1,
      totalsRow: TEMPLATE_TOTALS_ROW + delta,
      footerDateRow: 78 + delta,
      printEndRow: TEMPLATE_PRINT_END_ROW + delta,
      lastRow
    };
  }

  function updatePrintArea(workbookDoc, printEndRow) {
    const names = Array.from(workbookDoc.getElementsByTagNameNS('*', 'definedName'));
    const printArea = names.find(node => node.getAttribute('name') === '_xlnm.Print_Area' && node.getAttribute('localSheetId') === '0');
    if (printArea) printArea.textContent = `'${FORM_SHEET_NAME}'!$A$1:$M$${printEndRow}`;
  }

  function endOfMonth(date) {
    return new Date(date.getFullYear(), date.getMonth() + 1, 0);
  }

  function safeFilenamePart(text) {
    return text.replace(/[\\/:*?"<>|]/g, '-');
  }

  function updateCalculationMode(workbookDoc) {
    let calcPr = workbookDoc.getElementsByTagNameNS('*', 'calcPr')[0];
    if (!calcPr) {
      calcPr = createElementLike(workbookDoc, workbookDoc.documentElement, 'calcPr');
      workbookDoc.documentElement.appendChild(calcPr);
    }
    calcPr.setAttribute('calcMode', 'auto');
    calcPr.setAttribute('fullCalcOnLoad', '1');
    calcPr.setAttribute('forceFullCalc', '1');
  }

  async function removeCalculationChain(zip) {
    zip.remove('xl/calcChain.xml');

    const relPath = 'xl/_rels/workbook.xml.rels';
    const relFile = zip.file(relPath);
    if (relFile) {
      const relDoc = parseXml(await relFile.async('string'));
      for (const rel of Array.from(relDoc.getElementsByTagNameNS('*', 'Relationship'))) {
        if ((rel.getAttribute('Type') || '').endsWith('/calcChain')) rel.parentNode.removeChild(rel);
      }
      zip.file(relPath, serializeXml(relDoc));
    }

    const contentFile = zip.file('[Content_Types].xml');
    if (contentFile) {
      const contentDoc = parseXml(await contentFile.async('string'));
      for (const override of Array.from(contentDoc.getElementsByTagNameNS('*', 'Override'))) {
        if (override.getAttribute('PartName') === '/xl/calcChain.xml') override.parentNode.removeChild(override);
      }
      zip.file('[Content_Types].xml', serializeXml(contentDoc));
    }
  }

  async function loadTemplateRate() {
    try {
      const zip = await JSZip.loadAsync(window.MILEAGE_TEMPLATE_BASE64, { base64: true });
      const strings = await readSharedStrings(zip);
      const path = await findWorksheetPath(zip, FORM_SHEET_NAME);
      const doc = parseXml(await zip.file(path).async('string'));
      const cell = Array.from(doc.getElementsByTagNameNS('*', 'c')).find(c => c.getAttribute('r') === 'J4');
      const value = Number(cellValue(cell, strings));
      if (Number.isFinite(value) && value > 0) {
        state.rate = value;
        els.rateInput.value = value.toFixed(3);
      }
    } catch (err) {
      console.warn('Could not read template rate; using 0.625.', err);
    }
  }

  async function generateWorkbook() {
    const trips = selectedTrips();
    if (!trips.length) return;

    setLoading(true, 'Building reimbursement workbook…');
    updateSteps('generate');
    try {
      const zip = await JSZip.loadAsync(window.MILEAGE_TEMPLATE_BASE64, { base64: true });
      const sheetPath = await findWorksheetPath(zip, FORM_SHEET_NAME);
      const sheetFile = zip.file(sheetPath);
      if (!sheetFile) throw new Error('The embedded mileage template could not be opened.');

      const sheetDoc = parseXml(await sheetFile.async('string'));
      const sheetData = sheetDoc.getElementsByTagNameNS('*', 'sheetData')[0];
      if (!sheetData) throw new Error('The mileage template is missing its worksheet data area.');

      const layout = resizeMileageSection(sheetDoc, sheetData, trips.length);

      let totalMiles = 0;
      trips.forEach((trip, index) => {
        const row = FIRST_DATA_ROW + index;
        const pay = trip.miles * state.rate;
        totalMiles += trip.miles;
        setNumber(sheetDoc, sheetData, `B${row}`, dateToExcelSerial(trip.date));
        setNumber(sheetDoc, sheetData, `D${row}`, trip.startOdo);
        setNumber(sheetDoc, sheetData, `E${row}`, trip.endOdo);
        setInlineString(sheetDoc, sheetData, `F${row}`, trip.startAddress);
        setInlineString(sheetDoc, sheetData, `G${row}`, trip.endAddress);
        setFormula(sheetDoc, sheetData, `K${row}`, `E${row}-D${row}`, trip.miles);
        setFormula(sheetDoc, sheetData, `L${row}`, `K${row}*$J$4`, pay);
      });

      const totalPay = totalMiles * state.rate;
      const monthDate = trips[0].date;
      const monthName = monthLabel(monthDate);
      setNumber(sheetDoc, sheetData, 'D3', dateToExcelSerial(endOfMonth(monthDate)));
      setInlineString(sheetDoc, sheetData, 'J3', monthName);
      setNumber(sheetDoc, sheetData, 'J4', state.rate);
      setInlineString(sheetDoc, sheetData, `J${layout.totalsRow}`, 'Totals');
      setFormula(sheetDoc, sheetData, `K${layout.totalsRow}`, `SUM(K${FIRST_DATA_ROW}:K${layout.dataLastRow})`, totalMiles);
      setFormula(sheetDoc, sheetData, `L${layout.totalsRow}`, `SUM(L${FIRST_DATA_ROW}:L${layout.dataLastRow})`, totalPay);
      setFormula(sheetDoc, sheetData, 'J5', `K${layout.totalsRow}`, totalMiles);
      setFormula(sheetDoc, sheetData, 'J6', `L${layout.totalsRow}`, totalPay);
      setNumber(sheetDoc, sheetData, `F${layout.footerDateRow}`, dateToExcelSerial(endOfMonth(monthDate)));

      zip.file(sheetPath, serializeXml(sheetDoc));

      const workbookFile = zip.file('xl/workbook.xml');
      if (workbookFile) {
        const workbookDoc = parseXml(await workbookFile.async('string'));
        updateCalculationMode(workbookDoc);
        updatePrintArea(workbookDoc, layout.printEndRow);
        zip.file('xl/workbook.xml', serializeXml(workbookDoc));
      }
      await removeCalculationChain(zip);

      const blob = await zip.generateAsync({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `G.G. ${safeFilenamePart(monthName)} Mileage - Generated.xlsx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      showMessage(`Generated ${monthName}: ${trips.length} trips, ${totalMiles.toFixed(1)} miles, ${currency(totalPay)} reimbursement.`, 'success');
    } catch (err) {
      console.error(err);
      showMessage(err.message || 'Could not generate the reimbursement form.');
      updateSteps('review');
    } finally {
      setLoading(false);
    }
  }

  els.chooseButton.addEventListener('click', event => {
    event.preventDefault();
    els.fileInput.click();
  });
  els.fileInput.addEventListener('change', event => handleFile(event.target.files[0]));
  els.changeFile.addEventListener('click', () => {
    resetResults();
    clearMessage();
    els.fileInput.click();
  });
  els.dropZone.addEventListener('dragover', event => {
    event.preventDefault();
    els.dropZone.classList.add('dragover');
  });
  els.dropZone.addEventListener('dragleave', () => els.dropZone.classList.remove('dragover'));
  els.dropZone.addEventListener('drop', event => {
    event.preventDefault();
    els.dropZone.classList.remove('dragover');
    handleFile(event.dataTransfer.files[0]);
  });
  els.rateInput.addEventListener('input', recalculate);
  els.masterCheckbox.addEventListener('change', () => {
    state.trips.forEach(t => t.included = els.masterCheckbox.checked);
    renderTrips();
    recalculate();
  });
  els.selectAll.addEventListener('click', () => {
    state.trips.forEach(t => t.included = true);
    renderTrips();
    recalculate();
  });
  els.clearAll.addEventListener('click', () => {
    state.trips.forEach(t => t.included = false);
    renderTrips();
    recalculate();
  });
  els.generateButton.addEventListener('click', generateWorkbook);

  window.addEventListener('load', async () => {
    updateSteps('import');
    if (!window.JSZip) {
      showMessage('The local Excel processing library could not be loaded.');
      return;
    }
    await loadTemplateRate();
  });
})();

// (() => {
//   'use strict';

//   const REQUIRED_HEADERS = [
//     'Category',
//     'Started',
//     'Start odometer (miles)',
//     'Start address',
//     'End odometer (miles)',
//     'End address'
//   ];

//   const FORM_SHEET_NAME = '2022 Mileage Log';
//   const FIRST_DATA_ROW = 9;
//   const TEMPLATE_LAST_DATA_ROW = 72;
//   const TEMPLATE_TOTALS_ROW = 73;
//   const TEMPLATE_PRINT_END_ROW = 80;
//   const TEMPLATE_LAST_ROW = 87;

//   const els = {
//     fileInput: document.getElementById('fileInput'),
//     chooseButton: document.getElementById('chooseButton'),
//     dropZone: document.getElementById('dropZone'),
//     fileStatus: document.getElementById('fileStatus'),
//     fileName: document.getElementById('fileName'),
//     fileDetails: document.getElementById('fileDetails'),
//     changeFile: document.getElementById('changeFile'),
//     message: document.getElementById('message'),
//     resultsSection: document.getElementById('resultsSection'),
//     tripCount: document.getElementById('tripCount'),
//     excludedCount: document.getElementById('excludedCount'),
//     totalMiles: document.getElementById('totalMiles'),
//     rateInput: document.getElementById('rateInput'),
//     totalPay: document.getElementById('totalPay'),
//     detectedMonth: document.getElementById('detectedMonth'),
//     tripTableBody: document.getElementById('tripTableBody'),
//     masterCheckbox: document.getElementById('masterCheckbox'),
//     selectAll: document.getElementById('selectAll'),
//     clearAll: document.getElementById('clearAll'),
//     generateButton: document.getElementById('generateButton'),
//     loadingOverlay: document.getElementById('loadingOverlay'),
//     loadingText: document.getElementById('loadingText'),
//     stepImport: document.getElementById('stepImport'),
//     stepReview: document.getElementById('stepReview'),
//     stepGenerate: document.getElementById('stepGenerate')
//   };

//   let state = {
//     file: null,
//     trips: [],
//     excluded: 0,
//     detectedDate: null,
//     rate: 0.625
//   };

//   function setLoading(show, text = 'Processing trip log…') {
//     els.loadingText.textContent = text;
//     els.loadingOverlay.classList.toggle('hidden', !show);
//   }

//   function showMessage(text, type = 'error') {
//     els.message.textContent = text;
//     els.message.className = `message ${type}`;
//   }

//   function clearMessage() {
//     els.message.className = 'message hidden';
//     els.message.textContent = '';
//   }

//   function formatBytes(bytes) {
//     if (bytes < 1024) return `${bytes} bytes`;
//     if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
//     return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
//   }

//   function monthLabel(date) {
//     return date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
//   }

//   function shortDate(date) {
//     return date.toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
//   }

//   function currency(value) {
//     return `$${value.toFixed(3)}`;
//   }

//   function selectedTrips() {
//     return state.trips.filter(t => t.included);
//   }

//   function escapeHtml(value) {
//     return String(value ?? '')
//       .replace(/&/g, '&amp;')
//       .replace(/</g, '&lt;')
//       .replace(/>/g, '&gt;')
//       .replace(/"/g, '&quot;')
//       .replace(/'/g, '&#039;');
//   }

//   function parseXml(text) {
//     const doc = new DOMParser().parseFromString(text, 'application/xml');
//     if (doc.getElementsByTagName('parsererror').length) {
//       throw new Error('An Excel XML file could not be read.');
//     }
//     return doc;
//   }

//   function serializeXml(doc) {
//     return new XMLSerializer().serializeToString(doc);
//   }

//   function directChild(parent, tagName) {
//     for (const child of parent.children) {
//       if (child.localName === tagName) return child;
//     }
//     return null;
//   }

//   function createElementLike(doc, parent, localName) {
//     const namespace = parent.namespaceURI || doc.documentElement.namespaceURI;
//     return doc.createElementNS(namespace, localName);
//   }

//   function excelCellColumn(cellRef) {
//     return String(cellRef || '').replace(/\d+/g, '');
//   }

//   function columnToNumber(col) {
//     let n = 0;
//     for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
//     return n;
//   }

//   function excelSerialToDate(value) {
//     const serial = Number(value);
//     if (!Number.isFinite(serial)) return null;
//     const epoch = Date.UTC(1899, 11, 30);
//     return new Date(epoch + serial * 86400000);
//   }

//   function dateToExcelSerial(date) {
//     const utc = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
//     return (utc - Date.UTC(1899, 11, 30)) / 86400000;
//   }

//   function parseStarted(value) {
//     if (value instanceof Date) return value;
//     if (typeof value === 'number') return excelSerialToDate(value);
//     const text = String(value ?? '').trim();
//     if (!text) return null;
//     const m = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
//     if (m) {
//       return new Date(
//         Number(m[1]), Number(m[2]) - 1, Number(m[3]),
//         Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0)
//       );
//     }
//     const parsed = new Date(text);
//     return Number.isNaN(parsed.getTime()) ? null : parsed;
//   }

//   function cellValue(cell, sharedStrings) {
//     if (!cell) return '';
//     const type = cell.getAttribute('t') || '';
//     if (type === 'inlineStr') {
//       const isNode = directChild(cell, 'is');
//       if (!isNode) return '';
//       return Array.from(isNode.getElementsByTagNameNS('*', 't')).map(n => n.textContent || '').join('');
//     }
//     const v = directChild(cell, 'v');
//     const raw = v ? v.textContent : '';
//     if (type === 's') return sharedStrings[Number(raw)] ?? '';
//     if (type === 'str') return raw;
//     if (type === 'b') return raw === '1';
//     if (raw === '') return '';
//     const numeric = Number(raw);
//     return Number.isFinite(numeric) ? numeric : raw;
//   }

//   async function readSharedStrings(zip) {
//     const file = zip.file('xl/sharedStrings.xml');
//     if (!file) return [];
//     const doc = parseXml(await file.async('string'));
//     return Array.from(doc.getElementsByTagNameNS('*', 'si')).map(si =>
//       Array.from(si.getElementsByTagNameNS('*', 't')).map(t => t.textContent || '').join('')
//     );
//   }

//   function normalizeTarget(target) {
//     const clean = String(target || '').replace(/^\//, '');
//     return clean.startsWith('xl/') ? clean : `xl/${clean}`;
//   }

//   async function findWorksheetPath(zip, desiredName) {
//     const workbookFile = zip.file('xl/workbook.xml');
//     const relFile = zip.file('xl/_rels/workbook.xml.rels');
//     if (!workbookFile || !relFile) throw new Error('This file is not a standard Excel .xlsx workbook.');

//     const wbDoc = parseXml(await workbookFile.async('string'));
//     const relDoc = parseXml(await relFile.async('string'));
//     const relationshipMap = new Map();
//     for (const rel of Array.from(relDoc.getElementsByTagNameNS('*', 'Relationship'))) {
//       relationshipMap.set(rel.getAttribute('Id'), rel.getAttribute('Target'));
//     }

//     const sheets = Array.from(wbDoc.getElementsByTagNameNS('*', 'sheet'));
//     const chosen = sheets.find(s => s.getAttribute('name') === desiredName) || sheets[0];
//     if (!chosen) throw new Error('No worksheets were found in the Excel workbook.');

//     const relationshipId = chosen.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || chosen.getAttribute('r:id');
//     const target = relationshipMap.get(relationshipId);
//     if (!target) throw new Error('The workbook worksheet relationship could not be resolved.');
//     return normalizeTarget(target);
//   }

//   function rowCellsByHeader(sheetDoc, sharedStrings) {
//     const rows = Array.from(sheetDoc.getElementsByTagNameNS('*', 'row'));
//     const firstRow = rows.find(r => r.getAttribute('r') === '1') || rows[0];
//     if (!firstRow) throw new Error('The Volvo workbook has no header row.');

//     const map = new Map();
//     for (const cell of Array.from(firstRow.getElementsByTagNameNS('*', 'c'))) {
//       const header = String(cellValue(cell, sharedStrings) ?? '').trim();
//       if (header) map.set(header, excelCellColumn(cell.getAttribute('r')));
//     }
//     return map;
//   }

//   function findCellInRow(row, column) {
//     return Array.from(row.getElementsByTagNameNS('*', 'c')).find(c => excelCellColumn(c.getAttribute('r')) === column) || null;
//   }

//   async function parseVolvoWorkbook(file) {
//     const buffer = await file.arrayBuffer();
//     const zip = await JSZip.loadAsync(buffer);
//     const sharedStrings = await readSharedStrings(zip);
//     const sheetPath = await findWorksheetPath(zip, 'Sheet1');
//     const sheetFile = zip.file(sheetPath);
//     if (!sheetFile) throw new Error('The Volvo worksheet could not be found.');
//     const sheetDoc = parseXml(await sheetFile.async('string'));
//     const headers = rowCellsByHeader(sheetDoc, sharedStrings);

//     const missing = REQUIRED_HEADERS.filter(h => !headers.has(h));
//     if (missing.length) {
//       throw new Error(`Missing required Volvo column${missing.length > 1 ? 's' : ''}: ${missing.join(', ')}`);
//     }

//     const trips = [];
//     let excluded = 0;
//     const rows = Array.from(sheetDoc.getElementsByTagNameNS('*', 'row')).filter(r => Number(r.getAttribute('r')) >= 2);

//     for (const row of rows) {
//       const category = String(cellValue(findCellInRow(row, headers.get('Category')), sharedStrings) ?? '').trim();
//       if (!category) continue;
//       if (category.toLowerCase() !== 'business') {
//         excluded++;
//         continue;
//       }

//       const startOdo = Number(cellValue(findCellInRow(row, headers.get('Start odometer (miles)')), sharedStrings));
//       const endOdo = Number(cellValue(findCellInRow(row, headers.get('End odometer (miles)')), sharedStrings));
//       const miles = endOdo - startOdo;
//       if (!Number.isFinite(startOdo) || !Number.isFinite(endOdo) || !(miles > 0.000001)) {
//         excluded++;
//         continue;
//       }

//       const startedValue = cellValue(findCellInRow(row, headers.get('Started')), sharedStrings);
//       const started = parseStarted(startedValue);
//       if (!started || Number.isNaN(started.getTime())) {
//         excluded++;
//         continue;
//       }

//       trips.push({
//         id: `trip-${row.getAttribute('r')}`,
//         sourceRow: Number(row.getAttribute('r')),
//         date: new Date(started.getFullYear(), started.getMonth(), started.getDate()),
//         started,
//         startOdo,
//         endOdo,
//         startAddress: String(cellValue(findCellInRow(row, headers.get('Start address')), sharedStrings) ?? '').trim(),
//         endAddress: String(cellValue(findCellInRow(row, headers.get('End address')), sharedStrings) ?? '').trim(),
//         miles,
//         included: true
//       });
//     }

//     trips.sort((a, b) => b.started.getTime() - a.started.getTime());
//     return { trips, excluded };
//   }

//   function recalculate() {
//     state.rate = Math.max(0, Number(els.rateInput.value) || 0);
//     const selected = selectedTrips();
//     const miles = selected.reduce((sum, t) => sum + t.miles, 0);
//     const pay = miles * state.rate;

//     els.tripCount.textContent = String(selected.length);
//     els.excludedCount.textContent = `${state.excluded + (state.trips.length - selected.length)} excluded`;
//     els.totalMiles.textContent = miles.toFixed(1);
//     els.totalPay.textContent = currency(pay);

//     state.trips.forEach(t => {
//       const payCell = document.querySelector(`[data-pay-id="${t.id}"]`);
//       if (payCell) payCell.textContent = currency(t.miles * state.rate);
//     });

//     const checked = selected.length;
//     els.masterCheckbox.checked = state.trips.length > 0 && checked === state.trips.length;
//     els.masterCheckbox.indeterminate = checked > 0 && checked < state.trips.length;
//     els.generateButton.disabled = checked === 0;

//     if (checked > 0 && state.file) {
//       clearMessage();
//     }
//   }

//   function renderTrips() {
//     els.tripTableBody.innerHTML = '';
//     const fragment = document.createDocumentFragment();
//     state.trips.forEach(t => {
//       const tr = document.createElement('tr');
//       tr.innerHTML = `
//         <td class="check-col"><input type="checkbox" data-trip-id="${t.id}" ${t.included ? 'checked' : ''} aria-label="Include trip ${shortDate(t.date)}" /></td>
//         <td class="trip-date">${shortDate(t.date)}</td>
//         <td class="address" title="${escapeHtml(t.startAddress)}">${escapeHtml(t.startAddress)}</td>
//         <td class="address" title="${escapeHtml(t.endAddress)}">${escapeHtml(t.endAddress)}</td>
//         <td class="number-col">${Math.round(t.startOdo).toLocaleString()}</td>
//         <td class="number-col">${Math.round(t.endOdo).toLocaleString()}</td>
//         <td class="number-col"><strong>${t.miles.toFixed(1)}</strong></td>
//         <td class="number-col" data-pay-id="${t.id}">${currency(t.miles * state.rate)}</td>`;
//       fragment.appendChild(tr);
//     });
//     els.tripTableBody.appendChild(fragment);

//     els.tripTableBody.querySelectorAll('input[type="checkbox"]').forEach(box => {
//       box.addEventListener('change', event => {
//         const trip = state.trips.find(t => t.id === event.target.dataset.tripId);
//         if (trip) trip.included = event.target.checked;
//         recalculate();
//       });
//     });
//   }

//   function updateSteps(stage) {
//     [els.stepImport, els.stepReview, els.stepGenerate].forEach(el => el.classList.remove('active', 'complete'));
//     if (stage === 'import') {
//       els.stepImport.classList.add('active');
//     } else if (stage === 'review') {
//       els.stepImport.classList.add('complete');
//       els.stepReview.classList.add('active');
//     } else {
//       els.stepImport.classList.add('complete');
//       els.stepReview.classList.add('complete');
//       els.stepGenerate.classList.add('active');
//     }
//   }

//   function resetResults() {
//     state.file = null;
//     state.trips = [];
//     state.excluded = 0;
//     state.detectedDate = null;
//     els.resultsSection.classList.add('hidden');
//     els.fileStatus.classList.add('hidden');
//     els.dropZone.classList.remove('hidden');
//     els.fileInput.value = '';
//     updateSteps('import');
//   }

//   async function handleFile(file) {
//     clearMessage();
//     if (!file) return;
//     if (!/\.xls(x|m)$/i.test(file.name)) {
//       showMessage('Please select an Excel .xlsx or .xlsm Volvo trip log.');
//       return;
//     }

//     setLoading(true, 'Reading Volvo trip log…');
//     try {
//       const { trips, excluded } = await parseVolvoWorkbook(file);
//       if (!trips.length) throw new Error('No reimbursable Business trips with positive mileage were found.');

//       state.file = file;
//       state.trips = trips;
//       state.excluded = excluded;
//       state.detectedDate = trips[0].date;

//       els.fileName.textContent = file.name;
//       els.fileDetails.textContent = `${formatBytes(file.size)} • ${trips.length} eligible Business trips found`;
//       els.fileStatus.classList.remove('hidden');
//       els.dropZone.classList.add('hidden');
//       els.resultsSection.classList.remove('hidden');
//       els.detectedMonth.textContent = monthLabel(state.detectedDate);
//       renderTrips();
//       recalculate();
//       updateSteps('review');
//       showMessage(`Loaded ${trips.length} reimbursable trips for ${monthLabel(state.detectedDate)}. The Excel form will resize automatically.`, 'success');
//     } catch (err) {
//       console.error(err);
//       resetResults();
//       showMessage(err.message || 'Unable to read this Excel file.');
//     } finally {
//       setLoading(false);
//     }
//   }

//   function clearCellContents(cell) {
//     if (!cell) return;
//     for (const child of Array.from(cell.children)) cell.removeChild(child);
//     cell.removeAttribute('t');
//   }

//   function getOrCreateCell(doc, sheetData, ref) {
//     let cell = Array.from(sheetData.getElementsByTagNameNS('*', 'c')).find(c => c.getAttribute('r') === ref);
//     if (cell) return cell;

//     const rowNumber = Number(ref.match(/\d+/)[0]);
//     const col = excelCellColumn(ref);
//     let row = Array.from(sheetData.getElementsByTagNameNS('*', 'row')).find(r => Number(r.getAttribute('r')) === rowNumber);
//     if (!row) {
//       row = createElementLike(doc, sheetData, 'row');
//       row.setAttribute('r', String(rowNumber));
//       const rows = Array.from(sheetData.getElementsByTagNameNS('*', 'row'));
//       const next = rows.find(r => Number(r.getAttribute('r')) > rowNumber);
//       if (next) sheetData.insertBefore(row, next); else sheetData.appendChild(row);
//     }

//     cell = createElementLike(doc, row, 'c');
//     cell.setAttribute('r', ref);
//     const colNumber = columnToNumber(col);
//     const existing = Array.from(row.getElementsByTagNameNS('*', 'c'));
//     const nextCell = existing.find(c => columnToNumber(excelCellColumn(c.getAttribute('r'))) > colNumber);
//     if (nextCell) row.insertBefore(cell, nextCell); else row.appendChild(cell);
//     return cell;
//   }

//   function setNumber(doc, sheetData, ref, value) {
//     const cell = getOrCreateCell(doc, sheetData, ref);
//     clearCellContents(cell);
//     const v = createElementLike(doc, cell, 'v');
//     v.textContent = String(value);
//     cell.appendChild(v);
//   }

//   function setInlineString(doc, sheetData, ref, value) {
//     const cell = getOrCreateCell(doc, sheetData, ref);
//     clearCellContents(cell);
//     cell.setAttribute('t', 'inlineStr');
//     const isNode = createElementLike(doc, cell, 'is');
//     const t = createElementLike(doc, isNode, 't');
//     const text = String(value ?? '');
//     if (/^\s|\s$/.test(text)) t.setAttribute('xml:space', 'preserve');
//     t.textContent = text;
//     isNode.appendChild(t);
//     cell.appendChild(isNode);
//   }

//   function setFormula(doc, sheetData, ref, formula, cachedValue) {
//     const cell = getOrCreateCell(doc, sheetData, ref);
//     clearCellContents(cell);
//     const f = createElementLike(doc, cell, 'f');
//     f.textContent = formula;
//     cell.appendChild(f);
//     const v = createElementLike(doc, cell, 'v');
//     v.textContent = String(cachedValue);
//     cell.appendChild(v);
//   }

//   function clearRange(doc, sheetData, columns, startRow, endRow) {
//     for (let row = startRow; row <= endRow; row++) {
//       for (const col of columns) {
//         const cell = getOrCreateCell(doc, sheetData, `${col}${row}`);
//         clearCellContents(cell);
//       }
//     }
//   }

//   function shiftRowElement(row, newRowNumber) {
//     row.setAttribute('r', String(newRowNumber));
//     for (const cell of Array.from(row.getElementsByTagNameNS('*', 'c'))) {
//       const ref = cell.getAttribute('r');
//       if (!ref) continue;
//       cell.setAttribute('r', `${excelCellColumn(ref)}${newRowNumber}`);
//     }
//   }

//   function cloneBlankDataRow(templateRow, rowNumber) {
//     const row = templateRow.cloneNode(true);
//     shiftRowElement(row, rowNumber);
//     for (const cell of Array.from(row.getElementsByTagNameNS('*', 'c'))) {
//       clearCellContents(cell);
//     }
//     return row;
//   }

//   function resizeMileageSection(sheetDoc, sheetData, tripCount) {
//     const delta = tripCount - (TEMPLATE_LAST_DATA_ROW - FIRST_DATA_ROW + 1);
//     const originalRows = Array.from(sheetData.getElementsByTagNameNS('*', 'row'));
//     const templateDataRow = originalRows.find(row => Number(row.getAttribute('r')) === TEMPLATE_LAST_DATA_ROW)
//       || originalRows.find(row => Number(row.getAttribute('r')) === FIRST_DATA_ROW);
//     if (!templateDataRow) throw new Error('The mileage template is missing its trip-row formatting.');

//     const footerRows = originalRows
//       .filter(row => Number(row.getAttribute('r')) >= TEMPLATE_TOTALS_ROW)
//       .map(row => row.cloneNode(true));

//     // Remove the template's fixed trip section and footer. They are rebuilt below.
//     for (const row of originalRows) {
//       if (Number(row.getAttribute('r')) >= FIRST_DATA_ROW) sheetData.removeChild(row);
//     }

//     // Create exactly as many formatted trip rows as are needed.
//     for (let i = 0; i < tripCount; i++) {
//       sheetData.appendChild(cloneBlankDataRow(templateDataRow, FIRST_DATA_ROW + i));
//     }

//     // Move totals/signature/instructions down or up with the resized trip section.
//     for (const row of footerRows) {
//       shiftRowElement(row, Number(row.getAttribute('r')) + delta);
//       sheetData.appendChild(row);
//     }

//     const mergeCells = sheetDoc.getElementsByTagNameNS('*', 'mergeCells')[0];
//     if (mergeCells) {
//       for (const merge of Array.from(mergeCells.getElementsByTagNameNS('*', 'mergeCell'))) {
//         const ref = merge.getAttribute('ref') || '';
//         const dataMerge = /^B(\d+):C\1$/.exec(ref);
//         if (dataMerge) {
//           const rowNum = Number(dataMerge[1]);
//           if (rowNum >= FIRST_DATA_ROW && rowNum <= TEMPLATE_LAST_DATA_ROW) {
//             mergeCells.removeChild(merge);
//             continue;
//           }
//         }

//         const rangeMatch = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref);
//         if (rangeMatch) {
//           const startRow = Number(rangeMatch[2]);
//           const endRow = Number(rangeMatch[4]);
//           if (startRow >= TEMPLATE_TOTALS_ROW) {
//             merge.setAttribute('ref', `${rangeMatch[1]}${startRow + delta}:${rangeMatch[3]}${endRow + delta}`);
//           }
//         }
//       }

//       for (let row = FIRST_DATA_ROW; row < FIRST_DATA_ROW + tripCount; row++) {
//         const merge = createElementLike(sheetDoc, mergeCells, 'mergeCell');
//         merge.setAttribute('ref', `B${row}:C${row}`);
//         mergeCells.appendChild(merge);
//       }
//       mergeCells.setAttribute('count', String(mergeCells.getElementsByTagNameNS('*', 'mergeCell').length));
//     }

//     const lastRow = TEMPLATE_LAST_ROW + delta;
//     const dimension = sheetDoc.getElementsByTagNameNS('*', 'dimension')[0];
//     if (dimension) dimension.setAttribute('ref', `B1:P${lastRow}`);

//     return {
//       dataLastRow: FIRST_DATA_ROW + tripCount - 1,
//       totalsRow: TEMPLATE_TOTALS_ROW + delta,
//       footerDateRow: 78 + delta,
//       printEndRow: TEMPLATE_PRINT_END_ROW + delta,
//       lastRow
//     };
//   }

//   function updatePrintArea(workbookDoc, printEndRow) {
//     const names = Array.from(workbookDoc.getElementsByTagNameNS('*', 'definedName'));
//     const printArea = names.find(node => node.getAttribute('name') === '_xlnm.Print_Area' && node.getAttribute('localSheetId') === '0');
//     if (printArea) printArea.textContent = `'${FORM_SHEET_NAME}'!$A$1:$M$${printEndRow}`;
//   }

//   function endOfMonth(date) {
//     return new Date(date.getFullYear(), date.getMonth() + 1, 0);
//   }

//   function safeFilenamePart(text) {
//     return text.replace(/[\\/:*?"<>|]/g, '-');
//   }

//   function updateCalculationMode(workbookDoc) {
//     let calcPr = workbookDoc.getElementsByTagNameNS('*', 'calcPr')[0];
//     if (!calcPr) {
//       calcPr = createElementLike(workbookDoc, workbookDoc.documentElement, 'calcPr');
//       workbookDoc.documentElement.appendChild(calcPr);
//     }
//     calcPr.setAttribute('calcMode', 'auto');
//     calcPr.setAttribute('fullCalcOnLoad', '1');
//     calcPr.setAttribute('forceFullCalc', '1');
//   }

//   async function removeCalculationChain(zip) {
//     zip.remove('xl/calcChain.xml');

//     const relPath = 'xl/_rels/workbook.xml.rels';
//     const relFile = zip.file(relPath);
//     if (relFile) {
//       const relDoc = parseXml(await relFile.async('string'));
//       for (const rel of Array.from(relDoc.getElementsByTagNameNS('*', 'Relationship'))) {
//         if ((rel.getAttribute('Type') || '').endsWith('/calcChain')) rel.parentNode.removeChild(rel);
//       }
//       zip.file(relPath, serializeXml(relDoc));
//     }

//     const contentFile = zip.file('[Content_Types].xml');
//     if (contentFile) {
//       const contentDoc = parseXml(await contentFile.async('string'));
//       for (const override of Array.from(contentDoc.getElementsByTagNameNS('*', 'Override'))) {
//         if (override.getAttribute('PartName') === '/xl/calcChain.xml') override.parentNode.removeChild(override);
//       }
//       zip.file('[Content_Types].xml', serializeXml(contentDoc));
//     }
//   }

//   async function loadTemplateRate() {
//     try {
//       const zip = await JSZip.loadAsync(window.MILEAGE_TEMPLATE_BASE64, { base64: true });
//       const strings = await readSharedStrings(zip);
//       const path = await findWorksheetPath(zip, FORM_SHEET_NAME);
//       const doc = parseXml(await zip.file(path).async('string'));
//       const cell = Array.from(doc.getElementsByTagNameNS('*', 'c')).find(c => c.getAttribute('r') === 'J4');
//       const value = Number(cellValue(cell, strings));
//       if (Number.isFinite(value) && value > 0) {
//         state.rate = value;
//         els.rateInput.value = value.toFixed(3);
//       }
//     } catch (err) {
//       console.warn('Could not read template rate; using 0.625.', err);
//     }
//   }

//   async function generateWorkbook() {
//     const trips = selectedTrips();
//     if (!trips.length) return;

//     setLoading(true, 'Building reimbursement workbook…');
//     updateSteps('generate');
//     try {
//       const zip = await JSZip.loadAsync(window.MILEAGE_TEMPLATE_BASE64, { base64: true });
//       const sheetPath = await findWorksheetPath(zip, FORM_SHEET_NAME);
//       const sheetFile = zip.file(sheetPath);
//       if (!sheetFile) throw new Error('The embedded mileage template could not be opened.');

//       const sheetDoc = parseXml(await sheetFile.async('string'));
//       const sheetData = sheetDoc.getElementsByTagNameNS('*', 'sheetData')[0];
//       if (!sheetData) throw new Error('The mileage template is missing its worksheet data area.');

//       const layout = resizeMileageSection(sheetDoc, sheetData, trips.length);

//       let totalMiles = 0;
//       trips.forEach((trip, index) => {
//         const row = FIRST_DATA_ROW + index;
//         const pay = trip.miles * state.rate;
//         totalMiles += trip.miles;
//         setNumber(sheetDoc, sheetData, `B${row}`, dateToExcelSerial(trip.date));
//         setNumber(sheetDoc, sheetData, `D${row}`, trip.startOdo);
//         setNumber(sheetDoc, sheetData, `E${row}`, trip.endOdo);
//         setInlineString(sheetDoc, sheetData, `F${row}`, trip.startAddress);
//         setInlineString(sheetDoc, sheetData, `G${row}`, trip.endAddress);
//         setFormula(sheetDoc, sheetData, `K${row}`, `E${row}-D${row}`, trip.miles);
//         setFormula(sheetDoc, sheetData, `L${row}`, `K${row}*$J$4`, pay);
//       });

//       const totalPay = totalMiles * state.rate;
//       const monthDate = trips[0].date;
//       const monthName = monthLabel(monthDate);
//       setNumber(sheetDoc, sheetData, 'D3', dateToExcelSerial(endOfMonth(monthDate)));
//       setInlineString(sheetDoc, sheetData, 'J3', monthName);
//       setNumber(sheetDoc, sheetData, 'J4', state.rate);
//       setInlineString(sheetDoc, sheetData, `J${layout.totalsRow}`, 'Totals');
//       setFormula(sheetDoc, sheetData, `K${layout.totalsRow}`, `SUM(K${FIRST_DATA_ROW}:K${layout.dataLastRow})`, totalMiles);
//       setFormula(sheetDoc, sheetData, `L${layout.totalsRow}`, `SUM(L${FIRST_DATA_ROW}:L${layout.dataLastRow})`, totalPay);
//       setFormula(sheetDoc, sheetData, 'J5', `K${layout.totalsRow}`, totalMiles);
//       setFormula(sheetDoc, sheetData, 'J6', `L${layout.totalsRow}`, totalPay);
//       setNumber(sheetDoc, sheetData, `F${layout.footerDateRow}`, dateToExcelSerial(endOfMonth(monthDate)));

//       zip.file(sheetPath, serializeXml(sheetDoc));

//       const workbookFile = zip.file('xl/workbook.xml');
//       if (workbookFile) {
//         const workbookDoc = parseXml(await workbookFile.async('string'));
//         updateCalculationMode(workbookDoc);
//         updatePrintArea(workbookDoc, layout.printEndRow);
//         zip.file('xl/workbook.xml', serializeXml(workbookDoc));
//       }
//       await removeCalculationChain(zip);

//       const blob = await zip.generateAsync({
//         type: 'blob',
//         mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
//         compression: 'DEFLATE',
//         compressionOptions: { level: 6 }
//       });
//       const url = URL.createObjectURL(blob);
//       const link = document.createElement('a');
//       link.href = url;
//       link.download = `G.G. ${safeFilenamePart(monthName)} Mileage - Generated.xlsx`;
//       document.body.appendChild(link);
//       link.click();
//       link.remove();
//       setTimeout(() => URL.revokeObjectURL(url), 1500);
//       showMessage(`Generated ${monthName}: ${trips.length} trips, ${totalMiles.toFixed(1)} miles, ${currency(totalPay)} reimbursement.`, 'success');
//     } catch (err) {
//       console.error(err);
//       showMessage(err.message || 'Could not generate the reimbursement form.');
//       updateSteps('review');
//     } finally {
//       setLoading(false);
//     }
//   }

//   els.chooseButton.addEventListener('click', event => {
//     event.preventDefault();
//     els.fileInput.click();
//   });
//   els.fileInput.addEventListener('change', event => handleFile(event.target.files[0]));
//   els.changeFile.addEventListener('click', () => {
//     resetResults();
//     clearMessage();
//     els.fileInput.click();
//   });
//   els.dropZone.addEventListener('dragover', event => {
//     event.preventDefault();
//     els.dropZone.classList.add('dragover');
//   });
//   els.dropZone.addEventListener('dragleave', () => els.dropZone.classList.remove('dragover'));
//   els.dropZone.addEventListener('drop', event => {
//     event.preventDefault();
//     els.dropZone.classList.remove('dragover');
//     handleFile(event.dataTransfer.files[0]);
//   });
//   els.rateInput.addEventListener('input', recalculate);
//   els.masterCheckbox.addEventListener('change', () => {
//     state.trips.forEach(t => t.included = els.masterCheckbox.checked);
//     renderTrips();
//     recalculate();
//   });
//   els.selectAll.addEventListener('click', () => {
//     state.trips.forEach(t => t.included = true);
//     renderTrips();
//     recalculate();
//   });
//   els.clearAll.addEventListener('click', () => {
//     state.trips.forEach(t => t.included = false);
//     renderTrips();
//     recalculate();
//   });
//   els.generateButton.addEventListener('click', generateWorkbook);

//   window.addEventListener('load', async () => {
//     updateSteps('import');
//     if (!window.JSZip) {
//       showMessage('The local Excel processing library could not be loaded.');
//       return;
//     }
//     await loadTemplateRate();
//   });
// })();
