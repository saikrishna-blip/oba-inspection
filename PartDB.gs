/**
 * PART REFERENCE DATABASE  (single file - paste over the whole Code.gs)
 * ----------------------------------------------------------------------
 *  - Image linker            : Part DB > Link part images
 *  - Assign one part         : Part DB > Assign part to models
 *  - Bulk BOM builder        : Part DB > Build model BOM (bulk)
 *  - Compare model BOMs      : Part DB > Compare model BOMs
 *  - BOM lock / unlock (ECN) : Part DB > BOM lock / unlock
 *  - Lot-wise verification   : Part DB > BOM verification (lot-wise)
 *
 * SYSTEM SHEETS (created automatically, do not rename):
 *   BOM_CONTROL, ECN_HISTORY, BOM_VERIFY_LOG, LOT_BOM_DETAIL,
 *   LOT_BOM_VIEW, BOM_COMPARE
 *
 * IMAGE FILE NAMING RULE:
 *   <PART_NO>.jpg   -> primary image  (e.g. LH1G-OPT-0042.jpg)
 *   <PART_NO>_2.jpg -> alternate view (ignored by the linker)
 * ----------------------------------------------------------------------
 */

// ======================= CONFIG =======================
const IMAGE_FOLDER_ID = '1xwpe_z03u3dk2z40QHJ_559yDouuAFHM';
const SHEET_NAME      = 'PART_MASTER';
const HEADER_ROW      = 1;
const COL_PART_NO     = 1;   // A
const COL_IMAGE       = 3;   // C  (=IMAGE formula goes here)
const COL_IMAGE_URL   = 4;   // D  (raw link, used by other tools)
const ROW_HEIGHT_PX   = 90;  // row height so thumbnails are readable
const THUMB_WIDTH     = 400; // px width requested from Drive
const MAKE_PUBLIC     = true; // set false if your domain blocks link sharing
// Portal (LUMIO Quality Management - Projector) spreadsheet: its "Sessions" sheet
// decides who may open these screens on the web. Admins only.
const PORTAL_SS_ID    = '1WC6WGmwiqqytYobfoM9P4dP1RIptj5Rj17LSYtP_8tI';
// ======================================================

/**
 * Where the BOM data lives. Since the data was merged, it is the portal spreadsheet
 * (it has BOM_CONTROL); before that, the spreadsheet this script is bound to.
 */
let DATA_SS_CACHE_ = null;
function dataSS_() {
  if (!DATA_SS_CACHE_) {
    let p = null;
    try { p = SpreadsheetApp.openById(PORTAL_SS_ID); } catch (e) {}
    DATA_SS_CACHE_ = (p && p.getSheetByName('BOM_CONTROL')) ? p : SpreadsheetApp.getActiveSpreadsheet();
  }
  return DATA_SS_CACHE_;
}

/** Adds the custom menu when the spreadsheet opens. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Part DB')
    .addItem('Assign part to models', 'showAssignSidebar')
    .addItem('Build model BOM (bulk)', 'showBomDialog')
    .addItem('Compare model BOMs (all models)', 'buildComparison')
    .addItem('Compare two models', 'showTwoCompareDialog')
    .addSeparator()
    .addItem('BOM lock / unlock', 'showControlDialog')
    .addItem('Generate ECN (part replace / split)', 'showEcnWizardDialog')
    .addItem('BOM verification (lot-wise)', 'showVerifyDialog')
    .addItem('BOM version history (view / restore)', 'showBomHistoryDialog')
    .addItem('Print ECN report', 'showEcnReportDialog')
    .addSeparator()
    .addItem('Link part images', 'linkPartImages')
    .addItem('Report missing images', 'reportMissingImages')
    .addItem('Remove duplicate parts (PART_MASTER)', 'dedupePartMasterManual')
    .addSeparator()
    .addItem('Clear image column', 'clearImageColumn')
    .addToUi();
}

/**
 * Runs automatically whenever a cell is edited or pasted into. Whenever the
 * edit touches PART_MASTER's Part_No column (which is what a bulk
 * copy-paste from Excel lands on), it quietly cleans up any duplicate part
 * numbers that paste just created, so PART_MASTER never ends up holding the
 * same part twice.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;
    if (e.range.getRow() + e.range.getNumRows() - 1 <= HEADER_ROW) return;
    const firstCol = e.range.getColumn();
    const lastCol  = e.range.getColumn() + e.range.getNumColumns() - 1;
    if (COL_PART_NO < firstCol || COL_PART_NO > lastCol) return;
    const r = dedupePartMaster_();
    if (r.removed) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        r.removed + ' duplicate part row(s) merged and removed from PART_MASTER.', 'Part DB', 6);
    }
  } catch (err) {
    Logger.log('onEdit dedupe failed: ' + err.message);
  }
}

/**
 * Scans PART_MASTER for duplicate Part_No rows. Keeps the first row of each
 * duplicate group, filling in any of its blank cells (Description, Image,
 * Image_URL, etc.) from the later duplicate rows, then deletes the later
 * rows. Returns { removed, parts } - how many rows were removed and which
 * part numbers they belonged to.
 */
function dedupePartMaster_() {
  const sheet = dataSS_().getSheetByName(SHEET_NAME);
  if (!sheet) return { removed: 0, parts: [] };
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= HEADER_ROW || lastCol === 0) return { removed: 0, parts: [] };

  const numRows = lastRow - HEADER_ROW;
  const rng   = sheet.getRange(HEADER_ROW + 1, 1, numRows, lastCol);
  const vals  = rng.getValues();
  const forms = rng.getFormulas();

  // A column counts as a "formula column" (e.g. the =IMAGE() column) if any
  // row in it currently holds a formula - those are written back with
  // setFormulas, everything else with setValues, so neither call clobbers
  // the other's cells.
  const formulaCol = [];
  for (let c = 0; c < lastCol; c++) formulaCol.push(forms.some(r => r[c]));

  const firstRowOf = {};
  const dupRowIdx  = [];
  const removedParts = [];

  for (let i = 0; i < vals.length; i++) {
    const key = norm_(vals[i][COL_PART_NO - 1]);
    if (!key) continue;
    if (!(key in firstRowOf)) { firstRowOf[key] = i; continue; }

    const keepIdx = firstRowOf[key];
    for (let c = 0; c < lastCol; c++) {
      const keptIsBlank = formulaCol[c]
        ? !forms[keepIdx][c]
        : (vals[keepIdx][c] === '' || vals[keepIdx][c] === null || vals[keepIdx][c] === undefined);
      if (!keptIsBlank) continue;
      if (formulaCol[c]) {
        if (forms[i][c]) forms[keepIdx][c] = forms[i][c];
      } else if (vals[i][c] !== '' && vals[i][c] !== null && vals[i][c] !== undefined) {
        vals[keepIdx][c] = vals[i][c];
      }
    }
    dupRowIdx.push(i);
    removedParts.push(String(vals[i][COL_PART_NO - 1]).trim());
  }

  if (!dupRowIdx.length) return { removed: 0, parts: [] };

  for (let c = 0; c < lastCol; c++) {
    const colRange = sheet.getRange(HEADER_ROW + 1, c + 1, numRows, 1);
    if (formulaCol[c]) colRange.setFormulas(forms.map(r => [r[c]]));
    else colRange.setValues(vals.map(r => [r[c]]));
  }

  dupRowIdx.slice().sort((a, b) => b - a).forEach(i => sheet.deleteRow(HEADER_ROW + 1 + i));

  return { removed: dupRowIdx.length, parts: removedParts };
}

/** Menu entry: run the PART_MASTER dedupe on demand and report the result. */
function dedupePartMasterManual() {
  const r = dedupePartMaster_();
  const ui = SpreadsheetApp.getUi();
  const msg = r.removed
    ? r.removed + ' duplicate part row(s) were found and removed from PART_MASTER:\n\n' + r.parts.join(', ')
    : 'No duplicate part numbers were found in PART_MASTER.';
  ui.alert('Clean up duplicates', msg, ui.ButtonSet.OK);
}

/** Builds { PART_NO : fileId } from the Drive folder. */
function buildImageIndex_() {
  const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
  const files  = folder.getFiles();
  const index  = {};

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();

    // strip extension, uppercase, ignore "_2" / "_3" alternates
    const base = name.replace(/\.[^.]+$/, '').trim().toUpperCase();
    if (/_\d+$/.test(base)) continue;

    index[base] = file.getId();

    if (MAKE_PUBLIC) {
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        Logger.log('Sharing failed for ' + name + ': ' + e.message);
      }
    }
  }
  return index;
}

/** Drive thumbnail endpoint - this is the one that renders inside =IMAGE(). */
function thumbUrl_(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w' + THUMB_WIDTH;
}

/** Main routine: writes =IMAGE() formulas for every matched part number. */
function linkPartImages() {
  const ss    = dataSS_();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return { matched: 0, total: 0 };

  const numRows  = lastRow - HEADER_ROW;
  const partNos  = sheet.getRange(HEADER_ROW + 1, COL_PART_NO, numRows, 1).getValues();
  const index    = buildImageIndex_();

  const formulas = [];
  const urls     = [];
  let matched    = 0;

  for (let i = 0; i < numRows; i++) {
    const key = String(partNos[i][0]).trim().toUpperCase();

    if (key && index[key]) {
      const url = thumbUrl_(index[key]);
      formulas.push(['=IMAGE("' + url + '", 1)']);
      urls.push([url]);
      matched++;
    } else {
      formulas.push(['']);
      urls.push(['']);
    }
  }

  sheet.getRange(HEADER_ROW + 1, COL_IMAGE,     numRows, 1).setFormulas(formulas);
  sheet.getRange(HEADER_ROW + 1, COL_IMAGE_URL, numRows, 1).setValues(urls);
  sheet.setRowHeights(HEADER_ROW + 1, numRows, ROW_HEIGHT_PX);

  try {
    SpreadsheetApp.getActive().toast(
      matched + ' of ' + numRows + ' parts linked to an image.', 'Part DB', 6);
  } catch (e) { /* no toast outside the spreadsheet (web app) */ }
  return { matched: matched, total: numRows };
}

/** Lists part numbers with no matching image file, in a side panel. */
function reportMissingImages() {
  const sheet   = dataSS_().getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - HEADER_ROW;
  const partNos = sheet.getRange(HEADER_ROW + 1, COL_PART_NO, numRows, 1).getValues();
  const index   = buildImageIndex_();

  const missing = partNos
    .map(r => String(r[0]).trim())
    .filter(p => p && !index[p.toUpperCase()]);

  const html = '<div style="font:13px Arial;padding:8px">' +
    '<b>' + missing.length + ' part(s) without an image</b><br><br>' +
    (missing.length ? missing.join('<br>') : 'All parts have images.') +
    '</div>';

  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutput(html).setTitle('Missing Images'));
}

/** Wipes the image and URL columns (use before a full re-link). */
function clearImageColumn() {
  const sheet   = dataSS_().getSheetByName(SHEET_NAME);
  const numRows = sheet.getLastRow() - HEADER_ROW;
  if (numRows <= 0) return;
  sheet.getRange(HEADER_ROW + 1, COL_IMAGE,     numRows, 1).clearContent();
  sheet.getRange(HEADER_ROW + 1, COL_IMAGE_URL, numRows, 1).clearContent();
}

/** Optional: run once to refresh links automatically every night. */
function createNightlyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'linkPartImages')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('linkPartImages')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
}

// ======================================================================
//        SHARED HELPERS, SYSTEM SHEETS, BOM LOCK, ECN LOG
// ======================================================================
const CONTROL_SHEET = 'BOM_CONTROL';
const ECN_SHEET     = 'ECN_HISTORY';
const VERIFY_SHEET  = 'BOM_VERIFY_LOG';
const LOT_SHEET     = 'LOT_BOM_DETAIL';
const LOTVIEW_SHEET = 'LOT_BOM_VIEW';
const PROTECT_DESC  = 'BOM LOCKED - change only via Part DB menu with an ECN';
const QPS_HEADER    = 'QPS';
const DESC_COL      = 2;   // B = fallback description column, used only if no header matches /desc/i

const CONTROL_HDR = ['Model', 'Status', 'Rev', 'Locked_By', 'Locked_On', 'Last_ECN'];
const ECN_HDR     = ['Date_Time', 'ECN_No', 'Model', 'BOM_Rev', 'Part_No', 'Action',
                     'Old_QPS', 'New_QPS', 'Description', 'Effective_Date', 'Document_Link', 'Changed_By'];
const ECN_DOC_FOLDER = 'ECN_Documents'; // Drive folder name, created alongside the spreadsheet
const VERIFY_HDR  = ['Ver_ID', 'Date_Time', 'Model', 'Lot_No', 'BOM_Rev', 'BOM_Status',
                     'Parts_Checked', 'Deviations', 'Result', 'Verified_By', 'Remarks'];
// LOT_HDR is defined in the verification module (BOM-interface column order).

// ---------------------------- basic helpers ----------------------------
function norm_(v) { return String(v === null || v === undefined ? '' : v).trim().toUpperCase(); }

function txt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  return v === null || v === undefined ? '' : String(v);
}

function now_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function user_() {
  let u = '';
  try { u = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(); } catch (e) {}
  return u || 'unknown';
}

function withLock_(fn) {
  const l = LockService.getDocumentLock() || LockService.getScriptLock();
  l.waitLock(30000);
  try { return fn(); } finally { l.releaseLock(); }
}

function isSystemSheet_(name) {
  return [SHEET_NAME, COMPARE_SHEET, CONTROL_SHEET, ECN_SHEET, VERIFY_SHEET, LOT_SHEET, LOTVIEW_SHEET]
    .indexOf(name) >= 0;
}

/** Model sheets = every sheet that is not a system sheet and has a Part_No heading in row 1
 *  (the spreadsheet also holds the portal's own tabs: Batches, Master, Users, ...). */
function getModelSheets_() {
  return dataSS_().getSheets().filter(s => {
    const n = s.getName();
    if (isSystemSheet_(n) || n === TWO_SHEET) return false;
    if (s.getLastColumn() < 1) return false;
    const h = s.getRange(HEADER_ROW, 1, 1, Math.min(s.getLastColumn(), 40)).getValues()[0];
    return h.some(x => String(x).toLowerCase().replace(/[^a-z0-9]/g, '') === 'partno');
  });
}

function headersOf_(sheet) {
  const lc = sheet.getLastColumn();
  if (lc === 0) return [];
  return sheet.getRange(HEADER_ROW, 1, 1, lc).getValues()[0].map(h => String(h).trim());
}

/** Column number of the Part_No column inside a model sheet. */
function partColIn_(sheet, masterHeaders) {
  const target = norm_(masterHeaders[COL_PART_NO - 1] || '');
  const idx = headersOf_(sheet).findIndex(h => norm_(h) === target);
  return idx >= 0 ? idx + 1 : 1;
}

/** Creates a system sheet (plain-text columns, header row) if missing. */
function sysSheet_(name, headers) {
  const ss = dataSS_();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const prev = ss.getActiveSheet();
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, headers.length, 130);
  try { ss.setActiveSheet(prev); } catch (e) {}
  return sh;
}

function appendRows_(sh, rows) {
  if (!rows.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Gets a system sheet and makes sure its header row matches wantHeaders,
 * inserting any missing columns so old rows keep their data instead of
 * reading back shifted or blank once the schema gains a column. Each
 * missing header is anchored to insert right after whatever precedes it
 * in wantHeaders (so several new columns added at once still land in the
 * right order relative to each other). Existing sheets are migrated in
 * place; new sheets are created with the header already correct.
 */
function migratedSheet_(name, wantHeaders) {
  const sh = sysSheet_(name, wantHeaders);
  const have = headersOf_(sh).map(norm_);
  const want = wantHeaders.map(norm_);
  if (have.length && have.join('|') !== want.join('|')) {
    want.forEach((h, i) => {
      if (have.indexOf(h) < 0) {
        const anchor = i > 0 ? want[i - 1] : '';
        const afterIdx = anchor ? have.indexOf(anchor) : -1;
        const at = afterIdx >= 0 ? afterIdx + 1 : have.length;
        sh.insertColumnAfter(at || 1);
        have.splice(at, 0, h);
      }
    });
    sh.getRange(1, 1, 1, wantHeaders.length).setValues([wantHeaders])
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  }
  return sh;
}

// ---------------------------- BOM control (lock) ----------------------------
/** { MODEL_UPPER : {row, model, status, rev, by, on, ecn} } - read only. */
function controlMap_() {
  const sh = dataSS_().getSheetByName(CONTROL_SHEET);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, CONTROL_HDR.length).getValues().forEach((r, i) => {
    if (r[0]) map[norm_(r[0])] = {
      row: i + 2, model: String(r[0]), status: String(r[1]) || 'Unlocked',
      rev: Number(r[2]) || 0, by: txt_(r[3]), on: txt_(r[4]), ecn: txt_(r[5])
    };
  });
  return map;
}

function getControl_(model) {
  return controlMap_()[norm_(model)] || { row: 0, model: model, status: 'Unlocked', rev: 0, by: '', on: '', ecn: '' };
}

function setControl_(model, c) {
  const sh = sysSheet_(CONTROL_SHEET, CONTROL_HDR);
  const old = getControl_(model);
  const vals = [model, c.status, c.rev, c.by, c.on, c.ecn];
  if (old.row) sh.getRange(old.row, 1, 1, CONTROL_HDR.length).setValues([vals]);
  else appendRows_(sh, [vals]);
}

function countParts_(sheet, mh) {
  const n = sheet.getLastRow() - HEADER_ROW;
  if (n <= 0) return 0;
  return sheet.getRange(HEADER_ROW + 1, partColIn_(sheet, mh), n, 1).getValues()
    .filter(r => norm_(r[0])).length;
}

function showControlDialog() {
  const html = HtmlService.createHtmlOutput(getControlHtml_()).setWidth(1500).setHeight(900);
  SpreadsheetApp.getUi().showModelessDialog(html, 'BOM lock / unlock');
}

function getControlData() {
  const ctrl = controlMap_();
  const mh = headersOf_(dataSS_().getSheetByName(SHEET_NAME));
  return getModelSheets_().map(s => {
    const c = ctrl[norm_(s.getName())];
    return {
      model: s.getName(), parts: countParts_(s, mh),
      status: c ? c.status : 'Unlocked', rev: c ? c.rev : 0,
      by: c ? c.by : '', on: c ? c.on : '', ecn: c ? c.ecn : ''
    };
  });
}

function lockBom(model) {
  return withLock_(() => {
    const ss = dataSS_();
    const sh = ss.getSheetByName(model);
    if (!sh || isSystemSheet_(model)) throw new Error('Model sheet not found: ' + model);
    const mh = headersOf_(ss.getSheetByName(SHEET_NAME));
    if (countParts_(sh, mh) === 0) throw new Error('BOM is empty - add parts before locking.');
    const c = getControl_(model);
    if (c.status === 'Locked') throw new Error(model + ' is already locked.');
    const rev = c.row ? c.rev + 1 : 0;   // first lock = Rev 0, re-lock after unlock = next rev
    setControl_(model, { status: 'Locked', rev: rev, by: user_(), on: now_(), ecn: c.ecn });
    appendRows_(migratedSheet_(ECN_SHEET, ECN_HDR),
      [[now_(), '-', model, rev, '', 'BOM LOCKED', '', '', 'BOM baseline locked', '', '', user_()]]);
    try {
      const p = sh.protect();
      p.setDescription(PROTECT_DESC);
      p.setWarningOnly(true);
    } catch (e) { Logger.log('Protect failed: ' + e.message); }
    return { ok: true };
  });
}

function unlockBom(model, reason) {
  return withLock_(() => {
    const ss = dataSS_();
    const sh = ss.getSheetByName(model);
    if (!sh) throw new Error('Model sheet not found: ' + model);
    if (!String(reason || '').trim()) throw new Error('Enter a reason for unlocking.');
    const c = getControl_(model);
    if (c.status !== 'Locked') throw new Error(model + ' is not locked.');
    setControl_(model, { status: 'Unlocked', rev: c.rev, by: c.by, on: c.on, ecn: c.ecn });
    appendRows_(migratedSheet_(ECN_SHEET, ECN_HDR),
      [[now_(), '-', model, c.rev, '', 'BOM UNLOCKED', '', '', String(reason).trim(), '', '', user_()]]);
    try {
      sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
        if (p.getDescription() === PROTECT_DESC) p.remove();
      });
    } catch (e) { Logger.log('Unprotect failed: ' + e.message); }
    return { ok: true };
  });
}

// ---------------------------- master + BOM change engine ----------------------------
/** 1-based column number of the first header matching re, or 0 if none. */
function findCol_(headers, re) {
  const i = headers.findIndex(h => re.test(String(h)));
  return i >= 0 ? i + 1 : 0;
}

function readMaster_() {
  const master = dataSS_().getSheetByName(SHEET_NAME);
  const mh = headersOf_(master);
  const last = master.getLastRow();
  let mVals = [], mForm = [];
  if (last > HEADER_ROW) {
    const rng = master.getRange(HEADER_ROW + 1, 1, last - HEADER_ROW, mh.length);
    mVals = rng.getValues();
    mForm = rng.getFormulas();
  }
  const idx = {};
  mVals.forEach((r, i) => {
    const k = norm_(r[COL_PART_NO - 1]);
    if (k && !(k in idx)) idx[k] = i;
  });
  // Description / image-URL columns are found by header name (e.g. "Description",
  // "Part Description", "Image_URL"), so this works whatever order PART_MASTER uses.
  // Falls back to the fixed DESC_COL / COL_IMAGE_URL only if no header matches.
  const descCol = findCol_(mh, /desc/i) || DESC_COL;
  const urlCol  = findCol_(mh, /image.*url|url.*image/i) || COL_IMAGE_URL;
  return { mh: mh, mVals: mVals, mForm: mForm, idx: idx, descCol: descCol, urlCol: urlCol };
}

/** Description + image URL of PART_MASTER row i (0-based), or blanks if i is undefined. */
function masterExtra_(ctx, i) {
  if (i === undefined || i === null) return { desc: '', img: '' };
  return {
    desc: ctx.descCol ? String(ctx.mVals[i][ctx.descCol - 1] || '').trim() : '',
    img:  ctx.urlCol  ? String(ctx.mVals[i][ctx.urlCol  - 1] || '').trim() : ''
  };
}

/** Makes sure the model sheet has headers and a QPS column. */
function ensureHeaders_(sheet, masterHeaders) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, masterHeaders.length + 1)
      .setValues([masterHeaders.concat([QPS_HEADER])]);
  }
  const headers = headersOf_(sheet);
  let qpsCol = headers.findIndex(h => norm_(h) === QPS_HEADER) + 1;
  if (!qpsCol) {
    qpsCol = headers.length + 1;
    sheet.getRange(HEADER_ROW, qpsCol).setValue(QPS_HEADER);
    headers.push(QPS_HEADER);
  }
  return { headers: headers, qpsCol: qpsCol };
}

/** Finds { row, qps } of a part in a model sheet, or null. */
function findInModel_(sheet, partNo, masterHeaders) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return null;
  const pc   = partColIn_(sheet, masterHeaders);
  const vals = sheet.getRange(HEADER_ROW + 1, pc, lastRow - HEADER_ROW, 1).getValues();
  const key  = norm_(partNo);
  for (let i = 0; i < vals.length; i++) {
    if (norm_(vals[i][0]) === key) {
      const row = HEADER_ROW + 1 + i;
      const qc  = headersOf_(sheet).findIndex(h => norm_(h) === QPS_HEADER) + 1;
      return { row: row, qps: qc ? sheet.getRange(row, qc).getValue() : '' };
    }
  }
  return null;
}

/** Works out what would change in one model sheet. Writes nothing (except a missing header). */
function planBom_(sheet, ctx, items, remove) {
  const h  = ensureHeaders_(sheet, ctx.mh);
  const pc = partColIn_(sheet, ctx.mh);
  const n  = Math.max(sheet.getLastRow() - HEADER_ROW, 0);
  const existing = {};
  let qv = [];
  if (n > 0) {
    const pv = sheet.getRange(HEADER_ROW + 1, pc, n, 1).getValues();
    qv = sheet.getRange(HEADER_ROW + 1, h.qpsCol, n, 1).getValues();
    pv.forEach((r, i) => {
      const k = norm_(r[0]);
      if (k && !(k in existing)) existing[k] = i;
    });
  }
  const changes = [], newRows = [], missing = [], removeIdx = [];

  (items || []).forEach(it => {
    const k = norm_(it.partNo), q = Number(it.qps);
    if (!(q > 0)) throw new Error('Invalid QPS for ' + it.partNo);
    if (k in existing) {
      const old = qv[existing[k]][0];
      if (Number(old) !== q) {
        changes.push({ partNo: it.partNo, action: 'QPS changed', oldQ: old, newQ: q });
        qv[existing[k]][0] = q;
      }
    } else if (k in ctx.idx) {
      const i = ctx.idx[k];
      newRows.push(h.headers.map((name, j) => {
        if (j + 1 === h.qpsCol) return q;
        const m = ctx.mh.findIndex(x => norm_(x) === norm_(name));
        if (m < 0) return '';
        return ctx.mForm[i][m] || ctx.mVals[i][m];
      }));
      changes.push({ partNo: it.partNo, action: 'Added', oldQ: '', newQ: q });
    } else {
      missing.push(it.partNo);
    }
  });

  (remove || []).forEach(p => {
    const k = norm_(p);
    if (k in existing) {
      removeIdx.push(existing[k]);
      changes.push({ partNo: p, action: 'Removed', oldQ: qv[existing[k]][0], newQ: '' });
    }
  });
  return { h: h, n: n, qv: qv, newRows: newRows, removeIdx: removeIdx, changes: changes, missing: missing };
}

function applyBom_(sheet, plan) {
  if (plan.n > 0 && plan.changes.some(c => c.action === 'QPS changed'))
    sheet.getRange(HEADER_ROW + 1, plan.h.qpsCol, plan.n, 1).setValues(plan.qv);
  if (plan.newRows.length) {
    const start = Math.max(sheet.getLastRow(), HEADER_ROW) + 1;
    sheet.getRange(start, 1, plan.newRows.length, plan.h.headers.length).setValues(plan.newRows);
    sheet.setRowHeights(start, plan.newRows.length, ROW_HEIGHT_PX);
  }
  plan.removeIdx.slice().sort((a, b) => b - a).forEach(i => sheet.deleteRow(HEADER_ROW + 1 + i));
}

/**
 * Applies plans for one or more models. Locked models with changes
 * need an ECN (checked BEFORE anything is written) and get logged.
 */
function commitPlans_(list, ecn) {
  const ctrl = controlMap_();
  const locked = list.filter(x => x.plan.changes.length &&
    (ctrl[norm_(x.sheet.getName())] || {}).status === 'Locked');

  let e = null;
  if (locked.length) {
    const no   = String((ecn && ecn.no) || '').trim();
    const desc = String((ecn && ecn.desc) || '').trim();
    if (!no || !desc)
      throw new Error('BOM of ' + locked.map(x => x.sheet.getName()).join(', ') +
        ' is LOCKED. Enter the ECN number and a change description.');
    e = { no: no, desc: desc };
  }

  list.forEach(x => applyBom_(x.sheet, x.plan));

  if (e) {
    const rows = [];
    locked.forEach(x => {
      const name = x.sheet.getName();
      const c = ctrl[norm_(name)];
      const rev = (c.rev || 0) + 1;
      setControl_(name, { status: 'Locked', rev: rev, by: c.by, on: c.on, ecn: e.no });
      x.plan.changes.forEach(ch =>
        rows.push([now_(), e.no, name, rev, ch.partNo, ch.action, ch.oldQ, ch.newQ, e.desc, '', '', user_()]));
    });
    appendRows_(migratedSheet_(ECN_SHEET, ECN_HDR), rows);
  }
}

// ======================================================================
//        GENERATE ECN (part replacement / one-to-many split)
// ======================================================================
function showEcnWizardDialog() {
  const html = HtmlService.createHtmlOutput(getEcnHtml_()).setWidth(1600).setHeight(980);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Generate ECN');
}

function showBomHistoryDialog() {
  const html = HtmlService.createHtmlOutput(getBomHistoryHtml_()).setWidth(1600).setHeight(980);
  SpreadsheetApp.getUi().showModelessDialog(html, 'BOM version history');
}

function showEcnReportDialog() {
  const html = HtmlService.createHtmlOutput(getEcnReportHtml_()).setWidth(1100).setHeight(980);
  SpreadsheetApp.getUi().showModelessDialog(html, 'ECN report');
}

/** Next auto ECN number: one past the highest existing ECN-#### in ECN_HISTORY. */
function nextEcnNo_() {
  const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
  let max = 0;
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      const m = /^ECN-(\d+)$/i.exec(String(r[0]).trim());
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return 'ECN-' + ('0000' + (max + 1)).slice(-4);
}

/** Finds or creates the Drive folder ECN documents are saved into, next to this spreadsheet. */
function ecnFolder_() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const existing = parent.getFoldersByName(ECN_DOC_FOLDER);
  return existing.hasNext() ? existing.next() : parent.createFolder(ECN_DOC_FOLDER);
}

/** Saves a base64 PDF to the ECN documents folder and returns its Drive URL. */
function saveEcnDocument_(ecnNo, model, base64, mimeType, filename) {
  const bytes = Utilities.base64Decode(base64);
  const safeName = (filename || (ecnNo + '_' + model + '.pdf')).trim();
  const blob = Utilities.newBlob(bytes, mimeType || 'application/pdf', safeName);
  const file = ecnFolder_().createFile(blob);
  file.setName(ecnNo + '_' + model + '_' + safeName);
  return file.getUrl();
}

/** Looks up one part inside a model's current BOM, for the "before part" check in the ECN wizard. */
function getPartInModel(model, partNo) {
  const b = getVerifyBom(model);
  const key = norm_(partNo);
  const found = b.parts.find(p => norm_(p.partNo) === key);
  return {
    locked: b.locked, rev: b.rev,
    found: !!found,
    desc: found ? found.desc : '', img: found ? found.img : '', qps: found ? found.qps : ''
  };
}

/** Looks up one part in PART_MASTER itself, for live-checking new part numbers in the ECN wizard. */
function getMasterPart(partNo) {
  const ctx = readMaster_();
  const i = ctx.idx[norm_(partNo)];
  if (i === undefined) return { found: false, desc: '', img: '' };
  const ex = masterExtra_(ctx, i);
  return { found: true, desc: ex.desc, img: ex.img };
}

/**
 * Combined check for one "after" (new) part number in the ECN wizard: is it
 * in PART_MASTER at all, and - separately - is it already in the selected
 * model's current BOM (and at what QPS). The UI uses the second part to
 * offer "add to" / "subtract from" the existing QPS instead of just
 * overwriting it, for when the new part happens to already be on the BOM.
 */
function getAfterPartInfo(model, partNo) {
  const master = getMasterPart(partNo);
  let inModel = { found: false, qps: '' };
  if (model) {
    try {
      const r = getPartInModel(model, partNo);
      inModel = { found: r.found, qps: r.qps };
    } catch (e) { /* model not found yet, etc. - ignore, treat as not in model */ }
  }
  return { master: master, inModel: inModel };
}

/** Bulk version of getPartInModel - checks several "before" part numbers in one call (one sheet read). */
function getPartsInModel(model, partNos) {
  const b = getVerifyBom(model);
  const byKey = {};
  b.parts.forEach(p => { byKey[norm_(p.partNo)] = p; });
  return (partNos || []).map(partNo => {
    const found = byKey[norm_(partNo)];
    return {
      partNo: partNo, locked: b.locked, rev: b.rev, found: !!found,
      desc: found ? found.desc : '', img: found ? found.img : '', qps: found ? found.qps : ''
    };
  });
}

/** Bulk version of getMasterPart - checks several "after" part numbers in one call (one sheet read). */
function getMasterParts(partNos) {
  const ctx = readMaster_();
  return (partNos || []).map(partNo => {
    const i = ctx.idx[norm_(partNo)];
    if (i === undefined) return { partNo: partNo, found: false, desc: '', img: '' };
    const ex = masterExtra_(ctx, i);
    return { partNo: partNo, found: true, desc: ex.desc, img: ex.img };
  });
}

/** Today's date as yyyy-MM-dd in the spreadsheet's timezone, for the ECN wizard's effective-date default. */
function getTodayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Replaces one or more existing BOM parts with one or more new parts each
 * (a split, a straight swap, or a bulk batch of several old parts at once -
 * e.g. pasted in bulk from Excel). Always generates and logs ONE ECN
 * covering the whole batch, regardless of whether the model is currently
 * locked, and always bumps the model's BOM revision once - this tool is
 * the explicit "make an ECN" path.
 *
 * payload = { model, oldParts:[partNo,...] (or legacy oldPart: string),
 *             newParts:[{partNo,qps}], received (bool),
 *             docBase64, docMime, docName, reason, effectiveDate }
 */
function runEcnChange(payload) {
  return withLock_(() => {
    const ss = dataSS_();
    const model = String(payload.model || '').trim();
    const reason = String(payload.reason || '').trim();
    const effectiveDate = String(payload.effectiveDate || '').trim();
    const newParts = payload.newParts || [];

    // Accept either the new bulk "oldParts" array or the original single
    // "oldPart" string, de-duplicated and with blanks dropped.
    let oldParts = Array.isArray(payload.oldParts) ? payload.oldParts.slice() : [];
    if (!oldParts.length && payload.oldPart) oldParts = [payload.oldPart];
    oldParts = oldParts.map(p => String(p || '').trim()).filter(Boolean);
    const seenOld = {};
    oldParts = oldParts.filter(p => {
      const k = norm_(p);
      if (seenOld[k]) return false;
      seenOld[k] = true;
      return true;
    });

    const sheet = ss.getSheetByName(model);
    if (!sheet || isSystemSheet_(model)) throw new Error('Model sheet not found: ' + model);
    if (!oldParts.length) throw new Error('Enter at least one before (old) part number.');
    if (!newParts.length) throw new Error('Enter at least one after (new) part number.');
    if (!reason) throw new Error('Enter the reason for the change.');
    if (!effectiveDate) throw new Error('Enter the effective date.');
    if (payload.received && !payload.docBase64) throw new Error('Attach the ECN document, or answer "No" if you have not received it yet.');

    const ctx = readMaster_();
    const notFound = oldParts.filter(p => !findInModel_(sheet, p, ctx.mh));
    if (notFound.length)
      throw new Error('Not found in ' + model + '\'s current BOM, so they cannot be replaced: ' +
        notFound.join(', ') + '. Check the part number(s) and try again.');

    // If the same new part number is listed more than once (typed twice, or
    // split across rows by mistake), merge those rows into one by summing
    // their QPS rather than rejecting the ECN.
    const mergedByKey = {};
    const mergedOrder = [];
    newParts.forEach(p => {
      const k = norm_(p.partNo);
      if (!k) throw new Error('One of the new part rows has no part number.');
      const q = Number(p.qps);
      if (!(q > 0)) throw new Error('Enter a valid QPS for ' + p.partNo + '.');
      if (mergedByKey[k]) {
        mergedByKey[k].qps += q;
      } else {
        mergedByKey[k] = { partNo: p.partNo, qps: q };
        mergedOrder.push(k);
      }
    });
    const mergedNewParts = mergedOrder.map(k => mergedByKey[k]);

    // A part listed in BOTH the before and after lists (e.g. "replace P2
    // with P2 + P4", meaning "keep P2, also add P4") must not be deleted -
    // only parts that are being removed outright go to planBom_'s remove list.
    const keptKeys = {};
    mergedNewParts.forEach(p => { keptKeys[norm_(p.partNo)] = true; });
    const removeParts = oldParts.filter(p => !keptKeys[norm_(p)]);

    const plan = planBom_(sheet, ctx, mergedNewParts, removeParts);
    if (plan.missing.length)
      throw new Error('Not in PART_MASTER yet, so they cannot be added: ' + plan.missing.join(', ') +
        '. Add them to PART_MASTER first, then run this ECN again.');

    const ecnNo = nextEcnNo_();
    let docUrl = '';
    if (payload.received && payload.docBase64) {
      docUrl = saveEcnDocument_(ecnNo, model, payload.docBase64, payload.docMime, payload.docName);
    }

    applyBom_(sheet, plan);

    const c = getControl_(model);
    const rev = (c.row ? c.rev : -1) + 1;
    setControl_(model, { status: c.status || 'Unlocked', rev: rev, by: c.by, on: c.on, ecn: ecnNo });

    const rows = plan.changes.map(ch =>
      [now_(), ecnNo, model, rev, ch.partNo, ch.action, ch.oldQ, ch.newQ, reason, effectiveDate, docUrl, user_()]);
    appendRows_(migratedSheet_(ECN_SHEET, ECN_HDR), rows);

    return {
      ok: true, ecnNo: ecnNo, model: model, rev: rev, docUrl: docUrl,
      removed: removeParts, added: mergedNewParts.map(p => p.partNo + ' (QPS ' + p.qps + ')')
    };
  });
}

/** Recent ECN_HISTORY rows, newest first, optionally filtered to one model. */
function getEcnLog(model) {
  const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
  if (sh.getLastRow() < 2) return [];
  const want = model ? norm_(model) : '';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, ECN_HDR.length).getValues()
    .filter(r => !want || norm_(r[2]) === want)
    .map(r => ({
      when: txt_(r[0]), ecnNo: txt_(r[1]), model: txt_(r[2]), rev: txt_(r[3]), part: txt_(r[4]),
      action: txt_(r[5]), oldQ: txt_(r[6]), newQ: txt_(r[7]), reason: txt_(r[8]),
      effectiveDate: txt_(r[9]), docUrl: txt_(r[10]), by: txt_(r[11])
    }));
  return rows.reverse().slice(0, 300);
}

/**
 * All ECN_HISTORY rows sharing one ECN number, grouped for the "Edit ECN"
 * panel: the shared metadata (model, reason, effective date, document) plus
 * the list of individual part changes it covers (read-only - editing here
 * only corrects the metadata, never re-applies the BOM change).
 */
function getEcnByNo(ecnNo) {
  const want = norm_(ecnNo);
  const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
  if (sh.getLastRow() < 2) return null;
  const all = sh.getRange(2, 1, sh.getLastRow() - 1, ECN_HDR.length).getValues();
  const mine = all.filter(r => norm_(r[1]) === want);
  if (!mine.length) return null;
  const first = mine[0];
  return {
    ecnNo: txt_(first[1]), model: txt_(first[2]), rev: txt_(first[3]),
    reason: txt_(first[8]), effectiveDate: txt_(first[9]), docUrl: txt_(first[10]), by: txt_(first[11]),
    changes: mine.map(r => ({
      part: txt_(r[4]), action: txt_(r[5]), oldQ: txt_(r[6]), newQ: txt_(r[7]), when: txt_(r[0])
    }))
  };
}

/**
 * Data for a formatted, printable ECN report: one entry per ECN (not per
 * part-change row), with its Before (old/removed) parts and After
 * (new/added, or qty-changed-to) parts separated out - each with its
 * PART_MASTER description alongside the part number - plus the reason,
 * effective date, the BOM revision it produced, who raised it and the
 * saved document link. Newest first.
 *
 * model: filter to just that model, or '' / omit for every model.
 * ecnNo: filter to just that one ECN number, or '' / omit for every ECN
 *        (within whatever model filter also applies).
 */
function getEcnReportData(model, ecnNo) {
  const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
  if (sh.getLastRow() < 2) return [];
  const want = model ? norm_(model) : '';
  const wantEcn = ecnNo ? norm_(ecnNo) : '';
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, ECN_HDR.length).getValues()
    .filter(r => (!want || norm_(r[2]) === want) && (!wantEcn || norm_(r[1]) === wantEcn) && String(r[4]).trim()); // part-level rows only

  const ctx = readMaster_();
  const descOf = partNo => { const i = ctx.idx[norm_(partNo)]; return i === undefined ? '' : masterExtra_(ctx, i).desc; };

  const groups = {};
  const order = [];
  rows.forEach(r => {
    const no = txt_(r[1]);
    if (!groups[no]) {
      groups[no] = {
        ecnNo: no, model: txt_(r[2]), rev: Number(r[3]) || 0, when: txt_(r[0]),
        reasons: [], effectiveDate: txt_(r[9]), docUrl: txt_(r[10]), by: txt_(r[11]),
        before: [], after: []
      };
      order.push(no);
    }
    const g = groups[no];
    if (txt_(r[0]) > g.when) g.when = txt_(r[0]); // latest timestamp among this ECN's rows
    const reason = txt_(r[8]);
    if (reason && g.reasons.indexOf(reason) < 0) g.reasons.push(reason); // several when ECNs were merged
    const partNo = txt_(r[4]), action = txt_(r[5]), oldQ = r[6], newQ = r[7], desc = descOf(partNo);
    if (action === 'Removed') {
      g.before.push({ partNo: partNo, desc: desc, qps: oldQ, action: action });
    } else if (action === 'Added') {
      g.after.push({ partNo: partNo, desc: desc, qps: newQ, action: action });
    } else { // 'QPS changed' - the same part, before AND after at different quantities
      g.before.push({ partNo: partNo, desc: desc, qps: oldQ, action: action });
      g.after.push({ partNo: partNo, desc: desc, qps: newQ, action: action });
    }
  });

  return order.map(no => {
    const g = groups[no];
    g.reason = g.reasons.join(' | ');
    delete g.reasons;
    return g;
  }).sort((a, b) => (b.rev - a.rev) || (a.when < b.when ? 1 : -1));
}

/** Every ECN number on record (optionally for one model), for the report's ECN picker. */
function getAllEcnNos(model) {
  const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
  if (sh.getLastRow() < 2) return [];
  const want = model ? norm_(model) : '';
  const seen = {};
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, ECN_HDR.length).getValues().forEach(r => {
    const no = txt_(r[1]);
    if (!no || !String(r[4]).trim()) return;
    if (want && norm_(r[2]) !== want) return;
    if (seen[no]) return;
    seen[no] = true;
    out.push({ ecnNo: no, model: txt_(r[2]) });
  });
  return out.sort((a, b) => a.ecnNo < b.ecnNo ? 1 : -1);
}

/**
 * Corrects an ECN's metadata (reason / effective date / document) after the
 * fact, across every ECN_HISTORY row filed under that ECN number. Does NOT
 * touch which parts were added or removed - that record is the audit trail
 * of what actually happened on the sheet, and reopening it would break the
 * link between ECN_HISTORY and the BOM. To change the parts themselves,
 * generate a new ECN instead.
 *
 * payload = { ecnNo, reason, effectiveDate, received(bool), docBase64,
 *             docMime, docName, keepExistingDoc(bool) }
 */
function updateEcnMeta(payload) {
  return withLock_(() => {
    const ecnNo = String(payload.ecnNo || '').trim();
    const reason = String(payload.reason || '').trim();
    const effectiveDate = String(payload.effectiveDate || '').trim();
    if (!ecnNo) throw new Error('Missing ECN number.');
    if (!reason) throw new Error('Enter the reason for the change.');
    if (!effectiveDate) throw new Error('Enter the effective date.');
    if (payload.received && !payload.keepExistingDoc && !payload.docBase64)
      throw new Error('Attach the ECN document, or keep the existing one.');

    const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('ECN ' + ecnNo + ' was not found.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const want = norm_(ecnNo);
    const idxs = [];
    vals.forEach((r, i) => { if (norm_(r[1]) === want) idxs.push(i); });
    if (!idxs.length) throw new Error('ECN ' + ecnNo + ' was not found.');

    const model = txt_(vals[idxs[0]][2]);
    let docUrl = txt_(vals[idxs[0]][10]);
    if (payload.received && payload.docBase64) {
      docUrl = saveEcnDocument_(ecnNo, model, payload.docBase64, payload.docMime, payload.docName);
    } else if (!payload.received) {
      docUrl = '';
    }

    idxs.forEach(i => {
      vals[i][8] = reason;        // Description / reason
      vals[i][9] = effectiveDate; // Effective_Date
      vals[i][10] = docUrl;       // Document_Link
    });
    rng.setValues(vals);

    return { ok: true, ecnNo: ecnNo, model: model, docUrl: docUrl, rowsUpdated: idxs.length };
  });
}

/**
 * Deletes an ECN that was done by mistake: reverses what it changed on the
 * model's live BOM (an "Added" part is removed again, a "Removed" part is
 * put back at its old QPS, a "QPS changed" part is put back to its old
 * QPS), then deletes its rows from ECN_HISTORY and renumbers later
 * revisions to close the gap.
 *
 * This works on ANY ECN for the model, not only the latest one. It is safe
 * to do even when a later ECN exists, as long as that later ECN never
 * touched the same part(s) - in that case the live sheet still holds
 * exactly what this ECN set those parts to, so putting them back is
 * correct. If a later ECN DID touch the same part(s), reversing this one
 * blindly could clobber that later change, so it's refused instead, naming
 * which ECN(s) are in the way.
 */
function deleteEcn(ecnNo) {
  return withLock_(() => {
    const want = norm_(ecnNo);
    const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('ECN ' + ecnNo + ' was not found.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const idxs = [];
    vals.forEach((r, i) => { if (norm_(r[1]) === want) idxs.push(i); });
    if (!idxs.length) throw new Error('ECN ' + ecnNo + ' was not found.');

    const model = txt_(vals[idxs[0]][2]);
    const ecnRev = Number(vals[idxs[0]][3]) || 0;
    const sheet = dataSS_().getSheetByName(model);
    if (!sheet || isSystemSheet_(model)) throw new Error('Model sheet not found: ' + model);

    const hasPartRows = idxs.some(i => String(vals[i][4]).trim());
    if (!hasPartRows) throw new Error('ECN ' + ecnNo + ' has no part changes to undo.');

    const modelWant = norm_(model);
    const touchedParts = {};
    idxs.forEach(i => { const p = txt_(vals[i][4]); if (p) touchedParts[norm_(p)] = true; });

    // Any LATER row for this model (a different ECN, a higher Rev) that touched the same part(s)?
    const conflicts = {};
    vals.forEach((r, i) => {
      if (idxs.indexOf(i) >= 0) return;
      if (norm_(r[2]) !== modelWant) return;
      if (Number(r[3]) <= ecnRev) return;
      const p = txt_(r[4]);
      if (p && touchedParts[norm_(p)]) conflicts[txt_(r[1])] = true;
    });
    const conflictList = Object.keys(conflicts);
    if (conflictList.length)
      throw new Error('Cannot delete ' + ecnNo + ' - ' + conflictList.join(', ') +
        ' changed the same part(s) afterwards (' + Object.keys(touchedParts).join(', ') + '). ' +
        'Delete ' + (conflictList.length > 1 ? 'those ECNs' : conflictList[0]) + ' first (most recent first), ' +
        'or use "BOM version history" to restore an earlier revision instead.');

    const ctx = readMaster_();
    const toAdd = [], toRemove = [];
    idxs.forEach(i => {
      const r = vals[i];
      const partNo = txt_(r[4]);
      if (!partNo) return;
      const action = txt_(r[5]);
      const oldQ = r[6];
      if (action === 'Added') toRemove.push(partNo);
      else toAdd.push({ partNo: partNo, qps: oldQ }); // "Removed" or "QPS changed" -> put old QPS back
    });

    const plan = planBom_(sheet, ctx, toAdd, toRemove);
    if (plan.missing.length)
      throw new Error('Cannot undo ' + ecnNo + ' - part(s) no longer in PART_MASTER: ' + plan.missing.join(', '));
    applyBom_(sheet, plan);

    const keepRows = vals.filter((r, i) => idxs.indexOf(i) < 0);

    // Close the numbering gap: if other rows still occupy ecnRev (e.g. it was a merged
    // revision with more than one ECN in it), later revisions keep their numbers as-is;
    // otherwise everything after ecnRev shifts down by one.
    const stillAtEcnRev = keepRows.some(r => norm_(r[2]) === modelWant && Number(r[3]) === ecnRev);
    const shiftAmount = stillAtEcnRev ? 0 : 1;
    keepRows.forEach(r => {
      if (norm_(r[2]) !== modelWant) return;
      const rv = Number(r[3]);
      if (rv > ecnRev) r[3] = rv - shiftAmount;
    });

    sh.getRange(2, 1, last - 1, ECN_HDR.length).clearContent();
    if (keepRows.length) sh.getRange(2, 1, keepRows.length, ECN_HDR.length).setValues(keepRows);

    const c = getControl_(model);
    let newRev;
    if (c.rev === ecnRev) newRev = stillAtEcnRev ? ecnRev : ecnRev - 1;
    else if (c.rev > ecnRev) newRev = c.rev - shiftAmount;
    else newRev = c.rev;

    let lastEcnNo = '';
    keepRows.forEach(r => { if (norm_(r[2]) === modelWant && Number(r[3]) === newRev) lastEcnNo = txt_(r[1]); });
    setControl_(model, { status: c.status || 'Unlocked', rev: newRev, by: c.by, on: c.on, ecn: lastEcnNo });

    return { ok: true, ecnNo: ecnNo, model: model, newRev: newRev, partsReverted: idxs.length };
  });
}

/**
 * Removes an ECN's rows from ECN_HISTORY WITHOUT touching the live BOM at
 * all - for cleaning up a duplicate/accidental entry whose changes are
 * already reflected by another ECN (e.g. the same change got submitted
 * twice), where the current BOM is already correct and only the extra
 * history record needs to go. Unlike deleteEcn(), this does nothing to the
 * model's sheet, so it never conflicts with later ECNs - but for that same
 * reason it's on the caller to be sure the ECN being removed is genuinely
 * redundant. If its changes are NOT duplicated elsewhere, removing its
 * record this way leaves the BOM looking like it changed with no ECN to
 * explain why. Still closes the revision-numbering gap it leaves behind.
 */
function deleteEcnRecordOnly(ecnNo) {
  return withLock_(() => {
    const want = norm_(ecnNo);
    const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('ECN ' + ecnNo + ' was not found.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const idxs = [];
    vals.forEach((r, i) => { if (norm_(r[1]) === want) idxs.push(i); });
    if (!idxs.length) throw new Error('ECN ' + ecnNo + ' was not found.');

    const model = txt_(vals[idxs[0]][2]);
    const ecnRev = Number(vals[idxs[0]][3]) || 0;
    const modelWant = norm_(model);

    const keepRows = vals.filter((r, i) => idxs.indexOf(i) < 0);

    const stillAtEcnRev = keepRows.some(r => norm_(r[2]) === modelWant && Number(r[3]) === ecnRev);
    const shiftAmount = stillAtEcnRev ? 0 : 1;
    keepRows.forEach(r => {
      if (norm_(r[2]) !== modelWant) return;
      const rv = Number(r[3]);
      if (rv > ecnRev) r[3] = rv - shiftAmount;
    });

    sh.getRange(2, 1, last - 1, ECN_HDR.length).clearContent();
    if (keepRows.length) sh.getRange(2, 1, keepRows.length, ECN_HDR.length).setValues(keepRows);

    const c = getControl_(model);
    let newRev;
    if (c.rev === ecnRev) newRev = stillAtEcnRev ? ecnRev : ecnRev - 1;
    else if (c.rev > ecnRev) newRev = c.rev - shiftAmount;
    else newRev = c.rev;

    let lastEcnNo = '';
    keepRows.forEach(r => { if (norm_(r[2]) === modelWant && Number(r[3]) === newRev) lastEcnNo = txt_(r[1]); });
    setControl_(model, { status: c.status || 'Unlocked', rev: newRev, by: c.by, on: c.on, ecn: lastEcnNo });

    return { ok: true, ecnNo: ecnNo, model: model, newRev: newRev, rowsRemoved: idxs.length };
  });
}

// ======================================================================
//              BOM VERSION HISTORY (view / restore an old revision)
// ======================================================================

/** Every part-level ECN_HISTORY row for one model, in the order they were logged. */
function ecnPartRowsFor_(model) {
  const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
  if (sh.getLastRow() < 2) return [];
  const want = norm_(model);
  return sh.getRange(2, 1, sh.getLastRow() - 1, ECN_HDR.length).getValues()
    .filter(r => norm_(r[2]) === want && String(r[4]).trim()) // same model, Part_No not blank
    .map(r => ({
      when: txt_(r[0]), ecnNo: txt_(r[1]), rev: Number(r[3]) || 0, partNo: txt_(r[4]),
      action: txt_(r[5]), oldQ: r[6], newQ: r[7], reason: txt_(r[8])
    }));
}

/**
 * Every revision checkpoint recorded for a model (one entry per distinct
 * Rev value in ECN_HISTORY), newest first, each showing the ECN(s) and
 * reason(s) that produced it and how many parts it touched. The live
 * current BOM (its rev, from BOM_CONTROL) is flagged separately so the UI
 * can mark it "(current)".
 */
function getBomRevisions(model) {
  const rows = ecnPartRowsFor_(model);
  const byRev = {};
  const order = [];
  rows.forEach(r => {
    if (!byRev[r.rev]) { byRev[r.rev] = { rev: r.rev, when: r.when, ecnNos: [], reasons: [], count: 0 }; order.push(r.rev); }
    const g = byRev[r.rev];
    g.when = r.when; // keep the latest timestamp seen for this rev
    if (g.ecnNos.indexOf(r.ecnNo) < 0) g.ecnNos.push(r.ecnNo);
    if (g.reasons.indexOf(r.reason) < 0) g.reasons.push(r.reason);
    g.count++;
  });
  const curRev = getControl_(model).rev;
  return order.sort((a, b) => b - a).map(rev => {
    const g = byRev[rev];
    return { rev: rev, when: g.when, ecnNo: g.ecnNos.join(', '), reason: g.reasons.join(' | '), partsChanged: g.count, isCurrent: rev === curRev };
  });
}

/**
 * Reconstructs the model's BOM as it stood at a given revision by replaying
 * every part-level ECN_HISTORY row up to and including that revision, in
 * order. Description/image come from PART_MASTER as it is today (those
 * aren't tracked per-revision), so only part numbers and QPS are historical.
 */
function getBomAtRevision(model, rev) {
  const target = Number(rev);
  const ctx = readMaster_();
  const map = {}; // norm partNo -> { partNo, qps }
  let lastWhen = '', lastEcn = '';
  ecnPartRowsFor_(model).forEach(r => {
    if (r.rev > target) return;
    const key = norm_(r.partNo);
    if (r.action === 'Removed') {
      delete map[key];
    } else {
      map[key] = { partNo: r.partNo, qps: r.newQ };
    }
    lastWhen = r.when; lastEcn = r.ecnNo;
  });
  const parts = Object.keys(map).map(k => {
    const row = map[k];
    const ex = masterExtra_(ctx, ctx.idx[k]);
    return { partNo: row.partNo, desc: ex.desc, img: ex.img, qps: row.qps };
  }).sort((a, b) => a.partNo.localeCompare(b.partNo));
  return { rev: target, model: model, asOfWhen: lastWhen, asOfEcn: lastEcn, parts: parts };
}

/** CSV download of a model's BOM as it stood at a given historical revision. */
function getBomRevisionCsv(model, rev) {
  const snap = getBomAtRevision(model, rev);
  const rowsOut = [['Part No', 'Description', 'QPS']];
  snap.parts.forEach(p => rowsOut.push([p.partNo, p.desc, p.qps]));
  const csv = rowsOut.map(r => r.map(csvField_).join(',')).join('\r\n');
  return { csv: csv, filename: model + '_rev' + rev + '_BOM.csv' };
}

/**
 * Restores a model's live BOM to match how it stood at an older revision.
 * This does not rewind BOM_CONTROL's revision number - it computes the
 * difference between the target revision and today's BOM and applies that
 * as a brand-new ECN (so the restore itself is fully audited, same as any
 * other BOM change), which bumps the revision forward by one as usual.
 */
function restoreBomVersion(model, rev, reason) {
  return withLock_(() => {
    const ss = dataSS_();
    const sheet = ss.getSheetByName(model);
    if (!sheet || isSystemSheet_(model)) throw new Error('Model sheet not found: ' + model);
    const r = String(reason || '').trim() || ('Restored to Rev ' + rev);

    const target = getBomAtRevision(model, rev);
    const current = getVerifyBom(model);
    const targetKeys = {};
    target.parts.forEach(p => { targetKeys[norm_(p.partNo)] = true; });
    const removeList = current.parts
      .filter(p => !targetKeys[norm_(p.partNo)])
      .map(p => p.partNo);

    const ctx = readMaster_();
    const plan = planBom_(sheet, ctx, target.parts, removeList);
    if (plan.missing.length)
      throw new Error('Revision ' + rev + ' includes part(s) no longer in PART_MASTER: ' + plan.missing.join(', ') +
        '. Add them back to PART_MASTER first, then restore again.');
    if (!plan.changes.length)
      return { ok: true, ecnNo: '', model: model, rev: current.rev, changes: 0, note: 'Already matches revision ' + rev + ' - nothing to restore.' };

    applyBom_(sheet, plan);

    const ecnNo = nextEcnNo_();
    const c = getControl_(model);
    const newRev = (c.row ? c.rev : -1) + 1;
    setControl_(model, { status: c.status || 'Unlocked', rev: newRev, by: c.by, on: c.on, ecn: ecnNo });

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const rows = plan.changes.map(ch =>
      [now_(), ecnNo, model, newRev, ch.partNo, ch.action, ch.oldQ, ch.newQ, r, today, '', user_()]);
    appendRows_(migratedSheet_(ECN_SHEET, ECN_HDR), rows);

    return { ok: true, ecnNo: ecnNo, model: model, rev: newRev, restoredToRev: rev, changes: plan.changes.length };
  });
}

/**
 * Merges two or more ADJACENT (consecutive) revisions of a model's BOM
 * history into one: the same single revision number, AND the same single
 * ECN number, so "BOM version history" shows one combined entry and the
 * "ECN history / edit" tab in Generate ECN shows one combined ECN row
 * (with every part change that happened across the merged revisions listed
 * under it) instead of several. The individual part-level change rows
 * themselves are kept exactly as they happened - only the Rev and ECN_No
 * they're filed under change. The lowest-numbered of the ECNs being merged
 * is kept as the surviving ECN number (its reason/date/document become the
 * combined ECN's); every later revision number shifts down to stay
 * contiguous. The live BOM itself is never touched, because the final part
 * list only depends on the order the changes happened in, not how they
 * happen to be numbered or filed. This cannot be undone.
 */
function mergeBomRevisions(model, revs) {
  return withLock_(() => {
    model = String(model || '').trim();
    if (!model) throw new Error('Missing model.');
    let list = (revs || []).map(Number).filter(n => !isNaN(n));
    list = Array.from(new Set(list)).sort((a, b) => a - b);
    if (list.length < 2) throw new Error('Pick at least two revisions to merge.');
    for (let i = 1; i < list.length; i++) {
      if (list[i] !== list[i - 1] + 1)
        throw new Error('Only adjacent revisions can be merged together (e.g. Rev 2 and Rev 3) - not revisions with a gap between them.');
    }
    const minMerged = list[0];
    const maxMerged = list[list.length - 1];

    const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
    const last = sh.getLastRow();
    if (last < 2) throw new Error('No ECN history found for ' + model + '.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const modelWant = norm_(model);

    const foundRevs = {};
    vals.forEach(r => { if (norm_(r[2]) === modelWant) foundRevs[Number(r[3])] = true; });
    const missing = list.filter(rv => !foundRevs[rv]);
    if (missing.length) throw new Error('Revision(s) not found for ' + model + ': ' + missing.join(', '));

    // The surviving ECN number: the lowest-numbered ECN among the merged revisions.
    const ecnNum_ = no => { const m = /(\d+)\s*$/.exec(String(no || '')); return m ? Number(m[1]) : Infinity; };
    let targetEcnNo = '';
    vals.forEach(r => {
      if (norm_(r[2]) !== modelWant) return;
      const rv = Number(r[3]);
      if (rv < minMerged || rv > maxMerged) return;
      const no = txt_(r[1]);
      if (!targetEcnNo || ecnNum_(no) < ecnNum_(targetEcnNo)) targetEcnNo = no;
    });

    const mergedEcnNos = new Set();
    const shift = rv => {
      if (rv < minMerged) return rv;
      if (rv <= maxMerged) return minMerged;
      return rv - (list.length - 1);
    };
    vals.forEach(r => {
      if (norm_(r[2]) !== modelWant) return;
      const origRev = Number(r[3]);
      if (origRev >= minMerged && origRev <= maxMerged) {
        mergedEcnNos.add(txt_(r[1]));
        r[1] = targetEcnNo; // consolidate every row in range under the one surviving ECN number
      }
      r[3] = shift(origRev);
    });
    rng.setValues(vals);

    const c = getControl_(model);
    const newCurRev = shift(c.rev);
    const newLastEcn = (c.rev >= minMerged && c.rev <= maxMerged) ? targetEcnNo : c.ecn;
    if (newCurRev !== c.rev || newLastEcn !== c.ecn)
      setControl_(model, { status: c.status || 'Unlocked', rev: newCurRev, by: c.by, on: c.on, ecn: newLastEcn });

    return {
      ok: true, model: model, mergedInto: minMerged,
      targetEcnNo: targetEcnNo, revisionsMerged: list.length, ecnsMerged: mergedEcnNos.size, newCurrentRev: newCurRev
    };
  });
}

// ======================================================================
//                     ASSIGN ONE PART TO MODELS (sidebar)
// ======================================================================
/**
 * Web app entry point.
 *   .../exec               -> launcher page with links to every screen
 *   .../exec?page=bom      -> Build model BOM
 *   .../exec?page=lock     -> BOM lock / unlock
 *   .../exec?page=verify   -> BOM verification (lot-wise)
 *   .../exec?page=compare  -> Compare two models
 *   .../exec?page=ecn      -> Generate ECN
 *   .../exec?page=history  -> BOM version history
 *   .../exec?page=report   -> ECN report
 */
function doGet(e) {
  const pages = {
    bom:     { title: 'Build model BOM',         html: getBomHtml_ },
    lock:    { title: 'BOM lock / unlock',       html: getControlHtml_ },
    verify:  { title: 'BOM verification',        html: getVerifyHtml_ },
    compare: { title: 'Compare two models',      html: getTwoCompareHtml_ },
    ecn:     { title: 'Generate ECN',            html: getEcnHtml_ },
    history: { title: 'BOM version history',     html: getBomHistoryHtml_ },
    report:  { title: 'ECN report',              html: getEcnReportHtml_ },
    assign:  { title: 'Assign part to models',   html: getAssignWebHtml_ },
    compareall: { title: 'Compare model BOMs (all models)', html: getCompareAllHtml_ },
    images:  { title: 'Part images and clean-up', html: getImagesToolsHtml_ }
  };

  // Portal users with the BOM management module (admins always have it), or the script owner.
  const token = String((e && e.parameter && e.parameter.t) || '');
  const sess = portalSession_(token);
  if (!((sess && sess.bomManagement) || ownerVisitor_())) {
    // Plain text, not an HTML page: an HtmlService page would let anyone call this script's
    // server functions from the browser console, so people without access get no page at all.
    return ContentService.createTextOutput(
      'BOM management needs the BOM management module. Sign in to LUMIO Quality Management - Projector ' +
      'and open BOM management from there, or ask an admin for access.' + (sess ? '' : ' If you were signed in, your session may have expired.'))
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const key = String((e && e.parameter && e.parameter.page) || '').toLowerCase();
  let body, title;

  if (pages[key]) {
    body = pages[key].html();
    title = pages[key].title;
  } else {
    const base = ScriptApp.getService().getUrl();
    const links = Object.keys(pages).map(k =>
      '<a href="' + base + '?page=' + k + (token ? '&t=' + encodeURIComponent(token) : '') + '" target="_top" ' +
      'style="display:block;padding:14px 18px;margin:8px 0;background:#fff;border:1px solid #dadce0;' +
      'border-radius:10px;color:#1a73e8;font-weight:600;text-decoration:none">' + pages[k].title + '</a>'
    ).join('');
    body = '<!DOCTYPE html><html><head><base target="_top"></head>' +
      '<body style="font-family:Segoe UI,Roboto,Arial,sans-serif;background:#f6f8fb;margin:0;padding:24px">' +
      '<div style="max-width:420px;margin:0 auto"><h2 style="margin:0 0 12px">Part Reference Database</h2>' +
      links + '</div></body></html>';
    title = 'Part DB';
  }

  return HtmlService.createHtmlOutput(body)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function showAssignSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml_()).setTitle('Assign part to models');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Called by the sidebar: what part is selected and where is it used? */
function getSelectedPart() {
  const ss    = dataSS_();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME)
    return { ok: false, msg: 'Go to ' + SHEET_NAME + ' and click a part row.' };

  const row = sheet.getActiveCell().getRow();
  if (row <= HEADER_ROW) return { ok: false, msg: 'Click a part row (not the header).' };

  const partNo = String(sheet.getRange(row, COL_PART_NO).getValue()).trim();
  if (!partNo) return { ok: false, msg: 'Selected row has no Part No.' };

  const mh   = headersOf_(sheet);
  const ctrl = controlMap_();
  const models = getModelSheets_().map(s => {
    const f = findInModel_(s, partNo, mh);
    const c = ctrl[norm_(s.getName())];
    return { name: s.getName(), listed: !!f, qps: f ? f.qps : '', locked: !!(c && c.status === 'Locked') };
  });
  return { ok: true, partNo: partNo, row: row, models: models };
}

/** payload = { partNo, add:[{model,qps}], remove:[model], ecn:{no,desc} } */
function assignPart(payload) {
  return withLock_(() => {
    const ss  = dataSS_();
    const ctx = readMaster_();
    if (!(norm_(payload.partNo) in ctx.idx))
      throw new Error('Part ' + payload.partNo + ' not found in ' + SHEET_NAME);

    const per = {};   // model -> { items, remove }
    (payload.add || []).forEach(a => {
      (per[a.model] = per[a.model] || { items: [], remove: [] }).items.push({ partNo: payload.partNo, qps: a.qps });
    });
    (payload.remove || []).forEach(m => {
      (per[m] = per[m] || { items: [], remove: [] }).remove.push(payload.partNo);
    });

    const list = Object.keys(per).map(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || isSystemSheet_(name)) throw new Error('Sheet not found: ' + name);
      return { sheet: sheet, plan: planBom_(sheet, ctx, per[name].items, per[name].remove) };
    });

    commitPlans_(list, payload.ecn);

    const done = [];
    list.forEach(x => {
      const name = x.sheet.getName();
      if (!x.plan.changes.length) { done.push(name + ': no change'); return; }
      x.plan.changes.forEach(c => {
        if (c.action === 'Added') done.push(name + ': added with QPS ' + c.newQ);
        else if (c.action === 'Removed') done.push(name + ': removed');
        else done.push(name + ': QPS ' + c.oldQ + ' -> ' + c.newQ);
      });
    });
    return { ok: true, done: done };
  });
}

// ======================================================================
//                  BULK BOM BUILDER (many parts -> one model)
// ======================================================================
function showBomDialog() {
  const html = HtmlService.createHtmlOutput(getBomHtml_()).setWidth(1600).setHeight(980);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Build model BOM');
}

/** All model sheet names + every part in PART_MASTER (with description + image). */
function getBomData() {
  const ctx = readMaster_();
  const parts = [], seen = {};
  ctx.mVals.forEach((r, i) => {
    const p = String(r[COL_PART_NO - 1]).trim();
    if (!p || seen[norm_(p)]) return;
    seen[norm_(p)] = true;
    const ex = masterExtra_(ctx, i);
    parts.push({ partNo: p, desc: ex.desc, img: ex.img });
  });
  return { models: getModelSheets_().map(s => s.getName()), parts: parts };
}

/** { parts:{ PART_UPPER: qps }, locked, rev } for one model sheet. */
function getModelParts(name) {
  const ss    = dataSS_();
  const sheet = ss.getSheetByName(name);
  const out   = {};
  const c     = getControl_(name);
  if (!sheet || sheet.getLastRow() <= HEADER_ROW)
    return { parts: out, locked: c.status === 'Locked', rev: c.rev };

  const mh = headersOf_(ss.getSheetByName(SHEET_NAME));
  const pc = partColIn_(sheet, mh);
  const qc = headersOf_(sheet).findIndex(h => norm_(h) === QPS_HEADER) + 1;
  const n  = sheet.getLastRow() - HEADER_ROW;
  const pv = sheet.getRange(HEADER_ROW + 1, pc, n, 1).getValues();
  const qv = qc ? sheet.getRange(HEADER_ROW + 1, qc, n, 1).getValues() : null;
  pv.forEach((r, i) => {
    const k = norm_(r[0]);
    if (k) out[k] = qv ? qv[i][0] : '';
  });
  return { parts: out, locked: c.status === 'Locked', rev: c.rev };
}

/** payload = { model, newModel, items:[{partNo,qps}], remove:[partNo], ecn:{no,desc} } */
function saveBom(payload) {
  return withLock_(() => {
    const ss  = dataSS_();
    const ctx = readMaster_();

    let sheet, created = false;
    const nm = String(payload.newModel || '').trim();
    if (nm) {
      if (isSystemSheet_(nm)) throw new Error('"' + nm + '" is a reserved sheet name.');
      if (ss.getSheetByName(nm)) throw new Error('A sheet named "' + nm + '" already exists.');
      sheet = ss.insertSheet(nm);
      created = true;
    } else {
      sheet = ss.getSheetByName(payload.model);
      if (!sheet || isSystemSheet_(payload.model)) throw new Error('Model sheet not found: ' + payload.model);
    }

    try {
      const plan = planBom_(sheet, ctx, payload.items || [], payload.remove || []);
      commitPlans_([{ sheet: sheet, plan: plan }], payload.ecn);
      const cnt = a => plan.changes.filter(c => c.action === a).length;
      return {
        ok: true, model: sheet.getName(),
        models: getModelSheets_().map(s => s.getName()),
        added: cnt('Added'), updated: cnt('QPS changed'), removed: cnt('Removed'),
        missing: plan.missing
      };
    } catch (err) {
      if (created) ss.deleteSheet(sheet);
      throw err;
    }
  });
}

// ======================================================================
//                    LOT-WISE BOM VERIFICATION
// ======================================================================
function showVerifyDialog() {
  const html = HtmlService.createHtmlOutput(getVerifyHtml_()).setWidth(1600).setHeight(980);
  SpreadsheetApp.getUi().showModelessDialog(html, 'BOM verification (lot-wise)');
}

/** Model list for the dialog, with lock state. */
function getVerifyModels() {
  const ctrl = controlMap_();
  return getModelSheets_().map(s => {
    const c = ctrl[norm_(s.getName())];
    return { name: s.getName(), locked: !!(c && c.status === 'Locked'), rev: c ? c.rev : 0 };
  });
}

/** Current BOM of a model: [{partNo, desc, img, qps}] + lock info. */
function getVerifyBom(model) {
  const ss    = dataSS_();
  const sheet = ss.getSheetByName(model);
  if (!sheet || isSystemSheet_(model)) throw new Error('Model sheet not found: ' + model);
  const ctx = readMaster_();
  const c   = getControl_(model);
  const parts = [];
  const n = sheet.getLastRow() - HEADER_ROW;
  if (n > 0) {
    const pc = partColIn_(sheet, ctx.mh);
    const qc = headersOf_(sheet).findIndex(h => norm_(h) === QPS_HEADER) + 1;
    const pv = sheet.getRange(HEADER_ROW + 1, pc, n, 1).getValues();
    const qv = qc ? sheet.getRange(HEADER_ROW + 1, qc, n, 1).getValues() : null;
    pv.forEach((r, i) => {
      const p = String(r[0]).trim();
      if (!p) return;
      const ex = masterExtra_(ctx, ctx.idx[norm_(p)]);
      parts.push({ partNo: p, desc: ex.desc, img: ex.img, qps: qv ? qv[i][0] : '' });
    });
  }
  return { locked: c.status === 'Locked', rev: c.rev, parts: parts };
}

// BOM-interface style column order for the lot detail log: identifying
// columns first, then the same fields the BOM builder shows (Part No,
// Description, Image, QPS), then what the lot actually used, with Result last.
const LOT_HDR = ['Ver_ID', 'Model', 'Lot_No', 'BOM_Rev', 'Part_No', 'Description', 'Image_URL',
                 'BOM_QPS', 'Lot_Part_No', 'Lot_QPS', 'Remark', 'Result'];

/**
 * p = { model, lot, by, remarks, rows:[{partNo, desc, qps, ok, actualPart, actualQps, remark, extra}] }
 * Appends one line to BOM_VERIFY_LOG and one line per part to LOT_BOM_DETAIL.
 */
function saveVerification(p) {
  return withLock_(() => {
    const ss    = dataSS_();
    const model = String(p.model || '').trim();
    const lot   = String(p.lot || '').trim();
    if (!model || !ss.getSheetByName(model)) throw new Error('Choose a model.');
    if (!lot) throw new Error('Enter the Lot No.');

    const c    = getControl_(model);
    const vlog = sysSheet_(VERIFY_SHEET, VERIFY_HDR);
    const dsh  = migratedSheet_(LOT_SHEET, LOT_HDR);
    const verId = 'VER-' + ('0000' + vlog.getLastRow()).slice(-4);

    const detail = [];
    let deviations = 0, checked = 0;

    (p.rows || []).forEach(r => {
      const remark = String(r.remark || '').trim();
      let lotPart, lotQps, result, bomPart = String(r.partNo || '').trim(), bomQps = r.qps,
          desc = String(r.desc || '').trim(), img = String(r.img || '').trim();

      if (r.extra) {
        lotPart = String(r.actualPart || '').trim();
        if (!lotPart) return;               // empty extra row - ignore
        lotQps = r.actualQps === '' || r.actualQps === null ? '' : Number(r.actualQps);
        result = 'Extra part';
        bomPart = ''; bomQps = ''; desc = ''; img = '';
      } else {
        checked++;
        if (r.ok) {
          lotPart = bomPart; lotQps = bomQps; result = 'Matched';
        } else {
          lotPart = String(r.actualPart || '').trim();
          const aq = String(r.actualQps === null || r.actualQps === undefined ? '' : r.actualQps).trim();
          lotQps = aq === '' ? bomQps : Number(aq);
          if (!lotPart)                              { result = 'Not used in lot'; lotQps = ''; }
          else if (norm_(lotPart) !== norm_(bomPart)) result = 'Part changed';
          else if (Number(lotQps) !== Number(bomQps)) result = 'QPS changed';
          else                                        result = 'Matched';
        }
      }
      if (result !== 'Matched') deviations++;
      detail.push([verId, model, lot, c.rev, bomPart, desc, img, bomQps, lotPart, lotQps, remark, result]);
    });

    if (!detail.length) throw new Error('This model has no BOM parts to verify.');
    const overall = deviations === 0 ? 'All matched' : 'Deviation';

    appendRows_(dsh, detail);
    appendRows_(vlog, [[verId, now_(), model, lot, c.rev, c.status === 'Locked' ? 'Locked' : 'Not locked',
      checked, deviations, overall, String(p.by || '').trim() || user_(), String(p.remarks || '').trim()]]);
    return { ok: true, verId: verId, result: overall, deviations: deviations, checked: checked };
  });
}

function verifyRowToObj_(r) {
  return {
    verId: txt_(r[0]), when: txt_(r[1]), model: txt_(r[2]), lot: txt_(r[3]), rev: txt_(r[4]),
    bomStatus: txt_(r[5]), checked: txt_(r[6]), dev: txt_(r[7]), result: txt_(r[8]),
    by: txt_(r[9]), remarks: txt_(r[10])
  };
}

/** All verifications of a model, newest first. */
function getLotList(model) {
  const sh = sysSheet_(VERIFY_SHEET, VERIFY_HDR);
  if (sh.getLastRow() < 2) return [];
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, VERIFY_HDR.length).getValues().forEach(r => {
    if (norm_(r[2]) === norm_(model)) out.push(verifyRowToObj_(r));
  });
  return out.reverse();
}

function findVerification_(verId) {
  const sh = sysSheet_(VERIFY_SHEET, VERIFY_HDR);
  if (sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, VERIFY_HDR.length).getValues();
  for (let i = 0; i < v.length; i++) if (txt_(v[i][0]) === verId) return verifyRowToObj_(v[i]);
  return null;
}

/** Detail rows in BOM-interface order: Part_No, Description, Image, BOM_QPS, Lot_Part_No, Lot_QPS, Remark, Result. */
function getLotDetailRows_(verId) {
  const sh = migratedSheet_(LOT_SHEET, LOT_HDR);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, LOT_HDR.length).getValues()
    .filter(r => txt_(r[0]) === verId)
    .map(r => ({ part: txt_(r[4]), desc: txt_(r[5]), img: txt_(r[6]), bomQps: txt_(r[7]),
                 lotPart: txt_(r[8]), lotQps: txt_(r[9]), remark: txt_(r[10]), result: txt_(r[11]) }));
}

/** The BOM as verified for one lot (BOM + what was actually used). */
function getLotBom(verId) {
  return { info: findVerification_(verId), rows: getLotDetailRows_(verId) };
}

/** Writes a printable copy of one lot's BOM to LOT_BOM_VIEW, in the same column order as the log. */
function exportLotBom(verId) {
  const info = findVerification_(verId);
  if (!info) throw new Error('Verification not found: ' + verId);
  const rows = getLotDetailRows_(verId);
  const ss = dataSS_();
  let sh = ss.getSheetByName(LOTVIEW_SHEET);
  if (!sh) sh = ss.insertSheet(LOTVIEW_SHEET);
  sh.clear();

  const title = 'Model: ' + info.model + '   |   Lot: ' + info.lot + '   |   BOM Rev: ' + info.rev +
    ' (' + info.bomStatus + ')   |   Result: ' + info.result + ' (' + info.dev + ' deviation)' +
    '   |   ' + info.verId + ' on ' + info.when + ' by ' + info.by;
  sh.getRange(1, 1).setValue(title).setFontWeight('bold');
  if (info.remarks) sh.getRange(2, 1).setValue('Remarks: ' + info.remarks);

  const hdr = ['Part_No', 'Description', 'Image', 'BOM_QPS', 'Lot_Part_No', 'Lot_QPS', 'Remark', 'Result'];
  const IMG_COL = 3;
  sh.getRange(4, 1, 1, hdr.length).setValues([hdr])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  if (!rows.length) {
    sh.getRange(5, 1).setValue('No part rows were recorded for this verification ' +
      '(its BOM had 0 parts at the time, or it predates part-level logging). Re-verify this lot for a full record.');
  } else {
    // Part_No / Lot_Part_No as text so numeric-looking part numbers aren't mangled;
    // the Image column is left in the default format so its =IMAGE() formula renders.
    sh.getRange(5, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(5, 5, rows.length, 1).setNumberFormat('@');
    const body = rows.map(r => [r.part, r.desc, r.img ? '=IMAGE("' + r.img + '",1)' : '',
                                 r.bomQps, r.lotPart, r.lotQps, r.remark, r.result]);
    sh.getRange(5, 1, body.length, hdr.length).setValues(body);
    sh.getRange(5, 1, body.length, hdr.length).setBackgrounds(
      rows.map(r => new Array(hdr.length).fill(r.result === 'Matched' ? '#e6f4ea' : '#fad2cf')));
    sh.setRowHeights(5, body.length, ROW_HEIGHT_PX);
  }
  sh.setFrozenRows(4);
  sh.autoResizeColumns(1, hdr.length);
  sh.setColumnWidth(IMG_COL, 110);
  try { ss.setActiveSheet(sh); } catch (e) {}
  return { ok: true };
}

/** One CSV field, quoted if it needs it. */
function csvField_(v) {
  const s = txt_(v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/** CSV text of one lot's verification, same columns as exportLotBom, Result last. Client triggers the download. */
function getLotCsv(verId) {
  const info = findVerification_(verId);
  if (!info) throw new Error('Verification not found: ' + verId);
  const rows = getLotDetailRows_(verId);
  const hdr = ['Part_No', 'Description', 'Image_URL', 'BOM_QPS', 'Lot_Part_No', 'Lot_QPS', 'Remark', 'Result'];
  const lines = [
    ['Model', info.model].map(csvField_).join(','),
    ['Lot_No', info.lot].map(csvField_).join(','),
    ['BOM_Rev', info.rev, info.bomStatus].map(csvField_).join(','),
    ['Result', info.result, info.dev + ' deviation(s)'].map(csvField_).join(','),
    ['Verified', info.when, info.by].map(csvField_).join(','),
    '',
    hdr.map(csvField_).join(',')
  ];
  if (!rows.length) {
    lines.push(csvField_('No part rows were recorded for this verification ' +
      '(its BOM had 0 parts at the time, or it predates part-level logging). Re-verify this lot for a full record.'));
  }
  rows.forEach(r => lines.push([r.part, r.desc, r.img, r.bomQps, r.lotPart, r.lotQps, r.remark, r.result].map(csvField_).join(',')));
  return {
    filename: 'BOM_Verification_' + info.model + '_' + info.lot + '_' + info.verId + '.csv',
    csv: lines.join('\r\n')
  };
}

/** CSV text of a model's current BOM (Part_No, Description, Image_URL, QPS). Client triggers the download. */
function getModelBomCsv(model) {
  const b = getVerifyBom(model);
  const hdr = ['Part_No', 'Description', 'Image_URL', 'QPS'];
  const lines = [
    ['Model', model].map(csvField_).join(','),
    ['BOM_Rev', b.rev, b.locked ? 'Locked' : 'Not locked'].map(csvField_).join(','),
    ['Parts', b.parts.length].map(csvField_).join(','),
    '',
    hdr.map(csvField_).join(',')
  ];
  if (!b.parts.length) lines.push(csvField_('This model has no parts in its BOM yet.'));
  b.parts.forEach(p => lines.push([p.partNo, p.desc, p.img, p.qps].map(csvField_).join(',')));
  return { filename: 'BOM_' + model + '.csv', csv: lines.join('\r\n') };
}

// ======================================================================
//                     BOM COMPARISON (all models side by side)
// ======================================================================
const COMPARE_SHEET      = 'BOM_COMPARE';
const COMPARE_SHOW_IMAGE = true;   // true = add part thumbnail column (makes rows tall)

/** Rebuilds the BOM_COMPARE sheet from every model sheet. */
function buildComparison() {
  const ss     = dataSS_();
  const models = getModelSheets_();
  if (!models.length) throw new Error('No model sheets found.');
  const ctx = readMaster_();
  const mh  = ctx.mh;

  // ---- description / image from PART_MASTER (columns found by header name) ----
  const info = {};
  ctx.mVals.forEach((r, i) => {
    const k = norm_(r[COL_PART_NO - 1]);
    if (k && !(k in info)) {
      const ex = masterExtra_(ctx, i);
      info[k] = { desc: ex.desc, img: ex.img ? '=IMAGE("' + ex.img + '",1)' : '' };
    }
  });

  // ---- collect parts + QPS from each model ----
  const names = [];
  const parts = {};   // KEY -> { partNo, q: { model: qps } }
  models.forEach(s => {
    const name = s.getName();
    names.push(name);
    const n = s.getLastRow() - HEADER_ROW;
    if (n <= 0) return;
    const pc = partColIn_(s, mh);
    const qc = headersOf_(s).findIndex(h => norm_(h) === QPS_HEADER) + 1;
    const pv = s.getRange(HEADER_ROW + 1, pc, n, 1).getValues();
    const qv = qc ? s.getRange(HEADER_ROW + 1, qc, n, 1).getValues() : null;
    pv.forEach((r, i) => {
      const raw = String(r[0]).trim();
      const k = norm_(raw);
      if (!k) return;
      if (!parts[k]) parts[k] = { partNo: raw, q: {} };
      const q = qv ? qv[i][0] : '';
      parts[k].q[name] = (q === '' ? '\u2713' : q);
    });
  });

  const total = names.length;
  const rows = Object.keys(parts).map(k => {
    const p = parts[k];
    const used = names.filter(n => n in p.q);
    const qs = used.map(n => String(p.q[n]));
    const distinct = qs.filter((x, i) => qs.indexOf(x) === i).length;
    let status;
    if (used.length === total)      status = 'Common (all models)';
    else if (used.length === 1)     status = 'Only in ' + used[0];
    else                            status = 'Shared (' + used.length + ' models)';
    return {
      k: k, p: p, used: used.length, status: status,
      check: used.length < 2 ? '' : (distinct === 1 ? 'Same' : 'Different')
    };
  });
  rows.sort((a, b) => b.used - a.used || a.p.partNo.localeCompare(b.p.partNo));

  // ---- build the table ----
  const header = ['Part_No']
    .concat(COMPARE_SHOW_IMAGE ? ['Image'] : [], ['Description'], names, ['Used in', 'Status', 'QPS check']);
  const ncols = header.length;
  const data = rows.map(r => {
    const inf = info[r.k] || { desc: '', img: '' };
    return [r.p.partNo]
      .concat(COMPARE_SHOW_IMAGE ? [inf.img] : [], [inf.desc],
              names.map(n => (n in r.p.q ? r.p.q[n] : '')),
              [r.used + ' / ' + total, r.status, r.check]);
  });

  const nCommon = rows.filter(r => r.used === total).length;
  const nUnique = rows.filter(r => r.used === 1 && total > 1).length;
  const nShared = rows.length - nCommon - nUnique;
  const nDiff   = rows.filter(r => r.check === 'Different').length;
  const summary = 'Models: ' + total + '  |  Total unique parts: ' + rows.length +
    '  |  Common to all: ' + nCommon + '  |  Shared by some: ' + nShared +
    '  |  Model-specific: ' + nUnique + '  |  Common parts with different QPS: ' + nDiff +
    '  |  Generated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');

  // ---- write ----
  let sheet = ss.getSheetByName(COMPARE_SHEET);
  if (!sheet) sheet = ss.insertSheet(COMPARE_SHEET);
  const oldFilter = sheet.getFilter();
  if (oldFilter) oldFilter.remove();
  sheet.clear();
  sheet.setRowHeights(1, sheet.getMaxRows(), 21);

  sheet.getRange(1, 1).setValue(summary).setFontWeight('bold');
  const HDR = 3;
  sheet.getRange(HDR, 1, 1, ncols).setValues([header])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');

  if (data.length) {
    const body = sheet.getRange(HDR + 1, 1, data.length, ncols);
    body.setValues(data);

    const bg = rows.map(r => {
      const color = r.used === total ? '#e6f4ea' : (r.used === 1 && total > 1 ? '#ffffff' : '#fef7e0');
      const line = new Array(ncols).fill(color);
      if (r.check === 'Different') line[ncols - 1] = '#fad2cf';
      return line;
    });
    body.setBackgrounds(bg);

    // centre the model / count / check columns
    const firstModelCol = header.indexOf(names[0]) + 1;
    sheet.getRange(HDR + 1, firstModelCol, data.length, total).setHorizontalAlignment('center');
    sheet.getRange(HDR + 1, ncols - 2, data.length, 1).setHorizontalAlignment('center');
    sheet.getRange(HDR + 1, ncols,     data.length, 1).setHorizontalAlignment('center');

    if (COMPARE_SHOW_IMAGE) sheet.setRowHeights(HDR + 1, data.length, ROW_HEIGHT_PX);
    sheet.getRange(HDR, 1, data.length + 1, ncols).createFilter();
  }

  sheet.setFrozenRows(HDR);
  sheet.setFrozenColumns(1);
  sheet.autoResizeColumns(1, ncols);
  if (COMPARE_SHOW_IMAGE) sheet.setColumnWidth(2, 110);
  sheet.setColumnWidth(1, Math.max(sheet.getColumnWidth(1), 140));

  try {
    ss.setActiveSheet(sheet);
    ss.toast(nCommon + ' parts are common to all ' + total + ' models.', 'BOM comparison', 8);
  } catch (e) { /* not available in the web app */ }
}

// ======================================================================
//                     COMPARE TWO MODELS (interface)
// ======================================================================
const TWO_SHEET = 'MODEL_COMPARE';

function showTwoCompareDialog() {
  const html = HtmlService.createHtmlOutput(getTwoCompareHtml_()).setWidth(1600).setHeight(980);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Compare two models');
}

/** Common / only-in-A / only-in-B parts of two model sheets. */
function compareTwoModels(a, b) {
  if (!a || !b) throw new Error('Choose both models.');
  if (norm_(a) === norm_(b)) throw new Error('Choose two different models.');
  const A = getVerifyBom(a), B = getVerifyBom(b);

  const mapB = {};
  B.parts.forEach(p => { mapB[norm_(p.partNo)] = p; });
  const seenA = {};
  const common = [], onlyA = [], onlyB = [];

  A.parts.forEach(p => {
    const k = norm_(p.partNo);
    seenA[k] = true;
    const o = mapB[k];
    if (o) common.push({ part: p.partNo, desc: p.desc, img: p.img, qa: p.qps, qb: o.qps,
                         same: Number(p.qps) === Number(o.qps) });
    else   onlyA.push({ part: p.partNo, desc: p.desc, img: p.img, q: p.qps });
  });
  B.parts.forEach(p => {
    if (!seenA[norm_(p.partNo)]) onlyB.push({ part: p.partNo, desc: p.desc, img: p.img, q: p.qps });
  });

  return {
    a: { name: a, locked: A.locked, rev: A.rev, count: A.parts.length },
    b: { name: b, locked: B.locked, rev: B.rev, count: B.parts.length },
    common: common, onlyA: onlyA, onlyB: onlyB
  };
}

/** Writes the comparison to the MODEL_COMPARE sheet. */
function exportTwoCompare(a, b) {
  const r  = compareTwoModels(a, b);
  const ss = dataSS_();
  let sh = ss.getSheetByName(TWO_SHEET);
  if (!sh) sh = ss.insertSheet(TWO_SHEET);
  sh.clear();

  const diff = r.common.filter(x => !x.same).length;
  sh.getRange(1, 1).setValue('Model comparison: ' + a + ' vs ' + b +
    '   |   Common: ' + r.common.length + ' (' + diff + ' with different QPS)' +
    '   |   Only in ' + a + ': ' + r.onlyA.length +
    '   |   Only in ' + b + ': ' + r.onlyB.length).setFontWeight('bold');

  const N = 6, IMG_COL = 2, out = [], bg = [], titleRows = [], hdrRows = [], tallRows = [];
  const img_ = u => u ? '=IMAGE("' + u + '",1)' : '';
  const push = (arr, color, tall) => {
    const row = arr.slice();
    while (row.length < N) row.push('');
    out.push(row);
    bg.push(new Array(N).fill(color));
    if (tall) tallRows.push(out.length);
    return out.length;
  };

  titleRows.push(push(['COMMON PARTS (' + r.common.length + ')'], '#ffffff'));
  hdrRows.push(push(['Part_No', 'Image', 'Description', 'QPS in ' + a, 'QPS in ' + b, 'QPS check'], '#1a73e8'));
  r.common.forEach(x => push([x.part, img_(x.img), x.desc, x.qa, x.qb, x.same ? 'Same' : 'Different'],
                              x.same ? '#e6f4ea' : '#fad2cf', true));
  push([], '#ffffff');

  titleRows.push(push(['ONLY IN ' + a + ' (' + r.onlyA.length + ')'], '#ffffff'));
  hdrRows.push(push(['Part_No', 'Image', 'Description', 'QPS', '', ''], '#1a73e8'));
  r.onlyA.forEach(x => push([x.part, img_(x.img), x.desc, x.q], '#fef7e0', true));
  push([], '#ffffff');

  titleRows.push(push(['ONLY IN ' + b + ' (' + r.onlyB.length + ')'], '#ffffff'));
  hdrRows.push(push(['Part_No', 'Image', 'Description', 'QPS', '', ''], '#1a73e8'));
  r.onlyB.forEach(x => push([x.part, img_(x.img), x.desc, x.q], '#e8f0fe', true));

  // Part_No column as text; Image column left alone so its formula renders.
  sh.getRange(3, 1, out.length, 1).setNumberFormat('@');
  const rng = sh.getRange(3, 1, out.length, N);
  rng.setValues(out);
  rng.setBackgrounds(bg);
  titleRows.forEach(i => sh.getRange(2 + i, 1).setFontWeight('bold'));
  hdrRows.forEach(i => sh.getRange(2 + i, 1, 1, N).setFontWeight('bold').setFontColor('#ffffff'));
  tallRows.forEach(i => sh.setRowHeight(2 + i, ROW_HEIGHT_PX));

  sh.autoResizeColumns(1, N);
  sh.setColumnWidth(IMG_COL, 110);
  try { ss.setActiveSheet(sh); } catch (e) {}
  return { ok: true };
}

/** Screen HTML (embedded). */
function getSidebarHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 90px; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; }

input[type=text], input[type=number], select, textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
  transition: border-color .15s;
}
input[type=text]:focus, input[type=number]:focus, select:focus {
  outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg);
}
input[type=number] { width: 78px; }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }

.tabs { display: flex; gap: 4px; }
.tab {
  padding: 9px 16px; cursor: pointer; border-radius: var(--radius-sm) var(--radius-sm) 0 0;
  background: transparent; color: var(--sub); font-weight: 500; border-bottom: 2px solid transparent;
}
.tab:hover { background: #f1f3f4; }
.tab.on { color: var(--blue); border-bottom: 2px solid var(--blue); background: var(--blue-bg); }

.footbar {
  position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 14px; align-items: center;
  padding: 12px 22px; background: var(--surface); border-top: 1px solid var(--line-soft); box-shadow: 0 -2px 8px rgba(0,0,0,.05);
}
.footbar .sp { flex: 1; }
.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }

img.thumb { width: 40px; height: 40px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

    body { padding: 14px; }
    h3 { margin: 0 0 10px; font-size: 15px; font-weight: 700; }
    .part { background: var(--blue-bg); color: var(--blue-dark); border-radius: var(--radius); padding: 10px 12px; margin-bottom: 14px; font-weight: 700; word-break: break-all; }
    .msg { color: var(--sub); padding: 12px 0; }
    .row { display: flex; align-items: center; gap: 8px; padding: 8px 0; border-bottom: 1px solid var(--line-soft); }
    .row label { flex: 1; cursor: pointer; }
    .row input[type=number] { width: 64px; padding: 6px; }
    .bulk { display: flex; gap: 6px; align-items: center; margin: 6px 0 14px; background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius-sm); padding: 8px; }
    .bulk input { width: 64px; }
    .ecn input { width: 100%; box-sizing: border-box; margin-top: 6px; }
    button.primary { width: 100%; margin-top: 14px; padding: 10px; font-size: 14px; }
    #status { margin-top: 10px; white-space: pre-line; }
  </style>
</head>
<body>
  <h3>Assign part to models</h3>
  <div id="content"><div class="msg">Loading...</div></div>
  <div id="status"></div>

<script>
  var current = null, busy = false, armed = false;

  function poll() {
    if (busy) return;
    google.script.run
      .withSuccessHandler(onSelection)
      .withFailureHandler(function (e) {
        document.getElementById('content').innerHTML = '<div class="msg err">Error: ' + esc(e.message) + '</div>';
      })
      .getSelectedPart();
  }

  function onSelection(res) {
    var sig = res.ok ? res.partNo : 'none:' + res.msg;
    if (current && current.sig === sig) return;
    current = { sig: sig, data: res };
    armed = false;
    render(res);
  }

  function render(res) {
    var el = document.getElementById('content');
    if (!res.ok) { el.innerHTML = '<div class="msg">' + esc(res.msg) + '</div>'; return; }

    var anyLocked = false;
    var h = '<div class="part">' + esc(res.partNo) + '</div>';
    h += '<div class="bulk">All models QPS: <input type="number" min="0" step="any" id="bulkQps">' +
         '<button onclick="applyAll()">Apply to all</button></div>';

    res.models.forEach(function (m, i) {
      if (m.locked) anyLocked = true;
      h += '<div class="row">' +
        '<input type="checkbox" id="c' + i + '"' + (m.listed ? ' checked' : '') + ' onchange="armed=false">' +
        '<label for="c' + i + '">' + esc(m.name) +
          (m.locked ? ' <span class="pill warn">LOCKED</span>' : '') +
          (m.listed ? '<br><span class="pill ok">already listed</span>' : '') + '</label>' +
        '<input type="number" min="0" step="any" id="q' + i + '" placeholder="QPS" value="' + esc(m.qps) + '" oninput="armed=false">' +
        '</div>';
    });

    if (anyLocked) {
      h += '<div class="banner warn ecn"><b>Locked BOM — ECN required</b><br>' +
           'Any add / QPS change / removal on a LOCKED model needs an ECN.' +
           '<input type="text" id="ecnNo" placeholder="ECN No">' +
           '<input type="text" id="ecnDesc" placeholder="Change description / reason"></div>';
    }
    h += '<button class="primary" id="go" onclick="submit()">Save to model sheets</button>';
    el.innerHTML = h;
  }

  function applyAll() {
    var v = document.getElementById('bulkQps').value;
    if (v === '') return;
    current.data.models.forEach(function (m, i) {
      document.getElementById('c' + i).checked = true;
      document.getElementById('q' + i).value = v;
    });
    armed = false;
  }

  function submit() {
    var d = current.data, add = [], remove = [];
    for (var i = 0; i < d.models.length; i++) {
      var checked = document.getElementById('c' + i).checked;
      var qps = document.getElementById('q' + i).value;
      if (checked) {
        if (!(Number(qps) > 0)) return setStatus('Enter a QPS for ' + d.models[i].name, true);
        add.push({ model: d.models[i].name, qps: Number(qps) });
      } else if (d.models[i].listed) {
        remove.push(d.models[i].name);
      }
    }
    if (!add.length && !remove.length) return setStatus('Nothing to save.', true);
    if (remove.length && !armed) {
      armed = true;
      return setStatus('Part will be REMOVED from: ' + remove.join(', ') + '\\nPress Save again to confirm.', true);
    }
    armed = false;

    var noEl = document.getElementById('ecnNo'), dEl = document.getElementById('ecnDesc');
    var ecn = { no: noEl ? noEl.value : '', desc: dEl ? dEl.value : '' };

    busy = true;
    document.getElementById('go').disabled = true;
    setStatus('Saving...');
    google.script.run
      .withSuccessHandler(function (r) {
        busy = false; current = null;
        setStatus(r.done.join('\\n'), false);
        poll();
      })
      .withFailureHandler(function (e) {
        busy = false;
        document.getElementById('go').disabled = false;
        setStatus(e.message, true);
      })
      .assignPart({ partNo: d.partNo, add: add, remove: remove, ecn: ecn });
  }

  function setStatus(t, isErr) {
    var s = document.getElementById('status');
    s.textContent = t;
    s.className = isErr === true ? 'err' : (isErr === false ? 'ok' : '');
  }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }

  poll();
  setInterval(poll, 1500);
</script>
</body>
</html>
`;
}

/** Screen HTML (embedded). */
function getBomHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 90px; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; }

input[type=text], input[type=number], select, textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
  transition: border-color .15s;
}
input[type=text]:focus, input[type=number]:focus, select:focus {
  outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg);
}
input[type=number] { width: 78px; }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }

.tabs { display: flex; gap: 4px; }
.tab {
  padding: 9px 16px; cursor: pointer; border-radius: var(--radius-sm) var(--radius-sm) 0 0;
  background: transparent; color: var(--sub); font-weight: 500; border-bottom: 2px solid transparent;
}
.tab:hover { background: #f1f3f4; }
.tab.on { color: var(--blue); border-bottom: 2px solid var(--blue); background: var(--blue-bg); }

.footbar {
  position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 14px; align-items: center;
  padding: 12px 22px; background: var(--surface); border-top: 1px solid var(--line-soft); box-shadow: 0 -2px 8px rgba(0,0,0,.05);
}
.footbar .sp { flex: 1; }
.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }

img.thumb { width: 40px; height: 40px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

  .tag { font-size: 11px; color: var(--green); margin-left: 6px; font-weight: 600; }
  td input[type=number] { width: 100%; }
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>Build model BOM</h1>
    <span class="sub">Bulk add or update parts for one model</span>
    <div class="fill"></div>
  </div>

  <div class="appbody">
    <div class="toolbar">
      <label>Model sheet <select id="model" onchange="loadModel()"></select></label>
      <label>or new model <input type="text" id="newModel" placeholder="e.g. LFG3-XXXX" onchange="onNew()"></label>
      <label>Default QPS <input type="number" id="defQps" min="0" step="any" value="1"></label>
    </div>

    <div class="toolbar">
      <input type="file" id="csvFile" accept=".csv,text/csv" style="display:none" onchange="onCsvFile(this)">
      <button onclick="document.getElementById('csvFile').click()">Upload CSV to modify BOM</button>
      <span class="sub">Columns: Part_No, QPS (header row optional - a column named "part"/"code" and one named "qps"/"qty" are detected automatically)</span>
      <label style="margin-left:auto"><input type="checkbox" id="csvReplace"> Replace entire BOM (remove parts not in the file)</label>
    </div>

    <div class="banner warn" id="banner" style="display:none">
      <b id="lockTxt">BOM LOCKED</b> — every add / QPS change / removal needs an ECN.
      <div style="margin-top:8px;display:flex;gap:8px;flex-wrap:wrap">
        <input type="text" id="ecnNo" placeholder="ECN No" style="width:160px" oninput="armed=false">
        <input type="text" id="ecnDesc" placeholder="Change description / reason" style="width:460px" oninput="armed=false">
      </div>
    </div>

    <div class="toolbar">
      <input type="text" id="search" placeholder="Search part no / description" oninput="renderRows()" class="grow">
      <button onclick="checkVisible(true)">Check visible</button>
      <button onclick="checkVisible(false)">Uncheck visible</button>
      <label><input type="checkbox" id="onlySel" onchange="renderRows()"> Selected only</label>
    </div>

    <div class="card">
      <div class="tbl-wrap" style="max-height:calc(100vh - 320px)">
        <table>
          <thead><tr><th style="width:36px"></th><th style="width:60px">Img</th><th>Part No</th><th>Description</th><th style="width:100px">QPS</th></tr></thead>
          <tbody id="rows"><tr><td colspan="5">Loading parts...</td></tr></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="footbar">
    <span id="count" class="pill gray"></span>
    <span id="status" class="sp"></span>
    <button id="dl" onclick="downloadBom()">Download BOM</button>
    <button class="primary" id="save" onclick="save()">Save BOM</button>
  </div>

<script>
  var parts = [], models = [], wasNew = false, armed = false, modelLocked = false;

  function downloadBom() {
    var nn = newName(), m = document.getElementById('model').value;
    if (nn) return setStatus('Save this new model first, then download its BOM.', true);
    if (!m) return setStatus('Choose a model.', true);
    setStatus('Preparing ' + m + ' BOM...');
    google.script.run.withSuccessHandler(function (r) {
      var blob = new Blob([r.csv], { type: 'text/csv;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = r.filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
      setStatus('Downloaded ' + r.filename, false);
    }).withFailureHandler(fail).getModelBomCsv(m);
  }

  function parseCsvLine(line) {
    var result = [], cur = '', inQ = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (inQ) {
        if (ch === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else { inQ = false; } }
        else cur += ch;
      } else {
        if (ch === '"') inQ = true;
        else if (ch === ',') { result.push(cur); cur = ''; }
        else cur += ch;
      }
    }
    result.push(cur);
    return result;
  }

  /**
   * Scans every row (not just the first) for the real header row, so files
   * with metadata lines above the table - like a downloaded BOM verification
   * or lot-BOM CSV ("Model,...", "Lot_No,...", a blank line, THEN the real
   * "Part_No,Description,...,BOM_QPS,..." header) still get parsed correctly
   * instead of mistaking an early metadata row or a Description column for
   * the quantity column.
   */
  function findCsvHeader(rows) {
    for (var i = 0; i < rows.length; i++) {
      var header = rows[i].map(function (h) { return h.trim().toLowerCase(); });
      var partIdx = -1, qtyCandidates = [];
      header.forEach(function (h, j) {
        if (partIdx < 0 && /part|code|p\\/n|^pn$/.test(h)) partIdx = j;
        if (/qps|qty|quantity|qnty|count/.test(h)) qtyCandidates.push(j);
      });
      if (partIdx < 0 || !qtyCandidates.length) continue;
      // Several QPS-like columns can appear together (e.g. a lot-verification
      // export has both BOM_QPS and Lot_QPS) - prefer an exact "qps"/"qty"
      // match, then one named after the BOM/master quantity, then anything
      // that isn't a lot-specific column, else just take the first.
      var qtyIdx = qtyCandidates.filter(function (j) { return header[j] === 'qps' || header[j] === 'qty'; })[0];
      if (qtyIdx === undefined) qtyIdx = qtyCandidates.filter(function (j) { return header[j].indexOf('bom') >= 0; })[0];
      if (qtyIdx === undefined) qtyIdx = qtyCandidates.filter(function (j) { return header[j].indexOf('lot') < 0; })[0];
      if (qtyIdx === undefined) qtyIdx = qtyCandidates[0];
      return { partIdx: partIdx, qtyIdx: qtyIdx, dataRows: rows.slice(i + 1) };
    }
    return null;
  }

  function parseCsv(text) {
    var lines = text.split(/\\r\\n|\\r|\\n/).filter(function (l) { return l.trim() !== ''; });
    if (!lines.length) return [];
    var rows = lines.map(parseCsvLine);
    var found = findCsvHeader(rows);
    var partIdx, qtyIdx, dataRows;
    if (found) {
      partIdx = found.partIdx; qtyIdx = found.qtyIdx; dataRows = found.dataRows;
    } else {
      // No recognizable header anywhere - assume a plain, header-less file:
      // column 1 is the part number, column 2 (if present) is the QPS.
      partIdx = 0; qtyIdx = rows[0].length > 1 ? 1 : -1;
      dataRows = rows;
    }
    var out = [];
    dataRows.forEach(function (r) {
      var pn = (r[partIdx] || '').trim();
      if (!pn) return;
      var q = qtyIdx >= 0 ? parseFloat(r[qtyIdx]) : NaN;
      out.push({ partNo: pn, qps: q });
    });
    return out;
  }

  function onCsvFile(input) {
    var f = input.files[0];
    if (!f) return;
    if (!newName() && !document.getElementById('model').value) {
      setStatus('Choose or create a model first.', true); input.value = ''; return;
    }
    var reader = new FileReader();
    reader.onload = function (e) {
      var rows = parseCsv(String(e.target.result));
      input.value = '';
      if (!rows.length) { setStatus('No rows found in that CSV.', true); return; }
      applyCsvRows(rows, document.getElementById('csvReplace').checked);
    };
    reader.onerror = function () { setStatus('Could not read that file.', true); input.value = ''; };
    reader.readAsText(f);
  }

  function applyCsvRows(rows, replaceMode) {
    var byKey = {};
    parts.forEach(function (p, i) { byKey[p.key] = i; });
    var seenKeys = {}, matched = 0, unmatched = [];
    var defQ = Number(document.getElementById('defQps').value) || 1;
    rows.forEach(function (r) {
      var key = r.partNo.toUpperCase();
      seenKeys[key] = true;
      var idx = byKey[key];
      if (idx === undefined) { unmatched.push(r.partNo); return; }
      matched++;
      var q = (r.qps && r.qps > 0) ? r.qps : defQ;
      parts[idx].checked = true;
      parts[idx].qps = q;
    });
    if (replaceMode) {
      parts.forEach(function (p) {
        if (p.was && !seenKeys[p.key]) p.checked = false;
      });
    }
    armed = false;
    renderRows();
    var msg = matched + ' part(s) matched from the CSV' + (replaceMode ? ' - BOM will be replaced to match the file' : ' - added/updated, existing parts not in the file are kept');
    if (unmatched.length) msg += '. Not found in PART_MASTER (' + unmatched.length + '): ' + unmatched.slice(0, 15).join(', ') + (unmatched.length > 15 ? ', ...' : '');
    setStatus(msg, unmatched.length ? true : false);
  }

  function init() {
    google.script.run
      .withSuccessHandler(function (d) {
        models = d.models;
        parts = d.parts.map(function (p) {
          return { partNo: p.partNo, desc: p.desc, img: p.img, key: p.partNo.toUpperCase(), checked: false, qps: '', was: false };
        });
        fillModels();
        renderRows();
        if (models.length) loadModel();
      })
      .withFailureHandler(fail)
      .getBomData();
  }

  function fillModels(sel) {
    var el = document.getElementById('model');
    el.innerHTML = models.map(function (m) { return '<option>' + esc(m) + '</option>'; }).join('');
    if (sel) el.value = sel;
  }

  function newName() { return document.getElementById('newModel').value.trim(); }

  function showLock(locked, rev) {
    modelLocked = locked;
    document.getElementById('banner').style.display = locked ? 'block' : 'none';
    if (locked) document.getElementById('lockTxt').textContent = 'BOM LOCKED (Rev ' + rev + ')';
  }

  function onNew() {
    var nn = newName() !== '';
    document.getElementById('model').disabled = nn;
    if (nn === wasNew) return;
    wasNew = nn;
    if (nn) { showLock(false, 0); applyExisting({ parts: {}, locked: false, rev: 0 }); } else loadModel();
  }

  function loadModel() {
    var m = document.getElementById('model').value;
    if (!m) return;
    armed = false;
    setStatus('Loading ' + m + '...');
    google.script.run
      .withSuccessHandler(function (r) { applyExisting(r); setStatus(''); })
      .withFailureHandler(fail)
      .getModelParts(m);
  }

  function applyExisting(r) {
    showLock(r.locked, r.rev);
    parts.forEach(function (p) {
      var has = Object.prototype.hasOwnProperty.call(r.parts, p.key);
      p.was = has; p.checked = has; p.qps = has ? r.parts[p.key] : '';
    });
    renderRows();
  }

  function visibleIdx() {
    var q = document.getElementById('search').value.trim().toUpperCase();
    var only = document.getElementById('onlySel').checked;
    var out = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (q && p.key.indexOf(q) < 0 && p.desc.toUpperCase().indexOf(q) < 0) continue;
      if (only && !p.checked) continue;
      out.push(i);
    }
    return out;
  }

  function renderRows() {
    var idx = visibleIdx(), html = [];
    idx.forEach(function (i) {
      var p = parts[i];
      html.push('<tr><td><input type="checkbox" ' + (p.checked ? 'checked ' : '') + 'onchange="tog(' + i + ',this.checked)"></td>' +
        '<td>' + (p.img ? '<img class="thumb" src="' + esc(p.img) + '">' : '') + '</td>' +
        '<td>' + esc(p.partNo) + (p.was ? '<span class="tag">in BOM</span>' : '') + '</td>' +
        '<td>' + esc(p.desc) + '</td>' +
        '<td><input type="number" min="0" step="any" id="q' + i + '" value="' + esc(p.qps) + '" oninput="setQ(' + i + ',this.value)"></td></tr>');
    });
    document.getElementById('rows').innerHTML = html.length ? html.join('') : '<tr><td colspan="5">No parts match.</td></tr>';
    updateCount(idx.length);
  }

  function fillDefault(p, i) {
    if (p.qps === '' || p.qps === null) {
      p.qps = document.getElementById('defQps').value;
      var f = document.getElementById('q' + i);
      if (f) f.value = p.qps;
    }
  }

  function tog(i, c) {
    armed = false;
    parts[i].checked = c;
    if (c) fillDefault(parts[i], i);
    updateCount();
  }

  function setQ(i, v) { armed = false; parts[i].qps = v; }

  function checkVisible(flag) {
    armed = false;
    visibleIdx().forEach(function (i) {
      parts[i].checked = flag;
      if (flag) fillDefault(parts[i], i);
    });
    renderRows();
  }

  function updateCount(shown) {
    var sel = 0, rem = 0;
    parts.forEach(function (p) { if (p.checked) sel++; else if (p.was) rem++; });
    var t = sel + ' selected';
    if (rem) t += ' - ' + rem + ' will be removed';
    if (shown !== undefined) t += ' - ' + shown + ' shown';
    document.getElementById('count').textContent = t;
  }

  function save() {
    var items = [], remove = [];
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (p.checked) {
        var q = Number(p.qps);
        if (!(q > 0)) return setStatus('Enter a QPS for ' + p.partNo, true);
        items.push({ partNo: p.partNo, qps: q });
      } else if (p.was) {
        remove.push(p.partNo);
      }
    }
    var nn = newName(), m = document.getElementById('model').value;
    if (!nn && !m) return setStatus('Choose a model sheet.', true);
    if (!items.length && !remove.length) return setStatus('Nothing to save.', true);
    if (remove.length && !armed) {
      armed = true;
      return setStatus(remove.length + ' part(s) will be REMOVED. Press Save BOM again to confirm.', true);
    }
    armed = false;

    var ecn = { no: document.getElementById('ecnNo').value, desc: document.getElementById('ecnDesc').value };
    document.getElementById('save').disabled = true;
    setStatus('Saving...');
    google.script.run
      .withSuccessHandler(function (r) {
        document.getElementById('save').disabled = false;
        var msg = r.model + ': added ' + r.added + ', updated ' + r.updated + ', removed ' + r.removed;
        if (r.missing.length) msg += ' (not found in master: ' + r.missing.join(', ') + ')';
        if (nn) {
          models = r.models; fillModels(r.model);
          document.getElementById('newModel').value = '';
          document.getElementById('model').disabled = false;
          wasNew = false;
        }
        setStatus(msg, false);
        document.getElementById('ecnNo').value = '';
        document.getElementById('ecnDesc').value = '';
        google.script.run.withSuccessHandler(applyExisting).withFailureHandler(fail).getModelParts(r.model);
      })
      .withFailureHandler(function (e) {
        document.getElementById('save').disabled = false;
        fail(e);
      })
      .saveBom({ model: m, newModel: nn, items: items, remove: remove, ecn: ecn });
  }

  function fail(e) { setStatus('Error: ' + (e && e.message ? e.message : e), true); }
  function setStatus(t, isErr) {
    var s = document.getElementById('status');
    s.textContent = t;
    s.className = 'sp ' + (isErr === true ? 'err' : (isErr === false ? 'ok' : ''));
  }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }

  init();
</script>
</body>
</html>
`;
}

/** Screen HTML (embedded). */
function getControlHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 90px; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; }

input[type=text], input[type=number], select, textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
  transition: border-color .15s;
}
input[type=text]:focus, input[type=number]:focus, select:focus {
  outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg);
}
input[type=number] { width: 78px; }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }

.tabs { display: flex; gap: 4px; }
.tab {
  padding: 9px 16px; cursor: pointer; border-radius: var(--radius-sm) var(--radius-sm) 0 0;
  background: transparent; color: var(--sub); font-weight: 500; border-bottom: 2px solid transparent;
}
.tab:hover { background: #f1f3f4; }
.tab.on { color: var(--blue); border-bottom: 2px solid var(--blue); background: var(--blue-bg); }

.footbar {
  position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 14px; align-items: center;
  padding: 12px 22px; background: var(--surface); border-top: 1px solid var(--line-soft); box-shadow: 0 -2px 8px rgba(0,0,0,.05);
}
.footbar .sp { flex: 1; }
.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }

img.thumb { width: 40px; height: 40px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

  button.lockb { background: var(--blue); color: #fff; border-color: var(--blue); }
  button.lockb:hover { background: var(--blue-dark); }
  button.unlockb { background: #fff; color: var(--red); border-color: var(--red); }
  button.unlockb:hover { background: var(--red-bg); }
  .reason input { width: 320px; }
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>BOM lock / unlock</h1>
    <span class="sub">Locked BOMs can only change via an ECN — every lock, unlock and change is logged</span>
    <div class="fill"></div>
  </div>

  <div class="appbody">
    <div class="card">
      <div class="tbl-wrap">
        <table>
          <thead><tr><th>Model</th><th>Parts</th><th>Status</th><th>Rev</th><th>Locked by / on</th><th>Last ECN</th><th></th></tr></thead>
          <tbody id="rows"><tr><td colspan="7">Loading...</td></tr></tbody>
        </table>
      </div>
    </div>

    <div class="card reason" id="reasonBox" style="display:none">
      <div class="card-h">Unlock <b id="unlockName"></b></div>
      <div style="padding:14px;display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <span>Reason:</span>
        <input type="text" id="reason" placeholder="Why is the BOM being unlocked?">
        <button class="unlockb" onclick="doUnlock()">Confirm unlock</button>
        <button onclick="cancelUnlock()">Cancel</button>
      </div>
    </div>

    <div class="banner info">Locked BOMs can only be changed from the Part DB tools with an ECN number and description. Every change, lock and unlock is written to <b>ECN_HISTORY</b>.</div>
    <div id="status"></div>
  </div>

<script>
  var data = [], unlocking = '';

  function load() {
    google.script.run.withSuccessHandler(function (d) { data = d; render(); })
      .withFailureHandler(fail).getControlData();
  }

  function render() {
    if (!data.length) { document.getElementById('rows').innerHTML = '<tr><td colspan="7">No model sheets found.</td></tr>'; return; }
    document.getElementById('rows').innerHTML = data.map(function (r, i) {
      var isL = r.status === 'Locked';
      return '<tr><td>' + esc(r.model) + '</td><td>' + esc(r.parts) + '</td>' +
        '<td><span class="pill ' + (isL ? 'warn' : 'ok') + '">' + esc(r.status) + '</span></td>' +
        '<td>' + (isL || r.rev ? esc(r.rev) : '-') + '</td>' +
        '<td>' + (isL ? esc(r.by) + '<br>' + esc(r.on) : '') + '</td>' +
        '<td>' + esc(r.ecn) + '</td>' +
        '<td style="white-space:nowrap">' +
          (isL ? '<button class="unlockb" onclick="askUnlock(' + i + ')">Unlock</button>'
               : '<button class="lockb" onclick="doLock(' + i + ')">Lock BOM</button>') +
          ' <button onclick="downloadBom(' + i + ')">Download BOM</button>' +
        '</td></tr>';
    }).join('');
  }

  function downloadBom(i) {
    var model = data[i].model;
    setStatus('Preparing ' + model + ' BOM...');
    google.script.run.withSuccessHandler(function (r) {
      var blob = new Blob([r.csv], { type: 'text/csv;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = r.filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
      setStatus('Downloaded ' + r.filename, false);
    }).withFailureHandler(fail).getModelBomCsv(model);
  }

  function doLock(i) {
    setStatus('Locking ' + data[i].model + '...');
    google.script.run.withSuccessHandler(function () { setStatus(data[i].model + ' locked.', false); load(); })
      .withFailureHandler(fail).lockBom(data[i].model);
  }

  function askUnlock(i) {
    unlocking = data[i].model;
    document.getElementById('unlockName').textContent = unlocking;
    document.getElementById('reason').value = '';
    document.getElementById('reasonBox').style.display = 'block';
  }

  function cancelUnlock() { document.getElementById('reasonBox').style.display = 'none'; unlocking = ''; }

  function doUnlock() {
    var reason = document.getElementById('reason').value;
    setStatus('Unlocking ' + unlocking + '...');
    var name = unlocking;
    google.script.run.withSuccessHandler(function () {
      cancelUnlock(); setStatus(name + ' unlocked.', false); load();
    }).withFailureHandler(fail).unlockBom(name, reason);
  }

  function fail(e) { setStatus('Error: ' + (e && e.message ? e.message : e), true); }
  function setStatus(t, isErr) {
    var s = document.getElementById('status');
    s.textContent = t;
    s.className = isErr === true ? 'err' : (isErr === false ? 'ok' : '');
  }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  load();
</script>
</body>
</html>
`;
}

/** Screen HTML (embedded). */
function getVerifyHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 90px; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; }

input[type=text], input[type=number], select, textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
  transition: border-color .15s;
}
input[type=text]:focus, input[type=number]:focus, select:focus {
  outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg);
}
input[type=number] { width: 78px; }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }

.tabs { display: flex; gap: 4px; }
.tab {
  padding: 9px 16px; cursor: pointer; border-radius: var(--radius-sm) var(--radius-sm) 0 0;
  background: transparent; color: var(--sub); font-weight: 500; border-bottom: 2px solid transparent;
}
.tab:hover { background: #f1f3f4; }
.tab.on { color: var(--blue); border-bottom: 2px solid var(--blue); background: var(--blue-bg); }

.footbar {
  position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 14px; align-items: center;
  padding: 12px 22px; background: var(--surface); border-top: 1px solid var(--line-soft); box-shadow: 0 -2px 8px rgba(0,0,0,.05);
}
.footbar .sp { flex: 1; }
.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }

img.thumb { width: 40px; height: 40px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

  td input[type=text], td input[type=number] { width: 100%; box-sizing: border-box; }
  td input:disabled { background: #f1f3f4; color: #80868b; }
  tr.dev td { background: var(--red-bg); }
  tr.extra td { background: var(--amber-bg); }
  tr.st-ok td { background: var(--green-bg); }
  tr.st-dev td { background: var(--red-bg); }
  .hide { display: none; }
  a.lnk { color: var(--blue); cursor: pointer; text-decoration: underline; font-weight: 500; }
  .hint { color: var(--sub); font-size: 12px; }
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>BOM verification</h1>
    <span class="sub">Lot-wise check against the locked BOM</span>
    <div class="fill"></div>
    <div class="tabs">
      <div class="tab on" id="t1" onclick="showTab(1)">Verify a lot</div>
      <div class="tab" id="t2" onclick="showTab(2)">Lot-wise BOM</div>
    </div>
  </div>

  <div class="appbody">
  <!-- ===================== TAB 1 : VERIFY ===================== -->
  <div id="p1">
    <div class="toolbar">
      <label>Model <select id="model" onchange="loadBom()"></select></label>
      <label>Lot No <input type="text" id="lot" style="width:170px"></label>
      <label>Verified by <input type="text" id="by" style="width:160px"></label>
      <label>Remarks <input type="text" id="rem" class="grow"></label>
    </div>
    <div class="banner" id="banner">Loading...</div>
    <div class="toolbar">
      <button onclick="allOk()">Mark all OK</button>
      <button onclick="addExtra()">+ Extra part used in lot</button>
      <span class="hint">Untick OK when the lot used a different part / QPS, or did not use it. Then enter what the lot actually used.</span>
    </div>
    <div class="card">
      <div class="tbl-wrap" style="max-height:calc(100vh - 400px)">
        <table>
          <thead><tr><th style="width:44px">OK</th><th style="width:60px">Img</th><th style="width:150px">Part No</th><th>Description</th>
            <th style="width:80px">BOM QPS</th><th style="width:150px">Lot Part No</th><th style="width:90px">Lot QPS</th>
            <th style="width:160px">Remark</th><th style="width:110px">Result</th></tr></thead>
          <tbody id="rows"><tr><td colspan="9">Loading...</td></tr></tbody>
        </table>
      </div>
    </div>
    <div class="footbar">
      <span id="count" class="pill gray"></span>
      <span id="status" class="sp"></span>
      <button class="primary" id="save" onclick="save()">Save verification</button>
    </div>
  </div>

  <!-- ===================== TAB 2 : LOT-WISE BOM ===================== -->
  <div id="p2" class="hide">
    <div class="toolbar">
      <label>Model <select id="model2" onchange="loadLots()"></select></label>
      <label>Search lot <input type="text" id="lotSearch" oninput="renderLots()" class="grow"></label>
    </div>
    <div class="card">
      <div class="tbl-wrap" style="max-height:28vh">
        <table>
          <thead><tr><th>Lot No</th><th>Result</th><th>Deviations</th><th>BOM Rev</th><th>Date</th><th>Verified by</th><th>Ver ID</th><th></th></tr></thead>
          <tbody id="lots"><tr><td colspan="8">Choose a model.</td></tr></tbody>
        </table>
      </div>
    </div>
    <div id="lotInfo" style="margin:4px 0 10px;font-weight:600"></div>
    <div class="card">
      <div class="tbl-wrap" style="max-height:calc(100vh - 560px)">
        <table>
          <thead><tr><th style="width:60px">Img</th><th>Part No</th><th>Description</th><th>BOM QPS</th><th>Lot Part No</th><th>Lot QPS</th><th>Remark</th><th>Result</th></tr></thead>
          <tbody id="detail"><tr><td colspan="8">Select a lot and press View.</td></tr></tbody>
        </table>
      </div>
    </div>
    <div class="footbar">
      <span id="status2" class="sp"></span>
      <button id="exp" class="hide" onclick="exportLot()">Export this lot to sheet</button>
      <button id="csv" class="hide primary" onclick="downloadCsv()">Download CSV</button>
    </div>
  </div>
  </div>

<script>
  var rows = [], lots = [], models = [], curVer = '', armed = false;

  function el(id) { return document.getElementById(id); }

  function showTab(n) {
    el('p1').className = n === 1 ? '' : 'hide';
    el('p2').className = n === 2 ? '' : 'hide';
    el('t1').className = 'tab' + (n === 1 ? ' on' : '');
    el('t2').className = 'tab' + (n === 2 ? ' on' : '');
    if (n === 2) loadLots();
  }

  function init() {
    google.script.run.withSuccessHandler(function (m) {
      models = m;
      var opts = m.map(function (x) { return '<option>' + esc(x.name) + '</option>'; }).join('');
      el('model').innerHTML = opts;
      el('model2').innerHTML = opts;
      if (m.length) loadBom(); else el('banner').textContent = 'No model sheets found.';
    }).withFailureHandler(fail).getVerifyModels();
  }

  // ---------- verify tab ----------
  function loadBom() {
    var m = el('model').value;
    armed = false;
    el('rows').innerHTML = '<tr><td colspan="8">Loading...</td></tr>';
    google.script.run.withSuccessHandler(function (r) {
      rows = r.parts.map(function (p) {
        return { partNo: p.partNo, desc: p.desc, img: p.img, qps: p.qps, ok: true, ap: p.partNo, aq: p.qps, remark: '', extra: false };
      });
      var b = el('banner');
      if (r.locked) { b.className = 'banner good'; b.textContent = 'BOM Rev ' + r.rev + ' - LOCKED. Verifying against this revision.'; }
      else { b.className = 'banner warn'; b.textContent = 'This BOM is NOT locked. The log will record it as "Not locked". Lock the BOM first for controlled verification.'; }
      renderRows();
    }).withFailureHandler(fail).getVerifyBom(m);
  }

  function resultPill(res) {
    if (!res) return '';
    var cls = res === 'Matched' ? 'ok' : (res === 'Extra part' ? 'warn' : 'bad');
    return '<span class="pill ' + cls + '">' + esc(res) + '</span>';
  }

  function resultOf(r) {
    if (r.extra) return r.ap.trim() ? 'Extra part' : '';
    if (r.ok) return 'Matched';
    if (!r.ap.trim()) return 'Not used in lot';
    if (r.ap.trim().toUpperCase() !== r.partNo.toUpperCase()) return 'Part changed';
    if (Number(r.aq) !== Number(r.qps)) return 'QPS changed';
    return 'Matched';
  }

  function renderRows() {
    var html = rows.map(function (r, i) {
      var dis = (r.ok && !r.extra) ? ' disabled' : '';
      var res = resultOf(r);
      var cls = r.extra ? 'extra' : (r.ok ? '' : 'dev');
      return '<tr class="' + cls + '"><td>' +
        (r.extra ? '<button onclick="delExtra(' + i + ')">x</button>'
                 : '<input type="checkbox" ' + (r.ok ? 'checked ' : '') + 'onchange="setOk(' + i + ',this.checked)">') +
        '</td><td>' + (r.img ? '<img class="thumb" src="' + esc(r.img) + '">' : '') + '</td>' +
        '<td>' + (r.extra ? 'EXTRA' : esc(r.partNo)) + '</td><td>' + esc(r.desc) + '</td><td>' + esc(r.qps) + '</td>' +
        '<td><input type="text" value="' + esc(r.ap) + '"' + dis + ' oninput="rows[' + i + '].ap=this.value;renderRows()"></td>' +
        '<td><input type="number" step="any" value="' + esc(r.aq) + '"' + dis + ' oninput="rows[' + i + '].aq=this.value;renderRows()"></td>' +
        '<td><input type="text" value="' + esc(r.remark) + '" oninput="rows[' + i + '].remark=this.value"></td>' +
        '<td>' + resultPill(res) + '</td></tr>';
    });
    el('rows').innerHTML = html.length ? html.join('') : '<tr><td colspan="9">This model has no BOM parts.</td></tr>';
    updateCount();
  }

  function setOk(i, c) {
    armed = false;
    rows[i].ok = c;
    if (c) { rows[i].ap = rows[i].partNo; rows[i].aq = rows[i].qps; }
    renderRows();
  }

  function allOk() {
    rows.forEach(function (r) { if (!r.extra) { r.ok = true; r.ap = r.partNo; r.aq = r.qps; } });
    renderRows();
  }

  function addExtra() {
    rows.push({ partNo: '', desc: '', qps: '', ok: false, ap: '', aq: '', remark: '', extra: true });
    renderRows();
  }

  function delExtra(i) { rows.splice(i, 1); renderRows(); }

  function updateCount() {
    var dev = 0, n = 0;
    rows.forEach(function (r) { var res = resultOf(r); if (r.extra) { if (res) dev++; } else { n++; if (res !== 'Matched') dev++; } });
    el('count').textContent = n + ' BOM parts - ' + dev + ' marked as deviation';
  }

  function save() {
    var lot = el('lot').value.trim();
    if (!lot) return setStatus('status', 'Enter the Lot No.', true);
    if (!rows.length) return setStatus('status', 'Nothing to verify.', true);
    el('save').disabled = true;
    setStatus('status', 'Saving...');
    var payload = {
      model: el('model').value, lot: lot, by: el('by').value, remarks: el('rem').value,
      rows: rows.map(function (r) {
        return { partNo: r.partNo, desc: r.desc, img: r.img, qps: r.qps, ok: r.ok, actualPart: r.ap, actualQps: r.aq, remark: r.remark, extra: r.extra };
      })
    };
    google.script.run.withSuccessHandler(function (r) {
      el('save').disabled = false;
      setStatus('status', r.verId + ' saved for lot ' + lot + ': ' + r.result +
        (r.deviations ? ' (' + r.deviations + ' deviation)' : ''), r.deviations ? true : false);
      el('lot').value = ''; el('rem').value = '';
      loadBom();
    }).withFailureHandler(function (e) { el('save').disabled = false; fail(e); }).saveVerification(payload);
  }

  // ---------- lot-wise tab ----------
  function loadLots() {
    var m = el('model2').value;
    if (!m) return;
    el('lots').innerHTML = '<tr><td colspan="8">Loading...</td></tr>';
    el('exp').className = 'hide'; el('csv').className = 'hide';
    google.script.run.withSuccessHandler(function (l) { lots = l; renderLots(); })
      .withFailureHandler(fail).getLotList(m);
  }

  function renderLots() {
    var q = el('lotSearch').value.trim().toUpperCase();
    var list = lots.filter(function (l) { return !q || l.lot.toUpperCase().indexOf(q) >= 0; });
    el('lots').innerHTML = list.length ? list.map(function (l) {
      return '<tr><td>' + esc(l.lot) + '</td><td><span class="pill ' + (l.result === 'All matched' ? 'ok' : 'bad') + '">' + esc(l.result) + '</span>' +
        '</td><td>' + esc(l.dev) + '</td><td>' + esc(l.rev) + '</td><td>' + esc(l.when) + '</td><td>' + esc(l.by) +
        '</td><td>' + esc(l.verId) + '</td><td><a class="lnk" onclick="viewLot(\\'' + esc(l.verId) + '\\')">View</a></td></tr>';
    }).join('') : '<tr><td colspan="8">No verified lots for this model.</td></tr>';
  }

  function viewLot(verId) {
    curVer = verId;
    google.script.run.withSuccessHandler(function (r) {
      var i = r.info;
      el('lotInfo').textContent = i ? ('Lot ' + i.lot + ' - BOM Rev ' + i.rev + ' (' + i.bomStatus + ') - ' + i.result +
        (i.remarks ? ' - ' + i.remarks : '') + (i.checked === '0' ? ' - 0 parts checked' : '')) : '';
      if (!r.rows.length) {
        el('detail').innerHTML = '<tr><td colspan="8">No part rows were recorded for this verification '
          + '(its BOM had 0 parts at the time it was saved, or it predates part-level logging). '
          + 'Re-verify this lot to get a complete record.</td></tr>';
        el('exp').className = 'hide'; el('csv').className = 'hide';
      } else {
        el('detail').innerHTML = r.rows.map(function (d) {
          return '<tr>' +
            '<td>' + (d.img ? '<img class="thumb" src="' + esc(d.img) + '">' : '') + '</td>' +
            '<td>' + esc(d.part) + '</td><td>' + esc(d.desc) +
            '</td><td>' + esc(d.bomQps) + '</td><td>' + esc(d.lotPart) + '</td><td>' + esc(d.lotQps) + '</td><td>' + esc(d.remark) +
            '</td><td>' + resultPill(d.result) + '</td></tr>';
        }).join('');
        el('exp').className = ''; el('csv').className = '';
      }
    }).withFailureHandler(fail).getLotBom(verId);
  }

  function exportLot() {
    setStatus('status2', 'Exporting...');
    google.script.run.withSuccessHandler(function () { setStatus('status2', 'Exported to sheet LOT_BOM_VIEW.', false); })
      .withFailureHandler(fail).exportLotBom(curVer);
  }

  function downloadCsv() {
    setStatus('status2', 'Preparing CSV...');
    google.script.run.withSuccessHandler(function (r) {
      var blob = new Blob([r.csv], { type: 'text/csv;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = r.filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
      setStatus('status2', 'Downloaded ' + r.filename, false);
    }).withFailureHandler(fail).getLotCsv(curVer);
  }

  function fail(e) { setStatus('status', 'Error: ' + (e && e.message ? e.message : e), true); setStatus('status2', 'Error: ' + (e && e.message ? e.message : e), true); }
  function setStatus(id, t, isErr) {
    var s = el(id);
    s.textContent = t;
    s.className = 'sp ' + (isErr === true ? 'err' : (isErr === false ? 'ok' : ''));
  }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  init();
</script>
</body>
</html>
`;
}

/** Screen HTML (embedded). */
function getTwoCompareHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 90px; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; }

input[type=text], input[type=number], select, textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
  transition: border-color .15s;
}
input[type=text]:focus, input[type=number]:focus, select:focus {
  outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg);
}
input[type=number] { width: 78px; }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }

.tabs { display: flex; gap: 4px; }
.tab {
  padding: 9px 16px; cursor: pointer; border-radius: var(--radius-sm) var(--radius-sm) 0 0;
  background: transparent; color: var(--sub); font-weight: 500; border-bottom: 2px solid transparent;
}
.tab:hover { background: #f1f3f4; }
.tab.on { color: var(--blue); border-bottom: 2px solid var(--blue); background: var(--blue-bg); }

.footbar {
  position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 14px; align-items: center;
  padding: 12px 22px; background: var(--surface); border-top: 1px solid var(--line-soft); box-shadow: 0 -2px 8px rgba(0,0,0,.05);
}
.footbar .sp { flex: 1; }
.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }

img.thumb { width: 40px; height: 40px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

  .tab .n { font-weight: 700; margin-left: 5px; }
  tr.same td { background: var(--green-bg); }
  tr.diff td { background: var(--red-bg); }
  tr.a td { background: var(--amber-bg); }
  tr.b td { background: var(--blue-bg); }
  .ctr { text-align: center; }
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>Compare two models</h1>
    <span class="sub" id="info">Loading...</span>
    <div class="fill"></div>
  </div>

  <div class="appbody">
    <div class="toolbar">
      <label>Model A <select id="ma" onchange="run()"></select></label>
      <button onclick="swap()">⇄ Swap</button>
      <label>Model B <select id="mb" onchange="run()"></select></label>
    </div>

    <div class="tabs" style="margin-bottom:12px">
      <div class="tab on" id="tc" onclick="setTab('c')">Common parts <span class="n" id="nc">0</span></div>
      <div class="tab" id="ta" onclick="setTab('a')">Only in A <span class="n" id="na">0</span></div>
      <div class="tab" id="tb" onclick="setTab('b')">Only in B <span class="n" id="nb">0</span></div>
    </div>

    <div class="toolbar">
      <input type="text" id="search" placeholder="Search part no / description" oninput="render()" class="grow">
      <label id="diffLbl" style="font-weight:normal"><input type="checkbox" id="diffOnly" onchange="render()"> Show only different QPS</label>
    </div>

    <div class="card">
      <div class="tbl-wrap" style="max-height:calc(100vh - 330px)">
        <table>
          <thead id="head"></thead>
          <tbody id="body"><tr><td>Loading...</td></tr></tbody>
        </table>
      </div>
    </div>
  </div>

  <div class="footbar">
    <span id="status" class="sp"></span>
    <button class="primary" id="exp" onclick="exportIt()">Export to sheet</button>
  </div>

<script>
  var res = null, tab = 'c';
  function el(id) { return document.getElementById(id); }

  function init() {
    google.script.run.withSuccessHandler(function (m) {
      var opts = m.map(function (x) { return '<option>' + esc(x.name) + '</option>'; }).join('');
      el('ma').innerHTML = opts;
      el('mb').innerHTML = opts;
      if (m.length > 1) el('mb').selectedIndex = 1;
      if (m.length < 2) { el('info').textContent = 'You need at least two model sheets to compare.'; el('body').innerHTML = ''; return; }
      run();
    }).withFailureHandler(fail).getVerifyModels();
  }

  function swap() {
    var a = el('ma').value, b = el('mb').value;
    el('ma').value = b; el('mb').value = a;
    run();
  }

  function run() {
    var a = el('ma').value, b = el('mb').value;
    if (!a || !b) return;
    if (a === b) { el('info').textContent = 'Choose two different models.'; el('body').innerHTML = ''; return; }
    setStatus('Comparing...');
    google.script.run.withSuccessHandler(function (r) {
      res = r; setStatus('');
      var lk = function (m) { return m.name + ' (' + m.count + ' parts, ' + (m.locked ? 'LOCKED Rev ' + m.rev : 'not locked') + ')'; };
      var diff = r.common.filter(function (x) { return !x.same; }).length;
      el('info').textContent = 'A: ' + lk(r.a) + '   |   B: ' + lk(r.b) + '   |   ' + diff + ' common part(s) have a different QPS';
      el('nc').textContent = r.common.length;
      el('na').textContent = r.onlyA.length;
      el('nb').textContent = r.onlyB.length;
      render();
    }).withFailureHandler(fail).compareTwoModels(a, b);
  }

  function setTab(t) {
    tab = t;
    el('tc').className = 'tab' + (t === 'c' ? ' on' : '');
    el('ta').className = 'tab' + (t === 'a' ? ' on' : '');
    el('tb').className = 'tab' + (t === 'b' ? ' on' : '');
    render();
  }

  function render() {
    if (!res) return;
    var q = el('search').value.trim().toUpperCase();
    var match = function (x) { return !q || x.part.toUpperCase().indexOf(q) >= 0 || (x.desc || '').toUpperCase().indexOf(q) >= 0; };
    el('diffLbl').style.display = tab === 'c' ? '' : 'none';
    var a = res.a.name, b = res.b.name, html = '', list;

    var img = function (u) { return u ? '<img class="thumb" src="' + esc(u) + '">' : ''; };

    if (tab === 'c') {
      var dOnly = el('diffOnly').checked;
      list = res.common.filter(function (x) { return match(x) && (!dOnly || !x.same); });
      el('head').innerHTML = '<tr><th style="width:50px">Img</th><th>Part No</th><th>Description</th><th class="ctr">QPS in ' + esc(a) + '</th><th class="ctr">QPS in ' + esc(b) + '</th><th class="ctr">QPS check</th></tr>';
      html = list.map(function (x) {
        return '<tr class="' + (x.same ? 'same' : 'diff') + '"><td>' + img(x.img) + '</td><td>' + esc(x.part) + '</td><td>' + esc(x.desc) + '</td><td class="ctr">' + esc(x.qa) +
          '</td><td class="ctr">' + esc(x.qb) + '</td><td class="ctr"><span class="pill ' + (x.same ? 'ok' : 'bad') + '">' + (x.same ? 'Same' : 'Different') + '</span></td></tr>';
      }).join('');
    } else {
      list = (tab === 'a' ? res.onlyA : res.onlyB).filter(match);
      el('head').innerHTML = '<tr><th style="width:50px">Img</th><th>Part No</th><th>Description</th><th class="ctr">QPS in ' + esc(tab === 'a' ? a : b) + '</th></tr>';
      html = list.map(function (x) {
        return '<tr class="' + tab + '"><td>' + img(x.img) + '</td><td>' + esc(x.part) + '</td><td>' + esc(x.desc) + '</td><td class="ctr">' + esc(x.q) + '</td></tr>';
      }).join('');
    }
    el('body').innerHTML = html || '<tr><td colspan="5">Nothing to show.</td></tr>';
  }

  function exportIt() {
    if (!res) return;
    el('exp').disabled = true;
    setStatus('Exporting...');
    google.script.run.withSuccessHandler(function () {
      el('exp').disabled = false; setStatus('Exported to sheet MODEL_COMPARE.', false);
    }).withFailureHandler(function (e) { el('exp').disabled = false; fail(e); }).exportTwoCompare(res.a.name, res.b.name);
  }

  function fail(e) { setStatus('Error: ' + (e && e.message ? e.message : e), true); }
  function setStatus(t, isErr) {
    var s = el('status');
    s.textContent = t;
    s.className = 'sp ' + (isErr === true ? 'err' : (isErr === false ? 'ok' : ''));
  }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  init();
</script>
</body>
</html>
`;
}


/** Screen HTML (embedded): Generate ECN wizard. */
function getEcnHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 90px; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; }
.card .card-b { padding: 14px 16px; }

input[type=text], input[type=number], input[type=date], input[type=file], select, textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
  transition: border-color .15s;
}
input[type=text]:focus, input[type=number]:focus, input[type=date]:focus, select:focus, textarea:focus {
  outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg);
}
input[type=number] { width: 90px; }
textarea { width: 100%; min-height: 64px; resize: vertical; font-family: 'Roboto Mono', Consolas, monospace; }
textarea.prose { font-family: inherit; }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button.ghost { background: transparent; border-color: transparent; color: var(--blue); }
button.ghost:hover { background: var(--blue-bg); }
button.sm { padding: 5px 10px; font-size: 12px; }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }

.tabs { display: flex; gap: 4px; }
.tab {
  padding: 9px 16px; cursor: pointer; border-radius: var(--radius-sm) var(--radius-sm) 0 0;
  background: transparent; color: var(--sub); font-weight: 500; border-bottom: 2px solid transparent;
}
.tab:hover { background: #f1f3f4; }
.tab.on { color: var(--blue); border-bottom: 2px solid var(--blue); background: var(--blue-bg); }
.tabpane { display: none; } .tabpane.on { display: block; }

.footbar {
  position: fixed; left: 0; right: 0; bottom: 0; display: flex; gap: 14px; align-items: center;
  padding: 12px 22px; background: var(--surface); border-top: 1px solid var(--line-soft); box-shadow: 0 -2px 8px rgba(0,0,0,.05);
}
.footbar .sp { flex: 1; }
.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }

img.thumb { width: 40px; height: 40px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

.field { margin-bottom: 14px; }
.field > label { display: block; font-weight: 600; font-size: 12.5px; margin-bottom: 6px; }
.field .hint { color: var(--sub); font-size: 11.5px; margin-top: 4px; }
.row2 { display: flex; gap: 14px; flex-wrap: wrap; }
.row2 > .field { flex: 1; min-width: 220px; }
.check-list { margin-top: 8px; display: flex; flex-direction: column; gap: 6px; }
.check-row { padding: 8px 10px; border-radius: var(--radius-sm); border: 1px solid var(--line-soft); background: #fafbfc; display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }
.after-rows { display: flex; flex-direction: column; gap: 8px; }
.after-row { display: flex; gap: 10px; align-items: center; background: #fafbfc; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); padding: 8px 10px; }
.after-row input[type=text] { flex: 1; }
.after-row .afterDesc { flex: 1.4; color: var(--sub); font-size: 12px; }
.seg { display: inline-flex; border: 1px solid var(--line); border-radius: var(--radius-sm); overflow: hidden; }
.seg button { border: none; border-radius: 0; background: var(--surface); }
.seg button.on { background: var(--blue); color: #fff; }
.uploadbox { border: 1px dashed var(--line); border-radius: var(--radius-sm); padding: 12px; margin-top: 10px; background: #fafbfc; }
.resultCard { border: 1px solid #bfe4c8; background: var(--green-bg); border-radius: var(--radius); padding: 14px 16px; margin-bottom: 14px; }
.resultCard h3 { margin: 0 0 8px; color: var(--green); font-size: 15px; }
.resultCard a { color: var(--blue-dark); }
.bulkbox { border: 1px dashed var(--line); border-radius: var(--radius-sm); padding: 12px; margin-bottom: 10px; background: #fafbfc; }
.bulkbox .hd { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
.editRow td { background: #fafbfc; }
.editPanel { padding: 12px; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fafbfc; }
.partsTag { display: inline-block; background: #f1f3f4; border-radius: 6px; padding: 2px 7px; margin: 2px; font-size: 11.5px; }
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>Generate ECN</h1>
    <span class="sub">Replace or split parts on a model's BOM, with an auto-numbered Engineering Change Number</span>
    <div class="fill"></div>
  </div>

  <div class="appbody">
    <div class="tabs">
      <div class="tab on" id="tabNew" onclick="showTab('new')">New ECN</div>
      <div class="tab" id="tabLog" onclick="showTab('log')">ECN history / edit</div>
    </div>

    <div class="tabpane on" id="paneNew" style="margin-top:12px">
      <div id="resultArea"></div>

      <div class="card">
        <div class="card-h">1. Model &amp; part(s) being replaced</div>
        <div class="card-b">
          <div class="field">
            <label>Model</label>
            <select id="model" onchange="onModelChange()" style="width:100%;max-width:420px"><option value="">Loading...</option></select>
          </div>
          <div class="field">
            <label>Before (old) part number(s)</label>
            <textarea id="oldParts" class="prose" placeholder="One part number per line - paste a whole column straight from Excel if you're replacing several at once" oninput="debounceCheckOld()" style="min-height:84px"></textarea>
            <div class="hint">Each line is checked against the selected model's current BOM below.</div>
          </div>
          <div class="check-list" id="oldCheckList"></div>
        </div>
      </div>

      <div class="card">
        <div class="card-h">
          2. After (new) part number(s)
          <span class="sub" style="font-weight:400">— one old part can split into several new ones</span>
        </div>
        <div class="card-b">
          <div class="bulkbox">
            <div class="hd">
              <b style="font-size:12.5px">Bulk paste</b>
              <span class="sub" style="font-size:11.5px">One part per line - "PARTCODE" alone, or "PARTCODE, QPS" / "PARTCODE&nbsp;&nbsp;QPS" (tab or space works too)</span>
            </div>
            <textarea id="bulkAfter" class="prose" placeholder="e.g.&#10;NEW-PART-101, 2&#10;NEW-PART-102, 1" style="min-height:70px"></textarea>
            <div style="margin-top:8px"><button class="sm" onclick="addBulkAfterRows()">Add pasted rows</button></div>
          </div>
          <div class="after-rows" id="afterRows"></div>
          <div style="margin-top:10px"><button class="ghost" onclick="addAfterRow()">+ Add one more new part</button></div>
        </div>
      </div>

      <div class="card">
        <div class="card-h">3. ECN document</div>
        <div class="card-b">
          <div class="field">
            <label>Have you received the ECN document?</label>
            <div class="seg">
              <button type="button" id="recvYes" onclick="setReceived(true)">Yes</button>
              <button type="button" id="recvNo" onclick="setReceived(false)">No</button>
            </div>
          </div>
          <div class="uploadbox" id="uploadBox" style="display:none">
            <label style="font-weight:600;font-size:12.5px;display:block;margin-bottom:8px">Upload the ECN document (PDF)</label>
            <input type="file" id="docFile" accept="application/pdf" onchange="onFile()">
            <div class="hint" id="fileInfo"></div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-h">4. Reason &amp; effective date</div>
        <div class="card-b">
          <div class="row2">
            <div class="field" style="flex:2">
              <label>Reason for change</label>
              <textarea id="reason" class="prose" placeholder="Why is this part being changed?"></textarea>
            </div>
            <div class="field">
              <label>Effective date</label>
              <input type="date" id="effDate" style="width:100%">
            </div>
          </div>
        </div>
      </div>

      <div class="banner info">Generating an ECN always writes a new row to <b>ECN_HISTORY</b> per part changed and bumps the model's BOM revision once, whether or not the BOM is currently locked. Listing several before-parts or several after-parts runs them all under one ECN number.</div>
    </div>

    <div class="tabpane" id="paneLog">
      <div class="toolbar">
        <label class="grow">Filter by model
          <select id="logModel" onchange="loadLog()" style="width:100%"><option value="">All models</option></select>
        </label>
        <button onclick="loadLog()">Refresh</button>
      </div>
      <div class="banner info">Editing an ECN here only corrects its reason, effective date or saved document. It never changes which parts were added or removed - that's the permanent record of what happened on the BOM. To change the parts themselves, generate a new ECN.</div>
      <div class="card">
        <div class="tbl-wrap">
          <table>
            <thead><tr><th>Date</th><th>ECN No</th><th>Model</th><th>Rev</th><th>Parts changed</th><th>Reason</th><th>Effective</th><th>Document</th><th>By</th><th></th></tr></thead>
            <tbody id="logRows"><tr><td colspan="10">Switch to this tab to load.</td></tr></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <div class="footbar">
    <span id="status" class="sp"></span>
    <button onclick="resetForm()">Reset form</button>
    <button class="primary" id="submitBtn" onclick="submitEcn()">Generate ECN</button>
  </div>

<script>
  var models = [], received = null, fileBase64 = '', fileMime = '', fileName = '', afterSeq = 0;
  var oldCheckTimer = null;
  var ecnGroups = {}; // ecnNo -> group, for the history/edit tab
  var afterModes = {}; // after-row id -> { existingQps, mode: 'replace'|'add'|'subtract' } when that part is already on the model's BOM

  function showTab(t) {
    document.getElementById('tabNew').classList.toggle('on', t === 'new');
    document.getElementById('tabLog').classList.toggle('on', t === 'log');
    document.getElementById('paneNew').classList.toggle('on', t === 'new');
    document.getElementById('paneLog').classList.toggle('on', t === 'log');
    if (t === 'log' && !document.getElementById('logRows').dataset.loaded) loadLog();
  }

  function init() {
    google.script.run.withSuccessHandler(function (d) { models = d; renderModels(); })
      .withFailureHandler(fail).getVerifyModels();
    google.script.run.withSuccessHandler(function (d) {
      document.getElementById('effDate').value = d;
    }).withFailureHandler(function(){}).getTodayStr();
    addAfterRow();
  }

  function renderModels() {
    var opts = '<option value="">Select a model...</option>' +
      models.map(function (m) { return '<option value="' + esc(m.name) + '">' + esc(m.name) + (m.locked ? ' (Locked)' : '') + '</option>'; }).join('');
    document.getElementById('model').innerHTML = opts;
    document.getElementById('logModel').innerHTML = '<option value="">All models</option>' +
      models.map(function (m) { return '<option value="' + esc(m.name) + '">' + esc(m.name) + '</option>'; }).join('');
  }

  function splitLines(s) {
    return String(s || '').split(/[\\r\\n]+/).map(function (x) { return x.trim(); }).filter(Boolean);
  }

  function onModelChange() {
    document.getElementById('oldCheckList').innerHTML = '';
    checkOld();
    recheckAllAfterRows();
  }

  function recheckAllAfterRows() {
    document.querySelectorAll('.after-row').forEach(function (row) { checkAfter(row.id); });
  }

  function debounceCheckOld() {
    clearTimeout(oldCheckTimer);
    oldCheckTimer = setTimeout(checkOld, 400);
  }

  function checkOld() {
    var model = document.getElementById('model').value;
    var parts = splitLines(document.getElementById('oldParts').value);
    var box = document.getElementById('oldCheckList');
    if (!model || !parts.length) { box.innerHTML = ''; return; }
    google.script.run.withSuccessHandler(function (rows) {
      box.innerHTML = rows.map(function (r) {
        if (!r.found) {
          return '<div class="check-row"><span class="pill bad">Not found</span><span>' + esc(r.partNo) + '</span>' +
            '<span class="sub">not in ' + esc(model) + '\\'s current BOM</span></div>';
        }
        return '<div class="check-row"><span class="pill ok">Found</span>' +
          (r.img ? '<img class="thumb" src="' + esc(r.img) + '">' : '') +
          '<span><b>' + esc(r.partNo) + '</b> - ' + esc(r.desc || '(no description)') + '</span>' +
          '<span class="pill info">QPS ' + esc(r.qps) + '</span>' +
          (r.locked ? '<span class="pill warn">BOM locked - rev ' + esc(r.rev) + '</span>' : '<span class="pill gray">Rev ' + esc(r.rev) + '</span>') +
          '</div>';
      }).join('');
    }).withFailureHandler(function (e) {
      box.innerHTML = '<div class="check-row"><span class="pill bad">Error</span><span class="sub">' + esc(e && e.message ? e.message : e) + '</span></div>';
    }).getPartsInModel(model, parts);
  }

  function addAfterRow(partNo, qps) {
    afterSeq++;
    var id = 'ar' + afterSeq;
    var div = document.createElement('div');
    div.className = 'after-row';
    div.id = id;
    div.innerHTML =
      '<input type="text" placeholder="New part number" value="' + esc(partNo || '') + '" oninput="checkAfter(\\'' + id + '\\')">' +
      '<span>QPS</span><input type="number" min="0.0001" step="any" value="' + esc(qps || 1) + '" style="width:80px" oninput="onAfterQpsInput(\\'' + id + '\\')">' +
      '<span class="afterDesc"></span>' +
      '<span class="afterExist" id="exist_' + id + '"></span>' +
      '<button class="ghost" onclick="removeAfterRow(\\'' + id + '\\')" title="Remove">\u2715</button>';
    document.getElementById('afterRows').appendChild(div);
    if (partNo) checkAfter(id);
  }

  function onAfterQpsInput(id) { renderExistBox(id); }

  function addBulkAfterRows() {
    var lines = splitLines(document.getElementById('bulkAfter').value);
    if (!lines.length) return;
    lines.forEach(function (line) {
      var parts = line.split(/[,\\t]+|\\s{2,}/).map(function (x) { return x.trim(); }).filter(Boolean);
      if (parts.length < 2) parts = line.split(/\\s+/).map(function (x) { return x.trim(); }).filter(Boolean);
      var partNo = parts[0];
      var qps = parts.length > 1 && Number(parts[1]) > 0 ? Number(parts[1]) : 1;
      if (partNo) addAfterRow(partNo, qps);
    });
    document.getElementById('bulkAfter').value = '';
  }

  function removeAfterRow(id) {
    var rows = document.getElementById('afterRows');
    if (rows.children.length <= 1) return;
    var el = document.getElementById(id);
    if (el) el.remove();
    delete afterModes[id];
  }

  function checkAfter(id) {
    var row = document.getElementById(id);
    if (!row) return;
    var part = row.querySelector('input[type=text]').value.trim();
    var descEl = row.querySelector('.afterDesc');
    delete afterModes[id];
    renderExistBox(id);
    if (!part) { descEl.textContent = ''; return; }
    google.script.run.withSuccessHandler(function (r) {
      descEl.textContent = r.found ? ('\u2713 ' + (r.desc || '(no description)')) : '\u2717 not in PART_MASTER - add it there first';
      descEl.style.color = r.found ? '' : '#d93025';
    }).withFailureHandler(function () { descEl.textContent = ''; }).getMasterPart(part);

    var model = document.getElementById('model').value;
    if (!model) return;
    google.script.run.withSuccessHandler(function (r) {
      // Re-check the part number is still the one this response is for - the field may have changed since the call went out.
      var curRow = document.getElementById(id);
      if (!curRow || curRow.querySelector('input[type=text]').value.trim() !== part) return;
      if (r.found) {
        afterModes[id] = { existingQps: Number(r.qps) || 0, mode: (afterModes[id] && afterModes[id].mode) || 'replace' };
      } else {
        delete afterModes[id];
      }
      renderExistBox(id);
    }).withFailureHandler(function () {}).getPartInModel(model, part);
  }

  function renderExistBox(id) {
    var box = document.getElementById('exist_' + id);
    if (!box) return;
    var m = afterModes[id];
    if (!m) { box.innerHTML = ''; return; }
    var typed = Number((document.getElementById(id).querySelectorAll('input[type=number]')[0] || {}).value) || 0;
    var finalQ = m.mode === 'add' ? m.existingQps + typed : m.mode === 'subtract' ? m.existingQps - typed : typed;
    box.innerHTML =
      '<span class="pill warn">Already in BOM: QPS ' + esc(m.existingQps) + '</span>' +
      '<span class="seg" style="margin-left:6px">' +
        '<button type="button" class="sm' + (m.mode === 'add' ? ' on' : '') + '" onclick="setAfterMode(\\'' + id + '\\',\\'add\\')">+ Add</button>' +
        '<button type="button" class="sm' + (m.mode === 'subtract' ? ' on' : '') + '" onclick="setAfterMode(\\'' + id + '\\',\\'subtract\\')">- Subtract</button>' +
        '<button type="button" class="sm' + (m.mode === 'replace' ? ' on' : '') + '" onclick="setAfterMode(\\'' + id + '\\',\\'replace\\')">= Replace</button>' +
      '</span>' +
      '<span class="pill ' + (finalQ > 0 ? 'info' : 'bad') + '" style="margin-left:6px">New QPS will be ' + esc(finalQ) + '</span>';
  }

  function setAfterMode(id, mode) {
    if (!afterModes[id]) return;
    afterModes[id].mode = mode;
    renderExistBox(id);
  }

  function setReceived(v) {
    received = v;
    document.getElementById('recvYes').className = v ? 'on' : '';
    document.getElementById('recvNo').className = !v ? 'on' : '';
    document.getElementById('uploadBox').style.display = v ? 'block' : 'none';
    if (!v) { fileBase64 = ''; fileMime = ''; fileName = ''; document.getElementById('fileInfo').textContent = ''; document.getElementById('docFile').value = ''; }
  }

  function onFile() {
    var f = document.getElementById('docFile').files[0];
    var info = document.getElementById('fileInfo');
    if (!f) { fileBase64 = ''; fileMime = ''; fileName = ''; info.textContent = ''; return; }
    if (f.type !== 'application/pdf') { info.textContent = 'Please choose a PDF file.'; info.style.color = '#d93025'; document.getElementById('docFile').value = ''; return; }
    info.style.color = ''; info.textContent = 'Reading ' + f.name + '...';
    var reader = new FileReader();
    reader.onload = function (e) {
      var res = e.target.result;
      fileBase64 = res.substring(res.indexOf(',') + 1);
      fileMime = f.type; fileName = f.name;
      info.textContent = 'Ready: ' + f.name + ' (' + Math.round(f.size / 1024) + ' KB)';
    };
    reader.onerror = function () { info.textContent = 'Could not read that file.'; info.style.color = '#d93025'; };
    reader.readAsDataURL(f);
  }

  function collectAfterParts() {
    var out = [];
    document.querySelectorAll('.after-row').forEach(function (row) {
      var inputs = row.querySelectorAll('input');
      var part = inputs[0].value.trim();
      var typed = Number(inputs[1].value);
      if (!part) return;
      var m = afterModes[row.id];
      var qps = typed;
      if (m) {
        if (m.mode === 'add') qps = m.existingQps + typed;
        else if (m.mode === 'subtract') qps = m.existingQps - typed;
        if (!(qps > 0))
          throw new Error(part + ' is already in the BOM at QPS ' + m.existingQps + ' - that ' + m.mode + ' would bring it to ' + qps + ', which isn\\'t valid. Adjust the QPS or choose Replace.');
      }
      out.push({ partNo: part, qps: qps });
    });
    return out;
  }

  function submitEcn() {
    var model = document.getElementById('model').value;
    var oldParts = splitLines(document.getElementById('oldParts').value);
    var newParts;
    try { newParts = collectAfterParts(); } catch (e) { return fail(e.message); }
    var reason = document.getElementById('reason').value.trim();
    var effectiveDate = document.getElementById('effDate').value;

    if (!model) return fail('Select a model.');
    if (!oldParts.length) return fail('Enter at least one before (old) part number.');
    if (!newParts.length) return fail('Enter at least one after (new) part number.');
    if (received === null) return fail('Answer whether the ECN document has been received.');
    if (received && !fileBase64) return fail('Upload the ECN document, or answer "No" if you have not received it yet.');
    if (!reason) return fail('Enter the reason for the change.');
    if (!effectiveDate) return fail('Enter the effective date.');

    var payload = {
      model: model, oldParts: oldParts, newParts: newParts,
      received: !!received, docBase64: received ? fileBase64 : '', docMime: fileMime, docName: fileName,
      reason: reason, effectiveDate: effectiveDate
    };

    document.getElementById('submitBtn').disabled = true;
    setStatus('Generating ECN...');
    google.script.run.withSuccessHandler(function (r) {
      document.getElementById('submitBtn').disabled = false;
      setStatus('Done.', false);
      showResult(r);
      if (document.getElementById('logRows').dataset.loaded) loadLog();
    }).withFailureHandler(function (e) {
      document.getElementById('submitBtn').disabled = false;
      fail(e);
    }).runEcnChange(payload);
  }

  function showResult(r) {
    var removed = Array.isArray(r.removed) ? r.removed.join(', ') : r.removed;
    var html = '<div class="resultCard"><h3>' + esc(r.ecnNo) + ' generated for ' + esc(r.model) + ' (rev ' + esc(r.rev) + ')</h3>' +
      '<div>Removed: <b>' + esc(removed) + '</b></div>' +
      '<div>Added: <b>' + esc((r.added || []).join(', ')) + '</b></div>' +
      (r.docUrl ? '<div style="margin-top:6px"><a href="' + esc(r.docUrl) + '" target="_blank">Open saved ECN document</a></div>' : '') +
      '</div>';
    document.getElementById('resultArea').innerHTML = html;
  }

  function resetForm() {
    document.getElementById('oldParts').value = '';
    document.getElementById('oldCheckList').innerHTML = '';
    document.getElementById('bulkAfter').value = '';
    document.getElementById('afterRows').innerHTML = '';
    afterModes = {};
    addAfterRow();
    received = null;
    document.getElementById('recvYes').className = '';
    document.getElementById('recvNo').className = '';
    document.getElementById('uploadBox').style.display = 'none';
    fileBase64 = ''; fileMime = ''; fileName = '';
    document.getElementById('fileInfo').textContent = '';
    document.getElementById('docFile').value = '';
    document.getElementById('reason').value = '';
    document.getElementById('resultArea').innerHTML = '';
    setStatus('');
  }

  // ---------------------------- ECN history / edit ----------------------------

  function loadLog() {
    var model = document.getElementById('logModel').value;
    document.getElementById('logRows').innerHTML = '<tr><td colspan="10">Loading...</td></tr>';
    google.script.run.withSuccessHandler(function (rows) {
      document.getElementById('logRows').dataset.loaded = '1';
      ecnGroups = {};
      var order = [];
      rows.forEach(function (r) {
        if (!ecnGroups[r.ecnNo]) { ecnGroups[r.ecnNo] = { ecnNo: r.ecnNo, when: r.when, model: r.model, rev: r.rev, reason: r.reason, effectiveDate: r.effectiveDate, docUrl: r.docUrl, by: r.by, parts: [] }; order.push(r.ecnNo); }
        ecnGroups[r.ecnNo].parts.push(r.part + ' (' + r.action + ')');
      });
      renderLogGroups(order);
    }).withFailureHandler(fail).getEcnLog(model);
  }

  function renderLogGroups(order) {
    if (!order.length) { document.getElementById('logRows').innerHTML = '<tr><td colspan="10">No ECN history yet.</td></tr>'; return; }
    document.getElementById('logRows').innerHTML = order.map(function (no) {
      var g = ecnGroups[no];
      return '<tr id="row_' + esc(no) + '"><td>' + esc(g.when) + '</td><td>' + esc(g.ecnNo) + '</td><td>' + esc(g.model) + '</td><td>' + esc(g.rev) + '</td>' +
        '<td>' + g.parts.map(function (p) { return '<span class="partsTag">' + esc(p) + '</span>'; }).join('') + '</td>' +
        '<td>' + esc(g.reason) + '</td><td>' + esc(g.effectiveDate) + '</td>' +
        '<td>' + (g.docUrl ? '<a href="' + esc(g.docUrl) + '" target="_blank">Open</a>' : '-') + '</td>' +
        '<td>' + esc(g.by) + '</td>' +
        '<td><button class="sm" onclick="toggleEdit(\\'' + esc(no) + '\\')">Edit</button> ' +
        '<button class="sm danger" onclick="confirmDeleteEcn(\\'' + esc(no) + '\\')">Delete</button></td></tr>';
    }).join('');
  }

  function toggleEdit(ecnNo) {
    var existing = document.getElementById('edit_' + ecnNo);
    if (existing) { existing.remove(); return; }
    var g = ecnGroups[ecnNo];
    var row = document.getElementById('row_' + ecnNo);
    var tr = document.createElement('tr');
    tr.className = 'editRow';
    tr.id = 'edit_' + ecnNo;
    tr.innerHTML = '<td colspan="10"><div class="editPanel">' +
      '<div class="row2">' +
        '<div class="field" style="flex:2"><label>Reason</label><textarea class="prose" id="editReason_' + esc(ecnNo) + '" style="min-height:48px">' + esc(g.reason) + '</textarea></div>' +
        '<div class="field"><label>Effective date</label><input type="date" id="editDate_' + esc(ecnNo) + '" value="' + esc(g.effectiveDate) + '" style="width:100%"></div>' +
      '</div>' +
      '<div class="field"><label>Document</label>' +
        '<div class="seg">' +
          '<button type="button" id="editKeep_' + esc(ecnNo) + '" class="on" onclick="editDocMode(\\'' + esc(ecnNo) + '\\', \\'keep\\')">' + (g.docUrl ? 'Keep current' : 'None') + '</button>' +
          '<button type="button" id="editNew_' + esc(ecnNo) + '" onclick="editDocMode(\\'' + esc(ecnNo) + '\\', \\'new\\')">Upload new PDF</button>' +
          '<button type="button" id="editNone_' + esc(ecnNo) + '" onclick="editDocMode(\\'' + esc(ecnNo) + '\\', \\'none\\')">Remove</button>' +
        '</div>' +
        (g.docUrl ? '<div class="hint">Current: <a href="' + esc(g.docUrl) + '" target="_blank">open</a></div>' : '') +
        '<div class="uploadbox" id="editUpload_' + esc(ecnNo) + '" style="display:none"><input type="file" accept="application/pdf" onchange="onEditFile(\\'' + esc(ecnNo) + '\\')"><div class="hint" id="editFileInfo_' + esc(ecnNo) + '"></div></div>' +
      '</div>' +
      '<div style="display:flex;gap:10px"><button class="primary sm" onclick="saveEdit(\\'' + esc(ecnNo) + '\\')">Save changes</button><button class="sm" onclick="toggleEdit(\\'' + esc(ecnNo) + '\\')">Cancel</button><span class="sp"></span><span id="editStatus_' + esc(ecnNo) + '"></span></div>' +
      '</div></td>';
    row.parentNode.insertBefore(tr, row.nextSibling);
    window['editDoc_' + ecnNo] = { mode: 'keep', base64: '', mime: '', name: '' };
  }

  function confirmDeleteEcn(ecnNo) {
    var existing = document.getElementById('del_' + ecnNo);
    if (existing) { existing.remove(); return; }
    document.getElementById('edit_' + ecnNo) && toggleEdit(ecnNo);
    var g = ecnGroups[ecnNo];
    var row = document.getElementById('row_' + ecnNo);
    var tr = document.createElement('tr');
    tr.className = 'editRow';
    tr.id = 'del_' + ecnNo;
    tr.innerHTML = '<td colspan="10"><div class="editPanel">' +
      '<b>Delete ' + esc(ecnNo) + ' from ' + esc(g.model) + '?</b>' +
      '<div class="hint" style="margin:6px 0"><b>Reverse &amp; delete</b> undoes exactly what this ECN changed on the live BOM (an added part is removed again, a removed part is put back, a QPS change is put back to its old value), then deletes its rows from ECN history and rolls ' + esc(g.model) + '\\'s revision counter back by one - as if it never happened. This only works if no later ECN touched the same part(s) - if one did, delete that later ECN first (most recent first).</div>' +
      '<div style="display:flex;gap:10px;align-items:center;margin-bottom:10px"><button class="danger sm" onclick="doDeleteEcn(\\'' + esc(ecnNo) + '\\')">Reverse &amp; delete this ECN</button></div>' +
      '<div class="hint" style="margin:6px 0"><b>Delete record only</b> removes this ECN from history but leaves the live BOM exactly as it is right now - use this only when this ECN was a duplicate/accidental repeat whose changes are already reflected by another ECN (e.g. the same thing was entered twice). If this ECN\\'s changes are NOT duplicated elsewhere, this will leave the BOM looking changed with no ECN explaining why.</div>' +
      '<div style="display:flex;gap:10px;align-items:center">' +
        '<button class="danger sm" onclick="doDeleteEcnRecordOnly(\\'' + esc(ecnNo) + '\\')">Delete record only, keep BOM as-is</button>' +
        '<button class="sm" onclick="confirmDeleteEcn(\\'' + esc(ecnNo) + '\\')">Cancel</button><span class="sp"></span><span id="delStatus_' + esc(ecnNo) + '"></span>' +
      '</div>' +
      '</div></td>';
    row.parentNode.insertBefore(tr, row.nextSibling);
  }

  function doDeleteEcn(ecnNo) {
    var statusEl = document.getElementById('delStatus_' + ecnNo);
    statusEl.textContent = 'Deleting...'; statusEl.className = '';
    google.script.run.withSuccessHandler(function (r) {
      statusEl.textContent = 'Deleted - ' + r.model + ' is back to Rev ' + r.newRev + '.';
      statusEl.className = 'ok';
      loadLog();
    }).withFailureHandler(function (e) {
      statusEl.textContent = 'Error: ' + (e && e.message ? e.message : e); statusEl.className = 'err';
    }).deleteEcn(ecnNo);
  }

  function doDeleteEcnRecordOnly(ecnNo) {
    var statusEl = document.getElementById('delStatus_' + ecnNo);
    statusEl.textContent = 'Removing record...'; statusEl.className = '';
    google.script.run.withSuccessHandler(function (r) {
      statusEl.textContent = 'Record removed - ' + r.model + '\\'s history now ends at Rev ' + r.newRev + '. The live BOM was not changed.';
      statusEl.className = 'ok';
      loadLog();
    }).withFailureHandler(function (e) {
      statusEl.textContent = 'Error: ' + (e && e.message ? e.message : e); statusEl.className = 'err';
    }).deleteEcnRecordOnly(ecnNo);
  }

  function editDocMode(ecnNo, mode) {
    window['editDoc_' + ecnNo].mode = mode;
    document.getElementById('editKeep_' + ecnNo).className = mode === 'keep' ? 'on' : '';
    document.getElementById('editNew_' + ecnNo).className = mode === 'new' ? 'on' : '';
    document.getElementById('editNone_' + ecnNo).className = mode === 'none' ? 'on' : '';
    document.getElementById('editUpload_' + ecnNo).style.display = mode === 'new' ? 'block' : 'none';
  }

  function onEditFile(ecnNo) {
    var input = document.getElementById('editUpload_' + ecnNo).querySelector('input[type=file]');
    var f = input.files[0];
    var info = document.getElementById('editFileInfo_' + ecnNo);
    var store = window['editDoc_' + ecnNo];
    if (!f) return;
    if (f.type !== 'application/pdf') { info.textContent = 'Please choose a PDF file.'; info.style.color = '#d93025'; input.value = ''; return; }
    info.style.color = ''; info.textContent = 'Reading ' + f.name + '...';
    var reader = new FileReader();
    reader.onload = function (e) {
      var res = e.target.result;
      store.base64 = res.substring(res.indexOf(',') + 1);
      store.mime = f.type; store.name = f.name;
      info.textContent = 'Ready: ' + f.name;
    };
    reader.readAsDataURL(f);
  }

  function saveEdit(ecnNo) {
    var reason = document.getElementById('editReason_' + ecnNo).value.trim();
    var effectiveDate = document.getElementById('editDate_' + ecnNo).value;
    var store = window['editDoc_' + ecnNo] || { mode: 'keep' };
    var statusEl = document.getElementById('editStatus_' + ecnNo);
    if (!reason) { statusEl.textContent = 'Enter a reason.'; statusEl.className = 'err'; return; }
    if (!effectiveDate) { statusEl.textContent = 'Enter the effective date.'; statusEl.className = 'err'; return; }
    if (store.mode === 'new' && !store.base64) { statusEl.textContent = 'Upload the replacement PDF, or choose Keep/Remove.'; statusEl.className = 'err'; return; }

    var payload = {
      ecnNo: ecnNo, reason: reason, effectiveDate: effectiveDate,
      received: store.mode !== 'none',
      keepExistingDoc: store.mode === 'keep',
      docBase64: store.mode === 'new' ? store.base64 : '',
      docMime: store.mime, docName: store.name
    };
    statusEl.textContent = 'Saving...'; statusEl.className = '';
    google.script.run.withSuccessHandler(function (r) {
      statusEl.textContent = 'Saved.'; statusEl.className = 'ok';
      loadLog();
    }).withFailureHandler(function (e) {
      statusEl.textContent = 'Error: ' + (e && e.message ? e.message : e); statusEl.className = 'err';
    }).updateEcnMeta(payload);
  }

  function fail(e) { setStatus('Error: ' + (e && e.message ? e.message : e), true); }
  function setStatus(t, isErr) {
    var s = document.getElementById('status');
    s.textContent = t;
    s.className = 'sp ' + (isErr === true ? 'err' : (isErr === false ? 'ok' : ''));
  }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  init();
</script>
</body>
</html>
`;
}


/** Screen HTML (embedded): BOM version history (view / restore). */
function getBomHistoryHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
  --shadow-lift: 0 2px 6px rgba(60,64,67,.15), 0 4px 12px rgba(60,64,67,.1);
}
* { box-sizing: border-box; }
html, body {
  height: 100%; margin: 0;
  font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif;
  background: var(--canvas); color: var(--ink); font-size: 13px;
}
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar {
  flex: 0 0 auto; display: flex; align-items: center; gap: 14px;
  padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft);
  box-shadow: var(--shadow);
}
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.appbody { flex: 1 1 auto; overflow: auto; padding: 18px 22px 30px; display: flex; gap: 16px; }
.col-left { flex: 0 0 420px; display: flex; flex-direction: column; min-width: 0; }
.col-right { flex: 1 1 auto; min-width: 0; display: flex; flex-direction: column; }

.toolbar {
  display: flex; gap: 12px; align-items: center; flex-wrap: wrap;
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  padding: 12px 14px; margin-bottom: 14px; box-shadow: var(--shadow);
}
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar .grow { flex: 1; min-width: 160px; }

.card {
  background: var(--surface); border: 1px solid var(--line-soft); border-radius: var(--radius);
  box-shadow: var(--shadow); margin-bottom: 14px; overflow: hidden; display: flex; flex-direction: column; flex: 1 1 auto; min-height: 0;
}
.card .card-h { padding: 11px 16px; border-bottom: 1px solid var(--line-soft); font-weight: 600; font-size: 13.5px; display: flex; align-items: center; gap: 10px; flex: 0 0 auto; }
.card .card-b { padding: 14px 16px; overflow: auto; flex: 1 1 auto; }

select {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink);
}
textarea {
  padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm);
  font-family: inherit; font-size: 13px; width: 100%; min-height: 56px; resize: vertical;
}
select:focus, textarea:focus { outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg); }

button {
  padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink);
  border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500;
  transition: background .15s, box-shadow .15s;
}
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); box-shadow: var(--shadow-lift); }
button.danger { background: #fff; color: var(--red); border-color: var(--red); }
button.danger:hover { background: var(--red-bg); }
button.sm { padding: 5px 10px; font-size: 12px; }
button:disabled { opacity: .5; cursor: default; box-shadow: none; }

.pill {
  display: inline-flex; align-items: center; gap: 5px; padding: 3px 10px; border-radius: 999px;
  font-size: 11.5px; font-weight: 600; letter-spacing: .2px; white-space: nowrap;
}
.pill.ok { background: var(--green-bg); color: var(--green); }
.pill.bad { background: var(--red-bg); color: var(--red); }
.pill.warn { background: var(--amber-bg); color: var(--amber); }
.pill.info { background: var(--blue-bg); color: var(--blue-dark); }
.pill.gray { background: #f1f3f4; color: var(--sub); }

.tbl-wrap { border: 1px solid var(--line-soft); border-radius: var(--radius); overflow: auto; background: var(--surface); box-shadow: var(--shadow); flex: 1 1 auto; min-height: 0; }
table { border-collapse: collapse; width: 100%; }
th {
  position: sticky; top: 0; z-index: 2; background: #f8f9fb; text-align: left;
  padding: 10px 12px; font-size: 12px; font-weight: 600; color: var(--sub);
  text-transform: uppercase; letter-spacing: .3px; border-bottom: 1px solid var(--line-soft);
}
td { padding: 7px 12px; border-bottom: 1px solid var(--line-soft); vertical-align: middle; }
tbody tr:hover td { background: #f8f9fb; }
tr.revRow { cursor: pointer; }
tr.revRow.sel td { background: var(--blue-bg); }

img.thumb { width: 36px; height: 36px; object-fit: contain; border: 1px solid var(--line-soft); border-radius: var(--radius-sm); background: #fff; }
.banner { padding: 10px 14px; border-radius: var(--radius); margin-bottom: 14px; font-size: 12.5px; border: 1px solid transparent; flex: 0 0 auto; }
.banner.good { background: var(--green-bg); color: var(--green); border-color: #bfe4c8; }
.banner.warn { background: var(--amber-bg); color: var(--amber); border-color: #fde3a7; }
.banner.info { background: var(--blue-bg); color: var(--blue-dark); border-color: #c7dcfb; }
::-webkit-scrollbar { width: 10px; height: 10px; }
::-webkit-scrollbar-thumb { background: #c6cad0; border-radius: 6px; }
::-webkit-scrollbar-thumb:hover { background: #a9aeb5; }

.status.ok { color: var(--green); } .status.err { color: var(--red); }
.ok { color: var(--green); } .err { color: var(--red); }
.restoreBox { border: 1px solid var(--line-soft); border-radius: var(--radius); padding: 12px 14px; margin-top: 12px; background: #fafbfc; flex: 0 0 auto; }
.restoreBox.show { display: block; } .restoreBox.hide { display: none; }
.ctr { text-align: center; }
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>BOM version history</h1>
    <span class="sub">View any past revision of a model's BOM, or restore the live BOM back to it</span>
    <div class="fill"></div>
  </div>

  <div class="appbody">
    <div class="col-left">
      <div class="toolbar">
        <label class="grow">Model
          <select id="model" onchange="onModelChange()" style="width:100%"><option value="">Loading...</option></select>
        </label>
      </div>
      <div class="card">
        <div class="card-h">
          Revisions
          <span class="fill" style="flex:1"></span>
          <span class="hint" id="mergeHint" style="font-size:11px;color:var(--sub);margin-right:8px">Tick 2+ adjacent revisions to merge them into one</span>
          <button class="sm" id="mergeBtn" onclick="showMergeBox()" disabled>Merge selected</button>
        </div>
        <div class="tbl-wrap">
          <table>
            <thead><tr><th style="width:26px"></th><th>Rev</th><th>Date</th><th>ECN</th><th>Reason</th><th>Parts</th></tr></thead>
            <tbody id="revRows"><tr><td colspan="6">Select a model.</td></tr></tbody>
          </table>
        </div>
      </div>
      <div class="restoreBox hide" id="mergeBox">
        <b>Merge revisions <span id="mergeList"></span> of <span id="mergeModel"></span> into one?</b>
        <div style="margin:8px 0;color:var(--sub);font-size:12px">The ECNs inside these revisions keep their own records exactly as they happened - only the revision number they're grouped under changes, so they show as a single combined version afterwards. Later revisions shift down to stay numbered consecutively. The live BOM itself does not change. This cannot be undone.</div>
        <div style="margin-top:10px;display:flex;gap:10px;align-items:center">
          <button class="danger sm" onclick="doMerge()">Yes, merge these revisions</button>
          <button class="sm" onclick="hideMergeBox()">Cancel</button>
          <span id="mergeStatus"></span>
        </div>
      </div>
    </div>

    <div class="col-right">
      <div id="viewBanner" class="banner info" style="display:none"></div>
      <div class="card">
        <div class="card-h">
          <span id="viewTitle">Pick a revision on the left to view it</span>
          <span class="fill" style="flex:1"></span>
          <button class="sm" id="downloadBtn" onclick="downloadRev()" disabled>Download CSV</button>
          <button class="sm danger" id="deleteEcnBtn" onclick="showDeleteBox()" disabled title="Select a revision with a single ECN first">Delete this ECN</button>
          <button class="sm primary" id="restoreBtn" onclick="showRestoreBox()" disabled>Restore this version</button>
        </div>
        <div class="card-b" style="padding:0">
          <div class="tbl-wrap" style="border:none;box-shadow:none;border-radius:0">
            <table>
              <thead><tr><th style="width:46px">Img</th><th>Part No</th><th>Description</th><th class="ctr">QPS</th></tr></thead>
              <tbody id="partRows"><tr><td colspan="4">Nothing selected yet.</td></tr></tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="restoreBox hide" id="restoreBox">
        <b>Restore <span id="restoreModel"></span> to Rev <span id="restoreRev"></span>?</b>
        <div style="margin:8px 0;color:var(--sub);font-size:12px">This applies the difference as a brand-new ECN (auto-numbered), so the restore itself is fully logged. The revision number keeps counting forward - it does not go back to the old number.</div>
        <textarea id="restoreReason" placeholder="Reason (optional - defaults to 'Restored to Rev N')"></textarea>
        <div style="margin-top:10px;display:flex;gap:10px;align-items:center">
          <button class="primary sm" onclick="doRestore()">Confirm restore</button>
          <button class="sm" onclick="hideRestoreBox()">Cancel</button>
          <span id="restoreStatus"></span>
        </div>
      </div>
      <div class="restoreBox hide" id="deleteBox">
        <b>Delete <span id="deleteEcnNo"></span> from <span id="deleteModel"></span>?</b>
        <div style="margin:8px 0;color:var(--sub);font-size:12px">This reverses exactly what this ECN changed on the live BOM (an added part is removed again, a removed part is put back, a QPS change is put back to its old value), deletes its rows from ECN history, and rolls the revision counter back by one - as if it never happened. This cannot be undone.</div>
        <div style="margin-top:10px;display:flex;gap:10px;align-items:center">
          <button class="danger sm" onclick="doDeleteEcnHere()">Yes, delete this ECN</button>
          <button class="sm" onclick="hideDeleteBox()">Cancel</button>
          <span id="deleteStatus"></span>
        </div>
      </div>
    </div>
  </div>

<script>
  var models = [], revisions = [], selectedRev = null, curModel = '', checkedRevs = {};

  function init() {
    google.script.run.withSuccessHandler(function (d) { models = d; renderModels(); })
      .withFailureHandler(fail).getVerifyModels();
  }

  function renderModels() {
    document.getElementById('model').innerHTML = '<option value="">Select a model...</option>' +
      models.map(function (m) { return '<option value="' + esc(m.name) + '">' + esc(m.name) + (m.locked ? ' (Locked)' : '') + '</option>'; }).join('');
  }

  function onModelChange() {
    curModel = document.getElementById('model').value;
    selectedRev = null;
    checkedRevs = {};
    document.getElementById('partRows').innerHTML = '<tr><td colspan="4">Nothing selected yet.</td></tr>';
    document.getElementById('viewTitle').textContent = 'Pick a revision on the left to view it';
    document.getElementById('viewBanner').style.display = 'none';
    document.getElementById('downloadBtn').disabled = true;
    document.getElementById('restoreBtn').disabled = true;
    document.getElementById('deleteEcnBtn').disabled = true;
    document.getElementById('mergeBtn').disabled = true;
    hideRestoreBox(); hideDeleteBox(); hideMergeBox();
    if (!curModel) { document.getElementById('revRows').innerHTML = '<tr><td colspan="6">Select a model.</td></tr>'; return; }
    document.getElementById('revRows').innerHTML = '<tr><td colspan="6">Loading...</td></tr>';
    google.script.run.withSuccessHandler(function (rows) {
      revisions = rows;
      renderRevisions();
      if (rows.length) selectRev(rows[0].rev);
    }).withFailureHandler(fail).getBomRevisions(curModel);
  }

  function renderRevisions() {
    if (!revisions.length) { document.getElementById('revRows').innerHTML = '<tr><td colspan="6">No ECN history yet for this model - nothing to show besides the live BOM.</td></tr>'; return; }
    document.getElementById('revRows').innerHTML = revisions.map(function (r) {
      return '<tr class="revRow' + (selectedRev === r.rev ? ' sel' : '') + '" id="rev_' + r.rev + '" onclick="selectRev(' + r.rev + ')">' +
        '<td><input type="checkbox" class="revChk" data-rev="' + r.rev + '" onclick="event.stopPropagation()" onchange="onRevCheck()"' + (checkedRevs[r.rev] ? ' checked' : '') + '></td>' +
        '<td>' + r.rev + (r.isCurrent ? ' <span class="pill ok">current</span>' : '') + '</td>' +
        '<td>' + esc(r.when) + '</td><td>' + esc(r.ecnNo) + '</td><td>' + esc(r.reason) + '</td><td class="ctr">' + r.partsChanged + '</td></tr>';
    }).join('');
    onRevCheck();
  }

  function onRevCheck() {
    checkedRevs = {};
    document.querySelectorAll('.revChk:checked').forEach(function (el) { checkedRevs[Number(el.dataset.rev)] = true; });
    var list = Object.keys(checkedRevs).map(Number).sort(function (a, b) { return a - b; });
    var valid = list.length >= 2;
    if (valid) { for (var i = 1; i < list.length; i++) { if (list[i] !== list[i - 1] + 1) { valid = false; break; } } }
    document.getElementById('mergeBtn').disabled = !valid;
    document.getElementById('mergeHint').textContent = list.length === 0 ? 'Tick 2+ adjacent revisions to merge them into one'
      : (list.length === 1 ? 'Tick one more adjacent revision to merge'
        : (valid ? list.length + ' revisions selected - ready to merge' : 'Selected revisions must be adjacent (consecutive) to merge'));
  }

  function selectRev(rev) {
    selectedRev = rev;
    document.querySelectorAll('.revRow').forEach(function (el) { el.classList.remove('sel'); });
    var row = document.getElementById('rev_' + rev);
    if (row) row.classList.add('sel');
    hideRestoreBox(); hideDeleteBox();
    document.getElementById('partRows').innerHTML = '<tr><td colspan="4">Loading...</td></tr>';
    google.script.run.withSuccessHandler(function (snap) {
      var meta = revisions.filter(function (r) { return r.rev === rev; })[0];
      document.getElementById('viewTitle').textContent = curModel + ' - Rev ' + rev + (meta && meta.isCurrent ? ' (current)' : '');
      var banner = document.getElementById('viewBanner');
      banner.style.display = 'block';
      banner.textContent = snap.parts.length + ' part(s) in this revision' + (snap.asOfEcn ? ' - last change ' + snap.asOfEcn + ' on ' + snap.asOfWhen : '');
      document.getElementById('partRows').innerHTML = snap.parts.length ? snap.parts.map(function (p) {
        return '<tr><td>' + (p.img ? '<img class="thumb" src="' + esc(p.img) + '">' : '') + '</td><td>' + esc(p.partNo) + '</td><td>' + esc(p.desc) + '</td><td class="ctr">' + esc(p.qps) + '</td></tr>';
      }).join('') : '<tr><td colspan="4">This revision had no parts.</td></tr>';
      document.getElementById('downloadBtn').disabled = false;
      document.getElementById('restoreBtn').disabled = !!(meta && meta.isCurrent);
      var singleEcn = meta && meta.ecnNo && meta.ecnNo.indexOf(',') < 0;
      var delBtn = document.getElementById('deleteEcnBtn');
      delBtn.disabled = !(meta && singleEcn);
      delBtn.title = (meta && !singleEcn)
        ? 'This revision combines more than one ECN - delete one individually from the ECN history / edit tab in Generate ECN'
        : 'Reverses this ECN\\'s part changes and removes it from history (refused automatically if a later ECN touched the same part)';
    }).withFailureHandler(fail).getBomAtRevision(curModel, rev);
  }

  function downloadRev() {
    if (selectedRev === null || !curModel) return;
    google.script.run.withSuccessHandler(function (r) {
      var blob = new Blob([r.csv], { type: 'text/csv;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = r.filename;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
    }).withFailureHandler(fail).getBomRevisionCsv(curModel, selectedRev);
  }

  function showMergeBox() {
    var list = Object.keys(checkedRevs).map(Number).sort(function (a, b) { return a - b; });
    if (list.length < 2) return;
    document.getElementById('mergeModel').textContent = curModel;
    document.getElementById('mergeList').textContent = list.join(', ');
    document.getElementById('mergeStatus').textContent = '';
    document.getElementById('mergeBox').className = 'restoreBox show';
  }
  function hideMergeBox() { document.getElementById('mergeBox').className = 'restoreBox hide'; }

  function doMerge() {
    var list = Object.keys(checkedRevs).map(Number).sort(function (a, b) { return a - b; });
    var statusEl = document.getElementById('mergeStatus');
    statusEl.textContent = 'Merging...'; statusEl.className = '';
    google.script.run.withSuccessHandler(function (r) {
      statusEl.textContent = 'Merged into Rev ' + r.mergedInto + ' under ' + r.targetEcnNo +
        (r.ecnsMerged > 1 ? ' (' + r.ecnsMerged + ' ECNs combined into this one)' : '') + '.';
      statusEl.className = 'ok';
      checkedRevs = {};
      hideMergeBox();
      reloadRevisionsKeepingModel();
    }).withFailureHandler(function (e) {
      statusEl.textContent = 'Error: ' + (e && e.message ? e.message : e); statusEl.className = 'err';
    }).mergeBomRevisions(curModel, list);
  }

  function showDeleteBox() {
    if (selectedRev === null) return;
    var meta = revisions.filter(function (r) { return r.rev === selectedRev; })[0];
    if (!meta || meta.ecnNo.indexOf(',') >= 0) return;
    document.getElementById('deleteEcnNo').textContent = meta.ecnNo;
    document.getElementById('deleteModel').textContent = curModel;
    document.getElementById('deleteStatus').textContent = '';
    document.getElementById('deleteBox').className = 'restoreBox show';
  }
  function hideDeleteBox() { document.getElementById('deleteBox').className = 'restoreBox hide'; }

  function doDeleteEcnHere() {
    var meta = revisions.filter(function (r) { return r.rev === selectedRev; })[0];
    if (!meta) return;
    var statusEl = document.getElementById('deleteStatus');
    statusEl.textContent = 'Deleting...'; statusEl.className = '';
    google.script.run.withSuccessHandler(function (r) {
      statusEl.textContent = 'Deleted - ' + esc(r.model) + ' is back to Rev ' + r.newRev + '.'; statusEl.className = 'ok';
      hideDeleteBox();
      reloadRevisionsKeepingModel();
    }).withFailureHandler(function (e) {
      statusEl.textContent = 'Error: ' + (e && e.message ? e.message : e); statusEl.className = 'err';
    }).deleteEcn(meta.ecnNo);
  }

  function reloadRevisionsKeepingModel() {
    google.script.run.withSuccessHandler(function (rows) {
      revisions = rows;
      checkedRevs = {};
      renderRevisions();
      if (rows.length) { selectRev(rows[0].rev); }
      else {
        selectedRev = null;
        document.getElementById('partRows').innerHTML = '<tr><td colspan="4">Nothing selected yet.</td></tr>';
        document.getElementById('viewTitle').textContent = 'Pick a revision on the left to view it';
        document.getElementById('viewBanner').style.display = 'none';
        document.getElementById('downloadBtn').disabled = true;
        document.getElementById('restoreBtn').disabled = true;
        document.getElementById('deleteEcnBtn').disabled = true;
      }
    }).withFailureHandler(fail).getBomRevisions(curModel);
  }

  function showRestoreBox() {
    if (selectedRev === null) return;
    document.getElementById('restoreModel').textContent = curModel;
    document.getElementById('restoreRev').textContent = selectedRev;
    document.getElementById('restoreReason').value = '';
    document.getElementById('restoreStatus').textContent = '';
    document.getElementById('restoreBox').className = 'restoreBox show';
  }
  function hideRestoreBox() { document.getElementById('restoreBox').className = 'restoreBox hide'; }

  function doRestore() {
    var reason = document.getElementById('restoreReason').value.trim();
    var statusEl = document.getElementById('restoreStatus');
    statusEl.textContent = 'Restoring...'; statusEl.className = '';
    google.script.run.withSuccessHandler(function (r) {
      if (r.note) { statusEl.textContent = r.note; statusEl.className = 'ok'; return; }
      statusEl.textContent = 'Restored - ' + r.ecnNo + ', now Rev ' + r.rev + ' (' + r.changes + ' part change(s)).';
      statusEl.className = 'ok';
      hideRestoreBox();
      reloadRevisionsKeepingModel();
    }).withFailureHandler(function (e) {
      statusEl.textContent = 'Error: ' + (e && e.message ? e.message : e); statusEl.className = 'err';
    }).restoreBomVersion(curModel, selectedRev, reason);
  }

  function fail(e) { console.error(e); var b = document.getElementById('viewBanner'); b.style.display = 'block'; b.className = 'banner warn'; b.textContent = 'Error: ' + (e && e.message ? e.message : e); }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  init();
</script>
</body>
</html>
`;
}

/** Screen HTML (embedded): formatted, printable ECN report. */
function getEcnReportHtml_() {
  return `<!DOCTYPE html>
<html>
<head>
<base target="_top">
<style>
:root {
  --blue: #1a73e8; --blue-dark: #0b57cf; --blue-bg: #e8f0fe;
  --green: #188038; --green-bg: #e6f4ea;
  --red: #d93025; --red-bg: #fce8e6;
  --amber: #b06000; --amber-bg: #fef7e0;
  --ink: #202124; --sub: #5f6368; --line: #dadce0; --line-soft: #eceff1;
  --surface: #ffffff; --canvas: #f6f8fb;
  --radius: 10px; --radius-sm: 6px;
  --shadow: 0 1px 2px rgba(60,64,67,.12), 0 1px 3px rgba(60,64,67,.08);
}
* { box-sizing: border-box; }
html, body { height: 100%; margin: 0; font-family: 'Google Sans', 'Segoe UI', Roboto, Arial, sans-serif; background: var(--canvas); color: var(--ink); font-size: 13px; }
body.app { display: flex; flex-direction: column; height: 100vh; overflow: hidden; }

.appbar { flex: 0 0 auto; display: flex; align-items: center; gap: 14px; padding: 14px 22px; background: var(--surface); border-bottom: 1px solid var(--line-soft); box-shadow: var(--shadow); }
.appbar h1 { font-size: 17px; font-weight: 600; margin: 0; white-space: nowrap; }
.appbar .sub { color: var(--sub); font-size: 12px; }
.appbar .fill { flex: 1; }

.toolbar { flex: 0 0 auto; display: flex; gap: 12px; align-items: center; flex-wrap: wrap; background: var(--surface); border-bottom: 1px solid var(--line-soft); padding: 12px 22px; }
.toolbar label { display: flex; gap: 7px; align-items: center; font-weight: 500; color: var(--ink); white-space: nowrap; }
.toolbar input, .toolbar select { padding: 8px 10px; border: 1px solid var(--line); border-radius: var(--radius-sm); font-family: inherit; font-size: 13px; background: var(--surface); color: var(--ink); }
.toolbar input:focus, .toolbar select:focus { outline: none; border-color: var(--blue); box-shadow: 0 0 0 3px var(--blue-bg); }

button { padding: 8px 14px; border: 1px solid var(--line); background: var(--surface); color: var(--ink); border-radius: var(--radius-sm); cursor: pointer; font-family: inherit; font-size: 13px; font-weight: 500; }
button:hover { background: #f1f3f4; }
button.primary { background: var(--blue); color: #fff; border-color: var(--blue); }
button.primary:hover { background: var(--blue-dark); }
button.sm { padding: 5px 10px; font-size: 12px; }

.reportArea { flex: 1 1 auto; overflow: auto; padding: 22px; }
.reportHead { text-align: center; margin-bottom: 20px; }
.reportHead .co { font-size: 15px; font-weight: 700; letter-spacing: .3px; }
.reportHead .ti { font-size: 19px; font-weight: 700; margin-top: 2px; }
.reportHead .ge { color: var(--sub); font-size: 11.5px; margin-top: 4px; }

.ecnCard { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); box-shadow: var(--shadow); margin-bottom: 18px; overflow: hidden; }
.ecnCard .hd { display: flex; align-items: center; gap: 14px; padding: 12px 18px; background: #f8f9fb; border-bottom: 1px solid var(--line-soft); }
.ecnCard .hd .no { font-size: 15px; font-weight: 700; color: var(--blue-dark); }
.ecnCard .hd .fill { flex: 1; }
.ecnCard .hd .rev { font-size: 12px; font-weight: 600; background: var(--blue-bg); color: var(--blue-dark); padding: 4px 10px; border-radius: 999px; white-space: nowrap; }
.ecnCard .meta { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px 18px; padding: 14px 18px; border-bottom: 1px solid var(--line-soft); }
.ecnCard .meta .f label { display: block; font-size: 10.5px; text-transform: uppercase; letter-spacing: .3px; color: var(--sub); font-weight: 600; margin-bottom: 2px; }
.ecnCard .meta .f div { font-size: 13px; }
.ecnCard .reason { padding: 12px 18px; border-bottom: 1px solid var(--line-soft); }
.ecnCard .reason label { display: block; font-size: 10.5px; text-transform: uppercase; letter-spacing: .3px; color: var(--sub); font-weight: 600; margin-bottom: 3px; }
.ecnCard .reason div { font-size: 13px; white-space: pre-wrap; }
.ecnCard .parts { display: grid; grid-template-columns: 1fr 1fr; }
.ecnCard .parts .col { padding: 12px 18px; }
.ecnCard .parts .col:first-child { border-right: 1px solid var(--line-soft); }
.ecnCard .parts .colTitle { font-size: 10.5px; text-transform: uppercase; letter-spacing: .3px; font-weight: 700; margin-bottom: 8px; }
.ecnCard .parts .colTitle.before { color: var(--red); }
.ecnCard .parts .colTitle.after { color: var(--green); }
.partLine { display: flex; justify-content: space-between; align-items: baseline; gap: 10px; padding: 5px 0; border-bottom: 1px dashed var(--line-soft); font-size: 12.5px; }
.partLine:last-child { border-bottom: none; }
.partLine .pnWrap { min-width: 0; }
.partLine .pn { font-weight: 600; }
.partLine .pd { color: var(--sub); }
.partLine .qty { color: var(--sub); white-space: nowrap; flex: 0 0 auto; }
.noneRow { color: var(--sub); font-size: 12px; font-style: italic; }
.ecnCard .ft { display: flex; justify-content: space-between; gap: 10px; padding: 9px 18px; background: #fafbfc; font-size: 11.5px; color: var(--sub); }
.ecnCard .ft a { color: var(--blue-dark); }
.empty { text-align: center; color: var(--sub); padding: 60px 20px; font-size: 13px; }

@media print {
  body.app { height: auto; overflow: visible; }
  .toolbar, .appbar .fill, .no-print { display: none !important; }
  .appbar { box-shadow: none; border-bottom: 2px solid #000; }
  .reportArea { overflow: visible; padding: 0; }
  .ecnCard { box-shadow: none; border: 1px solid #999; page-break-inside: avoid; }
}
</style>
</head>
<body class="app">
  <div class="appbar">
    <h1>ECN report</h1>
    <span class="sub">A formatted, printable record of ECNs - before/after parts with descriptions, reason, effective date and BOM revision. Pick one ECN below to print just that one.</span>
    <div class="fill"></div>
    <button class="primary no-print" onclick="window.print()">Print / Save as PDF</button>
  </div>

  <div class="toolbar">
    <label>Company / site
      <input id="coName" type="text" placeholder="(shown on the printed report)" style="width:220px">
    </label>
    <label>Model
      <select id="model" onchange="onModelChange()" style="width:180px"><option value="">All models</option></select>
    </label>
    <label class="grow">ECN
      <select id="ecnPick" onchange="load()" style="width:180px"><option value="">All ECNs</option></select>
    </label>
    <label>Effective from <input id="fromDate" type="date" onchange="render()"></label>
    <label>Effective to <input id="toDate" type="date" onchange="render()"></label>
    <button onclick="load()">Refresh</button>
    <span class="fill" style="flex:1"></span>
    <span id="count" style="color:var(--sub);font-size:12px"></span>
  </div>

  <div class="reportArea">
    <div class="reportHead">
      <div class="co" id="coHead"></div>
      <div class="ti">Engineering Change Notice (ECN) Report</div>
      <div class="ge" id="genLine"></div>
    </div>
    <div id="cards"><div class="empty">Loading...</div></div>
  </div>

<script>
  var allRows = [];

  function init() {
    document.getElementById('coName').oninput = function () { document.getElementById('coHead').textContent = this.value; };
    google.script.run.withSuccessHandler(function (d) {
      document.getElementById('model').innerHTML = '<option value="">All models</option>' +
        d.map(function (m) { return '<option value="' + esc(m.name) + '">' + esc(m.name) + '</option>'; }).join('');
      loadEcnList();
    }).withFailureHandler(fail).getVerifyModels();
  }

  function onModelChange() {
    document.getElementById('ecnPick').value = '';
    loadEcnList();
  }

  function loadEcnList() {
    var model = document.getElementById('model').value;
    google.script.run.withSuccessHandler(function (list) {
      var keep = document.getElementById('ecnPick').value;
      document.getElementById('ecnPick').innerHTML = '<option value="">All ECNs</option>' +
        list.map(function (e) { return '<option value="' + esc(e.ecnNo) + '">' + esc(e.ecnNo) + (model ? '' : ' - ' + esc(e.model)) + '</option>'; }).join('');
      document.getElementById('ecnPick').value = keep;
      load();
    }).withFailureHandler(fail).getAllEcnNos(model);
  }

  function load() {
    var model = document.getElementById('model').value;
    var ecnNo = document.getElementById('ecnPick').value;
    document.getElementById('cards').innerHTML = '<div class="empty">Loading...</div>';
    google.script.run.withSuccessHandler(function (rows) {
      allRows = rows || [];
      render();
    }).withFailureHandler(fail).getEcnReportData(model, ecnNo);
  }

  function render() {
    var ecnNo = document.getElementById('ecnPick').value;
    var from = document.getElementById('fromDate').value;
    var to = document.getElementById('toDate').value;
    var rows = allRows.filter(function (g) {
      if (ecnNo) return true; // a single ECN was explicitly picked - the date range doesn't apply
      if (from && g.effectiveDate && g.effectiveDate < from) return false;
      if (to && g.effectiveDate && g.effectiveDate > to) return false;
      return true;
    });
    document.getElementById('count').textContent = rows.length + ' ECN(s)';
    document.getElementById('genLine').textContent = 'Generated ' + new Date().toLocaleString();
    if (!rows.length) { document.getElementById('cards').innerHTML = '<div class="empty">No ECNs match this filter.</div>'; return; }
    document.getElementById('cards').innerHTML = rows.map(renderCard).join('');
  }

  function renderCard(g) {
    var beforeHtml = g.before.length ? g.before.map(partLine).join('') : '<div class="noneRow">No parts removed or changed from</div>';
    var afterHtml = g.after.length ? g.after.map(partLine).join('') : '<div class="noneRow">No parts added or changed to</div>';
    return '<div class="ecnCard">' +
      '<div class="hd"><span class="no">' + esc(g.ecnNo) + '</span><span class="fill"></span><span class="rev">' + esc(g.model) + ' - BOM Rev ' + esc(g.rev) + '</span></div>' +
      '<div class="meta">' +
        '<div class="f"><label>Model</label><div>' + esc(g.model) + '</div></div>' +
        '<div class="f"><label>Date raised</label><div>' + esc(g.when) + '</div></div>' +
        '<div class="f"><label>Effective date</label><div>' + esc(g.effectiveDate || '-') + '</div></div>' +
        '<div class="f"><label>Changed by</label><div>' + esc(g.by) + '</div></div>' +
      '</div>' +
      '<div class="reason"><label>Reason for change</label><div>' + esc(g.reason || '-') + '</div></div>' +
      '<div class="parts">' +
        '<div class="col"><div class="colTitle before">Before (old parts)</div>' + beforeHtml + '</div>' +
        '<div class="col"><div class="colTitle after">After (new parts)</div>' + afterHtml + '</div>' +
      '</div>' +
      '<div class="ft"><span>BOM revision produced: Rev ' + esc(g.rev) + '</span><span>' + (g.docUrl ? '<a href="' + esc(g.docUrl) + '" target="_blank">Open attached ECN document</a>' : 'No document attached') + '</span></div>' +
      '</div>';
  }

  function partLine(p) {
    return '<div class="partLine"><div class="pnWrap"><span class="pn">' + esc(p.partNo) + '</span>' +
      (p.desc ? '<span class="pd"> - ' + esc(p.desc) + '</span>' : '') + '</div>' +
      '<span class="qty">QPS ' + esc(p.qps) + '</span></div>';
  }

  function fail(e) { document.getElementById('cards').innerHTML = '<div class="empty">Error: ' + esc(e && e.message ? e.message : e) + '</div>'; }
  function esc(v) {
    return String(v === null || v === undefined ? '' : v).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  init();
</script>
</body>
</html>
`;
}

// ======================================================================
//   WEB ACCESS + WEB VERSIONS OF THE MENU-ONLY TOOLS
//   (used by LUMIO Quality Management - Projector > BOM management)
// ======================================================================

/** Looks up a portal login token (only its SHA-256 hash is stored). Returns {username,name,role} or null. */
function portalSession_(token) {
  token = String(token || '').trim();
  if (!token) return null;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  try {
    const sh = SpreadsheetApp.openById(PORTAL_SS_ID).getSheetByName('Sessions');
    if (!sh || sh.getLastRow() < 2) return null;
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
    for (let i = v.length - 1; i >= 0; i--) {
      if (String(v[i][0]) !== hash) continue;
      if (!(v[i][5] instanceof Date) || v[i][5] <= new Date()) return null;
      const role = String(v[i][3]).toLowerCase();
      const acc = String(v[i][6] || '').toLowerCase().split(/[\s,;]+/);
      return { username: String(v[i][1]), name: String(v[i][2]), role: role,
               bomManagement: role === 'admin' || acc.indexOf('bommgmt') > -1 };
    }
  } catch (e) { Logger.log('Session check failed: ' + e.message); }
  return null;
}

/** True when the person opening the page is the script owner (direct use outside the portal). */
function ownerVisitor_() {
  try {
    const a = Session.getActiveUser().getEmail();
    return !!a && a === Session.getEffectiveUser().getEmail();
  } catch (e) { return false; }
}

/** Shared look of the Part DB screens (taken from the existing screens' stylesheet). */
function pdbCss_() {
  const s = getSidebarHtml_();
  const a = s.indexOf('<style>') + 7, b = s.indexOf('    body { padding: 14px; }');
  return s.substring(a, b > a ? b : s.indexOf('</style>'));
}

/** Where is one part used? Same shape as getSelectedPart(), but by part number (web version of the sidebar). */
function getPartUsage(partNo) {
  partNo = String(partNo || '').trim();
  if (!partNo) return { ok: false, msg: 'Enter a part number.' };
  const ctx = readMaster_();
  const i = ctx.idx[norm_(partNo)];
  if (i === undefined) return { ok: false, msg: partNo + ' is not in ' + SHEET_NAME + '.' };
  const real = String(ctx.mVals[i][COL_PART_NO - 1]).trim();
  const ex = masterExtra_(ctx, i);
  const ctrl = controlMap_();
  const models = getModelSheets_().map(s => {
    const f = findInModel_(s, real, ctx.mh);
    const c = ctrl[norm_(s.getName())];
    return { name: s.getName(), listed: !!f, qps: f ? f.qps : '', locked: !!(c && c.status === 'Locked') };
  });
  return { ok: true, partNo: real, desc: ex.desc, img: ex.img, models: models };
}

/** Rebuilds BOM_COMPARE and returns it for the web screen. */
function webCompareAll() {
  buildComparison();
  const sh = dataSS_().getSheetByName(COMPARE_SHEET);
  const v = sh.getDataRange().getValues();
  const summary = String(v[0][0] || '');
  const header = (v[2] || []).map(String);
  const keep = header.map((h, i) => h === 'Image' ? -1 : i).filter(i => i >= 0);
  const rows = v.slice(3).filter(r => String(r[0]).trim()).map(r => keep.map(i => txt_(r[i])));
  return { summary: summary, header: keep.map(i => header[i]), rows: rows, sheetUrl: dataSS_().getUrl() + '#gid=' + sh.getSheetId() };
}

function webLinkImages() { return linkPartImages(); }

function webMissingImages() {
  const sheet = dataSS_().getSheetByName(SHEET_NAME);
  const numRows = sheet.getLastRow() - HEADER_ROW;
  if (numRows <= 0) return { total: 0, missing: [] };
  const partNos = sheet.getRange(HEADER_ROW + 1, COL_PART_NO, numRows, 1).getValues();
  const index = buildImageIndex_();
  const missing = partNos.map(r => String(r[0]).trim()).filter(p => p && !index[p.toUpperCase()]);
  return { total: partNos.filter(r => String(r[0]).trim()).length, missing: missing };
}

function webDedupe() { return withLock_(() => dedupePartMaster_()); }

function webClearImages() {
  const sheet = dataSS_().getSheetByName(SHEET_NAME);
  const numRows = Math.max(sheet.getLastRow() - HEADER_ROW, 0);
  clearImageColumn();
  return { cleared: numRows };
}

/* ---------- Assign part to models (web) ---------- */
function getAssignWebHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width, initial-scale=1"><style>${pdbCss_()}
  .wrap{max-width:860px;margin:0 auto}
  .part{display:flex;gap:14px;align-items:center;background:var(--surface);border:1px solid var(--line-soft);border-radius:var(--radius);padding:12px 14px;margin-bottom:14px;box-shadow:var(--shadow)}
  .part img{width:64px;height:64px;object-fit:contain;border:1px solid var(--line-soft);border-radius:var(--radius-sm);background:#fff}
  .pn{font-weight:700;font-size:15px}.desc{color:var(--sub);margin-top:2px}
  .row{display:flex;align-items:center;gap:10px;padding:10px 16px;border-bottom:1px solid var(--line-soft)}
  .row:last-child{border-bottom:none}.row label{flex:1;cursor:pointer}
  .ecn input{width:100%;margin-top:6px}
  #status{white-space:pre-line}
  </style></head><body class="app">
  <div class="appbar"><h1>Assign part to models</h1><span class="sub">Add a part to model BOMs, change its QPS, or remove it</span></div>
  <div class="appbody"><div class="wrap">
    <div class="toolbar"><label for="pn">Part No</label><input type="text" id="pn" class="grow" list="parts" placeholder="Type or pick a part number from PART_MASTER" autocomplete="off"><datalist id="parts"></datalist><button class="primary" id="load">Show models</button></div>
    <div id="content"><div class="banner info">Pick a part to see which models use it.</div></div>
  </div></div>
  <div class="footbar"><span id="status"></span><span class="sp"></span><button class="primary" id="go" disabled>Save to model sheets</button></div>
<script>
  var cur=null, armed=false, busy=false;
  function esc(v){return String(v===null||v===undefined?'':v).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function setStatus(t,err){var s=document.getElementById('status');s.textContent=t||'';s.className=err===true?'err':(err===false?'ok':'');}
  google.script.run.withSuccessHandler(function(d){document.getElementById('parts').innerHTML=d.parts.map(function(p){return '<option value="'+esc(p.partNo)+'">'+esc(p.desc)+'</option>';}).join('');}).getBomData();
  function load(){var pn=document.getElementById('pn').value.trim(); if(!pn) return setStatus('Enter a part number.',true); setStatus('Loading...'); armed=false;
    google.script.run.withSuccessHandler(function(r){ if(!r.ok){cur=null;document.getElementById('go').disabled=true;document.getElementById('content').innerHTML='<div class="banner warn">'+esc(r.msg)+'</div>';return setStatus('');}
      cur=r; render(r); setStatus(''); }).withFailureHandler(function(e){setStatus(e.message,true);}).getPartUsage(pn);}
  function render(r){var anyLocked=false;
    var h='<div class="part">'+(r.img?'<img src="'+esc(r.img)+'" referrerpolicy="no-referrer" onerror="this.remove()">':'')+'<div><div class="pn">'+esc(r.partNo)+'</div><div class="desc">'+esc(r.desc)+'</div></div></div>';
    h+='<div class="toolbar">All models QPS <input type="number" min="0" step="any" id="bulkQps"><button onclick="applyAll()">Apply to all</button></div>';
    h+='<div class="card"><div class="card-h">Models</div>';
    r.models.forEach(function(m,i){ if(m.locked) anyLocked=true;
      h+='<div class="row"><input type="checkbox" id="c'+i+'"'+(m.listed?' checked':'')+' onchange="armed=false"><label for="c'+i+'">'+esc(m.name)+' '+(m.locked?'<span class="pill warn">Locked</span> ':'')+(m.listed?'<span class="pill ok">Listed</span>':'')+'</label>'+
         '<input type="number" min="0" step="any" id="q'+i+'" placeholder="QPS" value="'+esc(m.qps)+'" oninput="armed=false"></div>'; });
    h+='</div>';
    if(anyLocked) h+='<div class="banner warn ecn"><b>Locked BOM: an ECN is required</b><br>Any add, QPS change or removal on a locked model needs an ECN number and reason.<input type="text" id="ecnNo" placeholder="ECN No"><input type="text" id="ecnDesc" placeholder="Change description / reason"></div>';
    document.getElementById('content').innerHTML=h; document.getElementById('go').disabled=false;}
  function applyAll(){var v=document.getElementById('bulkQps').value; if(v==='')return; cur.models.forEach(function(m,i){document.getElementById('c'+i).checked=true;document.getElementById('q'+i).value=v;}); armed=false;}
  function submit(){ if(!cur||busy) return; var add=[],remove=[];
    for(var i=0;i<cur.models.length;i++){var ck=document.getElementById('c'+i).checked, q=document.getElementById('q'+i).value;
      if(ck){ if(!(Number(q)>0)) return setStatus('Enter a QPS for '+cur.models[i].name,true); add.push({model:cur.models[i].name,qps:Number(q)}); }
      else if(cur.models[i].listed) remove.push(cur.models[i].name); }
    if(!add.length&&!remove.length) return setStatus('Nothing to save.',true);
    if(remove.length&&!armed){armed=true;return setStatus('Part will be removed from: '+remove.join(', ')+'. Press Save again to confirm.',true);}
    armed=false; var n=document.getElementById('ecnNo'), d=document.getElementById('ecnDesc');
    busy=true; document.getElementById('go').disabled=true; setStatus('Saving...');
    google.script.run.withSuccessHandler(function(r){busy=false; setStatus(r.done.join('\\n'),false); load();})
      .withFailureHandler(function(e){busy=false;document.getElementById('go').disabled=false;setStatus(e.message,true);})
      .assignPart({partNo:cur.partNo,add:add,remove:remove,ecn:{no:n?n.value:'',desc:d?d.value:''}}); }
  document.getElementById('load').onclick=load;
  document.getElementById('pn').onkeydown=function(e){if(e.key==='Enter')load();};
  document.getElementById('go').onclick=submit;
</script></body></html>`;
}

/* ---------- Compare model BOMs, all models (web) ---------- */
function getCompareAllHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width, initial-scale=1"><style>${pdbCss_()}
  td.c{text-align:center} tr.common td{background:#f3fbf5} tr.shared td{background:#fffbf0} tr .diff{background:var(--red-bg)!important;color:var(--red);font-weight:600}
  </style></head><body class="app">
  <div class="appbar"><h1>Compare model BOMs</h1><span class="sub">All models side by side, from the model sheets</span><span class="fill"></span><a id="sheetLink" target="_blank" style="display:none">Open BOM_COMPARE sheet</a><button class="primary" id="run">Rebuild comparison</button></div>
  <div class="appbody">
    <div id="summary" class="banner info">Press "Rebuild comparison" to compare every model BOM. The result is also written to the BOM_COMPARE sheet.</div>
    <div class="toolbar"><label>Show <select id="f"><option value="">All parts</option><option value="Common">Common to all models</option><option value="Shared">Shared by some</option><option value="Only">Model-specific</option><option value="Different">Different QPS</option></select></label><input type="text" id="q" class="grow" placeholder="Search part number or description"><span id="count" class="sub"></span></div>
    <div class="tbl-wrap"><table id="tbl"></table></div>
  </div>
<script>
  var data=null;
  function esc(v){return String(v===null||v===undefined?'':v).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function run(){var b=document.getElementById('run');b.disabled=true;b.textContent='Comparing...';
    google.script.run.withSuccessHandler(function(d){data=d;b.disabled=false;b.textContent='Rebuild comparison';document.getElementById('summary').textContent=d.summary;var l=document.getElementById('sheetLink');l.href=d.sheetUrl;l.style.display='';draw();})
      .withFailureHandler(function(e){b.disabled=false;b.textContent='Rebuild comparison';document.getElementById('summary').className='banner warn';document.getElementById('summary').textContent=e.message;}).webCompareAll();}
  function draw(){ if(!data) return; var f=document.getElementById('f').value, q=document.getElementById('q').value.trim().toLowerCase(), H=data.header;
    var iS=H.indexOf('Status'), iC=H.indexOf('QPS check'), iD=H.indexOf('Description');
    var rows=data.rows.filter(function(r){ if(f==='Different'&&r[iC]!=='Different') return false; if(f&&f!=='Different'&&r[iS].indexOf(f)!==0) return false;
      return !q||String(r[0]).toLowerCase().indexOf(q)>-1||String(r[iD]).toLowerCase().indexOf(q)>-1;});
    document.getElementById('count').textContent=rows.length+' of '+data.rows.length+' parts';
    var h='<thead><tr>'+H.map(function(x){return '<th>'+esc(x)+'</th>';}).join('')+'</tr></thead><tbody>';
    rows.forEach(function(r){ var cls=r[iS].indexOf('Common')===0?'common':(r[iS].indexOf('Shared')===0?'shared':'');
      h+='<tr class="'+cls+'">'+r.map(function(v,i){ return '<td class="'+(i>iD&&i<iS-1||i===iC?'c':'')+(i===iC&&v==='Different'?' diff':'')+'">'+esc(v)+'</td>';}).join('')+'</tr>';});
    document.getElementById('tbl').innerHTML=h+'</tbody>';}
  document.getElementById('run').onclick=run; document.getElementById('f').onchange=draw; document.getElementById('q').oninput=draw;
</script></body></html>`;
}

/* ---------- Images and PART_MASTER clean-up (web) ---------- */
function getImagesToolsHtml_() {
  return `<!DOCTYPE html><html><head><base target="_top"><meta name="viewport" content="width=device-width, initial-scale=1"><style>${pdbCss_()}
  .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:14px}
  .card .b{padding:14px 16px}.card p{margin:0 0 12px;color:var(--sub);line-height:1.5}
  .out{margin-top:12px;white-space:pre-line;max-height:260px;overflow:auto}
  </style></head><body class="app">
  <div class="appbar"><h1>Part images and clean-up</h1><span class="sub">PART_MASTER image links and duplicate parts</span></div>
  <div class="appbody"><div class="grid">
    <div class="card"><div class="card-h">Link part images</div><div class="b"><p>Matches files in the part image folder (named &lt;PART_NO&gt;.jpg) to PART_MASTER and writes the image and image URL columns.</p><button class="primary" data-act="link">Link part images</button><div class="out" id="o_link"></div></div></div>
    <div class="card"><div class="card-h">Report missing images</div><div class="b"><p>Lists the parts in PART_MASTER that have no image file in the folder.</p><button data-act="missing">Report missing images</button><div class="out" id="o_missing"></div></div></div>
    <div class="card"><div class="card-h">Remove duplicate parts</div><div class="b"><p>Merges rows in PART_MASTER that have the same part number. The first row is kept and its blank cells are filled from the duplicates.</p><button data-act="dedupe">Remove duplicate parts</button><div class="out" id="o_dedupe"></div></div></div>
    <div class="card"><div class="card-h">Clear image column</div><div class="b"><p>Empties the image and image URL columns of PART_MASTER, before a full re-link.</p><button class="danger" data-act="clear">Clear image column</button><div class="out" id="o_clear"></div></div></div>
  </div></div>
<script>
  function esc(v){return String(v===null||v===undefined?'':v).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  var fn={link:'webLinkImages',missing:'webMissingImages',dedupe:'webDedupe',clear:'webClearImages'};
  var ask={dedupe:'Merge and remove duplicate part rows in PART_MASTER?',clear:'Clear the image and image URL columns of PART_MASTER?'};
  var show={
    link:function(r){return r.matched+' of '+r.total+' parts linked to an image.';},
    missing:function(r){return r.missing.length?r.missing.length+' of '+r.total+' parts have no image:\\n'+r.missing.join('\\n'):'All '+r.total+' parts have an image.';},
    dedupe:function(r){return r.removed?r.removed+' duplicate row(s) removed:\\n'+r.parts.join(', '):'No duplicate part numbers found.';},
    clear:function(r){return 'Image columns cleared for '+r.cleared+' rows.';}};
  document.querySelectorAll('button[data-act]').forEach(function(b){ b.onclick=function(){ var a=b.getAttribute('data-act'); if(ask[a]&&!confirm(ask[a])) return;
    var o=document.getElementById('o_'+a); b.disabled=true; o.className='out'; o.textContent='Working...';
    google.script.run.withSuccessHandler(function(r){b.disabled=false;o.className='out ok';o.textContent=show[a](r);})
      .withFailureHandler(function(e){b.disabled=false;o.className='out err';o.textContent=e.message;})[fn[a]](); }; });
</script></body></html>`;
}
