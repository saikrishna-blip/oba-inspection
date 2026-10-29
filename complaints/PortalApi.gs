

/* ========================================================================= *
 *  11. PORTAL API  (LUMIO Quality Management - Projector)
 *  The portal's Complaints screens call this web app with the portal login token.
 *  Only people with the "complaints" module (or admins) get an answer; everything
 *  else in this file is unchanged and still works from the sheet menu.
 * ========================================================================= */
const PORTAL_SS_ID = '1WC6WGmwiqqytYobfoM9P4dP1RIptj5Rj17LSYtP_8tI';
const PORTAL_MAX_BODY = 12 * 1024 * 1024;        // media uploads come as base64
const PORTAL_RATE = 150;                          // calls per session per minute
let PORTAL_EDIT_ = false;                         // set for a checked portal call: canEdit_() lets it through
let PORTAL_USER_ = null;

function portalJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function portalSess_(token) {
  token = String(token || '').trim();
  if (!token || token.length > 200) return null;
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  const sh = SpreadsheetApp.openById(PORTAL_SS_ID).getSheetByName('Sessions');
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]) !== hash) continue;
    if (!(v[i][5] instanceof Date) || v[i][5] <= new Date()) return null;
    const role = String(v[i][3]).toLowerCase();
    const acc = String(v[i][6] || '').toLowerCase().split(/[\s,;]+/).filter(Boolean);
    return { username: String(v[i][1]), name: String(v[i][2]), role: role, admin: role === 'admin',
             complaints: role === 'admin' || acc.indexOf('complaints') > -1 };
  }
  return null;
}
function portalRateOk_(key) {
  const c = CacheService.getScriptCache(), b = 'pr_' + key + '_' + Math.floor(Date.now() / 60000);
  const n = Number(c.get(b) || 0) + 1; c.put(b, String(n), 70);
  return n <= PORTAL_RATE;
}
// text going into a cell: never a formula
function pSafe_(v, max) {
  if (v === null || v === undefined) return '';
  if (typeof v !== 'string') return v;
  let t = v.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  if (max && t.length > max) t = t.slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
function pClean_(v, d) {
  if (d > 6) return null;
  if (Array.isArray(v)) return v.slice(0, 3000).map(x => pClean_(x, d + 1));
  if (v && typeof v === 'object') { const o = {}; Object.keys(v).slice(0, 80).forEach(k => { o[k] = pClean_(v[k], d + 1); }); return o; }
  if (typeof v === 'string') return pSafe_(v, 5000);
  if (typeof v === 'number') return isFinite(v) ? v : null;
  return typeof v === 'boolean' ? v : null;
}

// ----- web versions of the menu tools (same work, a result instead of a pop-up) -----
function pCleanup_(apply) {
  const sh = getSheet_(), last = lastDataRow_(sh);
  if (last < 2) return { ok: true, text: 'No data rows found.' };
  const rng = sh.getRange(2, 1, last - 1, FIELDS.length), vals = rng.getValues(), disp = rng.getDisplayValues();
  const MANUAL = [3, 10, 11, 12, 16, 17, 18, 19, 20, 21, 22, 23];
  const untouched = i => MANUAL.every(c => String(vals[i][c - 1]).trim() === '');
  const found = [], ids = [];
  vals.forEach((r, i) => {
    const model = String(r[COL_MODEL - 1]).trim(), serial = String(r[COL_SERIAL - 1]).trim(), bug = bugNumber_(disp[i][COL_BUG - 1]);
    if (!untouched(i)) return;
    if (!model && !serial && !bug) { found.push({ i: i, why: 'Empty row' }); return; }
    const cat = String(disp[i][COL_CATEGORY - 1]).trim();
    if (bug && (cat === '' || isTbd_(cat))) ids.push(bug);
  });
  const uniq = Array.from(new Set(ids)), wps = uniq.length ? opFetchMany_(uniq) : {}, cache = {}, map = getModelMap_(), ex = {};
  Object.keys(wps).forEach(id => { ex[id] = { wp: wps[id], x: opExtract_(wps[id], opSchemaFor_(wps[id], cache)) }; });
  vals.forEach((r, i) => {
    const bug = bugNumber_(disp[i][COL_BUG - 1]), e = ex[bug];
    if (!bug || !e || !untouched(i)) return;
    const cat = String(disp[i][COL_CATEGORY - 1]).trim();
    if (!(cat === '' || isTbd_(cat))) return;
    const serial = String(r[COL_SERIAL - 1]).trim() || String(e.x.data.serial || '').replace(/\s+/g, '');
    if (!opIsProjector_(e.wp, { data: Object.assign({}, e.x.data, { serial: serial }) }, map)) found.push({ i: i, why: serial ? 'Not a projector case' : 'No serial number' });
  });
  found.sort((a, b) => a.i - b.i);
  const list = found.map(f => ({ row: f.i + 2, bug: bugNumber_(disp[f.i][COL_BUG - 1]), model: disp[f.i][COL_MODEL - 1], serial: disp[f.i][COL_SERIAL - 1], why: f.why }));
  if (apply) {
    const rows = found.map(f => f.i + 2).sort((a, b) => b - a);
    for (let k = 0; k < rows.length;) { let n = 1; while (k + n < rows.length && rows[k + n] === rows[k] - n) n++; sh.deleteRows(rows[k] - n + 1, n); k += n; }
    SpreadsheetApp.flush();
  }
  const cnt = w => found.filter(f => f.why === w).length;
  return { ok: true, list: list, text: (apply ? 'Removed: ' : 'Would remove: ') + found.length + ' row(s). Not a projector case: ' + cnt('Not a projector case') +
    ', no serial number: ' + cnt('No serial number') + ', empty rows: ' + cnt('Empty row') + '. Rows filled in by hand are never touched.' };
}
function pRepair_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return { ok: false, error: 'Another update is running. Try again in a minute.' };
  try {
    const ss = SpreadsheetApp.openById(SS_ID), S = syncSource_(ss);
    if (!S) return { ok: true, text: 'No data rows found in "' + SRC_SHEET_NAME + '".' };
    const bySrc = {};
    S.vals.forEach((r, i) => { const k = rowKey_(S.disp[i][S.cBug], r[S.cSer], r[S.cCmp]); if (k && bySrc[k] === undefined) bySrc[k] = i; });
    const lines = SYNC_RULES.map(rule => repairRule_(ss, S, bySrc, rule));
    SpreadsheetApp.flush();
    return { ok: true, text: lines.join('\n') };
  } finally { lock.releaseLock(); }
}
function pBackfill_(apply) {
  const sh = getSheet_();
  const X = bfLoad_(sh);
  if (!X) return { ok: true, text: 'No data rows found.' };
  bfReadTickets_(X);
  X.vals.forEach((r, i) => bfRow_(X, i));
  if (apply) bfWrite_(X);
  const head = ['Sheet row', 'Sl No', 'Bug ID', 'Field', 'Old value', 'New value', 'Lumio model', 'Note'];
  const name = apply ? 'OP Backfill Log' : 'OP Backfill Preview';
  const lg = X.ss.getSheetByName(name) || X.ss.insertSheet(name);
  lg.clear();
  const rows = [head].concat(X.log.length ? X.log : [['', '', '', 'No changes needed', '', '', '', '']]);
  lg.getRange(1, 1, rows.length, head.length).setNumberFormat('@').setValues(rows.map(r => r.map(x => String(x == null ? '' : x))));
  lg.getRange(1, 1, 1, head.length).setFontWeight('bold'); lg.setFrozenRows(1);
  const st = X.stat;
  return { ok: true, log: X.log.slice(0, 400).map(r => r.map(x => String(x == null ? '' : x))), logTotal: X.log.length,
    text: 'Rows checked: ' + st.rows + '. Tickets read: ' + st.found + ', not found: ' + st.notFound + ', rows without Bug ID: ' + st.noBug + '.\n' +
      (apply ? 'Changed: ' : 'Would change: ') + 'Model ' + st.model + ', Serial ' + st.serial + ', Invoice date ' + st.inv + ', Complaint date ' + st.cmp +
      ', calculated cells ' + st.calc + ', Category ' + st.cat + '. Warnings: ' + st.flags + '. Details in the "' + name + '" tab.' };
}
// automatic jobs: which are on, and turning them on / off (the same triggers the sheet menu makes)
const P_AUTO_ = {
  import:  { handlers: ['autoImportTickets'], label: 'Ticket import every 15 minutes (also runs the data check)' },
  sync:    { handlers: ['onEditCategorySync', 'autoSyncCategories'], label: 'Category sync on edit and every 10 minutes' },
  model:   { handlers: ['onEditModelSheets', 'autoModelSheets'], label: 'Model sheets on Production Details edit and hourly' },
  checks:  { handlers: ['autoDataCheck'], label: 'Daily data check with email (about 8 am)' }
};
function pAutoState_() {
  const on = {};
  ScriptApp.getProjectTriggers().forEach(t => { on[t.getHandlerFunction()] = true; });
  const o = {}; Object.keys(P_AUTO_).forEach(k => { o[k] = { on: P_AUTO_[k].handlers.some(h => on[h]), label: P_AUTO_[k].label }; });
  return o;
}
function pSetAuto_(kind, turnOn) {
  const def = P_AUTO_[kind]; if (!def) throw new Error('Unknown automatic job.');
  ScriptApp.getProjectTriggers().forEach(t => { if (def.handlers.indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t); });
  if (turnOn) {
    const ss = SpreadsheetApp.openById(SS_ID);
    if (kind === 'import') ScriptApp.newTrigger('autoImportTickets').timeBased().everyMinutes(15).create();
    if (kind === 'sync') { ScriptApp.newTrigger('onEditCategorySync').forSpreadsheet(ss).onEdit().create(); ScriptApp.newTrigger('autoSyncCategories').timeBased().everyMinutes(10).create(); }
    if (kind === 'model') { ScriptApp.newTrigger('onEditModelSheets').forSpreadsheet(ss).onEdit().create(); ScriptApp.newTrigger('autoModelSheets').timeBased().everyHours(1).create(); }
    if (kind === 'checks') ScriptApp.newTrigger('autoDataCheck').timeBased().everyDays(1).atHour(8).create();
  }
  return pAutoState_();
}
function pStatus_() {
  const props = PropertiesService.getScriptProperties();
  return { auto: pAutoState_(), opKeySet: !!props.getProperty('OP_API_KEY'), importPage: Number(props.getProperty('OP_IMPORT_PAGE')) || 1,
    checks: dcOpts_(), checkEmails: dcEmails_(), sheetUrl: SpreadsheetApp.openById(SS_ID).getUrl(), opBase: OP_BASE,
    limits: { doa: DOA_LIMITS, afr: AFR_LIMITS, colors: LIMIT_COLORS, oqcTarget: OQC_TARGET, obaTarget: OBA_TARGET } };
}

// action -> [admin only?, function(args)]
const P_API_ = {
  // complaints and form
  getData:        [0, a => getData()],
  getStatus:      [0, a => pStatus_()],
  fetchBugData:   [0, a => fetchBugData(a.bugId)],
  fetchTicketView:[0, a => fetchTicketView(a.bugId)],
  fetchAttachment:[0, a => fetchAttachmentData(a.attId)],
  saveRecord:     [0, a => { const r = saveRecord(a.record || {}, a.mode === 'edit' ? 'edit' : 'new'); delete r.data; return r; }],
  uploadMedia:    [0, a => { const f = a.file || {};
                     if (!/^(image|video)\/[a-z0-9.+-]+$/i.test(String(f.mime || '')) && String(f.mime) !== 'application/pdf') throw new Error('Only images, videos and PDF files can be uploaded.');
                     return { url: uploadMedia({ name: String(f.name || 'file').replace(/^'/, '').slice(0, 120), mime: f.mime, data: String(f.data || '') }) }; }],
  // dashboard, AFR, data entry
  getAfr:         [0, a => getAfr()],
  getModelTrend:  [0, a => getModelTrend()],
  getProduction:  [0, a => getDashProduction()],
  getInspection:  [0, a => getInspection()],
  getEntryData:   [0, a => getEntryData()],
  saveProduction: [0, a => saveProductionData(a.model, a.rows)],
  saveSales:      [0, a => saveSalesData(a.model, a.rows)],
  saveOqc:        [0, a => saveOqcData(a.model, a.rows)],
  saveOba:        [0, a => saveObaData(a.model, a.rows)],
  // data checks
  getDataChecks:  [0, a => getDataChecks(null, a.opts)],
  runDataChecks:  [0, a => runDataChecksWeb(null, !!a.mail, a.opts)],
  // menu tools
  syncPreview:    [0, a => ({ ok: true, text: syncText_(syncCategories_(null, 'portal', false, true), true) })],
  syncNow:        [0, a => ({ ok: true, text: syncText_(syncCategories_(null, 'portal', false, false), false) })],
  syncRepair:     [0, a => pRepair_()],
  importPreview:  [0, a => { const r = importTickets_(true, 'portal'); return { ok: r.ok, text: importText_(r) }; }],
  importNow:      [0, a => { const r = importTickets_(false, 'portal'); return { ok: r.ok, text: importText_(r) }; }],
  importReset:    [1, a => { PropertiesService.getScriptProperties().deleteProperty('OP_IMPORT_PAGE'); return { ok: true, text: 'The next import reads the tickets again from the first one. Tickets already in the sheet are not added twice.' }; }],
  cleanupPreview: [0, a => pCleanup_(false)],
  cleanupApply:   [1, a => pCleanup_(true)],
  lotsPreview:    [0, a => { const r = fillLots_(true); return { ok: r.ok, text: lotsText_(r) }; }],
  lotsNow:        [0, a => { const r = fillLots_(false); return { ok: r.ok, text: lotsText_(r) }; }],
  modelPreview:   [0, a => { const r = syncModelSheets_(false, true, 'portal'); return { ok: r.ok, text: msText_(r) }; }],
  modelNow:       [0, a => { const r = syncModelSheets_(false, false, 'portal'); return { ok: r.ok, text: msText_(r) }; }],
  modelRebuild:   [1, a => { const r = syncModelSheets_(true, false, 'portal'); return { ok: r.ok, text: msText_(r) }; }],
  setCheckScope:  [0, a => ({ ok: true, opts: dcSaveOpts_({ scope: a.scope === 'all' ? 'all' : 'open', dup: a.dup }) })],
  backfillPreview:[0, a => pBackfill_(false)],
  backfillApply:  [1, a => pBackfill_(true)],
  setAuto:        [1, a => ({ ok: true, auto: pSetAuto_(String(a.kind), !!a.on) })],
  setOpKey:       [1, a => { const k = String(a.key || '').replace(/^'/, '').trim();
                     if (k.length < 10 || k.length > 200 || /\s/.test(k)) throw new Error('That does not look like an OpenProject API token.');
                     PropertiesService.getScriptProperties().setProperty('OP_API_KEY', k); return { ok: true, text: 'API key saved. Test a Bug ID to check the connection.' }; }],
  setCheckEmails: [1, a => { const list = String(a.list || '').split(/[\s,;]+/).filter(Boolean);
                     if (list.some(x => !/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(x)) || list.length > 10) throw new Error('Enter up to 10 email addresses, separated by commas.');
                     PropertiesService.getScriptProperties().setProperty('DC_EMAILS', list.join(',')); return { ok: true, text: 'Check emails go to: ' + (list.join(', ') || 'the owner') }; }]
};
const P_READ_ = ['getData', 'getStatus', 'fetchBugData', 'fetchTicketView', 'fetchAttachment', 'getAfr', 'getModelTrend', 'getProduction', 'getInspection', 'getEntryData', 'getDataChecks',
  'syncPreview', 'importPreview', 'cleanupPreview', 'lotsPreview', 'modelPreview', 'backfillPreview'];

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > PORTAL_MAX_BODY) return portalJson_({ ok: false, error: 'Request is too large (files up to 8 MB).' });
    let d; try { d = JSON.parse(raw); } catch (x) { return portalJson_({ ok: false, error: 'Bad request.' }); }
    if (!d || typeof d !== 'object' || Array.isArray(d)) return portalJson_({ ok: false, error: 'Bad request.' });
    const sess = portalSess_(d.token);
    if (!sess) return portalJson_({ ok: false, auth: true, error: 'Your session has expired. Please log in again.' });
    if (!sess.complaints) return portalJson_({ ok: false, error: 'You do not have access to this. Ask an admin to give you the Customer complaints module.' });
    if (!portalRateOk_(String(d.token).slice(0, 24))) return portalJson_({ ok: false, error: 'Too many requests. Wait a minute and try again.' });
    const action = String(d.action || ''), spec = P_API_[action];
    if (!spec) return portalJson_({ ok: false, error: 'Unknown action.' });
    if (spec[0] && !sess.admin) return portalJson_({ ok: false, error: 'Only an admin can do this.' });
    const args = action === 'uploadMedia' ? { file: d.args && d.args.file } : pClean_(d.args || {}, 0);
    PORTAL_EDIT_ = true; PORTAL_USER_ = sess;
    let res;
    try { res = spec[1](args); }
    catch (x) { return portalJson_({ ok: false, error: errText_(x) }); }
    finally { PORTAL_EDIT_ = false; }
    if (res && typeof res === 'object' && !Array.isArray(res) && 'ok' in res) return portalJson_(res);
    return portalJson_({ ok: true, data: res });
  } catch (err) {
    console.error('Portal API failed: ' + (err && err.stack || err));
    return portalJson_({ ok: false, error: 'Server error. Please try again.' });
  }
}
