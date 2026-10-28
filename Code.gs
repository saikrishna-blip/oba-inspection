/****************************************************
 * IS2500 OBA Batch Inspection — Google Apps Script
 * SIMPLE VERSION - No switch, just if/else
 * v4: Models + Black Dot spec are stored in the "Models" sheet
 *     and managed by admin from the app (nothing hardcoded).
 ****************************************************/

var SS_ID        = "1WC6WGmwiqqytYobfoM9P4dP1RIptj5Rj17LSYtP_8tI";
var FIXED_LOT    = 90;
var FIXED_SAMPLE = 13;

// BOM reference spreadsheet (Lumio_Part_Reference_DB): BOM tabs per model,
// BOM_CONTROL, BOM_VERIFY_LOG, LOT_BOM_DETAIL
var BOM_SS_ID    = "11BFVHMfnDGymLNYfBNLp9CRKIAE0X5fLX8qX-R8DeYk";
var TZ           = "Asia/Kolkata";

// Part DB web app (BOM management tools, bound to the BOM sheet). Admins open it through the portal.
var PARTDB_URL   = "https://script.google.com/macros/s/AKfycbyPiE-gcrUZH7cEA6KKBgXHJa93MUl2r7-GseHkZ0FGdKBmTJrOnl7r7SOpm3SNjdjZ/exec";

// Login sessions (sheet "Sessions"; the Part DB script reads it too)
var SESSION_HOURS = 12;
// Actions anyone signed in may call; every other action except login needs the admin role.
var MEMBER_ACTIONS = ["logout","getBatches","getModels","getBomList","getBom","getBomVerification",
                      "checkSerial","submitLot","submitBomVerification","getLotUnits"];

// Batches sheet stage columns (1-based)
var COL_BOM_STATUS = 16;   // "", Pending, Passed, Failed, Waived  ("" = created before the BOM step)
var COL_BOM_VER    = 17;   // Ver_ID of the verification that set the status
var COL_FAI_STATUS = 18;   // reserved for FAI inspection

function getSpreadsheet(){
  return SpreadsheetApp.openById(SS_ID);
}

function getOrCreateSheet(name, headers){
  var ss = getSpreadsheet();
  var s  = ss.getSheetByName(name);
  if(!s){
    s = ss.insertSheet(name);
    if(headers && headers.length > 0) s.appendRow(headers);
  }
  return s;
}

function setup(){
  getOrCreateSheet("Batches", [
    "BatchName","LotNumber","LotSize","SampleSize","BDMax",
    "Status","CreatedBy","CreatedAt","Result",
    "BatchQty","TotalLots","Model","ProdMonthCode","FGPassQty","ProdMonthCode2"
  ]);
  getOrCreateSheet("Master", [
    "Timestamp","Inspector","BatchName","LotNumber","SerialNumber",
    "Aesthetic","General","Functional","Screen","BlackDot",
    "Remark","SampleResult","LotResult",
    "BatchQty","TotalLots","Model","ProdMonthCode","FGPassQty","BDSpec"
  ]);
  getOrCreateSheet("Users", ["Username","Password","Role","Name"]);
  modelsSheet_();

  var users = getOrCreateSheet("Users", []);
  if(users.getLastRow() < 2){
    users.appendRow(["admin","admin123","admin","Administrator"]);
    users.appendRow(["Rahul","Rahul123","inspector","Rahul C"]);
    users.appendRow(["inspector1","insp123","inspector","Inspector 1"]);
    users.appendRow(["saikrishna","Sai@1996","inspector","Appireddy Sai Krishna"]);
  }
  return "Setup done";
}

function testCreateBatch(){
  var m = readModels_()[0];
  if(!m){ Logger.log("Add a model in the app first"); return; }
  var result = createBatch({
    batchName: "TEST_BATCH",
    batchQty: 200,
    model: m.name,
    prodMonthCode: "0626",
    createdBy: "admin"
  });
  Logger.log(JSON.stringify(result));
  return result;
}

function doGet(e){
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, msg:"LUMIO Quality Management - Projector API running", version:"v5"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try {
    var raw  = e.postData.contents;
    var data = JSON.parse(raw);
    var action = data.action;
    var result = {ok:false, error:"Unknown action: " + action};

    // Every action except login needs a valid session; most need the admin role.
    if(action !== "login"){
      var sess = sessionFor_(data.token);
      if(!sess) return json_({ok:false, auth:true, error:"Your session has expired. Please log in again."});
      if(MEMBER_ACTIONS.indexOf(action) === -1 && sess.role !== "admin")
        return json_({ok:false, error:"Only an admin can do this."});
      data._session = sess;
      // record who did it from the session, not from what the browser sent
      if(action === "submitLot") data.inspector = sess.name;
      if(action === "submitBomVerification") data.by = sess.name;
      if(["createBatch","saveModel","removeDuplicates","approveSerial","setBatchStage"].indexOf(action) > -1){
        data.createdBy = sess.username; data.updatedBy = sess.username; data.by = sess.username;
      }
    }

    if(action === "login"){
      result = login(data);
    }
    if(action === "addUser"){
      result = addUser(data);
    }
    if(action === "getBatches"){
      result = getBatches();
    }
    if(action === "createBatch"){
      result = createBatch(data);
    }
    if(action === "updateLot"){
      result = updateLot(data);
    }
    if(action === "deleteBatch"){
      result = deleteBatch(data);
    }
    if(action === "deleteLot"){
      result = deleteLot(data);
    }
    if(action === "checkSerial"){
      result = checkSerial(data);
    }
    if(action === "submitLot"){
      result = submitLot(data);
    }
    if(action === "getMaster"){
      result = getMaster();
    }
    if(action === "getModels"){
      result = getModels();
    }
    if(action === "saveModel"){
      result = saveModel(data);
    }
    if(action === "deleteModel"){
      result = deleteModel(data);
    }
    if(action === "getLotUnits"){
      result = getLotUnits(data);
    }
    if(action === "scanDuplicates"){
      result = scanDuplicates();
    }
    if(action === "removeDuplicates"){
      result = removeDuplicates(data);
    }
    if(action === "searchRecords"){
      result = searchRecords(data);
    }
    if(action === "approveSerial"){
      result = approveSerial(data);
    }
    if(action === "logout"){
      result = logout(data);
    }
    if(action === "getBomList"){
      result = getBomList();
    }
    if(action === "getBom"){
      result = getBom(data);
    }
    if(action === "submitBomVerification"){
      result = submitBomVerification(data);
    }
    if(action === "getBomVerification"){
      result = getBomVerification(data);
    }
    if(action === "setBatchStage"){
      result = setBatchStage(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error:"Server error: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function login(d){
  var rows = getOrCreateSheet("Users",[]).getDataRange().getValues();
  for(var i = 1; i < rows.length; i++){
    if(String(rows[i][0]) === String(d.username) && String(rows[i][1]) === String(d.password)){
      var role = String(rows[i][2]||"inspector").trim().toLowerCase();
      var s = createSession_(String(rows[i][0]), String(rows[i][3]||rows[i][0]), role);
      var out = {ok:true, username:rows[i][0], role:role, name:rows[i][3], token:s.token, expiresAt:s.expiresAt};
      if(role === "admin"){ out.partdbUrl = PARTDB_URL; out.bomSheetUrl = "https://docs.google.com/spreadsheets/d/" + BOM_SS_ID + "/edit"; }
      return out;
    }
  }
  return {ok:false, error:"Invalid username or password"};
}

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- sessions ----------------
 * Sessions sheet: TokenHash | Username | Name | Role | CreatedAt | ExpiresAt
 * Only a SHA-256 hash of each token is stored. */
function sha256_(s){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function(b){ return ("0" + (b & 0xff).toString(16)).slice(-2); }).join("");
}
function sessionsSheet_(){
  var sh = getSpreadsheet().getSheetByName("Sessions");
  if(!sh){
    sh = getSpreadsheet().insertSheet("Sessions");
    sh.appendRow(["TokenHash","Username","Name","Role","CreatedAt","ExpiresAt"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function createSession_(username, name, role){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var sh = sessionsSheet_();
    var now = new Date();
    // drop expired sessions
    var v = sh.getDataRange().getValues();
    var keep = [v[0]];
    for(var i = 1; i < v.length; i++){ if(v[i][5] instanceof Date && v[i][5] > now) keep.push(v[i]); }
    if(keep.length !== v.length){
      sh.getRange(1, 1, v.length, 6).clearContent();
      sh.getRange(1, 1, keep.length, 6).setValues(keep);
    }
    var token = Utilities.getUuid().replace(/-/g,"") + Utilities.getUuid().replace(/-/g,"");
    var exp = new Date(now.getTime() + SESSION_HOURS * 3600 * 1000);
    sh.appendRow([sha256_(token), username, name, role, now, exp]);
    return {token:token, expiresAt:exp.toISOString()};
  } finally { lock.releaseLock(); }
}
function sessionFor_(token){
  token = String(token||"").trim();
  if(!token) return null;
  var h = sha256_(token);
  var cache = CacheService.getScriptCache();
  var hit = cache.get("s_" + h);
  if(hit) return JSON.parse(hit);
  var v = sessionsSheet_().getDataRange().getValues();
  for(var i = v.length - 1; i >= 1; i--){
    if(String(v[i][0]) !== h) continue;
    if(!(v[i][5] instanceof Date) || v[i][5] <= new Date()) return null;
    var s = {username:String(v[i][1]), name:String(v[i][2]), role:String(v[i][3]).toLowerCase()};
    cache.put("s_" + h, JSON.stringify(s), 300);
    return s;
  }
  return null;
}
function logout(d){
  var h = sha256_(String(d.token||""));
  CacheService.getScriptCache().remove("s_" + h);
  var sh = sessionsSheet_(), v = sh.getDataRange().getValues();
  for(var i = v.length - 1; i >= 1; i--){
    if(String(v[i][0]) === h) sh.getRange(i+1, 6).setValue(new Date(0));
  }
  return {ok:true};
}

function addUser(d){
  getOrCreateSheet("Users",[]).appendRow([d.username, d.password, d.role||"inspector", d.name]);
  return {ok:true};
}

function getBatches(){
  var sheet = getOrCreateSheet("Batches",[]);
  var rows  = sheet.getDataRange().getValues();
  var out   = [];
  for(var i = 1; i < rows.length; i++){
    var r = rows[i];
    var rawPM = String(r[12]||"").replace(/^'/,'');
    if(rawPM && rawPM.length===3) rawPM = '0'+rawPM;
    var rawPM2 = String(r[14]||"").replace(/^'/,'');
    if(rawPM2 && rawPM2.length===3) rawPM2 = '0'+rawPM2;
    out.push({
      batchName     : String(r[0]||""),
      lotNumber     : String(r[1]||""),
      lotSize       : Number(r[2])||FIXED_LOT,
      sampleSize    : Number(r[3])||FIXED_SAMPLE,
      bdMax         : Number(r[4])||0,
      status        : String(r[5]||"Open"),
      createdBy     : String(r[6]||""),
      createdAt     : r[7] ? r[7].toString() : "",
      result        : String(r[8]||""),
      batchQty      : Number(r[9])||0,
      totalLots     : Number(r[10])||0,
      model         : String(r[11]||""),
      prodMonthCode : rawPM,
      fgPassQty     : isNaN(Number(r[13])) ? 0 : Number(r[13]),
      prodMonthCode2: rawPM2,
      bomStatus     : String(r[COL_BOM_STATUS-1]||""),
      bomVerId      : String(r[COL_BOM_VER-1]||""),
      faiStatus     : String(r[COL_FAI_STATUS-1]||"")
    });
  }
  return {ok:true, batches:out};
}

function createBatch(d){
  var batchName = String(d.batchName||d.kitName||"").trim();
  var batchQty  = parseInt(d.batchQty||d.kitSize)||0;
  var model     = String(d.model||"");
  var pm        = String(d.prodMonthCode||"");
  var pm2       = String(d.prodMonthCode2||"");
  var by        = String(d.createdBy||"");

  if(!batchName) return {ok:false, error:"Batch Name is required"};
  if(batchQty < 1) return {ok:false, error:"Batch Quantity must be greater than 0"};
  var spec = findModel_(model);
  if(!spec) return {ok:false, error:"Model \"" + model + "\" not found. Add it in Models first."};
  var bdMax = spec.maxD;

  var fullLots  = Math.floor(batchQty / FIXED_LOT);
  var remainder = batchQty % FIXED_LOT;
  var total     = fullLots + (remainder > 0 ? 1 : 0);
  var sheet     = getOrCreateSheet("Batches",[]);
  ensureStageHeaders_(sheet);
  var now       = new Date();

  for(var i = 1; i <= fullLots; i++){
    sheet.appendRow([
      batchName,
      batchName + " Lot " + i,
      FIXED_LOT, FIXED_SAMPLE, bdMax, "Open", by, now, "",
      batchQty, total, model, "'" + pm, "", "'" + pm2, "Pending", "", ""
    ]);
  }

  if(remainder > 0){
    sheet.appendRow([
      batchName,
      batchName + " Lot " + (fullLots + 1),
      remainder, FIXED_SAMPLE, bdMax, "Open", by, now, "",
      batchQty, total, model, "'" + pm, "", "'" + pm2, "Pending", "", ""
    ]);
  }

  return {ok:true, lotsCreated:total, batchQty:batchQty};
}

function updateLot(d){
  var sheet = getOrCreateSheet("Batches",[]);
  var rows  = sheet.getDataRange().getValues();
  for(var i = 1; i < rows.length; i++){
    if(rows[i][0] === d.batchName && rows[i][1] === d.lotNumber){
      if(d.status        !== undefined) sheet.getRange(i+1, 6).setValue(d.status);
      if(d.bdMax         !== undefined) sheet.getRange(i+1, 5).setValue(d.bdMax);
      if(d.model         !== undefined) sheet.getRange(i+1, 12).setValue(d.model);
      if(d.prodMonthCode !== undefined) sheet.getRange(i+1, 13).setValue(d.prodMonthCode);
      if(d.prodMonthCode2!== undefined) sheet.getRange(i+1, 15).setValue(d.prodMonthCode2);
      return {ok:true};
    }
  }
  return {ok:false, error:"Lot not found"};
}

function deleteBatch(d){
  var sheet = getOrCreateSheet("Batches",[]);
  var rows  = sheet.getDataRange().getValues();
  for(var i = rows.length - 1; i >= 1; i--){
    if(rows[i][0] === d.batchName) sheet.deleteRow(i + 1);
  }
  return {ok:true};
}

function deleteLot(d){
  var sheet = getOrCreateSheet("Batches",[]);
  var rows  = sheet.getDataRange().getValues();
  for(var i = rows.length - 1; i >= 1; i--){
    if(rows[i][0] === d.batchName && rows[i][1] === d.lotNumber){
      sheet.deleteRow(i + 1);
      return {ok:true};
    }
  }
  return {ok:false, error:"Lot not found"};
}

function checkSerial(d){
  if(!d.serial) return {ok:true, exists:false};
  var rows = getOrCreateSheet("Master",[]).getDataRange().getValues();
  for(var i = 1; i < rows.length; i++){
    if(String(rows[i][4]) === String(d.serial))
      return {ok:true, exists:true, batch:rows[i][2], lot:rows[i][3]};
  }
  return {ok:true, exists:false};
}

function submitLot(d){
  // Guard: check if this lot is already closed to prevent duplicate submissions
  var bsh   = getOrCreateSheet("Batches",[]);
  var bRows = bsh.getDataRange().getValues();
  for(var i=1;i<bRows.length;i++){
    if(String(bRows[i][0])===String(d.batchName) &&
       String(bRows[i][1])===String(d.lotNumber) &&
       String(bRows[i][5])==='Closed'){
      return {ok:false, error:'This lot is already submitted and closed. Duplicate submission blocked.'};
    }
  }
  // Flow gate: sampling needs BOM verification passed (or waived) for batches created with the flow
  for(var g=1; g<bRows.length; g++){
    if(String(bRows[g][0]) !== String(d.batchName)) continue;
    var bs = String(bRows[g][COL_BOM_STATUS-1]||"");
    if(bs === "Pending" || bs === "Failed")
      return {ok:false, error:"BOM verification for " + d.batchName + " is " + (bs === "Pending" ? "not done yet" : "failed") + ". Sampling can start after BOM verification passes."};
    break;
  }

  var master = getOrCreateSheet("Master",[]);
  if(master.getRange(1, 19).getValue() === "") master.getRange(1, 19).setValue("BDSpec");

  // Rule: a serial number is recorded only once, ever
  var seen = {}, existing = {}, clash = [];
  var mRows = master.getDataRange().getValues();
  for(var x = 1; x < mRows.length; x++){
    var sx = normSerial_(mRows[x][4]);
    if(sx && !existing[sx]) existing[sx] = String(mRows[x][3]||"");
  }
  (d.units||[]).forEach(function(u){
    var su = normSerial_(u.serial);
    if(!su) return;
    if(seen[su]) clash.push(su + " (repeated in this lot)");
    else if(existing[su]) clash.push(su + " (already recorded in " + existing[su] + ")");
    seen[su] = true;
  });
  if(clash.length) return {ok:false, error:"Duplicate serial number, nothing was saved: " + clash.join(", ")};

  var units = d.units || [];
  for(var i = 0; i < units.length; i++){
    var u = units[i];
    var sampleTs = u.inspectedAt ? new Date(u.inspectedAt) : new Date();
    master.appendRow([
      sampleTs, d.inspector, d.batchName, d.lotNumber, u.serial,
      u.aes, u.general, u.func, u.screen, u.bd,
      u.remark, u.sampleResult, d.lotResult,
      d.batchQty||"", d.totalLots||"",
      d.model||"", d.prodMonthCode||"", d.fgPassQty||0,
      d.bdSpec||""
    ]);
  }

  var sheet = getOrCreateSheet("Batches",[]);
  var rows  = sheet.getDataRange().getValues();
  for(var j = 1; j < rows.length; j++){
    if(rows[j][0] === d.batchName && rows[j][1] === d.lotNumber){
      sheet.getRange(j+1, 6).setValue("Closed");
      sheet.getRange(j+1, 9).setValue(d.lotResult);
      sheet.getRange(j+1, 14).setValue(d.fgPassQty||0);
      break;
    }
  }
  return {ok:true};
}

function getMaster(){
  var rows = getOrCreateSheet("Master",[]).getDataRange().getValues();
  return {ok:true, rows:rows.slice(1)};
}
/****************************************************
 * MODELS + BLACK DOT SPEC  (sheet: "Models")
 * name | code | maxD | smallD | maxN | minDS | nA | nB | nAB | updatedBy | updatedAt | lumioModel
 *   code   = first 4 digits of serial
 *   lumioModel = Lumio model number (product name/number shown to users)
 *   maxD   = FAIL if D > maxD
 *   smallD = OK if D <= smallD (no area check)
 *   maxN   = FAIL if N > maxN
 *   minDS  = FAIL if DS < minDS (only when N > 1)
 *   nA/nB/nAB = max N in Area A / B / A+B for smallD < D <= maxD
 ****************************************************/
var MODEL_COLS     = ["name","code","maxD","smallD","maxN","minDS","nA","nB","nAB","updatedBy","updatedAt","lumioModel"];
var MODEL_NUM_COLS = ["maxD","smallD","maxN","minDS","nA","nB","nAB"];

function modelsSheet_(){
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName("Models");
  if(!sh){
    sh = ss.insertSheet("Models");
    sh.appendRow(MODEL_COLS);
    sh.getRange("B:B").setNumberFormat("@");   // keep code as text (e.g. 0123)
    sh.setFrozenRows(1);
  }
  if(sh.getRange(1, MODEL_COLS.length).getValue() === "") sh.getRange(1, MODEL_COLS.length).setValue(MODEL_COLS[MODEL_COLS.length-1]);
  return sh;
}

function code4_(v){ return ("0000" + String(v).replace(/^'/,"").trim()).slice(-4); }

function readModels_(){
  var v = modelsSheet_().getDataRange().getValues();
  var out = [];
  for(var i = 1; i < v.length; i++){
    var r = v[i];
    if(String(r[0]).trim() === "") continue;
    out.push({
      name:String(r[0]).trim(), code:code4_(r[1]),
      maxD:Number(r[2]), smallD:Number(r[3]), maxN:Number(r[4]), minDS:Number(r[5]),
      nA:Number(r[6]), nB:Number(r[7]), nAB:Number(r[8]),
      lumioModel:String(r[11]||"").trim()
    });
  }
  return out;
}

function mostCommon_(counts){
  var best = "", n = 0;
  for(var k in counts){ if(counts[k] > n){ best = k; n = counts[k]; } }
  return best;
}

/* Models + what is already in Batches/Master (used to suggest models to add) */
function getModels(){
  var models = readModels_();
  var known = {};
  models.forEach(function(m){ known[m.name.toLowerCase()] = true; });

  var info = {};   // model name -> {lots, open, closed, bd:{}, codes:{}, months:{}}
  function slot(name){
    if(!info[name]) info[name] = {lots:0, open:0, closed:0, bd:{}, codes:{}, months:{}, batches:{}};
    return info[name];
  }
  var b = getOrCreateSheet("Batches",[]).getDataRange().getValues();
  for(var i = 1; i < b.length; i++){
    var name = String(b[i][11]||"").trim();
    if(!name) continue;
    var s = slot(name);
    s.lots++;
    if(String(b[i][5]) === "Open") s.open++; else s.closed++;
    s.batches[String(b[i][0])] = true;
    var bd = String(b[i][4]||""); if(bd) s.bd[bd] = (s.bd[bd]||0) + 1;
  }
  var m = getOrCreateSheet("Master",[]).getDataRange().getValues();
  for(var j = 1; j < m.length; j++){
    var mn = String(m[j][15]||"").trim();
    var serial = String(m[j][4]||"").trim();
    if(!mn || !/^\d{15}$/.test(serial)) continue;
    var ms = slot(mn);
    var c = serial.substring(0,4);
    ms.codes[c] = (ms.codes[c]||0) + 1;
  }

  var usage = {}, suggestions = [];
  for(var nm in info){
    var x = info[nm];
    var row = {
      name: nm,
      code: mostCommon_(x.codes),
      maxD: Number(mostCommon_(x.bd)) || "",
      lots: x.lots, openLots: x.open, closedLots: x.closed,
      batches: Object.keys(x.batches).length
    };
    usage[nm] = row;
    if(!known[nm.toLowerCase()]) suggestions.push(row);
  }
  return {ok:true, models:models, usage:usage, suggestions:suggestions};
}

function findModel_(name){
  var list = readModels_();
  for(var i = 0; i < list.length; i++){
    if(list[i].name === String(name).trim()) return list[i];
  }
  return null;
}

function saveModel(d){
  var m    = d.model || {};
  var name = String(m.name||"").trim();
  var code = String(m.code||"").trim();
  var lumio = String(m.lumioModel||"").trim();
  if(!name) return {ok:false, error:"Model name is required"};
  if(!lumio) return {ok:false, error:"Lumio model number is required"};
  if(!/^\d{4}$/.test(code)) return {ok:false, error:"Serial code must be exactly 4 digits"};
  for(var k = 0; k < MODEL_NUM_COLS.length; k++){
    var key = MODEL_NUM_COLS[k];
    if(m[key] === "" || m[key] === null || m[key] === undefined || isNaN(Number(m[key])) || Number(m[key]) < 0)
      return {ok:false, error:"Invalid value for " + key};
  }
  if(Number(m.smallD) > Number(m.maxD)) return {ok:false, error:"OK D limit cannot be bigger than Fail D limit"};

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var sh = modelsSheet_();
    var v  = sh.getDataRange().getValues();
    var original = String(d.originalName||"").trim().toLowerCase();
    var rowIdx = -1;
    for(var i = 1; i < v.length; i++){
      var rn = String(v[i][0]).trim().toLowerCase();
      if(original && rn === original){ rowIdx = i; continue; }
      if(rn === name.toLowerCase()) return {ok:false, error:"Model \"" + name + "\" already exists"};
      if(code4_(v[i][1]) === code)   return {ok:false, error:"Serial code " + code + " is already used by " + v[i][0]};
    }
    if(original && rowIdx === -1) return {ok:false, error:"Model to edit not found - refresh and try again"};

    var row = [name, code, Number(m.maxD), Number(m.smallD), Number(m.maxN), Number(m.minDS),
               Number(m.nA), Number(m.nB), Number(m.nAB), d.updatedBy||"", new Date(), lumio];
    var target;
    if(rowIdx === -1){
      sh.appendRow(row);
      target = sh.getLastRow();
    } else {
      target = rowIdx + 1;
      sh.getRange(target, 1, 1, row.length).setValues([row]);
    }
    sh.getRange(target, 2).setNumberFormat("@").setValue(code);

    // Renamed model: update open lots that used the old name
    if(original && original !== name.toLowerCase()){
      var bs = getOrCreateSheet("Batches",[]);
      var br = bs.getDataRange().getValues();
      for(var j = 1; j < br.length; j++){
        if(String(br[j][11]).trim().toLowerCase() === original) bs.getRange(j+1, 12).setValue(name);
      }
    }
    return {ok:true};
  } finally {
    lock.releaseLock();
  }
}

function deleteModel(d){
  var name = String(d.name||"").trim().toLowerCase();
  if(!name) return {ok:false, error:"Model name is required"};
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    var sh = modelsSheet_();
    var v  = sh.getDataRange().getValues();
    for(var i = v.length - 1; i >= 1; i--){
      if(String(v[i][0]).trim().toLowerCase() === name){
        sh.deleteRow(i + 1);
        return {ok:true};
      }
    }
    return {ok:false, error:"Model not found"};
  } finally {
    lock.releaseLock();
  }
}

/****************************************************
 * LOT RESULTS — sample rows from Master for one lot
 ****************************************************/
function getLotUnits(d){
  var bn = String(d.batchName||""), ln = String(d.lotNumber||"");
  if(!bn || !ln) return {ok:false, error:"Batch and lot are required"};
  var rows = getOrCreateSheet("Master",[]).getDataRange().getValues();
  var out = [];
  for(var i = 1; i < rows.length; i++){
    var r = rows[i];
    if(String(r[2]) !== bn || String(r[3]) !== ln) continue;
    out.push({
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]||""),
      inspector: String(r[1]||""), serial: String(r[4]||""),
      aes: String(r[5]||""), general: String(r[6]||""), func: String(r[7]||""), screen: String(r[8]||""),
      bd: String(r[9]||""), remark: String(r[10]||""),
      sampleResult: String(r[11]||""), lotResult: String(r[12]||"")
    });
  }
  return {ok:true, units:out};
}

/****************************************************
 * DUPLICATE SERIAL RULE
 * Each serial number may appear in Master only once.
 * The original is the earliest record (by timestamp, then row order).
 * Extra copies are moved to the "Duplicates" sheet, never deleted outright.
 ****************************************************/
function normSerial_(v){
  return String(v == null ? "" : v).replace(/^'/,"").replace(/\.0+$/,"").trim().toUpperCase();
}

function recOf_(r, rowNo){
  return {
    row: rowNo,
    time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]||""),
    inspector: String(r[1]||""), batch: String(r[2]||""), lot: String(r[3]||""),
    serial: normSerial_(r[4]), result: String(r[11]||""), lotResult: String(r[12]||""),
    remark: String(r[10]||"")
  };
}

/* Serials an admin approved to stay in more than one lot (one record per lot). */
function exceptionsSheet_(){
  var ss = getSpreadsheet();
  var sh = ss.getSheetByName("SerialExceptions");
  if(!sh){
    sh = ss.insertSheet("SerialExceptions");
    sh.appendRow(["SerialNumber","Reason","ApprovedBy","ApprovedAt"]);
    sh.getRange("A:A").setNumberFormat("@");
    sh.setFrozenRows(1);
  }
  return sh;
}
function readExceptions_(){
  var sh = getSpreadsheet().getSheetByName("SerialExceptions");
  var out = {};
  if(!sh) return out;
  var v = sh.getDataRange().getValues();
  for(var i = 1; i < v.length; i++){
    var s = normSerial_(v[i][0]);
    if(s) out[s] = {serial:s, reason:String(v[i][1]||""), by:String(v[i][2]||""),
                    at: v[i][3] instanceof Date ? v[i][3].toISOString() : String(v[i][3]||"")};
  }
  return out;
}
function approveSerial(d){
  var s = normSerial_(d.serial);
  if(!/^\d{15}$/.test(s)) return {ok:false, error:"Serial number must be 15 digits"};
  if(readExceptions_()[s]) return {ok:true, already:true};
  exceptionsSheet_().appendRow(["'" + s, String(d.reason||"Kept in more than one lot"), String(d.by||""), new Date()]);
  return {ok:true};
}

function findDuplicates_(rows, exc){
  exc = exc || {};
  var bySerial = {}, order = [];
  for(var i = 1; i < rows.length; i++){
    var sn = normSerial_(rows[i][4]);
    if(!sn) continue;
    // Approved exceptions: duplicates are only checked inside the same lot
    var s = exc[sn] ? sn + "||" + String(rows[i][3]||"") : sn;
    if(!bySerial[s]){ bySerial[s] = []; order.push(s); }
    bySerial[s].push(i);
  }
  var groups = [], dupIdx = {};
  order.forEach(function(s){
    var idx = bySerial[s];
    if(idx.length < 2) return;
    idx.sort(function(a, b){
      var ta = rows[a][0] instanceof Date ? rows[a][0].getTime() : Infinity;
      var tb = rows[b][0] instanceof Date ? rows[b][0].getTime() : Infinity;
      return ta === tb ? a - b : ta - tb;
    });
    var keep = idx[0];
    var dups = idx.slice(1);
    dups.forEach(function(k){ dupIdx[k] = keep; });
    groups.push({serial:s.split("||")[0], keep:recOf_(rows[keep], keep+1), dups:dups.map(function(k){ return recOf_(rows[k], k+1); })});
  });
  var uniq = {};
  order.forEach(function(k){ uniq[k.split("||")[0]] = true; });
  return {groups:groups, dupIdx:dupIdx, unique:Object.keys(uniq).length};
}

function scanDuplicates(){
  var rows = getOrCreateSheet("Master",[]).getDataRange().getValues();
  var exc = readExceptions_();
  var f = findDuplicates_(rows, exc);
  var extra = Object.keys(f.dupIdx).length;
  var archived = getSpreadsheet().getSheetByName("Duplicates");
  return {ok:true, totalRecords: rows.length - 1, uniqueSerials: f.unique,
          duplicateSerials: f.groups.length, extraRows: extra,
          archivedRows: archived ? Math.max(archived.getLastRow() - 1, 0) : 0,
          exceptions: Object.keys(exc).map(function(k){ return exc[k]; }),
          groups: f.groups.slice(0, 500)};
}

function removeDuplicates(d){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var master = getOrCreateSheet("Master",[]);
    var rows = master.getDataRange().getValues();
    var f = findDuplicates_(rows, readExceptions_());
    var extra = Object.keys(f.dupIdx).length;
    if(extra === 0) return {ok:true, removed:0, message:"No duplicates found"};
    if(Number(d.expected) !== extra)
      return {ok:false, error:"The data changed since you checked (" + extra + " duplicate rows now). Check again before removing."};

    var width = rows[0].length;
    var arch = getSpreadsheet().getSheetByName("Duplicates");
    if(!arch){
      arch = getSpreadsheet().insertSheet("Duplicates");
      arch.appendRow(rows[0].concat(["RemovedAt","RemovedBy","KeptMasterRow"]));
      arch.setFrozenRows(1);
    }
    var keepRows = [rows[0]], moved = [], now = new Date(), by = String(d.by||"");
    // first pass: final row number of each kept record after cleanup
    var newRowOf = {}, n = 1;
    for(var i = 1; i < rows.length; i++){ if(!(i in f.dupIdx)){ n++; newRowOf[i] = n; } }
    for(var j = 1; j < rows.length; j++){
      if(j in f.dupIdx){
        var r = rows[j].slice(0, width);
        while(r.length < width) r.push("");
        moved.push(r.concat([now, by, newRowOf[f.dupIdx[j]] || ""]));
      } else keepRows.push(rows[j]);
    }
    // archive first, then rewrite Master in one write
    arch.getRange(arch.getLastRow()+1, 1, moved.length, moved[0].length).setValues(moved);
    master.getRange(1, 1, rows.length, width).clearContent();
    master.getRange(1, 1, keepRows.length, width).setValues(keepRows);
    SpreadsheetApp.flush();
    return {ok:true, removed: moved.length, remaining: keepRows.length - 1};
  } finally {
    lock.releaseLock();
  }
}

function searchRecords(d){
  var q = normSerial_(d.q);
  if(q.length < 3) return {ok:false, error:"Type at least 3 characters"};
  var out = [];
  function scan(sheet, status){
    if(!sheet) return;
    var rows = sheet.getDataRange().getValues();
    var remCol = rows.length ? rows[0].indexOf("RemovedAt") : -1;
    for(var i = 1; i < rows.length && out.length < 300; i++){
      var s = normSerial_(rows[i][4]);
      var lot = String(rows[i][3]||"").toUpperCase();
      if(s.indexOf(q) === -1 && lot.indexOf(q) === -1) continue;
      var rec = recOf_(rows[i], i+1);
      rec.status = status;
      if(remCol > -1 && rows[i][remCol] instanceof Date) rec.removedAt = rows[i][remCol].toISOString();
      out.push(rec);
    }
  }
  scan(getOrCreateSheet("Master",[]), "Active");
  scan(getSpreadsheet().getSheetByName("Duplicates"), "Removed duplicate");
  return {ok:true, records: out, truncated: out.length >= 300};
}

/****************************************************
 * BATCH FLOW: 1 BOM verification  ->  2 FAI (later)  ->  3 Sampling
 * BOM data lives in the BOM reference spreadsheet (BOM_SS_ID).
 * Each verification is written to BOM_VERIFY_LOG + LOT_BOM_DETAIL there,
 * in the same format as earlier verifications.
 ****************************************************/
function ensureStageHeaders_(sheet){
  var h = sheet.getRange(1, COL_BOM_STATUS, 1, 3).getValues()[0];
  if(h[0] === "") sheet.getRange(1, COL_BOM_STATUS).setValue("BOMStatus");
  if(h[1] === "") sheet.getRange(1, COL_BOM_VER).setValue("BOMVerId");
  if(h[2] === "") sheet.getRange(1, COL_FAI_STATUS).setValue("FAIStatus");
}
function bomSS_(){ return SpreadsheetApp.openById(BOM_SS_ID); }
function headerIndex_(row){
  var m = {};
  row.forEach(function(h, i){ m[String(h).trim().toLowerCase()] = i; });
  return m;
}
function bomControl_(ss){
  var sh = ss.getSheetByName("BOM_CONTROL");
  var out = {};
  if(!sh) return out;
  var v = sh.getDataRange().getValues();
  var h = headerIndex_(v[0]);
  for(var i = 1; i < v.length; i++){
    var m = String(v[i][h["model"]]||"").trim();
    if(!m) continue;
    out[m] = {model:m, status:String(v[i][h["status"]]||""), rev:String(v[i][h["rev"]]),
              lockedBy:String(v[i][h["locked_by"]]||""),
              lockedOn: v[i][h["locked_on"]] instanceof Date ? v[i][h["locked_on"]].toISOString() : String(v[i][h["locked_on"]]||""),
              lastEcn:String(v[i][h["last_ecn"]]||"")};
  }
  return out;
}

function getBomList(){
  try{
    var ss = bomSS_();
    var ctl = bomControl_(ss);
    var models = Object.keys(ctl).map(function(k){
      var c = ctl[k];
      var tab = ss.getSheetByName(k);
      c.hasTab = !!tab;
      c.parts = tab ? Math.max(tab.getLastRow() - 1, 0) : 0;
      return c;
    });
    return {ok:true, models:models};
  }catch(e){
    return {ok:false, error:"Cannot open the BOM sheet: " + e.message};
  }
}

function readBom_(ss, model){
  var tab = ss.getSheetByName(model);
  if(!tab) return null;
  var v = tab.getDataRange().getValues();
  var h = headerIndex_(v[0]);
  var parts = [];
  for(var i = 1; i < v.length; i++){
    var pn = String(v[i][h["part_no"]]||"").trim();
    if(!pn) continue;
    var altRaw = h["alternate part"] !== undefined ? String(v[i][h["alternate part"]]||"") : "";
    parts.push({
      partNo: pn,
      description: String(v[i][h["part_description"]]||""),
      imageUrl: String(v[i][h["image_url"]]||""),
      alternates: altRaw.split(/[,;\/\n]+/).map(function(x){ return x.trim(); }).filter(String),
      supplier: String(v[i][h["supplier"]]||""),
      uom: String(v[i][h["uom"]]||""),
      qps: Number(v[i][h["qps"]])
    });
  }
  return parts;
}

function getBom(d){
  var model = String(d.model||"").trim();
  if(!model) return {ok:false, error:"No BOM model given. Set the Lumio model number for this model."};
  try{
    var ss = bomSS_();
    var parts = readBom_(ss, model);
    if(!parts) return {ok:false, error:"No BOM tab named \"" + model + "\" in the BOM sheet."};
    var ctl = bomControl_(ss)[model] || {model:model, status:"", rev:""};
    return {ok:true, model:model, control:ctl, parts:parts};
  }catch(e){
    return {ok:false, error:"Cannot open the BOM sheet: " + e.message};
  }
}

function nextVerId_(log){
  var v = log.getDataRange().getValues(), max = 0;
  for(var i = 1; i < v.length; i++){
    var m = String(v[i][0]||"").match(/(\d+)\s*$/);
    if(m) max = Math.max(max, Number(m[1]));
  }
  return "VER-" + ("0000" + (max + 1)).slice(-4);
}

function setBomStatus_(batchName, status, verId){
  var sh = getOrCreateSheet("Batches",[]);
  ensureStageHeaders_(sh);
  var v = sh.getDataRange().getValues();
  var n = 0;
  for(var i = 1; i < v.length; i++){
    if(String(v[i][0]) !== String(batchName)) continue;
    sh.getRange(i+1, COL_BOM_STATUS, 1, 2).setValues([[status, verId||""]]);
    n++;
  }
  return n;
}

function submitBomVerification(d){
  var batchName = String(d.batchName||"").trim();
  var model = String(d.model||"").trim();
  var items = d.items || [];
  if(!batchName || !model) return {ok:false, error:"Batch and BOM model are required"};

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    // batch must exist and still need BOM verification
    var bv = getOrCreateSheet("Batches",[]).getDataRange().getValues(), cur = null;
    for(var b = 1; b < bv.length; b++){ if(String(bv[b][0]) === batchName){ cur = String(bv[b][COL_BOM_STATUS-1]||""); break; } }
    if(cur === null) return {ok:false, error:"Batch " + batchName + " not found"};
    if(cur === "Passed") return {ok:false, error:"BOM verification for this batch has already passed."};

    var ss = bomSS_();
    var parts = readBom_(ss, model);
    if(!parts) return {ok:false, error:"No BOM tab named \"" + model + "\" in the BOM sheet."};
    var ctl = bomControl_(ss)[model] || {status:"", rev:""};
    if(ctl.status && ctl.status.toLowerCase() !== "locked")
      return {ok:false, error:"The BOM for " + model + " is " + ctl.status + " (being changed). Lock it before verifying."};

    var byPart = {};
    items.forEach(function(it){ byPart[String(it.partNo||"").trim()] = it; });
    var missing = parts.filter(function(p){ return !byPart[p.partNo]; });
    if(missing.length) return {ok:false, error:missing.length + " BOM parts were not checked (BOM may have changed). Reload the BOM and check again."};

    var rows = [], dev = 0;
    parts.forEach(function(p){
      var it = byPart[p.partNo];
      var lotPn = String(it.lotPartNo||"").trim();
      var lotQ = Number(it.lotQps);
      var okPn = lotPn === p.partNo || p.alternates.indexOf(lotPn) > -1;
      var okQ = !isNaN(lotQ) && Math.abs(lotQ - p.qps) < 1e-9;
      // same result words as the Part DB lot-wise verification
      var res = !lotPn ? "Not used in lot" : !okPn ? "Part changed" : !okQ ? "QPS changed" : "Matched";
      if(res !== "Matched"){
        dev++;
        if(!String(it.remark||"").trim()) throw new Error("Remark is required for deviation on " + p.partNo);
      }
      rows.push({p:p, lotPn:lotPn, lotQ:isNaN(lotQ) ? "" : lotQ, remark:String(it.remark||""), res:res});
    });

    var log = ss.getSheetByName("BOM_VERIFY_LOG");
    var det = ss.getSheetByName("LOT_BOM_DETAIL");
    if(!log || !det) return {ok:false, error:"BOM_VERIFY_LOG or LOT_BOM_DETAIL tab missing in the BOM sheet"};
    var verId = nextVerId_(log);
    var when = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    var result = dev === 0 ? "All matched" : "Deviation";
    log.appendRow([verId, when, model, batchName, ctl.rev, String(ctl.status).toLowerCase() === "locked" ? "Locked" : "Not locked", parts.length, dev, result,
                   String(d.by||""), String(d.remarks||"")]);
    var detail = rows.map(function(r){
      return [verId, model, batchName, ctl.rev, r.p.partNo, r.p.description, r.p.imageUrl, r.p.qps,
              r.lotPn, r.lotQ, r.remark, r.res];
    });
    det.getRange(det.getLastRow()+1, 1, detail.length, detail[0].length).setValues(detail);

    var status = dev === 0 ? "Passed" : "Failed";
    setBomStatus_(batchName, status, verId);
    return {ok:true, verId:verId, status:status, partsChecked:parts.length, deviations:dev, result:result};
  }catch(e){
    return {ok:false, error:String(e.message||e)};
  }finally{
    lock.releaseLock();
  }
}

function getBomVerification(d){
  try{
    var ss = bomSS_();
    var log = ss.getSheetByName("BOM_VERIFY_LOG"), det = ss.getSheetByName("LOT_BOM_DETAIL");
    if(!log || !det) return {ok:false, error:"Verification tabs missing in the BOM sheet"};
    var want = String(d.verId||""), batch = String(d.batchName||"");
    var lv = log.getDataRange().getValues(), head = null;
    for(var i = lv.length - 1; i >= 1; i--){
      if((want && String(lv[i][0]) === want) || (!want && batch && String(lv[i][3]) === batch)){
        head = {verId:String(lv[i][0]), time: lv[i][1] instanceof Date ? Utilities.formatDate(lv[i][1], TZ, "yyyy-MM-dd HH:mm:ss") : String(lv[i][1]),
                model:String(lv[i][2]), lot:String(lv[i][3]), rev:String(lv[i][4]), bomStatus:String(lv[i][5]),
                partsChecked:Number(lv[i][6]), deviations:Number(lv[i][7]), result:String(lv[i][8]),
                by:String(lv[i][9]), remarks:String(lv[i][10]||"")};
        break;
      }
    }
    if(!head) return {ok:true, found:false};
    var dv = det.getDataRange().getValues(), items = [];
    for(var j = 1; j < dv.length; j++){
      if(String(dv[j][0]) !== head.verId) continue;
      items.push({partNo:String(dv[j][4]), description:String(dv[j][5]), imageUrl:String(dv[j][6]),
                  bomQps:dv[j][7], lotPartNo:String(dv[j][8]), lotQps:dv[j][9], remark:String(dv[j][10]||""), result:String(dv[j][11])});
    }
    return {ok:true, found:true, verification:head, items:items};
  }catch(e){
    return {ok:false, error:"Cannot open the BOM sheet: " + e.message};
  }
}

/* Admin: change a batch's BOM step (Waived = not required, Pending = require verification) */
function setBatchStage(d){
  var st = String(d.bomStatus||"");
  if(["Waived","Pending"].indexOf(st) === -1) return {ok:false, error:"Unsupported status"};
  var n = setBomStatus_(String(d.batchName||""), st, st === "Waived" ? ("Waived by " + String(d.by||"admin")) : "");
  return n ? {ok:true, lots:n} : {ok:false, error:"Batch not found"};
}
