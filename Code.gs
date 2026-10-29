/****************************************************
 * LUMIO Quality Management - Projector — Google Apps Script backend
 * v4: Models + Black Dot spec are stored in the "Models" sheet
 *     and managed by admin from the app (nothing hardcoded).
 ****************************************************/

var SS_ID        = "1WC6WGmwiqqytYobfoM9P4dP1RIptj5Rj17LSYtP_8tI";
var FIXED_LOT    = 90;
var FIXED_SAMPLE = 13;

// All data lives in ONE spreadsheet (SS_ID): inspection tabs and the BOM tabs (model BOMs, PART_MASTER,
// BOM_CONTROL, ECN_HISTORY, BOM_VERIFY_LOG, LOT_BOM_DETAIL). The old Lumio_Part_Reference_DB is only read
// until migrateBomIntoPortal() has copied its tabs in, and stays as a backup afterwards.
var LEGACY_BOM_SS_ID = "11BFVHMfnDGymLNYfBNLp9CRKIAE0X5fLX8qX-R8DeYk";
var BOM_TABS_SYSTEM  = ["PART_MASTER","BOM_CONTROL","ECN_HISTORY","BOM_VERIFY_LOG","LOT_BOM_DETAIL","LOT_BOM_VIEW","BOM_COMPARE","MODEL_COMPARE"];
var TZ           = "Asia/Kolkata";

// Part DB web app (BOM management tools, bound to the BOM sheet). Admins open it through the portal.
var PARTDB_URL   = "https://script.google.com/macros/s/AKfycbyPiE-gcrUZH7cEA6KKBgXHJa93MUl2r7-GseHkZ0FGdKBmTJrOnl7r7SOpm3SNjdjZ/exec";

// Login sessions (sheet "Sessions"; the Part DB script reads it too)
var SESSION_HOURS = 12;
// Modules an admin can give each user (Users sheet, column E "Access", comma separated).
// Admins always have every module. An empty cell means the original inspector set.
var MODULES        = ["bomverify","fai","sampling","reports","createbatch","models","data","bommgmt"];
var DEFAULT_ACCESS = ["bomverify","fai","sampling"];
// Actions any signed-in user may call (reading, own password)
// Actions that only read; every other successful action clears the cache
var READ_ACTIONS = ["login","logout","whoami","bootstrap","getBatches","getModels","modelInsights","getBomList","getBom","getBomVerification",
                    "getLotUnits","checkSerial","getFaiSpec","getFai","listUsers","scanDuplicates","searchRecords"];
var ALWAYS_ACTIONS = ["bootstrap","modelInsights","logout","whoami","getBatches","getModels","getBomList","getBomVerification","getLotUnits","changePassword","getFaiSpec","getFai","getBom"];
// recordRework (rework data + reoffer) belongs to the sampling module
// Actions unlocked by a module; everything not listed here or above is admin only
var ACTION_MODULE  = {getBom:"bomverify", submitBomVerification:"bomverify",
                      checkSerial:"sampling", submitLot:"sampling", recordRework:"sampling", setOfferedQty:"sampling",
                      saveFai:"fai", saveFaiSpec:"models",
                      createBatch:"createbatch",
                      saveModel:"models", deleteModel:"models",
                      scanDuplicates:"data", searchRecords:"data", removeDuplicates:"data", approveSerial:"data",
                      bom:"bommgmt"};

// Security settings
var PW_ITERATIONS    = 1000;    // PBKDF2-HMAC-SHA256 rounds (each round is slow in Apps Script: 1000 = about 1.5 s per login)
var PW_MIN_LENGTH    = 8;
var LOGIN_MAX_FAILS  = 5;       // failed logins per username ...
var LOGIN_LOCK_SECS  = 900;     // ... locks that username for 15 minutes
var RATE_PER_MINUTE  = 120;     // API calls per signed-in session per minute
var MAX_BODY_CHARS   = 500000;  // largest request accepted

// Batches sheet stage columns (1-based)
var COL_BOM_STATUS = 16;   // "", Pending, Passed, Failed, Waived  ("" = created before the BOM step)
var COL_BOM_VER    = 17;   // Ver_ID of the verification that set the status
var COL_FAI_STATUS = 18;   // reserved for FAI inspection
var COL_ROUND      = 19;   // inspection round of the lot: 1 = first offer, 2+ = reoffer after rework
var COL_REOFFER    = 20;   // serial numbers picked for the current reoffer (comma separated)
var MASTER_ROUND_COL = 20; // Master: round the sample row belongs to (blank = 1)

// Lot status (Batches column F): Open -> (defect) Hold -> (rework saved) Reoffered -> ... -> Closed
var LOT_STATUSES   = ["Open","Hold","Reoffered","Closed"];
var CHECKPOINTS    = ["Aesthetic","General","Functional","Screen","Black dot"];
var UNIT_STATUSES  = ["Reworked OK","Scrapped"];

var SS_CACHE_ = null;
function getSpreadsheet(){
  if(!SS_CACHE_) SS_CACHE_ = SpreadsheetApp.openById(SS_ID);
  return SS_CACHE_;
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
  // No default accounts are created. Add the first admin by hand in the Users sheet
  // (any password; it is hashed automatically at the first login).
  return "Setup done";
}

function doGet(e){
  return ContentService
    .createTextOutput(JSON.stringify({ok:true, msg:"LUMIO Quality Management - Projector API running", version:"v5"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try {
    var raw  = (e && e.postData && e.postData.contents) || "";
    if(raw.length > MAX_BODY_CHARS) return json_({ok:false, error:"Request is too large."});
    var data;
    try { data = JSON.parse(raw); } catch(pe) { return json_({ok:false, error:"Bad request."}); }
    if(!data || typeof data !== "object" || Array.isArray(data)) return json_({ok:false, error:"Bad request."});
    var action = String(data.action || "");
    var result = {ok:false, error:"Unknown action: " + action};

    // Every action except login needs a valid session; most need the admin role.
    if(action !== "login"){
      var sess = sessionFor_(data.token);
      if(!sess) return json_({ok:false, auth:true, error:"Your session has expired. Please log in again."});
      if(!rateOk_("r_" + sha256_(data.token).slice(0, 32), RATE_PER_MINUTE, 60))
        return json_({ok:false, error:"Too many requests. Wait a minute and try again."});
      if(!allowed_(sess, action))
        return json_({ok:false, error: ACTION_MODULE[action] ? "You do not have access to this. Ask an admin to give you the " + moduleName_(ACTION_MODULE[action]) + " module." : "Only an admin can do this."});
      data._session = sess;
      // record who did it from the session, not from what the browser sent
      if(action === "submitLot" || action === "recordRework" || action === "holdLot" || action === "setOfferedQty" || action === "saveFai" || action === "saveFaiSpec") data.inspector = sess.name;
      if(action === "submitBomVerification") data.by = sess.name;
      if(["createBatch","saveModel","removeDuplicates","approveSerial","setBatchStage","resetPassword","deleteUser","setUserAccess"].indexOf(action) > -1){
        data.createdBy = sess.username; data.updatedBy = sess.username; data.by = sess.username;
      }
    }

    if(action === "login"){
      result = login(data);
    }
    if(action === "addUser"){
      result = addUser(data);
    }
    if(action === "listUsers"){
      result = listUsers();
    }
    if(action === "setUserAccess"){
      result = setUserAccess(data);
    }
    if(action === "whoami"){
      result = whoami(data._session);
    }
    if(action === "bootstrap"){
      result = bootstrap(data);
    }
    if(action === "modelInsights"){
      result = modelInsights();
    }
    if(action === "resetPassword"){
      result = resetPassword(data);
    }
    if(action === "deleteUser"){
      result = deleteUser(data);
    }
    if(action === "changePassword"){
      result = changePassword(data);
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
    if(action === "getModels"){
      result = getModels();
    }
    if(action === "saveModel"){
      result = saveModel(data);
    }
    if(action === "deleteModel"){
      result = deleteModel(data);
    }
    if(action === "recordRework"){
      result = recordRework(data);
    }
    if(action === "getFaiSpec"){
      result = getFaiSpec(data);
    }
    if(action === "saveFaiSpec"){
      result = saveFaiSpec(data);
    }
    if(action === "getFai"){
      result = getFai(data);
    }
    if(action === "saveFai"){
      result = saveFai(data);
    }
    if(action === "setOfferedQty"){
      result = setOfferedQty(data);
    }
    if(action === "holdLot"){
      result = holdLot(data);
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
    if(action === "bom"){
      result = bomCall(data);
    }

    // any change clears the cached lists so the next read is fresh
    if(result && result.ok && (action === "bom" ? result.write : READ_ACTIONS.indexOf(action) === -1)) dropCache_();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    // details go to the Apps Script execution log only, never to the browser
    console.error("doPost failed: " + (err && err.stack || err));
    return json_({ok:false, error:"Server error. Please try again, or ask an admin to check the Apps Script log."});
  }
}

function login(d){
  var username = String(d.username || "").trim();
  var password = String(d.password || "");
  if(!username || !password || username.length > 64 || password.length > 200)
    return {ok:false, error:"Invalid username or password"};
  var failKey = "lf_" + sha256_(username.toLowerCase()).slice(0, 40);
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get(failKey) || 0);
  if(fails >= LOGIN_MAX_FAILS)
    return {ok:false, error:"Too many failed attempts for this username. Try again in 15 minutes."};

  migratePasswords_();
  var rows = getOrCreateSheet("Users",[]).getDataRange().getValues();
  var row = null;
  for(var i = 1; i < rows.length; i++){
    if(String(rows[i][0]).trim() === username){ row = rows[i]; break; }
  }
  // the same amount of work whether or not the user exists, so timing does not reveal usernames
  var ok = row ? verifyPassword_(password, String(row[1])) : (verifyPassword_(password, DUMMY_HASH_()) && false);
  if(!ok){
    cache.put(failKey, String(fails + 1), LOGIN_LOCK_SECS);
    return {ok:false, error:"Invalid username or password"};
  }
  cache.remove(failKey);
  // a password typed into the sheet by hand, or hashed with a different round count: store it hashed now
  if(!isHashed_(String(row[1])) || String(row[1]).split("$")[1] !== String(PW_ITERATIONS)){
    var ush = getOrCreateSheet("Users",[]), uv = ush.getDataRange().getValues(), ui = findUserRow_(uv, username);
    if(ui > -1) ush.getRange(ui + 1, 2).setNumberFormat("@").setValue(hashPassword_(password));
  }
  var role = String(row[2]||"inspector").trim().toLowerCase() === "admin" ? "admin" : "inspector";
  var access = parseAccess_(row[4], role);
  var s = createSession_(String(row[0]).trim(), String(row[3]||row[0]), role, access);
  var out = whoami({username:String(row[0]).trim(), name:String(row[3]||row[0]), role:role, access:access});
  out.token = s.token; out.expiresAt = s.expiresAt;
  return out;
}

/* ---------------- passwords ----------------
 * Stored as  pbkdf2$<iterations>$<salt base64>$<hash base64>  (PBKDF2-HMAC-SHA256, 32-byte key).
 * Plain-text passwords left from before are converted the first time anyone logs in. */
function toArr_(b){ return Array.prototype.slice.call(b); }
function bytesOf_(str){ return toArr_(Utilities.newBlob(String(str)).getBytes()); }
function pbkdf2_(password, saltBytes, iterations){
  var pw = bytesOf_(password);
  var u = toArr_(Utilities.computeHmacSha256Signature(toArr_(saltBytes).concat([0, 0, 0, 1]), pw));
  var t = u.slice();
  for(var i = 1; i < iterations; i++){
    u = toArr_(Utilities.computeHmacSha256Signature(u, pw));
    for(var j = 0; j < t.length; j++) t[j] ^= u[j];
  }
  return t.map(function(b){ return ((b << 24) >> 24); });   // keep signed bytes for base64Encode
}
function randomBytes_(n){
  var hex = "";
  while(hex.length < n * 2) hex += Utilities.getUuid().replace(/-/g, "");
  var out = [];
  for(var i = 0; i < n; i++) out.push(((parseInt(hex.substr(i * 2, 2), 16)) << 24) >> 24);
  return out;
}
function hashPassword_(password){
  var salt = randomBytes_(16);
  return "pbkdf2$" + PW_ITERATIONS + "$" + Utilities.base64Encode(salt) + "$" +
         Utilities.base64Encode(pbkdf2_(password, salt, PW_ITERATIONS));
}
function isHashed_(stored){ return /^pbkdf2\$\d+\$[^$]+\$[^$]+$/.test(String(stored)); }
function sameText_(a, b){
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  for(var i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}
function verifyPassword_(password, stored){
  stored = String(stored || "");
  if(!isHashed_(stored)) return stored !== "" && sameText_(password, stored);   // not converted yet
  var p = stored.split("$");
  var iters = Number(p[1]);
  if(!(iters >= 1 && iters <= 1000000)) return false;
  var calc = Utilities.base64Encode(pbkdf2_(password, Utilities.base64Decode(p[2]), iters));
  return sameText_(calc, p[3]);
}
var DUMMY_HASH_CACHE_ = null;
function DUMMY_HASH_(){
  if(!DUMMY_HASH_CACHE_) DUMMY_HASH_CACHE_ = "pbkdf2$" + PW_ITERATIONS + "$AAAAAAAAAAAAAAAAAAAAAA==$AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA=";
  return DUMMY_HASH_CACHE_;
}
/* Converts any plain-text password in the Users sheet to a hash (only does work when one is found). */
function migratePasswords_(){
  var sh = getOrCreateSheet("Users",[]);
  var pending = sh.getDataRange().getValues().some(function(r, i){
    return i > 0 && String(r[1] == null ? "" : r[1]) !== "" && !isHashed_(String(r[1]));
  });
  if(!pending) return 0;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var v = sh.getDataRange().getValues(), n = 0;
    for(var i = 1; i < v.length; i++){
      var cur = String(v[i][1] == null ? "" : v[i][1]);
      if(cur !== "" && !isHashed_(cur)){
        sh.getRange(i + 1, 2).setNumberFormat("@").setValue(hashPassword_(cur));
        n++;
      }
    }
    return n;
  } finally { lock.releaseLock(); }
}

/* ---------------- cache: batches, models and BOM list are read on every page load ---------------- */
function cached_(key, secs, fresh, fn){
  var cache = CacheService.getScriptCache();
  if(!fresh){ var hit = cache.get(key); if(hit){ try { return JSON.parse(hit); } catch(e){} } }
  var val = fn();
  try { var j = JSON.stringify(val); if(val && val.ok && j.length < 95000) cache.put(key, j, secs); } catch(e){}
  return val;
}
function dropCache_(){ CacheService.getScriptCache().removeAll(["c_batches","c_models","c_bomlist"]); }
/* Everything the portal needs to start, in one request. fresh = skip the cache (Refresh button). */
function bootstrap(d){
  var fresh = !!d.fresh;
  return {ok:true, me:whoami(d._session),
          batches:cached_("c_batches", 600, fresh, getBatches),
          models:cached_("c_models", 600, fresh, getModels),
          bomList:cached_("c_bomlist", 300, fresh, getBomList)};
}

/* ---------------- request limits + input safety ---------------- */
function rateOk_(key, max, windowSecs){
  var cache = CacheService.getScriptCache();
  var bucket = key + "_" + Math.floor(Date.now() / (windowSecs * 1000));
  var n = Number(cache.get(bucket) || 0) + 1;
  cache.put(bucket, String(n), windowSecs + 5);
  return n <= max;
}
/* Text going into a sheet cell: never let it start a formula (=, +, -, @), cap its length. */
function safe_(v, maxLen){
  if(v === null || v === undefined) return "";
  if(typeof v === "number" || typeof v === "boolean" || v instanceof Date) return v;
  var t = String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  if(maxLen && t.length > maxLen) t = t.substring(0, maxLen);
  if(/^[=+\-@\t\r]/.test(t)) t = "'" + t;
  return t;
}
var NAME_RE_ = /^[A-Za-z0-9][A-Za-z0-9 _.\-\/()#&]{0,79}$/;

function json_(o){
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------- module access ---------------- */
function parseAccess_(cell, role){
  if(role === "admin") return MODULES.slice();
  var t = String(cell == null ? "" : cell).trim().toLowerCase();
  if(t === "") return DEFAULT_ACCESS.slice();
  if(t === "none") return [];
  return t.split(/[\s,;]+/).filter(function(m){ return MODULES.indexOf(m) > -1; });
}
function allowed_(sess, action){
  if(sess.role === "admin") return true;
  if(ALWAYS_ACTIONS.indexOf(action) > -1) return true;
  var m = ACTION_MODULE[action];
  return !!m && (sess.access || []).indexOf(m) > -1;
}
function moduleName_(m){
  return {bomverify:"BOM verification", fai:"FAI inspection", sampling:"Sampling inspection", reports:"Reports", createbatch:"Create batches",
          models:"Models and black dot spec", data:"Data records", bommgmt:"BOM management"}[m] || m;
}
/* Who is signed in, and what they may use (the portal calls this on every refresh). */
function whoami(sess){
  var out = {ok:true, username:sess.username, name:sess.name, role:sess.role, access:sess.access || []};
  if(sess.role === "admin" || out.access.indexOf("bommgmt") > -1){
    out.partdbUrl = PARTDB_URL; out.bomSheetUrl = "https://docs.google.com/spreadsheets/d/" + (bomMigrated_() ? SS_ID : LEGACY_BOM_SS_ID) + "/edit";
  }
  return out;
}

/* ---------------- sessions ----------------
 * Sessions sheet: TokenHash | Username | Name | Role | CreatedAt | ExpiresAt | Access
 * Only a SHA-256 hash of each token is stored. */
function sha256_(s){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function(b){ return ("0" + (b & 0xff).toString(16)).slice(-2); }).join("");
}
function sessionsSheet_(){
  var sh = getSpreadsheet().getSheetByName("Sessions");
  if(!sh){
    sh = getSpreadsheet().insertSheet("Sessions");
    sh.appendRow(["TokenHash","Username","Name","Role","CreatedAt","ExpiresAt","Access"]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function createSession_(username, name, role, access){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var sh = sessionsSheet_();
    var now = new Date();
    // drop expired sessions
    var v = sh.getDataRange().getValues();
    var w = v[0].length;
    if(String(v[0][6] || "") === ""){ sh.getRange(1, 7).setValue("Access"); }
    var keep = [v[0]];
    for(var i = 1; i < v.length; i++){ if(v[i][5] instanceof Date && v[i][5] > now) keep.push(v[i]); }
    if(keep.length !== v.length){
      sh.getRange(1, 1, v.length, w).clearContent();
      sh.getRange(1, 1, keep.length, w).setValues(keep);
      if(String(v[0][6] || "") === "") sh.getRange(1, 7).setValue("Access");
    }
    var token = Utilities.getUuid().replace(/-/g,"") + Utilities.getUuid().replace(/-/g,"");
    var exp = new Date(now.getTime() + SESSION_HOURS * 3600 * 1000);
    sh.appendRow([sha256_(token), username, name, role, now, exp, (access && access.length) ? access.join(",") : "none"]);
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
    var role = String(v[i][3]).toLowerCase() === "admin" ? "admin" : "inspector";
    var s = {username:String(v[i][1]), name:String(v[i][2]), role:role, access:parseAccess_(v[i][6], role)};
    cache.put("s_" + h, JSON.stringify(s), 300);
    return s;
  }
  return null;
}
/* Ends every session of one user (password reset, user removed). */
function revokeSessions_(username){
  var sh = sessionsSheet_(), v = sh.getDataRange().getValues(), cache = CacheService.getScriptCache();
  for(var i = 1; i < v.length; i++){
    if(String(v[i][1]) === String(username) && v[i][5] instanceof Date && v[i][5] > new Date()){
      sh.getRange(i+1, 6).setValue(new Date(0));
      cache.remove("s_" + String(v[i][0]));
    }
  }
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

function findUserRow_(v, username){
  for(var i = 1; i < v.length; i++) if(String(v[i][0]).trim().toLowerCase() === String(username).trim().toLowerCase()) return i;
  return -1;
}
function pwProblem_(pw){
  pw = String(pw || "");
  if(pw.length < PW_MIN_LENGTH) return "Password must be at least " + PW_MIN_LENGTH + " characters";
  if(pw.length > 200) return "Password is too long";
  return "";
}
function addUser(d){
  var username = String(d.username || "").trim(), name = String(d.name || "").trim();
  var role = String(d.role || "inspector").toLowerCase();
  if(!/^[A-Za-z0-9._-]{3,32}$/.test(username)) return {ok:false, error:"Username: 3 to 32 letters, numbers, dot, dash or underscore"};
  if(!name || name.length > 80) return {ok:false, error:"Full name is required (up to 80 characters)"};
  if(["inspector","admin"].indexOf(role) === -1) return {ok:false, error:"Role must be inspector or admin"};
  var p = pwProblem_(d.password); if(p) return {ok:false, error:p};
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var sh = getOrCreateSheet("Users",[]);
    if(findUserRow_(sh.getDataRange().getValues(), username) > -1) return {ok:false, error:"Username " + username + " already exists"};
    var acc = Array.isArray(d.access) ? MODULES.filter(function(m){ return d.access.indexOf(m) > -1; }) : DEFAULT_ACCESS.slice();
    if(String(sh.getRange(1, 5).getValue() || "") === "") sh.getRange(1, 5).setValue("Access");
    sh.appendRow([username, hashPassword_(String(d.password)), role, safe_(name, 80), acc.length ? acc.join(",") : "none"]);
    sh.getRange(sh.getLastRow(), 2).setNumberFormat("@");
    return {ok:true};
  } finally { lock.releaseLock(); }
}
function listUsers(){
  var v = getOrCreateSheet("Users",[]).getDataRange().getValues(), out = [];
  for(var i = 1; i < v.length; i++){
    if(String(v[i][0]).trim() === "") continue;
    var role = String(v[i][2]||"inspector").trim().toLowerCase() === "admin" ? "admin" : "inspector";
    out.push({username:String(v[i][0]).trim(), name:String(v[i][3]||""), role:role, access:parseAccess_(v[i][4], role)});
  }
  return {ok:true, users:out, modules:MODULES};
}
/* Admin: set a user's role and modules. Signed-in sessions of that user change at once. */
function setUserAccess(d){
  var role = String(d.role || "inspector").toLowerCase();
  if(["inspector","admin"].indexOf(role) === -1) return {ok:false, error:"Role must be inspector or admin"};
  var list = Array.isArray(d.access) ? d.access.map(String) : [];
  for(var k = 0; k < list.length; k++) if(MODULES.indexOf(list[k]) === -1) return {ok:false, error:"Unknown module " + list[k]};
  var access = MODULES.filter(function(m){ return list.indexOf(m) > -1; });
  var me = d._session.username;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var sh = getOrCreateSheet("Users",[]), v = sh.getDataRange().getValues();
    if(String(v[0][4] || "") === "") sh.getRange(1, 5).setValue("Access");
    var i = findUserRow_(v, d.username);
    if(i < 0) return {ok:false, error:"User not found"};
    var uname = String(v[i][0]).trim();
    var oldRole = String(v[i][2]||"").trim().toLowerCase() === "admin" ? "admin" : "inspector";
    if(uname.toLowerCase() === String(me).toLowerCase() && role !== "admin") return {ok:false, error:"You cannot remove your own admin role"};
    if(oldRole === "admin" && role !== "admin"){
      var admins = 0;
      for(var r = 1; r < v.length; r++) if(String(v[r][2]||"").trim().toLowerCase() === "admin" && String(v[r][0]).trim() !== "") admins++;
      if(admins < 2) return {ok:false, error:"There must be at least one admin"};
    }
    sh.getRange(i + 1, 3).setValue(role);
    sh.getRange(i + 1, 5).setValue(access.length ? access.join(",") : "none");
    // update live sessions of this user
    var ss = sessionsSheet_(), sv = ss.getDataRange().getValues(), cache = CacheService.getScriptCache();
    for(var j = 1; j < sv.length; j++){
      if(String(sv[j][1]) !== uname) continue;
      ss.getRange(j + 1, 4).setValue(role);
      ss.getRange(j + 1, 7).setValue(access.length ? access.join(",") : "none");
      cache.remove("s_" + String(sv[j][0]));
    }
    return {ok:true, role:role, access:parseAccess_(access.length ? access.join(",") : "none", role)};
  } finally { lock.releaseLock(); }
}
function resetPassword(d){
  var p = pwProblem_(d.password); if(p) return {ok:false, error:p};
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var sh = getOrCreateSheet("Users",[]), v = sh.getDataRange().getValues();
    var i = findUserRow_(v, d.username);
    if(i < 0) return {ok:false, error:"User not found"};
    sh.getRange(i + 1, 2).setNumberFormat("@").setValue(hashPassword_(String(d.password)));
    revokeSessions_(String(v[i][0]).trim());
    return {ok:true};
  } finally { lock.releaseLock(); }
}
function deleteUser(d){
  var me = d._session ? d._session.username : "";
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var sh = getOrCreateSheet("Users",[]), v = sh.getDataRange().getValues();
    var i = findUserRow_(v, d.username);
    if(i < 0) return {ok:false, error:"User not found"};
    var uname = String(v[i][0]).trim();
    if(uname.toLowerCase() === String(me).toLowerCase()) return {ok:false, error:"You cannot remove your own account"};
    sh.deleteRow(i + 1);
    revokeSessions_(uname);
    return {ok:true};
  } finally { lock.releaseLock(); }
}
function changePassword(d){
  var me = d._session.username;
  var p = pwProblem_(d.newPassword); if(p) return {ok:false, error:p};
  var failKey = "cp_" + sha256_(me.toLowerCase()).slice(0, 40), cache = CacheService.getScriptCache();
  var fails = Number(cache.get(failKey) || 0);
  if(fails >= LOGIN_MAX_FAILS) return {ok:false, error:"Too many wrong attempts. Try again in 15 minutes."};
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var sh = getOrCreateSheet("Users",[]), v = sh.getDataRange().getValues();
    var i = findUserRow_(v, me);
    if(i < 0) return {ok:false, error:"User not found"};
    if(!verifyPassword_(String(d.currentPassword || ""), String(v[i][1]))){
      cache.put(failKey, String(fails + 1), LOGIN_LOCK_SECS);
      return {ok:false, error:"Current password is wrong"};
    }
    cache.remove(failKey);
    sh.getRange(i + 1, 2).setNumberFormat("@").setValue(hashPassword_(String(d.newPassword)));
    revokeSessions_(String(v[i][0]).trim());
    var s = createSession_(String(v[i][0]).trim(), d._session.name, d._session.role, d._session.access);
    return {ok:true, token:s.token, expiresAt:s.expiresAt};
  } finally { lock.releaseLock(); }
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
      faiStatus     : String(r[COL_FAI_STATUS-1]||""),
      round         : Number(r[COL_ROUND-1]) || 1,
      reofferSerials: String(r[COL_REOFFER-1]||"").split(",").map(function(x){ return x.trim(); }).filter(String)
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
  var by        = safe_(String(d.createdBy||""), 64);

  if(!batchName) return {ok:false, error:"Batch Name is required"};
  if(!NAME_RE_.test(batchName)) return {ok:false, error:"Batch Name: up to 80 letters, numbers, spaces and _ - . / ( ) # &, starting with a letter or number"};
  if(batchQty < 1 || batchQty > 1000000) return {ok:false, error:"Batch Quantity must be between 1 and 1000000"};
  if(pm && !/^\d{4}$/.test(pm)) return {ok:false, error:"Production month is not valid"};
  if(pm2 && !/^\d{4}$/.test(pm2)) return {ok:false, error:"Production month 2 is not valid"};
  var existing = getOrCreateSheet("Batches",[]).getDataRange().getValues();
  for(var x = 1; x < existing.length; x++)
    if(String(existing[x][0]).trim().toLowerCase() === batchName.toLowerCase()) return {ok:false, error:"A batch named " + batchName + " already exists"};
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
      if(d.status !== undefined && LOT_STATUSES.indexOf(String(d.status)) === -1) return {ok:false, error:"Status must be Open, Hold, Reoffered or Closed"};
      if(d.bdMax !== undefined && (isNaN(Number(d.bdMax)) || Number(d.bdMax) < 0)) return {ok:false, error:"Black dot limit is not valid"};
      if(d.model !== undefined && !findModel_(d.model)) return {ok:false, error:"Model not found"};
      if(d.prodMonthCode !== undefined && !/^'?\d{0,4}$/.test(String(d.prodMonthCode))) return {ok:false, error:"Production month is not valid"};
      if(d.prodMonthCode2 !== undefined && !/^'?\d{0,4}$/.test(String(d.prodMonthCode2))) return {ok:false, error:"Production month 2 is not valid"};
      if(d.status        !== undefined) sheet.getRange(i+1, 6).setValue(String(d.status));
      if(d.bdMax         !== undefined) sheet.getRange(i+1, 5).setValue(Number(d.bdMax));
      if(d.model         !== undefined) sheet.getRange(i+1, 12).setValue(safe_(d.model, 80));
      if(d.prodMonthCode !== undefined) sheet.getRange(i+1, 13).setValue(safe_(d.prodMonthCode, 5));
      if(d.prodMonthCode2!== undefined) sheet.getRange(i+1, 15).setValue(safe_(d.prodMonthCode2, 5));
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
  var sn = normSerial_(d.serial);
  if(!sn) return {ok:true, exists:false};
  var bn = String(d.batchName||""), ln = String(d.lotNumber||"");
  var master = getOrCreateSheet("Master",[]);
  var last = master.getLastRow();
  if(last < 2) return {ok:true, exists:false};
  // search only the serial column for the exact value (much faster than reading the whole sheet)
  var hits = master.getRange(2, 5, last - 1, 1).createTextFinder(sn).matchEntireCell(true).findAll();
  var sameLot = false;
  for(var i = 0; i < hits.length; i++){
    var row = master.getRange(hits[i].getRow(), 3, 1, 2).getValues()[0];
    if(bn && String(row[0]) === bn && String(row[1]) === ln){ sameLot = true; continue; }
    return {ok:true, exists:true, batch:row[0], lot:row[1]};
  }
  return {ok:true, exists:false, sameLot:sameLot};
}

function submitLot(d){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{ return submitLotLocked_(d); } finally { lock.releaseLock(); }
}
function lotRowOf_(bRows, batchName, lotNumber){
  for(var i = 1; i < bRows.length; i++)
    if(String(bRows[i][0]) === String(batchName) && String(bRows[i][1]) === String(lotNumber)) return i;
  return -1;
}
function logLot_(batch, lot, round, event, details, by){
  var sh = getOrCreateSheet("LotHistory", ["Time","Batch","Lot","Round","Event","Details","By"]);
  sh.appendRow([new Date(), safe_(batch, 80), safe_(lot, 100), round, event, safe_(details, 1000), safe_(by, 80)]);
}
/* Any defective sample puts the whole lot on HOLD straight away (submitted with the samples inspected so far).
   PASS needs every sample inspected and none defective. */
function submitLotLocked_(d){
  var bsh   = getOrCreateSheet("Batches",[]);
  ensureStageHeaders_(bsh);
  var bRows = bsh.getDataRange().getValues();
  var lotRow = lotRowOf_(bRows, d.batchName, d.lotNumber);
  if(lotRow < 0) return {ok:false, error:"Lot not found. Refresh and try again."};
  var status = String(bRows[lotRow][5]);
  if(status === 'Closed') return {ok:false, error:'This lot is already submitted and closed. Duplicate submission blocked.'};
  if(status === 'Hold') return {ok:false, error:'This lot is on hold. Record the rework data to reoffer it.'};
  // Flow gate: sampling needs BOM verification passed (or waived) for batches created with the flow
  var bs = String(bRows[lotRow][COL_BOM_STATUS-1]||"");
  if(bs === "Pending" || bs === "Failed")
    return {ok:false, error:"BOM verification for " + d.batchName + " is " + (bs === "Pending" ? "not done yet" : "failed") + ". Sampling can start after BOM verification passes."};

  var lotResult = String(d.lotResult||"");
  if(lotResult !== "PASS" && lotResult !== "HOLD") return {ok:false, error:"Lot result must be PASS or HOLD"};
  var round = Number(bRows[lotRow][COL_ROUND-1]) || 1;
  var lotSize = Number(bRows[lotRow][2]) || FIXED_LOT;
  var sampleSize = Number(bRows[lotRow][3]) || FIXED_SAMPLE;
  var units = Array.isArray(d.units) ? d.units : [];
  var pf = function(v){ return String(v) === "Pass" ? "Pass" : "Fail"; };
  var defective = units.filter(function(u){ return pf(u && u.sampleResult) === "Fail"; }).length;
  if(lotResult === "PASS"){
    if(units.length !== sampleSize) return {ok:false, error:"Expected " + sampleSize + " samples, got " + units.length};
    if(defective) return {ok:false, error:"A defective sample puts the lot on hold; it cannot pass."};
    var need = String(bRows[lotRow][COL_REOFFER-1]||"").split(",").map(function(x){ return normSerial_(x); }).filter(String);
    var got = {}; units.forEach(function(u){ got[normSerial_(u.serial)] = true; });
    var missing = need.filter(function(x){ return !got[x]; });
    if(missing.length) return {ok:false, error:"The reoffer must include these reworked units: " + missing.join(", ")};
  } else {
    if(!defective) return {ok:false, error:"Hold needs at least one defective sample."};
    if(units.length < 1 || units.length > sampleSize) return {ok:false, error:"Between 1 and " + sampleSize + " samples expected"};
  }
  var fgPassQty = lotResult === "PASS" ? lotSize : 0;

  var master = getOrCreateSheet("Master",[]);
  if(master.getRange(1, 19).getValue() === "") master.getRange(1, 19).setValue("BDSpec");
  if(master.getRange(1, MASTER_ROUND_COL).getValue() === "") master.getRange(1, MASTER_ROUND_COL).setValue("Round");

  // Rule: a serial number is recorded only once, except when the same lot re-inspects it after rework
  var seen = {}, existing = {}, clash = [];
  var mRows = master.getDataRange().getValues();
  var thisBatch = String(bRows[lotRow][0]), thisLot = String(bRows[lotRow][1]);
  for(var x = 1; x < mRows.length; x++){
    var sx = normSerial_(mRows[x][4]);
    if(!sx) continue;
    var same = String(mRows[x][2]) === thisBatch && String(mRows[x][3]) === thisLot;
    var rx = Number(mRows[x][MASTER_ROUND_COL-1]) || 1;
    if(same && rx < round) continue;              // inspected in an earlier round of this lot: fine
    if(!existing[sx]) existing[sx] = String(mRows[x][3]||"");
  }
  for(var k = 0; k < units.length; k++){
    var su = normSerial_(units[k] && units[k].serial);
    if(!/^[A-Z0-9]{1,40}$/.test(su)) return {ok:false, error:"Sample " + (k+1) + ": serial number is missing or not valid"};
    if(seen[su]) clash.push(su + " (repeated in this lot)");
    else if(existing[su]) clash.push(su + " (already recorded in " + existing[su] + ")");
    seen[su] = true;
  }
  if(clash.length) return {ok:false, error:"Duplicate serial number, nothing was saved: " + clash.join(", ")};

  var out = units.map(function(u){
    var ts = u.inspectedAt ? new Date(u.inspectedAt) : new Date();
    if(isNaN(ts.getTime()) || ts > new Date(Date.now() + 86400000)) ts = new Date();
    return [ts, safe_(d.inspector, 80), thisBatch, thisLot, safe_(normSerial_(u.serial), 40),
      pf(u.aes), pf(u.general), pf(u.func), pf(u.screen), safe_(u.bd, 300),
      safe_(u.remark, 500), pf(u.sampleResult), lotResult,
      Number(bRows[lotRow][9]) || "", Number(bRows[lotRow][10]) || "",
      safe_(bRows[lotRow][11], 80), safe_(d.prodMonthCode, 5), fgPassQty,
      safe_(d.bdSpec, 500), round];
  });
  master.getRange(master.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);

  if(lotResult === "PASS"){
    bsh.getRange(lotRow+1, 6).setValue("Closed");
    bsh.getRange(lotRow+1, 9).setValue("PASS");
    bsh.getRange(lotRow+1, 14).setValue(fgPassQty);
    bsh.getRange(lotRow+1, COL_REOFFER).setValue("");
    logLot_(thisBatch, thisLot, round, "Passed", units.length + " samples, none defective" + (round > 1 ? " (reoffer)" : ""), d.inspector);
  } else {
    bsh.getRange(lotRow+1, 6).setValue("Hold");
    bsh.getRange(lotRow+1, 9).setValue("HOLD");
    bsh.getRange(lotRow+1, 14).setValue(0);
    var bad = units.filter(function(u){ return pf(u.sampleResult) === "Fail"; }).map(function(u){ return normSerial_(u.serial); });
    logLot_(thisBatch, thisLot, round, "Hold", "Defective: " + bad.join(", ") + " (" + units.length + " samples inspected)", d.inspector);
  }
  return {ok:true, result:lotResult, round:round};
}

/* Rework of a lot on hold. Saves every reworked unit against its inspection check point,
   then reoffers the lot: the defective sample units always go into the new sample,
   the rest are picked at random from the reworked units; any slots left are scanned fresh. */
function recordRework(d){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var bsh = getOrCreateSheet("Batches",[]);
    ensureStageHeaders_(bsh);
    var bRows = bsh.getDataRange().getValues();
    var lotRow = lotRowOf_(bRows, d.batchName, d.lotNumber);
    if(lotRow < 0) return {ok:false, error:"Lot not found. Refresh and try again."};
    if(String(bRows[lotRow][5]) !== "Hold") return {ok:false, error:"Only a lot on hold can be reworked."};
    var thisBatch = String(bRows[lotRow][0]), thisLot = String(bRows[lotRow][1]);
    var round = Number(bRows[lotRow][COL_ROUND-1]) || 1;
    var sampleSize = Number(bRows[lotRow][3]) || FIXED_SAMPLE;
    var lotSize = Number(bRows[lotRow][2]) || FIXED_LOT;

    var items = Array.isArray(d.items) ? d.items : [];
    var count = parseInt(d.unitsReworked, 10);
    if(!(count >= 1)) return {ok:false, error:"Enter the number of units reworked."};
    if(count > lotSize) return {ok:false, error:"Units reworked cannot be more than the lot size (" + lotSize + ")."};
    if(items.length !== count) return {ok:false, error:"You entered " + count + " units reworked but listed " + items.length + " serial numbers. They must match."};
    var seen = {}, clean = [];
    for(var i = 0; i < items.length; i++){
      var it = items[i] || {}, sn = normSerial_(it.serial);
      if(!/^[A-Z0-9]{1,40}$/.test(sn)) return {ok:false, error:"Row " + (i+1) + ": serial number is missing or not valid"};
      if(seen[sn]) return {ok:false, error:"Serial " + sn + " is listed twice"};
      seen[sn] = true;
      var cp = String(it.checkpoint||"");
      if(CHECKPOINTS.indexOf(cp) === -1) return {ok:false, error:"Row " + (i+1) + ": choose the inspection check point"};
      var st = String(it.status||"");
      if(UNIT_STATUSES.indexOf(st) === -1) return {ok:false, error:"Row " + (i+1) + ": choose the unit status"};
      if(!String(it.defect||"").trim()) return {ok:false, error:"Row " + (i+1) + ": describe the defect"};
      if(!String(it.action||"").trim()) return {ok:false, error:"Row " + (i+1) + ": describe the rework done"};
      clean.push({serial:sn, checkpoint:cp, defect:String(it.defect), action:String(it.action), status:st});
    }

    // defective sample units of this round must all be in the rework list
    var mRows = getOrCreateSheet("Master",[]).getDataRange().getValues(), bad = [];
    for(var m = 1; m < mRows.length; m++){
      if(String(mRows[m][2]) !== thisBatch || String(mRows[m][3]) !== thisLot) continue;
      if((Number(mRows[m][MASTER_ROUND_COL-1]) || 1) !== round) continue;
      if(String(mRows[m][11]) === "Fail") bad.push(normSerial_(mRows[m][4]));
    }
    var notListed = bad.filter(function(x){ return !seen[x]; });
    if(notListed.length) return {ok:false, error:"Add the defective sample units to the rework list: " + notListed.join(", ")};

    // reoffer sample: defective units first (unless scrapped), then random reworked units
    var byS = {}; clean.forEach(function(c){ byS[c.serial] = c; });
    var must = bad.filter(function(x){ return byS[x] && byS[x].status !== "Scrapped"; });
    var pool = clean.filter(function(c){ return c.status !== "Scrapped" && must.indexOf(c.serial) === -1; }).map(function(c){ return c.serial; });
    for(var p = pool.length - 1; p > 0; p--){ var q = Math.floor(Math.random() * (p + 1)); var t = pool[p]; pool[p] = pool[q]; pool[q] = t; }
    var pick = must.concat(pool).slice(0, sampleSize);

    var rid = "RW-" + Utilities.formatDate(new Date(), TZ, "yyyyMMdd-HHmmss");
    var rsh = getOrCreateSheet("Rework", ["ReworkID","Time","Batch","Lot","Round","UnitsReworked","SerialNumber","CheckPoint","Defect","ReworkAction","UnitStatus","InReoffer","Remarks","By"]);
    var now = new Date();
    var rows = clean.map(function(c){
      return [rid, now, safe_(thisBatch, 80), safe_(thisLot, 100), round, count, "'" + c.serial, c.checkpoint, safe_(c.defect, 300),
              safe_(c.action, 300), c.status, pick.indexOf(c.serial) > -1 ? "Yes" : "No", safe_(d.remarks||"", 500), safe_(d.inspector, 80)];
    });
    rsh.getRange(rsh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    bsh.getRange(lotRow+1, 6).setValue("Reoffered");
    bsh.getRange(lotRow+1, 9).setValue("");
    bsh.getRange(lotRow+1, COL_ROUND).setValue(round + 1);
    bsh.getRange(lotRow+1, COL_REOFFER).setValue(pick.join(","));
    var scrapped = clean.filter(function(c){ return c.status === "Scrapped"; }).length;
    logLot_(thisBatch, thisLot, round, "Rework", rid + ": " + count + " units reworked" + (scrapped ? ", " + scrapped + " scrapped" : ""), d.inspector);
    logLot_(thisBatch, thisLot, round + 1, "Reoffered", pick.length + " units picked (" + must.length + " defective" + (pick.length < sampleSize ? ", " + (sampleSize - pick.length) + " to scan fresh" : "") + ")", d.inspector);
    return {ok:true, reworkId:rid, round:round + 1, reoffer:pick, fresh:sampleSize - pick.length};
  } finally { lock.releaseLock(); }
}

/* Offered quantity of a lot (1 to 90), changeable any time before the lot is inspected.
   Fewer units: the lot keeps only the offered units (13 samples, or all if fewer);
   the units not offered go back to the batch: every untouched open lot is recalculated into lots of 90 plus
   one remainder lot, and those lots are renumbered into the free lot numbers. Lots already inspected keep
   their numbers and data. */
function lotNum_(ln){ var m = String(ln).match(/(\d+)\s*$/); return m ? Number(m[1]) : 0; }
function setOfferedQty(d){
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    var bsh = getOrCreateSheet("Batches",[]);
    ensureStageHeaders_(bsh);
    var v = bsh.getDataRange().getValues();
    var lotRow = lotRowOf_(v, d.batchName, d.lotNumber);
    if(lotRow < 0) return {ok:false, error:"Lot not found. Refresh and try again."};
    var batch = String(v[lotRow][0]);
    var qty = parseInt(d.qty, 10), size = Number(v[lotRow][2]) || FIXED_LOT;
    if(!(qty >= 1) || qty > FIXED_LOT) return {ok:false, error:"Offered quantity must be between 1 and " + FIXED_LOT + "."};
    if(qty === size) return {ok:false, error:"The lot already has " + size + " units."};

    // which lots of this batch are untouched (never inspected, never reworked)
    var inspected = {};
    var m = getOrCreateSheet("Master",[]).getDataRange().getValues();
    for(var i = 1; i < m.length; i++) if(String(m[i][2]) === batch) inspected[String(m[i][3])] = true;
    // other lots whose offered quantity was set keep their size and number when the batch is recalculated
    var offered = {};
    var hs = getSpreadsheet().getSheetByName("LotHistory");
    if(hs){ var hv = hs.getDataRange().getValues(); for(var j = 1; j < hv.length; j++) if(String(hv[j][1]) === batch && String(hv[j][4]) === "Offered") offered[String(hv[j][2])] = true; }
    var untouched = function(r){
      return String(r[5]) === "Open" && String(r[8]||"") === "" && (Number(r[COL_ROUND-1]) || 1) === 1 &&
             String(r[COL_REOFFER-1]||"") === "" && !inspected[String(r[1])];
    };
    if(!untouched(v[lotRow])) return {ok:false, error:"The offered quantity can only be changed before the lot is inspected."};

    var rows = [];   // sheet row numbers (1-based) of this batch
    for(var r = 1; r < v.length; r++) if(String(v[r][0]) === batch) rows.push(r);
    var fixedNums = {}, pool = [], poolQty = size - qty;
    rows.forEach(function(r){
      if(r === lotRow) { fixedNums[lotNum_(v[r][1])] = true; return; }
      if(untouched(v[r]) && !offered[String(v[r][1])]) { pool.push(r); poolQty += Number(v[r][2]) || 0; }
      else fixedNums[lotNum_(v[r][1])] = true;
    });

    if(poolQty < 0) return {ok:false, error:"Not enough units in the open lots of this batch to offer " + qty + " (only " + (qty + poolQty) + " available)."};
    // recalculate the untouched units into lots of 90 + remainder
    var sizes = [];
    for(var left = poolQty; left > 0; left -= FIXED_LOT) sizes.push(Math.min(FIXED_LOT, left));
    var nums = [], n = 1;
    while(nums.length < sizes.length){ if(!fixedNums[n]) nums.push(n); n++; }
    var total = Object.keys(fixedNums).length + sizes.length;
    var tpl = v[lotRow].slice();
    var lotName = function(k){ return batch + " Lot " + k; };

    // 1. the offered lot keeps its number with the offered quantity
    bsh.getRange(lotRow + 1, 3, 1, 2).setValues([[qty, Math.min(FIXED_SAMPLE, qty)]]);
    // 2. reuse the untouched rows for the new lots, add or remove rows as needed
    var w = tpl.length;
    var newRow = function(k, sz){
      var x = tpl.slice(); while(x.length < w) x.push("");
      x[1] = lotName(k); x[2] = sz; x[3] = Math.min(FIXED_SAMPLE, sz); x[5] = "Open"; x[8] = ""; x[13] = "";
      x[COL_ROUND-1] = ""; x[COL_REOFFER-1] = "";
      return x;
    };
    for(var k = 0; k < sizes.length; k++){
      var data = newRow(nums[k], sizes[k]);
      if(k < pool.length) bsh.getRange(pool[k] + 1, 1, 1, w).setValues([data]);
      else bsh.appendRow(data);
    }
    for(var z = pool.length - 1; z >= sizes.length; z--) bsh.deleteRow(pool[z] + 1);   // bottom-up keeps row numbers valid
    // 3. total lots for every lot of the batch
    var v2 = bsh.getDataRange().getValues();
    for(var y = 1; y < v2.length; y++) if(String(v2[y][0]) === batch) bsh.getRange(y + 1, 11).setValue(total);

    logLot_(batch, String(v[lotRow][1]), 1, "Offered", qty + " units offered, was " + size + " (" + Math.min(FIXED_SAMPLE, qty) + " samples). " +
            (qty < size ? (size - qty) + " units moved back to" : (qty - size) + " units taken from") + " the open lots; recalculated: " + sizes.length + " lot" + (sizes.length === 1 ? "" : "s") +
            (sizes.length ? " (" + nums.map(function(x, i){ return "Lot " + x + ": " + sizes[i]; }).join(", ") + ")" : "") + ". Batch now has " + total + " lots.", d.inspector);
    return {ok:true, qty:qty, was:size, samples:Math.min(FIXED_SAMPLE, qty), moved:size - qty, lots:nums.map(function(x, i){ return {lotNumber:lotName(x), size:sizes[i]}; }), totalLots:total};
  } finally { lock.releaseLock(); }
}

/****************************************************
 * FAI (first article inspection): 5 samples per batch, in the format of the
 * "Projection Unit Finished Product Outgoing Inspection Report".
 *  FAISpec   : Model | SpecJSON | UpdatedBy | UpdatedAt      (check items + limits per model)
 *  FAI       : ReportID | Batch | Lot | Model | Result | Defective | Inspector | Reviewer | InspectionDate |
 *              Revision | CreatedBy | CreatedAt | UpdatedBy | UpdatedAt | DataJSON
 *  FAI_Detail: ReportID | Batch | Lot | Sample | Serial | Section | Item | Value | Judgment   (one row per sample and item)
 *  One report per batch (Lot column left blank). Batches column R (FAIStatus) holds the batch FAI result on every lot.
 ****************************************************/
var FAI_SAMPLES = 5;
var FAI_RESULTS = ["Pass","Fail","Special procurement"];
var FAI_TYPES   = ["check","text","number","xy"];
function faiItem_(id, section, name, content, type, extra){
  var o = {id:id, section:section, name:name, content:content||"", type:type||"check", na:false};
  for(var k in (extra||{})) o[k] = extra[k];
  return o;
}
function defaultFaiSpec_(){
  return {customer:"Lumio", samples:FAI_SAMPLES, items:[
    faiItem_("pack","Appearance","Packaging inspection","Printing content, dimensions, colours etc. of the packaging materials comply with the SIP or specification."),
    faiItem_("appear","Appearance","Appearance inspection","Product colour, surface spots, scratches, silk screen, lens etc. meet the SIP or specification."),
    faiItem_("struct","Appearance","Structure inspection","Each component is compact; clearances and step differences comply with the SIP."),
    faiItem_("accModel","Appearance","Accessory inspection: model, quantity, size","Model, quantity, dimensions etc. of accessories comply with the SIP or specification."),
    faiItem_("accPack","Appearance","Accessory inspection: packaging, appearance","Packaging and appearance of accessories comply with the SIP or specification."),
    faiItem_("model","Functional test","Product model","Model shown matches the order.","text",{expected:""}),
    faiItem_("sw","Functional test","Software version","Software version matches the release.","text",{expected:""}),
    faiItem_("led","Functional test","Indicator light check"),
    faiItem_("res","Functional test","Resolution"),
    faiItem_("av","Functional test","Video (AV)","", "check",{na:true}),
    faiItem_("usb","Functional test","USB"),
    faiItem_("chan","Functional test","Channel test","", "check",{na:true}),
    faiItem_("headset","Functional test","Headset test"),
    faiItem_("spdif","Functional test","Optical fibre (SPDIF)","", "check",{na:true}),
    faiItem_("af","Functional test","Auto focus test"),
    faiItem_("mf","Functional test","Manual focus test"),
    faiItem_("ak","Functional test","Auto keystone test"),
    faiItem_("mk","Functional test","Manual keystone test"),
    faiItem_("wifi","Functional test","Wi-Fi test"),
    faiItem_("bt","Functional test","Bluetooth"),
    faiItem_("hdmi","Functional test","HDMI"),
    faiItem_("reset","Functional test","Restore factory settings"),
    faiItem_("remote","Functional test","Remote control test"),
    faiItem_("power","Functional test","Power cord / adapter"),
    faiItem_("otherAcc","Functional test","Other accessories"),
    faiItem_("pure","Picture inspection","Pure colour pattern (R/G/B/Y/W/B)"),
    faiItem_("zoom","Picture inspection","Picture zoom (max, min)"),
    faiItem_("bright","Optical testing","Brightness","Calculation inputs to be added; enter the final value for now.","number",{unit:"lm", min:"", max:"", calcPending:true}),
    faiItem_("contrast","Optical testing","Contrast ratio","Calculation inputs to be added; enter the final value for now.","number",{unit:":1", min:"", max:"", calcPending:true}),
    faiItem_("color","Optical testing","Colour coordinates","x and y of white.","xy",{x:"", y:"", tol:""}),
    faiItem_("dimL","Finished product dimensions","L (length)","","number",{unit:"mm", nom:"", tol:"", min:"", max:""}),
    faiItem_("dimW","Finished product dimensions","W (width)","","number",{unit:"mm", nom:"", tol:"", min:"", max:""}),
    faiItem_("dimH","Finished product dimensions","H (height)","","number",{unit:"mm", nom:"", tol:"", min:"", max:""}),
    faiItem_("prodLabel","Packaging","Product label"),
    faiItem_("packLabel","Packaging","Packaging label"),
    faiItem_("packAppear","Packaging","Packaging appearance"),
    faiItem_("packMethod","Packaging","Packaging method")
  ]};
}
function faiSpecSheet_(){ return getOrCreateSheet("FAISpec", ["Model","SpecJSON","UpdatedBy","UpdatedAt"]); }
function getFaiSpec(d){
  var model = String(d.model||"").trim();
  var v = faiSpecSheet_().getDataRange().getValues();
  for(var i = 1; i < v.length; i++){
    if(String(v[i][0]).trim() === model && model){
      try { return {ok:true, model:model, spec:JSON.parse(String(v[i][1])), saved:true, updatedBy:String(v[i][2]||""),
                    updatedAt: v[i][3] instanceof Date ? v[i][3].toISOString() : String(v[i][3]||"")}; } catch(e){}
    }
  }
  return {ok:true, model:model, spec:defaultFaiSpec_(), saved:false};
}
function cleanStr_(v, n){ return String(v == null ? "" : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").substring(0, n || 200); }
function cleanNum_(v){ if(v === "" || v == null) return ""; var x = Number(v); return isFinite(x) ? x : ""; }
function cleanFaiSpec_(sp){
  sp = sp || {};
  var items = Array.isArray(sp.items) ? sp.items : [];
  if(!items.length || items.length > 150) throw new Error("A specification needs between 1 and 150 check items");
  var ids = {};
  var out = {customer:cleanStr_(sp.customer||"Lumio", 80), samples:FAI_SAMPLES, items:items.map(function(it, i){
    var id = cleanStr_(it.id || ("item" + i), 40).replace(/[^A-Za-z0-9_-]/g, "") || ("item" + i);
    while(ids[id]) id = id + "_"; ids[id] = true;
    var type = FAI_TYPES.indexOf(it.type) > -1 ? it.type : "check";
    var o = {id:id, section:cleanStr_(it.section, 60) || "Other", name:cleanStr_(it.name, 120), content:cleanStr_(it.content, 400), type:type, na:!!it.na};
    if(!o.name) throw new Error("Item " + (i + 1) + " needs a name");
    if(type === "text") o.expected = cleanStr_(it.expected, 120);
    if(type === "number"){ o.unit = cleanStr_(it.unit, 12); ["min","max","nom","tol"].forEach(function(k){ o[k] = cleanNum_(it[k]); }); o.calcPending = !!it.calcPending; }
    if(type === "xy"){ ["x","y","tol"].forEach(function(k){ o[k] = cleanNum_(it[k]); }); }
    return o;
  })};
  return out;
}
function saveFaiSpec(d){
  var model = String(d.model||"").trim();
  if(!findModel_(model)) return {ok:false, error:"Model not found"};
  var spec;
  try { spec = cleanFaiSpec_(d.spec); } catch(e){ return {ok:false, error:e.message}; }
  var json = JSON.stringify(spec);
  if(json.length > 45000) return {ok:false, error:"The specification is too large"};
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var sh = faiSpecSheet_(), v = sh.getDataRange().getValues(), row = -1;
    for(var i = 1; i < v.length; i++) if(String(v[i][0]).trim() === model){ row = i + 1; break; }
    var data = [[model, json, safe_(d.inspector, 80), new Date()]];
    if(row < 0) sh.appendRow(data[0]); else sh.getRange(row, 1, 1, 4).setValues(data);
    return {ok:true};
  } finally { lock.releaseLock(); }
}
function faiSheet_(){ return getOrCreateSheet("FAI", ["ReportID","Batch","Lot","Model","Result","Defective","Inspector","Reviewer","InspectionDate","Revision","CreatedBy","CreatedAt","UpdatedBy","UpdatedAt","DataJSON"]); }
function getFai(d){
  var bn = String(d.batchName||""), ln = "";
  var sh = getSpreadsheet().getSheetByName("FAI");
  if(!sh) return {ok:true, found:false};
  var v = sh.getDataRange().getValues();
  for(var i = v.length - 1; i >= 1; i--){
    if(String(v[i][1]) !== bn || String(v[i][2]) !== ln) continue;
    var data = {}; try { data = JSON.parse(String(v[i][14]||"{}")); } catch(e){}
    var t = function(x){ return x instanceof Date ? x.toISOString() : String(x||""); };
    return {ok:true, found:true, report:{id:String(v[i][0]), result:String(v[i][4]), defective:Number(v[i][5])||0, inspector:String(v[i][6]),
            reviewer:String(v[i][7]), revision:Number(v[i][9])||1, createdBy:String(v[i][10]), createdAt:t(v[i][11]),
            updatedBy:String(v[i][12]), updatedAt:t(v[i][13]), data:data}};
  }
  return {ok:true, found:false};
}
/* Saves or updates the FAI report of one batch. The judgment of every measured item is worked out here again
   from the lot model's specification, so the stored result does not depend on the browser. */
function faiJudge_(it, val){
  if(it.na) return "N/A";
  if(it.type === "check") return val === "Pass" || val === "Fail" ? val : "";
  if(it.type === "text"){
    var t = String(val && val.v != null ? val.v : "").trim();
    if(!t) return "";
    if(it.expected) return t.toLowerCase() === String(it.expected).trim().toLowerCase() ? "Pass" : "Fail";
    return val && (val.j === "Pass" || val.j === "Fail") ? val.j : "";
  }
  if(it.type === "number"){
    var raw = val && val.v != null ? String(val.v).trim() : "";
    if(raw === "") return "";
    var x = Number(raw); if(!isFinite(x)) return "Fail";
    var hasNom = it.nom !== "" && it.tol !== "" && it.nom != null && it.tol != null;
    var hasMin = it.min !== "" && it.min != null, hasMax = it.max !== "" && it.max != null;
    if(hasNom) return Math.abs(x - Number(it.nom)) <= Number(it.tol) + 1e-9 ? "Pass" : "Fail";
    if(hasMin || hasMax) return (!hasMin || x >= Number(it.min)) && (!hasMax || x <= Number(it.max)) ? "Pass" : "Fail";
    return val && (val.j === "Pass" || val.j === "Fail") ? val.j : "";
  }
  if(it.type === "xy"){
    var xs = val && val.x != null ? String(val.x).trim() : "", ys = val && val.y != null ? String(val.y).trim() : "";
    if(xs === "" || ys === "") return "";
    var X = Number(xs), Y = Number(ys); if(!isFinite(X) || !isFinite(Y)) return "Fail";
    if(it.x !== "" && it.y !== "" && it.tol !== "" && it.x != null && it.y != null && it.tol != null)
      return Math.abs(X - Number(it.x)) <= Number(it.tol) + 1e-9 && Math.abs(Y - Number(it.y)) <= Number(it.tol) + 1e-9 ? "Pass" : "Fail";
    return val && (val.j === "Pass" || val.j === "Fail") ? val.j : "";
  }
  return "";
}
function saveFai(d){
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try{
    var bsh = getOrCreateSheet("Batches",[]);
    ensureStageHeaders_(bsh);
    var bv = bsh.getDataRange().getValues();
    var batch = String(d.batchName||""), lots = [];
    for(var q = 1; q < bv.length; q++) if(String(bv[q][0]) === batch) lots.push(q);
    if(!lots.length) return {ok:false, error:"Batch not found. Refresh and try again."};
    var lot = "", model = String(bv[lots[0]][11]||"");
    var batchQty = Number(bv[lots[0]][9]) || lots.reduce(function(t, r){ return t + (Number(bv[r][2]) || 0); }, 0);
    var rep = d.report || {};
    var spec = getFaiSpec({model:model}).spec;
    var nS = Math.min(FAI_SAMPLES, batchQty || FAI_SAMPLES);
    var samples = Array.isArray(rep.samples) ? rep.samples.slice(0, nS) : [];
    if(samples.length !== nS) return {ok:false, error:"FAI needs " + nS + " samples"};
    var seen = {};
    for(var s = 0; s < nS; s++){
      var sn = normSerial_(samples[s] && samples[s].serial);
      if(!/^[A-Z0-9]{1,40}$/.test(sn)) return {ok:false, error:"Sample " + (s+1) + ": enter the serial number"};
      if(seen[sn]) return {ok:false, error:"Serial " + sn + " is used for two samples"};
      seen[sn] = true; samples[s].serial = sn;
    }
    // judge every item for every sample against the specification
    var detail = [], defective = 0, missing = [];
    var clean = {header:{}, samples:[], remarks:{}, others:cleanStr_(rep.others, 1000), defects:cleanStr_(rep.defects, 1000)};
    var H = rep.header || {};
    ["salesOrder","customer","reviewer","inspectionDate","firstSerial","lastSerial","approval"].forEach(function(k){ clean.header[k] = cleanStr_(H[k], 120); });
    if(clean.header.inspectionDate && !/^\d{4}-\d{2}-\d{2}$/.test(clean.header.inspectionDate)) return {ok:false, error:"Inspection date is not valid"};
    for(var k2 in (rep.remarks || {})) if(spec.items.some(function(it){ return it.id === k2; })) clean.remarks[k2] = cleanStr_(rep.remarks[k2], 300);
    for(var j = 0; j < nS; j++){
      var res = (samples[j] && samples[j].results) || {}, cs = {serial:samples[j].serial, results:{}}, bad = false;
      spec.items.forEach(function(it){
        var val = res[it.id], jd = faiJudge_(it, val);
        var keep = it.type === "check" ? jd : it.type === "xy" ? {x:cleanStr_(val && val.x, 20), y:cleanStr_(val && val.y, 20), j:jd} : {v:cleanStr_(val && val.v, 60), j:jd};
        cs.results[it.id] = keep;
        if(jd === "") missing.push(it.name + " (sample " + (j+1) + ")");
        if(jd === "Fail") bad = true;
        var shown = it.type === "check" ? "" : it.type === "xy" ? ("x " + keep.x + ", y " + keep.y) : keep.v;
        detail.push(["", safe_(batch, 80), safe_(lot, 100), j + 1, "'" + samples[j].serial, safe_(it.section, 60), safe_(it.name, 120), safe_(shown, 80), jd]);
      });
      if(bad) defective++;
      clean.samples.push(cs);
    }
    if(missing.length) return {ok:false, error:"Not filled yet: " + missing.slice(0, 6).join(", ") + (missing.length > 6 ? " and " + (missing.length - 6) + " more" : "")};
    var result = String(rep.result||"");
    if(FAI_RESULTS.indexOf(result) === -1) return {ok:false, error:"Choose the final determination"};
    if(result === "Pass" && defective) return {ok:false, error:defective + " sample(s) failed; the final determination cannot be Pass."};
    if(result === "Special procurement" && (!clean.header.approval || !clean.defects)) return {ok:false, error:"Special procurement needs the defect description and the customer approval reference"};
    clean.result = result; clean.defective = defective;
    var json = JSON.stringify(clean);
    if(json.length > 45000) return {ok:false, error:"The report is too large"};

    var sh = faiSheet_(), v = sh.getDataRange().getValues(), row = -1;
    for(var i = 1; i < v.length; i++) if(String(v[i][1]) === batch && String(v[i][2]) === lot){ row = i + 1; }
    var now = new Date(), id, rev;
    if(row < 0){
      id = "FAI-" + Utilities.formatDate(now, TZ, "yyyyMMdd-HHmmss");
      rev = 1;
      sh.appendRow([id, safe_(batch, 80), safe_(lot, 100), safe_(model, 80), result, defective, safe_(d.inspector, 80), safe_(clean.header.reviewer, 80),
                    clean.header.inspectionDate, rev, safe_(d.inspector, 80), now, safe_(d.inspector, 80), now, json]);
    } else {
      id = String(v[row-1][0]); rev = (Number(v[row-1][9]) || 1) + 1;
      sh.getRange(row, 5, 1, 11).setValues([[result, defective, String(v[row-1][6]), safe_(clean.header.reviewer, 80), clean.header.inspectionDate, rev,
                                              String(v[row-1][10]), v[row-1][11], safe_(d.inspector, 80), now, json]]);
    }
    // detail rows: replace this report's rows
    var ds = getOrCreateSheet("FAI_Detail", ["ReportID","Batch","Lot","Sample","SerialNumber","Section","Item","Value","Judgment"]);
    var dv = ds.getDataRange().getValues();
    for(var r = dv.length - 1; r >= 1; r--) if(String(dv[r][0]) === id) ds.deleteRow(r + 1);
    detail.forEach(function(x){ x[0] = id; });
    if(detail.length) ds.getRange(ds.getLastRow() + 1, 1, detail.length, detail[0].length).setValues(detail);

    lots.forEach(function(r){ bsh.getRange(r + 1, COL_FAI_STATUS).setValue(result); });
    logLot_(batch, "", 1, rev === 1 ? "FAI" : "FAI edited",
            id + " revision " + rev + ": " + result + ", " + defective + " of " + nS + " samples defective", d.inspector);
    return {ok:true, id:id, revision:rev, result:result, defective:defective};
  } finally { lock.releaseLock(); }
}

/* Admin: put a lot that failed under the old rule on hold, so it can be reworked and reoffered. */
function holdLot(d){
  var bsh = getOrCreateSheet("Batches",[]);
  ensureStageHeaders_(bsh);
  var bRows = bsh.getDataRange().getValues();
  var lotRow = lotRowOf_(bRows, d.batchName, d.lotNumber);
  if(lotRow < 0) return {ok:false, error:"Lot not found"};
  if(String(bRows[lotRow][8]) !== "FAIL") return {ok:false, error:"Only a failed lot can be moved to hold"};
  bsh.getRange(lotRow+1, 6).setValue("Hold");
  bsh.getRange(lotRow+1, 9).setValue("HOLD");
  logLot_(String(bRows[lotRow][0]), String(bRows[lotRow][1]), Number(bRows[lotRow][COL_ROUND-1]) || 1, "Hold", "Failed lot moved to hold for rework", d.inspector);
  return {ok:true};
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

/* Models only (fast). Usage and suggestions come from modelInsights() on the Models page. */
function getModels(){ return {ok:true, models:readModels_()}; }
/* Models + what is already in Batches/Master (used to suggest models to add) */
function modelInsights(){
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
  if(!NAME_RE_.test(name)) return {ok:false, error:"Model name: up to 80 letters, numbers, spaces and _ - . / ( ) # &"};
  if(!NAME_RE_.test(lumio)) return {ok:false, error:"Lumio model number: up to 80 letters, numbers, spaces and _ - . / ( ) # &"};
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
               Number(m.nA), Number(m.nB), Number(m.nAB), safe_(d.updatedBy||"", 64), new Date(), lumio];
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
      inspector: String(r[1]||""), serial: normSerial_(r[4]),
      aes: String(r[5]||""), general: String(r[6]||""), func: String(r[7]||""), screen: String(r[8]||""),
      bd: String(r[9]||""), remark: String(r[10]||""),
      sampleResult: String(r[11]||""), lotResult: String(r[12]||""),
      round: Number(r[MASTER_ROUND_COL-1]) || 1
    });
  }
  var hist = [], rw = [];
  var hs = getSpreadsheet().getSheetByName("LotHistory");
  if(hs){
    var hv = hs.getDataRange().getValues();
    for(var h = 1; h < hv.length; h++){
      if(String(hv[h][1]) !== bn || String(hv[h][2]) !== ln) continue;
      hist.push({time: hv[h][0] instanceof Date ? hv[h][0].toISOString() : String(hv[h][0]||""), round:Number(hv[h][3])||1,
                 event:String(hv[h][4]||""), details:String(hv[h][5]||""), by:String(hv[h][6]||"")});
    }
  }
  var rs = getSpreadsheet().getSheetByName("Rework");
  if(rs){
    var rv = rs.getDataRange().getValues();
    for(var k = 1; k < rv.length; k++){
      if(String(rv[k][2]) !== bn || String(rv[k][3]) !== ln) continue;
      rw.push({id:String(rv[k][0]), time: rv[k][1] instanceof Date ? rv[k][1].toISOString() : String(rv[k][1]||""),
               round:Number(rv[k][4])||1, unitsReworked:Number(rv[k][5])||0, serial:normSerial_(rv[k][6]), checkpoint:String(rv[k][7]||""),
               defect:String(rv[k][8]||""), action:String(rv[k][9]||""), status:String(rv[k][10]||""), inReoffer:String(rv[k][11]||"") === "Yes",
               remarks:String(rv[k][12]||""), by:String(rv[k][13]||"")});
    }
  }
  return {ok:true, units:out, history:hist, rework:rw};
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
  exceptionsSheet_().appendRow(["'" + s, safe_(d.reason||"Kept in more than one lot", 200), safe_(d.by||"", 64), new Date()]);
  return {ok:true};
}

function findDuplicates_(rows, exc){
  exc = exc || {};
  var bySerial = {}, order = [];
  for(var i = 1; i < rows.length; i++){
    var sn = normSerial_(rows[i][4]);
    if(!sn) continue;
    // Approved exceptions: duplicates are only checked inside the same lot
    var rnd = Number(rows[i][MASTER_ROUND_COL-1]) || 1;
    // approved exceptions are only checked inside the same lot; reoffer rounds re-inspect the same units on purpose
    var s = exc[sn] ? sn + "||" + String(rows[i][3]||"") : sn;
    if(rnd > 1) s = sn + "||" + String(rows[i][3]||"") + "||r" + rnd;
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
    var keepRows = [rows[0]], moved = [], now = new Date(), by = safe_(d.by||"", 64);
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
 * BOM data lives in this spreadsheet (after migrateBomIntoPortal), see bomSS_().
 * Each verification is written to BOM_VERIFY_LOG + LOT_BOM_DETAIL there,
 * in the same format as earlier verifications.
 ****************************************************/
function ensureStageHeaders_(sheet){
  var h = sheet.getRange(1, COL_BOM_STATUS, 1, 3).getValues()[0];
  if(h[0] === "") sheet.getRange(1, COL_BOM_STATUS).setValue("BOMStatus");
  if(h[1] === "") sheet.getRange(1, COL_BOM_VER).setValue("BOMVerId");
  if(h[2] === "") sheet.getRange(1, COL_FAI_STATUS).setValue("FAIStatus");
  var h2 = sheet.getRange(1, COL_ROUND, 1, 2).getValues()[0];
  if(h2[0] === "") sheet.getRange(1, COL_ROUND).setValue("Round");
  if(h2[1] === "") sheet.getRange(1, COL_REOFFER).setValue("ReofferSerials");
}
var BOM_SS_CACHE_ = null;
function bomMigrated_(){ return !!getSpreadsheet().getSheetByName("BOM_CONTROL"); }
function bomSS_(){
  if(!BOM_SS_CACHE_) BOM_SS_CACHE_ = bomMigrated_() ? getSpreadsheet() : SpreadsheetApp.openById(LEGACY_BOM_SS_ID);
  return BOM_SS_CACHE_;
}
/* One-off: copies every tab of the old Lumio_Part_Reference_DB into this spreadsheet (same names).
   Tabs that already exist here are left alone, so it is safe to run again. The old file is not changed.
   Run it from the Apps Script editor (select migrateBomIntoPortal, Run). */
function migrateBomIntoPortal(){
  var src = SpreadsheetApp.openById(LEGACY_BOM_SS_ID), dst = getSpreadsheet(), done = [], skipped = [];
  var tabs = src.getSheets();
  tabs.sort(function(a, b){ return (a.getName() === "BOM_CONTROL") - (b.getName() === "BOM_CONTROL"); });
  tabs.forEach(function(sh){
    var name = sh.getName();
    if(dst.getSheetByName(name)){ skipped.push(name); return; }
    var c = sh.copyTo(dst);
    c.setName(name);
    done.push(name + " (" + Math.max(sh.getLastRow() - 1, 0) + " rows)");
  });
  // BOM_CONTROL is copied last: the portal switches to this spreadsheet only once every other tab is here
  CacheService.getScriptCache().removeAll(["c_batches","c_models","c_bomlist"]);
  var msg = "Copied: " + (done.join(", ") || "nothing") + ". Already here: " + (skipped.join(", ") || "none") + ".";
  console.log(msg);
  return msg;
}
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
    console.error(e); return {ok:false, error:"Cannot open the BOM sheet. Ask an admin to check access to it."};
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
    console.error(e); return {ok:false, error:"Cannot open the BOM sheet. Ask an admin to check access to it."};
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
    var bv = getOrCreateSheet("Batches",[]).getDataRange().getValues(), cur = null, bModel = "";
    for(var b = 1; b < bv.length; b++){ if(String(bv[b][0]) === batchName){ cur = String(bv[b][COL_BOM_STATUS-1]||""); bModel = String(bv[b][11]||""); break; } }
    if(cur === null) return {ok:false, error:"Batch " + batchName + " not found"};
    if(cur === "Passed") return {ok:false, error:"BOM verification for this batch has already passed."};
    if(cur !== "Pending" && cur !== "Failed") return {ok:false, error:"This batch does not need BOM verification."};
    var spec = findModel_(bModel);
    if(!spec || spec.lumioModel !== model) return {ok:false, error:"This batch must be verified against the BOM of its own model" + (spec && spec.lumioModel ? " (" + spec.lumioModel + ")" : "") + "."};

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
      rows.push({p:p, lotPn:safe_(lotPn, 80), lotQ:isNaN(lotQ) ? "" : lotQ, remark:safe_(it.remark||"", 500), res:res});
    });

    var log = ss.getSheetByName("BOM_VERIFY_LOG");
    var det = ss.getSheetByName("LOT_BOM_DETAIL");
    if(!log || !det) return {ok:false, error:"BOM_VERIFY_LOG or LOT_BOM_DETAIL tab missing in the BOM sheet"};
    var verId = nextVerId_(log);
    var when = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
    var result = dev === 0 ? "All matched" : "Deviation";
    log.appendRow([verId, when, model, safe_(batchName, 80), ctl.rev, String(ctl.status).toLowerCase() === "locked" ? "Locked" : "Not locked", parts.length, dev, result,
                   safe_(d.by||"", 80), safe_(d.remarks||"", 1000)]);
    var detail = rows.map(function(r){
      return [verId, model, safe_(batchName, 80), ctl.rev, safe_(r.p.partNo, 80), safe_(r.p.description, 500), safe_(r.p.imageUrl, 500), r.p.qps,
              r.lotPn, r.lotQ, r.remark, r.res];
    });
    det.getRange(det.getLastRow()+1, 1, detail.length, detail[0].length).setValues(detail);

    var status = dev === 0 ? "Passed" : "Failed";
    setBomStatus_(batchName, status, verId);
    return {ok:true, verId:verId, status:status, partsChecked:parts.length, deviations:dev, result:result};
  }catch(e){
    var msg = String(e && e.message || e);
    if(/^Remark is required/.test(msg)) return {ok:false, error:msg};
    console.error("submitBomVerification failed: " + (e && e.stack || e));
    return {ok:false, error:"Could not save the BOM verification. Please try again."};
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
    console.error(e); return {ok:false, error:"Cannot open the BOM sheet. Ask an admin to check access to it."};
  }
}

/* Admin: change a batch's BOM step (Waived = not required, Pending = require verification) */
function setBatchStage(d){
  var st = String(d.bomStatus||"");
  if(["Waived","Pending"].indexOf(st) === -1) return {ok:false, error:"Unsupported status"};
  var n = setBomStatus_(String(d.batchName||""), st, st === "Waived" ? safe_("Waived by " + String(d.by||"admin"), 80) : "");
  return n ? {ok:true, lots:n} : {ok:false, error:"Batch not found"};
}

/* =====================================================================
   BOM MANAGEMENT (native). Same rules as the Part DB tools; data in this spreadsheet.
   Ported from PartDB.gs. Everything here is reached through the "bom" action only.
   ===================================================================== */
var BOMX = (function(){
const IMAGE_FOLDER_ID = '1xwpe_z03u3dk2z40QHJ_559yDouuAFHM';
const SHEET_NAME      = 'PART_MASTER';
const HEADER_ROW      = 1;
const COL_PART_NO     = 1;   // A
const COL_IMAGE       = 3;   // C  (=IMAGE formula goes here)
const COL_IMAGE_URL   = 4;   // D  (raw link, used by other tools)
const ROW_HEIGHT_PX   = 90;  // row height so thumbnails are readable
const THUMB_WIDTH     = 400; // px width requested from Drive
const MAKE_PUBLIC     = true; // set false if your domain blocks link sharing

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
    if (!sh || isSystemSheet_(model)) throw uerr_('Model sheet not found: ' + model);
    const mh = headersOf_(ss.getSheetByName(SHEET_NAME));
    if (countParts_(sh, mh) === 0) throw uerr_('BOM is empty - add parts before locking.');
    const c = getControl_(model);
    if (c.status === 'Locked') throw uerr_(model + ' is already locked.');
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
    if (!sh) throw uerr_('Model sheet not found: ' + model);
    if (!String(reason || '').trim()) throw uerr_('Enter a reason for unlocking.');
    const c = getControl_(model);
    if (c.status !== 'Locked') throw uerr_(model + ' is not locked.');
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
    if (!(q > 0)) throw uerr_('Invalid QPS for ' + it.partNo);
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
      throw uerr_('BOM of ' + locked.map(x => x.sheet.getName()).join(', ') +
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
    if (!sheet || isSystemSheet_(model)) throw uerr_('Model sheet not found: ' + model);
    if (!oldParts.length && !newParts.length) throw uerr_('Enter the parts to remove or replace, the parts to add, or both.');
    if (!reason) throw uerr_('Enter the reason for the change.');
    if (!effectiveDate) throw uerr_('Enter the effective date.');
    if (payload.received && !docLink_(payload.docUrl)) throw uerr_('Paste the link to the ECN document, or answer "No" if you have not received it yet.');

    const ctx = readMaster_();
    const notFound = oldParts.filter(p => !findInModel_(sheet, p, ctx.mh));
    if (notFound.length)
      throw uerr_('Not found in ' + model + '\'s current BOM, so they cannot be replaced: ' +
        notFound.join(', ') + '. Check the part number(s) and try again.');

    // If the same new part number is listed more than once (typed twice, or
    // split across rows by mistake), merge those rows into one by summing
    // their QPS rather than rejecting the ECN.
    const mergedByKey = {};
    const mergedOrder = [];
    newParts.forEach(p => {
      const k = norm_(p.partNo);
      if (!k) throw uerr_('One of the new part rows has no part number.');
      const q = Number(p.qps);
      if (!(q > 0)) throw uerr_('Enter a valid QPS for ' + p.partNo + '.');
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
      throw uerr_('Not in PART_MASTER yet, so they cannot be added: ' + plan.missing.join(', ') +
        '. Add them to PART_MASTER first, then run this ECN again.');

    const ecnNo = nextEcnNo_();
    const docUrl = payload.received ? docLink_(payload.docUrl) : '';

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
    if (!ecnNo) throw uerr_('Missing ECN number.');
    if (!reason) throw uerr_('Enter the reason for the change.');
    if (!effectiveDate) throw uerr_('Enter the effective date.');
    if (payload.received && !payload.keepExistingDoc && !docLink_(payload.docUrl))
      throw uerr_('Paste the link to the ECN document, or keep the existing one.');

    const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
    const last = sh.getLastRow();
    if (last < 2) throw uerr_('ECN ' + ecnNo + ' was not found.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const want = norm_(ecnNo);
    const idxs = [];
    vals.forEach((r, i) => { if (norm_(r[1]) === want) idxs.push(i); });
    if (!idxs.length) throw uerr_('ECN ' + ecnNo + ' was not found.');

    const model = txt_(vals[idxs[0]][2]);
    let docUrl = txt_(vals[idxs[0]][10]);
    if (payload.received && !payload.keepExistingDoc) {
      docUrl = docLink_(payload.docUrl);
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
    if (last < 2) throw uerr_('ECN ' + ecnNo + ' was not found.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const idxs = [];
    vals.forEach((r, i) => { if (norm_(r[1]) === want) idxs.push(i); });
    if (!idxs.length) throw uerr_('ECN ' + ecnNo + ' was not found.');

    const model = txt_(vals[idxs[0]][2]);
    const ecnRev = Number(vals[idxs[0]][3]) || 0;
    const sheet = dataSS_().getSheetByName(model);
    if (!sheet || isSystemSheet_(model)) throw uerr_('Model sheet not found: ' + model);

    const hasPartRows = idxs.some(i => String(vals[i][4]).trim());
    if (!hasPartRows) throw uerr_('ECN ' + ecnNo + ' has no part changes to undo.');

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
      throw uerr_('Cannot delete ' + ecnNo + ' - ' + conflictList.join(', ') +
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
      throw uerr_('Cannot undo ' + ecnNo + ' - part(s) no longer in PART_MASTER: ' + plan.missing.join(', '));
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
    if (last < 2) throw uerr_('ECN ' + ecnNo + ' was not found.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const idxs = [];
    vals.forEach((r, i) => { if (norm_(r[1]) === want) idxs.push(i); });
    if (!idxs.length) throw uerr_('ECN ' + ecnNo + ' was not found.');

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
    if (!sheet || isSystemSheet_(model)) throw uerr_('Model sheet not found: ' + model);
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
      throw uerr_('Revision ' + rev + ' includes part(s) no longer in PART_MASTER: ' + plan.missing.join(', ') +
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
    if (!model) throw uerr_('Missing model.');
    let list = (revs || []).map(Number).filter(n => !isNaN(n));
    list = Array.from(new Set(list)).sort((a, b) => a - b);
    if (list.length < 2) throw uerr_('Pick at least two revisions to merge.');
    for (let i = 1; i < list.length; i++) {
      if (list[i] !== list[i - 1] + 1)
        throw uerr_('Only adjacent revisions can be merged together (e.g. Rev 2 and Rev 3) - not revisions with a gap between them.');
    }
    const minMerged = list[0];
    const maxMerged = list[list.length - 1];

    const sh = migratedSheet_(ECN_SHEET, ECN_HDR);
    const last = sh.getLastRow();
    if (last < 2) throw uerr_('No ECN history found for ' + model + '.');
    const rng = sh.getRange(2, 1, last - 1, ECN_HDR.length);
    const vals = rng.getValues();
    const modelWant = norm_(model);

    const foundRevs = {};
    vals.forEach(r => { if (norm_(r[2]) === modelWant) foundRevs[Number(r[3])] = true; });
    const missing = list.filter(rv => !foundRevs[rv]);
    if (missing.length) throw uerr_('Revision(s) not found for ' + model + ': ' + missing.join(', '));

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

function assignPart(payload) {
  return withLock_(() => {
    const ss  = dataSS_();
    const ctx = readMaster_();
    if (!(norm_(payload.partNo) in ctx.idx))
      throw uerr_('Part ' + payload.partNo + ' not found in ' + SHEET_NAME);

    const per = {};   // model -> { items, remove }
    (payload.add || []).forEach(a => {
      (per[a.model] = per[a.model] || { items: [], remove: [] }).items.push({ partNo: payload.partNo, qps: a.qps });
    });
    (payload.remove || []).forEach(m => {
      (per[m] = per[m] || { items: [], remove: [] }).remove.push(payload.partNo);
    });

    const list = Object.keys(per).map(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet || isSystemSheet_(name)) throw uerr_('Sheet not found: ' + name);
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
      if (isSystemSheet_(nm)) throw uerr_('"' + nm + '" is a reserved sheet name.');
      if (ss.getSheetByName(nm)) throw uerr_('A sheet named "' + nm + '" already exists.');
      sheet = ss.insertSheet(nm);
      created = true;
    } else {
      sheet = ss.getSheetByName(payload.model);
      if (!sheet || isSystemSheet_(payload.model)) throw uerr_('Model sheet not found: ' + payload.model);
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
  if (!sheet || isSystemSheet_(model)) throw uerr_('Model sheet not found: ' + model);
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
    if (!model || !ss.getSheetByName(model)) throw uerr_('Choose a model.');
    if (!lot) throw uerr_('Enter the Lot No.');

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

    if (!detail.length) throw uerr_('This model has no BOM parts to verify.');
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


/** One CSV field, quoted if it needs it. */
function csvField_(v) {
  const s = txt_(v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

/** CSV text of one lot's verification, same columns as exportLotBom, Result last. Client triggers the download. */
function getLotCsv(verId) {
  const info = findVerification_(verId);
  if (!info) throw uerr_('Verification not found: ' + verId);
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

const COMPARE_SHEET      = 'BOM_COMPARE';
const COMPARE_SHOW_IMAGE = true;   // true = add part thumbnail column (makes rows tall)


const TWO_SHEET = 'MODEL_COMPARE';

/** Common / only-in-A / only-in-B parts of two model sheets. */
function compareTwoModels(a, b) {
  if (!a || !b) throw uerr_('Choose both models.');
  if (norm_(a) === norm_(b)) throw uerr_('Choose two different models.');
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

// ---------------- portal overrides ----------------
let CUR_USER_ = '';
function uerr_(m) { const e = new Error(m); e.userMsg = true; return e; }
function dataSS_() { return getSpreadsheet(); }
function user_() { return CUR_USER_ || 'unknown'; }
function withLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(30000);
  try { return fn(); } finally { l.releaseLock(); }
}
function docLink_(u) {
  u = String(u || '').replace(/^'/, '').trim();
  return /^https:\/\/[A-Za-z0-9.-]+\.[A-Za-z]{2,}(\/[^\s"<>]*)?$/.test(u) && u.length < 600 ? u : '';
}
// The portal's own tabs are never model BOMs, whatever they hold.
const PORTAL_TABS_ = ['batches','master','users','sessions','models','rework','lothistory','fai','fai_detail','faispec',
                      'duplicates','serialexceptions'];
function hasPartNo_(s) {
  if (s.getLastColumn() < 1) return false;
  const h = s.getRange(HEADER_ROW, 1, 1, Math.min(s.getLastColumn(), 40)).getValues()[0];
  return h.some(x => String(x).toLowerCase().replace(/[^a-z0-9]/g, '') === 'partno');
}
function reservedName_(name) {
  name = String(name || '');
  return [SHEET_NAME, COMPARE_SHEET, CONTROL_SHEET, ECN_SHEET, VERIFY_SHEET, LOT_SHEET, LOTVIEW_SHEET, TWO_SHEET, 'PART_LOG'].indexOf(name) >= 0 ||
         PORTAL_TABS_.indexOf(name.toLowerCase()) >= 0;
}
function isSystemSheet_(name) {
  if (reservedName_(name)) return true;
  const s = dataSS_().getSheetByName(name);
  return !!s && s.getLastRow() > 0 && !hasPartNo_(s);   // an existing tab that is not a BOM
}
function getModelSheets_() {
  return dataSS_().getSheets().filter(s => !reservedName_(s.getName()) && hasPartNo_(s));
}
function sysSheet_(name, headers) {
  const ss = dataSS_();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}
function modelOk_(model) {
  const sh = dataSS_().getSheetByName(String(model || ''));
  if (!sh || isSystemSheet_(model) || !hasPartNo_(sh)) throw uerr_('Model BOM not found: ' + model);
  return sh;
}

/** Every model side by side (read only; nothing is written to a sheet). */
function compareAll() {
  const ctx = readMaster_();
  const names = [], parts = {};
  getModelSheets_().forEach(s => {
    const name = s.getName();
    names.push(name);
    const n = s.getLastRow() - HEADER_ROW;
    if (n <= 0) return;
    const pc = partColIn_(s, ctx.mh);
    const qc = headersOf_(s).findIndex(h => norm_(h) === QPS_HEADER) + 1;
    const pv = s.getRange(HEADER_ROW + 1, pc, n, 1).getValues();
    const qv = qc ? s.getRange(HEADER_ROW + 1, qc, n, 1).getValues() : null;
    pv.forEach((r, i) => {
      const raw = String(r[0]).trim(), k = norm_(raw);
      if (!k) return;
      if (!parts[k]) parts[k] = { partNo: raw, q: {} };
      parts[k].q[name] = qv ? qv[i][0] : '';
    });
  });
  const total = names.length;
  const rows = Object.keys(parts).map(k => {
    const p = parts[k];
    const used = names.filter(n => n in p.q);
    const qs = used.map(n => String(p.q[n]));
    const ex = masterExtra_(ctx, ctx.idx[k]);
    return { partNo: p.partNo, desc: ex.desc, img: ex.img, qps: names.map(n => n in p.q ? p.q[n] : null),
             used: used.length,
             status: used.length === total ? 'Common' : used.length === 1 ? 'Only in ' + used[0] : 'Shared',
             check: used.length < 2 ? '' : (qs.filter((x, i) => qs.indexOf(x) === i).length === 1 ? 'Same' : 'Different') };
  }).sort((a, b) => b.used - a.used || a.partNo.localeCompare(b.partNo));
  return { models: names, rows: rows };
}

// ---------------- parts (PART_MASTER) ----------------
const PART_RE_ = /^[A-Za-z0-9][A-Za-z0-9._\-\/ ]{0,59}$/;
const PART_LOG_HDR = ['Date_Time', 'Part_No', 'Field', 'Old_Value', 'New_Value', 'Changed_By'];
function partFields_(ctx) {
  return ctx.mh.map((h, i) => ({ h: h, i: i }))
    .filter(x => x.h && x.i !== COL_PART_NO - 1 && ['IMAGE', 'BLANK'].indexOf(norm_(x.h)) < 0);
}
function imgFormula_(u) { u = docLink_(u); return u ? '=IMAGE("' + u.replace(/"/g, '%22') + '",1)' : ''; }
function usageMap_(mh) {
  const m = {};
  getModelSheets_().forEach(s => {
    const n = s.getLastRow() - HEADER_ROW;
    if (n <= 0) return;
    s.getRange(HEADER_ROW + 1, partColIn_(s, mh), n, 1).getValues().forEach(r => {
      const k = norm_(r[0]);
      if (k) (m[k] = m[k] || []).indexOf(s.getName()) < 0 && m[k].push(s.getName());
    });
  });
  return m;
}
/** Every part with its editable fields and the model BOMs that use it. */
function listParts() {
  const ctx = readMaster_(), f = partFields_(ctx), use = usageMap_(ctx.mh), rows = [];
  ctx.mVals.forEach(r => {
    const p = String(r[COL_PART_NO - 1]).trim();
    if (p) rows.push({ partNo: p, v: f.map(x => txt_(r[x.i])), used: use[norm_(p)] || [] });
  });
  return { fields: f.map(x => x.h), descField: ctx.descCol ? ctx.mh[ctx.descCol - 1] : '',
           urlField: ctx.urlCol ? ctx.mh[ctx.urlCol - 1] : '', rows: rows };
}
/** payload = { partNo, isNew, values:{ header: value } }. Edits are copied to every model BOM that lists the part. */
function savePart(p) {
  return withLock_(() => {
    const sh = dataSS_().getSheetByName(SHEET_NAME);
    const ctx = readMaster_(), fields = partFields_(ctx);
    const partNo = String(p.partNo || '').replace(/^'/, '').trim();
    if (!PART_RE_.test(partNo)) throw uerr_('Part numbers may use letters, numbers, . _ - / and spaces (up to 60 characters).');
    const k = norm_(partNo), exists = k in ctx.idx;
    if (p.isNew && exists) throw uerr_(partNo + ' is already in PART_MASTER.');
    if (!p.isNew && !exists) throw uerr_(partNo + ' was not found in PART_MASTER.');
    const vals = (p.values && typeof p.values === 'object') ? p.values : {};
    const descH = ctx.descCol ? ctx.mh[ctx.descCol - 1] : '', urlH = ctx.urlCol ? ctx.mh[ctx.urlCol - 1] : '';
    if (descH && descH in vals && !String(vals[descH]).trim()) throw uerr_('Enter the part description.');
    if (p.isNew && descH && !String(vals[descH] || '').trim()) throw uerr_('Enter the part description.');
    if (urlH && String(vals[urlH] || '').trim() && !docLink_(vals[urlH])) throw uerr_('The image link must be an https link.');
    if (urlH && urlH in vals) vals[urlH] = docLink_(vals[urlH]);
    const imgCol = ctx.mh.findIndex(h => norm_(h) === 'IMAGE') + 1;
    const log = migratedSheet_('PART_LOG', PART_LOG_HDR), when = now_(), by = user_();
    if (p.isNew) {
      const row = ctx.mh.map((h, i) => i === COL_PART_NO - 1 ? partNo : (fields.some(x => x.i === i) && vals[h] != null ? vals[h] : ''));
      const r = Math.max(sh.getLastRow(), HEADER_ROW) + 1;
      sh.getRange(r, COL_PART_NO).setNumberFormat('@');
      sh.getRange(r, 1, 1, row.length).setValues([row]);
      if (imgCol && urlH && vals[urlH]) sh.getRange(r, imgCol).setFormula(imgFormula_(vals[urlH]));
      appendRows_(log, [[when, partNo, 'Created', '', String(vals[descH] || ''), by]]);
      return { partNo: partNo, created: true, changed: 0, models: 0 };
    }
    const i = ctx.idx[k], r = HEADER_ROW + 1 + i, changed = [];
    fields.forEach(x => {
      if (!(x.h in vals)) return;
      const nv = vals[x.h] == null ? '' : vals[x.h];
      if (txt_(ctx.mVals[i][x.i]) === String(nv).replace(/^'/, '')) return;
      changed.push({ h: x.h, i: x.i, old: txt_(ctx.mVals[i][x.i]), nv: nv });
    });
    if (!changed.length) return { partNo: partNo, created: false, changed: 0, models: 0 };
    changed.forEach(c => sh.getRange(r, c.i + 1).setValue(c.nv));
    const urlChanged = changed.some(c => c.h === urlH);
    if (urlChanged && imgCol) sh.getRange(r, imgCol).setFormula(imgFormula_(vals[urlH]));
    let models = 0;
    getModelSheets_().forEach(s => {
      const f = findInModel_(s, partNo, ctx.mh);
      if (!f) return;
      models++;
      const hs = headersOf_(s);
      changed.forEach(c => { const j = hs.findIndex(h => norm_(h) === norm_(c.h)); if (j >= 0) s.getRange(f.row, j + 1).setValue(c.nv); });
      if (urlChanged) { const j = hs.findIndex(h => norm_(h) === 'IMAGE'); if (j >= 0) s.getRange(f.row, j + 1).setFormula(imgFormula_(vals[urlH])); }
    });
    appendRows_(log, changed.map(c => [when, partNo, c.h, c.old, String(c.nv).replace(/^'/, ''), by]));
    return { partNo: partNo, created: false, changed: changed.length, models: models };
  });
}
/** Rewrites the Image column (=IMAGE of Image_URL) in PART_MASTER and every model BOM. */
function repairImages() {
  return withLock_(() => {
    const mh = headersOf_(dataSS_().getSheetByName(SHEET_NAME));
    let fixed = 0;
    [dataSS_().getSheetByName(SHEET_NAME)].concat(getModelSheets_()).forEach(s => {
      const hs = headersOf_(s), n = s.getLastRow() - HEADER_ROW;
      const ic = hs.findIndex(h => norm_(h) === 'IMAGE') + 1, uc = findCol_(hs, /image.*url|url.*image/i);
      if (n <= 0 || !ic || !uc) return;
      const urls = s.getRange(HEADER_ROW + 1, uc, n, 1).getValues();
      s.getRange(HEADER_ROW + 1, ic, n, 1).setFormulas(urls.map(u => [imgFormula_(u[0])]));
      fixed += urls.filter(u => docLink_(u[0])).length;
    });
    return { fixed: fixed };
  });
}
/** New model BOM from a part list; it is locked straight away as Rev 0 so BOM verification can use it. */
function createModel(p) {
  const name = String(p.newModel || '').trim();
  if (!(p.items || []).length) throw uerr_('Add at least one part with its QPS.');
  const r = saveBom({ newModel: name, items: p.items });
  if (r.missing && r.missing.length) {
    const ss = dataSS_(), sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
    throw uerr_('Not in PART_MASTER yet: ' + r.missing.join(', ') + '. Create these parts first.');
  }
  lockBom(name);
  return { model: name, parts: r.added };
}

/** PART_MASTER as a list (part, description, image link) for pickers. */
function masterList() {
  return getBomData();
}
function dedupe() { return withLock_(() => dedupePartMaster_()); }

// fn -> [write?, argument names]; anything else cannot be called
const API_ = {
  getControlData: [0, []], getVerifyModels: [0, []], getBomData: [0, []], getTodayStr: [0, []],
  getVerifyBom: [0, ['model']], getModelParts: [0, ['model']], getModelBomCsv: [0, ['model']],
  getEcnLog: [0, ['model']], getAllEcnNos: [0, ['model']], getEcnByNo: [0, ['ecnNo']],
  getEcnReportData: [0, ['model', 'ecnNo']], getBomRevisions: [0, ['model']],
  getBomAtRevision: [0, ['model', 'rev']], getBomRevisionCsv: [0, ['model', 'rev']],
  getPartInModel: [0, ['model', 'partNo']], getMasterPart: [0, ['partNo']],
  getAfterPartInfo: [0, ['model', 'partNo']], getPartsInModel: [0, ['model', 'partNos']],
  getMasterParts: [0, ['partNos']], getLotList: [0, ['model']], getLotBom: [0, ['verId']],
  getLotCsv: [0, ['verId']], compareTwoModels: [0, ['a', 'b']], compareAll: [0, []],
  getPartUsage: [0, ['partNo']],
  lockBom: [1, ['model']], unlockBom: [1, ['model', 'reason']], runEcnChange: [1, ['payload']],
  updateEcnMeta: [1, ['payload']], deleteEcn: [1, ['ecnNo']], deleteEcnRecordOnly: [1, ['ecnNo']],
  restoreBomVersion: [1, ['model', 'rev', 'reason']], mergeBomRevisions: [1, ['model', 'revs']],
  assignPart: [1, ['payload']], saveBom: [1, ['payload']], dedupe: [1, []],
  listParts: [0, []], savePart: [1, ['payload']], repairImages: [1, []], createModel: [1, ['payload']]
};
const FNS_ = {
  getControlData, getVerifyModels, getBomData, getTodayStr, getVerifyBom, getModelParts, getModelBomCsv,
  getEcnLog, getAllEcnNos, getEcnByNo, getEcnReportData, getBomRevisions, getBomAtRevision, getBomRevisionCsv,
  getPartInModel, getMasterPart, getAfterPartInfo, getPartsInModel, getMasterParts, getLotList, getLotBom,
  getLotCsv, compareTwoModels, compareAll, getPartUsage, lockBom, unlockBom, runEcnChange, updateEcnMeta,
  deleteEcn, deleteEcnRecordOnly, restoreBomVersion, mergeBomRevisions, assignPart, saveBom, dedupe,
  listParts, savePart, repairImages, createModel
};
// Functions that take a model name check it is a real model BOM first.
const MODEL_ARG_ = { model: 1, a: 1, b: 1 };

function clean_(v, depth) {
  if (depth > 6) return null;
  if (Array.isArray(v)) return v.slice(0, 2000).map(x => clean_(x, depth + 1));
  if (v && typeof v === 'object') {
    const o = {};
    Object.keys(v).slice(0, 50).forEach(k => { o[k] = clean_(v[k], depth + 1); });
    return o;
  }
  if (typeof v === 'string') return safe_(v, 2000);
  if (typeof v === 'number') return isFinite(v) ? v : null;
  return typeof v === 'boolean' ? v : (v == null ? '' : null);
}

function call(d) {
  const fn = String(d.fn || '');
  const spec = API_[fn];
  if (!spec || !FNS_[fn]) return { ok: false, error: 'Unknown BOM tool.' };
  const a = (d.args && typeof d.args === 'object' && !Array.isArray(d.args)) ? clean_(d.args, 0) : {};
  CUR_USER_ = String(d._session && (d._session.name || d._session.username) || '');
  try {
    const args = spec[1].map(name => {
      const v = a[name];
      if (MODEL_ARG_[name] && v !== '' && v != null) modelOk_(v);
      if (name === 'payload' && v && v.model && !v.newModel && fn !== 'savePart') modelOk_(v.model);
      if (name === 'payload' && v && v.newModel) {
        if (!NAME_RE_.test(String(v.newModel))) throw uerr_('Model names may use letters, numbers, spaces and - _ . / ( ) # & only.');
      }
      return v == null ? '' : v;
    });
    if (fn === 'assignPart') ((a.payload || {}).add || []).concat(((a.payload || {}).remove || []).map(m => ({ model: m })))
      .forEach(x => modelOk_(x.model));
    const data = FNS_[fn].apply(null, args);
    return { ok: true, write: !!spec[0], data: data === undefined ? null : data };
  } catch (e) {
    if (e && e.userMsg) return { ok: false, error: e.message };
    console.error('BOM ' + fn + ' failed: ' + (e && e.stack || e));
    return { ok: false, error: 'Could not complete that. Please try again, or ask an admin to check the Apps Script log.' };
  }
}
return { call: call, isWrite: fn => !!(API_[fn] && API_[fn][0]), repairImages: repairImages };
})();
function bomCall(d){ return BOMX.call(d); }
/* Run once from the editor: fixes the Image column (#REF! after the move into this spreadsheet). */
function repairBomImages(){ var r = BOMX.repairImages(); console.log('Images rewritten: ' + r.fixed); return r; }
