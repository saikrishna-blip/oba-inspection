/****************************************************
 * LUMIO Quality Management - Projector — Google Apps Script backend
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
// Modules an admin can give each user (Users sheet, column E "Access", comma separated).
// Admins always have every module. An empty cell means the original inspector set.
var MODULES        = ["bomverify","sampling","reports","createbatch","models","data","bommgmt"];
var DEFAULT_ACCESS = ["bomverify","sampling"];
// Actions any signed-in user may call (reading, own password)
var ALWAYS_ACTIONS = ["logout","whoami","getBatches","getModels","getBomList","getBomVerification","getLotUnits","changePassword"];
// recordRework (rework data + reoffer) belongs to the sampling module
// Actions unlocked by a module; everything not listed here or above is admin only
var ACTION_MODULE  = {getBom:"bomverify", submitBomVerification:"bomverify",
                      checkSerial:"sampling", submitLot:"sampling", recordRework:"sampling", setOfferedQty:"sampling",
                      createBatch:"createbatch",
                      saveModel:"models", deleteModel:"models",
                      scanDuplicates:"data", searchRecords:"data", removeDuplicates:"data", approveSerial:"data"};

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
      if(action === "submitLot" || action === "recordRework" || action === "holdLot" || action === "setOfferedQty") data.inspector = sess.name;
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
  return {bomverify:"BOM verification", sampling:"Sampling inspection", reports:"Reports", createbatch:"Create batches",
          models:"Models and black dot spec", data:"Data records", bommgmt:"BOM management"}[m] || m;
}
/* Who is signed in, and what they may use (the portal calls this on every refresh). */
function whoami(sess){
  var out = {ok:true, username:sess.username, name:sess.name, role:sess.role, access:sess.access || []};
  if(sess.role === "admin" || out.access.indexOf("bommgmt") > -1){
    out.partdbUrl = PARTDB_URL; out.bomSheetUrl = "https://docs.google.com/spreadsheets/d/" + BOM_SS_ID + "/edit";
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
  var rows = getOrCreateSheet("Master",[]).getDataRange().getValues(), sameLot = false;
  for(var i = 1; i < rows.length; i++){
    if(normSerial_(rows[i][4]) !== sn) continue;
    if(bn && String(rows[i][2]) === bn && String(rows[i][3]) === ln){ sameLot = true; continue; }
    return {ok:true, exists:true, batch:rows[i][2], lot:rows[i][3]};
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
 * BOM data lives in the BOM reference spreadsheet (BOM_SS_ID).
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
