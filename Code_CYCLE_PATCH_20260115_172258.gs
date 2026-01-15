
/****************************************************
 * Meeting Expense Claim – Code.gs (FULL) ✅ GitHub Pages + JSONP
 * Project: ระบบส่งข้อมูลเบิกเงินประชุมร้าน
 *
 * ✅ Uses Google Sheets as Master Data:
 *   - BRANCH_MASTER (branchCode, branchName, brand, districtManager, isActive)
 *   - ADMIN_MASTER  (adminCode, adminName, isActive)
 *   - CYCLE_MASTER  (cycleKey, cycleName, startDate, endDate, isOpen, note)
 *
 * ✅ Creates/uses CLAIMS sheet automatically (and adds brand & districtManager columns)
 * ✅ Stores images in Drive Folder (ANYONE_WITH_LINK view, so GitHub can preview)
 * ✅ Cross-domain calling from GitHub Pages via JSONP dispatcher (doGet)
 *
 * Deploy:
 *  - Deploy as Web App
 *  - Execute as: Me
 *  - Who has access: Anyone (or Anyone with the link)
 *
 ****************************************************/

const SPREADSHEET_ID   = "1wDMKb_TCd3o0Bn-PQTsS8k5k0MU_dGab8IWVZgGhVKw";
const DRIVE_FOLDER_ID  = "1rFc-oAyDfDaNoIwSPxOhSYt_n5modwZp";

// Sheet names
const SHEET_CLAIMS  = "CLAIMS";
const SHEET_BRANCH  = "BRANCH_MASTER";
const SHEET_ADMIN   = "ADMIN_MASTER";
const SHEET_CYCLE   = "CYCLE_MASTER";



/* =========================
   Cache Helpers (v2.0)
   - ScriptCache is fast and shared across users
   - Use TTL to speed up master lookups and list queries
========================= */
function cache_(){ return CacheService.getScriptCache(); }
function cacheKey_(parts){ return parts.map(p=>String(p)).join('|'); }
function cacheGet_(key){
  try{
    const v = cache_().get(key);
    return v ? JSON.parse(v) : null;
  }catch(e){ return null; }
}
function cacheSet_(key, obj, ttlSec){
  try{
    cache_().put(key, JSON.stringify(obj), ttlSec || 60);
  }catch(e){}
}
function cacheRemove_(key){
  try{ cache_().remove(key); }catch(e){}
}
function cacheClearPrefix_(prefix){
  // ScriptCache doesn't support prefix delete; keep a small index to clear common keys
  try{
    const idxKey = 'IDX:'+prefix;
    const idx = cacheGet_(idxKey) || [];
    idx.forEach(k=>cacheRemove_(k));
    cacheRemove_(idxKey);
  }catch(e){}
}
function cacheIndexAdd_(prefix, key){
  try{
    const idxKey = 'IDX:'+prefix;
    const idx = cacheGet_(idxKey) || [];
    if(idx.indexOf(key) < 0){
      idx.push(key);
      cacheSet_(idxKey, idx, 3600);
    }
  }catch(e){}
}

/* =========================
   JSONP Dispatcher (Web App)
   =========================
   Call format:
     /exec?fn=apiLogin&args=BASE64(JSON_ARRAY)&callback=cb
   Example args:
     []  -> "W10="
     ["5001"] -> base64(JSON.stringify(["5001"]))
========================= */
function doGet(e) {
  e = e || {};
  const p = (e && e.parameter) ? e.parameter : {};

  // JSONP API mode
  if (p.fn) {
    const fn = String(p.fn);
    const cb = String(p.callback || "callback");

    let args = [];
    try {
      if (p.args) {
        const json = Utilities.newBlob(Utilities.base64Decode(p.args)).getDataAsString("utf-8");
        args = JSON.parse(json);
        if (!Array.isArray(args)) args = [args];
      }
    } catch (err) {
      return jsonp_(cb, { ok: false, error: "BAD_ARGS", detail: String(err) });
    }

    try {
      const target = this[fn];
      if (typeof target !== "function") {
        return jsonp_(cb, { ok: false, error: "FN_NOT_FOUND", fn });
      }
      const result = target.apply(null, args);
      return jsonp_(cb, normalizeResult_(result));
    } catch (err) {
      return jsonp_(cb, { ok: false, error: "EXCEPTION", detail: String(err) });
    }
  }

  // Not API call: lightweight healthcheck
  return ContentService
    .createTextOutput("OK")
    .setMimeType(ContentService.MimeType.TEXT);
}

function jsonp_(callbackName, obj) {
  const cb = String(callbackName || "callback").replace(/[^\w$]/g, "");
  const out = `${cb}(${JSON.stringify(obj)});`;
  return ContentService
    .createTextOutput(out)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function normalizeResult_(result) {
  if (result && typeof result === "object" && Object.prototype.hasOwnProperty.call(result, "ok")) return result;
  return { ok: true, data: result };
}

/* =========================
   Helpers: Spreadsheet/Sheets
========================= */
function ss_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function sheet_(name) {
  const ss = ss_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`SHEET_NOT_FOUND:${name}`);
  return sh;
}

/* =========================
   Helpers: Date utilities
========================= */
function startOfDay_(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}

function endOfDay_(d) {
  const x = new Date(d);
  x.setHours(23, 59, 59, 999);
  return x;
}

function nowIso_() {
  return new Date().toISOString();
}

/* =========================
   Master Data: Branch
   BRANCH_MASTER columns:
     A branchCode (4 digits)
     B branchName
     C brand
     D districtManager
     E isActive (TRUE/FALSE)
========================= */
function getBranchByCode_(branchCode) {
  const code = String(branchCode || "").trim();
  if (!/^\d{4}$/.test(code)) return null;

  const ck = cacheKey_(['BRANCH', code]);
  const cached = cacheGet_(ck);
  if (cached) return cached;

  const sh = sheet_(SHEET_BRANCH);
  const last = sh.getLastRow();
  if (last < 2) return null;

  const rows = sh.getRange(2, 1, last - 1, 5).getValues();
  const hit = rows.find(r => String(r[0]).trim() === code);
  if (!hit) return null;

  const isActive = String(hit[4]).toUpperCase() === "TRUE";
  if (!isActive) return null;

  const obj = {
    branchCode: code,
    branchName: String(hit[1] || "").trim(),
    brand: String(hit[2] || "").trim(),
    districtManager: String(hit[3] || "").trim(),
    isActive: true
  };

  cacheSet_(ck, obj, 300); // 5 นาที
  cacheIndexAdd_('BRANCH', ck);
  return obj;
}

/* =========================
   Master Data: Admin
   ADMIN_MASTER columns:
     A adminCode
     B adminName
     C isActive (TRUE/FALSE)
========================= */
function getAdminByCode_(adminCode) {
  const code = String(adminCode || "").trim();
  if (!code) return null;

  const ck = cacheKey_(['ADMIN', code]);
  const cached = cacheGet_(ck);
  if (cached) return cached;

  const sh = sheet_(SHEET_ADMIN);
  const last = sh.getLastRow();
  if (last < 2) return null;

  const rows = sh.getRange(2, 1, last - 1, 3).getValues();
  const hit = rows.find(r => String(r[0]).trim() === code);
  if (!hit) return null;

  const isActive = String(hit[2]).toUpperCase() === "TRUE";
  if (!isActive) return null;

  const obj = {
    adminCode: code,
    adminName: String(hit[1] || "").trim(),
    isActive: true
  };

  cacheSet_(ck, obj, 300); // 5 นาที
  cacheIndexAdd_('ADMIN', ck);
  return obj;
}

/* =========================
   Master Data: Cycle (Primary: date range)
   CYCLE_MASTER columns:
     A cycleKey
     B cycleName
     C startDate (Date)
     D endDate (Date)
     E isOpen (TRUE/FALSE)
     F note (optional)
========================= */
function getCycleFromSheet_() {
  const ck = cacheKey_(['CYCLE', 'NOW']);
  const cached = cacheGet_(ck);
  if (cached) return cached;

  const sh = sheet_(SHEET_CYCLE);
  const last = sh.getLastRow();
  if (last < 2) return null;

  const rows = sh.getRange(2, 1, last - 1, 6).getValues();
  const now = new Date();

  const hit = rows.find(r => {
    const s = (r[2] instanceof Date) ? r[2] : (r[2] ? new Date(r[2]) : null);
    const e = (r[3] instanceof Date) ? r[3] : (r[3] ? new Date(r[3]) : null);
    if (!s || !e || isNaN(s.getTime()) || isNaN(e.getTime())) return false;
    // Prefer only cycles that are within range (date-driven)
    return now >= startOfDay_(s) && now <= endOfDay_(e);
  });

  if (!hit) return null;

  const s = (hit[2] instanceof Date) ? hit[2] : new Date(hit[2]);
  const e = (hit[3] instanceof Date) ? hit[3] : new Date(hit[3]);
  const isOpen = String(hit[4]).toUpperCase() === "TRUE";

  const obj = {
    cycleKey: String(hit[0] || "").trim(),
    cycleName: String(hit[1] || "").trim(),
    startDate: s,
    endDate: e,
    isOpen: isOpen,
    note: String(hit[5] || "").trim()
  };

  cacheSet_(ck, obj, 60); // 60 วินาที
  cacheIndexAdd_('CYCLE', ck);
  return obj;
}

function fallbackCycle_() {
  // Fallback: current month (Thailand year)
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const thYear = yyyy + 543;
  const thMonths = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];

  return {
    cycleKey: `${yyyy}-${mm}`,
    cycleName: `${thMonths[d.getMonth()]} ${thYear}`,
    startDate: null,
    endDate: null,
    isOpen: true,
    note: ""
  };
}

function getCycle_() {
  const c = getCycleFromSheet_();
  return c || fallbackCycle_();
}

/* =========================
   CLAIMS Sheet: ensure + columns
   Creates CLAIMS if missing, adds columns brand & districtManager.
========================= */
function ensureClaimsSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_CLAIMS);
  if (!sh) sh = ss.insertSheet(SHEET_CLAIMS);

  const header = [
    "id",
    "submittedAt",
    "cycleKey",
    "cycleName",
    "branchCode",
    "branchName",
    "brand",
    "districtManager",
    "claimantName",
    "amount",
    "note",
    "status",
    "imageFileIds",
    "imageUrls"
  ];

  const isEmpty = sh.getLastRow() === 0 || sh.getLastColumn() === 0;
  if (isEmpty) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
    return sh;
  }

  const row1 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || "").trim());
  const existing = new Set(row1.filter(Boolean));
  const toAdd = header.filter(h => !existing.has(h));
  if (toAdd.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  }
  return sh;
}

function headerMap_(sh) {
  const row1 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  row1.forEach((v, i) => {
    const key = String(v || "").trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function makeId_() {
  const ts = new Date().getTime();
  const rnd = Math.floor(Math.random() * 1e6).toString().padStart(6, "0");
  return `CLM-${ts}-${rnd}`;
}

function toNumber_(x) {
  const n = Number(x);
  return Number.isFinite(n) ? n : 0;
}

function findClaimRowById_(sh, id) {
  const map = headerMap_(sh);
  const idCol = map["id"];
  if (!idCol) throw new Error("CLAIMS_HEADER_MISSING:id");

  const last = sh.getLastRow();
  if (last < 2) return -1;

  const ids = sh.getRange(2, idCol, last - 1, 1).getValues().map(r => String(r[0]));
  const idx = ids.indexOf(String(id));
  return (idx < 0) ? -1 : (2 + idx);
}

function findClaimRowByIdFast_(sh, id){
  const map = headerMap_(sh);
  const idCol = map["id"];
  if (!idCol) throw new Error("CLAIMS_HEADER_MISSING:id");
  const last = sh.getLastRow();
  if (last < 2) return -1;

  try{
    const tf = sh.getRange(2, idCol, last - 1, 1)
      .createTextFinder(String(id))
      .matchEntireCell(true)
      .findNext();
    return tf ? tf.getRow() : -1;
  }catch(e){
    // fallback to slow scan
    return findClaimRowById_(sh, id);
  }
}



function readClaimRow_(sh, rowNumber) {
  const map = headerMap_(sh);
  const lastCol = sh.getLastColumn();
  const row = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];

  function g(name) {
    const c = map[name];
    return c ? row[c - 1] : "";
  }

  return {
    id: String(g("id") || ""),
    submittedAt: String(g("submittedAt") || ""),
    cycleKey: String(g("cycleKey") || ""),
    cycleName: String(g("cycleName") || ""),
    branchCode: String(g("branchCode") || ""),
    branchName: String(g("branchName") || ""),
    brand: String(g("brand") || ""),
    districtManager: String(g("districtManager") || ""),
    claimantName: String(g("claimantName") || ""),
    amount: toNumber_(g("amount")),
    note: String(g("note") || ""),
    status: String(g("status") || ""),
    imageFileIds: String(g("imageFileIds") || "").split("|").filter(Boolean),
    imageUrls: String(g("imageUrls") || "").split("|").filter(Boolean)
  };
}

/* =========================
   Drive: Images
========================= */
function folder_() {
  return DriveApp.getFolderById(DRIVE_FOLDER_ID);
}

function saveImages_(claimId, images) {
  const outIds = [];
  const outUrls = [];

  if (!images || !Array.isArray(images) || images.length === 0) {
    return { fileIds: [], urls: [] };
  }

  const folder = folder_();

  images.forEach((img, idx) => {
    const name = String(img && img.name ? img.name : `image_${idx + 1}.jpg`);
    const mimeType = String(img && img.mimeType ? img.mimeType : "image/jpeg");
    const b64 = String(img && img.base64 ? img.base64 : "");

    if (!b64) return;

    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mimeType, `${claimId}__${name}`);
    const file = folder.createFile(blob);

    // GitHub viewable
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    outIds.push(file.getId());
    outUrls.push(file.getUrl());
  });

  return { fileIds: outIds, urls: outUrls };
}

function trashFiles_(fileIds) {
  let count = 0;
  (fileIds || []).forEach(id => {
    if (!id) return;
    try {
      DriveApp.getFileById(id).setTrashed(true);
      count++;
    } catch (e) {}
  });
  return count;
}

/* =========================
   API: Login (Branch/Admin)
========================= */
function apiLogin(code) {
  const c = String(code || "").trim();
  if (!c) return { ok: false, error: "EMPTY_CODE" };

  // Branch (4 digits)
  if (/^\d{4}$/.test(c)) {
    const b = getBranchByCode_(c);
    if (!b) return { ok: false, error: "BRANCH_NOT_ALLOWED" };
    return {
      ok: true,
      role: "branch",
      user: {
        branchCode: b.branchCode,
        branchName: b.branchName,
        brand: b.brand,
        districtManager: b.districtManager,
        name: b.branchName || b.branchCode
      }
    };
  }

  // Admin
  const a = getAdminByCode_(c);
  if (!a) return { ok: false, error: "ADMIN_NOT_ALLOWED" };
  return {
    ok: true,
    role: "admin",
    user: {
      adminCode: a.adminCode,
      adminName: a.adminName,
      name: a.adminName || "Administrator"
    }
  };
}

/* =========================
   API: Cycle
========================= */
function apiGetCurrentCycle() {
  const c = getCycle_();
  return {
    ok: true,
    cycle: {
      cycleKey: c.cycleKey,
      cycleName: c.cycleName,
      startDate: c.startDate ? c.startDate.toISOString() : null,
      endDate: c.endDate ? c.endDate.toISOString() : null,
      isOpen: !!c.isOpen,
      note: c.note || ""
    }
  };
}

/* =========================
   API: List Claims
========================= */
/* =========================
   API: List Claims (LITE) – FAST
   - Returns NO imageUrls to keep payload small
   - Reads only last N rows (SCAN_ROWS) for speed (most recent data)
   - Supports limit/offset
========================= */
const CLAIMS_SCAN_ROWS = 800;          // อ่านย้อนหลังล่าสุด 800 แถว (ปรับได้)
const CLAIMS_LIMIT_DEF = 200;          // จำนวนรายการเริ่มต้น

function toInt_(x, def){
  const n = parseInt(String(x||''), 10);
  return Number.isFinite(n) ? n : def;
}

function apiListClaimsLite(mode, requesterCode, filters) {
  const m = String(mode || "").toLowerCase();
  const code = String(requesterCode || "").trim();

  let branch = null;
  let admin = null;

  if (m === "branch") {
    branch = getBranchByCode_(code);
    if (!branch) return { ok: false, error: "INVALID_BRANCH" };
  } else if (m === "admin") {
    admin = getAdminByCode_(code);
    if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };
  } else {
    return { ok: false, error: "BAD_MODE" };
  }

  const sh = ensureClaimsSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, claims: [], meta:{ returned:0, totalMatched:0 } };

  const f = filters || {};
  const limit = Math.max(1, Math.min(500, toInt_(f.limit, CLAIMS_LIMIT_DEF)));
  const offset = Math.max(0, toInt_(f.offset, 0));

  const fCycleKey = f.cycleKey ? String(f.cycleKey).trim() : "";
  const fBrand = f.brand ? String(f.brand).trim() : "";
  const fDM = f.districtManager ? String(f.districtManager).trim() : "";
  const fStatus = f.status ? String(f.status).trim() : "";
  const fBranchCode = f.branchCode ? String(f.branchCode).trim() : "";

  // Cache key (include lastRow to auto-bust when new rows appended)
  const cacheKey = cacheKey_(['CLAIMS_LITE', m, code, fCycleKey, fBrand, fDM, fStatus, fBranchCode, limit, offset, last]);
  const cached = cacheGet_(cacheKey);
  if (cached) return cached;

  const map = headerMap_(sh);
  const lastCol = sh.getLastColumn();

  // Read only recent rows for speed
  const available = last - 1;
  const scan = Math.min(CLAIMS_SCAN_ROWS, available);
  const startRow = Math.max(2, last - scan + 1);
  const values = sh.getRange(startRow, 1, scan, lastCol).getValues();

  function cell(row, name) {
    const c = map[name];
    return c ? row[c - 1] : "";
  }

  let claims = values
    .map(row => {
      const imageFileIds = String(cell(row, "imageFileIds") || "");
      const imageCount = imageFileIds ? imageFileIds.split("|").filter(Boolean).length : 0;
      return {
        id: String(cell(row, "id") || ""),
        submittedAt: String(cell(row, "submittedAt") || ""),
        cycleKey: String(cell(row, "cycleKey") || ""),
        cycleName: String(cell(row, "cycleName") || ""),
        branchCode: String(cell(row, "branchCode") || ""),
        branchName: String(cell(row, "branchName") || ""),
        brand: String(cell(row, "brand") || ""),
        districtManager: String(cell(row, "districtManager") || ""),
        claimantName: String(cell(row, "claimantName") || ""),
        amount: toNumber_(cell(row, "amount")),
        status: String(cell(row, "status") || ""),
        imageCount: imageCount
      };
    })
    .filter(c => {
      if (m === "branch" && c.branchCode !== branch.branchCode) return false;
      if (fBranchCode && c.branchCode !== fBranchCode) return false;
      if (fCycleKey && c.cycleKey !== fCycleKey) return false;
      if (fBrand && c.brand !== fBrand) return false;
      if (fDM && c.districtManager !== fDM) return false;
      if (fStatus && c.status !== fStatus) return false;
      return true;
    });

  // sort newest first
  claims.sort((a, b) => (b.submittedAt || "").localeCompare(a.submittedAt || ""));

  const totalMatched = claims.length;
  claims = claims.slice(offset, offset + limit);

  const result = {
    ok: true,
    claims,
    meta: {
      returned: claims.length,
      totalMatched: totalMatched,
      offset: offset,
      limit: limit,
      scannedRows: scan,
      truncatedScan: startRow > 2
    }
  };

  cacheSet_(cacheKey, result, 15); // 15 วินาทีพอ (ลดโหลดชีตซ้ำ)
  cacheIndexAdd_('CLAIMS', cacheKey);

  return result;
}

/* Backward compatible alias (older HTML may call apiListClaims) */
function apiListClaims(mode, requesterCode, filters){
  return apiListClaimsLite(mode, requesterCode, filters);
}

/* =========================
   API: Claim Detail (Load images on demand)
========================= */
function apiGetClaimDetail(mode, requesterCode, id){
  const m = String(mode || "").toLowerCase();
  const code = String(requesterCode || "").trim();

  let branch = null;
  let admin = null;

  if (m === "branch") {
    branch = getBranchByCode_(code);
    if (!branch) return { ok: false, error: "INVALID_BRANCH" };
  } else if (m === "admin") {
    admin = getAdminByCode_(code);
    if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };
  } else {
    return { ok: false, error: "BAD_MODE" };
  }

  const claimId = String(id || "").trim();
  if (!claimId) return { ok:false, error:"MISSING_ID" };

  const sh = ensureClaimsSheet_();
  const rowNumber = findClaimRowByIdFast_(sh, claimId);
  if (rowNumber < 0) return { ok:false, error:"NOT_FOUND" };

  const claim = readClaimRow_(sh, rowNumber);

  // branch can only read its own claim
  if(m === 'branch' && claim.branchCode !== branch.branchCode) return { ok:false, error:"FORBIDDEN" };

  return { ok:true, claim: claim };
}


/* =========================
   API: Submit Claim (Branch)
========================= */
function apiSubmitClaim(payload) {
  const p = payload || {};
  const branchCode = String(p.branchCode || "").trim();
  const branch = getBranchByCode_(branchCode);
  if (!branch) return { ok: false, error: "INVALID_BRANCH" };

  const cycle = getCycle_();

  // Rule: date range is primary + must be isOpen=TRUE
  if (!cycle.isOpen) return { ok: false, error: "CYCLE_CLOSED", cycle };

  if (cycle.startDate && cycle.endDate) {
    const now = new Date();
    if (now < startOfDay_(cycle.startDate) || now > endOfDay_(cycle.endDate)) {
      return { ok: false, error: "OUT_OF_CYCLE_DATE", cycle };
    }
  }

  const claimantName = String(p.claimantName || "").trim();
  const amount = toNumber_(p.amount);
  const note = String(p.note || "").trim();

  if (!claimantName) return { ok: false, error: "MISSING_CLAIMANT_NAME" };
  if (!(amount > 0)) return { ok: false, error: "INVALID_AMOUNT" };

  const claimId = makeId_();
  const submittedAt = nowIso_();

  const images = saveImages_(claimId, p.images || []);

  const sh = ensureClaimsSheet_();
  const map = headerMap_(sh);

  const lastCol = sh.getLastColumn();
  const row = Array(lastCol).fill("");

  function set(name, val) {
    const col = map[name];
    if (col) row[col - 1] = val;
  }

  set("id", claimId);
  set("submittedAt", submittedAt);
  set("cycleKey", String(p.cycleKey || cycle.cycleKey));
  set("cycleName", String(p.cycleName || cycle.cycleName));
  set("branchCode", branch.branchCode);
  set("branchName", branch.branchName);
  set("brand", branch.brand);
  set("districtManager", branch.districtManager);
  set("claimantName", claimantName);
  set("amount", amount);
  set("note", note);
  set("status", "รอดำเนินการ");
  set("imageFileIds", images.fileIds.join("|"));
  set("imageUrls", images.urls.join("|"));

  sh.appendRow(row);

  try{ cacheClearPrefix_('CLAIMS'); }catch(e){}

  return {
    ok: true,
    claim: { id: claimId },
    cycle: {
      cycleKey: cycle.cycleKey,
      cycleName: cycle.cycleName,
      startDate: cycle.startDate ? cycle.startDate.toISOString() : null,
      endDate: cycle.endDate ? cycle.endDate.toISOString() : null,
      isOpen: !!cycle.isOpen
    },
    branch: branch
  };
}

/* =========================
   API: Update Claim (Admin)
========================= */
function apiUpdateClaim(adminCode, id, patch) {
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };

  const claimId = String(id || "").trim();
  if (!claimId) return { ok: false, error: "MISSING_ID" };

  const sh = ensureClaimsSheet_();
  const rowNumber = findClaimRowByIdFast_(sh, claimId);
  if (rowNumber < 0) return { ok: false, error: "NOT_FOUND" };

  const map = headerMap_(sh);
  const p = patch || {};

  const updates = [];
  if (Object.prototype.hasOwnProperty.call(p, "amount")) updates.push({ col: map["amount"], val: toNumber_(p.amount) });
  if (Object.prototype.hasOwnProperty.call(p, "note")) updates.push({ col: map["note"], val: String(p.note || "") });
  if (Object.prototype.hasOwnProperty.call(p, "status")) updates.push({ col: map["status"], val: String(p.status || "") });

  updates.forEach(u => {
    if (u.col) sh.getRange(rowNumber, u.col).setValue(u.val);
  });

  try{ cacheClearPrefix_('CLAIMS'); }catch(e){}
  return { ok: true };
}

/* =========================
   API: Delete Claim (Admin)
========================= */
function apiDeleteClaim(adminCode, id) {
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };

  const claimId = String(id || "").trim();
  if (!claimId) return { ok: false, error: "MISSING_ID" };

  const sh = ensureClaimsSheet_();
  const rowNumber = findClaimRowByIdFast_(sh, claimId);
  if (rowNumber < 0) return { ok: false, error: "NOT_FOUND" };

  const claim = readClaimRow_(sh, rowNumber);
  const deletedFileCount = trashFiles_(claim.imageFileIds);

  sh.deleteRow(rowNumber);

  try{ cacheClearPrefix_('CLAIMS'); }catch(e){}

  return { ok: true, deletedFileCount: deletedFileCount };
}

/* =========================
   OPTIONAL: Masters (Admin)
========================= */
function apiGetMasters(adminCode) {
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };

  const shB = sheet_(SHEET_BRANCH);
  const lastB = shB.getLastRow();
  const branches = (lastB < 2) ? [] : shB.getRange(2, 1, lastB - 1, 5).getValues()
    .filter(r => String(r[4]).toUpperCase() === "TRUE")
    .map(r => ({
      branchCode: String(r[0]).trim(),
      branchName: String(r[1] || "").trim(),
      brand: String(r[2] || "").trim(),
      districtManager: String(r[3] || "").trim()
    }));

  const shA = sheet_(SHEET_ADMIN);
  const lastA = shA.getLastRow();
  const admins = (lastA < 2) ? [] : shA.getRange(2, 1, lastA - 1, 3).getValues()
    .filter(r => String(r[2]).toUpperCase() === "TRUE")
    .map(r => ({
      adminCode: String(r[0]).trim(),
      adminName: String(r[1] || "").trim()
    }));

  const shC = sheet_(SHEET_CYCLE);
  const lastC = shC.getLastRow();
  const cycles = (lastC < 2) ? [] : shC.getRange(2, 1, lastC - 1, 6).getValues()
    .map(r => ({
      cycleKey: String(r[0] || "").trim(),
      cycleName: String(r[1] || "").trim(),
      startDate: (r[2] instanceof Date) ? r[2].toISOString() : String(r[2] || ""),
      endDate: (r[3] instanceof Date) ? r[3].toISOString() : String(r[3] || ""),
      isOpen: String(r[4]).toUpperCase() === "TRUE",
      note: String(r[5] || "").trim()
    }));

  return { ok: true, branches, admins, cycles };
}


/* =========================
   ADMIN: Cycle Management (FINAL)
   - Allow Administrator to create/update 12 months cycles
   - Allow open/close per cycle
   - Branches will reference these settings via apiGetCurrentCycle()
========================= */

function ensureCycleSheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_CYCLE);
  if (!sh) sh = ss.insertSheet(SHEET_CYCLE);

  const header = ["cycleKey", "cycleName", "startDate", "endDate", "isOpen", "note"];
  const isEmpty = sh.getLastRow() === 0 || sh.getLastColumn() === 0;
  if (isEmpty) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
    return sh;
  }

  const row1 = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || "").trim());
  const existing = new Set(row1.filter(Boolean));
  const toAdd = header.filter(h => !existing.has(h));
  if (toAdd.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
  }
  return sh;
}

function parseDate_(x) {
  if (!x) return null;
  if (x instanceof Date) return x;
  const s = String(x).trim();
  if (!s) return null;
  // accept yyyy-mm-dd or ISO
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function thMonthShort_(mIdx) {
  const th = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
  return th[mIdx] || "";
}

function makeCycleNameTH_(d) {
  const yyyy = d.getFullYear();
  const thYear = yyyy + 543;
  return `${thMonthShort_(d.getMonth())} ${thYear}`;
}

function monthStart_(y, m) {
  return new Date(y, m, 1, 0, 0, 0, 0);
}
function monthEnd_(y, m) {
  return new Date(y, m + 1, 0, 23, 59, 59, 999);
}

/* Admin: list cycles (for management UI) */
function apiListCycles(adminCode) {
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };

  const sh = ensureCycleSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, cycles: [] };

  const rows = sh.getRange(2, 1, last - 1, 6).getValues();
  const cycles = rows
    .map(r => ({
      cycleKey: String(r[0] || "").trim(),
      cycleName: String(r[1] || "").trim(),
      startDate: (r[2] instanceof Date) ? r[2].toISOString() : (parseDate_(r[2]) ? parseDate_(r[2]).toISOString() : ""),
      endDate: (r[3] instanceof Date) ? r[3].toISOString() : (parseDate_(r[3]) ? parseDate_(r[3]).toISOString() : ""),
      isOpen: String(r[4]).toUpperCase() === "TRUE",
      note: String(r[5] || "").trim()
    }))
    .filter(c => c.cycleKey);

  cycles.sort((a, b) => String(a.cycleKey).localeCompare(String(b.cycleKey)));
  return { ok: true, cycles };
}

/* Admin: upsert cycle */
function apiUpsertCycle(adminCode, cycle) {
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };

  const c = cycle || {};
  const cycleKey = String(c.cycleKey || "").trim();
  if (!cycleKey) return { ok: false, error: "MISSING_CYCLE_KEY" };

  const cycleName = String(c.cycleName || "").trim();
  const s = parseDate_(c.startDate);
  const e = parseDate_(c.endDate);
  if (!s || !e) return { ok: false, error: "INVALID_DATE" };

  const isOpen = !!c.isOpen;
  const note = String(c.note || "").trim();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureCycleSheet_();
    const last = sh.getLastRow();

    // read keys (A)
    const keys = (last < 2) ? [] : sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || "").trim());
    const idx = keys.indexOf(cycleKey);

    const row = [cycleKey, (cycleName || makeCycleNameTH_(s)), s, e, isOpen, note];

    if (idx >= 0) {
      sh.getRange(2 + idx, 1, 1, 6).setValues([row]);
    } else {
      sh.appendRow(row);
    }

    try { cacheClearPrefix_('CYCLE'); } catch (e2) {}
    return { ok: true };
  } finally {
    try { lock.releaseLock(); } catch (e3) {}
  }
}

/* Admin: generate 12 months (current month + next 11) */
function apiGenerate12Months(adminCode) {
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok: false, error: "UNAUTHORIZED_ADMIN" };

  const now = new Date();
  const baseY = now.getFullYear();
  const baseM = now.getMonth();

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ensureCycleSheet_();
    const last = sh.getLastRow();
    const existingKeys = (last < 2) ? new Set() : new Set(sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0] || "").trim()).filter(Boolean));

    let created = 0;
    for (let i = 0; i < 12; i++) {
      const y = baseY + Math.floor((baseM + i) / 12);
      const m = (baseM + i) % 12;
      const start = monthStart_(y, m);
      const end = monthEnd_(y, m);
      const key = `${y}-${String(m + 1).padStart(2, '0')}`;
      if (existingKeys.has(key)) continue;

      const name = makeCycleNameTH_(start);
      const isOpen = (i === 0); // เปิดเดือนปัจจุบันไว้ก่อน
      sh.appendRow([key, name, start, end, isOpen, ""]);
      created++;
    }

    try { cacheClearPrefix_('CYCLE'); } catch (e2) {}
    return { ok: true, created };
  } finally {
    try { lock.releaseLock(); } catch (e3) {}
  }
}

/* Backward compatibility: apiGetMasters already returns cycles; keep it. */



/* =========================
   Cycle Date Parse (Safe)
========================= */
function parseYmd_(s){
  const str = String(s || "").trim();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str);
  if(!m) return null;
  const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
  const dt = new Date(y, mo-1, d, 0, 0, 0, 0);
  return isNaN(dt.getTime()) ? null : dt;
}




/* =========================
   Cycle Columns Mapping (Robust)
   - Supports English headers and Thai headers
   - Normalizes header names: lower + remove spaces/_/-
========================= */
function normKey_(s){
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\s_\-]+/g, "");
}

function cycleCols_(sh){
  const row1 = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const cols = {};
  row1.forEach((v, idx)=>{
    const raw = String(v || "").trim();
    if(!raw) return;
    cols[normKey_(raw)] = idx + 1;
  });

  function pick(keys, fallback){
    for(const k of keys){
      const c = cols[normKey_(k)];
      if(c) return c;
    }
    return fallback || null;
  }

  return {
    cycleKey: pick(["cycleKey","รอบ(key)","รอบkey","รอบ","key"], 1),
    cycleName: pick(["cycleName","ชื่อรอบ","รอบชื่อ","cycle"], 2),
    startDate: pick(["startDate","วันเริ่มต้น","วันเริ่ม","start"], 3),
    endDate: pick(["endDate","วันสิ้นสุด","วันจบ","end"], 4),
    isOpen: pick(["isOpen","เปิดใช้งาน","เปิดใช้งาน?","เปิดปิด","สถานะ","open"], 5),
    note: pick(["note","หมายเหตุ"], 6),
    lastCol: sh.getLastColumn()
  };
}

function toBool_(v){
  if (typeof v === "boolean") return v;
  const s = String(v || "").trim().toLowerCase();
  if (!s) return false;
  return (s === "true" || s === "1" || s === "yes" || s === "y" || s === "on" || s === "เปิด" || s === "ใช่");
}


/* =========================
   API: List Cycles (Admin) – Robust
========================= */
function apiListCycles(adminCode){
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok:false, error:"UNAUTHORIZED_ADMIN" };

  const sh = ensureCycleSheet_();
  const cols = cycleCols_(sh);
  const last = sh.getLastRow();
  if (last < 2) return { ok:true, cycles:[] };

  const values = sh.getRange(2, 1, last-1, cols.lastCol).getValues();

  const cycles = values.map(r=>{
    const s = r[cols.startDate-1];
    const e = r[cols.endDate-1];
    return {
      cycleKey: String(r[cols.cycleKey-1] || "").trim(),
      cycleName: String(r[cols.cycleName-1] || "").trim(),
      startDate: (s instanceof Date) ? s.toISOString() : (String(s||"").trim() || null),
      endDate:   (e instanceof Date) ? e.toISOString() : (String(e||"").trim() || null),
      isOpen: toBool_(r[cols.isOpen-1]),
      note: String(r[cols.note-1] || "").trim()
    };
  }).filter(c=>c.cycleKey);

  cycles.sort((a,b)=>String(a.cycleKey).localeCompare(String(b.cycleKey)));
  return { ok:true, cycles: cycles };
}


/* =========================
   API: Upsert Cycles (Admin) – Batch (Robust)
   - Accepts isOpen as boolean or string
   - Writes TRUE/FALSE consistently
   - Optional: enforce only one open cycle (latest open wins)
========================= */
function apiUpsertCycles(adminCode, cycles){
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok:false, error:"UNAUTHORIZED_ADMIN" };

  const list = Array.isArray(cycles) ? cycles : [];
  if (!list.length) return { ok:false, error:"EMPTY_CYCLES" };

  const sh = ensureCycleSheet_();
  const cols = cycleCols_(sh);

  // Index existing rows by cycleKey
  const last = sh.getLastRow();
  const existingKeys = (last < 2) ? [] :
    sh.getRange(2, cols.cycleKey, last-1, 1).getValues().map(r=>String(r[0]||"").trim());
  const rowByKey = {};
  existingKeys.forEach((k,i)=>{ if(k) rowByKey[k] = 2+i; });

  // Normalize input
  const cleaned = list.map(o=>{
    const cycleKey = String(o.cycleKey || "").trim();
    if(!cycleKey) return null;
    return {
      cycleKey,
      cycleName: String(o.cycleName || cycleKey).trim(),
      startDate: (o.startDate instanceof Date) ? o.startDate : (parseYmd_(o.startDate) || (String(o.startDate||"").trim() || "")),
      endDate:   (o.endDate instanceof Date) ? o.endDate : (parseYmd_(o.endDate)   || (String(o.endDate||"").trim()   || "")),
      isOpen: toBool_(o.isOpen),
      note: String(o.note || "").trim()
    };
  }).filter(Boolean);

  if(!cleaned.length) return { ok:false, error:"EMPTY_CYCLES" };
function writeRow_(rowNum, c){
    const row = sh.getRange(rowNum, 1, 1, cols.lastCol).getValues()[0];
    row[cols.cycleKey-1] = c.cycleKey;
    row[cols.cycleName-1] = c.cycleName;
    row[cols.startDate-1] = c.startDate;
    row[cols.endDate-1] = c.endDate;
    row[cols.isOpen-1] = c.isOpen; // boolean -> checkbox/TRUE
    row[cols.note-1] = c.note;
    sh.getRange(rowNum, 1, 1, cols.lastCol).setValues([row]);
  }

  cleaned.forEach(c=>{
    const rowNum = rowByKey[c.cycleKey];
    if(rowNum){
      writeRow_(rowNum, c);
    }else{
      // append new row
      const row = Array(cols.lastCol).fill("");
      row[cols.cycleKey-1] = c.cycleKey;
      row[cols.cycleName-1] = c.cycleName;
      row[cols.startDate-1] = c.startDate;
      row[cols.endDate-1] = c.endDate;
      row[cols.isOpen-1] = c.isOpen;
      row[cols.note-1] = c.note;
      sh.appendRow(row);
    }
  });

  SpreadsheetApp.flush();
  try{ cacheClearPrefix_('CYCLE'); }catch(e){}
  return { ok:true };
}


function apiDeleteCycles(adminCode, cycleKeys){
  const admin = getAdminByCode_(adminCode);
  if (!admin) return { ok:false, error:"UNAUTHORIZED" };

  const keys = (cycleKeys||[]).map(String);
  const sh = ensureCycleSheet_();
  const data = sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat();

  let del = [];
  data.forEach((k,i)=>{ if(keys.includes(k)) del.push(i+2); });
  del.reverse().forEach(r=>sh.deleteRow(r));

  return { ok:true, deleted: del.length };
}
