/**
 * SISTEM MONITORING TRAINING & DATA PESERTA
 * Basis visual: Index.html lama dari pengguna, dengan arsitektur backend
 * disesuaikan ke sheet:
 * 1) Data
 * 2) DataPeserta
 * 3) NamaPerusahaan
 * 4) UserPemakai
 * 5) UserPeserta
 *
 * Catatan:
 * - Login Pemakai: Username + Password dari UserPemakai.
 * - Login Peserta: NIK + Password dari UserPeserta.
 * - Password mengikuti isi sheet apa adanya sesuai kebutuhan saat ini.
 * - Operasi database divalidasi lagi di server.
 */

const APP = {
  TITLE: 'Training Management System',
  TIMEZONE: 'Asia/Makassar',
  SHEETS: {
    DATA: 'Data',
    PESERTA: 'DataPeserta',
    PERUSAHAAN: 'NamaPerusahaan',
    PEMAKAI: 'UserPemakai',
    USER_PESERTA: 'UserPeserta',
    USER_AKTIVITAS: 'UserAktivitas',
    TUGAS: 'Tugas',
    PENGUMPULAN_TUGAS: 'PengumpulanTugas',
    SETTING_TUGAS: 'SettingTugas',
    SETTING_NILAI: 'SettingNilai',
    KALENDER_AGENDA: 'KalenderAgenda',
    DOKUMEN_REPORT: 'DokumenReport',
    SETTINGS: 'LogoAplikasi',
    DB_TRAINING: 'DbTraining',
    SCOPES: 'Scopes',
    SOAL_TEST: 'SoalTest',
    SOAL_TEST_RESPONSES: 'SoalTestResponses',
    SOAL_TEST_ANSWERS: 'SoalTestAnswers'
  },
  DB_KEYS: {
    Data: 'Data',
    DataPeserta: 'DataPeserta',
    NamaPerusahaan: 'NamaPerusahaan',
    UserPemakai: 'UserPemakai',
    UserPeserta: 'UserPeserta',
    DbTraining: 'DbTraining',
    Tugas: 'Tugas'
  }
};

/** ========================= SCOPES / HAK AKSES MENU ========================= */
function _isProgramerRole_(role) {
  return _norm_(role) === 'programer';
}

function _scopeHeaders_() {
  return ['Username','Scope Key','Menu','Parent','Type','Allowed','Programer Only','Updated At'];
}

function _ensureScopesSheet_() {
  let sh = _sheet_(APP.SHEETS.SCOPES);
  const headers = _scopeHeaders_();
  if (!sh) {
    sh = _ss_().insertSheet(APP.SHEETS.SCOPES);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    return sh;
  }
  if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(headers.length,sh.getLastColumn())).getDisplayValues()[0].map(String);
    if (current.slice(0,headers.length).some((v,i)=>_norm_(v)!==_norm_(headers[i]))) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }
  return sh;
}

function _scopeCatalogNormalize_(catalog) {
  const arr = Array.isArray(catalog) ? catalog : [];
  const seen = new Set();
  const out = [];
  arr.forEach(item => {
    if (!item) return;
    const key = String(item.key || '').trim();
    const menu = String(item.menu || '').trim();
    if (!key || !menu || seen.has(key)) return;
    seen.add(key);
    out.push({
      key,
      menu,
      parentKey: String(item.parentKey || '').trim(),
      parent: String(item.parent || '').trim(),
      type: String(item.type || 'submenu').trim().toLowerCase() === 'menu' ? 'menu' : 'submenu',
      programerOnly: Boolean(item.programerOnly)
    });
  });
  return out;
}

function _getPemakaiRowsForScopes_() {
  const result = _records_(APP.SHEETS.PEMAKAI);
  if (!result.success) throw new Error(result.message);
  const h = result.headers;
  const iUsername = _headerIndex_(h,['Username','User Name']);
  const iNama = _headerIndex_(h,['Nama Lengkap','Nama']);
  const iRole = _headerIndex_(h,['Role','Peran']);
  if (iUsername < 0) throw new Error('Kolom Username pada UserPemakai tidak ditemukan.');
  return result.rows.map(r => ({
    username:String(r.values[iUsername] || '').trim(),
    name:String(iNama>=0 ? r.values[iNama] || '' : r.values[iUsername] || '').trim(),
    role:String(iRole>=0 ? r.values[iRole] || '' : '').trim()
  })).filter(x=>x.username);
}

function _syncScopesCatalogForUsers_(catalog, targetUsername) {
  const normalized = _scopeCatalogNormalize_(catalog);
  const sh = _ensureScopesSheet_();
  const users = _getPemakaiRowsForScopes_();
  const target = targetUsername ? _norm_(targetUsername) : '';
  const eligibleUsers = target ? users.filter(u=>_norm_(u.username)===target) : users;
  if (!eligibleUsers.length || !normalized.length) return;

  const lastRow = sh.getLastRow();
  const existing = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,8).getValues() : [];
  const map = new Map();
  existing.forEach((r,i)=>{
    const k = _norm_(r[0]) + '|' + String(r[1] || '').trim();
    if (r[0] && r[1]) map.set(k,{row:i+2,values:r});
  });

  const toAppend = [];
  const now = new Date();
  eligibleUsers.forEach(user => {
    normalized.forEach(item => {
      const k = _norm_(user.username) + '|' + item.key;
      const found = map.get(k);
      if (!found) {
        // Default ON untuk menjaga akses menu yang sebelumnya sudah tersedia.
        // Menu khusus Programer tetap OFF untuk user biasa.
        toAppend.push([
          user.username, item.key, item.menu, item.parent, item.type,
          item.programerOnly ? false : true, item.programerOnly ? true : false, now
        ]);
      } else {
        // Metadata menu mengikuti versi UI terbaru, status Allowed tidak diubah.
        sh.getRange(found.row,3,1,5).setValues([[item.menu,item.parent,item.type,found.values[5],item.programerOnly]]);
      }
    });
  });
  if (toAppend.length) sh.getRange(sh.getLastRow()+1,1,toAppend.length,8).setValues(toAppend);
  SpreadsheetApp.flush();
}

function _isProgramerUsername_(username) {
  const user = _authorizePemakai_(username);
  const h = _records_(APP.SHEETS.PEMAKAI).headers;
  const iRole = _headerIndex_(h,['Role','Peran']);
  return _isProgramerRole_(iRole>=0 ? user.values[iRole] : '');
}

/**
 * Validasi Scope untuk operasi database di server.
 * Programer tetap memiliki akses penuh. Untuk user biasa, jika baris Scope
 * sudah tersedia maka status Allowed menjadi sumber otorisasi utama.
 * Jika belum ada baris Scope sama sekali, akses lama tetap dipertahankan.
 */
function _assertDatabaseScope_(sheetName, username) {
  const cleanSheet = String(sheetName || '').trim();
  if (!cleanSheet || !APP.DB_KEYS[cleanSheet]) {
    throw new Error('Sheet database tidak diizinkan.');
  }

  const user = _authorizePemakai_(username);
  const h = _records_(APP.SHEETS.PEMAKAI).headers;
  const iRole = _headerIndex_(h,['Role','Peran']);
  const iUsername = _headerIndex_(h,['Username','User Name']);
  const role = iRole >= 0 ? String(user.values[iRole] || '') : '';
  if (_isProgramerRole_(role)) return true;

  // Database tertentu memang khusus Programer.
  const programerOnlySheets = ['DataPeserta','UserPemakai','UserPeserta'];
  if (programerOnlySheets.includes(cleanSheet)) {
    throw new Error('Akses database ini hanya diperbolehkan untuk user dengan role Programer.');
  }

  const canonicalUsername = iUsername >= 0
    ? String(user.values[iUsername] || username).trim()
    : String(username || '').trim();
  const sh = _ensureScopesSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return true;

  const rows = sh.getRange(2,1,lastRow-1,8).getValues();
  const targetKey = 'db:' + cleanSheet;
  let found = false;
  let allowed = false;
  rows.forEach(r => {
    if (_norm_(r[0]) !== _norm_(canonicalUsername)) return;
    if (String(r[1] || '').trim() !== targetKey) return;
    found = true;
    allowed = Boolean(r[5]) && !Boolean(r[6]);
  });

  // Bila katalog Scope belum tersinkron untuk user ini, pertahankan kompatibilitas.
  if (!found) return true;
  if (!allowed) throw new Error('Database ini tidak diberikan kepada user Anda.');
  return true;
}

function getUserScopeAccess(username, catalog, sessionToken) {
  try {
    // Scope memakai username kanonis dari UserPemakai sebagai identitas utama.
    // sessionToken tetap diterima untuk kompatibilitas, tetapi tidak boleh membuat
    // Scope gagal hanya karena token sesi browser lama/tidak sinkron.
    const cleanUsername = String(username || '').trim();
    if (!cleanUsername) {
      return {success:false,sessionExpired:true,all:false,allowedKeys:[],message:'Username Pemakai tidak tersedia. Silakan login ulang.'};
    }

    const user = _authorizePemakai_(cleanUsername);
    const h = _records_(APP.SHEETS.PEMAKAI).headers;
    const iUsername = _headerIndex_(h,['Username','User Name']);
    const iRole = _headerIndex_(h,['Role','Peran']);
    const canonicalUsername = iUsername>=0
      ? String(user.values[iUsername] || cleanUsername).trim()
      : cleanUsername;
    const role = iRole>=0 ? String(user.values[iRole] || '') : '';

    if (_isProgramerRole_(role)) {
      return {success:true,all:true,allowedKeys:[],role,username:canonicalUsername};
    }

    const normalized = _scopeCatalogNormalize_(catalog);
    const scopeCacheKey = 'TMS_SCOPE_ACCESS_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, _norm_(canonicalUsername), Utilities.Charset.UTF_8)).replace(/=+$/,'');
    try {
      const cached = CacheService.getScriptCache().get(scopeCacheKey);
      if (cached) {
        const parsed=JSON.parse(cached);
        if (parsed && Array.isArray(parsed.allowedKeys)) return {success:true,all:false,allowedKeys:parsed.allowedKeys,role,username:canonicalUsername};
      }
    } catch(e) {}

    const sh = _sheet_(APP.SHEETS.SCOPES);
    const allowed = [];
    if (sh && sh.getLastRow() >= 2) {
      const values = sh.getRange(2,1,sh.getLastRow()-1,8).getValues();
      values.forEach(r=>{
        if (_norm_(r[0]) !== _norm_(canonicalUsername)) return;
        if (Boolean(r[5]) && !Boolean(r[6])) allowed.push(String(r[1] || '').trim());
      });
    } else if (!sh) {
      // Kompatibilitas akun lama: bila Sheet Scopes belum ada, default-kan semua
      // menu non-Programer tanpa menulis Sheet pada jalur login.
      normalized.forEach(item=>{ if(!item.programerOnly) allowed.push(item.key); });
    }

    // Bila belum ada baris Scope untuk user, perilaku default lama tetap dipertahankan
    // (semua menu non-Programer boleh digunakan), tetapi tidak ada write pada login.
    const hasUserScope = sh && sh.getLastRow() >= 2 && sh.getRange(2,1,sh.getLastRow()-1,8).getValues().some(r=>_norm_(r[0])===_norm_(canonicalUsername));
    if (!hasUserScope && normalized.length) {
      normalized.forEach(item=>{ if(!item.programerOnly && !allowed.includes(item.key)) allowed.push(item.key); });
    }
    const scopePayload=JSON.stringify({allowedKeys:allowed});
    try { if(scopePayload.length<90000) CacheService.getScriptCache().put(scopeCacheKey,scopePayload,120); } catch(e) {}
    return {success:true,all:false,allowedKeys:allowed,role,username:canonicalUsername};
  } catch (err) {
    return {success:false,sessionExpired:false,all:false,allowedKeys:[],message:err.message || 'Hak akses Scope tidak dapat dimuat.'};
  }
}

function getScopesAdminData(adminUsername, catalog) {
  const admin = _authorizePemakai_(adminUsername);
  const h = _records_(APP.SHEETS.PEMAKAI).headers;
  const iRole = _headerIndex_(h,['Role','Peran']);
  if (!_isProgramerRole_(iRole>=0 ? admin.values[iRole] : '')) throw new Error('Menu Scopes hanya dapat dikelola oleh user dengan role Programer.');

  const normalized = _scopeCatalogNormalize_(catalog);
  if (normalized.length) _syncScopesCatalogForUsers_(normalized);
  const users = _getPemakaiRowsForScopes_().map(u=>({username:u.username,name:u.name,role:u.role}));
  const sh = _ensureScopesSheet_();
  const rows = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,8).getValues() : [];
  const entries = rows.map(r=>({username:String(r[0]||''),key:String(r[1]||''),menu:String(r[2]||''),parent:String(r[3]||''),type:String(r[4]||''),allowed:Boolean(r[5]),programerOnly:Boolean(r[6])}));
  return {success:true,users,items:normalized,entries};
}

function saveUserScopes(adminUsername, targetUsername, catalog, allowedKeys) {
  const admin = _authorizePemakai_(adminUsername);
  const h = _records_(APP.SHEETS.PEMAKAI).headers;
  const iRole = _headerIndex_(h,['Role','Peran']);
  if (!_isProgramerRole_(iRole>=0 ? admin.values[iRole] : '')) throw new Error('Hanya user dengan role Programer yang dapat menyimpan Scopes.');

  const target = String(targetUsername || '').trim();
  if (!target) throw new Error('User Pemakai belum dipilih.');
  const users = _getPemakaiRowsForScopes_();
  if (!users.some(u=>_norm_(u.username)===_norm_(target))) throw new Error('User Pemakai tidak ditemukan.');

  const normalized = _scopeCatalogNormalize_(catalog);
  _syncScopesCatalogForUsers_(normalized, target);
  const allowed = new Set((Array.isArray(allowedKeys)?allowedKeys:[]).map(v=>String(v||'').trim()));
  const sh = _ensureScopesSheet_();
  const lastRow = sh.getLastRow();
  const rows = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,8).getValues() : [];
  const now = new Date();
  rows.forEach((r,i)=>{
    if (_norm_(r[0]) !== _norm_(target)) return;
    const key = String(r[1] || '').trim();
    const programerOnly = Boolean(r[6]);
    const canAllow = !programerOnly;
    sh.getRange(i+2,6,1,2).setValues([[canAllow && allowed.has(key),programerOnly]]);
    sh.getRange(i+2,8).setValue(now);
  });
  SpreadsheetApp.flush();
  return {success:true,message:'Scopes berhasil disimpan.'};
}


function _handleSoalExamCheatRequest_(e) {
  try {
    const p=(e && e.parameter)||{};
    if(String(p.action||'').trim().toLowerCase()!=='soal_exam_cheat') return null;
    const nik=String(p.nik||'').trim();
    const attemptId=String(p.attemptId||'').trim();
    const reason=String(p.reason||'Melakukan kecurangan pada saat ujian.').trim();
    let answers={};
    try{answers=p.answers?JSON.parse(p.answers):{};}catch(err){answers={};}
    if(nik && attemptId) markSoalTestCheating(nik,attemptId,answers,reason);
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
  }
}

function _handleAutoLogoutRequest_(e) {
  try {
    const p = (e && e.parameter) || {};
    if (String(p.action || '').trim().toLowerCase() !== 'auto_logout') return null;

    const username = String(p.username || '').trim();
    const userType = String(p.userType || '').trim();
    const sessionToken = String(p.sessionToken || '').trim();
    if (username && sessionToken) {
      // Jangan langsung menulis Logout At. Saat refresh, pagehide/beforeunload
      // juga terpanggil. Logout dijadwalkan dulu dan akan dibatalkan oleh
      // heartbeat halaman yang baru jika ini hanya refresh.
      _schedulePendingAutoLogout_(username, userType, sessionToken);
    }
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput('ERROR').setMimeType(ContentService.MimeType.TEXT);
  }
}

function _pendingAutoLogoutKeyFromActivityKey_(activityKey) {
  return activityKey ? 'PENDING_AUTO_LOGOUT_' + String(activityKey) : '';
}

function _pendingAutoLogoutKey_(sessionToken) {
  return _pendingAutoLogoutKeyFromActivityKey_(_activitySessionKey_(sessionToken));
}

function _schedulePendingAutoLogout_(username, userType, sessionToken) {
  const key = _pendingAutoLogoutKey_(sessionToken);
  if (!key || key.length < 20) return;
  CacheService.getScriptCache().put(key, JSON.stringify({
    username:String(username || ''),
    userType:String(userType || ''),
    sessionToken:String(sessionToken || ''),
    createdAt:Date.now()
  }), 120);
}

function _cancelPendingAutoLogout_(sessionToken) {
  const key = _pendingAutoLogoutKey_(sessionToken);
  if (key && key.length >= 20) CacheService.getScriptCache().remove(key);
}

function _finalizePendingAutoLogoutByActivityKey_(activityKey) {
  const key = _pendingAutoLogoutKeyFromActivityKey_(activityKey);
  if (!key || key.length < 20) return;
  const cache = CacheService.getScriptCache();
  const raw = cache.get(key);
  if (!raw) return;
  try {
    const p = JSON.parse(raw);
    // Beri grace period supaya refresh sempat mengirim heartbeat halaman baru.
    if (!p.createdAt || Date.now() - Number(p.createdAt) < 2500) return;
    if (p.sessionToken) recordUserActivity(p.username, p.userType, p.sessionToken, 'logout');
  } catch (e) {}
  cache.remove(key);
}

function doGet(e) {
  const autoLogout = _handleAutoLogoutRequest_(e);
  if (autoLogout) return autoLogout;

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * sendBeacon() menggunakan HTTP POST, sehingga logout saat tab/browser ditutup
 * harus ditangani oleh doPost(), bukan hanya doGet().
 */
function doPost(e) {
  const cheat = _handleSoalExamCheatRequest_(e);
  if (cheat) return cheat;
  const autoLogout = _handleAutoLogoutRequest_(e);
  if (autoLogout) return autoLogout;
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

/** URL Web App untuk kebutuhan logout saat pagehide/beforeunload. */
function getAppWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl() || '';
  } catch (e) {
    return '';
  }
}

function _ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _sheet_(name) {
  return _ss_().getSheetByName(name);
}

function _norm_(value) {
  return String(value == null ? '' : value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

// Normalisasi NIK khusus agar perbedaan tipe angka/text atau leading zero
// (mis. 001234 vs 1234) tidak memblokir peserta yang sama.
function _normNik_(value) {
  const raw=String(value == null ? '' : value).trim().replace(/\s+/g,'');
  if(!raw) return '';
  if(/^\d+$/.test(raw)) return raw.replace(/^0+(?=\d)/,'');
  return raw.toLowerCase();
}

function _headerIndex_(headers, aliases) {
  const normalized = headers.map(_norm_);
  for (const alias of aliases) {
    const idx = normalized.indexOf(_norm_(alias));
    if (idx >= 0) return idx;
  }
  return -1;
}

function _records_(sheetName) {
  const sh = _sheet_(sheetName);
  if (!sh) return { success: false, message: `Sheet "${sheetName}" tidak ditemukan.` };

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { success: true, headers: [], rows: [] };
  }

  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(v => String(v || '').trim());
  const rows = values.slice(1).map((row, i) => ({
    rowNumber: i + 2,
    values: row
  })).filter(item => item.values.some(v => String(v || '').trim() !== ''));

  return { success: true, headers, rows };
}

/* ===== LOGIN CONCURRENCY / FAST AUTH INDEX =====
 * Login hanya perlu membaca data akun. Jangan melakukan sinkronisasi/write ke
 * Google Sheets pada setiap login karena 100+ browser dapat datang bersamaan.
 * CacheService dipakai sebagai shared read cache. Cache dibangun satu kali
 * ketika cold-start dengan ScriptLock, lalu seluruh login membaca cache.
 */
const LOGIN_AUTH_CACHE_TTL_SEC_ = 300;
const LOGIN_AUTH_CACHE_MAX_BYTES_ = 90000;
const LOGIN_AUTH_CACHE_WAIT_MS_ = 7000;

function _loginAuthCacheKey_(type) {
  return 'TMS_LOGIN_AUTH_V4_' + (_norm_(type) === 'peserta' ? 'PESERTA' : 'PEMAKAI');
}

function _readLoginAuthSheet_(type) {
  const sheetName = _norm_(type) === 'peserta' ? APP.SHEETS.USER_PESERTA : APP.SHEETS.PEMAKAI;
  const sh = _sheet_(sheetName);
  if (!sh) return {success:false, message:'Sheet "' + sheetName + '" tidak ditemukan.'};

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return {success:true, headers:[], rows:[]};

  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v => String(v || '').trim());
  const values = lastRow >= 2 ? sh.getRange(2,1,lastRow-1,lastCol).getDisplayValues() : [];
  const idx = aliases => _headerIndex_(headers, aliases);

  if (_norm_(type) === 'peserta') {
    const iNik = idx(['NIK']);
    const iNama = idx(['Nama Peserta','Nama']);
    const iPerusahaan = idx(['Nama Perusahaan','Perusahaan']);
    const iJabatan = idx(['Jabatan']);
    const iPassword = idx(['Password']);
    const iStatus = idx(['Status']);
    const iFoto = idx(['Foto Profil','Foto Profile','Foto']);
    if (iNik < 0 || iPassword < 0) return {success:false,message:'Kolom NIK/Password pada UserPeserta tidak ditemukan.'};
    const rows = values.map(r => ({
      nik:String(r[iNik] || '').trim(),
      password:String(r[iPassword] || ''),
      nama:String(iNama >= 0 ? r[iNama] || '' : ''),
      perusahaan:String(iPerusahaan >= 0 ? r[iPerusahaan] || '' : ''),
      jabatan:String(iJabatan >= 0 ? r[iJabatan] || '' : ''),
      status:String(iStatus >= 0 ? r[iStatus] || '' : 'Aktif'),
      foto:(iFoto >= 0 ? String(r[iFoto] || '') : '').length <= 2048 ? String(iFoto >= 0 ? r[iFoto] || '' : '') : ''
    })).filter(x => x.nik);
    return {success:true, rows};
  }

  const iUsername = idx(['Username','User Name']);
  const iPassword = idx(['Password']);
  if (iUsername < 0 || iPassword < 0) return {success:false,message:'Kolom Username/Password pada UserPemakai tidak ditemukan.'};
  const iNama = idx(['Nama Lengkap','Nama']);
  const iJabatan = idx(['Jabatan']);
  const iUnit = idx(['Unit Usaha','Unit']);
  const iEmail = idx(['Email']);
  const iHp = idx(['No HP','Nomor HP','No. HP']);
  const iGender = idx(['Jensi Kelamin','Jenis Kelamin','Gender']);
  const iRole = idx(['Role','Peran']);
  const iStatus = idx(['Status']);
  const iFoto = idx(['Foto Profi"l','Foto Profil','Foto Profile','Foto']);
  const rows = values.map(r => ({
    username:String(r[iUsername] || '').trim(),
    password:String(r[iPassword] || ''),
    nama:String(iNama >= 0 ? r[iNama] || '' : ''),
    jabatan:String(iJabatan >= 0 ? r[iJabatan] || '' : ''),
    unitUsaha:String(iUnit >= 0 ? r[iUnit] || '' : ''),
    email:String(iEmail >= 0 ? r[iEmail] || '' : ''),
    noHp:String(iHp >= 0 ? r[iHp] || '' : ''),
    jenisKelamin:String(iGender >= 0 ? r[iGender] || '' : ''),
    role:String(iRole >= 0 ? r[iRole] || 'Pemakai' : 'Pemakai'),
    status:String(iStatus >= 0 ? r[iStatus] || '' : 'Aktif'),
    foto:(iFoto >= 0 ? String(r[iFoto] || '') : '').length <= 2048 ? String(iFoto >= 0 ? r[iFoto] || '' : '') : ''
  })).filter(x => x.username);
  return {success:true, rows};
}

function _getLoginAuthRows_(type) {
  const key = _loginAuthCacheKey_(type);
  const cache = CacheService.getScriptCache();

  try {
    const raw = cache.get(key);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.rows)) return parsed.rows;
    }
  } catch (e) {}

  // Cold-cache stampede protection: hanya satu execution membaca Sheet;
  // execution lain menunggu sebentar lalu membaca cache yang baru dibuat.
  const lock = LockService.getScriptLock();
  if (lock.tryLock(150)) {
    try {
      try {
        const again = cache.get(key);
        if (again) {
          const parsedAgain = JSON.parse(again);
          if (parsedAgain && Array.isArray(parsedAgain.rows)) return parsedAgain.rows;
        }
      } catch (e) {}

      const result = _readLoginAuthSheet_(type);
      if (!result.success) throw new Error(result.message);
      const lightRows=result.rows.map(r=>{const copy={...r};delete copy.foto;return copy;});
      const payload = JSON.stringify({version:5, rows:lightRows});
      if (payload.length <= LOGIN_AUTH_CACHE_MAX_BYTES_) {
        try { cache.put(key, payload, LOGIN_AUTH_CACHE_TTL_SEC_); } catch (e) {}
      }
      // Foto disimpan sebagai cache per akun supaya foto tidak membuat payload
      // autentikasi bersama melewati batas 100 KB CacheService.
      try {
        const photoEntries={};
        result.rows.forEach(r=>{const identity=_norm_(type)==='peserta'?r.nik:r.username;const photo=String(r.foto||'');if(identity&&photo)photoEntries['TMS_LOGIN_PHOTO_'+_norm_(type)+'_'+_norm_(identity)]=photo;});
        Object.keys(photoEntries).slice(0,900).forEach(k=>{try{cache.put(k,photoEntries[k],LOGIN_AUTH_CACHE_TTL_SEC_);}catch(e){}});
      } catch (e) {}
      return lightRows;
    } finally {
      try { lock.releaseLock(); } catch (e) {}
    }
  }

  // Jangan langsung membaca Sheet saat lock sedang dipegang execution lain.
  // Poll cache sebentar agar 100+ login cold-start tidak berubah menjadi 100+ pembacaan Sheet.
  const polled=_cachePollJson_(key,LOGIN_AUTH_CACHE_WAIT_MS_);
  if(polled&&Array.isArray(polled.rows))return polled.rows;
  const fallback = _readLoginAuthSheet_(type);
  if (!fallback.success) throw new Error(fallback.message);
  return fallback.rows.map(r=>{const copy={...r};delete copy.foto;return copy;});
}

function _invalidateLoginAuthCache_(type) {
  try {
    const c = CacheService.getScriptCache();
    if (!type) c.removeAll([_loginAuthCacheKey_('pemakai'), _loginAuthCacheKey_('peserta')]);
    else c.remove(_loginAuthCacheKey_(type));
  } catch (e) {}
}

function _findCachedLoginRow_(type, identity) {
  const rows = _getLoginAuthRows_(type);
  const wanted = _norm_(identity);
  if (!wanted) return null;
  if (_norm_(type) === 'peserta') return rows.find(r => _norm_(r.nik) === wanted) || null;
  return rows.find(r => _norm_(r.username) === wanted) || null;
}
function _getCachedLoginPhoto_(type,identity){
  try{return CacheService.getScriptCache().get('TMS_LOGIN_PHOTO_'+_norm_(type)+'_'+_norm_(identity))||'';}catch(e){return '';}
}

function _findUserPemakai_(username, password) {
  try {
    const user = _findCachedLoginRow_('pemakai', username);
    if (!user || String(user.password || '') !== String(password == null ? '' : password)) {
      return { success:false, message:'Username atau Password Pemakai salah.' };
    }
    if (/^non\s*aktif$/i.test(String(user.status || '').trim())) {
      return {
        success:false,
        code:'ACCOUNT_INACTIVE',
        message:'Akun Pemakai berstatus Non Aktif. Silakan minta administrator melakukan pengaktifan kembali agar dapat login.'
      };
    }
    return {
      success:true,
      userType:'pemakai',
      username:user.username,
      nama:user.nama || user.username,
      jabatan:user.jabatan || '',
      unitUsaha:user.unitUsaha || '',
      email:user.email || '',
      noHp:user.noHp || '',
      jenisKelamin:user.jenisKelamin || '',
      role:user.role || 'Pemakai',
      foto:'' // Foto dimuat lewat endpoint ringan setelah dashboard tampil; mengurangi payload login.
    };
  } catch (err) {
    return {success:false,message:'Username atau Password Pemakai salah.'};
  }
}

function _findUserPeserta_(nik, password) {
  try {
    const user = _findCachedLoginRow_('peserta', nik);
    if (!user || String(user.password || '') !== String(password == null ? '' : password)) {
      return { success:false, message:'NIK atau Password Peserta salah.' };
    }
    if (/^non\s*aktif$/i.test(String(user.status || '').trim())) {
      return { success:false, message:'Akun peserta berstatus Non Aktif dan tidak dapat login. Silakan hubungi administrator.' };
    }

    // UserPeserta sudah menyimpan Nama/Perusahaan/Jabatan. Fallback ke Data hanya
    // jika ada data profil yang memang kosong; jalur login normal tidak menyentuh Data.
    let jabatan=String(user.jabatan || '').trim();
    let nama=String(user.nama || '').trim();
    let perusahaan=String(user.perusahaan || '').trim();
    if (!jabatan || !nama || !perusahaan) {
      try {
        const dataRes = _records_(APP.SHEETS.DATA);
        if (dataRes.success) {
          const dh=dataRes.headers;
          const diNik=_headerIndex_(dh,['NIK']);
          const diNama=_headerIndex_(dh,['Nama Peserta','Nama']);
          const diPerusahaan=_headerIndex_(dh,['Nama Perusahaan','Perusahaan']);
          const diJabatan=_headerIndex_(dh,['Jabatan']);
          const n=_norm_(nik);
          const drow=(dataRes.rows||[]).find(x=>diNik>=0 && _norm_(x.values[diNik])===n);
          if(drow){
            if(!nama && diNama>=0) nama=String(drow.values[diNama]||'').trim();
            if(!perusahaan && diPerusahaan>=0) perusahaan=String(drow.values[diPerusahaan]||'').trim();
            if(!jabatan && diJabatan>=0) jabatan=String(drow.values[diJabatan]||'').trim();
          }
        }
      } catch(e) {}
    }
    return {
      success:true,
      userType:'peserta',
      nik:user.nik,
      username:user.nik,
      nama:nama || user.nik,
      jabatan,
      perusahaan,
      status:user.status || 'Aktif',
      role:'Peserta',
      foto:'' // Foto dimuat lewat endpoint ringan setelah dashboard tampil; mengurangi payload login.
    };
  } catch (err) {
    return {success:false,message:'NIK atau Password Peserta salah.'};
  }
}


function _sessionSecret_() {
  const props=PropertiesService.getScriptProperties();
  let secret=props.getProperty('TMS_SESSION_SECRET');
  if(secret) return secret;

  // Hanya serialisasi saat secret belum pernah dibuat. Tanpa ini, 100+ login
  // pertama yang benar-benar bersamaan dapat membuat beberapa execution
  // menandatangani token dengan secret berbeda lalu token lama menjadi invalid.
  const lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    secret=props.getProperty('TMS_SESSION_SECRET');
    if(!secret){
      secret=Utilities.getUuid()+'-'+Utilities.getUuid();
      props.setProperty('TMS_SESSION_SECRET',secret);
    }
    return secret;
  } finally {
    try{lock.releaseLock();}catch(e){}
  }
}

function _createSessionToken_(userType, identity) {
  const payload=JSON.stringify({userType:String(userType||''),identity:String(identity||''),iat:Date.now()});
  const encoded=Utilities.base64EncodeWebSafe(payload, Utilities.Charset.UTF_8);
  const sig=Utilities.computeHmacSha256Signature(encoded,_sessionSecret_());
  const signature=Utilities.base64EncodeWebSafe(sig);
  return encoded+'.'+signature;
}

function _readSessionPayload_(token) {
  try {
    const parts=String(token||'').split('.');
    if(parts.length!==2)return null;
    const encoded=parts[0], supplied=parts[1];
    const expected=Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(encoded,_sessionSecret_()));
    if(expected!==supplied)return null;
    const payload=JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(encoded)).getDataAsString('UTF-8'));
    if(!payload || !payload.userType || !payload.identity || !payload.iat)return null;
    return payload;
  }catch(err){return null;}
}

function _verifySessionToken_(token,userType,identity,maxAgeMs) {
  try{
    const payload=_readSessionPayload_(token);
    if(!payload)return false;
    const age=Date.now()-Number(payload.iat||0);
    return payload.userType===String(userType||'') && _norm_(payload.identity)===_norm_(identity) && age>=0 && age<=Number(maxAgeMs||43200000);
  }catch(err){return false;}
}


/**
 * Membuat sheet log aktivitas user jika belum tersedia.
 * Satu baris merepresentasikan satu sesi login; heartbeat hanya memperbarui baris tersebut.
 */
function _ensureUserActivitySheet_() {
  let sh = _sheet_(APP.SHEETS.USER_AKTIVITAS);
  const headers = ['Session Key','Jenis User','Username / NIK','Nama','Role','Unit / Perusahaan','Login At','Last Active At','Logout At','Status'];
  if (!sh) {
    sh = _ss_().insertSheet(APP.SHEETS.USER_AKTIVITAS);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _activityUserInfo_(username, userType) {
  const type = _norm_(userType) === 'peserta' ? 'peserta' : 'pemakai';
  const user=_findCachedLoginRow_(type,username);
  if(!user)throw new Error('Sesi user tidak ditemukan.');
  if(type==='pemakai'){
    return {userType:'Pemakai',username:String(user.username||username),nik:'',name:String(user.nama||user.username||username),role:String(user.role||'Pemakai'),unit:String(user.unitUsaha||'')};
  }
  return {userType:'Peserta',username:String(user.nik||username),nik:String(user.nik||username),name:String(user.nama||username),role:'Peserta',unit:String(user.perusahaan||'')};
}

function _activitySessionKey_(sessionToken) {
  const raw = String(sessionToken || '').trim();
  if (!raw) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(digest);
}

/**
 * Mencatat login, heartbeat, dan logout user.
 * Heartbeat 30 detik dipakai frontend agar status "Aktif" dapat dipantau real-time.
 */
function recordUserActivity(username, userType, sessionToken, eventType) {
  try {
    const type = _norm_(userType) === 'peserta' ? 'peserta' : 'pemakai';
    const identity = type === 'peserta' ? String(username || '') : String(username || '');
    const maxAge = 43200000; // 12 jam mengikuti umur sesi aplikasi.
    if (!_verifySessionToken_(sessionToken, type, identity, maxAge)) {
      throw new Error('Sesi user tidak valid atau sudah kedaluwarsa.');
    }
    // Heartbeat/login dari halaman baru membatalkan pending logout.
    if (_norm_(eventType || 'heartbeat') !== 'logout') {
      _cancelPendingAutoLogout_(sessionToken);
    }
    const key = _activitySessionKey_(sessionToken);
    if (!key) throw new Error('Session key tidak tersedia.');

    const info = _activityUserInfo_(identity, type);
    const sh = _ensureUserActivitySheet_();
    const now = new Date();
    const ev = _norm_(eventType || 'heartbeat');
    const lastRow = sh.getLastRow();
    let rowNumber = -1;
    if (lastRow >= 2) {
      const finder = sh.getRange(2,1,lastRow-1,1).createTextFinder(key).matchEntireCell(true);
      const found = finder.findNext();
      if (found) rowNumber = found.getRow();
    }

    if (ev === 'login') {
      const row = [key, info.userType, info.username, info.name, info.role, info.unit, now, now, '', 'Aktif'];
      if (rowNumber > 1) sh.getRange(rowNumber,1,1,row.length).setValues([row]);
      else sh.getRange(sh.getLastRow()+1,1,1,row.length).setValues([row]);
      return {success:true,status:'Aktif'};
    }

    if (rowNumber < 2) {
      // Recovery: browser mungkin masih menyimpan sesi lama tanpa baris aktivitas.
      const row = [key, info.userType, info.username, info.name, info.role, info.unit, now, now, ev === 'logout' ? now : '', ev === 'logout' ? 'Tidak Aktif' : 'Aktif'];
      sh.getRange(sh.getLastRow()+1,1,1,row.length).setValues([row]);
      return {success:true,status:ev === 'logout' ? 'Tidak Aktif' : 'Aktif'};
    }

    sh.getRange(rowNumber,8).setValue(now);
    if (ev === 'logout') {
      sh.getRange(rowNumber,9).setValue(now);
      sh.getRange(rowNumber,10).setValue('Tidak Aktif');
    } else {
      sh.getRange(rowNumber,10).setValue('Aktif');
    }
    return {success:true,status:ev === 'logout' ? 'Tidak Aktif' : 'Aktif'};
  } catch (err) {
    return {success:false,message:'Gagal mencatat aktivitas user: ' + err.message};
  }
}

/**
 * Mengambil kondisi user terbaru untuk dashboard Pemakai.
 * Aktif = heartbeat terakhir <= 90 detik dan sesi belum logout.
 */

/**
 * ===== AUDIT LOG AKTIVITAS USER =====
 * Mencatat seluruh aktivitas yang dikirim aplikasi setelah user berhasil login.
 * Satu baris = satu aktivitas. Detail sensitif seperti password/token tidak pernah
 * disimpan; session token hanya dipakai untuk validasi dan disimpan dalam bentuk hash.
 */
function _ensureUserActivityLogSheet_() {
  const ss = _ss_();
  const name = 'LogAktivitasUser';
  let sh = ss.getSheetByName(name);
  const headers = ['Timestamp','Jenis User','Username / NIK','Nama Lengkap','Role','Unit / Perusahaan','Aktivitas','Detail','Status','Session Key'];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    try { sh.getRange(1,1,1,headers.length).setFontWeight('bold'); } catch (e) {}
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn())).getDisplayValues()[0];
    headers.forEach((h,i)=>{
      if (_norm_(current[i]) !== _norm_(h)) sh.getRange(1,i+1).setValue(h);
    });
  }
  return sh;
}

function _sanitizeUserActivityDetail_(value) {
  return String(value ?? '')
    .replace(/[\r\n\t]+/g,' ')
    .replace(/\s+/g,' ')
    .trim()
    .slice(0,500);
}

function _userActivityActionLabel_(fn) {
  const map = {
    getDatabasePage:'Melihat Data Base',
    getDatabaseExportData:'Export Data Base XLSX',
    getDatabaseTable:'Melihat Data Base',
    saveDatabaseRow:'Mengubah Data',
    deleteDatabaseRow:'Menghapus Data',
    importDatabaseRows:'Import Data',
    uploadTaskMaterial:'Upload Materi Tugas',
    deleteTaskMaterialFile:'Menghapus File Materi',
    deleteTaskMaterialFolder:'Menghapus Folder Materi',
    getTaskMaterialFiles:'Melihat Materi Tugas',
    uploadReportDocument:'Upload Dokumen Report',
    renameReportDocument:'Mengubah Nama Dokumen Report',
    deleteReportDocument:'Menghapus Dokumen Report',
    getReportDocuments:'Melihat Rekap Dokumen Report',
    saveTaskSetting:'Menyimpan Setting Tugas',
    deleteTaskSetting:'Menghapus Setting Tugas',
    saveSettingNilai:'Menyimpan Setting Nilai',
    postSettingNilai:'Posting Nilai',
    cancelSettingNilai:'Cancel Posting Nilai',
    updateMyProfile:'Mengubah Profil',
    updateSystemLogo:'Mengubah Logo Sistem',
    saveUserScopes:'Mengubah Scopes User',
    getScopesAdminData:'Melihat Scopes User',
    getUserScopeAccess:'Memeriksa Hak Akses',
    getTaskSubmissionMonitoring:'Melihat Pantauan Pengiriman Tugas',
    getUserActivityMonitoring:'Melihat Pantauan User Aktif',
    getCalendarAgenda:'Melihat Kalender & Agenda',
    saveCalendarAgenda:'Menyimpan Agenda Kalender',
    updateCalendarAgenda:'Mengubah Agenda Kalender',
    deleteCalendarAgenda:'Menghapus Agenda Kalender'
  };
  if (map[fn]) return map[fn];
  const raw=String(fn||'').trim();
  if (!raw) return 'Aktivitas Aplikasi';
  return raw
    .replace(/([a-z])([A-Z])/g,'$1 $2')
    .replace(/^./,c=>c.toUpperCase());
}

function recordUserAction(username,userType,sessionToken,activity,detail,status) {
  try {
    const type = _norm_(userType) === 'peserta' ? 'peserta' : 'pemakai';
    const identity = String(username || '').trim();
    if (!identity || !sessionToken) throw new Error('Identitas sesi tidak tersedia.');
    if (!_verifySessionToken_(sessionToken,type,identity,43200000)) {
      throw new Error('Sesi user tidak valid atau sudah kedaluwarsa.');
    }
    const info = _activityUserInfo_(identity,type);
    const sh = _ensureUserActivityLogSheet_();
    const now = new Date();
    const sessionKey = _activitySessionKey_(sessionToken);
    sh.appendRow([
      now,
      info.userType,
      info.username,
      info.name,
      info.role,
      info.unit,
      _sanitizeUserActivityDetail_(activity || 'Aktivitas Aplikasi'),
      _sanitizeUserActivityDetail_(detail || ''),
      _sanitizeUserActivityDetail_(status || 'Berhasil') || 'Berhasil',
      sessionKey
    ]);
    return {success:true};
  } catch (err) {
    return {success:false,message:'Gagal mencatat audit aktivitas user: '+err.message};
  }
}

function getUserActivityLog(username, filters) {
  _authorizePemakai_(username);
  try {
    const sh = _ensureUserActivityLogSheet_();
    const last = sh.getLastRow();
    if (last < 2) return {success:true,rows:[],summary:{total:0,users:0,success:0,failed:0}};
    const vals = sh.getRange(2,1,last-1,10).getValues();
    const f = filters || {};
    const q = _norm_(f.search || '');
    const type = _norm_(f.userType || '');
    const statusFilter = _norm_(f.status || '');
    const dateFrom = String(f.dateFrom || '').trim();
    const dateTo = String(f.dateTo || '').trim();
    const fmt = d => d instanceof Date && !isNaN(d.getTime()) ? Utilities.formatDate(d,APP.TIMEZONE || Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm:ss') : '';
    const iso = d => d instanceof Date && !isNaN(d.getTime()) ? Utilities.formatDate(d,APP.TIMEZONE || Session.getScriptTimeZone(),'yyyy-MM-dd') : '';
    const rows = [];
    vals.reverse().forEach(r=>{
      const stamp=r[0] instanceof Date ? r[0] : new Date(r[0]);
      const row={
        timestamp:fmt(stamp),
        dateKey:iso(stamp),
        userType:String(r[1]||''), username:String(r[2]||''), name:String(r[3]||''),
        role:String(r[4]||''), unit:String(r[5]||''), activity:String(r[6]||''),
        detail:String(r[7]||''), status:String(r[8]||'Berhasil')
      };
      if(q){
        const hay=[row.timestamp,row.userType,row.username,row.name,row.role,row.unit,row.activity,row.detail,row.status].map(_norm_).join(' | ');
        if(!hay.includes(q)) return;
      }
      if(type && _norm_(row.userType)!==type) return;
      if(statusFilter && _norm_(row.status)!==statusFilter) return;
      if(dateFrom && row.dateKey < dateFrom) return;
      if(dateTo && row.dateKey > dateTo) return;
      rows.push(row);
    });
    const userSet={};
    rows.forEach(r=>{userSet[_norm_(r.userType)+'|'+_norm_(r.username)]=true;});
    return {success:true,rows,summary:{
      total:rows.length,
      users:Object.keys(userSet).length,
      success:rows.filter(r=>_norm_(r.status)!=='gagal').length,
      failed:rows.filter(r=>_norm_(r.status)==='gagal').length
    }};
  } catch (err) {
    return {success:false,message:'Gagal memuat log aktivitas user: '+err.message};
  }
}

function getUserActivityMonitoring(username) {
  _authorizePemakai_(username);
  try {
    const users = [];
    const addUsers = (sheetName, type) => {
      const res = _records_(sheetName);
      if (!res.success) return;
      const h = res.headers;
      const iUsername = _headerIndex_(h, type === 'Pemakai' ? ['Username','User Name'] : ['NIK']);
      const iNama = _headerIndex_(h, type === 'Pemakai' ? ['Nama Lengkap','Nama'] : ['Nama Peserta','Nama']);
      const iRole = _headerIndex_(h, ['Role','Peran']);
      const iUnit = _headerIndex_(h, type === 'Pemakai' ? ['Unit Usaha','Unit'] : ['Nama Perusahaan','Perusahaan']);
      if (iUsername < 0) return;
      res.rows.forEach(item => {
        const row = item.values;
        const ident = String(row[iUsername] || '').trim();
        if (!ident) return;
        users.push({
          username:ident,
          nik:type === 'Peserta' ? ident : '',
          name:String(iNama>=0 ? row[iNama] || '' : ident),
          role:String(iRole>=0 ? row[iRole] || '' : type),
          unit:String(iUnit>=0 ? row[iUnit] || '' : ''),
          userTypeLabel:type
        });
      });
    };
    addUsers(APP.SHEETS.PEMAKAI,'Pemakai');
    addUsers(APP.SHEETS.USER_PESERTA,'Peserta');

    const sh = _sheet_(APP.SHEETS.USER_AKTIVITAS);
    // Finalkan pending logout yang tidak dibatalkan oleh halaman baru.
    // Saat refresh, heartbeat halaman baru menghapus cache pending terlebih dahulu.
    try {
      if (sh && sh.getLastRow() >= 2) {
        const keys = sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
        keys.forEach(r => {
          const activityKey = String(r[0] || '').trim();
          if (activityKey) _finalizePendingAutoLogoutByActivityKey_(activityKey);
        });
      }
    } catch (e) {}
    const sessionRows = [];
    if (sh && sh.getLastRow() >= 2) {
      const vals = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
      vals.forEach(r => {
        if (!r[2]) return;
        sessionRows.push({key:String(r[0]||''),type:String(r[1]||''),username:String(r[2]||''),name:String(r[3]||''),role:String(r[4]||''),unit:String(r[5]||''),login:r[6],last:r[7],logout:r[8],status:String(r[9]||'')});
      });
    }

    const now = Date.now();
    const ACTIVE_MS = 90000;
    const fmt = d => d instanceof Date && !isNaN(d.getTime()) ? Utilities.formatDate(d,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss') : '';
    const byUser = new Map();
    sessionRows.forEach(s => {
      const key = _norm_(s.type) + '|' + _norm_(s.username);
      let a = byUser.get(key);
      if (!a) {
        a = {lastLogin:null,lastActive:null,lastLogout:null,latestStatus:'',sessionCount:0};
        byUser.set(key,a);
      }
      a.sessionCount++;
      if (s.login instanceof Date && (!a.lastLogin || s.login.getTime()>a.lastLogin.getTime())) a.lastLogin=s.login;
      if (s.last instanceof Date && (!a.lastActive || s.last.getTime()>a.lastActive.getTime())) a.lastActive=s.last;
      if (s.logout instanceof Date && (!a.lastLogout || s.logout.getTime()>a.lastLogout.getTime())) a.lastLogout=s.logout;
      if (s.last instanceof Date && (!a._latestTime || s.last.getTime()>a._latestTime)) { a._latestTime=s.last.getTime(); a.latestStatus=s.status; }
    });

    const out = users.map(u => {
      const a = byUser.get(_norm_(u.userTypeLabel)+'|'+_norm_(u.username)) || {lastLogin:null,lastActive:null,lastLogout:null,latestStatus:'',sessionCount:0};
      const lastMs = a.lastActive instanceof Date ? a.lastActive.getTime() : 0;
      const active = lastMs > 0 && (now - lastMs) <= ACTIVE_MS && a.latestStatus !== 'Tidak Aktif';
      return {
        name:u.name, username:u.username, nik:u.nik, role:u.role || u.userTypeLabel, unit:u.unit,
        userTypeLabel:u.userTypeLabel, currentStatus:active ? 'Aktif' : 'Tidak Aktif',
        lastLoginAt:fmt(a.lastLogin), lastActiveAt:fmt(a.lastActive), lastLogoutAt:fmt(a.lastLogout), sessionCount:a.sessionCount
      };
    });
    out.sort((a,b) => {
      if (a.currentStatus !== b.currentStatus) return a.currentStatus === 'Aktif' ? -1 : 1;
      const at = a.lastActiveAt || '';
      const bt = b.lastActiveAt || '';
      return bt.localeCompare(at);
    });
    const active = out.filter(x => x.currentStatus === 'Aktif').length;
    return {success:true,rows:out,summary:{active,total:out.length,peserta:out.filter(x=>x.userTypeLabel==='Peserta').length,pemakai:out.filter(x=>x.userTypeLabel==='Pemakai').length}};
  } catch (err) {
    return {success:false,message:'Gagal memuat pantauan user: ' + err.message};
  }
}

/**
 * Endpoint ringan untuk foto profil setelah login berhasil.
 * Memisahkan data gambar dari payload autentikasi agar dashboard bisa terbuka lebih cepat.
 */
function getLoginProfilePhoto(userType, identity, sessionToken) {
  try {
    const type = _norm_(userType) === 'peserta' ? 'peserta' : 'pemakai';
    const id = String(identity || '').trim();
    if (!id) return {success:true, foto:''};
    if (!_verifySessionToken_(sessionToken, type, id, 12 * 60 * 60 * 1000)) {
      return {success:false, message:'Sesi tidak valid untuk memuat foto profil.'};
    }
    const cachedPhoto = _getCachedLoginPhoto_(type, id);
    if (cachedPhoto && /^data:image\//i.test(cachedPhoto)) return {success:true, foto:cachedPhoto};
    const sheetName = type === 'peserta' ? APP.SHEETS.USER_PESERTA : APP.SHEETS.PEMAKAI;
    const sh = _sheet_(sheetName);
    if (!sh || sh.getLastRow() < 2) return {success:true, foto:''};
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
    const loginIdx = _headerIndex_(headers, type === 'peserta' ? ['NIK'] : ['Username','User Name']);
    const photoIdx = _headerIndex_(headers, ['Foto Profil','Foto Profile','Foto Profi','Foto']);
    if (loginIdx < 0 || photoIdx < 0) return {success:true, foto:''};
    // Baca hanya kolom identitas dan foto, bukan seluruh baris pengguna.
    const loginValues = sh.getRange(2, loginIdx + 1, lastRow - 1, 1).getDisplayValues();
    let matchedRow = -1;
    for (let i = 0; i < loginValues.length; i++) {
      if (_norm_(loginValues[i][0]) === _norm_(id)) { matchedRow = i + 2; break; }
    }
    const photo = matchedRow > 0 ? String(sh.getRange(matchedRow, photoIdx + 1).getDisplayValue() || '').trim() : '';
    if (/^data:image\//i.test(photo) && photo.length < 95000) {
      try { CacheService.getScriptCache().put('TMS_LOGIN_PHOTO_' + _norm_(type) + '_' + _norm_(id), photo, LOGIN_AUTH_CACHE_TTL_SEC_); } catch (e) {}
    }
    return {success:true, foto:/^data:image\//i.test(photo) ? photo : ''};
  } catch (err) {
    return {success:false, message:'Gagal memuat foto profil: ' + err.message};
  }
}

function verifyLogin(loginType, username, password) {
  try {
    const type = _norm_(loginType) === 'peserta' ? 'peserta' : 'pemakai';
    const result = type === 'peserta'
      ? _findUserPeserta_(username, password)
      : _findUserPemakai_(username, password);

    if (!result.success) return result;

    result.sessionToken = _createSessionToken_(result.userType, result.userType === 'peserta' ? result.nik : result.username);
    result.logo = getSystemLogo();
    result.loginAt = new Date().toISOString();
    return result;
  } catch (err) {
    return { success: false, message: 'Error server: ' + err.message };
  }
}




function _isDeveloperRole_(role) {
  return _norm_(role) === 'developer';
}

function _isProgramerRole_(role) {
  return _norm_(role) === 'programer';
}

function verifyDeveloperAccess(username, password) {
  try {
    const result = _findUserPemakai_(username, password);
    if (!result.success) {
      return {success:false, message:'Username atau Password Developer salah.'};
    }
    if (!_isDeveloperRole_(result.role)) {
      return {success:false, message:'Akses pendaftaran hanya diperbolehkan untuk user dengan role Developer.'};
    }
    return {
      success:true,
      username:result.username,
      role:result.role
    };
  } catch (err) {
    return {success:false, message:'Gagal memverifikasi akses Developer: ' + err.message};
  }
}

/**
 * Pendaftaran akun Pemakai dari form login.
 * Mengikuti header yang sudah tersedia pada sheet UserPemakai dan tidak
 * mengganggu kolom tambahan yang mungkin dimiliki pengguna.
 */
function registerPemakai(payload) {
  try {
    payload = payload || {};
    const nama = String(payload.nama || '').trim();
    const username = String(payload.username || '').trim();
    const jabatan = String(payload.jabatan || '').trim();
    const unitUsaha = String(payload.unitUsaha || '').trim();
    const email = String(payload.email || '').trim();
    const noHp = String(payload.noHp || '').trim();
    const password = String(payload.password || '');
    const confirm = String(payload.confirm || '');
    const developerUsername = String(payload.developerUsername || '').trim();
    const developerPassword = String(payload.developerPassword || '');
    if (!developerUsername || !developerPassword) {
      return {success:false,message:'Akses pendaftaran memerlukan kredensial user Developer.'};
    }
    const developer = _findUserPemakai_(developerUsername, developerPassword);
    if (!developer.success) {
      return {success:false,message:'Kredensial Developer tidak valid.'};
    }
    if (!_isDeveloperRole_(developer.role)) {
      return {success:false,message:'Hanya user dengan role Developer yang dapat membuat akun baru.'};
    }
    if (!nama || !username || !email || !password || !confirm) return {success:false,message:'Nama, username, email, password, dan konfirmasi password wajib diisi.'};
    const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
    if (!emailPattern.test(email)) return {success:false,message:'Format email tidak valid. Email wajib menggunakan domain, contoh nama@perusahaan.com.'};
    if (password.length < 6) return {success:false,message:'Password saat pendaftaran minimal 6 karakter.'};
    if (password !== confirm) return {success:false,message:'Konfirmasi password tidak sama.'};
    const sh = _sheet_(APP.SHEETS.PEMAKAI);
    if (!sh) return {success:false,message:'Sheet UserPemakai tidak ditemukan.'};
    let lastCol = Math.max(1, sh.getLastColumn());
    if (sh.getLastRow() < 1) {
      sh.getRange(1,1,1,6).setValues([['Username','Password','Nama Lengkap','Jabatan','Unit Usaha','Email']]);
      lastCol = 6;
    }
    let headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
    const idx = (aliases) => _headerIndex_(headers, aliases);
    const ensure = (aliases, preferred) => {
      let i=idx(aliases);
      if(i>=0) return i;
      const c=headers.length;
      sh.getRange(1,c+1).setValue(preferred); headers.push(preferred); return c;
    };
    const iUsername=ensure(['Username','User Name'],'Username');
    const iPassword=ensure(['Password'],'Password');
    const iNama=ensure(['Nama Lengkap','Nama'],'Nama Lengkap');
    const iJabatan=idx(['Jabatan']);
    const iUnit=idx(['Unit Usaha','Unit']);
    const iEmail=idx(['Email']);
    const iHp=idx(['No HP','Nomor HP','No. HP']);
    const iRole=idx(['Role','Peran']);
    const iStatus=idx(['Status']);
    const data=sh.getLastRow()>=2 ? sh.getRange(2,1,sh.getLastRow()-1,headers.length).getDisplayValues() : [];
    if(data.some(r=>_norm_(r[iUsername])===_norm_(username))) return {success:false,message:'Username sudah digunakan. Silakan pilih username lain.'};
    const row=Array(headers.length).fill('');
    // Pastikan No HP selalu tersimpan sebagai teks dan tetap memiliki angka 0 di depan.
    let normalizedHp = String(noHp || '').replace(/\D/g,'');
    if (!normalizedHp) normalizedHp = '0';
    if (normalizedHp.charAt(0) !== '0') normalizedHp = '0' + normalizedHp;
    row[iUsername]=username; row[iPassword]=password; row[iNama]=nama;
    if(iJabatan>=0) row[iJabatan]=jabatan;
    if(iUnit>=0) row[iUnit]=unitUsaha;
    if(iEmail>=0) row[iEmail]=email;
    if(iHp>=0) row[iHp]=normalizedHp;
    if(iRole>=0) row[iRole]='Pemakai';
    if(iStatus>=0) row[iStatus]='Aktif';
    const targetRow = sh.getLastRow()+1;
    if(iHp>=0) sh.getRange(targetRow,iHp+1).setNumberFormat('@');
    sh.getRange(targetRow,1,1,headers.length).setValues([row]);
    SpreadsheetApp.flush();
    _invalidateLoginAuthCache_('pemakai');
    return {success:true,message:'Akun Pemakai berhasil didaftarkan.',username};
  } catch(err){ return {success:false,message:'Gagal mendaftarkan akun Pemakai: '+err.message}; }
}

/** Reset password Pemakai dengan verifikasi username + email. */
function resetPemakaiPassword(username, email, newPassword) {
  try {
    const u=String(username||'').trim(), e=String(email||'').trim(), p=String(newPassword||'');
    if(!u||!e||!p) return {success:false,message:'Username, email, dan password baru wajib diisi.'};
    if(p.length<6) return {success:false,message:'Password baru minimal 6 karakter.'};
    const sh=_sheet_(APP.SHEETS.PEMAKAI);
    if(!sh) return {success:false,message:'Sheet UserPemakai tidak ditemukan.'};
    const lastRow=sh.getLastRow(), lastCol=sh.getLastColumn();
    if(lastRow<2||lastCol<1) return {success:false,message:'Data UserPemakai belum tersedia.'};
    const headers=sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
    const iu=_headerIndex_(headers,['Username','User Name']);
    const ip=_headerIndex_(headers,['Password']);
    const ie=_headerIndex_(headers,['Email']);
    if(iu<0||ip<0||ie<0) return {success:false,message:'Kolom Username, Password, atau Email pada UserPemakai tidak ditemukan.'};
    const vals=sh.getRange(2,1,lastRow-1,lastCol).getDisplayValues();
    for(let i=0;i<vals.length;i++){
      if(_norm_(vals[i][iu])===_norm_(u) && _norm_(vals[i][ie])===_norm_(e)){
        sh.getRange(i+2,ip+1).setValue(p); SpreadsheetApp.flush();
        _invalidateLoginAuthCache_('pemakai');
        return {success:true,message:'Password berhasil diubah.'};
      }
    }
    return {success:false,message:'Username dan email tidak cocok dengan data UserPemakai.'};
  } catch(err){return {success:false,message:'Gagal reset password: '+err.message};}
}

const SHEET_CELL_CHAR_LIMIT = 50000;

function _fitSheetCellText_(value, limit) {
  const max = Number(limit) > 0 ? Number(limit) : SHEET_CELL_CHAR_LIMIT;
  const text = String(value || '');
  return text.length > max ? text.slice(0, max) : text;
}

function getSystemLogo() {
  const CACHE_KEY = 'TMS_SYSTEM_LOGO_V2';
  try {
    const cached = CacheService.getScriptCache().get(CACHE_KEY);
    if (cached) return cached === '__EMPTY__' ? '' : cached;
  } catch (e) {}
  try {
    const sh = _sheet_(APP.SHEETS.SETTINGS);
    if (!sh || sh.getLastRow() < 1) { try { CacheService.getScriptCache().put(CACHE_KEY,'__EMPTY__',600); } catch(e) {} return ''; }
    const cols = Math.max(2, sh.getLastColumn());
    const values = sh.getRange(1,1,sh.getLastRow(),cols).getDisplayValues();
    const h = values[0] || [];
    const ki = _headerIndex_(h,['Key','Kunci','Nama']);
    const vi = _headerIndex_(h,['Value','Nilai','Logo']);
    if(ki>=0 && vi>=0){
      for(let i=1;i<values.length;i++) if(_norm_(values[i][ki])==='logo') { const logo=String(values[i][vi]||''); try{CacheService.getScriptCache().put(CACHE_KEY,logo||'__EMPTY__',600);}catch(e){} return logo; }
    }
    for(let i=0;i<values.length;i++) for(let j=0;j<values[i].length;j++)
      if(String(values[i][j]||'').startsWith('data:image/')) { const logo=String(values[i][j]); try{CacheService.getScriptCache().put(CACHE_KEY,logo,600);}catch(e){} return logo; }
    try { CacheService.getScriptCache().put(CACHE_KEY,'__EMPTY__',600); } catch(e) {}
    return '';
  } catch(err){ return ''; }
}

function updateSystemLogo(username, logoBase64) {
  try {
    const result=_records_(APP.SHEETS.PEMAKAI);
    if(!result.success)return result;
    const h=result.headers, iu=_headerIndex_(h,['Username','User Name']), ir=_headerIndex_(h,['Role','Peran']);
    if(iu<0||ir<0)return {success:false,message:'Kolom Username/Role pada UserPemakai tidak ditemukan.'};
    const match=result.rows.find(x=>_norm_(x.values[iu])===_norm_(username));
    if(!match)return {success:false,message:'Akun tidak ditemukan.'};
    if(_norm_(match.values[ir])!=='programer')return {success:false,message:'Hanya user dengan role Programer yang dapat mengganti logo aplikasi.'};
    const logo=_fitSheetCellText_(logoBase64,SHEET_CELL_CHAR_LIMIT);
    if(!logo.startsWith('data:image/'))return {success:false,message:'Format logo tidak valid.'};

    let sh=_sheet_(APP.SHEETS.SETTINGS);
    if(!sh)sh=_ss_().insertSheet(APP.SHEETS.SETTINGS);
    if(sh.getLastRow()<1)sh.getRange(1,1,1,2).setValues([['Key','Value']]);

    const vals=sh.getRange(1,1,Math.max(2,sh.getLastRow()),2).getDisplayValues();
    let row=-1;
    for(let i=1;i<vals.length;i++)if(_norm_(vals[i][0])==='logo'){row=i+1;break;}
    if(row<0){row=Math.max(2,sh.getLastRow()+1);sh.getRange(row,1,1,2).setValues([['Logo',logo]]);}
    else sh.getRange(row,2).setValue(logo);
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().put('TMS_SYSTEM_LOGO_V2',logo,600); } catch(e) {}
    return {success:true,message:'Logo aplikasi berhasil diperbarui.',logo};
  }catch(err){return {success:false,message:'Gagal memperbarui logo: '+err.message};}
}

function _authorizePemakai_(username) {
  const result = _records_(APP.SHEETS.PEMAKAI);
  if (!result.success) throw new Error(result.message);

  const h = result.headers;
  const iUsername = _headerIndex_(h, ['Username', 'User Name']);
  if (iUsername < 0) throw new Error('Kolom Username pada UserPemakai tidak ditemukan.');

  const wanted = _norm_(username);
  if (!wanted) throw new Error('Sesi Pemakai tidak valid: username sesi kosong.');

  const match = result.rows.find(item => _norm_(item.values[iUsername]) === wanted);
  if (!match) {
    throw new Error('Sesi Pemakai tidak valid. Username sesi tidak ditemukan pada UserPemakai.');
  }
  return match;
}

/**
 * Otorisasi Pemakai berbasis token sesi. Token dibuat saat login dan ditandatangani
 * server, sehingga nilai identitas yang dipakai Scope tetap kanonis walaupun objek
 * currentUser pada browser sedang tidak sinkron.
 */
function _authorizePemakaiSession_(username, sessionToken) {
  const tokenPayload=_readSessionPayload_(sessionToken);
  if(tokenPayload && tokenPayload.userType==='pemakai' && tokenPayload.identity) {
    const age=Date.now()-Number(tokenPayload.iat||0);
    if(age>=0 && age<=43200000) {
      return _authorizePemakai_(tokenPayload.identity);
    }
  }

  // Fallback hanya bila username masih tersedia. Ini menjaga kompatibilitas
  // dengan pemanggilan backend lama, tetapi tidak pernah mengotorisasi sesi kosong.
  const fallbackUsername=String(username||'').trim();
  if(fallbackUsername) return _authorizePemakai_(fallbackUsername);

  throw new Error('Sesi Pemakai tidak tersedia atau sudah kedaluwarsa. Silakan login ulang.');
}

function _authorizePeserta_(nik) {
  const result = _records_(APP.SHEETS.USER_PESERTA);
  if (!result.success) throw new Error(result.message);

  const h = result.headers;
  const iNik = _headerIndex_(h, ['NIK']);
  const match = result.rows.find(item => _norm_(item.values[iNik]) === _norm_(nik));
  if (!match) throw new Error('Sesi Peserta tidak valid.');
  return match;
}
function _authorizePesertaFast_(nik) {
  const user=_findCachedLoginRow_('peserta',nik);
  if(!user)throw new Error('Sesi Peserta tidak valid.');
  return user;
}

function getKumpulanTrainingDashboard(username, filters) {
  _authorizePemakai_(username);
  filters = filters || {};
  const result = _records_(APP.SHEETS.DB_TRAINING || 'DbTraining');
  if (!result.success) return result;
  const h = result.headers || [];
  const rows = (result.rows || []).map(x => x.values || []);
  const idx = aliases => _headerIndex_(h, aliases);
  const iTraining = idx(['Nama Agenda','Agenda','Nama Training','Training','Nama Kegiatan','Kegiatan','Judul Training']);
  const iCompany = idx(['Nama Perusahaan','Perusahaan','Company','Nama Company']);
  const iPeserta = idx(['Jumlah Peserta','Total Peserta','Peserta','Jumlah Peserta Training','Total Peserta Training']);
  const iNamaPeserta = idx(['Nama Peserta','Nama Lengkap Peserta','Nama']);
  const iNik = idx(['NIK','Nik']);
  const iJam = idx(['Total Jam','Jumlah Jam','Jam','Jam Aktivitas','Durasi Jam','Durasi']);
  const iHari = idx(['Total Hari','Jumlah Hari','Hari Aktivitas','Jumlah Hari Aktivitas','Hari']);
  const iTanggal = idx(['Tanggal Pelaksanaan','Tanggal','Tanggal Mulai','Start Date','Date']);
  const iTanggalMulai = idx(['Tanggal Mulai','Start Date','Tanggal Awal','Mulai']);
  const iTanggalSelesai = idx(['Tanggal Selesai','End Date','Tanggal Akhir','Selesai']);
  const iJenis = idx(['Jenis','Jenis Training','Kategori','Category','Type','Jenis Kegiatan','Sifat']);
  const iMode = idx(['Internal / External','Internal External','Internal/External','Kategori Peserta','Scope']);
  const iMonth = idx(['Bulan','Month']);
  const iYear = idx(['Tahun','Year']);
  const iProject = idx(['Project / Training','Project/Training','Jenis Aktivitas','Activity Type','Jenis Agenda']);
  const iDetail = idx(['Detail Aktivitas','Detail Aktivitas Training','Detail Activity','Activity Detail']);
  // Jenis Agenda adalah sumber utama untuk warna/label Nama Agenda.
  // Fallback menggunakan alias Project / Training / Jenis Aktivitas bila header sheet memakai nama tersebut.
  const iJenisAgenda = idx(['Jenis Agenda']);
  const isKelasIndustriText = v => /kelas\s+industri/i.test(String(v || ''));
  // Index Batch wajib didefinisikan sebelum dipakai pada filter dan filterOptions.
  const iBatch = idx(['Batch','Batch Training','No Batch','Nomor Batch']);

  const parseNum = v => {
    const n = Number(String(v == null ? '' : v).replace(/\./g,'').replace(',','.'));
    return Number.isFinite(n) ? n : 0;
  };
  const parseDate = v => {
    if (v instanceof Date && !isNaN(v.getTime())) return v;
    const s = String(v == null ? '' : v).trim();
    if (!s) return null;

    // Dukungan format tanggal dari getDisplayValues() Google Sheets,
    // termasuk format Indonesia seperti: 04 Mei 26, 05 Agu 26, 13 Jul 26.
    const monthMap = {
      jan:0, januari:0, january:0,
      feb:1, februari:1, february:1,
      mar:2, maret:2, march:2,
      apr:3, april:3,
      mei:4, may:4,
      jun:5, juni:5, june:5,
      jul:6, juli:6, july:6,
      agu:7, agustus:7, aug:7, august:7,
      sep:8, september:8,
      okt:9, oktober:9, oct:9, october:9,
      nov:10, november:10,
      des:11, desember:11, dec:11, december:11
    };

    // dd MMM yy / dd MMM yyyy / dd-MMM-yy / dd/MMM/yyyy
    let m = s.match(/^(\d{1,2})[\s\-\/]+([A-Za-zÀ-ÿ]+)[\s\-\/]+(\d{2}|\d{4})$/i);
    if (m) {
      const day = Number(m[1]);
      const monthKey = m[2].toLowerCase();
      if (Object.prototype.hasOwnProperty.call(monthMap, monthKey)) {
        let year = Number(m[3]);
        if (year < 100) year += 2000;
        const d = new Date(year, monthMap[monthKey], day);
        if (d.getFullYear() === year && d.getMonth() === monthMap[monthKey] && d.getDate() === day) return d;
      }
    }

    // Format numerik dd/mm/yyyy atau yyyy-mm-dd.
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) return new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));

    const d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  };
  const monthNames=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const monthAliases={january:'Januari',february:'Februari',march:'Maret',april:'April',may:'Mei',june:'Juni',july:'Juli',august:'Agustus',september:'September',october:'Oktober',november:'November',december:'Desember'};
  const isInternal = v => /internal|intern|dalam/i.test(String(v || ''));
  const isExternal = v => /external|ekstern|luar/i.test(String(v || ''));
  const detailOf = r => String(iDetail>=0?r[iDetail]:'').trim();
  const classifyDetail = r => {
    const detail = detailOf(r);
    if (/kelas\s+industri/i.test(detail)) return 'Kelas Industri';
    if (/external|ekstern|luar/i.test(detail)) return 'External';
    if (/internal|intern|dalam/i.test(detail)) return 'Internal';
    return detail || 'Internal';
  };
  const categoryOf = r => classifyDetail(r);
  const trainingOf = r => String(iTraining>=0?r[iTraining]:'').trim() || 'Tanpa Nama Training';
  const companyOf = r => String(iCompany>=0?r[iCompany]:'').trim() || 'Tanpa Company';
  const monthFromValue = v => {
    const s=String(v==null?'':v).trim();
    if(!s) return '';
    const low=s.toLowerCase();
    if(monthAliases[low]) return monthAliases[low];
    const idx=monthNames.findIndex(m=>m.toLowerCase()===low);
    if(idx>=0) return monthNames[idx];
    const shortIdx=monthNames.findIndex(m=>m.slice(0,3).toLowerCase()===low.slice(0,3));
    if(shortIdx>=0 && low.length>=3) return monthNames[shortIdx];
    const n=Number(s);
    return n>=1 && n<=12 ? monthNames[n-1] : s;
  };
  const monthNumberFromValue = v => {
    const s=String(v==null?'':v).trim();
    if(!s) return 0;
    const normalized=monthFromValue(s);
    const n=monthNames.findIndex(m=>m.toLowerCase()===String(normalized).toLowerCase());
    if(n>=0) return n+1;
    const direct=Number(s);
    return direct>=1 && direct<=12 ? direct : 0;
  };
  const dateYearMonth = v => {
    const d=parseDate(v);
    return d ? {year:Utilities.formatDate(d,APP.TIMEZONE,'yyyy'),month:monthNames[d.getMonth()]} : {year:'',month:''};
  };
  const rowYear = r => {
    if(iYear>=0 && String(r[iYear]||'').trim()) return String(r[iYear]).trim();
    return dateYearMonth(iTanggal>=0?r[iTanggal]:'').year;
  };
  const rowMonth = r => {
    if(iMonth>=0 && String(r[iMonth]||'').trim()) return monthFromValue(r[iMonth]);
    return dateYearMonth(iTanggal>=0?r[iTanggal]:'').month;
  };
  const rowMonthNumber = r => {
    if(iMonth>=0 && String(r[iMonth]||'').trim()) return monthNumberFromValue(r[iMonth]);
    const d=parseDate(iTanggal>=0?r[iTanggal]:'');
    return d ? d.getMonth()+1 : 0;
  };
  const rowType = r => classifyDetail(r);
  const rowBatch = r => String(iBatch>=0 ? r[iBatch] : '').trim();
  const matchesFilter = r => {
    if(filters.year && rowYear(r)!==String(filters.year).trim()) return false;
    const monthNo=rowMonthNumber(r);
    const startMonth=monthNumberFromValue(filters.monthStart);
    const endMonth=monthNumberFromValue(filters.monthEnd);
    if(startMonth && (!monthNo || monthNo<startMonth)) return false;
    if(endMonth && (!monthNo || monthNo>endMonth)) return false;
    if(filters.company && _norm_(companyOf(r))!==_norm_(filters.company)) return false;
    if(filters.training && _norm_(trainingOf(r))!==_norm_(filters.training)) return false;
    if(filters.batch && _normalizeBatch_(rowBatch(r))!==_normalizeBatch_(filters.batch)) return false;
    if(filters.type && _norm_(rowType(r))!==_norm_(filters.type)) return false;
    return true;
  };
  const filteredRows = rows.filter(matchesFilter);
  const sourceRows = filteredRows;

  let totalPeserta=0, totalJam=0, totalHari=0;
  // TOTAL HARI AKTIVITAS dihitung per kombinasi Nama Agenda + Batch.
  // Durasi bersifat inklusif: 1 hari jika Tanggal Selesai = Tanggal Mulai.
  // TOTAL JAM AKTIVITAS kemudian dihitung dari TOTAL HARI AKTIVITAS x JAM
  // untuk setiap kombinasi Nama Agenda + Batch.
  const activityDayGroups = new Map();
  const companyMap = new Map();
  const companyDetailMap = new Map();
  const companyDetailParticipantSets = new Map();
  const companyAgendaParticipantSets = new Map();
  const companyUniqueParticipantSets = new Map();
  const agendaLabelsMap = new Map();
  const detailLabelsMap = new Map();
  const internalMap = new Map();
  const kelasIndustriMap = new Map();
  const externalMap = new Map();
  const activityMap = new Map();
  const seenDate = new Set();
  const participantFallback = new Set();
  let minDate=null,maxDate=null;

  sourceRows.forEach(r => {
    const training=trainingOf(r), company=companyOf(r), cat=categoryOf(r);
    let peserta = iPeserta>=0 ? parseNum(r[iPeserta]) : 0;
    if (peserta<=0 && iNik>=0 && String(r[iNik]||'').trim()) {
      const k=`${String(r[iNik]).trim()}|${training}|${String(r[iTanggal>=0?iTanggal:0]||'').trim()}`;
      if (!participantFallback.has(k)) { participantFallback.add(k); peserta=1; }
    }
    if (peserta<=0 && iPeserta<0) peserta=1;
    totalPeserta += peserta;

    // TOTAL JAM AKTIVITAS tidak lagi menjumlahkan kolom Jam per baris,
    // karena satu kombinasi Nama Agenda + Batch dapat muncul pada banyak baris peserta.
    // Nilai Jam disimpan satu kali pada group tersebut dan akan dikalikan dengan jumlah hari.

    // TOTAL HARI AKTIVITAS:
    // dihitung untuk setiap kombinasi Nama Agenda + Batch.
    // Nama Perusahaan tidak menjadi bagian kunci agar agenda + batch
    // yang sama tidak dihitung berulang saat muncul di beberapa perusahaan.
    const groupKey = `${_norm_(training)}|${_normalizeBatch_(rowBatch(r))}`;
    const explicitStart = iTanggalMulai>=0 ? parseDate(r[iTanggalMulai]) : null;
    const explicitEnd = iTanggalSelesai>=0 ? parseDate(r[iTanggalSelesai]) : null;
    const activityDate = parseDate(iTanggal>=0 ? r[iTanggal] : '');
    if(!activityDayGroups.has(groupKey)) activityDayGroups.set(groupKey,{start:null,end:null,jam:0});
    const dayGroup=activityDayGroups.get(groupKey);

    // Ambil JAM satu kali per Nama Agenda + Batch.
    // Umumnya nilainya sama pada seluruh baris peserta dalam group yang sama.
    // Jika ada beberapa nilai, gunakan nilai JAM positif pertama yang ditemukan.
    if(dayGroup.jam<=0 && iJam>=0){
      const jamValue=parseNum(r[iJam]);
      if(jamValue>0) dayGroup.jam=jamValue;
    }

    // Tanggal Mulai/Tanggal Selesai menjadi sumber utama.
    // Fallback ke Tanggal Pelaksanaan hanya jika kolom masing-masing tidak tersedia.
    if(explicitStart){
      if(!dayGroup.start || explicitStart < dayGroup.start) dayGroup.start = explicitStart;
    }else if(iTanggalMulai<0 && activityDate){
      if(!dayGroup.start || activityDate < dayGroup.start) dayGroup.start = activityDate;
    }
    if(explicitEnd){
      if(!dayGroup.end || explicitEnd > dayGroup.end) dayGroup.end = explicitEnd;
    }else if(iTanggalSelesai<0 && activityDate){
      if(!dayGroup.end || activityDate > dayGroup.end) dayGroup.end = activityDate;
    }
    if (iHari<0) {
      const d=parseDate(iTanggal>=0?r[iTanggal]:'');
      if (d) seenDate.add(Utilities.formatDate(d, APP.TIMEZONE, 'yyyy-MM-dd'));
    }
    const d=parseDate(iTanggal>=0?r[iTanggal]:'');
    if (d) { if(!minDate||d<minDate)minDate=d; if(!maxDate||d>maxDate)maxDate=d; }

    const detailType = classifyDetail(r);
    if(!companyMap.has(company)) companyMap.set(company,{internal:0,kelasIndustri:0,external:0,other:{},total:0});
    const cm=companyMap.get(company); cm.total+=peserta;
    if(detailType==='External') cm.external+=peserta;
    else if(detailType==='Kelas Industri') cm.kelasIndustri+=peserta;
    else if(detailType==='Internal') cm.internal+=peserta;
    else { cm.other[detailType]=(cm.other[detailType]||0)+peserta; }

    // GRAFIK JUMLAH PESERTA: unique Nama Peserta berdasarkan
    // kombinasi Nama Perusahaan + Nama Agenda (bukan penjumlahan baris).
    const rawDetail=detailOf(r) || 'Tanpa Detail Aktivitas';
    const detailKeyNorm=_norm_(rawDetail);
    if(!detailLabelsMap.has(detailKeyNorm)) detailLabelsMap.set(detailKeyNorm, rawDetail);
    if(!companyDetailMap.has(company)) companyDetailMap.set(company, new Map());
    if(!companyDetailParticipantSets.has(company)) companyDetailParticipantSets.set(company, new Map());
    const companyDetails=companyDetailMap.get(company);
    const companyDetailSets=companyDetailParticipantSets.get(company);
    if(!companyDetailSets.has(detailKeyNorm)) companyDetailSets.set(detailKeyNorm, new Set());
    const participantName=String(iNamaPeserta>=0 ? r[iNamaPeserta] : '').trim();
    const participantKey=participantName ? _norm_(participantName) : (iNik>=0 ? _norm_(r[iNik]) : '');
    if(participantKey) companyDetailSets.get(detailKeyNorm).add(participantKey);

    // Simpan unique peserta per Nama Perusahaan untuk GRAFIK JUMLAH PESERTA.
    if(!companyUniqueParticipantSets.has(company)) companyUniqueParticipantSets.set(company,new Set());
    if(participantKey) companyUniqueParticipantSets.get(company).add(participantKey);

    // Simpan unique peserta per Nama Perusahaan + Nama Agenda untuk kebutuhan dataset agenda.
    if(!companyAgendaParticipantSets.has(company)) companyAgendaParticipantSets.set(company, new Map());
    const agendaSets=companyAgendaParticipantSets.get(company);
    const agendaKey=_norm_(training);
    if(!agendaLabelsMap.has(agendaKey)) agendaLabelsMap.set(agendaKey,training);
    if(!agendaSets.has(agendaKey)) agendaSets.set(agendaKey,new Set());
    if(participantKey) agendaSets.get(agendaKey).add(participantKey);

    // Pengelompokan DETAIL AKTIVITAS tetap mengikuti kolom "Detail Aktivitas".
    // Namun WARNA/LABEL setiap Nama Agenda mengambil HANYA kolom "Jenis Agenda".
    const agendaType = String((iJenisAgenda>=0 && String(r[iJenisAgenda]||'').trim()) ? r[iJenisAgenda] : (iProject>=0 ? r[iProject] : '')).trim() || 'Training';
    const map = detailType==='Kelas Industri' ? kelasIndustriMap : (detailType==='External' ? externalMap : internalMap);
    const detailKey = `${training}|${agendaType}`;
    const current = map.get(detailKey) || {name:training, participants:0, agendaType};
    current.participants += peserta;
    map.set(detailKey,current);

    const dateKey=d ? Utilities.formatDate(d, APP.TIMEZONE, 'yyyy-MM') : (iMonth>=0 ? String(r[iMonth]||'').trim() : '');
    const activityType=String((iJenisAgenda>=0 && String(r[iJenisAgenda]||'').trim()) ? r[iJenisAgenda] : (iProject>=0 ? r[iProject] : '')).trim() || 'Training';
    const akey=`${dateKey}|${training}|${activityType}`;
    if(!activityMap.has(akey)) activityMap.set(akey,{month:dateKey,training,type:activityType,value:0,monthSort:d?d.getTime():9999999999999});
    activityMap.get(akey).value+=peserta;
  });

  // TOTAL HARI AKTIVITAS = jumlah durasi inklusif
  // (Tanggal Selesai - Tanggal Mulai) + 1 hari kalender
  // untuk setiap kombinasi Nama Agenda + Batch.
  // Contoh: 13 Jul - 13 Jul = 1 hari, 21 Apr - 22 Apr = 2 hari.
  totalHari = Array.from(activityDayGroups.values()).reduce((sum,g)=>{
    if(!g.start || !g.end) return sum;
    const diff=Math.round((g.end.getTime()-g.start.getTime())/(24*60*60*1000)) + 1;
    return sum + Math.max(0,diff);
  },0);

  // TOTAL JAM AKTIVITAS = TOTAL HARI AKTIVITAS x JAM
  // dihitung per kombinasi Nama Agenda + Batch agar Jam peserta tidak terduplikasi.
  totalJam = Array.from(activityDayGroups.values()).reduce((sum,g)=>{
    if(!g.start || !g.end) return sum;
    const diff=Math.round((g.end.getTime()-g.start.getTime())/(24*60*60*1000)) + 1;
    const hari=Math.max(0,diff);
    const jam=Number(g.jam||0);
    return sum + (hari * (Number.isFinite(jam) && jam>0 ? jam : 0));
  },0);
  const companies=Array.from(companyMap.entries()).sort((a,b)=>b[1].total-a[1].total).slice(0,18);
  const companyLabels=companies.map(x=>x[0]);
  const companyInternal=companies.map(x=>x[1].internal);
  const companyKelasIndustri=companies.map(x=>x[1].kelasIndustri);
  const companyExternal=companies.map(x=>x[1].external);
  const detailPriority=(label)=>{
    const n=_norm_(label);
    if(n==='internal') return 0;
    if(n==='external') return 1;
    if(n==='kelas industri') return 2;
    return 3;
  };
  const companyDetailLabels=Array.from(detailLabelsMap.values()).sort((a,b)=>detailPriority(a)-detailPriority(b)||String(a).localeCompare(String(b),'id'));
  const companyDetailDatasets=companyDetailLabels.map(label=>({
    label,
    data:companyLabels.map(company=>Number(companyDetailParticipantSets.get(company)?.get(_norm_(label))?.size||0))
  }));

  // Dataset grafik mengikuti NAMA AGENDA. Setiap nilai adalah UNIQUE Nama Peserta
  // berdasarkan kombinasi Nama Perusahaan + Nama Agenda.
  const companyAgendaLabels=Array.from(agendaLabelsMap.values()).sort((a,b)=>String(a).localeCompare(String(b),'id'));
  const companyAgendaDatasets=companyAgendaLabels.map(label=>({
    label,
    data:companyLabels.map(company=>Number(companyAgendaParticipantSets.get(company)?.get(_norm_(label))?.size||0))
  }));
  const detail=(map)=>Array.from(map.values()).map(x=>({
    name:x.name,
    participants:Number(x.participants||0),
    agendaType:String(x.agendaType||'Training').trim() || 'Training'
  })).sort((a,b)=>b.participants-a.participants || String(a.name).localeCompare(String(b.name),'id'));

  const activities=Array.from(activityMap.values()).sort((a,b)=>a.monthSort-b.monthSort || a.training.localeCompare(b.training,'id')).slice(0,24);
  const activityLabels=activities.map(x=>x.training);
  const activityValues=activities.map(x=>x.value);
  const activityTypes=activities.map(x=>x.type);
  const activityMonths=activities.map(x=>{
    if(/^\d{4}-\d{2}$/.test(x.month)) return `${monthNames[Number(x.month.slice(5))-1]} ${x.month.slice(0,4)}`;
    return x.month || '';
  });

  const filterStartMonth=monthNumberFromValue(filters.monthStart);
  const filterEndMonth=monthNumberFromValue(filters.monthEnd);
  const titleYear=String(filters.year || (maxDate ? Utilities.formatDate(maxDate, APP.TIMEZONE, 'yyyy') : Utilities.formatDate(new Date(), APP.TIMEZONE, 'yyyy')));
  const titleStartMonth=filterStartMonth ? monthNames[filterStartMonth-1] : (minDate ? monthNames[minDate.getMonth()] : '');
  const titleEndMonth=filterEndMonth ? monthNames[filterEndMonth-1] : (maxDate ? monthNames[maxDate.getMonth()] : '');
  const rangeText = titleStartMonth && titleEndMonth
    ? `${titleStartMonth.toUpperCase()} – ${titleEndMonth.toUpperCase()}`
    : '';

  // Nilai GRAFIK JUMLAH PESERTA = jumlah unique Nama Peserta per Nama Perusahaan.
  const companyUniqueParticipantDatasets=[{
    label:'UNIQUE PESERTA',
    data:companyLabels.map(company=>Number(companyUniqueParticipantSets.get(company)?.size||0))
  }];

  return {
    success:true,
    dashboard:{
      title: rangeText ? `DASHBOARD REPORT ${rangeText}` : 'DASHBOARD REPORT TRAINING',
      subtitle:'JHONLIN INSTITUTE TAHUN ' + titleYear,
      totalPeserta: Math.round(totalPeserta),
      totalJam: Math.round(totalJam),
      totalCompany: new Set(sourceRows.map(r=>companyOf(r)).filter(Boolean)).size,
      totalHari: Math.round(totalHari),
      totalRecords: sourceRows.length,
      internal: detail(internalMap),
      kelasIndustri: detail(kelasIndustriMap),
      external: detail(externalMap),
      companyLabels,
      companyInternal,
      companyKelasIndustri,
      companyExternal,
      companyDetailLabels,
      companyDetailDatasets,
      companyAgendaDatasets,
      companyUniqueParticipantDatasets,
      activityLabels,
      activityValues,
      activityTypes,
      activityMonths
    },
    filterOptions: {
      years: [...new Set(rows.map(rowYear).filter(Boolean))].sort((a,b)=>String(b).localeCompare(String(a))),
      months: monthNames.slice(),
      companies: [...new Set(rows.map(companyOf).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id')),
      trainings: [...new Set(rows.map(trainingOf).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id')),
      batches: [...new Set(rows.map(rowBatch).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'id', {numeric:true}))
    },
    meta:{headers:h}
  };
}

function getKumpulanTraining(username) {
  _authorizePemakai_(username);
  const result = _records_('DbTraining');
  if (!result.success) return result;
  return {
    success: true,
    headers: result.headers || [],
    rows: result.rows || [],
    total: (result.rows || []).length
  };
}

function getDashboardPemakai(username, filters) {
  _authorizePemakai_(username);
  filters = filters || {};

  const result = _records_(APP.SHEETS.DATA);
  if (!result.success) return result;

  const h = result.headers;
  const rows = result.rows.map(item => item.values);
  const iNik = _headerIndex_(h, ['NIK']);
  const iNama = _headerIndex_(h, ['Nama Peserta']);
  const iPerusahaan = _headerIndex_(h, ['Nama Perusahaan']);
  const iDepartemen = _headerIndex_(h, ['Departemen']);
  const iJabatan = _headerIndex_(h, ['Jabatan']);
  const iBatch = _headerIndex_(h, ['Batch']);
  const iTraining = _headerIndex_(h, ['Nama Training']);
  const iTanggal = _headerIndex_(h, ['Tanggal Pelaksanaan']);
  const iJenisNilai = _headerIndex_(h, ['Jenis Nilai']);
  const iNilai = _headerIndex_(h, ['Nilai']);

  const filtered = rows.filter(r => {
    const yearOk = !filters.year || String(r[iTanggal] || '').includes(String(filters.year));
    const trainingOk = !filters.training || _norm_(r[iTraining]) === _norm_(filters.training);
    const companyOk = !filters.company || _norm_(r[iPerusahaan]) === _norm_(filters.company);
    const batchOk = !filters.batch || _normalizeBatch_(r[iBatch]) === _normalizeBatch_(filters.batch);
    return yearOk && trainingOk && companyOk && batchOk;
  });

  // RATA-RATA NILAI mengikuti logika TOTAL NILAI pada menu Report.
  // Setiap peserta dibuat menjadi satu baris berdasarkan kombinasi NIK + Nama + Perusahaan + Dept + Training + Batch.
  // Komponen yang dihitung: Pre Test + Post Test setiap hari + Tugas Akhir.
  // Nilai kosong/NULL tetap menjadi 0 dan seluruh komponen tetap menjadi pembagi.
  const reportNormType = value => String(value || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ');
  const reportClassifyType = value => {
    const t = reportNormType(value);
    if (/tugas\s*akhir|final\s*task|final\s*assignment/.test(t)) return {kind:'task', final:true};
    const isPre = /pre\s*test|pretest/.test(t);
    const isPost = /post\s*test|posttest/.test(t);
    if (isPre || isPost) {
      let day = 0;
      let m = t.match(/(?:hari|day)\s*(?:ke\s*)?-?\s*(\d+)/);
      if (!m) m = t.match(/(?:^|\s)ke\s*-?\s*(\d+)/);
      if (m) day = Number(m[1] || 0);
      const wordDays = ['pertama','kedua','ketiga','keempat','kelima','keenam','ketujuh','kedelapan','kesembilan','kesepuluh','kesebelas','keduabelas','ketigabelas','keempatbelas','kelimabelas','keenambelas','ketujuhbelas','kedelapanbelas','kesembilanbelas','keduapuluh','keduapuluh satu','keduapuluh dua','keduapuluh tiga','keduapuluh empat','keduapuluh lima','keduapuluh enam','keduapuluh tujuh','keduapuluh delapan','keduapuluh sembilan','ketigapuluh'];
      if (!day) {
        const wm = t.match(/(?:hari|day)\s+(pertama|kedua|ketiga|keempat|kelima|keenam|ketujuh|kedelapan|kesembilan|kesepuluh|kesebelas|keduabelas|ketigabelas|keempatbelas|kelimabelas|keenambelas|ketujuhbelas|keduapuluh)/);
        if (wm) day = wordDays.indexOf(wm[1]) + 1;
      }
      return {kind:isPre ? 'pre' : 'post', day};
    }
    return {kind:'other', day:0};
  };
  const reportDayNumber = c => c && Number(c.day) > 0 ? Number(c.day) : 0;
  const foundDays = new Set([1,2]);
  filtered.forEach(r => {
    const c = reportClassifyType(r[iJenisNilai]);
    if ((c.kind === 'pre' || c.kind === 'post') && reportDayNumber(c) > 0) foundDays.add(reportDayNumber(c));
  });
  const reportDaysForAverage = [...foundDays].sort((a,b) => a-b);
  const scoreKey = (day, kind) => `d${Number(day)}${kind}`;
  const groupedScores = new Map();
  filtered.forEach(r => {
    const key = [r[iNik], r[iNama], r[iPerusahaan], r[iDepartemen], r[iTraining], r[iBatch]]
      .map(v => String(v ?? '').trim()).join('||');
    let g = groupedScores.get(key);
    if (!g) {
      g = {scores:{task:null}};
      reportDaysForAverage.forEach(d => {
        g.scores[scoreKey(d,'pre')] = null;
        g.scores[scoreKey(d,'post')] = null;
      });
      groupedScores.set(key, g);
    }
    const val = Number(String(r[iNilai] || '').replace(',', '.'));
    if (!Number.isFinite(val)) return;
    const c = reportClassifyType(r[iJenisNilai]);
    if (c.kind === 'task' && c.final === true && g.scores.task == null) {
      g.scores.task = val;
    } else if (c.kind === 'pre') {
      const d = reportDayNumber(c);
      if (d > 0 && reportDaysForAverage.includes(d)) {
        const k = scoreKey(d,'pre');
        if (g.scores[k] == null) g.scores[k] = val;
      }
    } else if (c.kind === 'post') {
      const d = reportDayNumber(c);
      if (d > 0 && reportDaysForAverage.includes(d)) {
        const k = scoreKey(d,'post');
        if (g.scores[k] == null) g.scores[k] = val;
      }
    }
  });
  const totalScoreKeys = reportDaysForAverage.flatMap(d => [scoreKey(d,'pre'), scoreKey(d,'post')]).concat('task');
  const perParticipantTotals = [...groupedScores.values()].map(g => {
    const sum = totalScoreKeys.reduce((acc,k) => acc + (Number.isFinite(Number(g.scores[k])) ? Number(g.scores[k]) : 0), 0);
    return totalScoreKeys.length ? sum / totalScoreKeys.length : 0;
  });
  const averageNilaiReportTotal = perParticipantTotals.length
    ? +(perParticipantTotals.reduce((a,b) => a + b, 0) / perParticipantTotals.length).toFixed(2)
    : 0;

  const countUnique = idx => new Set(filtered.map(r => String(r[idx] || '').trim()).filter(Boolean)).size;

  // Distribusi Training = jumlah peserta unik pada masing-masing training.
  // Definisi peserta mengikuti KPI TOTAL PESERTA: kombinasi NIK + jenis training.
  // Dengan demikian, satu peserta yang memiliki banyak baris nilai pada training
  // yang sama tetap dihitung satu kali pada training tersebut.
  // Jika peserta yang sama mengikuti training berbeda, masing-masing training
  // mendapatkan hitungan peserta tersendiri.
  const trainingParticipantSets = {};
  filtered.forEach(r => {
    const training = String(r[iTraining] || '').trim() || 'Tanpa Training';
    const nik = String(r[iNik] || '').trim();
    // Abaikan baris tanpa NIK agar tidak dihitung sebagai peserta.
    if (!nik) return;
    if (!trainingParticipantSets[training]) trainingParticipantSets[training] = new Set();
    trainingParticipantSets[training].add(`${_norm_(nik)}|${_norm_(training)}`);
  });
  const trainingCounts = Object.fromEntries(
    Object.entries(trainingParticipantSets).map(([training, participants]) => [training, participants.size])
  );

  const companyCounts = {};
  filtered.forEach(r => {
    const name = String(r[iPerusahaan] || '').trim() || 'Tanpa Perusahaan';
    companyCounts[name] = (companyCounts[name] || 0) + 1;
  });

  // RATA-RATA NILAI JENIS TEST PER PESERTA mengikuti komponen yang sama dengan Report.
  // Setiap peserta diringkas dari groupedScores, sehingga:
  // - Pre Test = nilai Pre Test per hari yang dipakai Report
  // - Post Test = nilai Post Test per hari yang dipakai Report
  // - Tugas Akhir = nilai Tugas Akhir yang dipakai Report
  // - Jika komponen kosong, nilainya 0 dan tetap masuk pembagi.
  // Satu kombinasi peserta + training + batch diproses satu kali seperti pada Report.
  const participantTestTotals = new Map();
  const participantTestMeta = new Map();
  const testAverageOrder = [
    {key:'pre', label:'Pre Test'},
    {key:'post', label:'Post Test'},
    {key:'task', label:'Tugas Akhir'}
  ];


  // Metadata group dibuat dari key groupedScores agar identik dengan Report.
  const dashboardGroupMeta = new Map();
  filtered.forEach(r => {
    const key = [r[iNik], r[iNama], r[iPerusahaan], r[iDepartemen], r[iTraining], r[iBatch]]
      .map(v => String(v ?? '').trim()).join('||');
    if (!dashboardGroupMeta.has(key)) {
      dashboardGroupMeta.set(key, {
        nik: String(r[iNik] || '').trim(),
        nama: String(r[iNama] || '').trim()
      });
    }
  });

  groupedScores.forEach((g, groupKey) => {
    const meta = dashboardGroupMeta.get(groupKey);
    const nik = String(meta?.nik || '').trim();
    const nama = String(meta?.nama || '').trim() || nik || 'Tanpa Nama';
    if (!nik) return;
    const participantKey = _norm_(nik);
    let item = participantTestTotals.get(participantKey);
    if (!item) {
      item = { nik, nama, sums:{pre:0,post:0,task:0}, counts:{pre:0,post:0,task:0}, groups:0 };
      participantTestTotals.set(participantKey, item);
    }
    if (!participantTestMeta.has(participantKey)) participantTestMeta.set(participantKey, {nik, nama});
    item.groups += 1;

    reportDaysForAverage.forEach(d => {
      const pre = Number(g.scores[scoreKey(d,'pre')]);
      item.sums.pre += Number.isFinite(pre) ? pre : 0;
      item.counts.pre += 1;
      const post = Number(g.scores[scoreKey(d,'post')]);
      item.sums.post += Number.isFinite(post) ? post : 0;
      item.counts.post += 1;
    });

    const task = Number(g.scores.task);
    item.sums.task += Number.isFinite(task) ? task : 0;
    item.counts.task += 1;
  });

  const participantOrderTest = Array.from(participantTestTotals.values());
  const testAverageLabels = participantOrderTest.map(p => p.nama);
  const testAverageParticipants = participantOrderTest.map(p => ({
    nik: p.nik,
    nama: p.nama,
    averages: testAverageOrder.map(type =>
      p.counts[type.key] ? +(p.sums[type.key] / p.counts[type.key]).toFixed(2) : 0
    )
  }));
  const testAverageSeries = testAverageOrder.map((type, typeIndex) => ({
    label: type.label,
    values: testAverageParticipants.map(p => p.averages[typeIndex])
  }));

  // Top 3 peserta HARUS memakai komponen Total Nilai yang sama dengan Report.
  // Report menghasilkan total per kombinasi NIK + Nama + Perusahaan + Dept + Training + Batch.
  // Untuk Top 3, satu NIK tetap ditampilkan satu kali dengan nilai rata-rata dari
  // seluruh total per-group milik NIK tersebut.
  const reportGroupMeta = new Map();
  filtered.forEach(r => {
    const key = [r[iNik], r[iNama], r[iPerusahaan], r[iDepartemen], r[iTraining], r[iBatch]]
      .map(v => String(v ?? '').trim()).join('||');
    if (!reportGroupMeta.has(key)) {
      reportGroupMeta.set(key, {
        nik: String(r[iNik] || '').trim(),
        nama: String(r[iNama] || '').trim(),
        perusahaan: String(r[iPerusahaan] || '').trim(),
        training: String(r[iTraining] || '').trim()
      });
    }
  });

  const topByNik = new Map();
  groupedScores.forEach((g, key) => {
    const totalKeys = totalScoreKeys;
    const total = totalKeys.length
      ? totalKeys.reduce((acc,k) => acc + (Number.isFinite(Number(g.scores[k])) ? Number(g.scores[k]) : 0), 0) / totalKeys.length
      : 0;
    const meta = reportGroupMeta.get(key) || {};
    const nik = String(meta.nik || '').trim();
    if (!nik) return;
    const nikKey = _norm_(nik);
    if (!topByNik.has(nikKey)) {
      topByNik.set(nikKey, {nik, nama:meta.nama || nik, totals:[], trainings:new Set(), perusahaan:new Set()});
    }
    const item = topByNik.get(nikKey);
    item.totals.push(total);
    if (meta.training) item.trainings.add(meta.training);
    if (meta.perusahaan) item.perusahaan.add(meta.perusahaan);
    if (!item.nama || item.nama===item.nik) item.nama=meta.nama || item.nama;
  });

  const topParticipants = Array.from(topByNik.values())
    .map(item => ({
      nik:item.nik,
      nama:item.nama,
      nilai:item.totals.length ? +(item.totals.reduce((a,b)=>a+b,0)/item.totals.length).toFixed(2) : 0,
      training:Array.from(item.trainings).join(' • ') || '-',
      perusahaan:Array.from(item.perusahaan).join(' • ') || '-'
    }))
    .sort((a,b)=>b.nilai-a.nilai)
    .slice(0,3);

  return {
    success: true,
    summary: {
      totalRecords: filtered.length,
      totalPeserta: new Set(filtered.map(r => `${String(r[iNik] || '').trim()}|${_norm_(r[iTraining])}`).filter(key => key !== '|')).size,
      totalPerusahaan: countUnique(iPerusahaan),
      totalTraining: countUnique(iTraining),
      averageNilai: averageNilaiReportTotal
    },
    charts: {
      trainingLabels: Object.keys(trainingCounts).slice(0, 12),
      trainingValues: Object.values(trainingCounts).slice(0, 12),
      companyLabels: Object.keys(companyCounts).slice(0, 10),
      companyValues: Object.values(companyCounts).slice(0, 10),
      testAverageLabels,
      testAverageParticipants,
      testAverageSeries
    },
    filterOptions: {
      years: [...new Set(rows.map(r => String(r[iTanggal] || '').match(/\d{4}/)?.[0]).filter(Boolean))].sort().reverse(),
      companies: [...new Set(rows.map(r => String(r[iPerusahaan] || '').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id')),
      trainings: [...new Set(rows.map(r => String(r[iTraining] || '').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id')),
      batches: [...new Set(rows.map(r => String(r[iBatch] || '').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'id'))
    },
    topParticipants,
    latestRows: result.rows
      .slice(-8)
      .reverse()
      .map(item => ({ rowNumber: item.rowNumber, values: item.values })),
    meta: { headers: h, indexes: { iNama, iDepartemen, iJabatan, iBatch, iTraining, iTanggal, iJenisNilai, iNilai } }
  };
}

function _ensureDataStatusColumn_() {
  const sh = _sheet_(APP.SHEETS.DATA);
  if (!sh) throw new Error(`Sheet "${APP.SHEETS.DATA}" tidak ditemukan.`);
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || '').trim());
  let idx = _headerIndex_(headers, ['Status Nilai', 'Status']);
  if (idx < 0) {
    idx = headers.length;
    sh.getRange(1, idx + 1).setValue('Status');
    return { sh, headers: headers.concat('Status'), index: idx };
  }
  // Standarkan nama header menjadi Status tanpa mengganggu posisi kolom yang sudah ada.
  if (String(headers[idx]).trim() !== 'Status') sh.getRange(1, idx + 1).setValue('Status');
  return { sh, headers, index: idx };
}

function _syncDataStatusFromSettingNilai_() {
  const dataSh = _sheet_(APP.SHEETS.DATA);
  if (!dataSh) return {success:false, message:`Sheet "${APP.SHEETS.DATA}" tidak ditemukan.`, updated:0};
  const settingSh = _sheet_(APP.SHEETS.SETTING_NILAI);
  if (!settingSh) {
    const info = _ensureDataStatusColumn_();
    const lastRow = info.sh.getLastRow();
    if (lastRow >= 2) info.sh.getRange(2, info.index + 1, lastRow - 1, 1).setValue('Belum Posting');
    SpreadsheetApp.flush();
    return {success:true, updated:Math.max(0,lastRow-1)};
  }

  const dataLastRow = dataSh.getLastRow();
  const dataLastCol = Math.max(1, dataSh.getLastColumn());
  const dataHeaders = dataSh.getRange(1,1,1,dataLastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  const statusIndex = _headerIndex_(dataHeaders, ['Status Nilai','Status']);
  const actualStatusIndex = statusIndex >= 0 ? statusIndex : dataHeaders.length;
  if (statusIndex < 0) dataSh.getRange(1, actualStatusIndex + 1).setValue('Status');
  else if (String(dataHeaders[statusIndex]).trim() !== 'Status') dataSh.getRange(1, actualStatusIndex + 1).setValue('Status');

  const iTraining = _headerIndex_(dataHeaders, ['Nama Training','Jenis Training','Training']);
  const iCompany  = _headerIndex_(dataHeaders, ['Nama Perusahaan','Perusahaan','Company']);
  const iBatch    = _headerIndex_(dataHeaders, ['Batch']);
  if (iTraining < 0 || iCompany < 0 || iBatch < 0) {
    throw new Error('Kolom Nama Training, Nama Perusahaan, dan Batch pada sheet Data wajib tersedia.');
  }

  const settingLastRow = settingSh.getLastRow();
  const settingLastCol = Math.max(6, settingSh.getLastColumn());
  const settingHeaders = settingSh.getRange(1,1,1,settingLastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  const siTraining = _headerIndex_(settingHeaders, ['Jenis Training','Nama Training','Training']);
  const siCompany  = _headerIndex_(settingHeaders, ['Nama Perusahaan','Unit Usaha','Perusahaan','Company']);
  const siBatch    = _headerIndex_(settingHeaders, ['Batch']);
  const siStatus   = _headerIndex_(settingHeaders, ['Status Posting','Status']);
  const postedKeys = new Set();
  if (settingLastRow >= 2 && siTraining >= 0 && siCompany >= 0 && siBatch >= 0 && siStatus >= 0) {
    const vals = settingSh.getRange(2,1,settingLastRow-1,settingLastCol).getDisplayValues();
    vals.forEach(r => {
      if (_norm_(r[siStatus]) !== 'posted') return;
      const t=String(r[siTraining]||'').trim(), c=String(r[siCompany]||'').trim(), b=String(r[siBatch]||'').trim();
      if (!t || !c || !b) return;
      postedKeys.add([_norm_(t),_norm_(c),_normalizeBatch_(b)].join('|'));
    });
  }

  if (dataLastRow < 2) return {success:true,updated:0};
  const dataValues = dataSh.getRange(2,1,dataLastRow-1,Math.max(dataLastCol, actualStatusIndex+1)).getDisplayValues();
  const statuses = dataValues.map(r => {
    const t=String(r[iTraining]||'').trim(), c=String(r[iCompany]||'').trim(), b=String(r[iBatch]||'').trim();
    const key=[_norm_(t),_norm_(c),_normalizeBatch_(b)].join('|');
    return [t && c && b && postedKeys.has(key) ? 'Posted' : 'Belum Posting'];
  });
  dataSh.getRange(2, actualStatusIndex + 1, statuses.length, 1).setValues(statuses);
  SpreadsheetApp.flush();
  return {success:true,updated:statuses.length};
}

function _setSettingNilaiAndSyncData_(training, company, batch, status) {
  const t=String(training||'').trim(), c=String(company||'').trim(), b=String(batch||'').trim();
  if(!t || !c || !b) throw new Error('Jenis Training, Nama Perusahaan, dan Batch wajib tersedia.');
  const sh=_ensureSettingNilaiSheet_();
  const last=Math.max(sh.getLastRow(),1);
  let target=-1;
  if(last>=2){
    const vals=sh.getRange(2,1,last-1,Math.max(6,sh.getLastColumn())).getDisplayValues();
    for(let i=0;i<vals.length;i++){
      const key=[_norm_(vals[i][1]),_norm_(vals[i][2]),_normalizeBatch_(vals[i][3])].join('|');
      if(key===[_norm_(t),_norm_(c),_normalizeBatch_(b)].join('|')){target=i+2;break;}
    }
  }
  const now = Utilities.formatDate(new Date(), APP.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  const id = target>0 ? (String(sh.getRange(target,1).getDisplayValue()||'') || ('NILAI-'+Utilities.getUuid().slice(0,8).toUpperCase())) : ('NILAI-'+Utilities.getUuid().slice(0,8).toUpperCase());
  if(target<0) target=sh.getLastRow()+1;
  sh.getRange(target,1,1,6).setValues([[id,t,c,b,status,status==='Posted'?now:'']]);
  _syncDataStatusFromSettingNilai_();
  SpreadsheetApp.flush();
  return {rowNumber:target,status,postedAt:status==='Posted'?now:''};
}

function getDashboardPeserta(nik) {
  _authorizePesertaFast_(nik);

  const peserta = _records_(APP.SHEETS.PESERTA);
  const data = _records_(APP.SHEETS.DATA);
  if (!peserta.success) return peserta;
  if (!data.success) return data;

  const ph = peserta.headers;
  const piNik = _headerIndex_(ph, ['NIK']);
  const myProfile = peserta.rows.find(item => _norm_(item.values[piNik]) === _norm_(nik));

  const h = data.headers;
  const iNik = _headerIndex_(h, ['NIK']);
  const iTraining = _headerIndex_(h, ['Nama Training']);
  const iTanggal = _headerIndex_(h, ['Tanggal Pelaksanaan']);
  const iBatch = _headerIndex_(h, ['Batch']);
  const iNilai = _headerIndex_(h, ['Nilai']);
  const iJenisNilai = _headerIndex_(h, ['Jenis Nilai']);
  const iStatus = _headerIndex_(h, ['Status']);
  const myRows = data.rows
    .filter(item => _norm_(item.values[iNik]) === _norm_(nik))
    .filter(item => iStatus < 0 || _norm_(item.values[iStatus]) === 'posted')
    .reverse();

  const scores = myRows
    .map(item => Number(String(item.values[iNilai] || '').replace(',', '.')))
    .filter(v => Number.isFinite(v));

  const trainingCounts = {};
  myRows.forEach(item => {
    const name = String(item.values[iTraining] || '').trim() || 'Training';
    trainingCounts[name] = (trainingCounts[name] || 0) + 1;
  });

  // Training terakhir mengikuti record terbaru yang digunakan dashboard.
  // Rata-rata nilai kartu dihitung khusus untuk seluruh nilai pada training terakhir tersebut.
  const lastTraining = myRows.length ? String(myRows[0].values[iTraining] || '-').trim() || '-' : '-';
  const lastTrainingScores = myRows
    .filter(item => String(item.values[iTraining] || '').trim() === lastTraining)
    .map(item => Number(String(item.values[iNilai] || '').replace(',', '.')))
    .filter(v => Number.isFinite(v));
  const lastTrainingAverageNilai = lastTrainingScores.length
    ? +(lastTrainingScores.reduce((a, b) => a + b, 0) / lastTrainingScores.length).toFixed(2)
    : 0;

  return {
    success: true,
    profile: myProfile ? myProfile.values : [],
    profileHeaders: ph,
    summary: {
      totalTraining: Object.keys(trainingCounts).length,
      totalRecord: myRows.length,
      averageNilai: scores.length ? +(scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(2) : 0,
      lastTraining,
      lastTrainingAverageNilai
    },
    chart: (() => {
      // Rata-rata nilai berdasarkan kombinasi Batch + Nama Training.
      // Semua jenis nilai (mis. Pre Test/Post Test) digabung dalam satu rata-rata.
      const groups = {};
      myRows.forEach(item => {
        const training = String(item.values[iTraining] || 'Training').trim() || 'Training';
        const batch = iBatch >= 0 ? String(item.values[iBatch] || '').trim() : '';
        const key = `${_norm_(batch)}|${_norm_(training)}`;
        const n = Number(String(item.values[iNilai] || '').replace(',', '.'));
        if (!Number.isFinite(n)) return;
        if (!groups[key]) groups[key] = { batch, training, sum: 0, count: 0 };
        groups[key].sum += n;
        groups[key].count += 1;
      });
      const items = Object.values(groups).sort((a,b) => {
        const ta = `${a.training} ${a.batch}`.toLowerCase();
        const tb = `${b.training} ${b.batch}`.toLowerCase();
        return ta.localeCompare(tb, 'id');
      });
      return {
        labels: items.map(x => x.batch ? `${x.training} • Batch ${x.batch}` : x.training),
        values: items.map(x => +(x.sum / x.count).toFixed(2))
      };
    })(),
    rows: myRows.map(item => ({ rowNumber: item.rowNumber, values: item.values })),
    dataHeaders: h
  };
}


function _profileSheetInfo_(userType) {
  const sheetName = userType === 'peserta' ? APP.SHEETS.USER_PESERTA : APP.SHEETS.PEMAKAI;
  const result = _records_(sheetName);
  if (!result.success) throw new Error(result.message);

  const h = result.headers;
  const loginHeaderAliases = userType === 'peserta'
    ? ['NIK']
    : ['Username', 'User Name'];
  const loginIdx = _headerIndex_(h, loginHeaderAliases);
  if (loginIdx < 0) throw new Error(`Kolom ${userType === 'peserta' ? 'NIK' : 'Username'} tidak ditemukan pada ${sheetName}.`);

  return { sheetName, result, headers: h, loginIdx };
}

function _ensureProfileColumn_(sheet, headers, aliases, preferredName) {
  let idx = _headerIndex_(headers, aliases);
  if (idx >= 0) return idx;

  const newCol = headers.length + 1;
  sheet.getRange(1, newCol).setValue(preferredName);
  headers.push(preferredName);
  return newCol - 1;
}

function getMyProfile(userType, username) {
  try {
    const type = _norm_(userType) === 'peserta' ? 'peserta' : 'pemakai';
    const info = _profileSheetInfo_(type);
    const match = info.result.rows.find(item => _norm_(item.values[info.loginIdx]) === _norm_(username));
    if (!match) return { success:false, message:'Akun tidak ditemukan.' };

    const row = match.values, h = info.headers;
    const idx = aliases => _headerIndex_(h, aliases);
    const value = aliases => { const i=idx(aliases); return i>=0 ? String(row[i] ?? '').trim() : ''; };
    let jabatan = value(['Jabatan']);
    let nama = value(['Nama Lengkap','Nama Peserta','Nama']);
    let perusahaan = value(['Nama Perusahaan','Perusahaan']);

    // Untuk peserta, UserPeserta boleh menjadi sumber utama; bila Jabatan/Nama/Perusahaan
    // kosong, ambil fallback dari Data berdasarkan NIK.
    if (type === 'peserta' && (!jabatan || !nama || !perusahaan)) {
      const data = _records_(APP.SHEETS.DATA);
      if (data.success) {
        const dh=data.headers;
        const diNik=_headerIndex_(dh,['NIK']);
        const diNama=_headerIndex_(dh,['Nama Peserta','Nama']);
        const diPerusahaan=_headerIndex_(dh,['Nama Perusahaan','Perusahaan']);
        const diJabatan=_headerIndex_(dh,['Jabatan']);
        const nik=String(row[info.loginIdx] || username).trim();
        const matches=data.rows.filter(x=>diNik>=0 && _norm_(x.values[diNik])===_norm_(nik));
        const best=matches.find(x=>diJabatan>=0 && String(x.values[diJabatan]||'').trim()) || matches[0];
        if(best){
          if(!nama && diNama>=0) nama=String(best.values[diNama]||'').trim();
          if(!perusahaan && diPerusahaan>=0) perusahaan=String(best.values[diPerusahaan]||'').trim();
          if(!jabatan && diJabatan>=0) jabatan=String(best.values[diJabatan]||'').trim();
        }
      }
    }

    const user={
      userType:type,
      username:type==='peserta' ? value(['NIK']) || String(username) : value(['Username','User Name']) || String(username),
      nik:type==='peserta' ? (value(['NIK']) || String(username)) : undefined,
      nama, jabatan,
      unitUsaha:value(['Unit Usaha','Unit']),
      perusahaan,
      email:value(['Email']),
      noHp:value(['No HP','Nomor HP','No. HP']),
      jenisKelamin:value(['Jensi Kelamin','Jenis Kelamin','Gender']),
      role:value(['Role','Peran']) || (type==='peserta'?'Peserta':'Pemakai'),
      status:value(['Status']) || 'Aktif',
      foto:value(['Foto Profil','Foto Profile','Foto'])
    };
    return {success:true,user};
  } catch(err) {
    return {success:false,message:'Gagal mengambil profil: '+err.message};
  }
}

function updateMyProfile(userType, username, noHp, newPassword, fotoBase64) {
  try {
    let normalizedNoHp = String(noHp || '').replace(/\D/g, '');
    if (normalizedNoHp) {
      if (normalizedNoHp.indexOf('62') === 0) normalizedNoHp = '0' + normalizedNoHp.slice(2);
      else if (normalizedNoHp.indexOf('8') === 0) normalizedNoHp = '0' + normalizedNoHp;
      else if (normalizedNoHp.indexOf('0') !== 0) normalizedNoHp = '0' + normalizedNoHp;
      normalizedNoHp = normalizedNoHp.slice(0, 15);
    }
    const type = _norm_(userType) === 'peserta' ? 'peserta' : 'pemakai';
    const info = _profileSheetInfo_(type);
    const sh = _sheet_(info.sheetName);
    if (!sh) return { success: false, message: `Sheet "${info.sheetName}" tidak ditemukan.` };

    const match = info.result.rows.find(item => _norm_(item.values[info.loginIdx]) === _norm_(username));
    if (!match) return { success: false, message: 'Akun tidak ditemukan.' };

    const headers = info.headers.slice();
    const iHp = _ensureProfileColumn_(sh, headers, ['No HP', 'Nomor HP', 'No. HP'], 'No HP');
    const iFoto = _ensureProfileColumn_(sh, headers, ['Foto Profil', 'Foto Profile', 'Foto'], 'Foto Profil');

    let iPassword = _headerIndex_(headers, ['Password']);
    if (iPassword < 0) throw new Error('Kolom Password tidak ditemukan.');

    const rowNumber = match.rowNumber;
    const hpRange = sh.getRange(rowNumber, iHp + 1);
    hpRange.setNumberFormat('@');
    hpRange.setValue(normalizedNoHp || '');

    if (String(newPassword || '').trim()) {
      if (String(newPassword).length < 4) {
        return { success: false, message: 'Password baru minimal 4 karakter.' };
      }
      sh.getRange(rowNumber, iPassword + 1).setValue(String(newPassword));
    }

    // Hanya simpan data URL gambar yang benar-benar dikirim oleh browser.
    if (String(fotoBase64 || '').startsWith('data:image/')) {
      const foto=_fitSheetCellText_(fotoBase64,SHEET_CELL_CHAR_LIMIT);
      sh.getRange(rowNumber, iFoto + 1).setValue(foto);
    }

    const fresh = _records_(info.sheetName);
    if (!fresh.success) return fresh;
    const updated = fresh.rows.find(item => item.rowNumber === rowNumber);
    if (!updated) return { success: false, message: 'Profil gagal dibaca ulang.' };

    const row = updated.values;
    const iNama = _headerIndex_(fresh.headers, ['Nama Lengkap', 'Nama Peserta', 'Nama']);
    const iJabatan = _headerIndex_(fresh.headers, ['Jabatan']);
    const iUnit = _headerIndex_(fresh.headers, ['Unit Usaha', 'Unit']);
    const iPerusahaan = _headerIndex_(fresh.headers, ['Nama Perusahaan', 'Perusahaan']);
    const iEmail = _headerIndex_(fresh.headers, ['Email']);
    const iGender = _headerIndex_(fresh.headers, ['Jensi Kelamin', 'Jenis Kelamin', 'Gender']);
    const iRole = _headerIndex_(fresh.headers, ['Role', 'Peran']);
    const iStatus = _headerIndex_(fresh.headers, ['Status']);
    const iFotoFresh = _headerIndex_(fresh.headers, ['Foto Profil', 'Foto Profile', 'Foto']);
    const iHpFresh = _headerIndex_(fresh.headers, ['No HP', 'Nomor HP', 'No. HP']);

    const user = {
      userType: type,
      username: type === 'peserta'
        ? String(row[_headerIndex_(fresh.headers, ['NIK'])] || username)
        : String(row[_headerIndex_(fresh.headers, ['Username', 'User Name'])] || username),
      nama: String(iNama >= 0 ? row[iNama] : ''),
      jabatan: String(iJabatan >= 0 ? row[iJabatan] : ''),
      unitUsaha: String(iUnit >= 0 ? row[iUnit] : ''),
      perusahaan: String(iPerusahaan >= 0 ? row[iPerusahaan] : ''),
      email: String(iEmail >= 0 ? row[iEmail] : ''),
      noHp: String(iHpFresh >= 0 ? row[iHpFresh] : ''),
      jenisKelamin: String(iGender >= 0 ? row[iGender] : ''),
      role: String(iRole >= 0 ? row[iRole] : (type === 'peserta' ? 'Peserta' : 'Pemakai')),
      status: String(iStatus >= 0 ? row[iStatus] : 'Aktif'),
      foto: String(iFotoFresh >= 0 ? row[iFotoFresh] : ''),
      logo: getSystemLogo()
    };

    _invalidateLoginAuthCache_(type);
    try { CacheService.getScriptCache().remove('TMS_LOGIN_PHOTO_' + _norm_(type) + '_' + _norm_(username)); } catch (e) {}
    return { success: true, message: 'Profil berhasil diperbarui.', user };
  } catch (err) {
    return { success: false, message: 'Gagal memperbarui profil: ' + err.message };
  }
}

function _normalizeBatch_(value) {
  return _norm_(value).replace(/^batch\s*/i,'').replace(/\s+/g,' ');
}

function _parseDateValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    // Normalisasi Date object ke tanggal lokal Asia/Makassar.
    const ds = Utilities.formatDate(value, APP.TIMEZONE, 'yyyy-MM-dd');
    const parts = ds.split('-').map(Number);
    return {y:parts[0], m:parts[1], d:parts[2]};
  }
  const s = String(value == null ? '' : value).trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return {y:Number(m[1]), m:Number(m[2]), d:Number(m[3])};
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return {y:Number(m[3]), m:Number(m[2]), d:Number(m[1])};
  const parsed = new Date(s);
  if (isNaN(parsed.getTime())) return null;
  const ds = Utilities.formatDate(parsed, APP.TIMEZONE, 'yyyy-MM-dd');
  const parts = ds.split('-').map(Number);
  return {y:parts[0], m:parts[1], d:parts[2]};
}

function _deadlineFromRawValues_(dateValue, timeValue) {
  const p = _parseDateValue_(dateValue);
  if (!p) return null;
  let hh = 23, mm = 59;
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    const ts = Utilities.formatDate(timeValue, APP.TIMEZONE, 'HH:mm');
    const tm = ts.match(/^(\d{2}):(\d{2})$/);
    if (tm) { hh=Number(tm[1]); mm=Number(tm[2]); }
  } else {
    const t = String(timeValue == null ? '' : timeValue).trim();
    const m = t.match(/^(\d{1,2}):(\d{2})$/);
    if (m) { hh=Number(m[1]); mm=Number(m[2]); }
  }
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return null;
  // Asia/Makassar = UTC+08:00. Simpan sebagai absolute instant agar hasil sama
  // walaupun timezone project Apps Script berbeda.
  const utcMillis = Date.UTC(p.y, p.m-1, p.d, hh-8, mm, 0, 0);
  return new Date(utcMillis);
}

function _isTaskActiveRaw_(dateValue, timeValue) {
  const deadline = _deadlineFromRawValues_(dateValue, timeValue);
  return !deadline || Date.now() <= deadline.getTime();
}

function _taskDeadlineStatusMap_() {
  const sh = _ensureTaskSettingSheet_();
  const last = sh.getLastRow();
  const map = new Map();
  if (last < 2) return map;
  // Gunakan nilai mentah agar tanggal/jam tidak bergantung pada locale tampilan Sheet.
  const vals = sh.getRange(2,1,last-1,6).getValues();
  vals.forEach(r => {
    const company = String(r[1] || '').trim();
    const training = String(r[2] || '').trim();
    const batch = String(r[3] || '').trim();
    if (!company || !training || !batch) return;
    const key = [_norm_(training), _norm_(company), _normalizeBatch_(batch)].join('|');
    const active = _isTaskActiveRaw_(r[4], r[5]);
    map.set(key, Boolean(map.get(key) || active));
  });
  return map;
}

function _syncUserPesertaStatusFromDeadline_() {
  const dataRes = _records_(APP.SHEETS.DATA);
  if (!dataRes.success) throw new Error(dataRes.message);
  const userRes = _records_(APP.SHEETS.USER_PESERTA);
  if (!userRes.success) throw new Error(userRes.message);

  const dh = dataRes.headers;
  const uh = userRes.headers;
  const diNik = _headerIndex_(dh, ['NIK']);
  const diTraining = _headerIndex_(dh, ['Jenis Training','Nama Training','Training']);
  const diCompany = _headerIndex_(dh, ['Nama Perusahaan','Perusahaan','Company']);
  const diBatch = _headerIndex_(dh, ['Batch']);
  const uiNik = _headerIndex_(uh, ['NIK']);
  const uiStatus = _headerIndex_(uh, ['Status']);
  if (diNik < 0 || diTraining < 0 || diCompany < 0 || diBatch < 0) {
    return {updated:0, skipped:true, reason:'Kolom Data tidak lengkap'};
  }
  if (uiNik < 0 || uiStatus < 0) {
    return {updated:0, skipped:true, reason:'Kolom UserPeserta tidak lengkap'};
  }

  const deadlineMap = _taskDeadlineStatusMap_();
  const participantState = new Map();

  // Mulai dari seluruh peserta yang terdaftar agar peserta yang TIDAK memiliki
  // kombinasi training + perusahaan + batch pada SettingTugas aktif juga
  // otomatis diproses menjadi Non Aktif.
  userRes.rows.forEach(item => {
    const nik = String(item.values[uiNik] || '').trim();
    if (nik) participantState.set(_norm_(nik), {active:false, matched:false});
  });

  dataRes.rows.forEach(item => {
    const r = item.values;
    const nik = String(r[diNik] || '').trim();
    const training = String(r[diTraining] || '').trim();
    const company = String(r[diCompany] || '').trim();
    const batch = String(r[diBatch] || '').trim();
    if (!nik || !training || !company || !batch) return;

    const key = [_norm_(training), _norm_(company), _normalizeBatch_(batch)].join('|');
    const nkey = _norm_(nik);
    if (!participantState.has(nkey)) {
      participantState.set(nkey, {active:false, matched:false});
    }
    const state = participantState.get(nkey);

    // Hanya kombinasi yang ada di SettingTugas dan deadline-nya masih aktif
    // yang boleh membuat akun peserta tetap Aktif.
    if (deadlineMap.has(key) && deadlineMap.get(key) === true) {
      state.matched = true;
      state.active = true;
    }
  });

  const sh = _sheet_(APP.SHEETS.USER_PESERTA);
  let updated = 0;
  userRes.rows.forEach(item => {
    const nik = String(item.values[uiNik] || '').trim();
    if (!nik) return;
    const state = participantState.get(_norm_(nik));

    // Aktif hanya jika peserta memiliki minimal satu kombinasi yang benar-benar
    // masuk SettingTugas dan masih berstatus aktif. Selain itu wajib Non Aktif.
    const newStatus = state && state.active ? 'Aktif' : 'Non Aktif';
    const oldStatus = String(item.values[uiStatus] || '').trim();
    if (oldStatus !== newStatus) {
      sh.getRange(item.rowNumber, uiStatus + 1).setValue(newStatus);
      updated++;
    }
  });
  if (updated) SpreadsheetApp.flush();
  return {updated, skipped:false};
}

function refreshUserPesertaDeadlineStatus(username) {
  _authorizePemakai_(username);
  const sync = _syncUserPesertaFromData_();
  return {success:true, message:'Status UserPeserta berhasil disegarkan.', detail:sync};
}

// Fungsi publik untuk memaksa sinkronisasi status dari deadline.
function setupUserPesertaDeadlineSync() {
  // Jalankan sekali sebagai user/owner proyek untuk membuat trigger berkala.
  const fn = 'syncUserPesertaStatusAutomatically';
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === fn).forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger(fn).timeBased().everyMinutes(5).create();
  syncUserPesertaStatusAutomatically();
  return {success:true,message:'Sinkronisasi status UserPeserta otomatis setiap 5 menit telah diaktifkan.'};
}

function syncUserPesertaStatusAutomatically() {
  try {
    _syncUserPesertaFromData_();
    return {success:true};
  } catch (err) {
    return {success:false,message:err.message};
  }
}

function syncUserPesertaStatus(username) {
  _authorizePemakai_(username);
  _syncUserPesertaFromData_();
  return {success:true, ..._syncUserPesertaStatusFromDeadline_()};
}

/**
 * Sinkronisasi master DataPeserta dari sheet Data.
 * Sumber kolom:
 * NIK, Nama Peserta, Nama Perusahaan, Departemen, Jabatan.
 *
 * Aturan:
 * - Satu peserta hanya boleh muncul satu kali berdasarkan NIK.
 * - Jika NIK muncul beberapa kali di Data, nilai non-kosong pertama
 *   untuk masing-masing kolom dipertahankan.
 * - Isi DataPeserta untuk 5 kolom master selalu dibuat ulang dari Data,
 *   sehingga penambahan/perubahan/penghapusan pada Data tercermin otomatis.
 */
function _syncDataPesertaFromData_() {
  const dataRes = _records_(APP.SHEETS.DATA);
  if (!dataRes.success) throw new Error(dataRes.message);

  const dataH = dataRes.headers || [];
  const iNik = _headerIndex_(dataH, ['NIK']);
  const iNama = _headerIndex_(dataH, ['Nama Peserta', 'Nama']);
  const iPerusahaan = _headerIndex_(dataH, ['Nama Perusahaan', 'Perusahaan']);
  const iDepartemen = _headerIndex_(dataH, ['Departemen', 'Department']);
  const iJabatan = _headerIndex_(dataH, ['Jabatan', 'Position']);

  const missing = [];
  if (iNik < 0) missing.push('NIK');
  if (iNama < 0) missing.push('Nama Peserta');
  if (iPerusahaan < 0) missing.push('Nama Perusahaan');
  if (iDepartemen < 0) missing.push('Departemen');
  if (iJabatan < 0) missing.push('Jabatan');
  if (missing.length) {
    throw new Error('Kolom berikut pada sheet Data tidak ditemukan: ' + missing.join(', ') + '.');
  }

  let sh = _sheet_(APP.SHEETS.PESERTA);
  if (!sh) sh = _ss_().insertSheet(APP.SHEETS.PESERTA);

  const masterHeaders = ['NIK', 'Nama Peserta', 'Nama Perusahaan', 'Departemen', 'Jabatan'];
  const requiredCols = masterHeaders.length;

  // Pastikan header DataPeserta konsisten dengan struktur master yang diminta.
  if (sh.getMaxColumns() < requiredCols) {
    sh.insertColumnsAfter(sh.getMaxColumns(), requiredCols - sh.getMaxColumns());
  }
  sh.getRange(1, 1, 1, requiredCols).setValues([masterHeaders]);

  // Gabungkan baris Data yang memiliki NIK sama menjadi satu master peserta.
  const byNik = new Map();
  const sourceRows = dataRes.rows || [];
  sourceRows.forEach(item => {
    const r = item.values || [];
    const nik = String(r[iNik] == null ? '' : r[iNik]).trim();
    if (!nik) return;

    const key = _norm_(nik);
    const current = byNik.get(key) || {
      nik: '', nama: '', perusahaan: '', departemen: '', jabatan: ''
    };

    // Nilai dari kemunculan pertama yang tidak kosong menjadi nilai master.
    if (!current.nik) current.nik = nik;
    if (!current.nama) current.nama = String(r[iNama] == null ? '' : r[iNama]).trim();
    if (!current.perusahaan) current.perusahaan = String(r[iPerusahaan] == null ? '' : r[iPerusahaan]).trim();
    if (!current.departemen) current.departemen = String(r[iDepartemen] == null ? '' : r[iDepartemen]).trim();
    if (!current.jabatan) current.jabatan = String(r[iJabatan] == null ? '' : r[iJabatan]).trim();

    byNik.set(key, current);
  });

  const output = Array.from(byNik.values()).map(x => [
    x.nik, x.nama, x.perusahaan, x.departemen, x.jabatan
  ]);

  // DataPeserta adalah master turunan dari Data; hapus isi lama agar tidak ada
  // peserta yang tertinggal dan tidak ada baris duplikat.
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, requiredCols).clearContent();
  }

  if (output.length) {
    sh.getRange(2, 1, output.length, requiredCols).setValues(output);
  }

  // Rapikan format NIK sebagai teks agar angka 0 di depan tetap aman.
  sh.getRange(2, 1, Math.max(1, output.length), 1).setNumberFormat('@');
  SpreadsheetApp.flush();

  return {
    success: true,
    synchronized: output.length,
    sourceRows: sourceRows.length
  };
}

/** Sinkronisasi yang dapat dipanggil dari UI/server. */
function syncDataPeserta(username) {
  _authorizePemakai_(username);
  return _syncDataPesertaFromData_();
}

/**
 * Sinkronisasi ringan saat spreadsheet sumber Data diedit langsung.
 * Tidak memicu dirinya sendiri karena perubahan dari script tidak menjalankan onEdit.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    // Perubahan pada Data harus memperbarui master DataPeserta sekaligus
    // menghitung ulang status UserPeserta berdasarkan kombinasi training aktif.
    if (sheetName === APP.SHEETS.DATA) {
      _syncDataPesertaFromData_();
      try { _syncUserPesertaFromData_(); } catch (statusErr) {}
      return;
    }

    // Perubahan pada SettingTugas juga langsung menghitung ulang status
    // UserPeserta agar peserta tanpa tugas aktif segera menjadi Non Aktif.
    if (sheetName === APP.SHEETS.SETTING_TUGAS) {
      try { _syncUserPesertaStatusFromDeadline_(); } catch (statusErr) {
        console.error('Gagal memperbarui status UserPeserta: ' + statusErr.message);
      }
    }
  } catch (err) {
    console.error('Gagal sinkronisasi otomatis: ' + err.message);
  }
}

function _syncUserPesertaFromData_() {
  const dataRes = _records_(APP.SHEETS.DATA);
  if (!dataRes.success) throw new Error(dataRes.message);
  const dataH = dataRes.headers;
  const iNik = _headerIndex_(dataH, ['NIK']);
  const iNama = _headerIndex_(dataH, ['Nama Peserta','Nama']);
  const iPerusahaan = _headerIndex_(dataH, ['Nama Perusahaan','Perusahaan']);
  if (iNik < 0 || iNama < 0 || iPerusahaan < 0) {
    throw new Error('Kolom NIK, Nama Peserta, atau Nama Perusahaan pada sheet Data tidak ditemukan.');
  }

  let sh = _sheet_(APP.SHEETS.USER_PESERTA);
  if (!sh) sh = _ss_().insertSheet(APP.SHEETS.USER_PESERTA);

  let lastCol = Math.max(3, sh.getLastColumn());
  if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,4).setValues([['NIK','Nama Peserta','Nama Perusahaan','Password']]);
    lastCol = 4;
  }

  let headers = sh.getRange(1,1,1,Math.max(4,lastCol)).getDisplayValues()[0].map(v=>String(v||'').trim());
  const ensureCol = (aliases, preferred) => {
    let idx = _headerIndex_(headers, aliases);
    if (idx >= 0) return idx;
    const c = headers.length + 1;
    sh.getRange(1,c).setValue(preferred);
    headers.push(preferred);
    return c-1;
  };
  const uNik = ensureCol(['NIK'], 'NIK');
  const uNama = ensureCol(['Nama Peserta','Nama'], 'Nama Peserta');
  const uPerusahaan = ensureCol(['Nama Perusahaan','Perusahaan'], 'Nama Perusahaan');
  const uPassword = ensureCol(['Password'], 'Password');
  const uStatus = ensureCol(['Status'], 'Status');

  const last = sh.getLastRow();
  const existing = last >= 2 ? sh.getRange(2,1,last-1,headers.length).getDisplayValues() : [];
  const byNik = new Map();
  existing.forEach((r,i)=>{
    const key=_norm_(r[uNik]);
    if(key) byNik.set(key,{rowNumber:i+2,row:r});
  });

  const source = dataRes.rows.map(x=>x.values);
  const seen = new Set();
  const updates = [];
  source.forEach(r=>{
    const nik=String(r[iNik]||'').trim();
    const nama=String(r[iNama]||'').trim();
    const perusahaan=String(r[iPerusahaan]||'').trim();
    if(!nik || !nama || !perusahaan) return;
    const key=_norm_(nik);
    if(seen.has(key)) return;
    seen.add(key);
    const found=byNik.get(key);
    if(found){
      const rowNum=found.rowNumber;
      const oldPass=String(found.row[uPassword]||'').trim();
      sh.getRange(rowNum,uNik+1).setValue(nik);
      sh.getRange(rowNum,uNama+1).setValue(nama);
      sh.getRange(rowNum,uPerusahaan+1).setValue(perusahaan);
      if(!oldPass) sh.getRange(rowNum,uPassword+1).setValue('user123');
      const oldStatus=String(found.row[uStatus]||'').trim();
      if(!oldStatus) sh.getRange(rowNum,uStatus+1).setValue('Aktif');
    } else {
      const out=Array(headers.length).fill('');
      out[uNik]=nik; out[uNama]=nama; out[uPerusahaan]=perusahaan; out[uPassword]='user123'; out[uStatus]='Aktif';
      updates.push(out);
    }
  });
  if(updates.length){
    const start=Math.max(2,sh.getLastRow()+1);
    sh.getRange(start,1,updates.length,headers.length).setValues(updates);
  }
  const deadlineSync = _syncUserPesertaStatusFromDeadline_();
  _invalidateLoginAuthCache_('peserta');
  SpreadsheetApp.flush();
  return {success:true, synchronized:seen.size, added:updates.length, statusUpdated:deadlineSync.updated};
}

function syncUserPeserta(username) {
  _authorizePemakai_(username);
  return _syncUserPesertaFromData_();
}

function _calendarDateKeyServer_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, APP.TIMEZONE, 'yyyy-MM-dd');
  const s=String(value||'').trim();
  if(!s) return '';
  let m=s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  m=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if(m) return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  const months={januari:1,februari:2,maret:3,april:4,mei:5,juni:6,juli:7,agustus:8,september:9,oktober:10,november:11,desember:12,jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,sept:9,oct:10,nov:11,dec:12};
  m=s.toLowerCase().match(/^(\d{1,2})\s+([a-z]+)\s+(\d{4})/);
  if(m && months[m[2]]) return `${m[3]}-${String(months[m[2]]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  return '';
}

function getDatabaseTable(sheetName, username) {
  _assertDatabaseScope_(sheetName, username);
  if (sheetName === 'DataPeserta') _syncDataPesertaFromData_();
  if (sheetName === 'UserPeserta') _syncUserPesertaFromData_();
  if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
  if (!APP.DB_KEYS[sheetName]) return { success: false, message: 'Sheet database tidak diizinkan.' };
  return _records_(APP.DB_KEYS[sheetName]);
}

function getDatabaseExportData(sheetName, username, search) {
  _assertDatabaseScope_(sheetName, username);
  if (sheetName === 'DataPeserta') _syncDataPesertaFromData_();
  if (sheetName === 'UserPeserta') _syncUserPesertaFromData_();
  if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
  if (sheetName === 'Tugas') _syncTugasFromSetting_();
  if (!APP.DB_KEYS[sheetName]) return { success:false, message:'Sheet database tidak diizinkan.' };

  const result = _records_(APP.DB_KEYS[sheetName]);
  if (!result.success) return result;

  const q = _norm_(search || '');
  const rows = !q ? result.rows : result.rows.filter(item =>
    (item.values || []).some(v => _norm_(v).includes(q))
  );

  return {
    success:true,
    sheetName,
    title: APP.DB_KEYS[sheetName],
    headers: result.headers || [],
    rows
  };
}

function getDatabasePage(sheetName, username, page, pageSize, search) {
  _assertDatabaseScope_(sheetName, username);
  if (sheetName === 'DataPeserta') _syncDataPesertaFromData_();
  if (sheetName === 'UserPeserta') _syncUserPesertaFromData_();
  if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
  if (sheetName === 'Tugas') _syncTugasFromSetting_();
  if (!APP.DB_KEYS[sheetName]) return { success:false, message:'Sheet database tidak diizinkan.' };

  const sh = _sheet_(APP.DB_KEYS[sheetName]);
  if (!sh) return { success:false, message:`Sheet "${APP.DB_KEYS[sheetName]}" tidak ditemukan.` };

  page = Math.max(1, Number(page) || 1);
  pageSize = Math.min(10, Math.max(1, Number(pageSize) || 10));
  search = _norm_(search || '');

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return {success:true,headers:[],rows:[],total:0,page:1,pageSize};
  }

  const headers = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());

  // Jalur normal: hanya baca 10 baris untuk halaman yang diminta.
  if (!search) {
    const total = Math.max(0,lastRow-1);
    const startRow = 2 + ((page-1)*pageSize);
    const count = startRow <= lastRow ? Math.min(pageSize,lastRow-startRow+1) : 0;
    const values = count ? sh.getRange(startRow,1,count,lastCol).getDisplayValues() : [];
    const rows = values.map((row,i)=>({rowNumber:startRow+i,values:row}))
      .filter(item=>item.values.some(v=>String(v||'').trim()!==''));
    return {success:true,headers,rows,total,page,pageSize};
  }

  // Jalur pencarian: scan sheet hanya ketika benar-benar diperlukan.
  const values = lastRow>1 ? sh.getRange(2,1,lastRow-1,lastCol).getDisplayValues() : [];
  const matched=[];
  for(let i=0;i<values.length;i++){
    const row=values[i];
    if(row.some(v=>_norm_(v).includes(search))) matched.push({rowNumber:i+2,values:row});
  }
  const total=matched.length;
  const start=(page-1)*pageSize;
  return {success:true,headers,rows:matched.slice(start,start+pageSize),total,page,pageSize};
}


function saveDatabaseRow(sheetName, rowNumber, values, username) {
  _assertDatabaseScope_(sheetName, username);
  const realSheet = APP.DB_KEYS[sheetName];
  if (!realSheet) return { success: false, message: 'Sheet database tidak diizinkan.' };

  const sh = _sheet_(realSheet);
  if (!sh) return { success: false, message: `Sheet "${realSheet}" tidak ditemukan.` };

  // Data selalu mempunyai kolom Status yang dikendalikan oleh SettingNilai.
  if (sheetName === 'Data') _ensureDataStatusColumn_();
  const lastCol = Math.max(sh.getLastColumn(), values.length);
  const row = Array.from({ length: lastCol }, (_, i) => values[i] == null ? '' : values[i]);

  // No. HP wajib disimpan sebagai TEKS agar angka 0 di depan tidak dihilangkan
  // oleh Google Sheets (misalnya 08123456789 tetap 08123456789).
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(v => String(v || '').trim());
  const iHp = _headerIndex_(headers, ['No HP', 'Nomor HP', 'No. HP']);
  if (iHp >= 0) {
    let hp = String(row[iHp] == null ? '' : row[iHp]).replace(/\D/g, '');
    if (hp) {
      if (hp.indexOf('62') === 0) hp = '0' + hp.slice(2);
      else if (hp.indexOf('8') === 0) hp = '0' + hp;
      else if (hp.indexOf('0') !== 0) hp = '0' + hp;
      hp = hp.slice(0, 15);
    }
    row[iHp] = hp;
  }

  if (rowNumber && Number(rowNumber) > 1) {
    const targetRow = Number(rowNumber);
    sh.getRange(targetRow, 1, 1, lastCol).clearContent();
    if (iHp >= 0) sh.getRange(targetRow, iHp + 1).setNumberFormat('@');
    sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
    if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
    if (sheetName === 'UserPemakai') _invalidateLoginAuthCache_('pemakai');
    if (sheetName === 'UserPeserta') _invalidateLoginAuthCache_('peserta');
    return { success: true, message: 'Data berhasil diperbarui.' };
  }

  const targetRow = sh.getLastRow() + 1;
  if (iHp >= 0) sh.getRange(targetRow, iHp + 1).setNumberFormat('@');
  sh.getRange(targetRow, 1, 1, row.length).setValues([row]);
  if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
  if (sheetName === 'UserPemakai') _invalidateLoginAuthCache_('pemakai');
  if (sheetName === 'UserPeserta') _invalidateLoginAuthCache_('peserta');
  return { success: true, message: 'Data baru berhasil ditambahkan.' };
}

function deleteDatabaseRow(sheetName, rowNumber, username) {
  _assertDatabaseScope_(sheetName, username);
  const realSheet = APP.DB_KEYS[sheetName];
  if (!realSheet) return { success: false, message: 'Sheet database tidak diizinkan.' };

  const sh = _sheet_(realSheet);
  if (!sh) return { success: false, message: `Sheet "${realSheet}" tidak ditemukan.` };

  const n = Number(rowNumber);
  if (!Number.isFinite(n) || n < 2 || n > sh.getLastRow()) {
    return { success: false, message: 'Nomor baris tidak valid.' };
  }
  sh.deleteRow(n);
  if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
  if (sheetName === 'UserPemakai') _invalidateLoginAuthCache_('pemakai');
  if (sheetName === 'UserPeserta') _invalidateLoginAuthCache_('peserta');
  return { success: true, message: 'Data berhasil dihapus.' };
}

/**
 * Import data XLSX/CSV yang sudah diparse di browser.
 * Kolom dipetakan berdasarkan nama header yang dinormalisasi.
 * Operasi bersifat APPEND: tidak menghapus/menimpa data lama.
 */
function importDatabaseRows(sheetName, headers, rows, username) {
  try {
    _assertDatabaseScope_(sheetName, username);
    const realSheet = APP.DB_KEYS[sheetName];
    if (!realSheet) return { success:false, message:'Sheet database tidak diizinkan.' };

    const sh = _sheet_(realSheet);
    if (!sh) return { success:false, message:`Sheet "${realSheet}" tidak ditemukan.` };

    if (sheetName === 'Data') _ensureDataStatusColumn_();
    const targetLastCol = sh.getLastColumn();
    if (targetLastCol < 1) return { success:false, message:`Sheet "${realSheet}" belum memiliki header.` };

    const targetHeaders = sh.getRange(1,1,1,targetLastCol).getDisplayValues()[0]
      .map(v => String(v == null ? '' : v).trim());
    const sourceHeaders = (headers || []).map(v => String(v == null ? '' : v).trim());
    if (!sourceHeaders.length) return { success:false, message:'Header XLSX tidak ditemukan.' };

    const sourceIndex = {};
    sourceHeaders.forEach((h,i) => {
      const n = _norm_(h);
      if (n && sourceIndex[n] == null) sourceIndex[n] = i;
    });

    const matched = targetHeaders.filter(h => h && sourceIndex[_norm_(h)] != null);
    const unmatchedTarget = targetHeaders.filter(h => h && sourceIndex[_norm_(h)] == null);
    if (!matched.length) {
      return { success:false, message:'Tidak ada kolom XLSX yang cocok dengan header database "' + realSheet + '". Gunakan Master Template sebagai dasar import.' };
    }

    const cleanRows = (rows || []).filter(row =>
      Array.isArray(row) && row.some(v => String(v == null ? '' : v).trim() !== '')
    );
    if (!cleanRows.length) return { success:false, message:'Tidak ada data yang dapat diimpor.' };
    if (cleanRows.length > 5000) return { success:false, message:'Maksimal 5.000 baris per proses import.' };

    const output = cleanRows.map(srcRow => targetHeaders.map(h => {
      const srcIdx = sourceIndex[_norm_(h)];
      return srcIdx == null ? '' : (srcRow[srcIdx] == null ? '' : srcRow[srcIdx]);
    }));

    const startRow = Math.max(2, sh.getLastRow() + 1);
    sh.getRange(startRow, 1, output.length, targetHeaders.length).setValues(output);
    if (sheetName === 'Data') _syncDataStatusFromSettingNilai_();
    if (sheetName === 'UserPemakai') _invalidateLoginAuthCache_('pemakai');
    if (sheetName === 'UserPeserta') _invalidateLoginAuthCache_('peserta');
    SpreadsheetApp.flush();

    return {
      success:true,
      message:`Import berhasil: ${output.length} baris ditambahkan ke ${realSheet}.`,
      imported:output.length,
      targetSheet:realSheet,
      matchedColumns:matched.length,
      unmatchedTargetColumns:unmatchedTarget
    };
  } catch (err) {
    return { success:false, message:'Gagal import XLSX: ' + err.message };
  }
}




function _ensureTaskSettingSheet_() {
  const ss = _ss_();
  let sh = ss.getSheetByName(APP.SHEETS.SETTING_TUGAS);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.SETTING_TUGAS);
  if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,6).setValues([['ID','Nama Perusahaan','Jenis Training','Batch','Tanggal','Jam']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function getTaskSettingOptions(username) {
  _authorizePemakai_(username);
  try {
    const result = {companies: [], trainings: [], batches: []};
    const records = {};
    [APP.SHEETS.PERUSAHAAN, APP.SHEETS.DATA].forEach(name => {
      const r = _records_(name);
      if (r.success) records[name] = r;
    });
    if (records[APP.SHEETS.PERUSAHAAN]) {
      const h = records[APP.SHEETS.PERUSAHAAN].headers;
      const i = _headerIndex_(h, ['Nama Perusahaan','Perusahaan','Company']);
      if (i >= 0) result.companies = [...new Set(records[APP.SHEETS.PERUSAHAAN].rows.map(x => String(x.values[i]||'').trim()).filter(Boolean))].sort();
    }
    if (records[APP.SHEETS.DATA]) {
      const h = records[APP.SHEETS.DATA].headers;
      const it = _headerIndex_(h, ['Nama Training','Training','Jenis Training']);
      const ib = _headerIndex_(h, ['Batch']);
      const ic = _headerIndex_(h, ['Nama Perusahaan','Perusahaan','Company']);
      if (!result.companies.length && ic >= 0) result.companies = [...new Set(records[APP.SHEETS.DATA].rows.map(x=>String(x.values[ic]||'').trim()).filter(Boolean))].sort();
      if (it >= 0) result.trainings = [...new Set(records[APP.SHEETS.DATA].rows.map(x=>String(x.values[it]||'').trim()).filter(Boolean))].sort();
      if (ib >= 0) result.batches = [...new Set(records[APP.SHEETS.DATA].rows.map(x=>String(x.values[ib]||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    }
    return {success:true, ...result};
  } catch(err) {
    return {success:false, message:'Gagal memuat opsi Setting Tugas: '+err.message};
  }
}

function getTaskSettings(username) {
  _authorizePemakai_(username);
  try {
    const sh = _ensureTaskSettingSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {success:true, rows:[]};
    const values = sh.getRange(1,1,lastRow,6).getDisplayValues();
    const rows = values.slice(1).map((r,i) => ({
      rowNumber: i+2, id:r[0]||'', company:r[1]||'', training:r[2]||'', batch:r[3]||'', date:r[4]||'', time:r[5]||''
    })).filter(r=>r.company||r.training||r.batch||r.date||r.time);
    return {success:true, rows};
  } catch(err) { return {success:false,message:'Gagal memuat Setting Tugas: '+err.message}; }
}

function saveTaskSetting(username, payload, rowNumber) {
  _authorizePemakai_(username);
  try {
    const data = payload || {};
    const company = String(data.company||'').trim();
    const training = String(data.training||'').trim();
    const batch = String(data.batch||'').trim();
    const date = String(data.date||'').trim();
    const time = String(data.time||'').trim();
    if (!company || !training || !batch || !date || !time) return {success:false,message:'Nama perusahaan, jenis training, batch, tanggal, dan jam wajib diisi.'};
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return {success:false,message:'Format tanggal tidak valid.'};
    if (!/^\d{2}:\d{2}$/.test(time)) return {success:false,message:'Format jam tidak valid.'};
    const sh = _ensureTaskSettingSheet_();
    const id = rowNumber ? String(sh.getRange(Number(rowNumber),1).getDisplayValue()||'') : '';
    const finalId = id || ('SET-'+Utilities.getUuid().slice(0,8).toUpperCase());
    const row = [finalId,company,training,batch,date,time];
    const target = rowNumber ? Number(rowNumber) : sh.getLastRow()+1;
    if (target < 2) throw new Error('Baris data tidak valid.');
    sh.getRange(target,1,1,6).setValues([row]);
    sh.getRange(target,5).setNumberFormat('yyyy-mm-dd');
    sh.getRange(target,6).setNumberFormat('hh:mm');
    SpreadsheetApp.flush();
    _syncTugasFromSetting_();
    _syncUserPesertaStatusFromDeadline_();
    return {success:true,message:rowNumber?'Setting deadline berhasil diperbarui.':'Setting deadline berhasil ditambahkan.',rowNumber:target,id:finalId};
  } catch(err) { return {success:false,message:'Gagal menyimpan Setting Tugas: '+err.message}; }
}

function deleteTaskSetting(username, rowNumber) {
  _authorizePemakai_(username);
  try {
    const sh = _ensureTaskSettingSheet_();
    const row = Number(rowNumber);
    if (!Number.isInteger(row) || row < 2 || row > sh.getLastRow()) return {success:false,message:'Baris setting tidak valid.'};
    sh.deleteRow(row);
    _syncTugasFromSetting_();
    _syncUserPesertaStatusFromDeadline_();
    return {success:true,message:'Setting deadline berhasil dihapus.'};
  } catch(err) { return {success:false,message:'Gagal menghapus Setting Tugas: '+err.message}; }
}



function _ensureCalendarAgendaSheet_() {
  const ss=_ss_();
  let sh=ss.getSheetByName(APP.SHEETS.KALENDER_AGENDA);
  if(!sh) sh=ss.insertSheet(APP.SHEETS.KALENDER_AGENDA);
  const wanted=['ID','Tanggal Mulai','Tanggal Selesai','Unit Usaha','Jenis Agenda','Status Agenda','Agenda'];
  if(sh.getLastRow()<1){ sh.getRange(1,1,1,7).setValues([wanted]); sh.setFrozenRows(1); sh.getRange(1,1,1,7).setFontWeight('bold'); return sh; }
  const hdr=sh.getRange(1,1,1,Math.max(7,sh.getLastColumn())).getDisplayValues()[0].map(v=>String(v||'').trim());
  const oldSchema=hdr.some(x=>/^Nama Training$/i.test(x))||hdr.some(x=>/^Batch$/i.test(x))||hdr.some(x=>/^Jam$/i.test(x));
  if(oldSchema){
    const last=sh.getLastRow(); const old=last>=2?sh.getRange(2,1,last-1,Math.max(8,sh.getLastColumn())).getValues():[];
    const migrated=old.map(r=>[r[0]||'',r[1]||'',r[2]||r[1]||'',r[3]||'',r[4]||'Training','Progress',r[7]||'']);
    sh.clearContents(); sh.getRange(1,1,1,7).setValues([wanted]); if(migrated.length)sh.getRange(2,1,migrated.length,7).setValues(migrated); sh.setFrozenRows(1);
  }else{ const current=sh.getRange(1,1,1,Math.max(7,sh.getLastColumn())).getDisplayValues()[0]; if(current.slice(0,7).join('|')!==wanted.join('|')) sh.getRange(1,1,1,7).setValues([wanted]); }
  sh.getRange(1,1,1,7).setFontWeight('bold');
  return sh;
}

function getCalendarAgenda(userType, identity) {
  const type=_norm_(userType)==='peserta'?'peserta':'pemakai';
  if(type==='peserta') _authorizePeserta_(identity); else _authorizePemakai_(identity);
  try{
    const sh=_ensureCalendarAgendaSheet_();
    const last=sh.getLastRow();
    const manual=[];
    if(last>=2){
      const vals=sh.getRange(2,1,last-1,7).getDisplayValues();
      vals.forEach((r,i)=>{
        const start=_parseDateKeyLoose_(r[1]), end=_parseDateKeyLoose_(r[2])||start;
        if(!start||!r[3]||!r[4]||!r[6])return;
        manual.push({rowNumber:i+2,id:r[0]||'',startKey:start,endKey:end,company:r[3]||'',training:r[4]||'',jenisAgenda:r[4]||'',status:r[5]||'Progress',agenda:r[6]||'',source:'manual'});
      });
    }
    // Kalender & Agenda hanya mengambil jadwal yang benar-benar dibuat
    // melalui menu "Tambah Jadwal" dan tersimpan di sheet KalenderAgenda.
    // Sheet Data tidak digunakan sebagai sumber agenda otomatis.
    return {success:true,events:manual};
  }catch(err){return {success:false,message:'Gagal memuat Kalender & Agenda: '+err.message};}
}

function _parseDateKeyLoose_(value){
  if(value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value,APP.TIMEZONE,'yyyy-MM-dd');
  const s=String(value||'').trim(); if(!s)return '';
  let m=s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m)return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  m=s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})/);
  if(m)return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  const d=new Date(s); if(isNaN(d.getTime()))return '';
  return Utilities.formatDate(d,APP.TIMEZONE,'yyyy-MM-dd');
}

function getCalendarAgendaSheet(username) {
  _authorizePemakai_(username);
  try{
    const sh=_ensureCalendarAgendaSheet_();
    const last=sh.getLastRow();
    const rows=[];
    if(last>=2){
      const vals=sh.getRange(2,1,last-1,7).getDisplayValues();
      vals.forEach((r,i)=>{
        const start=_parseDateKeyLoose_(r[1]), end=_parseDateKeyLoose_(r[2])||start;
        if(!start||!r[3]||!r[4]||!r[6]) return;
        rows.push({rowNumber:i+2,id:r[0]||'',startKey:start,endKey:end,dateKey:start,company:r[3]||'',jenisAgenda:r[4]||'',status:r[5]||'Progress',agenda:r[6]||'',source:'manual'});
      });
    }
    return {success:true,events:rows};
  }catch(err){return {success:false,message:'Gagal memuat data KalenderAgenda: '+err.message};}
}

function saveCalendarAgenda(username,payload,rowNumber){
  _authorizePemakai_(username);
  try{
    const data=payload||{};
    const start=_parseDateKeyLoose_(data.startKey||data.dateStart||data.date);
    const end=_parseDateKeyLoose_(data.endKey||data.dateEnd||start)||start;
    const company=String(data.company||'').trim();
    const jenisAgendaRaw=String(data.jenisAgenda||data.training||'').trim();
    const statusRaw=String(data.status||'Progress').trim();
    const agenda=String(data.agenda||'').trim();
    const jenisMap={training:'Training',project:'Project',event:'Event'};
    const statusMap={cancel:'Cancel',progress:'Progress',complete:'Complete',reschedule:'Reschedule'};
    const jenisAgenda=jenisMap[jenisAgendaRaw.toLowerCase()]||jenisAgendaRaw;
    const status=statusMap[statusRaw.toLowerCase()]||statusRaw;
    const validJenis=Object.values(jenisMap); const validStatus=Object.values(statusMap);
    if(!start||!company||!jenisAgenda||!agenda)return {success:false,message:'Tanggal mulai, Unit Usaha, Jenis Agenda, dan Agenda wajib diisi.'};
    if(end<start)return {success:false,message:'Tanggal selesai tidak boleh lebih awal dari tanggal mulai.'};
    if(!validJenis.includes(jenisAgenda))return {success:false,message:'Jenis Agenda tidak valid.'};
    if(!validStatus.includes(status))return {success:false,message:'Status Agenda tidak valid.'};
    const sh=_ensureCalendarAgendaSheet_(); const target=rowNumber?Number(rowNumber):sh.getLastRow()+1; if(target<2)throw new Error('Baris tujuan tidak valid.');
    const id=String(rowNumber?sh.getRange(target,1).getDisplayValue():'').trim()||('KAL-'+Utilities.getUuid().slice(0,8).toUpperCase());
    sh.getRange(target,1,1,7).setValues([[id,start,end,company,jenisAgenda,status,agenda]]); sh.getRange(target,2,1,2).setNumberFormat('yyyy-mm-dd'); SpreadsheetApp.flush();
    return {success:true,message:rowNumber?'Jadwal berhasil diperbarui.':'Jadwal berhasil ditambahkan.',rowNumber:target,id};
  }catch(err){return {success:false,message:'Gagal menyimpan jadwal: '+err.message};}
}

function deleteCalendarAgenda(username,rowNumber){
  _authorizePemakai_(username);
  try{
    const sh=_ensureCalendarAgendaSheet_();
    const row=Number(rowNumber);
    if(!Number.isInteger(row)||row<2||row>sh.getLastRow())return {success:false,message:'Baris agenda tidak valid.'};
    sh.deleteRow(row); SpreadsheetApp.flush();
    return {success:true,message:'Agenda berhasil dihapus.'};
  }catch(err){return {success:false,message:'Gagal menghapus agenda: '+err.message};}
}

function _ensureSettingNilaiSheet_() {
  const ss = _ss_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    let sh = ss.getSheetByName(APP.SHEETS.SETTING_NILAI);
    if (!sh) sh = ss.insertSheet(APP.SHEETS.SETTING_NILAI);
    if (sh.getLastRow() < 1) {
      sh.getRange(1,1,1,6).setValues([['ID','Jenis Training','Nama Perusahaan','Batch','Status Posting','Waktu Posting']]);
      sh.setFrozenRows(1);
    } else {
      const header = sh.getRange(1,1,1,Math.max(6,sh.getLastColumn())).getDisplayValues()[0];
      if (_norm_(header[2]) === 'unit usaha') sh.getRange(1,3).setValue('Nama Perusahaan');
    }
    return sh;
  } finally {
    lock.releaseLock();
  }
}

function getSettingNilai(username) {
  _authorizePemakai_(username);
  try {
    const dataRes = _records_(APP.SHEETS.DATA);
    if (!dataRes.success) throw new Error(dataRes.message);
    const h = dataRes.headers;
    // Sumber langsung dari sheet Data: Nama Training + Nama Perusahaan + Batch.
    const iTraining = _headerIndex_(h, ['Nama Training','Jenis Training','Training']);
    const iCompany  = _headerIndex_(h, ['Nama Perusahaan','Perusahaan','Company']);
    const iBatch    = _headerIndex_(h, ['Batch']);
    if (iTraining < 0 || iCompany < 0 || iBatch < 0) {
      throw new Error('Kolom Nama Training, Nama Perusahaan, dan Batch pada sheet Data wajib tersedia.');
    }

    const comboMap = new Map();
    dataRes.rows.forEach(item => {
      const r = item.values || [];
      const training = String(r[iTraining] || '').trim();
      const company  = String(r[iCompany] || '').trim();
      const batch    = String(r[iBatch] || '').trim();
      if (!training || !company || !batch) return;
      const key = [_norm_(training), _norm_(company), _normalizeBatch_(batch)].join('|');
      if (!comboMap.has(key)) comboMap.set(key, {training, company, batch});
    });

    const sh = _ensureSettingNilaiSheet_();
    const last = sh.getLastRow();
    sh.getRange(1,1,1,6).setValues([['ID','Jenis Training','Nama Perusahaan','Batch','Status Posting','Waktu Posting']]);

    const existing = new Map();
    if (last >= 2) {
      sh.getRange(2,1,last-1,6).getDisplayValues().forEach((r,i) => {
        const training=String(r[1]||'').trim(), company=String(r[2]||'').trim(), batch=String(r[3]||'').trim();
        if (!training || !company || !batch) return;
        const key=[_norm_(training),_norm_(company),_normalizeBatch_(batch)].join('|');
        existing.set(key,{rowNumber:i+2,id:r[0]||'',status:r[4]||'Belum Posting',postedAt:r[5]||''});
      });
    }

    const rows=[...comboMap.values()].map(item=>{
      const key=[_norm_(item.training),_norm_(item.company),_normalizeBatch_(item.batch)].join('|');
      const ex=existing.get(key);
      return {rowNumber:ex?.rowNumber||null,id:ex?.id||'',training:item.training,company:item.company,unit:item.company,batch:item.batch,status:ex?.status||'Belum Posting',postedAt:ex?.postedAt||''};
    }).sort((a,b)=>a.training.localeCompare(b.training,'id')||a.company.localeCompare(b.company,'id')||a.batch.localeCompare(b.batch,undefined,{numeric:true}));
    _syncDataStatusFromSettingNilai_();
    return {success:true,rows};
  } catch(err) { return {success:false,message:'Gagal memuat Setting Nilai: '+err.message}; }
}


function postSettingNilai(username, training, company, batch) {
  _authorizePemakai_(username);
  try {
    const result = _setSettingNilaiAndSyncData_(training, company, batch, 'Posted');
    return {success:true,message:'Nilai berhasil diposting.',rowNumber:result.rowNumber,status:'Posted',postedAt:result.postedAt};
  } catch(err) { return {success:false,message:'Gagal posting nilai: '+err.message}; }
}

function cancelSettingNilai(username, training, company, batch) {
  _authorizePemakai_(username);
  try {
    const result = _setSettingNilaiAndSyncData_(training, company, batch, 'Belum Posting');
    return {success:true,message:'Posting nilai berhasil dibatalkan.',rowNumber:result.rowNumber,status:'Belum Posting'};
  } catch(err) { return {success:false,message:'Gagal membatalkan posting nilai: '+err.message}; }
}

function _ensureTaskSheets_() {
  const ss = _ss_();
  let tugas = ss.getSheetByName(APP.SHEETS.TUGAS);
  if (!tugas) {
    tugas = ss.insertSheet(APP.SHEETS.TUGAS);
    tugas.getRange(1,1,1,5).setValues([['ID Tugas','Nama Tugas','Deskripsi','Deadline','Status']]);
  } else if (tugas.getLastRow() < 1) {
    tugas.getRange(1,1,1,5).setValues([['ID Tugas','Nama Tugas','Deskripsi','Deadline','Status']]);
  }
  let peng = ss.getSheetByName(APP.SHEETS.PENGUMPULAN_TUGAS);
  if (!peng) {
    peng = ss.insertSheet(APP.SHEETS.PENGUMPULAN_TUGAS);
    peng.getRange(1,1,1,10).setValues([['ID Pengumpulan','Timestamp','NIK','Nama Peserta','ID Tugas','Nama Tugas','Keterangan','Nama File','URL File','Status']]);
  } else if (peng.getLastRow() < 1) {
    peng.getRange(1,1,1,10).setValues([['ID Pengumpulan','Timestamp','NIK','Nama Peserta','ID Tugas','Nama Tugas','Keterangan','Nama File','URL File','Status']]);
  }
  return {tugas, peng};
}

// Compatibility alias: gunakan nama ini bila deployment lama memanggil daftar tugas peserta.
function loadParticipantTasks(nik) {
  return getParticipantTasks(nik);
}

function _taskDateTime_(dateValue,timeValue){
  if(dateValue instanceof Date && !isNaN(dateValue.getTime())){
    const dKey=Utilities.formatDate(dateValue,APP.TIMEZONE,'yyyy-MM-dd');
    let hh=23,mm=59;
    if(timeValue instanceof Date && !isNaN(timeValue.getTime())){
      const t=Utilities.formatDate(timeValue,APP.TIMEZONE,'HH:mm'); const m=t.match(/^(\d{2}):(\d{2})$/);
      if(m){hh=Number(m[1]);mm=Number(m[2]);}
    }else{
      const ts=String(timeValue??'').trim(); const m=ts.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
      if(m){hh=Number(m[1]);mm=Number(m[2]);}
    }
    const m=dKey.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(!m)return null;
    const y=Number(m[1]),mo=Number(m[2]),d=Number(m[3]);
    if(hh<0||hh>23||mm<0||mm>59)return null;
    const dt=new Date(Date.UTC(y,mo-1,d,hh,mm,0,0)-8*60*60*1000);
    return isNaN(dt.getTime())?null:dt;
  }
  const ds=String(dateValue??'').trim(), ts=String(timeValue??'').trim(); if(!ds)return null;
  let y,m,d,hh=23,mm=59;
  let mt=ds.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::\d{2})?)?$/);
  if(mt){y=Number(mt[1]);m=Number(mt[2]);d=Number(mt[3]);if(mt[4]!=null){hh=Number(mt[4]);mm=Number(mt[5]);}}
  else{
    mt=ds.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::\d{2})?)?$/);
    if(mt){d=Number(mt[1]);m=Number(mt[2]);y=Number(mt[3]);if(mt[4]!=null){hh=Number(mt[4]);mm=Number(mt[5]);}}
    else{
      const date=new Date(ds); if(isNaN(date.getTime()))return null;
      const key=Utilities.formatDate(date,APP.TIMEZONE,'yyyy-MM-dd'); const dm=key.match(/^(\d{4})-(\d{2})-(\d{2})$/); if(!dm)return null;
      y=Number(dm[1]);m=Number(dm[2]);d=Number(dm[3]);
    }
  }
  const tt=ts.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/); if(tt){hh=Number(tt[1]);mm=Number(tt[2]);}
  if(hh<0||hh>23||mm<0||mm>59)return null;
  const dt=new Date(Date.UTC(y,m-1,d,hh,mm,0,0)-8*60*60*1000); return isNaN(dt.getTime())?null:dt;
}

function _isTaskActive_(dateValue, timeValue) {
  const deadline = _taskDateTime_(dateValue, timeValue);
  if (!deadline) return true;
  return new Date().getTime() <= deadline.getTime();
}

function _taskDisplayDeadline_(dateValue, timeValue) {
  const dt = _taskDateTime_(dateValue, timeValue);
  if (!dt) return '';
  return Utilities.formatDate(dt, 'Asia/Makassar', 'dd/MM/yyyy HH:mm');
}

function _syncTugasFromSetting_() {
  const {tugas} = _ensureTaskSheets_();
  const settings = _ensureTaskSettingSheet_();
  const last = settings.getLastRow();

  // Pastikan header sheet Tugas sesuai format master hasil sinkronisasi.
  tugas.getRange(1,1,1,5).setValues([['ID Tugas','Nama Tugas','Deskripsi','Deadline','Status']]);
  tugas.setFrozenRows(1);

  if (last < 2) {
    if (tugas.getLastRow() > 1) tugas.getRange(2,1,tugas.getLastRow()-1,5).clearContent();
    return;
  }

  const vals = settings.getRange(1,1,last,6).getDisplayValues();
  const output = [];
  vals.slice(1).forEach((r, idx) => {
    let id = String(r[0]||'').trim();
    const company = String(r[1]||'').trim();
    const training = String(r[2]||'').trim();
    const batch = String(r[3]||'').trim();
    const date = String(r[4]||'').trim();
    const time = String(r[5]||'').trim();
    if (!id && !company && !training && !batch && !date && !time) return;

    // Pastikan ID tersimpan di SettingTugas. Dengan begitu ID tugas tetap stabil
    // untuk dropdown peserta, download, upload, hapus file, dan submit.
    if (!id) {
      id = 'SET-' + Utilities.getUuid().slice(0,8).toUpperCase();
      settings.getRange(idx + 2, 1).setValue(id);
    }

    const name = [training, company, batch].filter(Boolean).join(' - ');
    const dt = _taskDateTime_(date, time);
    const deadline = dt ? Utilities.formatDate(dt, Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : '';
    const status = _isTaskActive_(date, time) ? 'Aktif' : 'Non Aktif';
    output.push([id, name, '', deadline, status]);
  });

  const existing = Math.max(0, tugas.getLastRow()-1);
  if (existing > 0) tugas.getRange(2,1,existing,5).clearContent();
  if (output.length) tugas.getRange(2,1,output.length,5).setValues(output);
}

function _readTaskRecords_() {
  const {tugas} = _ensureTaskSheets_();
  const lastRow = tugas.getLastRow();
  const lastCol = Math.max(5, tugas.getLastColumn());
  if (lastRow < 2) return {headers: tugas.getRange(1,1,1,lastCol).getDisplayValues()[0], rows: []};
  const vals = tugas.getRange(1,1,lastRow,lastCol).getDisplayValues();
  return {headers: vals[0], rows: vals.slice(1)};
}

function getParticipantTasks(nik) {
  try {
    const participantNik = String(nik == null ? '' : nik).trim();
    if (!participantNik) return {success:true, tasks:[]};

    // Pastikan master peserta dan status terbaru tersinkron tanpa mengubah
    // mekanisme login / password / status yang sudah berjalan.
    try { _syncUserPesertaFromData_(); } catch (e) {}

    // Sumber kecocokan peserta adalah DATA training: NIK + Perusahaan + Training + Batch.
    // Ini memungkinkan satu peserta melihat beberapa tugas sesuai seluruh training
    // yang memang pernah diikuti, bukan seluruh isi sheet Tugas.
    const participantCombos = new Set();
    const dataSources = [APP.SHEETS.DATA, APP.SHEETS.PESERTA];

    dataSources.forEach(sheetName => {
      const result = _records_(sheetName);
      if (!result.success) return;
      const h = result.headers || [];
      const iNik = _headerIndex_(h, ['NIK']);
      const iCompany = _headerIndex_(h, ['Nama Perusahaan','Perusahaan','Company']);
      const iTraining = _headerIndex_(h, ['Jenis Training','Nama Training','Training']);
      const iBatch = _headerIndex_(h, ['Batch']);
      if (iNik < 0 || iCompany < 0 || iTraining < 0 || iBatch < 0) return;

      (result.rows || []).forEach(item => {
        const r = item.values || [];
        if (_norm_(r[iNik]) !== _norm_(participantNik)) return;
        const company = String(r[iCompany] || '').trim();
        const training = String(r[iTraining] || '').trim();
        const batch = String(r[iBatch] || '').trim();
        if (!company || !training || !batch) return;
        participantCombos.add([
          _norm_(training),
          _norm_(company),
          _normalizeBatch_(batch)
        ].join('|'));
      });
    });

    if (!participantCombos.size) return {success:true, tasks:[]};

    // SettingTugas adalah master konfigurasi deadline. Dibaca langsung agar
    // pasangan perusahaan/training/batch tidak hilang ketika disinkronkan ke Tugas.
    const settings = _ensureTaskSettingSheet_();
    const last = settings.getLastRow();
    if (last < 2) return {success:true, tasks:[]};

    const vals = settings.getRange(1,1,last,6).getDisplayValues();
    const headers = vals[0] || [];
    const iId = _headerIndex_(headers, ['ID']);
    const iCompany = _headerIndex_(headers, ['Nama Perusahaan','Perusahaan','Company']);
    const iTraining = _headerIndex_(headers, ['Jenis Training','Nama Training','Training']);
    const iBatch = _headerIndex_(headers, ['Batch']);
    const iDate = _headerIndex_(headers, ['Tanggal','Tanggal Deadline','Deadline']);
    const iTime = _headerIndex_(headers, ['Jam','Jam Deadline','Waktu']);

    const tasks = [];
    const seen = new Set();
    for (let rowIndex = 1; rowIndex < vals.length; rowIndex++) {
      const r = vals[rowIndex] || [];
      const company = iCompany >= 0 ? String(r[iCompany] || '').trim() : '';
      const training = iTraining >= 0 ? String(r[iTraining] || '').trim() : '';
      const batch = iBatch >= 0 ? String(r[iBatch] || '').trim() : '';
      if (!company || !training || !batch) continue;

      const comboKey = [_norm_(training), _norm_(company), _normalizeBatch_(batch)].join('|');
      if (!participantCombos.has(comboKey)) continue;

      // Status tugas tetap sepenuhnya mengikuti deadline yang sudah digunakan
      // aplikasi saat ini. Tugas yang lewat deadline tidak ditampilkan untuk peserta.
      const dateValue = iDate >= 0 ? r[iDate] : '';
      const timeValue = iTime >= 0 ? r[iTime] : '';
      if (!_isTaskActive_(dateValue, timeValue)) continue;

      const configuredId = iId >= 0 ? String(r[iId] || '').trim() : '';
      const id = configuredId || ('SET-' + Utilities.getUuid().slice(0,8).toUpperCase());
      if (seen.has(id)) continue;
      seen.add(id);

      const name = [training, company, batch].filter(Boolean).join(' - ');
      tasks.push({
        id,
        nama: name,
        deskripsi: '',
        deadline: _taskDisplayDeadline_(dateValue, timeValue),
        status: 'Aktif',
        company,
        training,
        batch
      });
    }

    return {success:true, tasks};
  } catch(err) {
    return {success:false,message:'Gagal memuat daftar tugas: '+err.message};
  }
}

function _findParticipantName_(nik) {
  const result=_records_(APP.SHEETS.USER_PESERTA);
  if(!result.success) return '';
  const h=result.headers, iNik=_headerIndex_(h,['NIK']), iNama=_headerIndex_(h,['Nama Peserta','Nama']);
  const row=result.rows.find(x=>iNik>=0 && _norm_(x.values[iNik])===_norm_(nik));
  return row && iNama>=0 ? String(row.values[iNama]||'') : '';
}



/**
 * ===== AKSES GOOGLE DRIVE APLIKASI =====
 * Semua folder dan file yang berada di dalam folder aplikasi dibuat/diatur
 * dengan akses: Anyone + Editor.
 *
 * Root aplikasi saat ini:
 * - Tugas Training
 * - Pengumpulan Tugas
 * - Dokumen Report
 *
 * Catatan: kebijakan Google Workspace/administrator domain dapat membatasi
 * sharing public. Error sharing sengaja ditangani agar proses utama aplikasi
 * tetap berjalan.
 */
function _setDriveAnyoneEditor_(resource) {
  if (!resource) return false;
  try {
    resource.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.EDIT);
    return true;
  } catch (e) {
    console.warn('Gagal mengatur sharing Drive: ' + e.message);
    return false;
  }
}

function _shareDriveTreeRecursive_(folder, visited) {
  if (!folder) return;
  const seen = visited || {};
  const folderId = folder.getId();
  if (seen[folderId]) return;
  seen[folderId] = true;

  _setDriveAnyoneEditor_(folder);

  try {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      _setDriveAnyoneEditor_(file);
    }
  } catch (e) {
    console.warn('Gagal memproses file pada folder ' + folderId + ': ' + e.message);
  }

  try {
    const children = folder.getFolders();
    while (children.hasNext()) {
      _shareDriveTreeRecursive_(children.next(), seen);
    }
  } catch (e) {
    console.warn('Gagal memproses subfolder pada folder ' + folderId + ': ' + e.message);
  }
}

function _ensureDriveAnyoneEditorForRoots_() {
  const rootNames = ['Tugas Training', 'Pengumpulan Tugas', 'Dokumen Report'];
  const visited = {};
  rootNames.forEach(function(name) {
    try {
      const it = DriveApp.getFoldersByName(name);
      while (it.hasNext()) {
        _shareDriveTreeRecursive_(it.next(), visited);
      }
    } catch (e) {
      console.warn('Gagal sinkronisasi folder root ' + name + ': ' + e.message);
    }
  });
  return true;
}

/**
 * Jalankan sekali dari editor Apps Script untuk memperbaiki akses seluruh
 * folder/file lama. Setelah itu folder/file baru juga akan langsung diatur
 * melalui helper di bawah.
 */
function setupAllAppDriveAccess() {
  try {
    _ensureDriveAnyoneEditorForRoots_();
    let migration = null;
    try { migration = migrateReportDocumentsToManagedStorage(); } catch (e) { migration = {success:false,message:e.message}; }
    return {
      success: true,
      migration: migration,
      message: 'Akses folder/file aplikasi telah disetel ke Anyone sebagai Editor dan dokumen report lama telah dipindahkan ke storage yang dikelola akun eksekusi sejauh file dapat dibaca.'
    };
  } catch (e) {
    return {success:false, message:'Gagal mengatur akses Google Drive: ' + e.message};
  }
}

function _ensureTaskDownloadLogSheet_() {
  const ss = _ss_();
  let sh = ss.getSheetByName('LogUnduhTugas');
  if (!sh) sh = ss.insertSheet('LogUnduhTugas');
  if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,5).setValues([['Timestamp','NIK','ID Tugas','Nama Tugas','Folder URL']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _getTugasTrainingFolder_() {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty('TASK_TRAINING_FOLDER_ID');
  let folder = null;
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = null; }
  }
  if (!folder) {
    const it = DriveApp.getFoldersByName('Tugas Training');
    folder = it.hasNext() ? it.next() : DriveApp.createFolder('Tugas Training');
    props.setProperty('TASK_TRAINING_FOLDER_ID', folder.getId());
  }
  _setDriveAnyoneEditor_(folder);
  return folder;
}

function _findTaskById_(taskId) {
  const wantedId = String(taskId || '').trim();
  if (!wantedId) return null;

  // Sumber deadline yang sebenarnya adalah SettingTugas. Jangan memakai
  // nilai Deadline pada sheet Tugas karena nilainya sudah diformat sebagai teks
  // dan dapat mengalami pergeseran interpretasi timezone/locale.
  try {
    const st = _ensureTaskSettingSheet_();
    const sl = st.getLastRow();
    if (sl >= 2) {
      const sv = st.getRange(2,1,sl-1,6).getValues();
      for (const r of sv) {
        const id = String(r[0] || '').trim();
        if (id !== wantedId) continue;
        const company = String(r[1] || '').trim();
        const training = String(r[2] || '').trim();
        const batch = String(r[3] || '').trim();
        const deadlineDate = _taskDateTime_(r[4], r[5]);
        const active = deadlineDate ? Date.now() <= deadlineDate.getTime() : true;
        const name = [training, company, batch].filter(Boolean).join(' - ');
        if (!name) return null;
        return {id, name, active, deadlineDate, company, training, batch};
      }
    }
  } catch(e) {}

  // Fallback kompatibilitas untuk data lama yang mungkin belum memiliki
  // pasangan lengkap di SettingTugas.
  try {
    _syncTugasFromSetting_();
    const rec = _readTaskRecords_();
    const h = rec.headers || [];
    const ii = _headerIndex_(h,['ID Tugas','ID']);
    const iname = _headerIndex_(h,['Nama Tugas','Tugas','Judul']);
    if (ii < 0) return null;
    for (const r of rec.rows || []) {
      const id = String(r[ii] || '').trim();
      if (id !== wantedId) continue;
      const name = iname >= 0 ? String(r[iname] || '').trim() : '';
      if (!name) return null;
      return {id, name, active:true, deadlineDate:null};
    }
  } catch(e) {}
  return null;
}

function getTaskDownloadStatus(nik,sessionToken,taskId){
  try{
    if(!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000))return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir. Silakan login kembali.'};
    const task=_findTaskById_(taskId); if(!task)return {success:false,message:'Tugas tidak ditemukan.'};
    const nowMs=Date.now(),deadlineDate=task.deadlineDate instanceof Date&&!isNaN(task.deadlineDate.getTime())?task.deadlineDate:null;
    const deadlineSecondsRemaining=deadlineDate?Math.max(0,Math.ceil((deadlineDate.getTime()-nowMs)/1000)):0;
    const taskActive=!!task.active&&(!deadlineDate||deadlineSecondsRemaining>0);
    const canDownload=taskActive&&(!deadlineDate||deadlineSecondsRemaining>=2*60*60);
    const sh=_ensureTaskDownloadLogSheet_(); let latest=null; const last=sh.getLastRow();
    if(last>=2){
      const vals=sh.getRange(2,1,last-1,3).getValues();
      for(const r of vals){if(_norm_(r[1])!==_norm_(nik)||String(r[2]||'').trim()!==String(task.id).trim())continue;const ts=r[0] instanceof Date?r[0]:new Date(r[0]);if(!isNaN(ts.getTime())&&(!latest||ts.getTime()>latest.getTime()))latest=ts;}
    }
    const base={success:true,serverNowMs:nowMs,taskActive,canDownload,deadlineSecondsRemaining:taskActive?deadlineSecondsRemaining:0,deadlineAt:deadlineDate?deadlineDate.toISOString():'',folderUrl:'',folderName:''};
    if(!taskActive)return {...base,downloaded:false,canSubmit:false,secondsRemaining:0,downloadSecondsRemaining:0};
    if(!latest)return {...base,downloaded:false,canSubmit:false,secondsRemaining:0,downloadSecondsRemaining:Math.max(0,deadlineSecondsRemaining-2*60*60)};
    const eligibleAt=new Date(latest.getTime()+60*60*1000),waitSeconds=Math.max(0,Math.ceil((eligibleAt.getTime()-nowMs)/1000));
    return {...base,downloaded:true,downloadedAt:Utilities.formatDate(latest,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss'),eligibleAt:Utilities.formatDate(eligibleAt,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss'),canSubmit:waitSeconds<=0&&deadlineSecondsRemaining>0,secondsRemaining:waitSeconds,downloadSecondsRemaining:Math.max(0,deadlineSecondsRemaining-2*60*60)};
  }catch(e){return {success:false,message:'Gagal memeriksa status pengunduhan: '+e.message};}
}

function prepareTaskDownload(nik, sessionToken, taskId){
  try{
    if(!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000)) return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir. Silakan login kembali.'};
    const task=_findTaskById_(taskId); if(!task) return {success:false,message:'Tugas tidak ditemukan.'};
    if(!task.active) return {success:false,message:'Tugas tersebut sudah Non Aktif karena deadline telah terlampaui.'};
    if(task.deadlineDate){ const remaining=Math.max(0,Math.ceil((task.deadlineDate.getTime()-Date.now())/1000)); if(remaining<2*60*60)return {success:false,message:'Download tugas otomatis ditutup karena sisa deadline kurang dari 2 jam.',deadlineSecondsRemaining:remaining,canDownload:false}; }
    const info=_ensureTaskFolderBySetting_(task.id); const it=info.folder.getFiles(); const files=[]; while(it.hasNext()){ const f=it.next(); if(!f.isTrashed())files.push(f); }
    if(!files.length)return {success:false,message:'Folder tugas belum memiliki file soal.'};
    let blob,name,mime;
    if(files.length===1){ blob=files[0].getBlob(); name=files[0].getName(); mime=files[0].getMimeType()||'application/octet-stream'; }
    else { blob=Utilities.zip(files.map(f=>f.getBlob().setName(f.getName())),info.folderName+' - Soal.zip'); name=info.folderName+' - Soal.zip'; mime='application/zip'; }
    const bytes=blob.getBytes();
    return {success:true,fileName:name,mimeType:mime,data:Utilities.base64Encode(bytes),folderName:info.folderName};
  }catch(e){return {success:false,message:'Gagal menyiapkan file download: '+e.message};}
}
function recordTaskDownload(nik,sessionToken,taskId){
  try{
    if(!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000))return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir.'};
    const task=_findTaskById_(taskId);if(!task)return {success:false,message:'Tugas tidak ditemukan.'};
    const sh=_ensureTaskDownloadLogSheet_(),now=new Date();
    sh.getRange(sh.getLastRow()+1,1,1,5).setValues([[now,String(nik),task.id,task.name,'']]);
    const eligibleAt=new Date(now.getTime()+60*60*1000);
    return {success:true,downloaded:true,downloadedAt:Utilities.formatDate(now,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss'),eligibleAt:Utilities.formatDate(eligibleAt,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss'),secondsRemaining:3600,serverNowMs:now.getTime()};
  }catch(e){return {success:false,message:'Gagal mencatat download tugas: '+e.message};}
}

function registerTaskDownload(nik, sessionToken, taskId) {
  try {
    if (!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000)) {
      return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir. Silakan login kembali.'};
    }
    const task = _findTaskById_(taskId);
    if (!task) return {success:false,message:'Tugas tidak ditemukan.'};
    if (!task.active) return {success:false,message:'Tugas tersebut sudah Non Aktif karena deadline telah terlampaui.'};
    if (task.deadlineDate) {
      const remaining = Math.max(0, Math.ceil((task.deadlineDate.getTime()-new Date().getTime())/1000));
      if (remaining < 2*60*60) return {success:false,message:'Download tugas otomatis ditutup karena sisa deadline kurang dari 2 jam.',deadlineSecondsRemaining:remaining,canDownload:false};
    }
    const taskFolderInfo = _ensureTaskFolderBySetting_(task.id);
    const folder = taskFolderInfo.folder;
    const sh = _ensureTaskDownloadLogSheet_();
    const now = new Date();
    sh.getRange(sh.getLastRow()+1,1,1,5).setValues([[now,String(nik),task.id,task.name,folder.getUrl()]]);
    SpreadsheetApp.flush();
    const eligibleAt = new Date(now.getTime()+60*60*1000);
    return {
      success:true,downloaded:true,
      downloadedAt:Utilities.formatDate(now,'Asia/Makassar','dd/MM/yyyy HH:mm:ss'),
      eligibleAt:Utilities.formatDate(eligibleAt,'Asia/Makassar','dd/MM/yyyy HH:mm:ss'),
      canSubmit:false,
      canDownload:task.deadlineDate ? Math.max(0, Math.ceil((task.deadlineDate.getTime()-now.getTime())/1000)) >= 2*60*60 : true,
      secondsRemaining:3600,
      deadlineSecondsRemaining:task.deadlineDate ? Math.max(0, Math.ceil((task.deadlineDate.getTime()-now.getTime())/1000)) : 0,
      folderUrl:folder.getUrl(),
      folderName:taskFolderInfo.folderName
    };
  } catch(err) {
    return {success:false,message:'Gagal mencatat pengunduhan tugas: '+err.message};
  }
}

function getParticipantTaskSubmission(nik, sessionToken, taskId) {
  try {
    if(!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000)) {
      return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir. Silakan login kembali.'};
    }
    const sh=SpreadsheetApp.getActive().getSheetByName(APP.SHEETS.PENGUMPULAN_TUGAS || 'PengumpulanTugas');
    if(!sh || sh.getLastRow()<2) return {success:true,exists:false};
    const data=sh.getRange(1,1,sh.getLastRow(),Math.max(10,sh.getLastColumn())).getDisplayValues();
    const h=data[0];
    const iN=_headerIndex_(h,['NIK']);
    const iTask=_headerIndex_(h,['ID Tugas','ID']);
    const iTime=_headerIndex_(h,['Timestamp','Waktu']);
    const iName=_headerIndex_(h,['Nama File','File']);
    const iUrl=_headerIndex_(h,['URL File','URL']);
    const iStatus=_headerIndex_(h,['Status']);
    let latest=null;
    for(let i=1;i<data.length;i++){
      const r=data[i];
      if(iN>=0 && _norm_(r[iN])!==_norm_(nik)) continue;
      if(iTask>=0 && String(r[iTask]||'').trim()!==String(taskId||'').trim()) continue;
      const ts=iTime>=0 ? new Date(r[iTime]) : new Date(0);
      if(!latest || ts.getTime()>=latest.ts.getTime()) latest={ts,row:r};
    }
    if(!latest) return {success:true,exists:false};
    const r=latest.row;
    return {
      success:true,exists:true,
      fileName:iName>=0?String(r[iName]||''):'',
      url:iUrl>=0?String(r[iUrl]||''):'',
      status:iStatus>=0?String(r[iStatus]||'Dikirim'):'Dikirim',
      submittedAt:iTime>=0?String(r[iTime]||''):''
    };
  } catch(err) {
    return {success:false,message:'Gagal memuat file tugas yang sudah diupload: '+err.message};
  }
}

function deleteParticipantTaskSubmission(nik, sessionToken, taskId) {
  try {
    if(!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000)) {
      return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir. Silakan login kembali.'};
    }
    const sh=SpreadsheetApp.getActive().getSheetByName((APP.SHEETS && APP.SHEETS.PENGUMPULAN_TUGAS) || 'PengumpulanTugas');
    if(!sh || sh.getLastRow()<2) return {success:false,message:'File tugas yang akan dihapus tidak ditemukan.'};
    const data=sh.getRange(1,1,sh.getLastRow(),sh.getLastColumn()).getDisplayValues();
    const h=data[0];
    const iN=_headerIndex_(h,['NIK']);
    const iTask=_headerIndex_(h,['ID Tugas','ID']);
    const iTime=_headerIndex_(h,['Timestamp','Waktu','Tanggal']);
    const iUrl=_headerIndex_(h,['URL File','URL']);
    if(iN<0 || iTask<0) return {success:false,message:'Kolom NIK atau ID Tugas tidak ditemukan pada sheet PengumpulanTugas.'};
    let latestIndex=-1, latestMs=-1;
    for(let i=1;i<data.length;i++){
      const r=data[i];
      if(_norm_(r[iN])!==_norm_(nik)) continue;
      if(String(r[iTask]||'').trim()!==String(taskId||'').trim()) continue;
      let ms=0;
      if(iTime>=0){ const d=new Date(r[iTime]); ms=isNaN(d.getTime())?0:d.getTime(); }
      if(latestIndex<0 || ms>=latestMs){ latestIndex=i; latestMs=ms; }
    }
    if(latestIndex<0) return {success:false,message:'File tugas yang akan dihapus tidak ditemukan.'};
    const row=data[latestIndex];
    const url=iUrl>=0?String(row[iUrl]||'').trim():'';
    let trashed=false;
    const m=url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if(m && m[1]){ try{const f=DriveApp.getFileById(m[1]);f.setTrashed(true);_logTaskDriveDelete_(String(nik||''),f.getId(),f.getName(),'File','peserta');trashed=true;}catch(e){} }
    sh.deleteRow(latestIndex+1);
    return {success:true,trashed,message:trashed?'File tugas dan catatan pengumpulan berhasil dihapus.':'Catatan pengumpulan berhasil dihapus. File Drive tidak ditemukan atau sudah terhapus.'};
  } catch(err) { return {success:false,message:'Gagal menghapus file tugas: '+err.message}; }
}

function submitParticipantTask(nik, sessionToken, taskId, note, fileName, dataUrl, mimeType, selectedTaskNameFromClient) {
  try {
    if(!_verifySessionToken_(sessionToken,'peserta',nik,12*60*60*1000)) return {success:false,message:'Sesi login peserta tidak valid atau sudah berakhir. Silakan login kembali.'};
    const name=_findParticipantName_(nik);
    if(!name) return {success:false,message:'Peserta tidak ditemukan.'};
    const {tugas,peng}=_ensureTaskSheets_();
    // Validasi submit harus membaca deadline langsung dari SettingTugas agar
    // tanggal/jam menggunakan nilai mentah dan timezone WITA yang konsisten.
    const task = _findTaskById_(taskId);
    if(!task) return {success:false,message:'Tugas tidak ditemukan.'};
    if(!task.active) return {success:false,message:'Tugas tersebut sudah melewati deadline dan tidak dapat dikumpulkan.'};
    // Peserta wajib menunggu minimal 1 jam sejak unduhan soal tercatat.
    const dl = getTaskDownloadStatus(nik, sessionToken, task.id);
    if(!dl || !dl.success) return {success:false,message:dl?.message||'Status pengunduhan soal tidak dapat diverifikasi.'};
    if(!dl.taskActive) return {success:false,message:'Tugas sudah Non Aktif karena deadline telah terlampaui.'};
    if(!dl.downloaded) return {success:false,message:'Soal belum diunduh. Download soal terlebih dahulu.'};
    if(!dl.canSubmit){
      if(Number(dl.deadlineSecondsRemaining||0)<=0) return {success:false,message:'Deadline pengiriman tugas sudah berakhir.'};
      return {success:false,message:'Tugas baru dapat dikirim setelah 1 jam sejak soal diunduh. Waktu yang tersisa: '+Math.ceil(Number(dl.secondsRemaining||0)/60)+' menit.'};
    }
    // Satu file submission aktif per peserta + tugas. Hapus submission lama terlebih dahulu.
    const existing=getParticipantTaskSubmission(nik, sessionToken, task.id);
    if(existing && existing.success && existing.exists){
      return {success:false,message:'File tugas untuk tugas ini sudah pernah dikirim. Hapus file yang sudah dikirim terlebih dahulu sebelum mengunggah file baru.'};
    }
    const nm=String(fileName||'').trim();
    const raw=String(dataUrl||'');
    if(!nm || !raw.startsWith('data:')) return {success:false,message:'File tugas tidak valid.'};
    const parts=raw.split(',');
    if(parts.length<2) return {success:false,message:'Data file tidak valid.'};
    const bytes=Utilities.base64Decode(parts[1]);
    if(bytes.length>10*1024*1024) return {success:false,message:'Ukuran file maksimal 10 MB.'};
    const blob=Utilities.newBlob(bytes, mimeType || 'application/octet-stream', nm);
    const props=PropertiesService.getScriptProperties();
    let rootFolderId=props.getProperty('TASK_SUBMISSION_FOLDER_ID');
    let rootFolder=null;
    if(rootFolderId){ try { rootFolder=DriveApp.getFolderById(rootFolderId); } catch(e) { rootFolder=null; } }
    if(!rootFolder){
      const rootName='Pengumpulan Tugas';
      const it=DriveApp.getFoldersByName(rootName);
      rootFolder=it.hasNext()?it.next():DriveApp.createFolder(rootName);
      props.setProperty('TASK_SUBMISSION_FOLDER_ID',rootFolder.getId());
    }
    _setDriveAnyoneEditor_(rootFolder);
    // Folder hasil pengiriman: [Nama Training] – [Nama Perusahaan] – [Batch].
    // Contoh: Coba – PT ABC – Batch 1
    const sanitizeName = value => String(value || '')
      .replace(/[\\/:*?"<>|#%{}]/g,'_')
      .replace(/\s+/g,' ')
      .trim()
      .slice(0,120) || 'Tugas';

    const trainingName = sanitizeName(task.training || task.name || 'Training');
    const companyName = sanitizeName(task.company || 'Nama Perusahaan');
    let batchName = sanitizeName(String(task.batch || '').trim() || 'Batch');
    const batchLabel = /^batch\b/i.test(batchName) ? batchName : `Batch ${batchName}`;

    const taskFolderName = sanitizeName(`${trainingName} – ${companyName} – ${batchLabel}`);
    const folders = rootFolder.getFoldersByName(taskFolderName);
    const taskFolder = folders.hasNext() ? folders.next() : rootFolder.createFolder(taskFolderName);
    _setDriveAnyoneEditor_(taskFolder);

    // Nama file: [Nama Peserta] – Tugas Training [Training] – [Nama Perusahaan] – [Batch].ext
    // Contoh: Jarwo – Tugas Training Coba – PT ABC – Batch 1.pdf
    const participantName = sanitizeName(name);
    const originalName = nm.replace(/[\\/:*?"<>|#%{}]/g,'_').trim();
    const dot = originalName.lastIndexOf('.');
    const extension = (dot>0 && dot<originalName.length-1) ? originalName.slice(dot) : '';
    const finalBase = sanitizeName(`${participantName} – Tugas Training ${trainingName} – ${companyName} – ${batchLabel}`).slice(0,180);
    const finalName = finalBase + extension;
    blob.setName(finalName);
    const file = taskFolder.createFile(blob);
    _setDriveAnyoneEditor_(file);
    const id='SUB-'+Utilities.getUuid().slice(0,8).toUpperCase();
    const selectedTaskName = String(selectedTaskNameFromClient || task.name || task.nama || [task.training, task.company, task.batch].filter(Boolean).join(' - ') || 'Tugas Training').trim();
    const row=[id,new Date(),String(nik||''),name,task.id,selectedTaskName,String(note||''),finalName,file.getUrl(),'Dikirim'];
    const nextRow=peng.getLastRow()+1;
    peng.getRange(nextRow,1,1,row.length).setValues([row]);
    return {success:true,message:'Tugas berhasil dikumpulkan.',url:file.getUrl(),submissionId:id};
  } catch(err){ return {success:false,message:'Gagal menyimpan tugas: '+err.message}; }
}


/**
 * Pantauan pengiriman tugas untuk Pemakai.
 * Sumber:
 * - SettingTugas: seluruh setting, baik yang masih aktif maupun sudah Non Aktif berdasarkan deadline.
 * - Data: peserta yang terdaftar pada kombinasi perusahaan + training + batch.
 * - PengumpulanTugas: menentukan sudah/belum mengirim berdasarkan ID Tugas.
 */


/**
 * ===== SOAL TEST • CUSTOM EXAM =====
 * XLSX -> Master Soal -> Ujian custom di dalam aplikasi.
 * Tidak lagi menggunakan Google Form.
 */
function _ensureSoalTestSheet_() {
  const ss=_ss_();
  let sh=ss.getSheetByName(APP.SHEETS.SOAL_TEST);
  const headers=[
    'ID','Tanggal Dibuat','Nama Perusahaan','Nama Agenda','Jenis Soal','Batch','Durasi (Menit)',
    'Nama File XLSX','Jumlah Soal','Total Poin','Nama Soal Test','URL Soal Test','URL Admin Soal Test',
    'URL File XLSX','Status Posting','Waktu Posting','Dibuat Oleh'
  ];
  if(sh && sh.getLastRow()>0)return sh;
  if(!sh)sh=ss.insertSheet(APP.SHEETS.SOAL_TEST);
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  return sh;
}

function _ensureSoalTestDetailSheet_(){
  const ss=_ss_(); const name='SoalTestDetail';
  let sh=ss.getSheetByName(name); if(sh)return sh;
  sh=ss.insertSheet(name);
  const headers=['Test ID','No','Soal','A','B','C','D','E','Kunci Jawaban','Poin'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  return sh;
}

function _ensureSoalTestResponseSheets_(){
  const ss=_ss_();
  let resp=ss.getSheetByName(APP.SHEETS.SOAL_TEST_RESPONSES);
  let ans=ss.getSheetByName(APP.SHEETS.SOAL_TEST_ANSWERS);
  const lockNeeded=!resp||!ans;
  if(lockNeeded){
    const lock=LockService.getScriptLock();
    lock.waitLock(10000);
    try{
      resp=ss.getSheetByName(APP.SHEETS.SOAL_TEST_RESPONSES); if(!resp)resp=ss.insertSheet(APP.SHEETS.SOAL_TEST_RESPONSES);
      ans=ss.getSheetByName(APP.SHEETS.SOAL_TEST_ANSWERS); if(!ans)ans=ss.insertSheet(APP.SHEETS.SOAL_TEST_ANSWERS);
      const headers=['Attempt ID','Soal Test ID','NIK','Nama Peserta','Nama Perusahaan','Batch','Jenis Soal','Nama Agenda','Mulai','Terakhir Autosave','Selesai','Status','Nilai','Total Poin','Persentase','Departemen','Jabatan','NIK Akun (Internal)','Status Posting','Keterangan','Answer Start Row','Answer JSON'];
      if(resp.getLastRow()<1){resp.getRange(1,1,1,headers.length).setValues([headers]);resp.setFrozenRows(1);}
      else if(resp.getLastColumn()<headers.length){resp.getRange(1,1,1,headers.length).setValues([headers]);resp.setFrozenRows(1);}
      if(ans.getLastRow()<1){ans.getRange(1,1,1,7).setValues([['Attempt ID','Soal Test ID','NIK','No Soal','Jawaban','Poin Diperoleh','Waktu Simpan']]);ans.setFrozenRows(1);}
    }finally{try{lock.releaseLock();}catch(e){}}
  }
  return {resp,ans};
}

function _ensureSoalTestAttemptIndexSheet_(){
  const ss=_ss_(); let sh=ss.getSheetByName('SoalTestAttemptIndex');
  if(sh&&sh.getLastRow()>0)return sh;
  const lock=LockService.getScriptLock(); lock.waitLock(10000);
  try{
    sh=ss.getSheetByName('SoalTestAttemptIndex');
    if(!sh)sh=ss.insertSheet('SoalTestAttemptIndex');
    if(sh.getLastRow()<1){
      sh.getRange(1,1,1,7).setValues([['Attempt Key','Attempt ID','Response Row','Soal Test ID','NIK','Status','Updated At']]);
      sh.setFrozenRows(1);sh.getRange(1,1,1,7).setFontWeight('bold');
    }
    return sh;
  }finally{try{lock.releaseLock();}catch(e){}}
}
function _soalTestAttemptKey_(testId,nik){return 'TMS_ST_IDX_'+_soalTestTextKey_(testId)+'_'+_soalTestTextKey_(nik);}

/**
 * Upgrade/rebuild the attempt index in one bulk pass. Older code scanned the full
 * SoalTestResponses sheet for every participant whose index key was missing.
 * That repeated scan becomes expensive exactly when a cohort starts together.
 * Run `_prepareSoalTestConcurrency_` from the Apps Script editor before an exam;
 * startSoalTestAttempt retains this one-time migration as a safe fallback.
 */
function _rebuildSoalTestAttemptIndexFromResponses_(resp,indexSheet,force){
  const props=PropertiesService.getScriptProperties();
  const flag='TMS_SOAL_TEST_ATTEMPT_INDEX_MIGRATED_V2';
  if(!force&&props.getProperty(flag)==='1')return {migrated:false,indexed:Math.max(0,(indexSheet.getLastRow()||1)-1)};
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    if(!force&&props.getProperty(flag)==='1')return {migrated:false,indexed:Math.max(0,(indexSheet.getLastRow()||1)-1)};
    const last=resp.getLastRow(),rows=last>=2?resp.getRange(2,1,last-1,18).getValues():[];
    const chosen=new Map(),allAttemptIds=[];
    rows.forEach((v,i)=>{
      const attemptId=String(v[0]||'').trim(),testId=String(v[1]||'').trim(),nik=String(v[17]||v[2]||'').trim();
      if(!attemptId||!testId||!nik)return;
      allAttemptIds.push(attemptId);
      const key=_soalTestAttemptKey_(testId,nik),status=String(v[11]||'').trim();
      const previous=chosen.get(key);
      // Preserve historical behavior: an in-progress attempt takes precedence over
      // submitted attempts for the same test/NIK; among in-progress attempts use newest.
      if(!previous || status.toUpperCase()!=='SUBMITTED'){
        chosen.set(key,{key,attemptId,responseRow:i+2,testId,nik,status,updatedAt:new Date()});
      }
    });
    const indexLast=indexSheet.getLastRow();
    if(indexLast>1)indexSheet.getRange(2,1,indexLast-1,7).clearContent();
    const entries=Array.from(chosen.values());
    const matrix=entries.map(e=>[e.key,e.attemptId,e.responseRow,e.testId,e.nik,e.status,e.updatedAt]);
    if(matrix.length)indexSheet.getRange(2,1,matrix.length,7).setValues(matrix);
    const cache=CacheService.getScriptCache();
    allAttemptIds.forEach(id=>{try{cache.remove('TMS_ST_ATT_'+id);cache.remove('TMS_ST_META_'+id);}catch(e){}});
    entries.forEach((e,i)=>{
      const entry={row:i+2,attemptId:e.attemptId,responseRow:e.responseRow,testId:e.testId,nik:e.nik,status:e.status,updatedAt:e.updatedAt.toISOString()};
      try{cache.remove(e.key);}catch(err){}
      _cacheJsonPut_(e.key,entry,SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
      _cacheJsonPut_('TMS_ST_ATT_'+e.attemptId,{row:e.responseRow},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
    });
    props.setProperty(flag,'1');
    return {migrated:true,indexed:entries.length};
  }finally{try{lock.releaseLock();}catch(e){}}
}
function _ensureSoalTestAttemptIndexMigrated_(resp,indexSheet){
  return _rebuildSoalTestAttemptIndexFromResponses_(resp,indexSheet,false);
}
function _soalTestAttemptIndexFind_(testId,nik,forceLookup){
  const key=_soalTestAttemptKey_(testId,nik),cached=_cacheJsonGet_(key);
  if(cached&&Number(cached.row)>1)return cached;
  if(cached&&cached.notFound&&!forceLookup)return null;
  const sh=_ss_().getSheetByName('SoalTestAttemptIndex');
  if(!sh||sh.getLastRow()<2){_cacheJsonPut_(key,{row:0,notFound:true},8);return null;}
  const found=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(key).matchEntireCell(true).findNext();
  if(!found){_cacheJsonPut_(key,{row:0,notFound:true},8);return null;}
  const v=sh.getRange(found.getRow(),1,1,7).getValues()[0];
  const entry={row:found.getRow(),attemptId:String(v[1]||''),responseRow:Number(v[2]||0),testId:String(v[3]||''),nik:String(v[4]||''),status:String(v[5]||''),updatedAt:v[6] instanceof Date?v[6].toISOString():String(v[6]||'')};
  _cacheJsonPut_(key,entry,SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
  if(entry.responseRow>1)_cacheJsonPut_('TMS_ST_ATT_'+entry.attemptId,{row:entry.responseRow},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
  return entry;
}
function _soalTestAttemptIndexPut_(testId,nik,attemptId,responseRow,status){
  const sh=_ensureSoalTestAttemptIndexSheet_(),key=_soalTestAttemptKey_(testId,nik),now=new Date();
  const cached=_cacheJsonGet_(key);
  let row=cached&&Number(cached.row)>1?Number(cached.row):-1;
  // The cache row is authoritative for the current index version; rebuild/delete paths
  // refresh it. Fall back to TextFinder only on cold cache to avoid repeated full scans
  // when many participants submit at the same time.
  if(row<2&&sh.getLastRow()>=2){const found=sh.getRange(2,1,sh.getLastRow()-1,1).createTextFinder(key).matchEntireCell(true).findNext();if(found)row=found.getRow();}
  const values=[[key,String(attemptId||''),Number(responseRow||0),String(testId||''),String(nik||''),String(status||''),now]];
  if(row>1)sh.getRange(row,1,1,7).setValues(values);else{row=sh.getLastRow()+1;sh.getRange(row,1,1,7).setValues(values);}
  const entry={row,attemptId:String(attemptId||''),responseRow:Number(responseRow||0),testId:String(testId||''),nik:String(nik||''),status:String(status||''),updatedAt:now.toISOString()};
  _cacheJsonPut_(key,entry,SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
  _cacheJsonPut_('TMS_ST_ATT_'+String(attemptId||''),{row:Number(responseRow||0)},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
  return entry;
}
function _soalTestAttemptIndexAppend_(testId,nik,attemptId,responseRow,status){
  // Called inside _createSoalTestAttempt_'s script lock, so row allocation is deterministic
  // and duplicate concurrent starts cannot point their cache/index at the wrong attempt.
  const sh=_ensureSoalTestAttemptIndexSheet_(),key=_soalTestAttemptKey_(testId,nik),now=new Date();
  const row=sh.getLastRow()+1;
  sh.getRange(row,1,1,7).setValues([[key,String(attemptId||''),Number(responseRow||0),String(testId||''),String(nik||''),String(status||''),now]]);
  const entry={row,attemptId:String(attemptId||''),responseRow:Number(responseRow||0),testId:String(testId||''),nik:String(nik||''),status:String(status||''),updatedAt:now.toISOString()};
  _cacheJsonPut_(key,entry,SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
  _cacheJsonPut_('TMS_ST_ATT_'+String(attemptId||''),{row:Number(responseRow||0)},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
  return entry;
}
function _soalTestAttemptIndexUpdateStatus_(testId,nik,attemptId,responseRow,status){try{_soalTestAttemptIndexPut_(testId,nik,attemptId,responseRow,status);}catch(e){}}

/* ===== SOAL TEST PERFORMANCE CACHE ===== */
const SOAL_TEST_CACHE_TTL_SEC_=120;
const SOAL_TEST_QUESTION_CACHE_TTL_SEC_=1800;
const SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_=43200;
const SOAL_TEST_WAIT_CACHE_MS_=8000;
function _cacheJsonGet_(key){try{const raw=CacheService.getScriptCache().get(key);return raw?JSON.parse(raw):null;}catch(e){return null;}}
function _cacheJsonPut_(key,value,ttl){try{const raw=JSON.stringify(value);if(raw.length<95000)CacheService.getScriptCache().put(key,raw,ttl||60);}catch(e){}}
function _soalTestQuestionCacheGet_(key){
  const direct=_cacheJsonGet_(key);
  if(Array.isArray(direct))return direct;
  const meta=_cacheJsonGet_(key+'_META');
  if(!meta||!Number(meta.chunks))return null;
  try{
    const cache=CacheService.getScriptCache(),parts=[];
    for(let i=0;i<Number(meta.chunks);i++){
      const piece=cache.get(key+'_C'+i);
      if(!piece)return null;
      parts.push(JSON.parse(piece));
    }
    const parsed=JSON.parse(parts.join(''));
    return Array.isArray(parsed)?parsed:null;
  }catch(e){return null;}
}
function _soalTestQuestionCachePut_(key,questions,ttl){
  const cache=CacheService.getScriptCache(),raw=JSON.stringify(questions),life=ttl||SOAL_TEST_QUESTION_CACHE_TTL_SEC_;
  try{
    _soalTestQuestionCacheRemove_(key);
    if(raw.length<40000){cache.put(key,raw,life);return true;}
    // Chunk unusually large exams so each CacheService entry stays comfortably under 100 KB,
    // including JSON escaping and multi-byte question text.
    const chunkSize=15000,chunks=Math.ceil(raw.length/chunkSize);
    for(let i=0;i<chunks;i++)cache.put(key+'_C'+i,JSON.stringify(raw.slice(i*chunkSize,(i+1)*chunkSize)),life);
    cache.put(key+'_META',JSON.stringify({chunks}),life);
    return true;
  }catch(e){return false;}
}
function _soalTestQuestionCacheRemove_(key){
  try{
    const cache=CacheService.getScriptCache(),meta=_cacheJsonGet_(key+'_META');
    const keys=[key,key+'_META'];
    if(meta&&Number(meta.chunks)>0)for(let i=0;i<Number(meta.chunks);i++)keys.push(key+'_C'+i);
    cache.removeAll(keys);
  }catch(e){}
}
function _cachePollJson_(key,waitMs){
  const end=Date.now()+Math.max(0,Number(waitMs)||0);
  while(Date.now()<end){
    const hit=_cacheJsonGet_(key);
    if(hit!==null)return hit;
    Utilities.sleep(100);
  }
  return null;
}
function _invalidateSoalTestCaches_(testId){
  try{
    const c=CacheService.getScriptCache();
    c.removeAll(['TMS_ST_HISTORY','TMS_ST_HAS_POSTED']);
    if(testId){const id=String(testId).trim();_soalTestQuestionCacheRemove_('TMS_ST_Q_'+id);c.remove('TMS_ST_MASTER_'+id);}
  }catch(e){}
}
function _getCachedSoalTestHistoryRows_(){
  const key='TMS_ST_HISTORY',cached=_cacheJsonGet_(key);
  if(Array.isArray(cached))return cached;
  const lock=LockService.getScriptLock();
  if(lock.tryLock(150)){
    try{
      const again=_cacheJsonGet_(key);
      if(Array.isArray(again))return again;
      const rows=_getSoalTestHistoryRowsRaw_();
      _cacheJsonPut_(key,rows,SOAL_TEST_CACHE_TTL_SEC_);
      return rows;
    }finally{try{lock.releaseLock();}catch(e){}}
  }
  const polled=_cachePollJson_(key,SOAL_TEST_WAIT_CACHE_MS_);
  if(Array.isArray(polled))return polled;
  return _getSoalTestHistoryRowsRaw_();
}
function _getCachedSoalTestQuestions_(testId){
  const id=String(testId||'').trim(); if(!id)return [];
  const key='TMS_ST_Q_'+id,cached=_soalTestQuestionCacheGet_(key);
  if(Array.isArray(cached))return cached;
  const lock=LockService.getScriptLock();
  if(lock.tryLock(150)){
    try{
      const again=_soalTestQuestionCacheGet_(key); if(Array.isArray(again))return again;
      const ds=_ensureSoalTestDetailSheet_(),last=ds.getLastRow(),out=[];
      if(last>=2)ds.getRange(2,1,last-1,10).getValues().forEach(v=>{
        if(String(v[0])!==id)return;
        out.push({no:String(v[1]||''),question:String(v[2]||''),options:{A:String(v[3]||''),B:String(v[4]||''),C:String(v[5]||''),D:String(v[6]||''),E:String(v[7]||'')},point:Number(v[9]||0),key:String(v[8]||'').trim().toUpperCase()});
      });
      _soalTestQuestionCachePut_(key,out,SOAL_TEST_QUESTION_CACHE_TTL_SEC_);
      return out;
    }finally{try{lock.releaseLock();}catch(e){}}
  }
  const polled=_soalTestQuestionCacheGet_(key);
  if(Array.isArray(polled))return polled;
  const ds=_ensureSoalTestDetailSheet_(),last=ds.getLastRow(),out=[];
  if(last>=2)ds.getRange(2,1,last-1,10).getValues().forEach(v=>{if(String(v[0])!==id)return;out.push({no:String(v[1]||''),question:String(v[2]||''),options:{A:String(v[3]||''),B:String(v[4]||''),C:String(v[5]||''),D:String(v[6]||''),E:String(v[7]||'')},point:Number(v[9]||0),key:String(v[8]||'').trim().toUpperCase()});});
  _soalTestQuestionCachePut_(key,out,SOAL_TEST_QUESTION_CACHE_TTL_SEC_);
  return out;
}
function _getSoalTestHistoryRowsRaw_(){const sh=_ensureSoalTestSheet_(),last=sh.getLastRow();if(last<2)return [];return sh.getRange(2,1,last-1,17).getDisplayValues().map(r=>({id:r[0]||'',createdAt:r[1]||'',company:r[2]||'',agenda:r[3]||'',type:r[4]||'',batch:r[5]||'',duration:r[6]||'',fileName:r[7]||'',questionCount:r[8]||'0',totalPoints:r[9]||'0',formName:r[10]||'',formUrl:'',editUrl:'',fileUrl:r[13]||'',status:r[14]||'Belum Posting',postedAt:r[15]||'',createdBy:r[16]||''})).filter(r=>r.id);}
function _getCachedSoalTestMasterById_(testId){const id=String(testId||'').trim();if(!id)return null;const key='TMS_ST_MASTER_'+id,cached=_cacheJsonGet_(key);if(cached)return cached;const r=_getCachedSoalTestHistoryRows_().find(x=>String(x.id||'').trim()===id);if(!r)return null;const master=[r.id,r.createdAt,r.company,r.agenda,r.type,r.batch,Number(r.duration||0),r.fileName,Number(r.questionCount||0),Number(r.totalPoints||0),r.formName||'','','',r.fileUrl||'',r.status||'Belum Posting',r.postedAt||'',r.createdBy||''];_cacheJsonPut_(key,master,SOAL_TEST_CACHE_TTL_SEC_);return master;}
function _getCachedSoalTestParticipantProfileFast_(nik){const key='TMS_ST_PROF_'+_normNik_(nik),cached=_cacheJsonGet_(key);if(cached)return cached;const p=_getSoalTestParticipantProfileFast_(nik);_cacheJsonPut_(key,p,60);return p;}

function _getSoalTestRootFolder_(){
  const props=PropertiesService.getScriptProperties(); let id=props.getProperty('SOAL_TEST_FOLDER_ID'); let folder=null;
  if(id){try{folder=DriveApp.getFolderById(id);}catch(e){folder=null;}}
  if(!folder){const it=DriveApp.getFoldersByName('Soal Test'); folder=it.hasNext()?it.next():DriveApp.createFolder('Soal Test'); props.setProperty('SOAL_TEST_FOLDER_ID',folder.getId());}
  try{_setDriveAnyoneEditor_(folder);}catch(e){}
  return folder;
}

function getSoalTestOptions(username){
  _authorizePemakai_(username);
  try{
    const result={companies:[],agendas:[],batches:[]};
    const companyRec=_records_(APP.SHEETS.PERUSAHAAN);
    if(companyRec.success){const i=_headerIndex_(companyRec.headers,['Nama Perusahaan','Perusahaan','Company']);if(i>=0)result.companies=[...new Set(companyRec.rows.map(r=>String(r.values[i]||'').trim()).filter(Boolean))].sort();}
    const dataRec=_records_(APP.SHEETS.DATA);
    if(dataRec.success){const ic=_headerIndex_(dataRec.headers,['Nama Perusahaan','Perusahaan','Company']);const ib=_headerIndex_(dataRec.headers,['Batch']);const it=_headerIndex_(dataRec.headers,['Nama Agenda','Nama Training','Training','Jenis Training']);
      if(!result.companies.length&&ic>=0)result.companies=[...new Set(dataRec.rows.map(r=>String(r.values[ic]||'').trim()).filter(Boolean))].sort();
      if(ib>=0)result.batches=[...new Set(dataRec.rows.map(r=>String(r.values[ib]||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
      if(it>=0)result.agendas=[...new Set(dataRec.rows.map(r=>String(r.values[it]||'').trim()).filter(Boolean))].sort();
    }
    const cal=_records_(APP.SHEETS.KALENDER_AGENDA);
    if(cal.success){const ia=_headerIndex_(cal.headers,['Agenda','Nama Agenda']);if(ia>=0)result.agendas=[...new Set(result.agendas.concat(cal.rows.map(r=>String(r.values[ia]||'').trim()).filter(Boolean)))] .sort();}
    return {success:true,...result};
  }catch(err){return {success:false,message:'Gagal memuat opsi Soal Test: '+err.message};}
}

function _getSoalTestHistoryRows_(){return _getCachedSoalTestHistoryRows_();}

function _soalTestTextKey_(value){
  return String(value==null?'':value)
    .trim().toLowerCase()
    .replace(/[\u00a0\s]+/g,' ')
    .replace(/[.,'’"`´`]/g,'')
    .replace(/\s+/g,'');
}

function _parseSoalTestDateTime_(value){
  if(value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
  const s=String(value||'').trim();
  if(!s) return null;
  let m=s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if(m){
    const d=Number(m[1]),mo=Number(m[2])-1,y=Number(m[3]),hh=Number(m[4]||0),mm=Number(m[5]||0),ss=Number(m[6]||0);
    const dt=new Date(y,mo,d,hh,mm,ss);
    return isNaN(dt.getTime())?null:dt;
  }
  const dt=new Date(s);
  return isNaN(dt.getTime())?null:dt;
}
function _soalTestStatusInfo_(row, now){
  const status=String(row?.status||'').trim();
  const postedAt=_parseSoalTestDateTime_(row?.postedAt);
  const duration=Number(row?.duration||0);
  if(/^posted$/i.test(status) && postedAt && Number.isFinite(duration) && duration>0){
    const expiresAt=new Date(postedAt.getTime()+duration*60000);
    if((now||new Date()).getTime()>=expiresAt.getTime()) return {status:'Expired',expiresAt};
    return {status:'Posted',expiresAt};
  }
  return {status:status||'Belum Posting',expiresAt:null};
}
function _syncSoalTestExpiredStatuses_(){
  const sh=_ensureSoalTestSheet_();
  const last=sh.getLastRow();
  if(last<2) return {expired:0};
  const vals=sh.getRange(2,1,last-1,17).getValues();
  const now=new Date(); let expired=0; let changed=false;
  vals.forEach((r,i)=>{
    const info=_soalTestStatusInfo_({status:r[14],postedAt:r[15],duration:r[6]},now);
    if(info.status==='Expired' && String(r[14]||'').trim().toLowerCase()==='posted'){
      sh.getRange(i+2,15).setValue('Expired'); expired++; changed=true;
    }
  });
  if(changed) SpreadsheetApp.flush();
  return {expired};
}
function getSoalTestRecords(username){
  _authorizePemakai_(username);
  try{
    _syncSoalTestExpiredStatuses_();
    const rows=_getSoalTestHistoryRows_();
    rows.forEach(r=>{ const info=_soalTestStatusInfo_(r); r.status=info.status; r.expiresAt=info.expiresAt?Utilities.formatDate(info.expiresAt,APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss'):''; });
    rows.reverse();
    return {success:true,rows};
  }catch(err){
    return {success:false,message:'Gagal memuat histori Soal Test: '+err.message};
  }
}

function _ensureSoalTestKeteranganColumn_(){
  const sh=_sheet_(APP.SHEETS.SOAL_TEST_RESPONSES)||_ss_().getSheetByName(APP.SHEETS.SOAL_TEST_RESPONSES);
  if(!sh) throw new Error(`Sheet "${APP.SHEETS.SOAL_TEST_RESPONSES}" tidak ditemukan.`);
  const lastCol=Math.max(19,sh.getLastColumn());
  const h=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  let idx=h.findIndex(v=>v.toLowerCase()==='keterangan');
  if(idx<0){idx=h.length;sh.getRange(1,idx+1).setValue('Keterangan');}
  return {sh,index:idx};
}

function markSoalTestCheating(nik,attemptId,answers,reason){
  try{
    const ss=_ensureSoalTestResponseSheets_();
    const ket=_ensureSoalTestKeteranganColumn_();
    const lock=LockService.getScriptLock();
    lock.waitLock(10000);
    try{
      const row=_soalTestFindAttemptRow_(ss.resp,attemptId);
      if(row<0) throw new Error('Sesi Soal Test tidak ditemukan.');
      const meta=ss.resp.getRange(row,1,1,Math.max(22,ss.resp.getLastColumn())).getValues()[0];
      const ownerNik=String(meta[17]||meta[2]||'').trim();
      if(_normNik_(ownerNik)!==_normNik_(nik)) throw new Error('Sesi bukan milik peserta ini.');

      const existing=String(ss.resp.getRange(row,ket.index+1).getDisplayValue()||'').trim();
      // Request bisa tiba lebih dari sekali (mis. fullscreenchange lalu beforeunload).
      // Jangan menghitung jawaban ulang atau menggandakan catatan yang sama.
      if(/kecurangan|curang|pelanggaran ujian/i.test(existing)){
        return {success:true,status:String(meta[11]||'SUBMITTED'),cheating:true,keterangan:existing,alreadyMarked:true};
      }

      const rawReason=String(reason||'Melanggar aturan ujian').replace(/[\r\n]+/g,' ').trim().slice(0,300)||'Melanggar aturan ujian';
      const cheatNote='Melakukan kecurangan: '+rawReason;
      const finalKeterangan=existing?existing+' | '+cheatNote:cheatNote;
      // Persist the violation first. If scoring fails due to a server-side issue,
      // the violation must still be visible in Nilai Test.
      ss.resp.getRange(row,ket.index+1).setValue(finalKeterangan);
      SpreadsheetApp.flush();

      let result;
      if(String(meta[11]||'').toUpperCase()==='SUBMITTED'){
        result={success:true,status:'SUBMITTED',score:Number(meta[12]||0),totalPoints:Number(meta[13]||0),percentage:Number(meta[14]||0),alreadySubmitted:true};
      }else{
        try{
          result=_soalTestSubmitAttemptInternal_(attemptId,answers||{},false);
          if(!result?.success)throw new Error(result?.message||'Jawaban tidak dapat difinalisasi.');
        }catch(submitError){
          // Jangan biarkan attempt yang melanggar aturan dilanjutkan lagi walau
          // perhitungan nilai gagal. Nilai existing tidak dihapus.
          const now=new Date();
          ss.resp.getRange(row,11).setValue(now);
          ss.resp.getRange(row,12).setValue('SUBMITTED');
          try{_soalTestAttemptIndexUpdateStatus_(meta[1],ownerNik,attemptId,row,'SUBMITTED');}catch(indexError){}
          try{_soalTestAttemptMetaPut_(attemptId,{row,attemptId:String(attemptId),testId:String(meta[1]||''),nik:ownerNik,status:'SUBMITTED',startedAt:meta[8] instanceof Date?meta[8].toISOString():String(meta[8]||''),durationMs:0});}catch(metaError){}
          result={success:true,status:'SUBMITTED',score:Number(meta[12]||0),totalPoints:Number(meta[13]||0),percentage:Number(meta[14]||0),scoreWarning:String(submitError?.message||submitError)};
        }
      }
      SpreadsheetApp.flush();
      return {...result,success:true,cheating:true,keterangan:finalKeterangan};
    }finally{
      try{lock.releaseLock();}catch(e){}
    }
  }catch(e){
    return {success:false,message:'Gagal mencatat pelanggaran ujian: '+e.message};
  }
}
function getSoalTestScores(username){
  _authorizePemakai_(username);
  try{
    const headers=['Attempt ID','Soal Test ID','NIK','Nama Peserta','Nama Perusahaan','Batch','Jenis Soal','Nama Agenda','Mulai','Terakhir Autosave','Selesai','Status','Nilai','Total Poin','Persentase','Departemen','Jabatan','NIK Akun (Internal)','Status Posting','Keterangan / Pelanggaran'];
    const {sh}=_ensureSoalTestKeteranganColumn_();
    if(!sh || sh.getLastRow()<2) return {success:true,headers,rows:[],summary:{total:0,submitted:0,inProgress:0,posted:0}};
    _ensureSoalTestPostingColumn_();
    const lastCol=Math.max(20,sh.getLastColumn());
    const vals=sh.getRange(2,1,sh.getLastRow()-1,lastCol).getDisplayValues();
    const rows=vals.map(r=>({
      attemptId:r[0]||'', testId:r[1]||'', nik:r[2]||'', nama:r[3]||'', perusahaan:r[4]||'', batch:r[5]||'',
      jenisSoal:r[6]||'', agenda:r[7]||'', mulai:r[8]||'', autosave:r[9]||'', selesai:r[10]||'', status:r[11]||'',
      nilai:r[12]||'', totalPoin:r[13]||'', persentase:r[14]||'', departemen:r[15]||'', jabatan:r[16]||'', postingStatus:r[18]||'Belum Posting', keterangan:r[19]||''
    })).filter(r=>r.attemptId);
    rows.reverse();
    const submitted=rows.filter(r=>String(r.status||'').trim().toUpperCase()==='SUBMITTED').length;
    const inProgress=rows.filter(r=>String(r.status||'').trim().toUpperCase()==='IN_PROGRESS').length;
    const posted=rows.filter(r=>String(r.postingStatus||'').trim().toUpperCase()==='POSTED').length;
    return {success:true,headers,rows,summary:{total:rows.length,submitted,inProgress,posted}};
  }catch(err){ return {success:false,message:'Gagal memuat Nilai Test: '+err.message}; }
}

function _normalizePostingDate_(value){
  if(value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value,APP.TIMEZONE,'yyyy-MM-dd');
  const s=String(value||'').trim(); if(!s) return '';
  let m=s.match(/(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})/);
  if(m) return `${m[3]}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  m=s.match(/(\d{4})[\/.-](\d{1,2})[\/.-](\d{1,2})/);
  if(m) return `${m[1]}-${String(m[2]).padStart(2,'0')}-${String(m[3]).padStart(2,'0')}`;
  const d=new Date(s); return isNaN(d.getTime())?'':Utilities.formatDate(d,APP.TIMEZONE,'yyyy-MM-dd');
}
function _postingDateValue_(value){
  const key=_normalizePostingDate_(value); if(!key) return '';
  const m=key.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m?new Date(Number(m[1]),Number(m[2])-1,Number(m[3])):value;
}
function _ensureSoalTestPostingColumn_(){
  const sh=_sheet_(APP.SHEETS.SOAL_TEST_RESPONSES)||_ss_().getSheetByName(APP.SHEETS.SOAL_TEST_RESPONSES);
  if(!sh) throw new Error(`Sheet "${APP.SHEETS.SOAL_TEST_RESPONSES}" tidak ditemukan.`);
  const lastCol=Math.max(18,sh.getLastColumn());
  const h=sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  let idx=h.findIndex(v=>v.toLowerCase()==='status posting');
  if(idx<0){idx=h.length;sh.getRange(1,idx+1).setValue('Status Posting');}
  return {sh,index:idx};
}
function _findDataRowForPostedScore_(dataSh, score){
  if(!dataSh || dataSh.getLastRow()<2) return -1;
  const lastCol=Math.max(1,dataSh.getLastColumn());
  const h=dataSh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  const ix={
    nik:_headerIndex_(h,['NIK']), nama:_headerIndex_(h,['Nama Peserta','Nama','Nama Lengkap']), company:_headerIndex_(h,['Nama Perusahaan','Perusahaan','Company']),
    dept:_headerIndex_(h,['Departemen','Department']), pos:_headerIndex_(h,['Jabatan','Position']), batch:_headerIndex_(h,['Batch']), training:_headerIndex_(h,['Nama Training','Nama Traning','Training']),
    date:_headerIndex_(h,['Tanggal Pelaksanaan','Tanggal']), type:_headerIndex_(h,['Jenis Nilai','Jenis Soal','Jenis Test']), value:_headerIndex_(h,['Nilai','Score'])
  };
  const vals=dataSh.getRange(2,1,dataSh.getLastRow()-1,lastCol).getDisplayValues();
  const norm=x=>_norm_(x);
  const targetDate=_normalizePostingDate_(score.mulai);
  for(let i=0;i<vals.length;i++){
    const r=vals[i];
    if(ix.nik>=0 && norm(r[ix.nik])!==norm(score.nik)) continue;
    if(ix.company>=0 && norm(r[ix.company])!==norm(score.perusahaan)) continue;
    if(ix.batch>=0 && _normalizeBatch_(r[ix.batch])!==_normalizeBatch_(score.batch)) continue;
    if(ix.training>=0 && norm(r[ix.training])!==norm(score.agenda)) continue;
    if(ix.type>=0 && norm(r[ix.type])!==norm(score.jenisSoal)) continue;
    if(ix.date>=0 && targetDate && _normalizePostingDate_(r[ix.date])!==targetDate) continue;
    return i+2;
  }
  return -1;
}
function _upsertPostedScoreToData_(score){
  const dataSh=_sheet_(APP.SHEETS.DATA); if(!dataSh) throw new Error(`Sheet "${APP.SHEETS.DATA}" tidak ditemukan.`);
  const info=_ensureDataStatusColumn_();
  const lastCol=Math.max(dataSh.getLastColumn(),info.index+1);
  const h=dataSh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v=>String(v||'').trim());
  const ix={
    nik:_headerIndex_(h,['NIK']), nama:_headerIndex_(h,['Nama Peserta','Nama','Nama Lengkap']), company:_headerIndex_(h,['Nama Perusahaan','Perusahaan','Company']),
    dept:_headerIndex_(h,['Departemen','Department']), pos:_headerIndex_(h,['Jabatan','Position']), batch:_headerIndex_(h,['Batch']), training:_headerIndex_(h,['Nama Training','Nama Traning','Training']),
    date:_headerIndex_(h,['Tanggal Pelaksanaan','Tanggal']), type:_headerIndex_(h,['Jenis Nilai','Jenis Soal','Jenis Test']), value:_headerIndex_(h,['Nilai','Score'])
  };
  const rowNum=_findDataRowForPostedScore_(dataSh,score);
  const row=rowNum>1 ? dataSh.getRange(rowNum,1,1,lastCol).getValues()[0] : Array(lastCol).fill('');
  if(ix.nik>=0) row[ix.nik]=score.nik||'';
  if(ix.nama>=0) row[ix.nama]=score.nama||'';
  if(ix.company>=0) row[ix.company]=score.perusahaan||'';
  if(ix.dept>=0) row[ix.dept]=score.departemen||'';
  if(ix.pos>=0) row[ix.pos]=score.jabatan||'';
  if(ix.batch>=0) row[ix.batch]=score.batch||'';
  if(ix.training>=0) row[ix.training]=score.agenda||'';
  if(ix.date>=0) row[ix.date]=_postingDateValue_(score.mulai)||score.mulai||'';
  if(ix.type>=0) row[ix.type]=score.jenisSoal||'';
  if(ix.value>=0) row[ix.value]=score.nilai===''?0:score.nilai;
  row[info.index]='Posted';
  if(rowNum>1){
    dataSh.getRange(rowNum,1,1,lastCol).setValues([row]);
    return rowNum;
  }
  dataSh.getRange(dataSh.getLastRow()+1,1,1,lastCol).setValues([row]);
  return dataSh.getLastRow();
}
function postSoalTestScores(username,attemptIds){
  _authorizePemakai_(username);
  try{
    const ids=[...new Set((Array.isArray(attemptIds)?attemptIds:[attemptIds]).map(v=>String(v||'').trim()).filter(Boolean))];
    if(!ids.length) return {success:false,message:'Pilih minimal satu Nilai Test untuk diposting.'};
    const {sh,index}=_ensureSoalTestPostingColumn_();
    const lastCol=Math.max(19,sh.getLastColumn());
    const vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),lastCol).getValues();
    let posted=0,skipped=0;
    ids.forEach(id=>{
      const i=vals.findIndex(r=>String(r[0]||'').trim()===id);
      if(i<0){skipped++;return;}
      const r=vals[i];
      const status=String(r[11]||'').trim().toUpperCase();
      const nilai=r[12];
      if(status!=='SUBMITTED' || nilai==='' || nilai==null || !Number.isFinite(Number(nilai))){skipped++;return;}
      _upsertPostedScoreToData_({nik:String(r[2]||''),nama:String(r[3]||''),perusahaan:String(r[4]||''),batch:String(r[5]||''),jenisSoal:String(r[6]||''),agenda:String(r[7]||''),mulai:r[8],mulaiDate:_normalizePostingDate_(r[8]),nilai:Number(nilai),departemen:String(r[15]||''),jabatan:String(r[16]||'')});
      sh.getRange(i+2,index+1).setValue('Posted'); posted++;
    });
    SpreadsheetApp.flush();
    return {success:true,message:`${posted} Nilai Test berhasil diposting ke Data.${skipped?` ${skipped} data dilewati karena belum Submitted/tidak valid.`:''}`,posted,skipped};
  }catch(e){return {success:false,message:'Gagal posting Nilai Test: '+e.message};}
}
function cancelPostSoalTestScores(username,attemptIds){
  _authorizePemakai_(username);
  try{
    const ids=[...new Set((Array.isArray(attemptIds)?attemptIds:[attemptIds]).map(v=>String(v||'').trim()).filter(Boolean))];
    if(!ids.length) return {success:false,message:'Pilih minimal satu Nilai Test untuk cancel posting.'};
    const {sh,index}=_ensureSoalTestPostingColumn_();
    const dataSh=_sheet_(APP.SHEETS.DATA); let canceled=0;
    if(!dataSh) throw new Error(`Sheet "${APP.SHEETS.DATA}" tidak ditemukan.`);
    const lastCol=Math.max(19,sh.getLastColumn());
    const vals=sh.getRange(2,1,Math.max(0,sh.getLastRow()-1),lastCol).getValues();
    ids.forEach(id=>{
      const i=vals.findIndex(r=>String(r[0]||'').trim()===id); if(i<0)return;
      const r=vals[i];
      const postingStatus=String(r[index]||'').trim().toUpperCase();
      if(postingStatus!=='POSTED') return;
      const score={nik:String(r[2]||''),nama:String(r[3]||''),perusahaan:String(r[4]||''),batch:String(r[5]||''),jenisSoal:String(r[6]||''),agenda:String(r[7]||''),mulai:r[8]};
      const rowNum=_findDataRowForPostedScore_(dataSh,score);
      if(rowNum>1) dataSh.deleteRow(rowNum);
      sh.getRange(i+2,index+1).setValue('Belum Posting'); canceled++;
    });
    SpreadsheetApp.flush();
    return {success:true,message:`${canceled} Nilai Test berhasil dibatalkan postingnya dari Data.`,canceled};
  }catch(e){return {success:false,message:'Gagal cancel posting Nilai Test: '+e.message};}
}

function _soalTestSafeCell_(value){return String(value==null?'':value).trim();}
function _soalTestFindHeader_(headers,patterns){for(const p of patterns){const i=headers.findIndex(h=>p.test(String(h||'')));if(i>=0)return i;}return -1;}
function _soalTestQuestionRows_(headers,rows){
  const h=headers.map(_soalTestSafeCell_), iq=_soalTestFindHeader_(h,[/^(no|nomor|number)$/i,/^no[.]?$/i]), is=_soalTestFindHeader_(h,[/soal/i,/pertanyaan/i,/question/i]);
  const oi={};['A','B','C','D','E'].forEach(l=>{const i=h.findIndex(v=>new RegExp('^(opsi\\s*)?'+l+'$','i').test(v)||new RegExp('^pilihan\\s*'+l+'$','i').test(v));if(i>=0)oi[l]=i;});
  const out=[];rows.forEach((r,ri)=>{const q=is>=0?_soalTestSafeCell_(r[is]):'';const has=Object.values(oi).some(i=>_soalTestSafeCell_(r[i]));if(!q&&!has)return;out.push({no:iq>=0&&_soalTestSafeCell_(r[iq])?_soalTestSafeCell_(r[iq]):String(ri+1),text:q||'Pertanyaan',options:['A','B','C','D','E'].filter(l=>oi[l]!==undefined&&_soalTestSafeCell_(r[oi[l]])).map(l=>({letter:l,value:_soalTestSafeCell_(r[oi[l]])}))});});return out;
}
function _soalTestExtractDetails_(headers,rows){
  const h=(headers||[]).map(_soalTestSafeCell_); const qIdx=_soalTestFindHeader_(h,[/soal/i,/pertanyaan/i,/question/i]); const noIdx=_soalTestFindHeader_(h,[/^(no|nomor|number)$/i,/^no[.]?$/i]); const pIdx=_soalTestFindHeader_(h,[/^(poin|point|nilai per soal|score per question)$/i,/poin/i,/point/i]); const kIdx=_soalTestFindHeader_(h,[/kunci/i,/jawaban\\s*benar/i,/answer\\s*key/i,/correct/i]);
  const oi={};['A','B','C','D','E'].forEach(l=>{const i=h.findIndex(v=>new RegExp('^(opsi\\s*)?'+l+'$','i').test(v)||new RegExp('^pilihan\\s*'+l+'$','i').test(v));if(i>=0)oi[l]=i;});
  const details=[];let total=0;(rows||[]).forEach((r,ri)=>{const q=qIdx>=0?_soalTestSafeCell_(r[qIdx]):'';const has=Object.values(oi).some(i=>_soalTestSafeCell_(r[i]));if(!q&&!has)return;const raw=pIdx>=0?String(r[pIdx]??'').replace(',','.').replace(/[^0-9.\\-]/g,''):'';const point=Number(raw);const safe=Number.isFinite(point)&&point>=0?point:0;total+=safe;details.push({no:noIdx>=0&&_soalTestSafeCell_(r[noIdx])?_soalTestSafeCell_(r[noIdx]):String(details.length+1),question:q,A:oi.A!==undefined?_soalTestSafeCell_(r[oi.A]):'',B:oi.B!==undefined?_soalTestSafeCell_(r[oi.B]):'',C:oi.C!==undefined?_soalTestSafeCell_(r[oi.C]):'',D:oi.D!==undefined?_soalTestSafeCell_(r[oi.D]):'',E:oi.E!==undefined?_soalTestSafeCell_(r[oi.E]):'',key:kIdx>=0?_soalTestSafeCell_(r[kIdx]):'',point:safe});});return {details,totalPoints:total,pointsProvided:pIdx>=0};
}
function _soalTestWriteSourceXlsx_(folder,fileName,headers,rows){
  const safe=_sanitizeTaskDriveName_(fileName)||'Soal_Test.xlsx';const finalName=/\.xlsx$/i.test(safe)?safe:safe+'.xlsx';const temp=SpreadsheetApp.create('TMP_'+Utilities.getUuid());
  try{const sh=temp.getSheets()[0];sh.setName('Soal Test');const matrix=[headers].concat(rows);sh.getRange(1,1,matrix.length,headers.length).setValues(matrix.map(r=>headers.map((_,i)=>r?.[i]??'')));SpreadsheetApp.flush();const url='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';const resp=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(resp.getResponseCode()>=300)throw new Error('Ekspor XLSX gagal ('+resp.getResponseCode()+').');const file=folder.createFile(resp.getBlob().setName(finalName));try{_setDriveAnyoneEditor_(file);}catch(e){}return {id:file.getId(),name:file.getName(),url:file.getUrl()};}finally{try{DriveApp.getFileById(temp.getId()).setTrashed(true);}catch(e){}}
}
function _soalTestFindRowById_(sh,id){const last=sh.getLastRow();if(last<2)return -1;const vals=sh.getRange(2,1,last-1,1).getDisplayValues();for(let i=0;i<vals.length;i++)if(String(vals[i][0]).trim()===String(id||'').trim())return i+2;return -1;}
function updateSoalTest(username,id,changes){_authorizePemakai_(username);try{const sh=_ensureSoalTestSheet_();const row=_soalTestFindRowById_(sh,id);if(row<0)throw new Error('Data Soal Test tidak ditemukan.');const cur=sh.getRange(row,1,1,17).getValues()[0],p=changes||{};const company=_soalTestSafeCell_(p.company||cur[2]),agenda=_soalTestSafeCell_(p.agenda||cur[3]),type=_soalTestSafeCell_(p.type||cur[4]),batch=_soalTestSafeCell_(p.batch||cur[5]),duration=Number(p.duration||cur[6]);if(!company||!agenda||!batch||!Number.isFinite(duration)||duration<=0)throw new Error('Nama perusahaan, nama agenda, batch, dan durasi wajib diisi.');if(!/^(Pre Test|Post Test) Hari (Pertama|Kedua|Ketiga|Keempat|Kelima)$/i.test(type))throw new Error('Jenis soal tidak valid.');sh.getRange(row,3,1,5).setValues([[company,agenda,type,batch,duration]]);_invalidateSoalTestCaches_(id);return {success:true,message:'Setting Soal Test berhasil diperbarui.'};}catch(e){return {success:false,message:'Gagal mengedit Soal Test: '+e.message};}}
function deleteSoalTest(username,id){
  _authorizePemakai_(username);
  try{
    const sh=_ensureSoalTestSheet_();
    const row=_soalTestFindRowById_(sh,id);
    if(row<0)throw new Error('Data Soal Test tidak ditemukan.');
    const r=sh.getRange(row,1,1,17).getValues()[0];
    const fileIds=[];
    const urlId=String(r[13]||'').match(/[-\w]{20,}/);
    if(urlId)fileIds.push(urlId[0]);
    const folderIds=new Set();
    fileIds.forEach(fid=>{ try{ const f=DriveApp.getFileById(fid); const parents=f.getParents(); while(parents.hasNext()){ const parent=parents.next(); folderIds.add(parent.getId()); } f.setTrashed(true); }catch(e){ console.warn('Gagal menghapus file Soal Test: '+e.message); } });
    // Ikut hapus folder khusus Soal Test yang menampung file XLSX. Jangan menghapus root aplikasi.
    folderIds.forEach(fid=>{ try{ const folder=DriveApp.getFolderById(fid); const nm=String(folder.getName()||'').toLowerCase(); if(nm!=='soal test') folder.setTrashed(true); }catch(e){ console.warn('Gagal menghapus folder Soal Test: '+e.message); } });
    const detail=_ensureSoalTestDetailSheet_();for(let i=detail.getLastRow();i>=2;i--)if(String(detail.getRange(i,1).getDisplayValue())===String(id))detail.deleteRow(i);
    try{
      const ss=_ensureSoalTestResponseSheets_();
      for(let i=ss.ans.getLastRow();i>=2;i--)if(String(ss.ans.getRange(i,2).getDisplayValue())===String(id))ss.ans.deleteRow(i);
      for(let i=ss.resp.getLastRow();i>=2;i--)if(String(ss.resp.getRange(i,2).getDisplayValue())===String(id))ss.resp.deleteRow(i);
      const idx=_ensureSoalTestAttemptIndexSheet_();
      _rebuildSoalTestAttemptIndexFromResponses_(ss.resp,idx,true);
    }catch(e){console.warn('Gagal membangun ulang indeks Soal Test: '+e.message);}
    sh.deleteRow(row);_invalidateSoalTestCaches_(id);
    return {success:true,message:'Soal Test, file XLSX, folder Soal Test, Detail, dan seluruh Responses berhasil dihapus.'};
  }catch(e){return {success:false,message:'Gagal menghapus Soal Test: '+e.message};}
}

function createSoalTestFromXlsx(username,payload){_authorizePemakai_(username);try{const p=payload||{},company=_soalTestSafeCell_(p.company),agenda=_soalTestSafeCell_(p.agenda),type=_soalTestSafeCell_(p.type),batch=_soalTestSafeCell_(p.batch),duration=Number(p.duration),fileName=_soalTestSafeCell_(p.fileName)||'Soal_Test.xlsx',headers=Array.isArray(p.headers)?p.headers.map(_soalTestSafeCell_):[],rows=Array.isArray(p.rows)?p.rows:[];if(!company||!agenda||!type||!batch||!Number.isFinite(duration)||duration<=0)return {success:false,message:'Nama perusahaan, nama agenda, jenis soal, batch, dan durasi wajib diisi.'};if(!/^(Pre Test|Post Test) Hari (Pertama|Kedua|Ketiga|Keempat|Kelima)$/i.test(type))return {success:false,message:'Jenis soal tidak valid.'};if(duration>600)return {success:false,message:'Durasi maksimal 600 menit.'};if(!headers.length||!rows.length)return {success:false,message:'Header dan data XLSX wajib tersedia.'};const d=_soalTestExtractDetails_(headers,rows);if(!d.pointsProvided)return {success:false,message:'Kolom Poin wajib tersedia pada file XLSX.'};if(!d.details.length)return {success:false,message:'Soal tidak dapat dikenali dari file XLSX.'};const errs=[];d.details.forEach(q=>{const key=String(q.key||'').trim().toUpperCase();const opts=['A','B','C','D','E'].filter(k=>String(q[k]||'').trim());if(opts.length<2)errs.push(`Soal ${q.no}: minimal 2 pilihan jawaban.`);if(!key||!opts.includes(key))errs.push(`Soal ${q.no}: Kunci Jawaban harus A-E sesuai opsi.`);if(!(Number(q.point)>0))errs.push(`Soal ${q.no}: Poin harus lebih dari 0.`);});if(errs.length)return {success:false,message:errs.slice(0,10).join('\n')};const root=_getSoalTestRootFolder_(),folderName=_sanitizeTaskDriveName_(`${agenda} – ${company} – Batch ${batch} – ${type}`);const it=root.getFoldersByName(folderName);const folder=it.hasNext()?it.next():root.createFolder(folderName);try{_setDriveAnyoneEditor_(folder);}catch(e){}const source=_soalTestWriteSourceXlsx_(folder,fileName,headers,rows),sh=_ensureSoalTestSheet_(),id='ST-'+Utilities.getUuid().slice(0,8).toUpperCase(),name=_sanitizeTaskDriveName_(`${type} – ${agenda} – ${company} – Batch ${batch}`);sh.getRange(sh.getLastRow()+1,1,1,17).setValues([[id,new Date(),company,agenda,type,batch,duration,fileName,d.details.length,d.totalPoints,name,'','',source.url,'Belum Posting','',username]]);const ds=_ensureSoalTestDetailSheet_(),dr=d.details.map(q=>[id,q.no,q.question,q.A,q.B,q.C,q.D,q.E,q.key,q.point]);ds.getRange(ds.getLastRow()+1,1,dr.length,10).setValues(dr);_invalidateSoalTestCaches_(id);return {success:true,message:'Soal Test custom berhasil dibuat.',id,formName:name,formUrl:'',editUrl:'',fileUrl:source.url,folderName:folder.getName(),questionCount:d.details.length,totalPoints:d.totalPoints};}catch(e){return {success:false,message:'Gagal membuat Soal Test: '+e.message};}}
function postSoalTest(username,id){_authorizePemakai_(username);try{const sh=_ensureSoalTestSheet_();const row=_soalTestFindRowById_(sh,id);if(row<0)throw new Error('Data Soal Test tidak ditemukan.');const now=Utilities.formatDate(new Date(),APP.TIMEZONE,'dd/MM/yyyy HH:mm:ss');sh.getRange(row,15,1,2).setValues([['Posted',now]]);_invalidateSoalTestCaches_(id);return {success:true,message:'Soal Test berhasil diposting ke Dashboard Peserta.',status:'Posted',postedAt:now};}catch(e){return {success:false,message:'Gagal posting Soal Test: '+e.message};}}
function cancelPostSoalTest(username,id){_authorizePemakai_(username);try{const sh=_ensureSoalTestSheet_();const row=_soalTestFindRowById_(sh,id);if(row<0)throw new Error('Data Soal Test tidak ditemukan.');sh.getRange(row,15,1,2).setValues([['Belum Posting','']]);_invalidateSoalTestCaches_(id);return {success:true,message:'Posting Soal Test berhasil dibatalkan.',status:'Belum Posting',postedAt:''};}catch(e){return {success:false,message:'Gagal membatalkan posting Soal Test: '+e.message};}}

function _getSoalTestParticipantInfo_(nik){
  const auth=_authorizePeserta_(nik); const h=_records_(APP.SHEETS.USER_PESERTA).headers; const v=auth.values||[];
  const iNama=_headerIndex_(h,['Nama Peserta','Nama','Nama Lengkap']); const iComp=_headerIndex_(h,['Nama Perusahaan','Perusahaan','Company']);
  return {nama:iNama>=0?String(v[iNama]||'').trim():'',company:iComp>=0?String(v[iComp]||'').trim():''};
}
function _soalTestParticipantBelongs_(nik,company,batch){
  const data=_records_(APP.SHEETS.DATA); if(!data.success)return false; const h=data.headers||[];const iNik=_headerIndex_(h,['NIK']),iComp=_headerIndex_(h,['Nama Perusahaan','Perusahaan','Company']),iBatch=_headerIndex_(h,['Batch']);
  return (data.rows||[]).some(r=>{const v=r.values||[];return _normNik_(v[iNik])===_normNik_(nik)&&_norm_(v[iComp])===_norm_(company)&&_normalizeBatch_(v[iBatch])===_normalizeBatch_(batch);});
}
function _soalTestFindAttemptRow_(sh,attemptId){
  const key='TMS_ST_ATT_'+String(attemptId),cached=_cacheJsonGet_(key);
  if(cached&&Number(cached.row)>1)return Number(cached.row);
  const idx=_ss_().getSheetByName('SoalTestAttemptIndex');
  if(idx&&idx.getLastRow()>=2){
    const found=idx.getRange(2,2,idx.getLastRow()-1,1).createTextFinder(String(attemptId)).matchEntireCell(true).findNext();
    if(found){const row=Number(idx.getRange(found.getRow(),3).getValue()||0);if(row>1){_cacheJsonPut_(key,{row},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);return row;}}
  }
  const last=sh.getLastRow(); if(last<2)return -1;
  const v=sh.getRange(2,1,last-1,1).getDisplayValues();
  for(let i=0;i<v.length;i++)if(String(v[i][0])===String(attemptId)){const row=i+2;_cacheJsonPut_(key,{row},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);return row;}
  return -1;
}

function _soalTestQuestionPayload_(testId){return _getCachedSoalTestQuestions_(testId);}

function _soalTestLoadSavedAnswers_(ansSh,attemptId,attemptRow,questions){
  const out={},json=String(attemptRow?.[21]||'').trim();
  if(json){try{const o=JSON.parse(json);if(o&&typeof o==='object'){Object.keys(o).forEach(k=>{const v=String(o[k]??'').trim().toUpperCase();if(v)out[String(k).trim()]=v;});return out;}}catch(e){}}
  const start=Number(attemptRow?.[20]||0),qs=Array.isArray(questions)?questions:[];
  if(start>1&&qs.length){ansSh.getRange(start,1,qs.length,7).getValues().forEach(v=>{const no=String(v[3]||'').trim();if(no)out[no]=String(v[4]||'').trim().toUpperCase();});return out;}
  const last=ansSh.getLastRow();if(last>=2)ansSh.getRange(2,1,last-1,7).getValues().forEach(v=>{if(String(v[0])===String(attemptId)){const no=String(v[3]||'').trim();if(no)out[no]=String(v[4]||'').trim().toUpperCase();}});
  return out;
}

function _soalTestProperCaseText_(value){
  return String(value==null?'':value).trim().toLowerCase().replace(/\s+/g,' ').replace(/(^|[\s-])(\p{L})/gu,function(_,prefix,ch){return prefix+ch.toUpperCase();});
}
function _soalTestProperCase_(v){ return String(v==null?'':v).trim().replace(/\s+/g,' ').toLowerCase().replace(/(^|[\s-])(\p{L})/gu,function(m,p,c){return p+c.toUpperCase();}); }
function _soalTestAttemptMetaPut_(attemptId,meta){
  _cacheJsonPut_('TMS_ST_META_'+String(attemptId||''),meta,SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
}
function _soalTestAttemptMetaGet_(attemptId){return _cacheJsonGet_('TMS_ST_META_'+String(attemptId||''));}
function _createSoalTestAttempt_(resp,ans,tr,testId,nik,participantData){
  const lock=LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    // Recheck inside the critical section. Two tabs/requests for the same NIK/test
    // must reuse the same attempt rather than creating duplicate response rows.
    const idx=_soalTestAttemptIndexFind_(testId,nik);
    if(idx&&idx.responseRow>1){
      const existing=resp.getRange(idx.responseRow,1,1,22).getValues()[0];
      if(String(existing[0]||'')===String(idx.attemptId||'') && String(existing[11]||'').toUpperCase()!=='SUBMITTED'){
        _cacheJsonPut_('TMS_ST_ATT_'+String(existing[0]),{row:idx.responseRow},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
        _soalTestAttemptMetaPut_(String(existing[0]),{row:idx.responseRow,attemptId:String(existing[0]),testId:String(testId),nik:String(nik),status:String(existing[11]||'IN_PROGRESS'),startedAt:existing[8] instanceof Date?existing[8].toISOString():String(existing[8]||''),durationMs:Number(tr[6]||0)*60000});
        return {attemptId:String(existing[0]||idx.attemptId),latest:existing,responseRow:idx.responseRow,existing:true};
      }
      if(String(existing[0]||'')===String(idx.attemptId||'') && String(existing[11]||'').toUpperCase()==='SUBMITTED'){
        return {attemptId:String(existing[0]||idx.attemptId),latest:existing,responseRow:idx.responseRow,existing:true,alreadySubmitted:true};
      }
    }
    const pd=participantData||{},now=new Date(),attemptId='ATT-'+Utilities.getUuid().replace(/-/g,'').slice(0,24).toUpperCase();
    const row=[
      attemptId,String(testId||''),String(nik||pd.nik||''),_soalTestProperCaseText_(pd.name||''),String(pd.company||tr[2]||''),String(pd.batch||tr[5]||''),
      String(pd.type||pd.jenisNilai||tr[4]||''),String(pd.training||pd.agenda||tr[3]||''),now,'','', 'IN_PROGRESS',0,Number(tr[9]||0),0,
      _soalTestProperCaseText_(pd.department||''),_soalTestProperCaseText_(pd.position||''),String(nik||pd.nik||''),String(tr[14]||'Posted'),'','', '{}'
    ];
    const responseRow=resp.getLastRow()+1;
    resp.getRange(responseRow,1,1,row.length).setValues([row]);
    _soalTestAttemptIndexAppend_(testId,nik,attemptId,responseRow,'IN_PROGRESS');
    const latest=row.slice();
    _soalTestAttemptMetaPut_(attemptId,{row:responseRow,attemptId,testId:String(testId||''),nik:String(nik||pd.nik||''),status:'IN_PROGRESS',startedAt:now.toISOString(),durationMs:Number(tr[6]||0)*60000});
    return {attemptId,latest,responseRow,existing:false};
  }finally{try{lock.releaseLock();}catch(e){}}
}

/**
 * Jalankan manual dari editor Apps Script sebelum ujian penting.
 * Migrasi indeks dilakukan satu kali dan cache master/soal untuk ujian Posted dipanaskan.
 * Fungsi private (awalan underscore) sehingga tidak tersedia untuk google.script.run.
 */
function _prepareSoalTestConcurrency_(testId){
  const started=Date.now();
  try{
    const respSet=_ensureSoalTestResponseSheets_(),indexSheet=_ensureSoalTestAttemptIndexSheet_();
    const migration=_rebuildSoalTestAttemptIndexFromResponses_(respSet.resp,indexSheet,false);
    const history=_getCachedSoalTestHistoryRows_();
    const wanted=String(testId||'').trim();
    const tests=history.filter(r=>String(r.status||'').trim().toLowerCase()==='posted'&&(!wanted||String(r.id||'').trim()===wanted));
    if(wanted&&!tests.length)throw new Error('ID Soal Test tidak ditemukan atau statusnya bukan Posted: '+wanted);
    const warmed=[];
    tests.forEach(r=>{
      const master=_getCachedSoalTestMasterById_(r.id);
      const questions=_getCachedSoalTestQuestions_(r.id);
      const cacheCheck=_soalTestQuestionCacheGet_('TMS_ST_Q_'+String(r.id||'').trim());
      warmed.push({id:r.id,questions:questions.length,cached:Array.isArray(cacheCheck),ok:!!master&&questions.length>0&&Array.isArray(cacheCheck)});
    });
    const failures=warmed.filter(x=>!x.ok);
    if(!warmed.length)failures.push({reason:'Tidak ada Soal Test berstatus Posted untuk dipanaskan.'});
    const result={success:warmed.length>0&&failures.length===0,migrated:!!migration.migrated,indexed:Number(migration.indexed||0),postedExams:warmed.length,warmed,elapsedMs:Date.now()-started,failures};
    console.log('[TMS_EXAM_PREFLIGHT] '+JSON.stringify(result));
    return result;
  }catch(e){
    const result={success:false,message:e.message||String(e),elapsedMs:Date.now()-started};
    console.error('[TMS_EXAM_PREFLIGHT] '+JSON.stringify(result));
    return result;
  }
}

function startSoalTestAttempt(nik,testId,participantData){
  try{
    const tr=_getCachedSoalTestMasterById_(testId); if(!tr)throw new Error('Soal Test tidak ditemukan.');
    if(String(tr[14]).toLowerCase()!=='posted')return {success:false,message:'Soal Test belum diposting atau sudah dibatalkan.'};
    const respSet=_ensureSoalTestResponseSheets_(),resp=respSet.resp,cleanNik=_normNik_(nik);
    const indexSheet=_ensureSoalTestAttemptIndexSheet_();
    _ensureSoalTestAttemptIndexMigrated_(resp,indexSheet);
    const idxEntry=_soalTestAttemptIndexFind_(testId,cleanNik);
    if(idxEntry&&idxEntry.responseRow>1){
      let v=resp.getRange(idxEntry.responseRow,1,1,22).getValues()[0];
      if(String(v[11]||'').toUpperCase()==='SUBMITTED'){
        return {success:false,code:'ALREADY_SUBMITTED',message:'Soal Test ini sudah selesai dikerjakan dan tidak dapat dibuka kembali.'};
      }
      const started=new Date(v[8]),durationMs=Number(tr[6]||0)*60000;
      if(!isNaN(started.getTime()) && Date.now()<started.getTime()+durationMs){
        if(participantData){
          const pd=participantData||{};
          const merged=[String(pd.nik||v[2]||cleanNik).trim(),_soalTestProperCaseText_(pd.name||v[3]||''),String(pd.company||v[4]||tr[2]).trim(),String(pd.batch||v[5]||tr[5]).trim(),String(pd.type||pd.jenisNilai||v[6]||tr[4]).trim(),String(pd.training||pd.agenda||v[7]||tr[3]).trim()];
          v[2]=merged[0];v[3]=merged[1];v[4]=merged[2];v[5]=merged[3];v[6]=merged[4];v[7]=merged[5];
          v[15]=_soalTestProperCaseText_(pd.department||v[15]||'');v[16]=_soalTestProperCaseText_(pd.position||v[16]||'');
          resp.getRange(idxEntry.responseRow,3,1,6).setValues([[v[2],v[3],v[4],v[5],v[6],v[7]]]);
          resp.getRange(idxEntry.responseRow,16,1,2).setValues([[v[15],v[16]]]);
        }
        _cacheJsonPut_('TMS_ST_ATT_'+String(v[0]),{row:idxEntry.responseRow},SOAL_TEST_ATTEMPT_CACHE_TTL_SEC_);
        _soalTestAttemptMetaPut_(String(v[0]),{row:idxEntry.responseRow,attemptId:String(v[0]),testId:String(testId),nik:cleanNik,status:String(v[11]||'IN_PROGRESS'),startedAt:started.toISOString(),durationMs});
        return _soalTestAttemptPayload_(testId,String(v[0]),tr,v,respSet.ans);
      }
      return _soalTestSubmitAttemptInternal_(String(v[0]),{},true);
    }

    // Riwayat lama sudah diindeks satu kali secara bulk di atas. Hindari memindai
    // seluruh SoalTestResponses untuk setiap NIK baru saat satu batch mulai ujian.
    const created=_createSoalTestAttempt_(resp,respSet.ans,tr,String(testId),cleanNik,participantData);
    if(created.alreadySubmitted)return {success:false,code:'ALREADY_SUBMITTED',message:'Soal Test ini sudah selesai dikerjakan dan tidak dapat dibuka kembali.'};
    return _soalTestAttemptPayload_(testId,created.attemptId,tr,created.latest,respSet.ans);
  }catch(e){return {success:false,message:'Gagal memulai Soal Test: '+e.message};}
}

function _soalTestAttemptPayload_(testId,attemptId,tr,attemptRow,ansSh){const started=new Date(attemptRow[8]),duration=Number(tr[6]||0),questions=_soalTestQuestionPayload_(testId).map(q=>({no:q.no,question:q.question,options:q.options,point:q.point}));return {success:true,attemptId,testId,startedAt:started.toISOString(),endsAt:new Date(started.getTime()+duration*60000).toISOString(),duration,company:tr[2],agenda:tr[3],type:tr[4],batch:tr[5],serverNowMs:Date.now(),questions,answers:_soalTestLoadSavedAnswers_(ansSh,attemptId,attemptRow,questions)};}

function saveSoalTestAnswers(nik,attemptId,answers){
  try{
    const ss=_ensureSoalTestResponseSheets_();
    let metaInfo=_soalTestAttemptMetaGet_(attemptId);
    let row=Number(metaInfo?.row||0);
    if(row<2||String(metaInfo?.attemptId||'')!==String(attemptId)){
      row=_soalTestFindAttemptRow_(ss.resp,attemptId);
      if(row<2)throw new Error('Sesi Soal Test tidak ditemukan.');
      const m=ss.resp.getRange(row,1,1,18).getValues()[0];
      metaInfo={row,attemptId:String(attemptId),testId:String(m[1]||''),nik:String(m[17]||m[2]||''),status:String(m[11]||''),startedAt:m[8] instanceof Date?m[8].toISOString():String(m[8]||''),durationMs:0};
      _soalTestAttemptMetaPut_(attemptId,metaInfo);
    }
    const ownerNik=String(metaInfo.nik||'').trim();
    if(_normNik_(ownerNik)!==_normNik_(nik))throw new Error('Sesi bukan milik akun peserta ini.');
    if(String(metaInfo.status||'').toUpperCase()==='SUBMITTED')return {success:true,status:'SUBMITTED'};
    const test=_getCachedSoalTestMasterById_(metaInfo.testId);if(!test)throw new Error('Soal Test tidak ditemukan.');
    const started=new Date(metaInfo.startedAt);
    if(isNaN(started.getTime()))throw new Error('Waktu mulai ujian tidak valid.');
    const durationMs=Number(test[6]||0)*60000;
    if(Date.now()>=started.getTime()+durationMs){_soalTestAttemptMetaPut_(attemptId,{...metaInfo,status:'SUBMITTED'});return _soalTestSubmitAttemptInternal_(attemptId,answers,true);}

    // Browser selalu mengirim snapshot lengkap seluruh jawaban. Tidak perlu membaca
    // Answer JSON lama lalu menggabungkannya; overwrite atomik mencegah jawaban lama
    // tertinggal dan mengurangi satu operasi baca/parse per autosave.
    const clean=answers&&typeof answers==='object'?answers:{};
    const compact={};
    Object.keys(clean).forEach(k=>{const no=String(k||'').trim(),v=String(clean[k]??'').trim().toUpperCase();if(no&&v)compact[no]=v;});
    const serialized=JSON.stringify(compact);
    if(serialized.length>45000)throw new Error('Data jawaban terlalu besar untuk disimpan.');
    const now=new Date();
    ss.resp.getRange(row,10).setValue(now);
    ss.resp.getRange(row,22).setValue(serialized);
    return {success:true,status:'IN_PROGRESS',savedAt:now.toISOString(),savedCount:Object.keys(compact).length};
  }catch(e){return {success:false,message:'Gagal menyimpan jawaban: '+e.message};}
}

function _soalTestSubmitAttemptInternal_(attemptId,answers,autoSubmitted){
  const ss=_ensureSoalTestResponseSheets_(),row=_soalTestFindAttemptRow_(ss.resp,attemptId);if(row<0)throw new Error('Sesi Soal Test tidak ditemukan.');
  const meta=ss.resp.getRange(row,1,1,22).getValues()[0];if(String(meta[11]).toUpperCase()==='SUBMITTED')return {success:true,status:'SUBMITTED',score:Number(meta[12]||0),totalPoints:Number(meta[13]||0),percentage:Number(meta[14]||0)};
  const questions=_getCachedSoalTestQuestions_(meta[1]);if(!questions.length)throw new Error('Daftar soal tidak ditemukan.');const clean=answers&&typeof answers==='object'?answers:{},savedMap={},json=String(meta[21]||'').trim();
  if(json){try{const parsed=JSON.parse(json);if(parsed&&typeof parsed==='object')Object.keys(parsed).forEach(k=>{savedMap[String(k).trim()]=String(parsed[k]??'').trim().toUpperCase();});}catch(e){}}
  else{const start=Number(meta[20]||0);if(start>1){ss.ans.getRange(start,1,questions.length,7).getValues().forEach((v,i)=>{const no=String(v[3]||questions[i]?.no||'').trim();if(no)savedMap[no]=String(v[4]||'').trim().toUpperCase();});}else{const last=ss.ans.getLastRow();if(last>=2)ss.ans.getRange(2,1,last-1,7).getValues().forEach(v=>{if(String(v[0])!==String(attemptId))return;const no=String(v[3]||'').trim();if(no)savedMap[no]=String(v[4]||'').trim().toUpperCase();});}}
  Object.keys(clean).forEach(k=>{savedMap[String(k).trim()]=String(clean[k]??'').trim().toUpperCase();});const now=new Date();let score=0,total=0;const finalRows=questions.map(q=>{const no=String(q.no||'').trim(),key=String(q.key||'').trim().toUpperCase(),point=Number(q.point||0),answer=savedMap[no]||'',earned=answer&&key&&answer===key?point:0;score+=earned;total+=point;return [attemptId,meta[1],String(meta[2]||''),no,answer,earned,now];});
  const answerStart=Number(meta[20]||0);if(answerStart>1)ss.ans.getRange(answerStart,1,finalRows.length,7).setValues(finalRows);else{const st=ss.ans.getLastRow()+1;if(finalRows.length)ss.ans.getRange(st,1,finalRows.length,7).setValues(finalRows);ss.resp.getRange(row,21).setValue(st);}
  const pct=total>0?score/total*100:0;ss.resp.getRange(row,10,1,6).setValues([[now,now,'SUBMITTED',score,total,pct]]);ss.resp.getRange(row,22).setValue('');
  try{const ownerNik=String(meta[17]||meta[2]||'').trim();_soalTestAttemptIndexUpdateStatus_(meta[1],ownerNik,attemptId,row,'SUBMITTED');_soalTestAttemptMetaPut_(attemptId,{row,attemptId:String(attemptId),testId:String(meta[1]||''),nik:ownerNik,status:'SUBMITTED',startedAt:meta[8] instanceof Date?meta[8].toISOString():String(meta[8]||''),durationMs:0});}catch(e){}
  try{CacheService.getScriptCache().remove('TMS_ST_ATT_'+String(attemptId));}catch(e){}
  return {success:true,status:'SUBMITTED',score,totalPoints:total,percentage:pct,submittedAt:now.toISOString(),autoSubmitted:!!autoSubmitted};
}

function submitSoalTestAttempt(nik,attemptId,answers){try{const ss=_ensureSoalTestResponseSheets_(),row=_soalTestFindAttemptRow_(ss.resp,attemptId);if(row<0)throw new Error('Sesi Soal Test tidak ditemukan.');const meta=ss.resp.getRange(row,1,1,18).getValues()[0];const ownerNik=String(meta[17]||meta[2]||'').trim();if(_normNik_(ownerNik)!==_normNik_(nik))throw new Error('Sesi bukan milik akun peserta ini.');return _soalTestSubmitAttemptInternal_(attemptId,answers,false);}catch(e){return {success:false,message:'Gagal mengirim jawaban Soal Test: '+e.message};}}
function autoSubmitSoalTestAttempt(nik,attemptId,answers){try{const ss=_ensureSoalTestResponseSheets_(),row=_soalTestFindAttemptRow_(ss.resp,attemptId);if(row<0)throw new Error('Sesi Soal Test tidak ditemukan.');const meta=ss.resp.getRange(row,1,1,18).getValues()[0];const ownerNik=String(meta[17]||meta[2]||'').trim();if(_normNik_(ownerNik)!==_normNik_(nik))throw new Error('Sesi bukan milik peserta ini.');if(String(meta[11]).toUpperCase()==='SUBMITTED')return {success:true,status:'SUBMITTED',score:Number(meta[12]||0),totalPoints:Number(meta[13]||0),percentage:Number(meta[14]||0),alreadySubmitted:true};return _soalTestSubmitAttemptInternal_(attemptId,answers||{},true);}catch(e){return {success:false,message:'Gagal menyimpan otomatis jawaban Soal Test: '+e.message};}}
function _getPostedSoalTestHistoryRows_(){
  const now=new Date();
  // Use the effective status, not only the stored text. Expired postings must never
  // appear in the public exam list even if the management page has not refreshed yet.
  return _getSoalTestHistoryRows_().filter(r=>_soalTestStatusInfo_(r,now).status==='Posted');
}

function hasPostedSoalTestForExam(){
  // Fast availability check for the public landing page. Do not wait on the shared
  // history-cache lock (which can wait several seconds during a cold-cache burst).
  // Also evaluate expiry here: a row whose status text still says Posted is not
  // available once postedAt + duration has elapsed.
  try{
    const key='TMS_ST_HAS_POSTED',cache=CacheService.getScriptCache(),cached=cache.get(key);
    if(cached!==null)return {success:true,hasPosted:String(cached)==='1'};
    let rows=_cacheJsonGet_('TMS_ST_HISTORY');
    if(!Array.isArray(rows)) rows=_getSoalTestHistoryRowsRaw_();
    const now=new Date();
    const has=rows.some(r=>_soalTestStatusInfo_(r,now).status==='Posted');
    // Short TTL bounds stale availability without polling the sheet on every visit.
    try{cache.put(key,has?'1':'0',8);}catch(e){}
    return {success:true,hasPosted:has};
  }catch(e){ return {success:false,hasPosted:false,message:e.message}; }
}
function getPostedSoalTestHistoryForExam(){
  try{
    const history=_getPostedSoalTestHistoryRows_();
    const candidates=history.map(r=>({
      id:r.id, createdAt:r.createdAt, company:r.company, agenda:r.agenda, type:r.type,
      batch:r.batch, duration:Number(r.duration||0), questionCount:Number(r.questionCount||0),
      totalPoints:Number(r.totalPoints||0), formName:r.formName||`${r.type||'Soal Test'} – ${r.agenda||''}`,
      status:r.status, postedAt:r.postedAt,
      label:`${r.type||'Soal Test'} — ${r.agenda||'Soal Test'}`,
      training:r.agenda, dateTime:r.createdAt, jenisNilai:r.type
    }));
    candidates.sort((a,b)=>{
      const da=_parseDateValue_(a.postedAt), db=_parseDateValue_(b.postedAt);
      return (db instanceof Date?db.getTime():0)-(da instanceof Date?da.getTime():0);
    });
    return {success:true,candidates};
  }catch(e){
    return {success:false,message:'Gagal memuat Riwayat Soal Test Posted: '+e.message};
  }
}

function getPostedSoalTestHistoryForPeserta(nik){
  _authorizePesertaFast_(nik);
  try{
    const cleanNik=_normNik_(nik);
    const auth=_findCachedLoginRow_('peserta',cleanNik)||{};
    const profile={nik:cleanNik,name:String(auth.nama||''),company:String(auth.perusahaan||''),department:String(auth.department||auth.departemen||''),position:String(auth.jabatan||''),batches:[]};
    // Sumber master Exam = Riwayat Soal Test. Nama Perusahaan peserta TIDAK dicocokkan.
    // Sumber Soal Test yang Dipilih HARUS langsung dari Riwayat Soal Test.
    // Tidak ada filter Nama Perusahaan maupun Batch pada daftar pilihan.
    // Satu-satunya filter adalah Status Posting = Posted.
    const history=_getPostedSoalTestHistoryRows_();

    const candidates=history.map(r=>({
      id:r.id, createdAt:r.createdAt, company:r.company, agenda:r.agenda, type:r.type,
      batch:r.batch, duration:Number(r.duration||0), questionCount:Number(r.questionCount||0),
      totalPoints:Number(r.totalPoints||0), formName:r.formName||`${r.type||'Soal Test'} – ${r.agenda||''}`,
      formUrl:'', editUrl:'', status:r.status, postedAt:r.postedAt,
      label:`${r.type||'Soal Test'} — ${r.agenda||'Soal Test'}`,
      training:r.agenda, dateTime:r.createdAt, jenisNilai:r.type
    }));

    candidates.sort((a,b)=>{
      const da=_parseDateValue_(a.postedAt); const db=_parseDateValue_(b.postedAt);
      return (db instanceof Date?db.getTime():0)-(da instanceof Date?da.getTime():0);
    });
    return {success:true,profile,candidates};
  }catch(e){
    return {success:false,message:'Gagal memuat data Soal Test peserta: '+e.message};
  }
}

function _getSoalTestParticipantProfileFast_(cleanNik){
  let name='',company='',department='',position='';
  const batches=new Set();

  const up=_records_(APP.SHEETS.USER_PESERTA);
  if(up.success){
    const h=up.headers||[];
    const iNik=_headerIndex_(h,['NIK']);
    const iNama=_headerIndex_(h,['Nama Peserta','Nama','Nama Lengkap']);
    const iComp=_headerIndex_(h,['Nama Perusahaan','Perusahaan','Company']);
    const iDept=_headerIndex_(h,['Departemen','Department']);
    const iPos=_headerIndex_(h,['Jabatan','Position']);
    const row=(up.rows||[]).find(r=>iNik>=0&&_normNik_(r.values?.[iNik])===cleanNik);
    if(row){
      if(iNama>=0)name=String(row.values?.[iNama]||'').trim();
      if(iComp>=0)company=String(row.values?.[iComp]||'').trim();
      if(iDept>=0)department=String(row.values?.[iDept]||'').trim();
      if(iPos>=0)position=String(row.values?.[iPos]||'').trim();
    }
  }

  const data=_records_(APP.SHEETS.DATA);
  if(data.success){
    const h=data.headers||[];
    const iNik=_headerIndex_(h,['NIK']);
    const iNama=_headerIndex_(h,['Nama Peserta','Nama','Nama Lengkap']);
    const iComp=_headerIndex_(h,['Nama Perusahaan','Perusahaan','Company']);
    const iDept=_headerIndex_(h,['Departemen','Department']);
    const iPos=_headerIndex_(h,['Jabatan','Position']);
    const iBatch=_headerIndex_(h,['Batch']);
    (data.rows||[]).forEach(r=>{
      const v=r.values||[];
      if(iNik<0||_normNik_(v[iNik])!==cleanNik)return;
      if(!name&&iNama>=0)name=String(v[iNama]||'').trim();
      if(!company&&iComp>=0)company=String(v[iComp]||'').trim();
      if(!department&&iDept>=0)department=String(v[iDept]||'').trim();
      if(!position&&iPos>=0)position=String(v[iPos]||'').trim();
      if(iBatch>=0&&String(v[iBatch]||'').trim())batches.add(_normalizeBatch_(v[iBatch]));
    });
  }

  // Setting Tugas menjadi fallback batch, tetapi TIDAK menjadi sumber Soal Test.
  const st=_records_(APP.SHEETS.SETTING_TUGAS);
  if(st.success){
    const rows=st.rows||[];
    const iComp=_headerIndex_(st.headers||[],['Nama Perusahaan','Perusahaan','Company']);
    const iBatch=_headerIndex_(st.headers||[],['Batch']);
    if(iComp>=0&&iBatch>=0){
      rows.forEach(r=>{
        const v=r.values||[];
        if(company&&_soalTestTextKey_(v[iComp])!==_soalTestTextKey_(company))return;
        if(String(v[iBatch]||'').trim())batches.add(_normalizeBatch_(v[iBatch]));
      });
    }
  }
  return {nik:cleanNik,name,company,department,position,batches:[...batches].filter(Boolean)};
}

function getPesertaSoalTests(nik){
  _authorizePesertaFast_(nik);
  try{
    const base=getPostedSoalTestHistoryForPeserta(nik); if(!base.success)return base;
    _ensureSoalTestAttemptIndexSheet_();
    const cleanNik=_normNik_(nik);
    const rows=(base.candidates||[]).map(r=>{
      const a=_soalTestAttemptIndexFind_(r.id,cleanNik);
      return {...r,attemptId:a?.attemptId||'',attemptStatus:a?.status||'',score:0,percentage:0};
    });
    // Read score/percentage only for attempts that actually exist. Each is a single row.
    rows.forEach(r=>{
      if(!r.attemptId||!r.responseRow)return;
      const sh=_ss_().getSheetByName(APP.SHEETS.SOAL_TEST_RESPONSES); if(!sh)return;
      const v=sh.getRange(Number(r.responseRow),12,1,4).getValues()[0];
      r.attemptStatus=String(v[0]||r.attemptStatus||''); r.score=Number(v[1]||0); r.percentage=Number(v[3]||0);
    });
    return {success:true,rows};
  }catch(e){return {success:false,message:'Gagal memuat Soal Test peserta: '+e.message};}
}

function getPesertaSoalTestParticipantData(nik){
  return getPostedSoalTestHistoryForPeserta(nik);
}

function getPesertaSoalTestAgenda(nik,testId){
  _authorizePesertaFast_(nik);
  try{
    const sh=_ensureSoalTestSheet_();
    const row=_soalTestFindRowById_(sh,testId);
    if(row<0)return {success:false,message:'Soal Test tidak ditemukan.'};
    const r=sh.getRange(row,1,1,17).getValues()[0];
    if(String(r[14]||'').trim().toLowerCase()!=='posted')return {success:false,message:'Soal Test belum diposting atau sudah dibatalkan.'};
    return {success:true,id:String(r[0]||''),agenda:String(r[3]||'').trim(),company:String(r[2]||'').trim(),batch:String(r[5]||'').trim(),type:String(r[4]||'').trim(),postedAt:String(r[15]||'')};
  }catch(e){return {success:false,message:'Gagal mengambil Nama Agenda Soal Test: '+e.message};}
}

function _sanitizeTaskDriveName_(value) {
  return String(value || '')
    .replace(/[\\/:*?"<>|#%{}]/g,'_')
    .replace(/\s+/g,' ')
    .trim()
    .slice(0,120) || 'Tugas';
}

function _ensureTaskFolderBySetting_(settingId) {
  const task = _findTaskById_(settingId);
  if (!task) throw new Error('Setting tugas tidak ditemukan.');
  const root = _getTugasTrainingFolder_();
  const training = _sanitizeTaskDriveName_(task.training || task.name || 'Training');
  const company = _sanitizeTaskDriveName_(task.company || 'Nama Perusahaan');
  const rawBatch = _sanitizeTaskDriveName_(task.batch || 'Batch');
  const batch = /^batch\b/i.test(rawBatch) ? rawBatch : `Batch ${rawBatch}`;
  const folderName = _sanitizeTaskDriveName_(`${training} – ${company} – ${batch}`);

  // Utamakan folder dengan nama yang sama yang sudah ada di Drive,
  // sehingga upload tidak membuat folder duplikat hanya karena
  // TASK_TRAINING_FOLDER_ID menunjuk ke root yang berbeda.
  let folder = null;
  try {
    const existing = DriveApp.getFoldersByName(folderName);
    if (existing.hasNext()) folder = existing.next();
  } catch (e) {
    console.warn('Pencarian folder tugas global gagal: ' + e.message);
  }

  // Bila belum ditemukan, cari lagi di root aplikasi sebelum membuat baru.
  if (!folder) {
    const it = root.getFoldersByName(folderName);
    folder = it.hasNext() ? it.next() : null;
  }

  // Folder baru hanya dibuat jika memang belum ada folder dengan nama tersebut.
  if (!folder) folder = root.createFolder(folderName);

  _setDriveAnyoneEditor_(folder);
  return {task, root, folder, folderName};
}

function getTaskUploadOptions(username) {
  _authorizePemakai_(username);
  try {
    const sh = _ensureTaskSettingSheet_();
    const last = sh.getLastRow();
    if (last < 2) return {success:true, tasks:[]};
    const vals = sh.getRange(2,1,last-1,6).getValues();
    const tasks = vals.map((r, idx) => {
      const id=String(r[0]||'').trim();
      const company=String(r[1]||'').trim();
      const training=String(r[2]||'').trim();
      const batch=String(r[3]||'').trim();
      if(!id || !company || !training || !batch) return null;
      const rawBatch = String(batch);
      const batchLabel = /^batch\b/i.test(rawBatch) ? rawBatch : `Batch ${rawBatch}`;
      return {id, company, training, batch:batchLabel, folderName:`${training} – ${company} – ${batchLabel}`};
    }).filter(Boolean);
    return {success:true,tasks};
  } catch(err) { return {success:false,message:'Gagal memuat folder tugas: '+err.message}; }
}

function uploadTaskMaterial(username, settingId, fileName, dataUrl, mimeType) {
  _authorizePemakai_(username);
  try {
    const raw=String(dataUrl||'');
    const name=String(fileName||'').trim();
    const mime=String(mimeType||'').trim().toLowerCase();
    if(!name || !raw.startsWith('data:')) return {success:false,message:'File tugas tidak valid.'};
    if(name.length>180) return {success:false,message:'Nama file terlalu panjang.'};
    const parts=raw.split(',');
    if(parts.length<2) return {success:false,message:'Data file tidak valid.'};
    const bytes=Utilities.base64Decode(parts[1]);
    if(bytes.length>20*1024*1024) return {success:false,message:'Ukuran file maksimal 20 MB.'};
    const safeName=_sanitizeTaskDriveName_(name) || 'Materi Tugas';
    const finalName=safeName.includes('.') ? safeName : safeName + (mime==='application/pdf'?'.pdf':'');
    const {task,folder,folderName}=_ensureTaskFolderBySetting_(String(settingId||'').trim());
    const blob=Utilities.newBlob(bytes,mime||'application/octet-stream',finalName);
    const file=folder.createFile(blob);
    _setDriveAnyoneEditor_(file);
    file.setDescription(`Materi tugas untuk ${task.training} - ${task.company} - ${task.batch}.`);
    return {success:true,message:'Materi tugas berhasil diupload.',fileId:file.getId(),fileName:file.getName(),fileUrl:file.getUrl(),folderName,folderUrl:folder.getUrl()};
  } catch(err){ return {success:false,message:'Gagal mengupload materi tugas: '+err.message}; }
}


function getTaskMaterialFiles(username, settingId) {
  _authorizePemakai_(username);
  try {
    const wantedSettingId = String(settingId || '').trim();
    const files = [];

    if (wantedSettingId) {
      const {folder, folderName, task} = _ensureTaskFolderBySetting_(wantedSettingId);
      const it = folder.getFiles();
      while (it.hasNext()) {
        const f = it.next();
        files.push({
          id:f.getId(), name:f.getName(), url:f.getUrl(), folderName,
          size:f.getSize(), sizeText:_formatDriveFileSize_(f.getSize()),
          updatedAt:Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
          mimeType:f.getMimeType(), taskId:task.id
        });
      }
      files.sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)));
      return {success:true,files,folderName,folderUrl:folder.getUrl()};
    }

    // Tidak ada folder yang dipilih: tampilkan seluruh file materi dari semua
    // folder tugas yang sudah ada. Jangan membuat folder baru hanya untuk membaca data.
    const root = _getTugasTrainingFolder_();
    const folders = root.getFolders();
    while (folders.hasNext()) {
      const folder = folders.next();
      const folderName = folder.getName();
      const it = folder.getFiles();
      while (it.hasNext()) {
        const f = it.next();
        files.push({
          id:f.getId(), name:f.getName(), url:f.getUrl(), folderName,
          size:f.getSize(), sizeText:_formatDriveFileSize_(f.getSize()),
          updatedAt:Utilities.formatDate(f.getLastUpdated(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
          mimeType:f.getMimeType(), taskId:''
        });
      }
    }
    files.sort((a,b)=>String(b.updatedAt).localeCompare(String(a.updatedAt)));
    return {success:true,files,folderName:'',folderUrl:root.getUrl()};
  } catch(err){ return {success:false,message:'Gagal memuat file materi tugas: '+err.message}; }
}

function _formatDriveFileSize_(bytes){
  const n=Number(bytes)||0;
  if(n<1024) return `${n} B`;
  if(n<1024*1024) return `${(n/1024).toFixed(1)} KB`;
  if(n<1024*1024*1024) return `${(n/1024/1024).toFixed(2)} MB`;
  return `${(n/1024/1024/1024).toFixed(2)} GB`;
}

function deleteTaskMaterialFile(username, fileId) {
  _authorizePemakai_(username);
  try {
    const id=String(fileId||'').trim();
    if(!id) return {success:false,message:'ID file tidak valid.'};
    const file=DriveApp.getFileById(id);
    file.setTrashed(true);
    _logTaskDriveDelete_(username,file.getId(),file.getName(),'File','pemakai');
    return {success:true,message:'File materi berhasil dihapus.'};
  } catch(err){ return {success:false,message:'Gagal menghapus file materi: '+err.message}; }
}

function deleteTaskMaterialFolder(username, settingId) {
  _authorizePemakai_(username);
  try {
    const {folder,folderName}=_ensureTaskFolderBySetting_(String(settingId||'').trim());
    folder.setTrashed(true);
    _logTaskDriveDelete_(username,folder.getId(),folder.getName(),'Folder','pemakai');
    return {success:true,message:'Folder materi berhasil dihapus.',folderName};
  } catch(err){ return {success:false,message:'Gagal menghapus folder materi: '+err.message}; }
}

/**
 * ===== TRASH / DELETE GOOGLE DRIVE UNTUK MENU TUGAS =====
 * Menampilkan seluruh item yang berada di Sampah Google Drive akun yang
 * menjalankan Apps Script ini, baik file maupun folder.
 */
function _ensureTaskDeleteAuditSheet_() {
  const ss = _ss_();
  let sh = ss.getSheetByName('LogDeleteDrive');
  if (!sh) sh = ss.insertSheet('LogDeleteDrive');
  if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,7).setValues([['Timestamp','Item ID','Nama Item','Jenis','Username','Nama Lengkap','Aksi']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _findPemakaiDisplayName_(username) {
  try {
    const rec = _records_(APP.SHEETS.PEMAKAI);
    if (!rec.success) return String(username || '').trim();
    const h = rec.headers || [];
    const iu = _headerIndex_(h, ['Username','User Name']);
    const inama = _headerIndex_(h, ['Nama Lengkap','Nama','Full Name']);
    const row = rec.rows.find(r => iu >= 0 && _norm_(r.values[iu]) === _norm_(username));
    return row && inama >= 0 ? String(row.values[inama] || '').trim() : String(username || '').trim();
  } catch (e) {
    return String(username || '').trim();
  }
}

function _findPesertaDisplayName_(nik) {
  try {
    const name = _findParticipantName_(nik);
    return String(name || '').trim() || String(nik || '').trim();
  } catch (e) {
    return String(nik || '').trim();
  }
}

function _logTaskDriveDelete_(username, itemId, itemName, itemType, userType) {
  try {
    const sh = _ensureTaskDeleteAuditSheet_();
    const actor = String(username || '').trim();
    const fullName = String(userType || '').toLowerCase() === 'peserta'
      ? _findPesertaDisplayName_(actor)
      : _findPemakaiDisplayName_(actor);
    sh.appendRow([new Date(), String(itemId || '').trim(), String(itemName || '').trim(), String(itemType || '').trim(), actor, fullName, 'Delete']);
  } catch (e) {
    console.warn('Gagal mencatat audit delete Drive: ' + e.message);
  }
}

function _getTaskDeleteAuditMap_() {
  const map = {};
  try {
    const sh = _ensureTaskDeleteAuditSheet_();
    const last = sh.getLastRow();
    if (last < 2) return map;
    const vals = sh.getRange(2,1,last-1,7).getValues();
    vals.forEach(r => {
      const id = String(r[1] || '').trim();
      if (!id) return;
      const ts = r[0] instanceof Date ? r[0].getTime() : new Date(r[0]).getTime();
      if (!map[id] || ts >= map[id].ts) {
        map[id] = {ts:Number.isFinite(ts)?ts:0, fullName:String(r[5] || r[4] || '').trim()};
      }
    });
  } catch (e) {
    console.warn('Gagal membaca audit delete Drive: ' + e.message);
  }
  return map;
}

function _getTaskAdminTrashItems_() {
  const items = [];
  const fileIt = DriveApp.searchFiles("trashed = true");
  while (fileIt.hasNext()) {
    const f = fileIt.next();
    items.push({
      id:f.getId(),
      name:f.getName(),
      type:'File',
      updatedAt:Utilities.formatDate(f.getLastUpdated(), APP.TIMEZONE || Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    });
  }

  const folderIt = DriveApp.searchFolders("trashed = true");
  while (folderIt.hasNext()) {
    const f = folderIt.next();
    items.push({
      id:f.getId(),
      name:f.getName(),
      type:'Folder',
      updatedAt:Utilities.formatDate(f.getLastUpdated(), APP.TIMEZONE || Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    });
  }

  const auditMap = _getTaskDeleteAuditMap_();
  items.forEach(function(item){ item.deletedBy = auditMap[item.id] ? auditMap[item.id].fullName : ''; });

  items.sort(function(a,b){
    const ta = new Date(a.updatedAt.split('/').reverse().join('-').replace(' ','T'));
    const tb = new Date(b.updatedAt.split('/').reverse().join('-').replace(' ','T'));
    return tb.getTime() - ta.getTime();
  });
  return items;
}

function _authorizeProgramer_(username) {
  const result = _authorizePemakai_(username);
  const h = _records_(APP.SHEETS.PEMAKAI).headers || [];
  const row = result.values || [];
  const iRole = _headerIndex_(h, ['Role','Peran']);
  const role = String(iRole >= 0 ? row[iRole] : '').trim().toLowerCase();
  if (role !== 'programer') {
    throw new Error('Akses ditolak. Menu Delete hanya tersedia untuk user dengan Role Programer.');
  }
  return result;
}

function getTaskAdminTrash(username) {
  _authorizeProgramer_(username);
  try {
    return {success:true, items:_getTaskAdminTrashItems_()};
  } catch (err) {
    return {success:false, message:'Gagal memuat Sampah Google Drive: ' + err.message};
  }
}

function _parseReportDocumentName_(fileName) {
  const raw = String(fileName || '').replace(/\.pdf$/i,'').trim();
  const m = raw.match(/^(.*)-Batch\s+(.+)$/i);
  if (!m) return null;
  const prefix = String(m[1] || '').trim();
  const batch = String(m[2] || '').trim();
  if (!prefix || !batch) return null;
  const sep = prefix.indexOf('-');
  if (sep < 0) return null;
  const company = prefix.slice(0, sep).trim();
  const agenda = prefix.slice(sep + 1).trim();
  if (!company || !agenda) return null;
  return {company, agenda, batch:`Batch ${batch}`};
}

function _ensureReportDeleteArchiveSheet_() {
  const ss = _ss_();
  let sh = ss.getSheetByName('Arsip Delete Report');
  const headers = ['Timestamp Delete','File ID','Nama Perusahaan','Nama Agenda','Batch','Link Dokumen','Tanggal Upload','Nama Pengupload','Username Penghapus'];
  if (!sh) {
    sh = ss.insertSheet('Arsip Delete Report');
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() < 1) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function _findDeletedReportArchive_(fileId) {
  try {
    const sh = _ensureReportDeleteArchiveSheet_();
    const last = sh.getLastRow();
    if (last < 2) return null;
    const vals = sh.getRange(2,1,last-1,9).getDisplayValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (String(vals[i][1] || '').trim() === String(fileId || '').trim()) {
        return {
          company:String(vals[i][2] || '').trim(),
          agenda:String(vals[i][3] || '').trim(),
          batch:String(vals[i][4] || '').trim(),
          link:String(vals[i][5] || '').trim(),
          uploadDate:String(vals[i][6] || '').trim(),
          fullName:String(vals[i][7] || '').trim(),
          deletedBy:String(vals[i][8] || '').trim()
        };
      }
    }
  } catch (e) {
    console.warn('Gagal membaca arsip delete report: ' + e.message);
  }
  return null;
}

function _archiveReportBeforeDelete_(rowValues, fileId, username) {
  try {
    const sh = _ensureReportDeleteArchiveSheet_();
    const fullName = String(rowValues[5] || '').trim() || _findPemakaiDisplayName_(username);
    sh.appendRow([
      new Date(), String(fileId || '').trim(), String(rowValues[0] || '').trim(),
      String(rowValues[1] || '').trim(), String(rowValues[2] || '').trim(),
      String(rowValues[3] || '').trim(), String(rowValues[4] || '').trim(),
      fullName, String(username || '').trim()
    ]);
  } catch (e) {
    console.warn('Gagal mengarsipkan data report sebelum delete: ' + e.message);
  }
}

function _restoreReportRekapEntry_(file) {
  try {
    if (!file) return {isReport:false, restored:false};
    const name = String(file.getName() || '').trim();
    const parsed = _parseReportDocumentName_(name);
    if (!parsed) return {isReport:false, restored:false};

    const sh = _getReportDocumentSheet_();
    const lastRow = sh.getLastRow();
    const fileUrl = file.getUrl();
    if (lastRow >= 2) {
      const values = sh.getRange(2,1,lastRow-1,6).getDisplayValues();
      const exists = values.some(r => {
        const existingId = _extractDriveFileId_(r[3]);
        return existingId && existingId === file.getId();
      });
      if (exists) return {isReport:true, restored:false, duplicate:true, company:parsed.company, agenda:parsed.agenda, batch:parsed.batch};
    }

    const archived = _findDeletedReportArchive_(file.getId());
    let fullName = String(archived?.fullName || '').trim();
    let uploadedAt = String(archived?.uploadDate || '').trim();
    const company = String(archived?.company || parsed.company).trim();
    const agenda = String(archived?.agenda || parsed.agenda).trim();
    const batch = String(archived?.batch || parsed.batch).trim();

    // Fallback untuk dokumen lama yang belum pernah masuk arsip delete report.
    if (!fullName) {
      try {
        const auditMap = _getTaskDeleteAuditMap_();
        fullName = String(auditMap[file.getId()]?.fullName || '').trim();
      } catch (e) {}
    }
    if (!uploadedAt) {
      try {
        const created = file.getDateCreated();
        uploadedAt = Utilities.formatDate(created, APP.TIMEZONE || Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      } catch (e) {
        uploadedAt = Utilities.formatDate(new Date(), APP.TIMEZONE || Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
      }
    }
    if (!fullName) fullName = '-';

    sh.appendRow([company, agenda, batch, fileUrl, uploadedAt, fullName]);
    _clearReportRekapCache_();
    SpreadsheetApp.flush();
    return {isReport:true, restored:true, company, agenda, batch, fullName, uploadDate:uploadedAt};
  } catch (e) {
    console.warn('Gagal mengembalikan dokumen ke Rekap Dokumen: ' + e.message);
    return {isReport:false, restored:false, error:e.message};
  }
}

function restoreTaskTrashItem(username, fileId) {
  _authorizeProgramer_(username);
  try {
    const id = String(fileId || '').trim();
    if (!id) return {success:false, message:'ID item tidak valid.'};

    // Coba sebagai file terlebih dahulu, lalu sebagai folder.
    try {
      const file = DriveApp.getFileById(id);
      file.setTrashed(false);
      _setDriveAnyoneEditor_(file);
      const rekap = _restoreReportRekapEntry_(file);
      if (rekap.isReport && rekap.restored) {
        return {success:true, message:'Dokumen report berhasil dipulihkan dan dikembalikan ke Rekap Dokumen.', isReport:true, rekapRestored:true, company:rekap.company, agenda:rekap.agenda, batch:rekap.batch};
      }
      if (rekap.isReport && rekap.duplicate) {
        return {success:true, message:'Dokumen report berhasil dipulihkan. Data Rekap Dokumen sudah tersedia.', isReport:true, rekapRestored:false, duplicate:true};
      }
      return {success:true, message:'File berhasil dipulihkan dari Sampah Google Drive.'};
    } catch (fileErr) {}

    const folder = DriveApp.getFolderById(id);
    folder.setTrashed(false);
    _setDriveAnyoneEditor_(folder);
    return {success:true, message:'Folder berhasil dipulihkan dari Sampah Google Drive.'};
  } catch (err) {
    return {success:false, message:'Gagal memulihkan item: ' + err.message};
  }
}

function permanentlyDeleteTaskTrashItem(username, fileId) {
  _authorizeProgramer_(username);
  try {
    const id = String(fileId || '').trim();
    if (!id) return {success:false, message:'ID item tidak valid.'};

    // DriveApp hanya mendukung memindahkan ke Sampah. Untuk penghapusan
    // permanen, gunakan Advanced Drive Service bila tersedia; jika tidak,
    // fallback ke Drive API v3 melalui OAuth token Apps Script.
    if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.remove === 'function') {
      Drive.Files.remove(id);
      return {success:true, message:'Item berhasil dihapus permanen dari Google Drive.'};
    }

    const token = ScriptApp.getOAuthToken();
    const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?supportsAllDrives=true';
    const response = UrlFetchApp.fetch(url, {
      method:'delete',
      headers:{Authorization:'Bearer ' + token},
      muteHttpExceptions:true
    });
    const code = response.getResponseCode();
    if (code !== 204) {
      let detail = response.getContentText() || '';
      try {
        const obj = JSON.parse(detail);
        detail = obj?.error?.message || detail;
      } catch (e) {}
      throw new Error('Drive API mengembalikan HTTP ' + code + (detail ? ': ' + detail : '.'));
    }
    return {success:true, message:'Item berhasil dihapus permanen dari Google Drive.'};
  } catch (err) {
    return {success:false, message:'Gagal menghapus item secara permanen: ' + err.message};
  }
}

function getTaskSubmissionFolders(username) {
  // Kompatibilitas: fungsi tetap tersedia untuk versi frontend baru.
  // Data folder dibangun dari fungsi hasil pengumpulan yang sudah ada.
  _authorizePemakai_(username);
  try {
    const sh=_sheet_(APP.SHEETS.PENGUMPULAN_TUGAS);
    if(!sh || sh.getLastRow()<2) return {success:true,folders:[]};
    const vals=sh.getRange(1,1,sh.getLastRow(),Math.max(10,sh.getLastColumn())).getDisplayValues();
    const h=vals[0]||[];
    const iTime=_headerIndex_(h,['Timestamp','Waktu','Tanggal']);
    const iUrl=_headerIndex_(h,['URL File','URL']);
    const folders=new Map();

    // Bentuk folder dari taskId agar konsisten dengan folder yang dibuat saat submit.
    const tasksSheet=_ensureTaskSettingSheet_();
    const taskMap=new Map();
    if(tasksSheet && tasksSheet.getLastRow()>=2){
      const tv=tasksSheet.getRange(2,1,tasksSheet.getLastRow()-1,6).getDisplayValues();
      tv.forEach(r=>{
        const id=String(r[0]||'').trim();
        if(id) taskMap.set(id,{company:String(r[1]||'').trim(),training:String(r[2]||'').trim(),batch:String(r[3]||'').trim()});
      });
    }

    const sanitize=v=>String(v||'').replace(/[\\/:*?"<>|#%{}]/g,'_').replace(/\s+/g,' ').trim().slice(0,120)||'Tugas';
    const batchLabel=v=>{const b=String(v||'').trim();return /^batch\b/i.test(b)?b:(b?`Batch ${b}`:'Batch');};

    for(let row=1;row<vals.length;row++){
      const r=vals[row]||[];
      const url=iUrl>=0?String(r[iUrl]||'').trim():'';
      if(!url) continue;
      // URL file menyimpan ID Drive; parent folder aktual akan dibaca langsung dari Drive.
      const m=url.match(/[-\w]{20,}/);
      let folder=null;
      if(m){
        try{
          const file=DriveApp.getFileById(m[0]);
          const parents=file.getParents();
          if(parents.hasNext()) folder=parents.next();
        }catch(e){}
      }
      if(!folder){
        const taskId=String(r[4]||'').trim();
        const t=taskMap.get(taskId);
        if(!t) continue;
        const nm=sanitize(`${sanitize(t.training||'Training')} – ${sanitize(t.company||'Nama Perusahaan')} – ${batchLabel(t.batch)}`);
        const rootIt=DriveApp.getFoldersByName('Pengumpulan Tugas');
        if(!rootIt.hasNext()) continue;
        const root=rootIt.next();
        const fit=root.getFoldersByName(nm);
        if(!fit.hasNext()) continue;
        folder=fit.next();
      }
      const key=folder.getId();
      const current=folders.get(key)||{id:key,name:folder.getName(),url:folder.getUrl(),fileCount:0,latest:null};
      current.fileCount++;
      const d=iTime>=0?new Date(r[iTime]):null;
      if(d && !isNaN(d.getTime()) && (!current.latest || d.getTime()>current.latest.getTime())) current.latest=d;
      folders.set(key,current);
    }

    const tz=APP.TIMEZONE || Session.getScriptTimeZone();
    const out=[...folders.values()].map(f=>({
      id:f.id,name:f.name,url:f.url,fileCount:f.fileCount,
      updatedAt:f.latest?Utilities.formatDate(f.latest,tz,'dd/MM/yyyy HH:mm'):''
    }));
    out.sort((a,b)=>{
      const da=a.updatedAt?new Date(a.updatedAt.split('/').reverse().join('-')):new Date(0);
      const db=b.updatedAt?new Date(b.updatedAt.split('/').reverse().join('-')):new Date(0);
      return db.getTime()-da.getTime() || a.name.localeCompare(b.name,undefined,{sensitivity:'base'});
    });
    return {success:true,folders:out};
  } catch(err){ return {success:false,message:'Gagal memuat folder hasil pengumpulan tugas: '+err.message}; }
}

function downloadTaskSubmissionFolder(username, folderId) {
  _authorizePemakai_(username);
  try {
    const id=String(folderId||'').trim();
    if(!id) return {success:false,message:'Folder hasil pengumpulan tidak ditemukan.'};
    const rootIt=DriveApp.getFoldersByName('Pengumpulan Tugas');
    if(!rootIt.hasNext()) return {success:false,message:'Folder Pengumpulan Tugas belum tersedia.'};
    const root=rootIt.next();
    const folder=DriveApp.getFolderById(id);
    // Pastikan folder benar-benar berada di bawah root Pengumpulan Tugas.
    let isChild=false;
    const parents=folder.getParents();
    while(parents.hasNext()){
      if(parents.next().getId()===root.getId()){isChild=true;break;}
    }
    if(!isChild) return {success:false,message:'Folder tidak valid untuk didownload.'};

    const files=[];
    const it=folder.getFiles();
    while(it.hasNext()){
      const f=it.next();
      if(!f.isTrashed()) files.push(f);
    }
    if(!files.length) return {success:false,message:'Folder belum memiliki file hasil pengumpulan.'};

    const blobs=files.map(f=>f.getBlob().setName(f.getName()));
    const zipName=folder.getName()+' - Hasil Pengumpulan.zip';
    const zipBlob=Utilities.zip(blobs,zipName);
    const tempIt=root.getFoldersByName('_TEMP_DOWNLOAD');
    const tempFolder=tempIt.hasNext()?tempIt.next():root.createFolder('_TEMP_DOWNLOAD');
    _setDriveAnyoneEditor_(tempFolder);
    const zipFile=tempFolder.createFile(zipBlob);
    _setDriveAnyoneEditor_(zipFile);
    zipFile.setDescription('ZIP sementara hasil download folder tugas: '+folder.getName());
    return {success:true,downloadUrl:zipFile.getDownloadUrl(),fileName:zipFile.getName()};
  } catch(err){
    return {success:false,message:'Gagal menyiapkan download folder: '+err.message};
  }
}

function getTaskSubmissionResults(username) {
  _authorizePemakai_(username);
  try {
    const sh=_sheet_(APP.SHEETS.PENGUMPULAN_TUGAS);
    if(!sh || sh.getLastRow()<2) return {success:true,rows:[]};
    const vals=sh.getRange(1,1,sh.getLastRow(),Math.max(10,sh.getLastColumn())).getDisplayValues();
    const h=vals[0]||[];
    const iId=_headerIndex_(h,['ID','ID Submission']);
    const iTime=_headerIndex_(h,['Timestamp','Waktu','Tanggal']);
    const iNik=_headerIndex_(h,['NIK']);
    const iName=_headerIndex_(h,['Nama Peserta','Nama']);
    const iTask=_headerIndex_(h,['ID Tugas','ID Task']);
    const iTaskName=_headerIndex_(h,['Nama Tugas','Tugas']);
    const iNote=_headerIndex_(h,['Keterangan','Catatan']);
    const iFile=_headerIndex_(h,['Nama File','File']);
    const iUrl=_headerIndex_(h,['URL File','URL']);
    const iStatus=_headerIndex_(h,['Status']);
    const rows=[];
    for(let i=1;i<vals.length;i++){
      const r=vals[i]||[];
      if(r.every(v=>String(v||'').trim()==='')) continue;
      rows.push({
        id:iId>=0?String(r[iId]||''):'', uploadedAt:iTime>=0?String(r[iTime]||''):'',
        nik:iNik>=0?String(r[iNik]||''):'', participant:iName>=0?String(r[iName]||''):'',
        taskId:iTask>=0?String(r[iTask]||''):'', taskName:iTaskName>=0?String(r[iTaskName]||''):'',
        note:iNote>=0?String(r[iNote]||''):'', fileName:iFile>=0?String(r[iFile]||''):'',
        link:iUrl>=0?String(r[iUrl]||''):'', status:iStatus>=0?String(r[iStatus]||''):'Dikirim'
      });
    }
    rows.reverse();
    return {success:true,rows};
  } catch(err){ return {success:false,message:'Gagal memuat hasil pengumpulan tugas: '+err.message}; }
}

function getTaskSubmissionMonitoring(username) {
  _authorizePemakai_(username);
  try {
    const settings = _ensureTaskSettingSheet_();
    const lastSettingRow = settings.getLastRow();
    if (lastSettingRow < 2) {
      return {success:true, rows:[], summary:{activeTasks:0,required:0,submitted:0,pending:0}};
    }

    const settingValues = settings.getRange(1,1,lastSettingRow,6).getDisplayValues();
    const shData = _sheet_(APP.SHEETS.DATA);
    if (!shData || shData.getLastRow() < 2) {
      return {success:true, rows:[], summary:{activeTasks:0,required:0,submitted:0,pending:0}};
    }

    const data = shData.getRange(1,1,shData.getLastRow(),Math.max(1,shData.getLastColumn())).getDisplayValues();
    const dh = data[0] || [];
    const iNik = _headerIndex_(dh,['NIK']);
    const iName = _headerIndex_(dh,['Nama Peserta','Nama']);
    const iCompany = _headerIndex_(dh,['Nama Perusahaan','Perusahaan','Company']);
    const iTraining = _headerIndex_(dh,['Jenis Training','Nama Training','Training']);
    const iBatch = _headerIndex_(dh,['Batch']);
    if(iNik<0 || iName<0 || iCompany<0 || iTraining<0 || iBatch<0){
      return {success:false,message:'Kolom NIK, Nama Peserta, Nama Perusahaan, Jenis Training, atau Batch pada sheet Data tidak lengkap.'};
    }

    // Peserta unik berdasarkan NIK + kombinasi training/perusahaan/batch.
    const participantsByCombo = new Map();
    for(let i=1;i<data.length;i++){
      const r=data[i]||[];
      const nik=String(r[iNik]||'').trim();
      const name=String(r[iName]||'').trim();
      const company=String(r[iCompany]||'').trim();
      const training=String(r[iTraining]||'').trim();
      const batch=String(r[iBatch]||'').trim();
      if(!nik||!name||!company||!training||!batch) continue;
      const combo=[_norm_(training),_norm_(company),_normalizeBatch_(batch)].join('|');
      if(!participantsByCombo.has(combo)) participantsByCombo.set(combo,new Map());
      const m=participantsByCombo.get(combo);
      const key=_norm_(nik);
      if(!m.has(key)) m.set(key,{nik,name,company,training,batch});
    }

    // Baca semua submission sekali agar tidak melakukan getRange berulang.
    const peng = _sheet_(APP.SHEETS.PENGUMPULAN_TUGAS);
    const submissionMap = new Map();
    if(peng && peng.getLastRow()>=2){
      const vals=peng.getRange(1,1,peng.getLastRow(),Math.max(10,peng.getLastColumn())).getDisplayValues();
      const ph=vals[0]||[];
      const pNik=_headerIndex_(ph,['NIK']);
      const pTask=_headerIndex_(ph,['ID Tugas','ID']);
      const pTime=_headerIndex_(ph,['Timestamp','Waktu','Tanggal']);
      const pFile=_headerIndex_(ph,['Nama File','File']);
      const pUrl=_headerIndex_(ph,['URL File','URL']);
      const pStatus=_headerIndex_(ph,['Status']);
      if(pNik>=0 && pTask>=0){
        for(let i=1;i<vals.length;i++){
          const r=vals[i]||[];
          const nk=_norm_(r[pNik]);
          const tk=String(r[pTask]||'').trim();
          if(!nk||!tk) continue;
          const key=nk+'|'+tk;
          let ms=0;
          if(pTime>=0){ const d=new Date(r[pTime]); ms=isNaN(d.getTime())?0:d.getTime(); }
          const old=submissionMap.get(key);
          if(!old || ms>=old.ms){
            submissionMap.set(key,{ms,submittedAt:pTime>=0?String(r[pTime]||''):'',fileName:pFile>=0?String(r[pFile]||''):'',url:pUrl>=0?String(r[pUrl]||''):'',status:pStatus>=0?String(r[pStatus]||'Dikirim'):'Dikirim'});
          }
        }
      }
    }

    const rows=[];
    const activeSettingIds=new Set();
    const inactiveSettingIds=new Set();
    for(let rowIndex=1;rowIndex<settingValues.length;rowIndex++){
      const r=settingValues[rowIndex]||[];
      let id=String(r[0]||'').trim();
      const company=String(r[1]||'').trim();
      const training=String(r[2]||'').trim();
      const batch=String(r[3]||'').trim();
      const date=String(r[4]||'').trim();
      const time=String(r[5]||'').trim();
      if(!company||!training||!batch) continue;

      // Pantauan menampilkan SEMUA setting tugas, baik masih aktif maupun
      // sudah Non Aktif karena deadline berakhir.
      const taskActive=_isTaskActive_(date,time);
      const taskStatus=taskActive?'Aktif':'Non Aktif';

      // Jaga agar setting yang lama tetap memiliki ID yang stabil.
      if(!id){
        id='SET-'+Utilities.getUuid().slice(0,8).toUpperCase();
        settings.getRange(rowIndex+1,1).setValue(id);
      }
      if(taskActive) activeSettingIds.add(id);
      else inactiveSettingIds.add(id);
      const deadline=_taskDisplayDeadline_(date,time);
      const taskName=[training,company,batch].filter(Boolean).join(' - ');
      const combo=[_norm_(training),_norm_(company),_normalizeBatch_(batch)].join('|');
      const participants=participantsByCombo.get(combo);
      if(!participants) continue;

      participants.forEach(p=>{
        const key=_norm_(p.nik)+'|'+id;
        const sub=submissionMap.get(key);
        const submitted=!!sub;
        // Tugas yang sudah Non Aktif dan belum pernah dikirim tidak lagi berstatus
        // "Belum Mengirim" karena kesempatan pengirimannya sudah berakhir.
        // Riwayat pengiriman yang sudah berhasil tetap ditampilkan sebagai "Sudah Mengirim".
        const submissionStatus=submitted
          ? 'Sudah Mengirim'
          : (taskActive ? 'Belum Mengirim' : 'Tidak Mengirimkan');
        rows.push({
          taskId:id,
          taskName,
          taskStatus,
          company,
          training,
          batch,
          participantNik:p.nik,
          participantName:p.name,
          deadline,
          submissionStatus,
          submittedAt:sub?.submittedAt||'',
          fileName:sub?.fileName||'',
          url:sub?.url||'',
          status:sub?.status||''
        });
      });
    }

    rows.sort((a,b)=>{
      const t=String(a.training).localeCompare(String(b.training),undefined,{numeric:true,sensitivity:'base'});
      if(t) return t;
      const c=String(a.company).localeCompare(String(b.company),undefined,{numeric:true,sensitivity:'base'});
      if(c) return c;
      const ba=String(a.batch),bb=String(b.batch);
      const bc=ba.localeCompare(bb,undefined,{numeric:true,sensitivity:'base'});
      if(bc) return bc;
      return String(a.participantName).localeCompare(String(b.participantName),undefined,{sensitivity:'base'});
    });

    const submitted=rows.filter(r=>r.submissionStatus==='Sudah Mengirim').length;
    const notSending=rows.filter(r=>r.submissionStatus==='Tidak Mengirimkan').length;
    const pending=rows.filter(r=>r.submissionStatus==='Belum Mengirim').length;
    return {
      success:true,
      rows,
      summary:{
        activeTasks:activeSettingIds.size,
        inactiveTasks:inactiveSettingIds.size,
        totalTasks:activeSettingIds.size + inactiveSettingIds.size,
        required:rows.length,
        submitted,
        pending,
        notSending
      }
    };
  } catch(err){
    return {success:false,message:'Gagal memuat pantauan pengiriman tugas: '+err.message};
  }
}

function getDatabaseOptions(username) {
  _authorizePemakai_(username);
  try { _syncDataPesertaFromData_(); } catch (e) {}

  const result = {};
  [APP.SHEETS.PESERTA, APP.SHEETS.PERUSAHAAN, APP.SHEETS.DATA].forEach(sheetName => {
    const records = _records_(sheetName);
    if (records.success) result[sheetName] = records;
  });
  return result;
}

function _getDriveEffectiveUserEmail_() {
  try { return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase(); } catch (e) { return ''; }
}

function _isDriveFolderManagedByExecutionUser_(folder) {
  if (!folder) return false;
  const effectiveEmail = _getDriveEffectiveUserEmail_();
  if (!effectiveEmail) return false;
  try {
    const owner = folder.getOwner();
    const ownerEmail = String(owner && owner.getEmail ? owner.getEmail() : '').trim().toLowerCase();
    return !!(ownerEmail && ownerEmail === effectiveEmail);
  } catch (e) {
    return false;
  }
}

function _createManagedReportFolder_() {
  const folder = DriveApp.createFolder('Dokumen Report');
  _setDriveAnyoneEditor_(folder);
  return folder;
}

function _getReportDocumentFolder_() {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'REPORT_DOCUMENT_FOLDER_ID';
  const folderName = 'Dokumen Report';
  let folderId = String(props.getProperty(propKey) || '').trim();
  let folder = null;

  // Selalu prioritaskan folder yang sudah ada dengan nama yang sama.
  // Ini mencegah upload membuat/terus memakai folder duplikat hanya karena
  // Script Property menunjuk ke folder lain.
  const existing = DriveApp.getFoldersByName(folderName);
  if (existing.hasNext()) folder = existing.next();

  // Jika belum ada folder bernama tersebut, gunakan property folder yang valid.
  if (!folder && folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }

  // Folder baru hanya dibuat apabila benar-benar belum ada.
  if (!folder) folder = _createManagedReportFolder_();

  props.setProperty(propKey, folder.getId());
  _setDriveAnyoneEditor_(folder);
  return folder;
}

/**
 * Migrasikan dokumen report lama ke folder yang dikelola akun eksekusi aplikasi.
 * Jalankan sekali setelah deployment baru. File lama yang dapat dibaca akan
 * disalin ke storage baru, link pada sheet diperbarui, lalu file baru diberi
 * akses Anyone + Editor. File sumber lama sengaja tidak dihapus karena mungkin
 * dimiliki akun lain/shared drive.
 */
function migrateReportDocumentsToManagedStorage() {
  try {
    const targetFolder = _getReportDocumentFolder_();
    const sh = _getReportDocumentSheet_();
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return {success:true, migrated:0, skipped:0, message:'Tidak ada dokumen report lama yang perlu dimigrasikan.'};

    const values = sh.getRange(2,1,lastRow-1,6).getDisplayValues();
    let migrated = 0, skipped = 0;

    values.forEach((r, idx) => {
      const rowNumber = idx + 2;
      const link = String(r[3] || '').trim();
      const fileId = _extractDriveFileId_(link);
      if (!fileId) { skipped++; return; }

      try {
        const source = DriveApp.getFileById(fileId);
        let alreadyInTarget = false;
        try {
          const parents = source.getParents();
          while (parents.hasNext()) {
            if (parents.next().getId() === targetFolder.getId()) { alreadyInTarget = true; break; }
          }
        } catch (parentErr) {}
        if (alreadyInTarget) {
          _setDriveAnyoneEditor_(source);
        } else {
          const target = source.makeCopy(source.getName(), targetFolder);
          _setDriveAnyoneEditor_(target);
          sh.getRange(rowNumber,4).setValue(target.getUrl());
          migrated++;
        }
      } catch (e) {
        skipped++;
        console.warn('Gagal migrasi row ' + rowNumber + ': ' + e.message);
      }
    });

    SpreadsheetApp.flush();
    return {success:true,migrated,skipped,targetFolderId:targetFolder.getId(),targetFolderUrl:targetFolder.getUrl(),message:'Migrasi dokumen report selesai.'};
  } catch (err) {
    return {success:false,message:'Migrasi dokumen report gagal: ' + err.message};
  }
}

function _getReportDocumentSheet_() {
  const ss = _ss_();
  let sh = ss.getSheetByName(APP.SHEETS.DOKUMEN_REPORT);
  if (!sh) sh = ss.insertSheet(APP.SHEETS.DOKUMEN_REPORT);
  const headers = ['Nama Perusahaan','Nama Agenda','Batch','Link Dokumen','Tanggal Upload','Nama Lengkap'];
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    const existing = sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn())).getDisplayValues()[0];
    const normalized = existing.map(_norm_);
    headers.forEach((h,i) => {
      if (normalized[i] !== _norm_(h)) sh.getRange(1,i+1).setValue(h);
    });
  }
  return sh;
}

function _extractDriveFileId_(urlOrId) {
  const raw = String(urlOrId || '').trim();
  if (!raw) return '';
  if (/^[A-Za-z0-9_-]{20,}$/.test(raw)) return raw;
  const m = raw.match(/(?:\/d\/|id=|[?&]id=)([A-Za-z0-9_-]{20,})/);
  return m ? m[1] : '';
}

function _sanitizeReportNamePart_(value, label) {
  const s = String(value || '').replace(/[\/:*?"<>|]+/g,'_').replace(/\s+/g,' ').trim();
  if (!s) throw new Error(label + ' tidak boleh kosong.');
  return s;
}

function _reportRekapCacheKey_() { return 'REPORT_REKAP_ROWS_V2'; }
function _clearReportRekapCache_() { try { CacheService.getScriptCache().remove(_reportRekapCacheKey_()); } catch (e) {} }
function _readReportRekapRowsFromSheet_() {
  const sh = _getReportDocumentSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2,1,lastRow-1,6).getDisplayValues();
  return values.map((r,i) => ({rowNumber:i+2,company:r[0]||'',agenda:r[1]||'',batch:r[2]||'',link:r[3]||'',uploadDate:r[4]||'',fullName:r[5]||''}))
    .filter(r => r.company || r.agenda || r.batch || r.link || r.uploadDate || r.fullName).reverse();
}

function getReportDocuments(username) {
  try {
    _authorizePemakai_(username);
    const cache = CacheService.getScriptCache();
    const cached = cache.get(_reportRekapCacheKey_());
    let rows = null;
    if (cached) { try { rows = JSON.parse(cached); } catch (e) {} }
    if (!Array.isArray(rows)) {
      rows = _readReportRekapRowsFromSheet_();
      try { cache.put(_reportRekapCacheKey_(), JSON.stringify(rows), 30); } catch (e) {}
    }
    return {success:true,headers:['Nama Perusahaan','Nama Agenda','Batch','Link Dokumen','Tanggal Upload','Nama Lengkap'],rows};
  } catch (err) {
    return {success:false,message:'Gagal memuat rekap dokumen: ' + err.message};
  }
}

function renameReportDocument(username, rowNumber, company, agenda, batch) {
  try {
    _authorizePemakai_(username);
    const sh = _getReportDocumentSheet_();
    const row = Number(rowNumber);
    if (!Number.isInteger(row) || row < 2 || row > sh.getLastRow()) throw new Error('Baris dokumen tidak valid.');
    const current = sh.getRange(row,1,1,6).getDisplayValues()[0];
    const currentLink = String(current[3] || '').trim();
    if (!currentLink) throw new Error('Link dokumen tidak ditemukan pada rekap.');

    const safeCompany = _sanitizeReportNamePart_(company, 'Nama Perusahaan');
    const safeAgenda = _sanitizeReportNamePart_(agenda, 'Nama Agenda');
    let safeBatch = String(batch || '').replace(/\s+/g,' ').trim();
    if (!safeBatch) throw new Error('Batch tidak boleh kosong.');
    if (!/^Batch\s+/i.test(safeBatch)) safeBatch = 'Batch ' + safeBatch;
    safeBatch = _sanitizeReportNamePart_(safeBatch, 'Batch');

    const fileId = _extractDriveFileId_(currentLink);
    if (!fileId) throw new Error('ID file Google Drive tidak dapat dikenali dari link dokumen.');
    const file = DriveApp.getFileById(fileId);
    const newBaseName = safeCompany + '-' + safeAgenda + '-' + safeBatch;
    const newName = newBaseName.replace(/\.pdf$/i,'') + '.pdf';
    file.setName(newName);

    sh.getRange(row,1,1,3).setValues([[safeCompany,safeAgenda,safeBatch]]);
    _clearReportRekapCache_();
    SpreadsheetApp.flush();
    return {success:true,message:'Nama dokumen berhasil diubah.',rowNumber:row,fileId:fileId,fileName:newName,company:safeCompany,agenda:safeAgenda,batch:safeBatch,link:file.getUrl()};
  } catch (err) {
    return {success:false,message:'Gagal mengubah nama dokumen: ' + err.message};
  }
}

function _trashReportDriveFile_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return {success:true, method:'none'};

  // Penghapusan/trash HARUS dilakukan oleh identitas yang memang mempunyai
  // hak menghapus file. Status Anyone + Editor tidak mengubah kepemilikan file.
  // Karena itu semua percobaan tetap menggunakan akun eksekusi Apps Script.
  let errors = [];

  // Prioritas 1: Advanced Drive Service.
  if (typeof Drive !== 'undefined' && Drive.Files && typeof Drive.Files.update === 'function') {
    try {
      Drive.Files.update({trashed:true}, id);
      return {success:true, method:'Advanced Drive Service'};
    } catch (e) {
      errors.push('Advanced Drive Service: ' + (e?.message || e));
    }
  }

  // Prioritas 2: Drive API v3 dengan OAuth token Apps Script.
  try {
    const token = ScriptApp.getOAuthToken();
    const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(id) + '?supportsAllDrives=true';
    const response = UrlFetchApp.fetch(url, {
      method:'patch',
      contentType:'application/json',
      payload:JSON.stringify({trashed:true}),
      headers:{Authorization:'Bearer ' + token},
      muteHttpExceptions:true
    });
    const code = response.getResponseCode();
    if (code >= 200 && code < 300) return {success:true, method:'Drive API v3'};
    let detail = '';
    try { detail = JSON.parse(response.getContentText() || '{}')?.error?.message || ''; } catch (_) {}
    errors.push('Drive API v3: ' + (detail || 'HTTP ' + code));
  } catch (e) {
    errors.push('Drive API v3: ' + (e?.message || e));
  }

  // Prioritas 3: DriveApp sebagai fallback terakhir.
  try {
    DriveApp.getFileById(id).setTrashed(true);
    return {success:true, method:'DriveApp'};
  } catch (e) {
    errors.push('DriveApp: ' + (e?.message || e));
  }

  throw new Error(
    'Akun eksekusi aplikasi tidak memiliki hak untuk menghapus file ini. ' +
    'Pastikan Web App menggunakan Execute as: Me dan file dimiliki/berada pada Drive yang dapat dikelola akun tersebut. ' +
    'Anyone + Editor hanya mengatur akses berbagi, bukan hak kepemilikan/hapus. Detail: ' + errors.join(' | ')
  );
}

function deleteReportDocument(username, rowNumber) {
  try {
    _authorizePemakai_(username);
    const sh = _getReportDocumentSheet_();
    const row = Number(rowNumber);
    if (!Number.isInteger(row) || row < 2 || row > sh.getLastRow()) throw new Error('Baris dokumen tidak valid.');
    const current = sh.getRange(row,1,1,6).getDisplayValues()[0];
    const link = String(current[3] || '').trim();
    const fileId = _extractDriveFileId_(link);
    if (!fileId) throw new Error('Link dokumen tidak ditemukan atau tidak valid pada rekap.');

    // Simpan metadata asli sebelum baris Rekap dihapus, sehingga Restore
    // dapat mengembalikan Nama Pengupload dan tanggal upload secara akurat.
    _archiveReportBeforeDelete_(current, fileId, username);

    // Ambil nama file dan catat siapa yang benar-benar menghapusnya.
    // Nama ini dipakai oleh submenu Delete melalui LogDeleteDrive.
    let deletedItemName = String(current[3] || '').trim() || 'Dokumen Report';
    try { deletedItemName = DriveApp.getFileById(fileId).getName() || deletedItemName; } catch (e) {}

    _trashReportDriveFile_(fileId);
    _logTaskDriveDelete_(username, fileId, deletedItemName, 'File', 'pemakai');
    const deletedBy = _findPemakaiDisplayName_(username);

    sh.deleteRow(row);
    _clearReportRekapCache_();
    return {success:true,message:'Dokumen report berhasil dihapus dari Drive dan rekap.',rowNumber:row,fileId:fileId,deletedBy};
  } catch (err) {
    return {success:false,message:'Gagal menghapus dokumen report: ' + err.message};
  }
}

function uploadReportDocument(username, fileName, dataUrl, mimeType) {
  try {
    const user = _authorizePemakai_(username);
    let name = String(fileName || '').trim();
    const mime = String(mimeType || '').trim().toLowerCase();
    const raw = String(dataUrl || '');
    if (!name) return {success:false, message:'Nama file PDF tidak boleh kosong.'};
    if (!/\.pdf$/i.test(name) || mime !== 'application/pdf' || !/^data:application\/pdf;base64,/i.test(raw)) {
      return {success:false, message:'File yang dapat diupload hanya PDF (.pdf).'};
    }
    // Nama report wajib mengikuti pola: Perusahaan-Jenis Agenda/Training-Batch X.pdf
    name = name.replace(/\.pdf$/i,'').trim();
    if (!/^.+-.+-Batch\s+.+$/i.test(name)) {
      return {success:false, message:'Nama file report harus mengikuti format Perusahaan-Jenis Agenda/Training-Batch X.pdf.'};
    }
    const base64 = raw.replace(/^data:application\/pdf;base64,/i,'');
    const bytes = Utilities.base64Decode(base64);
    const MAX_BYTES = 20 * 1024 * 1024;
    if (bytes.length > MAX_BYTES) return {success:false, message:'Ukuran dokumen PDF maksimal 20 MB.'};
    if (bytes.length < 5 || bytes[0] !== 0x25 || bytes[1] !== 0x50 || bytes[2] !== 0x44 || bytes[3] !== 0x46) {
      return {success:false, message:'Isi file tidak dikenali sebagai dokumen PDF yang valid.'};
    }
    const safeName = name.replace(/[\\/:*?"<>|]+/g,'_').replace(/\s+/g,' ').trim() || 'Dokumen_Report.pdf';
    const finalName = /\.pdf$/i.test(safeName) ? safeName : safeName + '.pdf';
    const folder = _getReportDocumentFolder_();
    const blob = Utilities.newBlob(bytes, 'application/pdf', finalName);
    const file = folder.createFile(blob);
    _setDriveAnyoneEditor_(file);
    file.setDescription('Dokumen Report yang diupload melalui Training Management System.');

    const reportSheet = _getReportDocumentSheet_();
    const parsedName = _parseReportDocumentName_(name);
    if (!parsedName) return {success:false,message:'Nama file report tidak dapat dibaca. Pastikan format berakhir dengan -Batch <nomor/rentang>.'};
    const company = parsedName.company;
    const agenda = parsedName.agenda;
    const batch = parsedName.batch;
    const h = _records_(APP.SHEETS.PEMAKAI).headers || [];
    const idxName = _headerIndex_(h, ['Nama Lengkap','Nama Pemakai','Nama User','Nama']);
    const fullName = idxName >= 0 ? String(user.values[idxName] || '').trim() : String(username || '').trim();
    const tz = APP.TIMEZONE || Session.getScriptTimeZone();
    const uploadedAt = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
    reportSheet.appendRow([company, agenda, batch, file.getUrl(), uploadedAt, fullName]);
    _clearReportRekapCache_();
    return {success:true, message:'Dokumen report berhasil diupload.', fileId:file.getId(), fileName:file.getName(), fileUrl:file.getUrl(), company, agenda, batch, uploadDate:uploadedAt, fullName};
  } catch (err) {
    return {success:false, message:'Gagal mengupload dokumen report: ' + err.message};
  }
}