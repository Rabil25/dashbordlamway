// ================================================================
//  WOK Lampung Barat / Way Kanan – Dashboard API (Google Apps Script)
//  Sheet : "Monitoring Daily HH LPG"
//  Deploy: Extensions → Apps Script → Deploy → New deployment
//          Type      : Web App
//          Execute as: Me
//          Who can access: Anyone
// ================================================================

const CONFIG = {
  SHEET_ID: '1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM',
  TOKEN:    'lpgLampung2026',
  TIMEZONE: 'Asia/Jakarta',
  TOKEN_TTL_MS: 24 * 60 * 60 * 1000,  // 24 jam
};

// ── Index kolom dalam range CE2:DT (0-based dari CE=0) ──────────
const IDX_BLN_IO = 6;   // CK
const IDX_BLN_RE = 9;   // CN
const IDX_BLN_PS = 12;  // CQ

// ── Cluster / Wilayah config ──────────────────────────────────
const WILAYAH = {
  'Lampung Barat': {
    kabupaten: ['LAMPUNG BARAT', 'TANGGAMUS', 'PESISIR BARAT'],
  },
  'Way Kanan': {
    kabupaten: ['LAMPUNG UTARA', 'WAY KANAN'],
  },
};

// ── Users sheet columns (0-based) ────────────────────────────
// A=0: username, B=1: password_hash, C=2: role, D=3: status,
// E=4: nama, F=5: tanggal_daftar, G=6: session_token, H=7: token_expiry
const U = { USERNAME:0, HASH:1, ROLE:2, STATUS:3, NAMA:4, TGL_DAFTAR:5, TOKEN:6, EXPIRY:7 };
const USERS_SHEET = 'Users';

// ── Cache TTL (detik) ─────────────────────────────────────────
const CACHE_TTL_DATA    = 600;   // 10 menit – bulan_ini & update_sf
const CACHE_TTL_BULAN   = 3600;  // 1 jam    – list_bulan

// ================================================================
//  Entry Point GET
// ================================================================
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const type   = (params.type || '').toLowerCase();

  // ── Auth endpoints tidak butuh CONFIG.TOKEN ───────────────
  const AUTH_TYPES = ['auth_login','auth_register','auth_verify',
                      'auth_change_password','auth_pending_users',
                      'auth_approve','auth_logout'];
  if (AUTH_TYPES.indexOf(type) !== -1) {
    try {
      return jsonOut(handleAuth(type, params));
    } catch(err) {
      return jsonOut({ ok:false, error: err.message });
    }
  }

  // ── Data endpoints butuh CONFIG.TOKEN ─────────────────────
  if (params.token !== CONFIG.TOKEN) {
    return jsonOut({ ok:false, error:'Unauthorized' });
  }

  try {
    const bulan = (params.bulan || '').toString().trim();
    let data;

    switch (type) {
      case 'bulan_ini':
        data = getRawSheetCached('Bulan_Ini', 'CE2:DT', bulan);
        break;
      case 'list_bulan':
        data = getListBulanCached();
        break;
      case 'update_sf':
        data = getRawSheetCached('Update SF', 'A2:AG', '');
        break;
      case 'wilayah':
        data = { ok:true, wilayah: WILAYAH };
        break;
      case 'ping':
        data = { ok:true, message:'pong', updatedAt: new Date().toISOString() };
        break;
      // ── BARU: Gabungkan 3 request jadi 1 ─────────────────
      case 'all':
        data = getAllData(bulan);
        break;
      default:
        data = { ok:false, error:'Unknown type. Use: bulan_ini | list_bulan | update_sf | wilayah | ping | all' };
    }

    return jsonOut(data);
  } catch (err) {
    return jsonOut({ ok:false, error: err.message });
  }
}

// ================================================================
//  BARU: Combined endpoint – kembalikan semua data dalam 1 request
//  Parameter: bulan (bulan aktif, misal "8" untuk Agustus)
// ================================================================
function getAllData(bulan) {
  // Hitung bulan lalu
  const bulanNum = parseInt(bulan) || 0;
  const bulanLalu = bulanNum > 1 ? (bulanNum - 1).toString() : '';

  // Ambil semua data (masing-masing sudah pakai cache)
  const bulanIni  = getRawSheetCached('Bulan_Ini', 'CE2:DT', bulan);
  const bulanPrev = bulanLalu ? getRawSheetCached('Bulan_Ini', 'CE2:DT', bulanLalu) : { ok:true, data:[], rows:0 };
  const updateSf  = getRawSheetCached('Update SF', 'A2:AG', '');
  const listBulan = getListBulanCached();

  return {
    ok:        true,
    bulan:     bulan,
    bulanLalu: bulanLalu,
    bulanIni:  bulanIni,
    bulanPrev: bulanPrev,
    updateSf:  updateSf,
    listBulan: listBulan,
    updatedAt: new Date().toISOString(),
  };
}

// ================================================================
//  BARU: Warmup – mencegah GAS cold start
//  Pasang time-based trigger setiap 5 menit via setupWarmupTrigger()
// ================================================================
function warmupSelf() {
  // Hanya buka spreadsheet agar GAS instance tetap hangat
  try {
    SpreadsheetApp.openById(CONFIG.SHEET_ID);
    Logger.log('Warmup OK: ' + new Date().toISOString());
  } catch(e) {
    Logger.log('Warmup error: ' + e.message);
  }
}

// Jalankan fungsi ini SEKALI dari Editor GAS untuk memasang trigger warmup
// Extensions → Apps Script → Run → setupWarmupTrigger
function setupWarmupTrigger() {
  // Hapus trigger warmup lama jika ada
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'warmupSelf') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Pasang trigger baru setiap 5 menit
  ScriptApp.newTrigger('warmupSelf')
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log('Warmup trigger terpasang: setiap 5 menit');
}

// Hapus trigger warmup (jalankan jika ingin menonaktifkan)
function removeWarmupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'warmupSelf') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  Logger.log('Trigger warmup dihapus: ' + removed + ' trigger');
}

// ================================================================
//  AUTH HANDLERS
// ================================================================
function handleAuth(type, params) {
  switch(type) {

    // ── Login: username + password_hash (SHA-256 dari client) ──
    case 'auth_login': {
      const username = (params.username||'').trim().toLowerCase();
      const hash     = (params.hash||'').trim();
      if (!username || !hash) return { ok:false, error:'Username dan password wajib diisi.' };

      const row = findUser(username);
      if (!row) return { ok:false, error:'Username tidak ditemukan.' };
      if (row[U.HASH] !== hash) return { ok:false, error:'Password salah.' };
      if (row[U.STATUS] === 'pending') return { ok:false, error:'Akun kamu belum disetujui admin. Silakan tunggu.' };
      if (row[U.STATUS] === 'rejected') return { ok:false, error:'Akun kamu ditolak oleh admin.' };
      if (row[U.STATUS] !== 'active') return { ok:false, error:'Status akun tidak valid.' };

      const token  = Utilities.getUuid();
      const expiry = Date.now() + CONFIG.TOKEN_TTL_MS;
      updateUserField(username, U.TOKEN,  token);
      updateUserField(username, U.EXPIRY, expiry.toString());

      return {
        ok:    true,
        token: token,
        user:  { username: username, nama: row[U.NAMA], role: row[U.ROLE] },
      };
    }

    // ── Register: username + hash + nama ──────────────────────
    case 'auth_register': {
      const username = (params.username||'').trim().toLowerCase();
      const hash     = (params.hash||'').trim();
      const nama     = (params.nama||'').trim();
      if (!username || !hash || !nama) return { ok:false, error:'Username, password, dan nama wajib diisi.' };
      if (!/^[a-z0-9_]{3,20}$/.test(username)) return { ok:false, error:'Username hanya huruf kecil, angka, underscore (3-20 karakter).' };

      if (findUser(username)) return { ok:false, error:'Username sudah dipakai.' };

      const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      const ws = ss.getSheetByName(USERS_SHEET);
      if (!ws) return { ok:false, error:'Sheet Users tidak ditemukan.' };

      const tgl = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
      ws.appendRow([username, hash, 'user', 'pending', nama, tgl, '', '']);

      return { ok:true, message:'Pendaftaran berhasil. Tunggu persetujuan admin.' };
    }

    // ── Verify session token ───────────────────────────────────
    case 'auth_verify': {
      const token = (params.session_token||'').trim();
      if (!token) return { ok:false, error:'Token tidak ada.' };
      const result = verifyToken(token);
      return result;
    }

    // ── Logout: hapus token ────────────────────────────────────
    case 'auth_logout': {
      const token = (params.session_token||'').trim();
      if (!token) return { ok:true };
      const row = findUserByToken(token);
      if (row) {
        updateUserField(row[U.USERNAME], U.TOKEN,  '');
        updateUserField(row[U.USERNAME], U.EXPIRY, '');
      }
      return { ok:true };
    }

    // ── Change password ────────────────────────────────────────
    case 'auth_change_password': {
      const token   = (params.session_token||'').trim();
      const oldHash = (params.old_hash||'').trim();
      const newHash = (params.new_hash||'').trim();
      if (!token || !oldHash || !newHash) return { ok:false, error:'Parameter tidak lengkap.' };

      const verified = verifyToken(token);
      if (!verified.ok) return verified;

      const row = findUser(verified.user.username);
      if (!row) return { ok:false, error:'User tidak ditemukan.' };
      if (row[U.HASH] !== oldHash) return { ok:false, error:'Password lama salah.' };

      updateUserField(verified.user.username, U.HASH, newHash);
      return { ok:true, message:'Password berhasil diubah.' };
    }

    // ── List pending users (admin only) ───────────────────────
    case 'auth_pending_users': {
      const token = (params.session_token||'').trim();
      const verified = verifyToken(token);
      if (!verified.ok) return verified;
      if (verified.user.role !== 'admin') return { ok:false, error:'Akses ditolak.' };

      const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
      const ws = ss.getSheetByName(USERS_SHEET);
      if (!ws) return { ok:false, error:'Sheet Users tidak ditemukan.' };

      const rows = ws.getDataRange().getValues().slice(1); // skip header
      const filter  = (params.filter||'pending').trim();   // pending | all
      const result  = rows
        .filter(r => filter === 'all' || r[U.STATUS] === 'pending')
        .map(r => ({
          username:    r[U.USERNAME],
          nama:        r[U.NAMA],
          role:        r[U.ROLE],
          status:      r[U.STATUS],
          tgl_daftar:  r[U.TGL_DAFTAR],
        }));

      return { ok:true, users: result };
    }

    // ── Approve / reject user (admin only) ────────────────────
    case 'auth_approve': {
      const token    = (params.session_token||'').trim();
      const target   = (params.target_username||'').trim().toLowerCase();
      const action   = (params.action||'').trim(); // approve | reject
      const verified = verifyToken(token);
      if (!verified.ok) return verified;
      if (verified.user.role !== 'admin') return { ok:false, error:'Akses ditolak.' };
      if (!target || !action) return { ok:false, error:'target_username dan action wajib.' };
      if (action !== 'approve' && action !== 'reject') return { ok:false, error:'Action harus approve atau reject.' };

      const row = findUser(target);
      if (!row) return { ok:false, error:'User tidak ditemukan.' };

      const newStatus = action === 'approve' ? 'active' : 'rejected';
      updateUserField(target, U.STATUS, newStatus);

      return { ok:true, message: 'User ' + target + ' ' + (action === 'approve' ? 'disetujui' : 'ditolak') + '.' };
    }

    default:
      return { ok:false, error:'Auth type tidak dikenal.' };
  }
}

// ── Verifikasi token: cek ada + belum expired ────────────────
function verifyToken(token) {
  const row = findUserByToken(token);
  if (!row) return { ok:false, error:'Sesi tidak valid. Silakan login ulang.' };

  const expiry = parseInt(row[U.EXPIRY]||0);
  if (Date.now() > expiry) {
    updateUserField(row[U.USERNAME], U.TOKEN,  '');
    updateUserField(row[U.USERNAME], U.EXPIRY, '');
    return { ok:false, error:'Sesi kamu sudah habis. Silakan login ulang.' };
  }

  return {
    ok:   true,
    user: { username: row[U.USERNAME], nama: row[U.NAMA], role: row[U.ROLE] },
  };
}

// ── Cari user by username (lowercase) ───────────────────────
function findUser(username) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName(USERS_SHEET);
  if (!ws) return null;
  const rows = ws.getDataRange().getValues().slice(1);
  for (let i=0; i<rows.length; i++) {
    if ((rows[i][U.USERNAME]||'').toString().trim().toLowerCase() === username) {
      return rows[i];
    }
  }
  return null;
}

// ── Cari user by session token ───────────────────────────────
function findUserByToken(token) {
  if (!token) return null;
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName(USERS_SHEET);
  if (!ws) return null;
  const rows = ws.getDataRange().getValues().slice(1);
  for (let i=0; i<rows.length; i++) {
    if ((rows[i][U.TOKEN]||'').toString().trim() === token) return rows[i];
  }
  return null;
}

// ── Update satu field user di sheet ─────────────────────────
function updateUserField(username, colIdx, value) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName(USERS_SHEET);
  if (!ws) return;
  const rows = ws.getDataRange().getValues();
  for (let i=1; i<rows.length; i++) {
    if ((rows[i][U.USERNAME]||'').toString().trim().toLowerCase() === username) {
      ws.getRange(i+1, colIdx+1).setValue(value);
      return;
    }
  }
}

// ================================================================
//  DATA HANDLERS – dengan CacheService
// ================================================================

// ── BARU: getListBulan dengan cache ─────────────────────────
function getListBulanCached() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'list_bulan';

  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      parsed._fromCache = true;
      return parsed;
    }
  } catch(e) {
    // Cache miss atau parse error, lanjut baca sheet
  }

  const result = getListBulan();

  if (result.ok) {
    try {
      cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_BULAN);
    } catch(e) {
      // Cache put gagal (data terlalu besar dll), tidak masalah
    }
  }

  return result;
}

// ── BARU: getRawSheet dengan cache ───────────────────────────
function getRawSheetCached(sheetName, range, filterBulan) {
  const cache = CacheService.getScriptCache();
  // Key cache unik per sheet + bulan
  const cacheKey = 'sheet_' + sheetName.replace(/\s/g,'_') + '_' + (filterBulan||'all');

  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      parsed._fromCache = true;
      return parsed;
    }
  } catch(e) {
    // Cache miss atau parse error, lanjut baca sheet
  }

  const result = getRawSheet(sheetName, range, filterBulan);

  if (result.ok) {
    try {
      // GAS CacheService max value size: 100KB per entry
      const serialized = JSON.stringify(result);
      if (serialized.length < 90000) {  // beri margin
        cache.put(cacheKey, serialized, CACHE_TTL_DATA);
      }
    } catch(e) {
      // Data terlalu besar untuk cache, tidak masalah
    }
  }

  return result;
}

// ── getListBulan (tanpa cache) ────────────────────────────────
function getListBulan() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName('Bulan_Ini');
  if (!ws) return { ok:false, error:'Sheet "Bulan_Ini" tidak ditemukan.' };

  const values = ws.getRange('CK2:CQ').getValues();
  const bulanSet = new Set();
  for (let i=0; i<values.length; i++) {
    const row = values[i];
    if (row[0]!==''&&row[0]!==null) bulanSet.add(row[0].toString().trim());
    if (row[3]!==''&&row[3]!==null) bulanSet.add(row[3].toString().trim());
    if (row[6]!==''&&row[6]!==null) bulanSet.add(row[6].toString().trim());
  }
  const bulanList = Array.from(bulanSet)
    .filter(function(b){ return /^\d+$/.test(b)&&Number(b)>=1&&Number(b)<=12; })
    .map(Number)
    .sort(function(a,b){ return a-b; });

  return { ok:true, bulan:bulanList, updatedAt: new Date().toISOString() };
}

// ── getRawSheet (tanpa cache) ─────────────────────────────────
function getRawSheet(sheetName, range, filterBulan) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return { ok:false, error:'Sheet "'+sheetName+'" tidak ditemukan.' };

  let values;
  try {
    values = ws.getRange(range).getValues();
  } catch(e) {
    return { ok:false, error:'Gagal baca range '+range+': '+e.message };
  }

  const isBulanIni = (sheetName==='Bulan_Ini' && filterBulan!=='');
  const cleaned = [];
  for (let i=0; i<values.length; i++) {
    const row = values[i];
    const hasData = row.some(function(cell){ return cell!==''&&cell!==null&&cell!==undefined; });
    if (!hasData) continue;

    if (isBulanIni) {
      const blnIO = row[IDX_BLN_IO]!==undefined&&row[IDX_BLN_IO]!=='' ? row[IDX_BLN_IO].toString().trim() : '';
      const blnRE = row[IDX_BLN_RE]!==undefined&&row[IDX_BLN_RE]!=='' ? row[IDX_BLN_RE].toString().trim() : '';
      const blnPS = row[IDX_BLN_PS]!==undefined&&row[IDX_BLN_PS]!=='' ? row[IDX_BLN_PS].toString().trim() : '';
      if (blnIO!==filterBulan && blnRE!==filterBulan && blnPS!==filterBulan) continue;
    }

    const cleanRow = row.map(function(cell){
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      }
      return cell;
    });
    cleaned.push(cleanRow);
  }

  return {
    ok:        true,
    sheet:     sheetName,
    range:     range,
    bulan:     filterBulan||'all',
    rows:      cleaned.length,
    data:      cleaned,
    updatedAt: new Date().toISOString(),
  };
}

// ── Helper: output JSON ───────────────────────────────────────
function jsonOut(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
