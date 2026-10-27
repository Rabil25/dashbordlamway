// ================================================================
//  IndiHome LPG – Dashboard API (Google Apps Script)
//  Deploy: Extensions → Apps Script → Deploy → New deployment
//          Execute as: Me | Who has access: Anyone
// ================================================================

const CONFIG = {
  SHEET_ID: '1csOC-Uat1y4D1mXY1Y6uM0oSeUzG2iOxyoh81I92QQY',
  TOKEN:    'lpgLampung2026',
  TIMEZONE: 'Asia/Jakarta',
};

// ── KONFIGURASI WILAYAH — single source of truth ─────────────
// Cluster ditentukan oleh kabupaten order (VALIDASI KAB), bukan STO.
// STO adalah entitas teknis jaringan, bisa lintas cluster.
const WILAYAH = {
  'Lampung Barat': {
    tl:        'ANDRES WELSEN',
    kabupaten: ['LAMPUNG BARAT', 'TANGGAMUS', 'PESISIR BARAT'],
  },
  'Way Kanan': {
    tl:        'AGUS SETIAWAN',
    kabupaten: ['LAMPUNG UTARA', 'WAY KANAN'],
  },
};

// Mapping STO → Cluster (untuk fallback saat KAB tidak tersedia)
const STO_TO_CLUSTER = {
  'TGP': 'Lampung Barat',
  'KOT': 'Lampung Barat',
  'LIW': 'Lampung Barat',
  'KRU': 'Lampung Barat',
  'BKM': 'Way Kanan',
  'KBI': 'Way Kanan',
  'BDU': 'Way Kanan',
};

// Auto-build reverse lookup
const KAB_TO_CLUSTER = {};
const TL_TO_CLUSTER  = {};
Object.entries(WILAYAH).forEach(([cluster, cfg]) => {
  cfg.kabupaten.forEach(k => { KAB_TO_CLUSTER[k] = cluster; });
  TL_TO_CLUSTER[cfg.tl] = cluster;
});

// ── Entry Point (GET — data read-only) ───────────────────────
function doGet(e) {
  const params = e.parameter || {};
  if (params.token !== CONFIG.TOKEN) {
    return jsonOut({ ok: false, error: 'Unauthorized' });
  }
  try {
    const action = params.action || 'performance';
    let data;
    switch (action) {
      case 'performance':   data = getPerformance(params);           break;
      case 'sf':            data = getSFList();                      break;
      case 'absen':         data = getAbsen();                       break;
      case 'kendala':       data = getKendala();                     break;
      case 'wilayah':       data = { wilayah: WILAYAH };             break;
      case 'orders':        data = getOrders(params);                break;
      case 'getUserList':   data = getUserList(params);              break;
      case 'getPendingCount': data = getPendingCount(params);        break;
      case 'getTargets':    data = getTargetsGAS();                  break;
      case 'getMultipliers':data = getMultipliers();                 break;
      // Legacy GET support — migrasi ke POST disarankan
      case 'login':         data = loginUser(params);                break;
      case 'register':      data = registerUser(params);             break;
      case 'changepass':    data = changePassword(params);           break;
      default:              data = { error: 'Unknown action' };
    }
    return jsonOut({ ok: true, ...data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ── Entry Point (POST — auth & mutasi data) ───────────────────
function doPost(e) {
  let params = {};
  try {
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter || {};
    }
  } catch(parseErr) {
    params = e.parameter || {};
  }

  if (params.token !== CONFIG.TOKEN) {
    return jsonOut({ ok: false, error: 'Unauthorized' });
  }

  try {
    const action = params.action || '';
    let data;
    switch (action) {
      case 'login':          data = loginUser(params);           break;
      case 'register':       data = registerUser(params);        break;
      case 'changepass':     data = changePassword(params);      break;
      case 'getUserList':    data = getUserList(params);         break;
      case 'approveUser':    data = approveUser(params);         break;
      case 'rejectUser':     data = rejectUser(params);          break;
      case 'updateRole':     data = updateUserRole(params);      break;
      case 'deleteUser':     data = deleteUser(params);          break;
      case 'saveTargets':    data = saveTargets(params);         break;
      case 'saveMultipliers':data = saveMultipliers(params);     break;
      default:               data = { error: 'Unknown action' };
    }
    return jsonOut({ ok: true, ...data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function jsonOut(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

function ss() { return SpreadsheetApp.openById(CONFIG.SHEET_ID); }

function fmtDate(d) {
  return Utilities.formatDate(new Date(d), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

// ── Hitung Hari berdasarkan rentang filter (semua hari kalender) ──
function calcHariRange(startStr, endStr) {
  // startStr / endStr: "YYYYMMDD"
  const y  = parseInt(startStr.substring(0, 4));
  const m  = parseInt(startStr.substring(4, 6));
  const d1 = new Date(y, m - 1, parseInt(startStr.substring(6, 8)));
  const d2 = new Date(parseInt(endStr.substring(0, 4)),
                      parseInt(endStr.substring(4, 6)) - 1,
                      parseInt(endStr.substring(6, 8)));
  const elapsed  = Math.round((d2 - d1) / 86400000) + 1;  // inklusif
  const total    = new Date(y, m, 0).getDate();            // total hari di bulan
  return { total, elapsed, month: m, year: y };
}

// ── Date range untuk filter (MoM apple-to-apple) ─────────────
// Returns { start: "YYYYMM01", end: "YYYYMMDD" } untuk bulan tsb s/d tanggal yg sama
function getDateRange(now, offsetMonths) {
  const tz  = CONFIG.TIMEZONE;
  const y   = parseInt(Utilities.formatDate(now, tz, 'yyyy'));
  const m   = parseInt(Utilities.formatDate(now, tz, 'MM'));
  const d   = parseInt(Utilities.formatDate(now, tz, 'dd'));
  let yr = y, mo = m + offsetMonths;
  if (mo <= 0) { mo += 12; yr--; }
  if (mo > 12) { mo -= 12; yr++; }
  // Clamp day to last day of target month
  const lastDay = new Date(yr, mo, 0).getDate();
  const endDay  = Math.min(d, lastDay);
  const pad = n => n.toString().padStart(2, '0');
  return {
    start: `${yr}${pad(mo)}01`,
    end:   `${yr}${pad(mo)}${pad(endDay)}`,
    ym:    `${yr}${pad(mo)}`,
  };
}

// ── Akun Web — Login & Registrasi ───────────────────────────
// Sheet "Akun_Web" kolom: A=Timestamp, B=Username, C=Nama, D=PasswordHash, E=Role, F=Status

function getAkunSheet() {
  const ws = ss().getSheetByName('Akun_Web');
  if (!ws) throw new Error('Sheet Akun_Web tidak ditemukan. Buat sheet dengan kolom: Timestamp | Username | Nama | PasswordHash | Role | Status');
  return ws;
}

// ── Rate Limiting (PropertiesService) ────────────────────────
// key: 'rl_login_<username>'
// value: JSON { attempts: N, firstAt: timestamp }
const RATE_LIMIT_MAX  = 5;
const RATE_LIMIT_SECS = 15 * 60; // 15 menit

function checkRateLimit(key) {
  const ps  = PropertiesService.getScriptProperties();
  const raw = ps.getProperty(key);
  if (!raw) return true; // belum ada → boleh
  try {
    const data = JSON.parse(raw);
    const now  = Date.now();
    if (now - data.firstAt > RATE_LIMIT_SECS * 1000) {
      ps.deleteProperty(key);
      return true; // window expired
    }
    return data.attempts < RATE_LIMIT_MAX;
  } catch(e) { return true; }
}

function recordFailedAttempt(key) {
  const ps  = PropertiesService.getScriptProperties();
  const raw = ps.getProperty(key);
  const now = Date.now();
  let data  = { attempts: 0, firstAt: now };
  try { if (raw) data = JSON.parse(raw); } catch(e) {}
  if (now - data.firstAt > RATE_LIMIT_SECS * 1000) {
    data = { attempts: 0, firstAt: now };
  }
  data.attempts++;
  ps.setProperty(key, JSON.stringify(data));
}

function resetRateLimit(key) {
  PropertiesService.getScriptProperties().deleteProperty(key);
}

// ── Verifikasi Admin ──────────────────────────────────────────
function verifyAdmin(params) {
  const username = (params.adminUsername || params.username || '').toString().trim().toLowerCase();
  if (!username) return false;
  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if ((r[1]||'').toString().trim().toLowerCase() !== username) continue;
    const role   = (r[4] || '').toString().trim().toLowerCase();
    const status = (r[5] || '').toString().trim().toLowerCase();
    return role === 'admin' && status === 'active';
  }
  return false;
}

function loginUser(params) {
  const username = (params.username || '').toString().trim().toLowerCase();
  const passHash = (params.passHash || '').toString().trim();
  if (!username || !passHash) return { ok: false, error: 'Username / password kosong' };

  // Rate limiting
  const rlKey = 'rl_login_' + username;
  if (!checkRateLimit(rlKey)) {
    return { ok: false, error: 'Terlalu banyak percobaan login. Coba lagi setelah 15 menit.' };
  }

  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const r        = rows[i];
    const rowUser  = (r[1] || '').toString().trim().toLowerCase();
    const rowHash  = (r[3] || '').toString().trim();
    const rowStatus= (r[5] || '').toString().trim().toLowerCase();

    if (rowUser !== username) continue;
    if (rowHash  !== passHash) {
      recordFailedAttempt(rlKey);
      return { ok: false, error: 'Password salah.' };
    }
    if (rowStatus === 'pending') return { ok: false, error: 'Akun sedang menunggu persetujuan admin. Hubungi supervisor.' };
    if (rowStatus !== 'active')  return { ok: false, error: 'Akun dinonaktifkan. Hubungi admin.' };
    resetRateLimit(rlKey);
    return { ok: true, username: rowUser, nama: (r[2] || rowUser).toString().trim(), role: (r[4] || 'viewer').toString().trim() };
  }
  recordFailedAttempt(rlKey);
  return { ok: false, error: 'Username tidak ditemukan.' };
}

function registerUser(params) {
  const nama     = (params.nama     || '').toString().trim();
  const username = (params.username || '').toString().trim().toLowerCase();
  const passHash = (params.passHash || '').toString().trim();

  if (!nama || !username || !passHash) return { ok: false, error: 'Data tidak lengkap.' };
  if (!/^[a-z0-9._-]+$/.test(username)) return { ok: false, error: 'Username tidak valid.' };

  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();

  // Cek duplikat
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][1] || '').toString().trim().toLowerCase() === username) {
      return { ok: false, error: 'Username sudah digunakan.' };
    }
  }

  // Tentukan role: admin kalau belum ada user sama sekali, sisanya viewer
  // Status: admin langsung active, user biasa perlu approval (pending)
  const isFirst = rows.length <= 1;
  const role   = isFirst ? 'admin' : 'viewer';
  const status = isFirst ? 'active' : 'pending';
  ws.appendRow([ new Date(), username, nama, passHash, role, status ]);
  return { ok: true, message: isFirst
    ? 'Akun admin berhasil dibuat.'
    : 'Pendaftaran berhasil! Menunggu persetujuan admin sebelum bisa login.' };
}

function changePassword(params) {
  const username = (params.username || '').toString().trim().toLowerCase();
  const oldHash  = (params.oldHash  || '').toString().trim();
  const newHash  = (params.newHash  || '').toString().trim();
  if (!username || !oldHash || !newHash) return { ok: false, error: 'Data tidak lengkap.' };

  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const rowUser = (rows[i][1] || '').toString().trim().toLowerCase();
    if (rowUser !== username) continue;
    const rowHash = (rows[i][3] || '').toString().trim();
    if (rowHash !== oldHash) return { ok: false, error: 'Password lama salah.' };
    ws.getRange(i + 1, 4).setValue(newHash); // kolom D = index 4
    return { ok: true, message: 'Password berhasil diubah.' };
  }
  return { ok: false, error: 'User tidak ditemukan.' };
}

// ── User Management (admin only) ─────────────────────────────
function getUserList(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[1]) continue;
    users.push({
      row:      i + 1,
      username: (r[1] || '').toString().trim(),
      nama:     (r[2] || '').toString().trim(),
      role:     (r[4] || 'viewer').toString().trim(),
      status:   (r[5] || 'pending').toString().trim(),
    });
  }
  return { users };
}

function getPendingCount(params) {
  if (!verifyAdmin(params)) return { pendingCount: 0 };
  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();
  let count  = 0;
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][5] || '').toString().trim().toLowerCase() === 'pending') count++;
  }
  return { pendingCount: count };
}

function approveUser(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  return _setUserStatus(params.targetUsername, 'active');
}

function rejectUser(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  return _setUserStatus(params.targetUsername, 'rejected');
}

function updateUserRole(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  const target = (params.targetUsername || '').toString().trim().toLowerCase();
  const role   = (params.newRole || 'viewer').toString().trim().toLowerCase();
  if (!['admin','viewer'].includes(role)) return { ok: false, error: 'Role tidak valid.' };
  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][1]||'').toString().trim().toLowerCase() === target) {
      ws.getRange(i + 1, 5).setValue(role); // kolom E
      return { ok: true, message: `Role ${target} diubah ke ${role}.` };
    }
  }
  return { ok: false, error: 'User tidak ditemukan.' };
}

function deleteUser(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  const target = (params.targetUsername || '').toString().trim().toLowerCase();
  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][1]||'').toString().trim().toLowerCase() === target) {
      ws.deleteRow(i + 1);
      return { ok: true, message: `User ${target} berhasil dihapus.` };
    }
  }
  return { ok: false, error: 'User tidak ditemukan.' };
}

function _setUserStatus(targetUsername, newStatus) {
  const target = (targetUsername || '').toString().trim().toLowerCase();
  const ws   = getAkunSheet();
  const rows = ws.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][1]||'').toString().trim().toLowerCase() === target) {
      ws.getRange(i + 1, 6).setValue(newStatus); // kolom F
      return { ok: true, message: `Status ${target} diubah ke ${newStatus}.` };
    }
  }
  return { ok: false, error: 'User tidak ditemukan.' };
}

// ── Targets & Multipliers (PropertiesService) ─────────────────
function getTargetsGAS() {
  const ps  = PropertiesService.getScriptProperties();
  const raw = ps.getProperty('wok_targets');
  let targets = {};
  try { if (raw) targets = JSON.parse(raw); } catch(e) {}
  return { targets };
}

function saveTargets(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  const raw = params.targets;
  if (!raw) return { ok: false, error: 'Data targets kosong.' };
  let targets;
  try { targets = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch(e) {
    return { ok: false, error: 'Format targets tidak valid.' };
  }
  PropertiesService.getScriptProperties().setProperty('wok_targets', JSON.stringify(targets));
  return { ok: true, message: 'Targets berhasil disimpan.' };
}

function getMultipliers() {
  const ps  = PropertiesService.getScriptProperties();
  const raw = ps.getProperty('wok_multipliers');
  let mult  = { io: 1.5, re: 1.2 };
  try { if (raw) mult = { ...mult, ...JSON.parse(raw) }; } catch(e) {}
  return { multipliers: mult };
}

function saveMultipliers(params) {
  if (!verifyAdmin(params)) return { ok: false, error: 'Akses ditolak.' };
  const io = parseFloat(params.io);
  const re = parseFloat(params.re);
  if (isNaN(io) || isNaN(re) || io <= 0 || re <= 0) {
    return { ok: false, error: 'Nilai multiplier tidak valid.' };
  }
  PropertiesService.getScriptProperties().setProperty('wok_multipliers', JSON.stringify({ io, re }));
  return { ok: true, message: 'Multiplier berhasil disimpan.' };
}

// ── Master SF ────────────────────────────────────────────────
function getSFList() {
  const ws   = ss().getSheetByName('Data Sales');
  const rows = ws.getDataRange().getValues();
  const hdr  = rows[0];
  const col  = (name) => hdr.indexOf(name);
  const C = {
    no:     col('No'),
    nama:   col('Nama'),
    kode:   col('Kode'),
    sto:    col('STO'),
    tl:     col('Team Leader'),
    city:   col('CITY'),
    status: col('Status'),
    level:  col('Level SF'),
    agency: col('Agency'),
  };

  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.no]) continue;
    const tlUpper  = (r[C.tl]  || '').toString().toUpperCase().trim();
    const stoUpper = (r[C.sto] || '').toString().toUpperCase().trim();
    list.push({
      nama:    r[C.nama],
      kode:    (r[C.kode] || '').toString().trim(),
      sto:     stoUpper,
      tl:      tlUpper,
      cluster: TL_TO_CLUSTER[tlUpper] || '-',
      city:    r[C.city],
      status:  r[C.status],
      level:   r[C.level],
      agency:  r[C.agency],
    });
  }
  return { sfList: list };
}

// ── Shift a YYYYMMDD string by N months ─────────────────────
function shiftMonthStr(yyyymmdd, offsetMonths) {
  if (!yyyymmdd || yyyymmdd.length < 8) return '';
  const y = parseInt(yyyymmdd.substring(0, 4));
  const m = parseInt(yyyymmdd.substring(4, 6));
  const d = parseInt(yyyymmdd.substring(6, 8));
  let yr = y, mo = m + offsetMonths;
  while (mo <= 0)  { mo += 12; yr--; }
  while (mo > 12)  { mo -= 12; yr++; }
  const lastDay = new Date(yr, mo, 0).getDate();
  const dd = Math.min(d, lastDay);
  const pad = n => n.toString().padStart(2, '0');
  return `${yr}${pad(mo)}${pad(dd)}`;
}

// ── Performance Utama ────────────────────────────────────────
function getPerformance(params) {
  params = params || {};
  const now = new Date();
  const { sfList } = getSFList();
  const sfMap = {};
  sfList.forEach(sf => { sfMap[sf.kode.toLowerCase()] = sf; });

  // Normalize param dates (accept YYYY-MM-DD or YYYYMMDD)
  let startIni = (params.start || '').toString().replace(/-/g, '').substring(0, 8);
  let endIni   = (params.end   || '').toString().replace(/-/g, '').substring(0, 8);

  // If no custom range → default: 1st of current month → today
  if (!startIni || !endIni) {
    const def = getDateRange(now, 0);
    startIni = def.start;
    endIni   = def.end;
  }

  // MoM: same day-range, previous month (untuk perbandingan WOK apples-to-apples)
  const startLalu = shiftMonthStr(startIni, -1);
  const endLalu   = shiftMonthStr(endIni,   -1);

  // Full previous month — khusus untuk Level LM & M-1 SF scorecard
  const prevMonthBase = shiftMonthStr(startIni, -1);
  const prevY = parseInt(prevMonthBase.substring(0, 4));
  const prevM = parseInt(prevMonthBase.substring(4, 6));
  const prevLastDay = new Date(prevY, prevM, 0).getDate();
  const pad2 = n => n.toString().padStart(2, '0');
  const startLaluFull = `${prevY}${pad2(prevM)}01`;
  const endLaluFull   = `${prevY}${pad2(prevM)}${pad2(prevLastDay)}`;

  const rangeIni      = { start: startIni,      end: endIni      };
  const rangeLalu     = { start: startLalu,      end: endLalu     };
  const rangeLaluFull = { start: startLaluFull,  end: endLaluFull };

  const produkFilter  = (params.produk  || '').toString().trim().toLowerCase();
  const channelFilter = (params.channel || '').toString().trim().toLowerCase();
  const kendalaFilter = (params.kendala || '').toString().trim().toLowerCase();

  const stoFilter     = (params.sto     || '').toString().trim().toUpperCase();
  const clusterFilter = (params.cluster || '').toString().trim();
  const kabFilter     = (clusterFilter && WILAYAH[clusterFilter])
                          ? WILAYAH[clusterFilter].kabupaten
                          : [];

  const ini      = parseBulan('Bulan_Ini',  sfMap, rangeIni.start,      rangeIni.end,      produkFilter, channelFilter, stoFilter, kabFilter, kendalaFilter);
  const lalu     = parseBulan('Bulan_Lalu', sfMap, rangeLalu.start,     rangeLalu.end,     produkFilter, channelFilter, stoFilter, kabFilter, kendalaFilter);
  const laluFull = parseBulan('Bulan_Lalu', sfMap, rangeLaluFull.start, rangeLaluFull.end, produkFilter, channelFilter, stoFilter, kabFilter, kendalaFilter);

  const hariKerja = calcHariRange(startIni, endIni);
  const wsa       = calcWSATSA(sfMap);

  // Daily trend helper: map {YYYY-MM-DD: count} → sorted array
  const toTrendArr = (byDate) =>
    Object.entries(byDate || {})
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([date, count]) => ({ date, count }));

  const allKodes = new Set([
    ...Object.keys(ini.bySF),
    ...Object.keys(lalu.bySF),
    ...sfList.filter(s => s.status === 'ACTIVE').map(s => s.kode.toLowerCase()),
  ]);

  const sfPerf = [];
  allKodes.forEach(kode => {
    const info = sfMap[kode] || {};
    const i    = ini.bySF[kode]      || { io: 0, re: 0, ps: 0, kendala: 0, produk: {}, kab: {}, dailyRE: {}, dailyPS: {} };
    const l    = lalu.bySF[kode]     || { io: 0, re: 0, ps: 0, kendala: 0, produk: {} };
    const lf   = laluFull.bySF[kode] || { io: 0, re: 0, ps: 0 };
    const w    = wsa.bySF[kode]  || { bulanIni: 0, bulanLalu: 0 };

    sfPerf.push({
      kode:    kode.toUpperCase(),
      nama:    info.nama    || kode.toUpperCase(),
      sto:     info.sto     || '-',
      tl:      info.tl      || '-',
      cluster: info.cluster || '-',
      status:  info.status  || '-',
      level:   info.level   || '-',
      bulanIni: {
        io:      i.io,
        re:      i.re,
        ps:      i.ps,
        wsa:     w.bulanIni,
        kendala: i.kendala,
        produk:  i.produk,
        kab:     i.kab,
        dailyRE: i.dailyRE || {},
        dailyPS: i.dailyPS || {},
      },
      bulanLalu: {
        io:  l.io,
        re:  l.re,
        ps:  l.ps,
        wsa: w.bulanLalu,
      },
      m1ps: lf.ps,  // full previous month PS — untuk Level LM & growth% di scorecard
    });
  });

  sfPerf.sort((a, b) => b.bulanIni.ps - a.bulanIni.ps);

  return {
    sfPerformance:  sfPerf,
    rekapSTO:       rekapBySTO(sfPerf),
    rekapOrderSTO:  rekapByOrderSTO(ini, lalu),
    rekapCluster:   rekapByCluster(sfPerf),
    rekapKab:       rekapByKabFull(ini, lalu),
    summary: {
      bulanIni:     ini.total,
      bulanLalu:    lalu.total,
      digitalIni:   ini.digital,
      digitalLalu:  lalu.digital,
      agencyIni:    ini.agency,
      agencyLalu:   lalu.agency,
    },
    hariKerja,
    wsaTrend:   wsa.trend,
    ioTrend:    toTrendArr(ini.ioByDate),
    reTrend:    toTrendArr(ini.reByDate),
    psTrend:    toTrendArr(ini.psByDate),
    wilayah:    WILAYAH,
    kendalaList: ini.kendalaList || [],
    updatedAt:  new Date().toISOString(),
  };
}

// ── Parse Sheet dengan filter rentang tanggal (YYYYMMDD) ─────
// startStr & endStr: "YYYYMMDD" – apple-to-apple MoM comparison
function parseBulan(sheetName, sfMap, startStr, endStr, produkFilter, channelFilter, stoFilter, kabFilter, kendalaFilter) {
  produkFilter  = (produkFilter  || '').toLowerCase().trim();
  channelFilter = (channelFilter || '').toLowerCase().trim();
  kendalaFilter = (kendalaFilter || '').toLowerCase().trim();
  stoFilter     = (stoFilter     || '').toUpperCase().trim();
  kabFilter     = Array.isArray(kabFilter) ? kabFilter.map(k => k.toUpperCase().trim()) : [];
  const ws = ss().getSheetByName(sheetName);
  if (!ws) return {
    bySF:    {},
    total:   { io: 0, re: 0, ps: 0, byKab: {} },
    agency:  { io: 0, re: 0, ps: 0, byKab: {} },
    digital: { io: 0, re: 0, ps: 0 },
  };

  const rows = ws.getDataRange().getValues();
  const hdr  = rows[0];

  const findCol = (names) => {
    for (const n of names) {
      const idx = hdr.findIndex(h => h.toString().trim().toLowerCase() === n.trim().toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const C = {
    status:    findCol(['Status Order']),
    io:        findCol(['Validasi IO']),
    tanggalIO: findCol(['Tanggal IO']),
    re:        findCol(['Validasi RE']),
    tanggalRE: findCol(['Tanggal RE']),
    ps:        findCol(['Validasi PS']),
    tanggalPS: findCol(['Tanggal PS']),
    codeSF:    findCol(['CODE SF', 'Validasi SF']),
    kendala:   findCol(['Kategori Kendala']),
    produk:    findCol(['Validasi Produk']),
    kab:       findCol(['VALIDASI KAB']),
    sto:       findCol(['Validasi STO']),
    channel2:  findCol(['Channel 2']),
    kendalaKat:findCol(['Kategori Kendala']),
    kendTekCol:findCol(['Kendala Teknis']),
  };

  function mkStat() { return { io: 0, re: 0, ps: 0, byKab: {}, bySTO: {} }; }
  function addKab(stat, kab, field) {
    if (!kab) return;
    if (!stat.byKab[kab]) stat.byKab[kab] = { io: 0, re: 0, ps: 0 };
    stat.byKab[kab][field]++;
  }
  function addSTO(stat, sto, field) {
    if (!sto) return;
    if (!stat.bySTO[sto]) stat.bySTO[sto] = { io: 0, re: 0, ps: 0 };
    stat.bySTO[sto][field]++;
  }

  const bySF    = {};
  const total   = mkStat();
  const agency  = mkStat();
  const digital = { io: 0, re: 0, ps: 0 };
  const ioByDate = {};
  const reByDate = {};
  const psByDate = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.status]) continue;

    const ch       = (C.channel2 !== -1 ? r[C.channel2] : '').toString().trim().toLowerCase();
    const isAgency = ch === 'agency';

    // Filter channel (agency / digital)
    if (channelFilter === 'agency'  && !isAgency) continue;
    if (channelFilter === 'digital' && isAgency)  continue;

    // Filter produk
    if (produkFilter) {
      const rowProduk = (C.produk !== -1 ? r[C.produk] : '').toString().trim().toLowerCase();
      if (!rowProduk.includes(produkFilter)) continue;
    }

    // Filter STO
    if (stoFilter) {
      const rowSTO = (C.sto !== -1 ? r[C.sto] : '').toString().trim().toUpperCase();
      if (rowSTO !== stoFilter) continue;
    }

    // Filter Kendala Teknis
    if (kendalaFilter) {
      const rowKT = (C.kendTekCol !== -1 ? r[C.kendTekCol] : '').toString().trim().toLowerCase();
      if (!rowKT.includes(kendalaFilter)) continue;
    }

    // Tanggal masing-masing metric (format YYYYMMDD)
    // Handle both text strings ("2026-09-01", "2026/09/01") and Date objects
    const getTgl = (col) => {
      if (col === -1 || r[col] === null || r[col] === undefined || r[col] === '') return '';
      const v = r[col];
      if (v instanceof Date) {
        // Pakai Utilities.formatDate agar timezone Asia/Jakarta dihormati.
        // Raw getDate() di GAS V8 pakai UTC → tanggal bisa geser -1 hari.
        return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyyMMdd');
      }
      return v.toString().replace(/-/g, '').replace(/\//g, '').substring(0, 8);
    };
    const tglIO = getTgl(C.tanggalIO);
    const tglRE = getTgl(C.tanggalRE);
    const tglPS = getTgl(C.tanggalPS);

    // Tiap metric dihitung hanya kalau tanggalnya masuk range filter
    const inRange = (tgl) => !startStr || (tgl && tgl >= startStr && tgl <= endStr);
    const countIO = inRange(tglIO);
    const countRE = inRange(tglRE);
    const countPS = inRange(tglPS);

    // Skip baris kalau tidak ada metric yang masuk range sama sekali
    if (!countIO && !countRE && !countPS) continue;

    const kabRaw = C.kab !== -1 && r[C.kab] ? r[C.kab].toString().trim().toUpperCase() : '';
    const stoRaw = (C.sto !== -1 ? r[C.sto] : '').toString().trim().toUpperCase();

    // Filter Cluster (berdasarkan kabupaten order, bukan STO SF)
    if (kabFilter.length && !kabFilter.includes(kabRaw)) continue;

    const valIO = (C.io !== -1 ? r[C.io] : '').toString().toLowerCase();
    const valRE = (C.re !== -1 ? r[C.re] : '').toString().toLowerCase();
    const valPS = (C.ps !== -1 ? r[C.ps] : '').toString().toLowerCase();

    // Helper: YYYYMMDD → YYYY-MM-DD
    const toDash = (s) => s ? s.substring(0,4)+'-'+s.substring(4,6)+'-'+s.substring(6,8) : '';

    // ── Total semua channel ──
    if (countIO && valIO.includes('valid io')) {
      total.io++;
      addKab(total, kabRaw, 'io');
      addSTO(total, stoRaw, 'io');
      if (!isAgency) digital.io++;
      const d = toDash(tglIO); if (d) ioByDate[d] = (ioByDate[d]||0)+1;
    }
    if (countRE && valRE.includes('valid re')) {
      total.re++;
      addKab(total, kabRaw, 're');
      addSTO(total, stoRaw, 're');
      if (!isAgency) digital.re++;
      const d = toDash(tglRE); if (d) reByDate[d] = (reByDate[d]||0)+1;
    }
    if (countPS && valPS.includes('put in service')) {
      total.ps++;
      addKab(total, kabRaw, 'ps');
      addSTO(total, stoRaw, 'ps');
      if (!isAgency) digital.ps++;
      const d = toDash(tglPS); if (d) psByDate[d] = (psByDate[d]||0)+1;
    }

    // ── Agency only ──
    if (isAgency) {
      if (countIO && valIO.includes('valid io')) { agency.io++; addKab(agency, kabRaw, 'io'); addSTO(agency, stoRaw, 'io'); }
      if (countRE && valRE.includes('valid re')) { agency.re++; addKab(agency, kabRaw, 're'); addSTO(agency, stoRaw, 're'); }
      if (countPS && valPS.includes('put in service')) { agency.ps++; addKab(agency, kabRaw, 'ps'); addSTO(agency, stoRaw, 'ps'); }
    }

    // ── Per-SF (agency only) ──
    if (!isAgency) continue;
    const rawKode = (C.codeSF !== -1 ? r[C.codeSF] : '') || '';
    const kode    = rawKode.toString().trim().toLowerCase();
    if (!kode || kode === '#n/a') continue;

    if (!bySF[kode]) bySF[kode] = { io: 0, re: 0, ps: 0, kendala: 0, produk: {}, kab: {}, dailyRE: {}, dailyPS: {} };

    if (countIO && valIO.includes('valid io')) bySF[kode].io++;
    if (countRE && valRE.includes('valid re')) {
      bySF[kode].re++;
      const dayRE = tglRE ? parseInt(tglRE.substring(6, 8)) : 0;
      if (dayRE) bySF[kode].dailyRE[dayRE] = (bySF[kode].dailyRE[dayRE] || 0) + 1;
    }
    if (countPS && valPS.includes('put in service')) {
      bySF[kode].ps++;
      const dayPS = tglPS ? parseInt(tglPS.substring(6, 8)) : 0;
      if (dayPS) bySF[kode].dailyPS[dayPS] = (bySF[kode].dailyPS[dayPS] || 0) + 1;
      if (kabRaw) bySF[kode].kab[kabRaw] = (bySF[kode].kab[kabRaw] || 0) + 1;
    }
    if (C.kendala !== -1 && r[C.kendala]) bySF[kode].kendala++;
    if (C.produk  !== -1 && r[C.produk]) {
      const p = r[C.produk].toString().trim();
      if (p) bySF[kode].produk[p] = (bySF[kode].produk[p] || 0) + 1;
    }
    if (C.kendalaKat !== -1 && r[C.kendalaKat]) {
      const k = r[C.kendalaKat].toString().trim();
      if (k) bySF[kode].kendalaList = bySF[kode].kendalaList || {};
      if (k) bySF[kode].kendalaList[k] = (bySF[kode].kendalaList[k] || 0) + 1;
    }
  }

  // Kumpulkan unique Kendala Teknis dari Bulan_Ini (bukan Kategori Kendala)
  const kendalaSet = new Set();
  const colKT = C.kendTekCol !== -1 ? C.kendTekCol : C.kendalaKat; // fallback ke Kategori Kendala
  if (colKT !== -1) {
    for (let i = 1; i < rows.length; i++) {
      const v = rows[i][colKT];
      if (v) { const s = v.toString().trim(); if (s) kendalaSet.add(s); }
    }
  }

  return { bySF, total, agency, digital, kendalaList: [...kendalaSet].sort(), ioByDate, reByDate, psByDate };
}

// ── Rekap per Kabupaten (Full: all + agency, ini + lalu) ─────
function rekapByKabFull(ini, lalu) {
  const allKabs = new Set([
    ...Object.keys(ini.total.byKab),
    ...Object.keys(lalu.total.byKab),
  ]);
  const result = [];
  allKabs.forEach(kab => {
    const cluster = KAB_TO_CLUSTER[kab] || 'Lainnya';
    result.push({
      kabupaten: kab,
      cluster,
      allIni:    ini.total.byKab[kab]  || { io: 0, re: 0, ps: 0 },
      allLalu:   lalu.total.byKab[kab] || { io: 0, re: 0, ps: 0 },
      agencyIni:  ini.agency.byKab[kab]  || { io: 0, re: 0, ps: 0 },
      agencyLalu: lalu.agency.byKab[kab] || { io: 0, re: 0, ps: 0 },
    });
  });
  return result.sort((a, b) => b.allIni.ps - a.allIni.ps);
}

// ── Rekap per STO Order (berdasarkan Validasi STO di order) ──
function rekapByOrderSTO(ini, lalu) {
  const allSTOs = new Set([
    ...Object.keys(ini.total.bySTO  || {}),
    ...Object.keys(lalu.total.bySTO || {}),
  ]);
  const result = [];
  allSTOs.forEach(sto => {
    if (!sto) return;
    result.push({
      sto,
      allIni:     ini.total.bySTO[sto]   || { io: 0, re: 0, ps: 0 },
      allLalu:    lalu.total.bySTO[sto]  || { io: 0, re: 0, ps: 0 },
      agencyIni:  ini.agency.bySTO[sto]  || { io: 0, re: 0, ps: 0 },
      agencyLalu: lalu.agency.bySTO[sto] || { io: 0, re: 0, ps: 0 },
    });
  });
  return result.sort((a, b) => b.allIni.ps - a.allIni.ps);
}

// ── Rekap per STO SF ─────────────────────────────────────────
function rekapBySTO(sfPerf) {
  const map = {};
  sfPerf.forEach(sf => {
    const sto = sf.sto || '-';
    if (!map[sto]) map[sto] = {
      sto, cluster: sf.cluster, tl: sf.tl,
      bulanIni:  { io: 0, re: 0, ps: 0, wsa: 0 },
      bulanLalu: { io: 0, re: 0, ps: 0, wsa: 0 },
      sfAktif: 0,
    };
    ['io','re','ps','wsa'].forEach(k => {
      map[sto].bulanIni[k]  += sf.bulanIni[k]  || 0;
      map[sto].bulanLalu[k] += sf.bulanLalu[k] || 0;
    });
    if (sf.status === 'ACTIVE') map[sto].sfAktif++;
  });
  return Object.values(map).sort((a, b) => b.bulanIni.ps - a.bulanIni.ps);
}

// ── Rekap per Cluster ────────────────────────────────────────
function rekapByCluster(sfPerf) {
  const map = {};
  sfPerf.forEach(sf => {
    const c = sf.cluster || '-';
    if (!map[c]) map[c] = {
      cluster: c, tl: sf.tl,
      bulanIni:  { io: 0, re: 0, ps: 0, wsa: 0 },
      bulanLalu: { io: 0, re: 0, ps: 0, wsa: 0 },
      sfAktif: 0,
    };
    ['io','re','ps','wsa'].forEach(k => {
      map[c].bulanIni[k]  += sf.bulanIni[k]  || 0;
      map[c].bulanLalu[k] += sf.bulanLalu[k] || 0;
    });
    if (sf.status === 'ACTIVE') map[c].sfAktif++;
  });
  return Object.values(map);
}

// ── WSA / TSA ────────────────────────────────────────────────
function calcWSATSA(sfMap) {
  const ws   = ss().getSheetByName('WSA_TSA');
  const rows = ws.getDataRange().getValues();

  const now   = new Date();
  const tz    = CONFIG.TIMEZONE;
  const nowYM = Utilities.formatDate(now, tz, 'yyyy-MM');  // e.g. "2026-10"
  const nowY  = parseInt(nowYM.split('-')[0]);
  const nowM  = parseInt(nowYM.split('-')[1]);              // 1-based
  const lastM = nowM === 1 ? 12 : nowM - 1;
  const lastY = nowM === 1 ? nowY - 1 : nowY;
  const lastYM = `${lastY}-${lastM.toString().padStart(2,'0')}`;

  const bySF  = {};
  const trend = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    let d;
    if (r[0] instanceof Date) {
      d = r[0];
    } else {
      try { d = new Date(r[0]); } catch(e) { continue; }
    }
    if (isNaN(d)) continue;

    const dateStr = Utilities.formatDate(d, tz, 'yyyy-MM');  // timezone-aware

    const kode = (r[1] || '').toString().trim().toLowerCase();
    if (!kode) continue;

    const wsa = Number(r[2]) || 0;

    if (!bySF[kode]) bySF[kode] = { bulanIni: 0, bulanLalu: 0 };

    if (dateStr === nowYM)  bySF[kode].bulanIni  += wsa;
    if (dateStr === lastYM) bySF[kode].bulanLalu += wsa;

    const trendKey = dateStr; // already "yyyy-MM" timezone-aware
    if (!trend[trendKey]) trend[trendKey] = 0;
    trend[trendKey] += wsa;
  }

  const trendArr = Object.entries(trend)
    .sort(([a], [b]) => a.localeCompare(b))
    .slice(-12)
    .map(([date, count]) => ({ date, count }));

  return { bySF, trend: trendArr };
}

// ── Orders (raw per-row data dari Bulan_Ini) ─────────────────
// Params: start (YYYYMMDD), end (YYYYMMDD), channel (agency|digital|all)
function getOrders(params) {
  const ws = ss().getSheetByName('Bulan_Ini');
  if (!ws) return { orders: [] };

  const startStr = (params.start || '').toString().replace(/-/g, '').substring(0, 8);
  const endStr   = (params.end   || '').toString().replace(/-/g, '').substring(0, 8);
  const chFilter = (params.channel || '').toString().toLowerCase().trim();

  const rows = ws.getDataRange().getValues();
  const hdr  = rows[0];

  const findCol = (names) => {
    for (const n of names) {
      const idx = hdr.findIndex(h => h.toString().trim().toLowerCase() === n.trim().toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const C = {
    status:    findCol(['Status Order']),
    wok:       findCol(['Validasi WOK']),
    sto:       findCol(['Validasi STO']),
    valIO:     findCol(['Validasi IO']),
    tanggalIO: findCol(['Tanggal IO']),
    valRE:     findCol(['Validasi RE']),
    tanggalRE: findCol(['Tanggal RE']),
    valPS:     findCol(['Validasi PS']),
    tanggalPS: findCol(['Tanggal PS']),
    codeSF:    findCol(['CODE SF', 'Validasi SF']),
    namaSF:    findCol(['Nama SF']),
    channel:   findCol(['Validasi Channel']),
    channel2:  findCol(['Channel 2']),
    reasonK:   findCol(['Reason Kendala']),
    kendTek:   findCol(['Kendala Teknis']),
    ioTimes:   findCol(['IO Times']),
    timeRange: findCol(['Time_Range']),
    odp:       findCol(['ODP2']),
    orderId:   findCol(['ID Order']),
    tglKend:   findCol(['Tanggal Kendala']),
    produk:    findCol(['Validasi Produk']),
    kabupaten: findCol(['VALIDASI KAB']),
    koordinat: findCol(['Koordinat']),
  };

  function cleanDate(val) {
    if (!val) return '';
    if (val instanceof Date) {
      return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    }
    const s = val.toString().replace(/-/g, '').replace(/\//g, '').substring(0, 8);
    if (s.length === 8) return s.substring(0,4)+'-'+s.substring(4,6)+'-'+s.substring(6,8);
    return '';
  }

  const orders = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.status]) continue;

    // Date filter on Tanggal IO
    const tglIO = C.tanggalIO !== -1 ? r[C.tanggalIO] : '';
    const tglStr = tglIO
      ? (tglIO instanceof Date
          ? Utilities.formatDate(tglIO, CONFIG.TIMEZONE, 'yyyyMMdd')
          : tglIO.toString().replace(/-/g, '').replace(/\//g, '').substring(0, 8))
      : '';
    if (startStr && tglStr && tglStr < startStr) continue;
    if (endStr   && tglStr && tglStr > endStr)   continue;
    // Kalau tglStr kosong tapi ada filter tanggal, skip juga (order belum IO)
    if ((startStr || endStr) && !tglStr) continue;

    const ch2  = (C.channel2 !== -1 ? r[C.channel2] : '').toString().trim().toLowerCase();
    const isAg = ch2 === 'agency';

    // Channel filter
    if (chFilter === 'agency'  && !isAg) continue;
    if (chFilter === 'digital' && isAg)  continue;

    const valIO = (C.valIO !== -1 ? r[C.valIO] : '').toString().toLowerCase();
    const valRE = (C.valRE !== -1 ? r[C.valRE] : '').toString().toLowerCase();
    const valPS = (C.valPS !== -1 ? r[C.valPS] : '').toString().toLowerCase();

    const isIO = valIO.includes('valid io');
    const isRE = valRE.includes('valid re');
    const isPS = valPS.includes('put in service');

    const kabRaw = C.kabupaten !== -1 && r[C.kabupaten]
      ? r[C.kabupaten].toString().trim().toUpperCase() : '';
    const cluster = KAB_TO_CLUSTER[kabRaw] || '-';

    const stoRaw = C.sto !== -1 && r[C.sto]
      ? r[C.sto].toString().trim().toUpperCase() : '';

    // Parse koordinat "lat,lng"
    let lat = null, lng = null;
    if (C.koordinat !== -1 && r[C.koordinat]) {
      const raw = r[C.koordinat].toString().trim();
      const parts = raw.split(/[\s,]+/);
      if (parts.length >= 2) {
        lat = parseFloat(parts[0]);
        lng = parseFloat(parts[1]);
        if (isNaN(lat) || isNaN(lng)) { lat = null; lng = null; }
      }
    }

    orders.push({
      status:    (r[C.status]   || '').toString().trim(),
      wok:       (C.wok   !== -1 ? r[C.wok]   : '').toString().trim(),
      sto:       stoRaw,
      cluster:   cluster,
      isIO, isRE, isPS,
      tglIO:     cleanDate(tglIO),
      tglRE:     cleanDate(C.tanggalRE !== -1 ? r[C.tanggalRE] : ''),
      tglPS:     cleanDate(C.tanggalPS !== -1 ? r[C.tanggalPS] : ''),
      codeSF:    (C.codeSF !== -1 ? r[C.codeSF] : '').toString().trim(),
      namaSF:    (C.namaSF !== -1 ? r[C.namaSF] : '').toString().trim(),
      channel:   (C.channel  !== -1 ? r[C.channel]  : '').toString().trim(),
      channel2:  (C.channel2 !== -1 ? r[C.channel2] : '').toString().trim(),
      reasonK:   (C.reasonK  !== -1 ? r[C.reasonK]  : '').toString().trim(),
      kendTek:   (C.kendTek  !== -1 ? r[C.kendTek]  : '').toString().trim(),
      ioTimes:   (C.ioTimes  !== -1 ? r[C.ioTimes]  : '').toString().trim(),
      timeRange: (C.timeRange !== -1 ? r[C.timeRange] : '').toString().trim(),
      odp:       (C.odp       !== -1 ? r[C.odp]       : '').toString().trim(),
      orderId:   (C.orderId   !== -1 ? r[C.orderId]   : '').toString().trim(),
      tglKend:   cleanDate(C.tglKend !== -1 ? r[C.tglKend] : ''),
      produk:    (C.produk    !== -1 ? r[C.produk]    : '').toString().trim(),
      kabupaten: kabRaw,
      lat, lng,
    });
  }

  return { orders };
}

// ── Kendala ──────────────────────────────────────────────────
function getKendala() {
  const ws   = ss().getSheetByName('Kendala');
  if (!ws) return { kendala: [], fuKendala: [], byJenis: {}, total: 0, totalFU: 0 };
  const rows = ws.getDataRange().getValues();
  const hdr  = rows[0];
  const col  = (name) => hdr.indexOf(name);
  const C = {
    tgl:     col('Tanggal'),
    sto:     col('STO'),
    kendala: col('Jenis Kendala'),
    detail:  col('Detail'),
    ket:     col('Keterangan'),
    teknisi: col('Teknisi'),
    order:   col('Order ID'),
    solusi:  col('Solusi'),
    status:  col('Status FU'),
  };

  const list = [], fu = [];
  const byJenis = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.tgl]) continue;

    const sto     = (r[C.sto] || '').toString().toUpperCase().trim();
    const cluster = STO_TO_CLUSTER[sto] || KAB_TO_CLUSTER[sto] || '-';
    const jenis   = (r[C.kendala] || 'Lainnya').toString().trim();

    byJenis[jenis] = (byJenis[jenis] || 0) + 1;

    const row = {
      tanggal: r[C.tgl] ? fmtDate(r[C.tgl]) : '',
      sto, cluster, jenis,
      detail:   (r[C.detail]  || '').toString(),
      keterangan:(r[C.ket]    || '').toString(),
      teknisi:  (r[C.teknisi] || '').toString(),
      orderId:  (r[C.order]   || '').toString(),
    };
    list.push(row);

    const statusFU = (r[C.status] || '').toString().toLowerCase();
    if (statusFU === 'fu aktif' || statusFU === 'aktif') {
      fu.push({ ...row, solusi: (r[C.solusi] || '').toString() });
    }
  }

  list.sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  fu.sort((a, b) => b.tanggal.localeCompare(a.tanggal));

  return {
    kendala:   list,
    fuKendala: fu,
    byJenis,
    total:     list.length,
    totalFU:   fu.length,
  };
}

// ── Absensi ──────────────────────────────────────────────────
function getAbsen() {
  const ws   = ss().getSheetByName('Absensi');
  if (!ws) return { today: [], trend: [], count: 0 };
  const rows = ws.getDataRange().getValues();
  const hdr  = rows[0];
  const col  = (name) => hdr.indexOf(name);
  const C = {
    waktu:     col('Waktu'),
    nama:      col('Nama'),
    username:  col('Username'),
    kelurahan: col('Kelurahan'),
    deal:      col('Deal'),
  };

  const now      = new Date();
  const todayStr = fmtDate(now);
  const trend    = {};
  const today    = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.waktu]) continue;

    let d;
    try { d = new Date(r[C.waktu]); } catch(e) { continue; }
    if (isNaN(d)) continue;

    const dateStr = fmtDate(d);
    trend[dateStr] = (trend[dateStr] || 0) + 1;

    if (dateStr === todayStr) {
      today.push({
        waktu:     r[C.waktu],
        nama:      (r[C.nama]      || '').toString(),
        username:  (r[C.username]  || '').toString(),
        kelurahan: (r[C.kelurahan] || '').toString(),
        deal:      ['ya','yes','1','true'].includes((r[C.deal]||'').toString().trim().toLowerCase()) ? 1 : 0,
      });
    }
  }

  const trendArr = Object.entries(trend)
    .sort(([a], [b]) => a.localeCompare(b))
    .slice(-30)
    .map(([date, count]) => ({ date, count }));

  return { today, trend: trendArr, count: today.length };
}