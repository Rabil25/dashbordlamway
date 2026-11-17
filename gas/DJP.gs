// ================================================================
//  DJP – Daily Journey Plan  (Versi 23)
//  Spreadsheet Monitoring : 1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ
//  Sheet DJP_Plan         : TGL|USERNAME_SF|NAMA_SF|ODP|CLUSTER|KAB|KEC|DESA|LAT|LNG|AVAI|DIBUAT_OLEH|DIBUAT_AT|TAHUN
//  Sheet DJP_Realisasi    : SCHEDULE_ID|TUGAS_ID|NAMA_SF|TGL|ODP|STATUS|CHECKIN|CHECKOUT|LAT|LNG|ODP_CONDITION|PROSPECT_ESTIMATE|NOTES|UPLOADED_BY|UPLOADED_AT
//  Sheet ODPBot           : (header dinamis) ODP|AVAI|USED|OCC 2|LATITUDE|LONGITUDE|KAB|KEC|DESA|Telkomsel Cluster
//  Spreadsheet Users      : 1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM
//  Sheet Update SF        : (header baris 1) col[1]=NAMA, col[2]=CODE/NIK, col[15]=CLUSTER, col[16]=STATUS, col[19]=STO_CODE, col[20]=KAB_U, col[33]=KAB_AH
// ================================================================

const DJP_CONFIG = {
  SHEET_ID:     '1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ',
  SHEET_PLAN:   'DJP_Plan',
  SHEET_REAL:   'DJP_Realisasi',
  SHEET_ODP:    'ODPBot',
  SF_SHEET_ID:  '1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM',
  SF_SHEET_NAME:'Update SF',
  BATCH_SIZE:   500,
  // Bot Telegram IndiHmLPG — plan otomatis jadi tugas clock in/out
  BOT_SHEET_ID: '1csOC-Uat1y4D1mXY1Y6uM0oSeUzG2iOxyoh81I92QQY',
  BOT_TUGAS:    'Tugas_DJP',
  BOT_REKAP:    'Rekap_TugasDJP',
  BOT_MIN_OUT:  15,
  BOT_RADIUS:   '150 meter',
};

// Kolom DJP_Plan (0-based) — TAHUN ditambah di index 13 (kolom N)
const P = {
  TGL:0, USERNAME_SF:1, NAMA_SF:2, ODP:3,
  CLUSTER:4, KAB:5, KEC:6, DESA:7,
  LAT:8, LNG:9, AVAI:10,
  DIBUAT_OLEH:11, DIBUAT_AT:12, TAHUN:13,
};

// Kolom DJP_Realisasi (0-based)
const R = {
  SCHEDULE_ID:0, TUGAS_ID:1, NAMA_SF:2, TGL:3,
  ODP:4, STATUS:5, CHECKIN:6, CHECKOUT:7,
  LAT:8, LNG:9, ODP_CONDITION:10,
  PROSPECT_ESTIMATE:11, NOTES:12,
  UPLOADED_BY:13, UPLOADED_AT:14,
};

// ================================================================
//  HELPER BERSAMA
// ================================================================
// Normalisasi tanggal dari Sheets ke 'yyyy-MM-dd' (WIB).
// Sheets otomatis mengubah teks '2026-11-12' menjadi sel bertipe tanggal,
// jadi getValues() bisa mengembalikan Date ATAU string. Semua perbandingan
// tanggal di file ini WAJIB lewat fungsi ini.
function djpNormTgl(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) return Utilities.formatDate(val, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  return val.toString().trim().slice(0, 10);
}

// Hapus baris DJP_Plan yang cocok dengan kumpulan kunci 'tgl|username_sf'.
// Baris berurutan dihapus per blok (deleteRows) supaya cepat.
// Hanya baris yang cocok yang disentuh; baris lain tidak diubah.
function djpDeletePlanRows_(ws, keySet) {
  const values = ws.getDataRange().getValues();
  const hit = [];
  for (let i = 1; i < values.length; i++) {          // i=0 → header
    const key = djpNormTgl(values[i][P.TGL]) + '|' + (values[i][P.USERNAME_SF] || '').toString().trim();
    if (keySet[key]) hit.push(i + 1);                 // nomor baris sheet (1-based)
  }
  // Dari bawah ke atas, gabungkan baris berurutan jadi satu deleteRows
  let end = hit.length - 1;
  while (end >= 0) {
    let start = end;
    while (start > 0 && hit[start - 1] === hit[start] - 1) start--;
    ws.deleteRows(hit[start], end - start + 1);
    end = start - 1;
  }
  return hit.length;
}


// ================================================================
//  INTEGRASI BOT IndiHmLPG
//  Tugas_DJP : Lat|Long|Nama Tugas|ODP|Kategori|Min Clock Out|Mulai yyyyMMdd|Akhir yyyyMMdd|SF code|Radius
//  Rekap_TugasDJP : Tanggal|SF Code|Nama SF|ODP/Nama Tugas|Jam In|Koordinat In|Jam Out|Durasi|Status
// ================================================================
function djpBotYmd_(v) {           // apa pun formatnya → 'yyyyMMdd'
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyyMMdd');
  const t = (v || '').toString().trim();
  let m = t.match(/^(\d{4})-?(\d{2})-?(\d{2})/); if (m) return m[1] + m[2] + m[3];
  m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m) return m[3] + ('0' + m[1]).slice(-2) + ('0' + m[2]).slice(-2); // M/d/yyyy
  return t.replace(/\D/g, '').slice(0, 8);
}
// Ganti tugas bot untuk setiap kunci 'yyyyMMdd|SFCODE', lalu tambahkan baris baru (boleh kosong = hapus saja)
function djpSyncBotTugas_(keySet, newRows) {
  try {
    const ws = SpreadsheetApp.openById(DJP_CONFIG.BOT_SHEET_ID).getSheetByName(DJP_CONFIG.BOT_TUGAS);
    if (!ws) return { ok:false, error:'Tab Tugas_DJP tidak ditemukan' };
    const v = ws.getDataRange().getValues(), hit = [];
    for (let i = 1; i < v.length; i++) {
      const k = djpBotYmd_(v[i][6]) + '|' + (v[i][8] || '').toString().trim().toUpperCase();
      if (keySet[k]) hit.push(i + 1);
    }
    let end = hit.length - 1;
    while (end >= 0) { let st = end; while (st > 0 && hit[st - 1] === hit[st] - 1) st--; ws.deleteRows(hit[st], end - st + 1); end = st - 1; }
    if (newRows.length) {
      // Bot mencocokkan SF code PERSIS (huruf besar/kecil) dengan Userdata kolom F → pakai penulisan dari Userdata
      const kode = {};
      try {
        const ud = SpreadsheetApp.openById(DJP_CONFIG.BOT_SHEET_ID).getSheetByName('Userdata');
        ud.getRange(2, 6, Math.max(1, ud.getLastRow() - 1), 1).getValues().forEach(x => {
          const k = (x[0] || '').toString().trim(); if (!k) return;
          const a = kode[k.toUpperCase()] = kode[k.toUpperCase()] || []; if (a.indexOf(k) < 0) a.push(k); });
      } catch (e) {}
      // Satu SF bisa terdaftar dengan beberapa penulisan (mis. SPSAB71 & Spsab71) → tulis untuk tiap penulisan
      const out = [];
      newRows.forEach(r => { const d = djpBotYmd_(r.tgl), sf = (r.username_sf || '').toString().trim();
        (kode[sf.toUpperCase()] || [sf]).forEach(c => out.push([r.lat || '', r.lng || '', r.odp, r.odp, 'DJP', DJP_CONFIG.BOT_MIN_OUT, d, d, c, DJP_CONFIG.BOT_RADIUS])); });
      const rg = ws.getRange(ws.getLastRow() + 1, 1, out.length, out[0].length);
      rg.offset(0, 6, out.length, 2).setNumberFormat('@');     // tanggal tetap teks yyyyMMdd
      rg.setValues(out);
    }
    return { ok:true, removed: hit.length, added: newRows.length };
  } catch (e) { return { ok:false, error: e.message }; }
}
// Status clock in/out dari bot untuk satu tanggal → { 'SF|ODP': {in, out, durasi, status} }
function djpBotStatus_(tgl) {
  const res = {};
  try {
    const ws = SpreadsheetApp.openById(DJP_CONFIG.BOT_SHEET_ID).getSheetByName(DJP_CONFIG.BOT_REKAP);
    if (!ws || ws.getLastRow() < 2) return res;
    const want = djpBotYmd_(tgl), fmt = x => x instanceof Date ? ('0' + x.getHours()).slice(-2) + ':' + ('0' + x.getMinutes()).slice(-2) : (x || '').toString().trim().slice(0, 5);
    ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues().forEach(r => {
      if (djpBotYmd_(r[0]) !== want) return;
      const k = (r[1] || '').toString().trim().toUpperCase() + '|' + (r[3] || '').toString().trim().toUpperCase();
      res[k] = { in: fmt(r[4]), out: fmt(r[6]), durasi: (r[7] || '').toString(), status: (r[8] || '').toString().trim() };
    });
  } catch (e) {}
  return res;
}

// ================================================================
//  MAIN HANDLER – dipanggil dari doGet di Code.gs
// ================================================================
function handleDJP(type, params) {
  // Frontend kirim data via ?payload=JSON (GET) untuk hindari CORS preflight POST
  if (params.payload) {
    try {
      var p = JSON.parse(params.payload);
      for (var k in p) params[k] = p[k];
    } catch(e) {}
  }
  switch (type) {
    case 'djp_users':             return djpGetUsers(params);
    case 'djp_plan_save':         return djpPlanSave(params);
    case 'djp_plan_get':          return djpPlanGet(params);
    case 'djp_plan_delete':       return djpPlanDelete(params);
    case 'djp_realisasi_upload':  return djpRealisasiUpload(params);
    case 'djp_daily':             return djpDaily(params);
    case 'djp_odp_list':          return djpOdpList(params);
    case 'djp_plan_history':      return djpPlanHistory(params);
    default:
      return { ok:false, error:'DJP type tidak dikenal: ' + type };
  }
}

// ================================================================
//  djp_users – list SF aktif dari sheet "Update SF"
//  params.kabs (optional) = JSON array kabupaten untuk filter cluster
//    contoh: ["LAMPUNG BARAT","TANGGAMUS","PESISIR BARAT"]
// ================================================================
function djpGetUsers(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  // Parse filter kabupaten (opsional)
  let kabFilter = [];
  try {
    if (params.kabs) {
      kabFilter = typeof params.kabs === 'string' ? JSON.parse(params.kabs) : params.kabs;
      kabFilter = kabFilter.map(k => k.toUpperCase().trim());
    }
  } catch(e) {}

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SF_SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SF_SHEET_NAME);
    if (!ws) return { ok:false, error:'Sheet "Update SF" tidak ditemukan.' };

    const rows = ws.getDataRange().getValues().slice(1); // skip header baris 1

    const users = rows
      .filter(r => {
        const status = (r[16] || '').toString().toUpperCase().trim();
        if (status !== 'ACTIVE') return false;
        const nama = (r[1] || '').toString().trim();
        const code = (r[2] || '').toString().trim();
        if (!nama || !code) return false;
        if (kabFilter.length > 0) {
          // Kolom AH (index 33) = kabupaten, fallback ke kolom U (index 20)
          const kab = (r[33] || r[20] || '').toString().trim().toUpperCase();
          return kabFilter.some(k => kab.includes(k) || k.includes(kab));
        }
        return true;
      })
      .map(r => ({
        username: (r[2]  || '').toString().trim(),   // NIK/code
        nama:     (r[1]  || '').toString().trim(),   // nama SF
        kab:      (r[33] || r[20] || '').toString().trim(), // kabupaten (kolom AH, fallback U)
        sto:      (r[19] || '').toString().trim(),   // kode STO
        cluster:  (r[15] || '').toString().trim(),   // cluster string
        role:     'user',
      }))
      .filter(u => u.username);

    // Deduplikasi by username
    const seen = new Set();
    const unique = users.filter(u => {
      if (seen.has(u.username)) return false;
      seen.add(u.username);
      return true;
    });

    return { ok:true, users: unique };
  } catch(e) {
    return { ok:false, error:'djpGetUsers: ' + e.message };
  }
}

// ================================================================
//  djp_plan_save – simpan satu atau banyak baris plan DJP
//  params.rows = JSON array of plan objects
//  params.session_token = token user yang input
// ================================================================
function djpPlanSave(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  let rows;
  try {
    rows = (typeof params.rows === 'string') ? JSON.parse(params.rows) : (params.rows || []);
  } catch(e) {
    return { ok:false, error:'Format rows tidak valid (harus JSON array).' };
  }
  if (!Array.isArray(rows) || rows.length === 0) {
    return { ok:false, error:'rows kosong.' };
  }

  // Lock: cegah dua simpan/hapus berjalan bersamaan dan saling menimpa
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch(e) {
    return { ok:false, error:'Server sedang memproses simpan lain. Coba lagi sebentar.' };
  }

  const _t0 = Date.now(), _ms = {};
  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok:false, error:'Sheet DJP_Plan tidak ditemukan.' };
    _ms.open = Date.now() - _t0;

    const now = new Date().toISOString();
    const dibuat_oleh = auth.user.username;

    // Validasi minimum: tgl, username_sf, odp wajib ada di setiap baris
    const bad = rows.filter(r => !/^\d{4}-\d{2}-\d{2}$/.test((r.tgl || '').toString().trim())
                              || !(r.username_sf || '').toString().trim()
                              || !(r.odp || '').toString().trim());
    if (bad.length) {
      return { ok:false, error: bad.length + ' baris tidak valid (tgl/username_sf/odp kosong atau format tgl salah). Tidak ada yang disimpan.' };
    }

    // GANTI, bukan tambah: hapus dulu plan lama untuk setiap (tgl, SF) yang dikirim
    const keySet = {};
    const tglSet = {};
    rows.forEach(r => {
      const t = r.tgl.toString().trim();
      keySet[t + '|' + r.username_sf.toString().trim()] = true;
      tglSet[t] = true;
    });
    const replaced = djpDeletePlanRows_(ws, keySet);
    _ms.hapusLama = Date.now() - _t0;

    const toAppend = rows.map(r => [
      r.tgl          || '',
      r.username_sf  || '',
      r.nama_sf      || '',
      r.odp          || '',
      r.cluster      || '',
      r.kab          || '',
      r.kec          || '',
      r.desa         || '',
      r.lat          || '',
      r.lng          || '',
      r.avai         || '',
      dibuat_oleh,
      now,
      r.tahun        || (r.tgl ? r.tgl.toString().slice(0,4) : new Date().getFullYear().toString()),
    ]);

    // Append batch
    for (let i = 0; i < toAppend.length; i += DJP_CONFIG.BATCH_SIZE) {
      const batch = toAppend.slice(i, i + DJP_CONFIG.BATCH_SIZE);
      ws.getRange(ws.getLastRow() + 1, 1, batch.length, batch[0].length)
        .setValues(batch);
    }


    // Invalidate cache untuk semua tanggal yang tersentuh
    try {
      CacheService.getScriptCache().removeAll(Object.keys(tglSet).map(t => 'djp_plan_' + t));
    } catch(e) {}

    // Kirim jadi tugas bot (gagal di sini tidak membatalkan simpan plan)
    const botKeys = {};
    Object.keys(keySet).forEach(k => { const a = k.split('|'); botKeys[djpBotYmd_(a[0]) + '|' + a[1].toUpperCase()] = true; });
    const bot = djpSyncBotTugas_(botKeys, rows);
    _ms.total = Date.now() - _t0;
    return { ok:true, saved: toAppend.length, replaced: replaced, bot: bot, _ms: _ms };
  } catch(e) {
    return { ok:false, error:'djpPlanSave: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
//  djp_plan_get – ambil plan by tanggal (+ optional filter tahun)
//  params.tgl   = 'yyyy-MM-dd'
//  params.tahun = '2025' (opsional, default: ambil dari tgl)
// ================================================================
// Tempelkan status clock in/out bot ke tiap baris plan (selalu segar, tidak ikut cache)
function djpWithBot_(result, tgl) {
  if (!result || !result.ok || !result.data || !result.data.length) return result;
  const st = djpBotStatus_(tgl);
  result.data.forEach(r => { const b = st[r.username_sf.toUpperCase() + '|' + r.odp.toUpperCase()]; if (b) r.bot = b; });
  return result;
}
function djpPlanGet(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  const tgl = (params.tgl || '').trim();
  if (!tgl) return { ok:false, error:'Parameter tgl wajib (yyyy-MM-dd).' };
  const tahunFilter = (params.tahun || '').toString().trim();

  // Cache
  const cache = CacheService.getScriptCache();
  const cacheKey = 'djp_plan_' + tgl;
  try {
    const cached = cache.get(cacheKey);
    if (cached) return djpWithBot_(JSON.parse(cached), tgl);
  } catch(e) {}

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok:false, error:'Sheet DJP_Plan tidak ditemukan.' };

    const rows = ws.getDataRange().getValues().slice(1); // skip header

    // Tanggal dinormalkan lewat helper bersama (zona waktu CONFIG.TIMEZONE)
    const formatTgl = djpNormTgl;

    const data = rows
      .filter(r => {
        if (formatTgl(r[P.TGL]) !== tgl) return false;
        // Filter tahun: jika ada filter tahun, cocokkan kolom TAHUN atau 4 char awal TGL
        if (tahunFilter) {
          const tahunRow = (r[P.TAHUN] || '').toString().trim()
            || formatTgl(r[P.TGL]).slice(0,4);
          if (tahunRow && tahunRow !== tahunFilter) return false;
        }
        return true;
      })
      .map(r => ({
        tgl:          formatTgl(r[P.TGL]),
        username_sf:  (r[P.USERNAME_SF]||'').toString().trim(),
        nama_sf:      (r[P.NAMA_SF]||'').toString().trim(),
        odp:          (r[P.ODP]||'').toString().trim(),
        cluster:      (r[P.CLUSTER]||'').toString().trim(),
        kab:          (r[P.KAB]||'').toString().trim(),
        kec:          (r[P.KEC]||'').toString().trim(),
        desa:         (r[P.DESA]||'').toString().trim(),
        lat:          parseFloat(r[P.LAT]) || 0,
        lng:          parseFloat(r[P.LNG]) || 0,
        avai:         parseInt(r[P.AVAI]) || 0,
        dibuat_oleh:  (r[P.DIBUAT_OLEH]||'').toString().trim(),
        dibuat_at:    (r[P.DIBUAT_AT]||'').toString().trim(),
        tahun:        (r[P.TAHUN]||'').toString().trim() || formatTgl(r[P.TGL]).slice(0,4),
      }));

    const result = { ok:true, tgl, data, total: data.length };
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(e) {} // cache 5 menit
    return djpWithBot_(result, tgl);
  } catch(e) {
    return { ok:false, error:'djpPlanGet: ' + e.message };
  }
}

// ================================================================
//  djp_plan_delete – hapus semua plan untuk tgl + username_sf tertentu
//  params.tgl, params.username_sf
// ================================================================
function djpPlanDelete(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  const tgl        = (params.tgl || '').trim();
  const username_sf = (params.username_sf || '').trim();
  if (!tgl || !username_sf) return { ok:false, error:'tgl dan username_sf wajib.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch(e) {
    return { ok:false, error:'Server sedang memproses simpan lain. Coba lagi sebentar.' };
  }

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok:false, error:'Sheet DJP_Plan tidak ditemukan.' };

    const keySet = {};
    keySet[tgl + '|' + username_sf] = true;
    const deleted = djpDeletePlanRows_(ws, keySet);

    // Invalidate cache
    try { CacheService.getScriptCache().remove('djp_plan_' + tgl); } catch(e) {}
    const bk = {}; bk[djpBotYmd_(tgl) + '|' + username_sf.toUpperCase()] = true;
    djpSyncBotTugas_(bk, []);

    if (deleted === 0) {
      return { ok:false, deleted: 0, error:'Tidak ada baris plan untuk ' + username_sf + ' pada ' + tgl + '.' };
    }
    return { ok:true, deleted };
  } catch(e) {
    return { ok:false, error:'djpPlanDelete: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
//  djp_realisasi_upload – batch append data CSV dari Web Telkomsel
//  params.rows = JSON array, params.session_token
// ================================================================
function djpRealisasiUpload(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  let rows;
  try {
    rows = Array.isArray(params.rows) ? params.rows : JSON.parse(params.rows || '[]');
  } catch(e) {
    return { ok:false, error:'Format rows tidak valid.' };
  }
  if (!Array.isArray(rows) || rows.length === 0) {
    return { ok:false, error:'rows kosong.' };
  }

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_REAL);
    if (!ws) return { ok:false, error:'Sheet DJP_Realisasi tidak ditemukan.' };

    const now          = new Date().toISOString();
    const uploaded_by  = auth.user.username;

    const toAppend = rows.map(r => [
      r.schedule_id       || '',
      r.tugas_id          || '',
      r.nama_sf           || r['TUGAS NAME'] || '',
      r.tgl               || r['TUGAS DATE'] || '',
      r.odp               || r['ODP']        || '',
      r.status            || r['STATUS']     || '',
      r.checkin           || r['CHECKIN']    || '',
      r.checkout          || r['CHECKOUT']   || '',
      r.lat               || r['LATITUDE']   || '',
      r.lng               || r['LONGITUDE']  || '',
      r.odp_condition     || r['ODP CONDITION'] || '',
      r.prospect_estimate || r['PROSPECT ESTIMATE'] || '',
      r.notes             || r['NOTES']      || '',
      uploaded_by,
      now,
    ]);

    // Cek duplikasi berdasarkan SCHEDULE_ID yang sudah ada
    const existing = ws.getDataRange().getValues().slice(1);
    const existingIds = new Set(existing.map(r => (r[R.SCHEDULE_ID]||'').toString().trim()));
    const filtered = toAppend.filter(r => !existingIds.has((r[0]||'').toString().trim()));

    if (filtered.length === 0) {
      return { ok:true, saved:0, skipped: toAppend.length, message:'Semua data sudah ada (duplikat).' };
    }

    // Batch append
    for (let i = 0; i < filtered.length; i += DJP_CONFIG.BATCH_SIZE) {
      const batch = filtered.slice(i, i + DJP_CONFIG.BATCH_SIZE);
      ws.getRange(ws.getLastRow() + 1, 1, batch.length, batch[0].length)
        .setValues(batch);
    }

    // Invalidate daily cache
    try { CacheService.getScriptCache().remove('djp_daily_' + new Date().toISOString().slice(0,10)); } catch(e) {}

    return {
      ok: true,
      saved:   filtered.length,
      skipped: toAppend.length - filtered.length,
    };
  } catch(e) {
    return { ok:false, error:'djpRealisasiUpload: ' + e.message };
  }
}

// ================================================================
//  djp_daily – plan + realisasi untuk tanggal tertentu
//  params.tgl = 'yyyy-MM-dd' (default: hari ini)
// ================================================================
function djpDaily(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  const tgl = (params.tgl || Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd')).trim();

  // Cache 3 menit
  const cache    = CacheService.getScriptCache();
  const cacheKey = 'djp_daily_' + tgl;
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);

    const normTgl = djpNormTgl; // helper bersama

    // Baca Plan
    const wsPlan = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    const planRows = wsPlan ? wsPlan.getDataRange().getValues().slice(1) : [];
    const plan = planRows
      .filter(r => normTgl(r[P.TGL]) === tgl)
      .map(r => ({
        username_sf: (r[P.USERNAME_SF]||'').toString().trim(),
        nama_sf:     (r[P.NAMA_SF]||'').toString().trim(),
        odp:         (r[P.ODP]||'').toString().trim(),
        cluster:     (r[P.CLUSTER]||'').toString().trim(),
        kab:         (r[P.KAB]||'').toString().trim(),
        kec:         (r[P.KEC]||'').toString().trim(),
        desa:        (r[P.DESA]||'').toString().trim(),
        lat:         parseFloat(r[P.LAT]) || 0,
        lng:         parseFloat(r[P.LNG]) || 0,
      }));

    // Baca Realisasi
    const wsReal = ss.getSheetByName(DJP_CONFIG.SHEET_REAL);
    const realRows = wsReal ? wsReal.getDataRange().getValues().slice(1) : [];
    const realisasi = realRows
      .filter(r => normTgl(r[R.TGL]) === tgl)
      .map(r => ({
        schedule_id:      (r[R.SCHEDULE_ID]||'').toString().trim(),
        nama_sf:          (r[R.NAMA_SF]||'').toString().trim(),
        odp:              (r[R.ODP]||'').toString().trim(),
        status:           (r[R.STATUS]||'').toString().trim(),
        checkin:          (r[R.CHECKIN]||'').toString().trim(),
        checkout:         (r[R.CHECKOUT]||'').toString().trim(),
        odp_condition:    (r[R.ODP_CONDITION]||'').toString().trim(),
        prospect_estimate: parseInt(r[R.PROSPECT_ESTIMATE]) || 0,
        notes:            (r[R.NOTES]||'').toString().trim(),
      }));

    // Fungsi ekstrak jam dari berbagai format timestamp/waktu
    function extractJam(checkinStr) {
      if (!checkinStr) return '';
      const s = checkinStr.toString().trim();
      // Format "2025-01-15 08:30:00" atau "08:30:00" atau "08:30"
      const m = s.match(/(\d{1,2}):(\d{2})/);
      if (m) return m[1].padStart(2,'0') + ':' + m[2];
      return '';
    }

    // Gabungkan: per SF → list ODP plan + status realisasi
    const sfMap = {};
    plan.forEach(p => {
      if (!sfMap[p.username_sf]) {
        sfMap[p.username_sf] = { username_sf: p.username_sf, nama_sf: p.nama_sf, odp_list: [] };
      }
      // Cari realisasi untuk ODP ini
      const real = realisasi.find(r => r.odp === p.odp && r.nama_sf === p.nama_sf) || null;
      const checkin_jam = real ? extractJam(real.checkin) : '';
      const done = real && real.status && real.status !== '' && real.status !== '-';
      sfMap[p.username_sf].odp_list.push({
        odp:           p.odp,
        cluster:       p.cluster,
        kab:           p.kab,
        kec:           p.kec,
        desa:          p.desa,
        lat:           p.lat,
        lng:           p.lng,
        status:        real ? real.status : 'Belum',
        checkin:       real ? real.checkin : '',
        checkin_jam:   checkin_jam,                  // FIX #4: jam check-in saja
        realisasi:     done ? 'Done' : '',           // FIX #4: "Done" atau ""
        checkout:      real ? real.checkout : '',
        odp_condition: real ? real.odp_condition : '',
        prospect:      real ? real.prospect_estimate : 0,
        notes:         real ? real.notes : '',
      });
    });

    // Flat rows untuk tabel DJP Daily (Fix #4)
    const flatRows = [];
    Object.values(sfMap).forEach(sf => {
      sf.odp_list.forEach(o => {
        flatRows.push({
          nama_sf:    sf.nama_sf,
          username_sf:sf.username_sf,
          odp:        o.odp,
          kec:        o.kec,
          desa:       o.desa,
          checkin_jam:o.checkin_jam,
          realisasi:  o.realisasi,
        });
      });
    });

    const summary = Object.values(sfMap);
    const result = {
      ok: true,
      tgl,
      total_sf:    summary.length,
      total_plan:  plan.length,
      total_real:  realisasi.length,
      summary,      // nested (backward compat)
      rows: flatRows, // flat untuk tabel baru (Fix #4)
      updatedAt:   new Date().toISOString(),
    };

    try { cache.put(cacheKey, JSON.stringify(result), 180); } catch(e) {}
    return result;
  } catch(e) {
    return { ok:false, error:'djpDaily: ' + e.message };
  }
}

// ================================================================
//  djp_odp_list – daftar ODP + status Assign & Visit bulan ini
//  params.bulan = 'yyyy-MM' (opsional, default: bulan berjalan)
//  params.kabs  = JSON array kabupaten filter (opsional)
//  Return: { ok, bulan, odp_list: [{odp, occ, avail, kab, kec, desa, assign, visit}] }
// ================================================================
// djpOdpList — hanya return set assign & visit (ODP base data dari data/odp.json di frontend)
function djpOdpList(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  const now   = new Date();
  const bulan = (params.bulan || Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM')).trim();

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const assignSet = {};
    const visitSet  = {};

    // ── 1. Assign: ODP yang di-plan bulan ini ──
    const wsPlan = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (wsPlan) {
      wsPlan.getDataRange().getValues().slice(1).forEach(r => {
        if (djpNormTgl(r[P.TGL]).startsWith(bulan)) {
          const odp = (r[P.ODP]||'').toString().trim();
          if (odp) assignSet[odp] = true;
        }
      });
    }

    // ── 2. Visit: ODP yang ada di Realisasi bulan ini ──
    const wsReal = ss.getSheetByName(DJP_CONFIG.SHEET_REAL);
    if (wsReal) {
      wsReal.getDataRange().getValues().slice(1).forEach(r => {
        if (djpNormTgl(r[R.TGL]).startsWith(bulan)) {
          const odp = (r[R.ODP]||'').toString().trim();
          if (odp) visitSet[odp] = true;
        }
      });
    }

    return {
      ok:     true,
      bulan,
      assign: Object.keys(assignSet),   // array nama ODP yang sudah assign
      visit:  Object.keys(visitSet),    // array nama ODP yang sudah visit
    };
  } catch(e) {
    return { ok:false, error:'djpOdpList: ' + e.message };
  }
}

// ================================================================
//  djpPlanHistory – ambil last plan date per ODP untuk rotasi
//  Params: session_token, kab (string, opsional), desas (JSON array string, opsional)
//  Return: { ok, history: { 'ODP_NAME': 'yyyy-MM-dd' } }
// ================================================================
function djpPlanHistory(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  let desaFilter = [];
  try {
    if (params.desas) {
      const raw = typeof params.desas === 'string' ? JSON.parse(params.desas) : (params.desas || []);
      desaFilter = raw.map(d => d.toString().trim().toLowerCase());
    }
  } catch(e) {}

  const kabFilter = (params.kab || '').toString().trim().toUpperCase();

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok: true, history: {} };

    const rows = ws.getDataRange().getValues().slice(1);

    const normTgl = djpNormTgl; // helper bersama

    // history[odp] = tanggal paling akhir ODP ini di-plan
    const history = {};

    rows.forEach(function(r) {
      const odp  = (r[P.ODP]  || '').toString().trim();
      const kab  = (r[P.KAB]  || '').toString().trim().toUpperCase();
      const desa = (r[P.DESA] || '').toString().trim().toLowerCase();
      const tgl  = normTgl(r[P.TGL]);

      if (!odp || !tgl) return;
      if (kabFilter && kab !== kabFilter) return;
      if (desaFilter.length > 0 && !desaFilter.includes(desa)) return;

      // simpan tanggal terbaru per ODP (untuk rotasi: oldest = highest priority)
      if (!history[odp] || tgl > history[odp]) {
        history[odp] = tgl;
      }
    });

    return { ok: true, history: history };
  } catch(e) {
    return { ok: false, error: 'djpPlanHistory: ' + e.message };
  }
}

// ================================================================
//  IMPORT RIWAYAT KUNJUNGAN: Source WPI → DJP_Plan  (jalankan manual sekali dari editor)
//  - Hanya SCHEDULE CATEGORY = DJP, ODP dikenal di ODPBot, kabupaten WOK LB–WK
//  - 1 baris per SF + ODP + tanggal (check-in ganda dari WPI digabung)
//  - Tidak mengubah/menghapus plan yang sudah ada; aman dijalankan ulang (tidak dobel)
//  - Baris hasil import ditandai DIBUAT_OLEH = 'import Source WPI'
//  - Batalkan dengan djpHapusImportWPI()
// ================================================================
const DJP_IMPORT_TAG = 'import Source WPI';
const DJP_WOK_KAB = {
  'LAMPUNG BARAT': 'LAMPUNG BARAT', 'TANGGAMUS': 'LAMPUNG BARAT', 'PESISIR BARAT': 'LAMPUNG BARAT',
  'LAMPUNG UTARA': 'WAY KANAN', 'WAY KANAN': 'WAY KANAN',
};   // kabupaten → cluster (aturan wilayah per kabupaten)

// ODP-KBI-FP/97 -> ODP-KBI-FP/097 (samakan dengan penulisan ODPBot)
function djpNormOdp(v) {
  return (v || '').toString().trim().toUpperCase().replace(/\/(\d{1,2})$/, (m, n) => '/' + ('00' + n).slice(-3));
}

function djpImportSourceWPI() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const wsWPI = ss.getSheetByName('Source WPI');
    const wsODP = ss.getSheetByName(DJP_CONFIG.SHEET_ODP);
    const wsPlan = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!wsWPI || !wsODP || !wsPlan) throw new Error('Tab Source WPI / ODPBot / DJP_Plan tidak ditemukan');
    const U = s => (s === null || s === undefined ? '' : s.toString()).trim().toUpperCase();

    // ODPBot → peta ODP (kolom dicari dari header)
    const odpVals = wsODP.getDataRange().getValues();
    const oh = odpVals[0].map(U), oi = n => oh.indexOf(n);
    const iO = oi('ODP'), iLa = oi('LATITUDE'), iLn = oi('LONGITUDE'), iKab = oi('KAB'), iKec = oi('KEC'), iDes = oi('DESA');
    if ([iO, iLa, iLn, iKab, iKec, iDes].some(i => i < 0)) throw new Error('Header ODPBot tidak lengkap');
    const odp = {};
    for (let i = 1; i < odpVals.length; i++) {
      const k = djpNormOdp(odpVals[i][iO]); if (!k || odp[k]) continue;
      odp[k] = { lat: odpVals[i][iLa], lng: odpVals[i][iLn], kab: (odpVals[i][iKab] || '').toString().trim(), kec: (odpVals[i][iKec] || '').toString().trim(), desa: (odpVals[i][iDes] || '').toString().trim() };
    }

    // Kunci plan yang sudah ada: tgl|sf|odp
    const planVals = wsPlan.getDataRange().getValues();
    const ada = {};
    for (let i = 1; i < planVals.length; i++) {
      ada[djpNormTgl(planVals[i][P.TGL]) + '|' + U(planVals[i][P.USERNAME_SF]) + '|' + djpNormOdp(planVals[i][P.ODP])] = true;
    }

    // Source WPI (kolom dicari dari header)
    const w = wsWPI.getDataRange().getValues();
    const wh = w[0].map(U), wi = n => wh.indexOf(n);
    const jCat = wi('SCHEDULE CATEGORY'), jYm = wi('TAHUN BULAN'), jDay = wi('TANGGAL'), jCode = wi('USER CODE'), jName = wi('USER NAME'), jOdp = wi('ODP');
    if ([jCat, jYm, jDay, jCode, jName, jOdp].some(i => i < 0)) throw new Error('Header Source WPI tidak lengkap');
    // Tanggal diambil dari kolom 'Tahun Bulan' (yyyymm) + 'Tanggal' (dd)
    const tglOf = r => {
      const ym = (r[jYm] || '').toString().replace(/\D/g, ''), d = parseInt(r[jDay], 10);
      if (ym.length !== 6 || !(d >= 1 && d <= 31)) return '';
      return ym.slice(0, 4) + '-' + ym.slice(4) + '-' + ('0' + d).slice(-2);
    };

    const now = new Date(), out = [], stat = { djp: 0, tanpaTanggal: 0, odpTakDikenal: 0, luarWOK: 0, ganda: 0, sudahAda: 0, ditambah: 0 };
    for (let i = 1; i < w.length; i++) {
      const r = w[i];
      if (U(r[jCat]) !== 'DJP') continue;
      stat.djp++;
      const tgl = tglOf(r); if (!tgl) { stat.tanpaTanggal++; continue; }
      const odpK = djpNormOdp(r[jOdp]); const o = odp[odpK]; if (!o) { stat.odpTakDikenal++; continue; }
      const cluster = DJP_WOK_KAB[U(o.kab)]; if (!cluster) { stat.luarWOK++; continue; }
      const sf = (r[jCode] || '').toString().trim();
      const key = tgl + '|' + sf.toUpperCase() + '|' + odpK;
      if (ada[key]) { stat[ada[key] === 'baru' ? 'ganda' : 'sudahAda']++; continue; }
      ada[key] = 'baru';
      out.push([tgl, sf, (r[jName] || '').toString().trim(), odpK, cluster,
        o.kab.toUpperCase(), o.kec, o.desa, o.lat, o.lng, '', DJP_IMPORT_TAG, now, tgl.slice(0, 4)]);
    }

    for (let i = 0; i < out.length; i += DJP_CONFIG.BATCH_SIZE) {
      const batch = out.slice(i, i + DJP_CONFIG.BATCH_SIZE);
      wsPlan.getRange(wsPlan.getLastRow() + 1, 1, batch.length, batch[0].length).setValues(batch);
    }
    stat.ditambah = out.length;
    Logger.log(JSON.stringify(stat));
    return stat;
  } finally {
    lock.releaseLock();
  }
}

// Batalkan import: hapus hanya baris DJP_Plan bertanda 'import Source WPI'
function djpHapusImportWPI() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ws = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID).getSheetByName(DJP_CONFIG.SHEET_PLAN);
    const v = ws.getDataRange().getValues(), hit = [];
    for (let i = 1; i < v.length; i++) if ((v[i][P.DIBUAT_OLEH] || '').toString().trim() === DJP_IMPORT_TAG) hit.push(i + 1);
    let end = hit.length - 1;
    while (end >= 0) {
      let start = end;
      while (start > 0 && hit[start - 1] === hit[start] - 1) start--;
      ws.deleteRows(hit[start], end - start + 1);
      end = start - 1;
    }
    Logger.log('Dihapus: ' + hit.length + ' baris import');
    return hit.length;
  } finally {
    lock.releaseLock();
  }
}

// ================================================================
//  SEKALI JALAN: kirim plan yang SUDAH tersimpan (hari ini s.d. ke depan) ke Tugas_DJP bot.
//  Jalankan manual dari editor Apps Script: pilih fungsi djpKirimPlanKeBot → Jalankan.
//  Aman diulang: tugas lama per (tanggal, SF) diganti, tidak dobel.
// ================================================================
function djpKirimPlanKeBot() {
  const ws = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID).getSheetByName(DJP_CONFIG.SHEET_PLAN);
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const rows = [], keys = {};
  ws.getDataRange().getValues().slice(1).forEach(r => {
    const tgl = djpNormTgl(r[P.TGL]), sf = (r[P.USERNAME_SF] || '').toString().trim(), odp = (r[P.ODP] || '').toString().trim();
    if (!tgl || tgl < today || !sf || !odp) return;
    rows.push({ tgl: tgl, username_sf: sf, odp: odp, lat: r[P.LAT], lng: r[P.LNG] });
    keys[djpBotYmd_(tgl) + '|' + sf.toUpperCase()] = true;
  });
  const res = djpSyncBotTugas_(keys, rows);
  Logger.log('Plan ' + today + ' dst: ' + rows.length + ' ODP, ' + Object.keys(keys).length + ' (tanggal×SF) → ' + JSON.stringify(res));
  return res;
}
