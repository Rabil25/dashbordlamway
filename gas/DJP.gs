// ================================================================
//  DJP – Daily Journey Plan
//  Spreadsheet: Monitoring DJP
//  Sheet DJP_Plan     : TGL | USERNAME_SF | NAMA_SF | ODP | CLUSTER | KAB | KEC | DESA | LAT | LNG | AVAI | DIBUAT_OLEH | DIBUAT_AT
//  Sheet DJP_Realisasi: SCHEDULE_ID | TUGAS_ID | NAMA_SF | TGL | ODP | STATUS | CHECKIN | CHECKOUT | LAT | LNG | ODP_CONDITION | PROSPECT_ESTIMATE | NOTES | UPLOADED_BY | UPLOADED_AT
// ================================================================

const DJP_CONFIG = {
  SHEET_ID: '1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ',
  SHEET_PLAN: 'DJP_Plan',
  SHEET_REAL: 'DJP_Realisasi',
  SHEET_USERS: 'Users',           // sheet Users ada di spreadsheet utama
  BATCH_SIZE: 500,
};

// Kolom DJP_Plan (0-based)
const P = {
  TGL:0, USERNAME_SF:1, NAMA_SF:2, ODP:3,
  CLUSTER:4, KAB:5, KEC:6, DESA:7,
  LAT:8, LNG:9, AVAI:10,
  DIBUAT_OLEH:11, DIBUAT_AT:12,
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
//  MAIN HANDLER – dipanggil dari doGet di Code.gs
// ================================================================
function handleDJP(type, params) {
  switch (type) {
    case 'djp_users':             return djpGetUsers(params);
    case 'djp_plan_save':         return djpPlanSave(params);
    case 'djp_plan_get':          return djpPlanGet(params);
    case 'djp_plan_delete':       return djpPlanDelete(params);
    case 'djp_realisasi_upload':  return djpRealisasiUpload(params);
    case 'djp_daily':             return djpDaily(params);
    default:
      return { ok:false, error:'DJP type tidak dikenal: ' + type };
  }
}

// ================================================================
//  djp_users – list semua user aktif (untuk pilih SF)
//  Baca dari Users sheet di spreadsheet utama
// ================================================================
function djpGetUsers(params) {
  // Verifikasi session
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID); // spreadsheet utama
    const ws = ss.getSheetByName(USERS_SHEET);
    if (!ws) return { ok:false, error:'Sheet Users tidak ditemukan.' };

    const rows = ws.getDataRange().getValues().slice(1);
    const users = rows
      .filter(r => (r[U.STATUS]||'') === 'active')
      .map(r => ({
        username: (r[U.USERNAME]||'').toString().trim(),
        nama:     (r[U.NAMA]||'').toString().trim(),
        role:     (r[U.ROLE]||'').toString().trim(),
      }))
      .filter(u => u.username);

    return { ok:true, users };
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
    rows = JSON.parse(params.rows || '[]');
  } catch(e) {
    return { ok:false, error:'Format rows tidak valid (harus JSON array).' };
  }
  if (!Array.isArray(rows) || rows.length === 0) {
    return { ok:false, error:'rows kosong.' };
  }

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok:false, error:'Sheet DJP_Plan tidak ditemukan.' };

    const now = new Date().toISOString();
    const dibuat_oleh = auth.user.username;

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
    ]);

    // Append batch
    for (let i = 0; i < toAppend.length; i += DJP_CONFIG.BATCH_SIZE) {
      const batch = toAppend.slice(i, i + DJP_CONFIG.BATCH_SIZE);
      ws.getRange(ws.getLastRow() + 1, 1, batch.length, batch[0].length)
        .setValues(batch);
    }

    // Invalidate cache
    try {
      CacheService.getScriptCache().remove('djp_plan_' + (rows[0].tgl || ''));
    } catch(e) {}

    return { ok:true, saved: toAppend.length };
  } catch(e) {
    return { ok:false, error:'djpPlanSave: ' + e.message };
  }
}

// ================================================================
//  djp_plan_get – ambil plan by tanggal
//  params.tgl = 'yyyy-MM-dd'
// ================================================================
function djpPlanGet(params) {
  const auth = verifyToken(params.session_token || '');
  if (!auth.ok) return auth;

  const tgl = (params.tgl || '').trim();
  if (!tgl) return { ok:false, error:'Parameter tgl wajib (yyyy-MM-dd).' };

  // Cache
  const cache = CacheService.getScriptCache();
  const cacheKey = 'djp_plan_' + tgl;
  try {
    const cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok:false, error:'Sheet DJP_Plan tidak ditemukan.' };

    const rows = ws.getDataRange().getValues().slice(1); // skip header
    const data = rows
      .filter(r => (r[P.TGL]||'').toString().trim() === tgl)
      .map(r => ({
        tgl:          (r[P.TGL]||'').toString().trim(),
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
      }));

    const result = { ok:true, tgl, data, total: data.length };
    try { cache.put(cacheKey, JSON.stringify(result), 300); } catch(e) {} // cache 5 menit
    return result;
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

  try {
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
    const ws = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    if (!ws) return { ok:false, error:'Sheet DJP_Plan tidak ditemukan.' };

    const rows = ws.getDataRange().getValues();
    // Hapus dari bawah supaya index tidak bergeser
    let deleted = 0;
    for (let i = rows.length - 1; i >= 1; i--) {
      const rowTgl = (rows[i][P.TGL]||'').toString().trim();
      const rowSF  = (rows[i][P.USERNAME_SF]||'').toString().trim();
      if (rowTgl === tgl && rowSF === username_sf) {
        ws.deleteRow(i + 1);
        deleted++;
      }
    }

    // Invalidate cache
    try { CacheService.getScriptCache().remove('djp_plan_' + tgl); } catch(e) {}

    return { ok:true, deleted };
  } catch(e) {
    return { ok:false, error:'djpPlanDelete: ' + e.message };
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
    rows = JSON.parse(params.rows || '[]');
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

    // Baca Plan
    const wsPlan = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN);
    const planRows = wsPlan ? wsPlan.getDataRange().getValues().slice(1) : [];
    const plan = planRows
      .filter(r => (r[P.TGL]||'').toString().trim() === tgl)
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
      .filter(r => (r[R.TGL]||'').toString().trim() === tgl)
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

    // Gabungkan: per SF → list ODP plan + status realisasi
    const sfMap = {};
    plan.forEach(p => {
      if (!sfMap[p.username_sf]) {
        sfMap[p.username_sf] = { username_sf: p.username_sf, nama_sf: p.nama_sf, odp_list: [] };
      }
      // Cari realisasi untuk ODP ini
      const real = realisasi.find(r => r.odp === p.odp && r.nama_sf === p.nama_sf) || null;
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
        checkout:      real ? real.checkout : '',
        odp_condition: real ? real.odp_condition : '',
        prospect:      real ? real.prospect_estimate : 0,
        notes:         real ? real.notes : '',
      });
    });

    const summary = Object.values(sfMap);
    const result = {
      ok: true,
      tgl,
      total_sf:    summary.length,
      total_plan:  plan.length,
      total_real:  realisasi.length,
      summary,
      updatedAt:   new Date().toISOString(),
    };

    try { cache.put(cacheKey, JSON.stringify(result), 180); } catch(e) {}
    return result;
  } catch(e) {
    return { ok:false, error:'djpDaily: ' + e.message };
  }
}
