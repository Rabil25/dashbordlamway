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

    _ms.total = Date.now() - _t0;
    return { ok:true, saved: toAppend.length, replaced: replaced, _ms: _ms };
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
    if (cached) return JSON.parse(cached);
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
