// ================================================================
//  DJP – Daily Journey Plan  (Versi 23)
//  Spreadsheet Monitoring : 1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ
//  Sheet DJP_Plan         : TGL|USERNAME_SF|NAMA_SF|ODP|CLUSTER|KAB|KEC|DESA|LAT|LNG|AVAI|DIBUAT_OLEH|DIBUAT_AT|TAHUN
//  Sheet DJP_Realisasi_Baru: TGL|KODE_SF|NAMA_SF|PLAN|KAB|KEC|DESA|JML_ODP|DIKUNJUNGI|TARGET_DJP|SUMBER|UPLOADED_BY|UPLOADED_AT
//  Sheet ODPBot           : (header dinamis) ODP|AVAI|USED|OCC 2|LATITUDE|LONGITUDE|KAB|KEC|DESA|Telkomsel Cluster
//  Spreadsheet Users      : 1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM
//  Sheet Update SF        : (header baris 1) col[1]=NAMA, col[2]=CODE/NIK, col[15]=CLUSTER, col[16]=STATUS, col[19]=STO_CODE, col[20]=KAB_U, col[33]=KAB_AH
// ================================================================

const DJP_CONFIG = {
  SHEET_ID:     '1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ',
  SHEET_PLAN:   'DJP_Plan',
  SHEET_REAL2:  'DJP_Realisasi_Baru',   // kunjungan nyata: 1 baris = SF × desa × hari
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



// Kunjungan nyata dari DJP_Realisasi_Baru untuk tanggal/bulan (prefix 'yyyy-MM-dd' atau 'yyyy-MM')
// → { desa: {'SF|KEC|DESA|tgl': true}, sf: {'SF|tgl': {dik, target}} }
function djpRealBaruRead_(prefix) {
  const U = x => (x || '').toString().trim().toUpperCase();
  const out = { desa: {}, sf: {} };
  const ws = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID).getSheetByName(DJP_CONFIG.SHEET_REAL2);
  if (!ws || ws.getLastRow() < 2) return out;
  ws.getRange(2, 1, ws.getLastRow() - 1, RB_HEAD.length).getValues().forEach(r => {
    const tgl = djpNormTgl(r[0]); if (!tgl.startsWith(prefix)) return;
    const sf = U(r[1]), dik = +r[8] || 0;
    const k = sf + '|' + tgl; out.sf[k] = { dik: Math.max(dik, (out.sf[k] || {}).dik || 0), target: +r[9] || 0 };
    if (dik > 0 && (r[6] || '').toString().trim()) out.desa[sf + '|' + U(r[5]) + '|' + U(r[6]) + '|' + tgl] = true;
  });
  return out;
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
  if (params.format === 'detail_djp') return djpDetailSave_(rows, auth.user.username, djpNormTgl(params.range_min || ''), djpNormTgl(params.range_max || ''));
  // Format Laporan Kunjungan DJP Harian (Web Telkomsel)
  return djpLaporanHarianSave_(rows, auth.user.username);
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

    // Realisasi: SF dianggap sudah ke ODP plan bila desa ODP itu tercatat dikunjungi SF tsb. di DJP_Realisasi_Baru
    const RB = djpRealBaruRead_(tgl), U = x => (x || '').toString().trim().toUpperCase();
    const isImport = {}; planRows.forEach(r => { if (normTgl(r[P.TGL]) === tgl && (r[P.DIBUAT_OLEH] || '').toString().trim() === DJP_IMPORT_TAG) isImport[U(r[P.USERNAME_SF]) + '|' + U(r[P.ODP])] = true; });
    const sfMap = {};
    let totalReal = 0;
    plan.forEach(p => {
      const su = U(p.username_sf);
      if (!sfMap[p.username_sf]) {
        const L = RB.sf[su + '|' + tgl] || null;
        sfMap[p.username_sf] = { username_sf: p.username_sf, nama_sf: p.nama_sf, odp_list: [], dikunjungi: L ? L.dik : null, target: L ? L.target : null };
      }
      const done = !!(RB.desa[su + '|' + U(p.kec) + '|' + U(p.desa) + '|' + tgl] || isImport[su + '|' + U(p.odp)]);
      if (done) totalReal++;
      sfMap[p.username_sf].odp_list.push({ odp: p.odp, cluster: p.cluster, kab: p.kab, kec: p.kec, desa: p.desa, lat: p.lat, lng: p.lng,
        status: done ? 'Dikunjungi' : 'Belum', realisasi: done ? 'Done' : '' });
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
          dikunjungi: sf.dikunjungi,
          target:     sf.target,
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
      total_real:  totalReal,
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

    // ── 2. Visit: ODP plan yang desanya tercatat dikunjungi SF tsb. (DJP_Realisasi_Baru) ──
    const RB = djpRealBaruRead_(bulan), U = x => (x || '').toString().trim().toUpperCase();
    if (wsPlan) {
      wsPlan.getDataRange().getValues().slice(1).forEach(r => {
        const tgl = djpNormTgl(r[P.TGL]); if (!tgl.startsWith(bulan)) return;
        const odp = (r[P.ODP] || '').toString().trim(); if (!odp) return;
        if ((r[P.DIBUAT_OLEH] || '').toString().trim() === DJP_IMPORT_TAG ||
            RB.desa[U(r[P.USERNAME_SF]) + '|' + U(r[P.KEC]) + '|' + U(r[P.DESA]) + '|' + tgl]) visitSet[odp] = true;
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
//  DJP_Realisasi_Baru — kunjungan nyata (1 baris = SF × desa × hari)
//  TGL|KODE_SF|NAMA_SF|PLAN|KAB|KEC|DESA|JML_ODP|DIKUNJUNGI|TARGET_DJP|SUMBER|UPLOADED_BY|UPLOADED_AT
//  PLAN = 'Ada Plan' / kosong. KAB/KEC/DESA hanya diisi bila Ada Plan DAN DIKUNJUNGI > 0 (= kunjungan desa).
// ================================================================
const RB_HEAD = ['TGL','KODE_SF','NAMA_SF','PLAN','KAB','KEC','DESA','JML_ODP','DIKUNJUNGI','TARGET_DJP','SUMBER','UPLOADED_BY','UPLOADED_AT'];
const RB_ADA = 'Ada Plan';
const RB_LAPORAN = 'Laporan Harian', RB_WPI = 'Source WPI';

function djpRealBaruSheet_(ssOpt) {
  const ss = ssOpt || SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);
  let ws = ss.getSheetByName(DJP_CONFIG.SHEET_REAL2);
  if (!ws) ws = ss.insertSheet(DJP_CONFIG.SHEET_REAL2);
  if (ws.getLastRow() === 0) {
    ws.getRange(1, 1, 1, RB_HEAD.length).setValues([RB_HEAD]).setFontWeight('bold');
    ws.setFrozenRows(1);
    ws.getRange('A:A').setNumberFormat('@');               // tanggal disimpan teks yyyy-MM-dd
  }
  return ws;
}
// Hapus baris yang cocok: filter(row) → true
function djpRealBaruDelete_(ws, match) {
  // Cepat: baca sekali, saring, tulis ulang sekali (tanpa deleteRows berulang)
  const n = ws.getLastRow() - 1; if (n < 1) return 0;
  const W = RB_HEAD.length, rg = ws.getRange(2, 1, n, W), v = rg.getValues();
  const keep = v.filter(r => !match(r)), hapus = v.length - keep.length;
  if (!hapus) return 0;
  rg.clearContent();
  if (keep.length) ws.getRange(2, 1, keep.length, W).setValues(keep);
  return hapus;
}
function djpRealBaruAppend_(ws, out) {
  for (let i = 0; i < out.length; i += DJP_CONFIG.BATCH_SIZE) {
    const b = out.slice(i, i + DJP_CONFIG.BATCH_SIZE);
    ws.getRange(ws.getLastRow() + 1, 1, b.length, RB_HEAD.length).setValues(b);
  }
}

// Upload laporan harian Web Telkomsel: rows = [{tgl, kode_sf, nama_sf, target, dikunjungi}]
// Desa diambil dari DJP_Plan (SF yang sama, tanggal yang sama).
function djpLaporanHarianSave_(rows, by) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { ok:false, error:'Server sedang sibuk, coba lagi.' }; }
  try {
    const U = x => (x || '').toString().trim().toUpperCase();
    const lap = {}, tglSet = {};
    rows.forEach(r => {
      const tgl = djpNormTgl(r.tgl), sf = U(r.kode_sf); if (!/^\d{4}-\d{2}-\d{2}$/.test(tgl) || !sf) return;
      lap[tgl + '|' + sf] = { tgl, sf, nama: (r.nama_sf || '').toString().trim(), target: +r.target || 0, dik: +r.dikunjungi || 0 };
      tglSet[tgl] = true;
    });
    const keys = Object.keys(lap); if (!keys.length) return { ok:false, error:'Tidak ada baris valid (Tanggal / Kode SF kosong).' };

    // Plan per (tgl|SF) → desa
    const plan = {};
    SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID).getSheetByName(DJP_CONFIG.SHEET_PLAN).getDataRange().getValues().slice(1).forEach(p => {
      const tgl = djpNormTgl(p[P.TGL]); if (!tglSet[tgl]) return;
      if ((p[P.DIBUAT_OLEH] || '').toString().trim() === DJP_IMPORT_TAG) return;
      const k = tgl + '|' + U(p[P.USERNAME_SF]); if (!lap[k]) return;
      const dk = U(p[P.KEC]) + '|' + U(p[P.DESA]);
      const g = (plan[k] = plan[k] || {}); g[dk] = g[dk] || { kab: U(p[P.KAB]), kec: (p[P.KEC] || '').toString().trim(), desa: (p[P.DESA] || '').toString().trim(), n: 0 }; g[dk].n++;
    });

    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'), out = [];
    let terplan = 0, tanpaPlan = 0, dikunjungi = 0;
    keys.forEach(k => {
      const L = lap[k], desas = plan[k] ? Object.values(plan[k]) : [];
      if (desas.length) terplan++; else tanpaPlan++;
      if (L.dik > 0) dikunjungi++;
      if (desas.length && L.dik > 0) desas.forEach(d => out.push([L.tgl, L.sf, L.nama, RB_ADA, d.kab, d.kec, d.desa, d.n, L.dik, L.target, RB_LAPORAN, by, now]));
      else out.push([L.tgl, L.sf, L.nama, desas.length ? RB_ADA : '', '', '', '', '', L.dik, L.target, RB_LAPORAN, by, now]);
    });

    const ws = djpRealBaruSheet_();
    // (tanggal, SF) yang sudah punya data Detail DJP tidak ditimpa laporan harian (detail lebih akurat)
    const adaDetail = {};
    if (ws.getLastRow() > 1) ws.getRange(2, 1, ws.getLastRow() - 1, 11).getValues().forEach(r => { if (r[10] === 'Detail DJP') adaDetail[djpNormTgl(r[0]) + '|' + U(r[1])] = true; });
    const outF = out.filter(r => !adaDetail[r[0] + '|' + r[1]]);
    out.length = 0; outF.forEach(r => out.push(r));
    const replaced = djpRealBaruDelete_(ws, r => r[10] === RB_LAPORAN && lap[djpNormTgl(r[0]) + '|' + U(r[1])]);
    djpRealBaruAppend_(ws, out);
    try { CacheService.getScriptCache().removeAll(Object.keys(tglSet).map(t => 'djp_daily_' + t)); } catch (e) {}
    return { ok:true, sf: keys.length, baris: out.length, replaced, terplan, tanpaPlan, dikunjungi };
  } catch (e) {
    return { ok:false, error:'djpLaporanHarianSave: ' + e.message };
  } finally { lock.releaseLock(); }
}

// ================================================================
//  SEKALI JALAN: SALIN baris "import Source WPI" dari DJP_Plan → DJP_Realisasi_Baru
//  Jalankan manual dari editor: pilih djpSalinWPIKeRealisasi → Jalankan.
//  Baris di DJP_Plan TIDAK dihapus (tetap dipakai rotasi ODP / Last Plan). Aman diulang.
// ================================================================
function djpSalinWPIKeRealisasi() {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const U = x => (x || '').toString().trim().toUpperCase();
    const v = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID).getSheetByName(DJP_CONFIG.SHEET_PLAN).getDataRange().getValues();
    const grp = {}, perSFHari = {};
    for (let i = 1; i < v.length; i++) {
      const p = v[i]; if ((p[P.DIBUAT_OLEH] || '').toString().trim() !== DJP_IMPORT_TAG) continue;
      const tgl = djpNormTgl(p[P.TGL]), sf = U(p[P.USERNAME_SF]); if (!tgl || !sf) continue;
      const k = tgl + '|' + sf + '|' + U(p[P.KEC]) + '|' + U(p[P.DESA]);
      grp[k] = grp[k] || { tgl, sf, nama: (p[P.NAMA_SF] || '').toString().trim(), kab: U(p[P.KAB]), kec: (p[P.KEC] || '').toString().trim(), desa: (p[P.DESA] || '').toString().trim(), n: 0 };
      grp[k].n++; perSFHari[tgl + '|' + sf] = (perSFHari[tgl + '|' + sf] || 0) + 1;
    }
    const g = Object.values(grp);
    if (!g.length) { Logger.log('Tidak ada baris import Source WPI di DJP_Plan.'); return 0; }
    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    const out = g.map(x => [x.tgl, x.sf, x.nama, RB_ADA, x.kab, x.kec, x.desa, x.n, perSFHari[x.tgl + '|' + x.sf], '', RB_WPI, 'import Source WPI', now]);
    const ws = djpRealBaruSheet_();
    const old = djpRealBaruDelete_(ws, r => r[10] === RB_WPI);   // aman diulang
    djpRealBaruAppend_(ws, out);
    Logger.log('Disalin: ' + out.length + ' baris (SF×desa×hari) ke DJP_Realisasi_Baru; ganti lama: ' + old + '. DJP_Plan tidak diubah.');
    return out.length;
  } finally { try { lock.releaseLock(); } catch (e) {} }
}


// ================================================================
//  SEKALI JALAN: ubah DJP_Realisasi_Baru format lama (tanpa kolom PLAN) → format baru
//  Sisipkan kolom PLAN setelah NAMA_SF, rapikan isi sesuai aturan:
//   - Ada Plan + Dikunjungi > 0 → KAB/KEC/DESA terisi
//   - Ada Plan + Dikunjungi 0   → PLAN = 'Ada Plan', KAB/KEC/DESA kosong
//   - Tanpa plan                → semua kosong
//  Jalankan sekali: pilih djpRealBaruTambahKolomPlan → Jalankan.
// ================================================================
function djpRealBaruTambahKolomPlan() {
  const ws = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID).getSheetByName(DJP_CONFIG.SHEET_REAL2);
  if (!ws) throw new Error('Tab DJP_Realisasi_Baru tidak ada');
  if ((ws.getRange(1, 4).getValue() || '').toString().trim().toUpperCase() === 'PLAN') { Logger.log('Kolom PLAN sudah ada.'); return; }
  ws.insertColumnAfter(3);
  ws.getRange(1, 1, 1, RB_HEAD.length).setValues([RB_HEAD]).setFontWeight('bold');
  const n = ws.getLastRow() - 1; if (n < 1) return;
  const rg = ws.getRange(2, 4, n, 4), v = rg.getValues();        // PLAN | KAB | KEC | DESA
  const dik = ws.getRange(2, 9, n, 1).getValues();
  const out = v.map((r, i) => {
    const kab = (r[1] || '').toString().trim(), desa = (r[3] || '').toString().trim();
    const adaPlan = !!(kab || desa);                               // format lama: KAB terisi = SF punya plan
    if (adaPlan && desa && (+dik[i][0] > 0)) return [RB_ADA, r[1], r[2], r[3]];
    return [adaPlan ? RB_ADA : '', '', '', ''];
  });
  rg.setValues(out);
  Logger.log('Kolom PLAN ditambahkan, ' + n + ' baris dirapikan.');
}


// ================================================================
//  Upload Detail DJP (per homepass, sudah dirangkum di web menjadi SF × desa × hari)
//  rows = [{tgl, kode_sf, nama_sf, kab, kec, desa, jml_odp, dikunjungi, target}]
//  Mengganti baris Laporan Harian / Detail DJP untuk (tanggal, SF) yang sama.
// ================================================================
const RB_DETAIL = 'Detail DJP';
function djpDetailSave_(rows, by, rMin, rMax) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return { ok:false, error:'Server sedang sibuk, coba lagi.' }; }
  try {
    const U = x => (x || '').toString().trim().toUpperCase();
    const keys = {}, tglSet = {};
    rows.forEach(r => { const t = djpNormTgl(r.tgl); if (t && r.kode_sf) { keys[t + '|' + U(r.kode_sf)] = true; tglSet[t] = true; } });
    if (!Object.keys(keys).length) return { ok:false, error:'Tidak ada baris valid.' };
    // SF punya plan web di tanggal itu?
    const ada = {};
    const ss = SpreadsheetApp.openById(DJP_CONFIG.SHEET_ID);           // buka sekali saja (spreadsheet besar)
    const wp = ss.getSheetByName(DJP_CONFIG.SHEET_PLAN), np = wp.getLastRow() - 1;
    if (np > 0) {
      const ab = wp.getRange(2, 1, np, 2).getValues(), by_ = wp.getRange(2, P.DIBUAT_OLEH + 1, np, 1).getValues();   // TGL, USERNAME_SF, DIBUAT_OLEH saja
      for (let i = 0; i < np; i++) {
        const t = djpNormTgl(ab[i][0]); if (!tglSet[t]) continue;
        if ((by_[i][0] || '').toString().trim() === DJP_IMPORT_TAG) continue;
        ada[t + '|' + U(ab[i][1])] = true;
      }
    }
    const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    let terplan = 0, dikunjungi = 0;
    const out = rows.filter(r => djpNormTgl(r.tgl) && r.kode_sf).map(r => {
      const t = djpNormTgl(r.tgl), sf = U(r.kode_sf), pl = ada[t + '|' + sf] ? RB_ADA : '';
      if (pl) terplan++; if (+r.dikunjungi > 0) dikunjungi++;
      return [t, sf, (r.nama_sf || '').toString().trim(), pl, U(r.kab), (r.kec || '').toString().trim(), (r.desa || '').toString().trim(),
              +r.jml_odp || 0, +r.dikunjungi || 0, +r.target || 0, RB_DETAIL, by, now];
    });
    const ws = djpRealBaruSheet_(ss);
    // File detail mewakili seluruh periodenya: data SF yang sama dalam rentang tanggal file diganti semua
    const sfSet = {}; Object.keys(keys).forEach(k => { sfSet[k.split('|')[1]] = true; });
    const inRange = t => rMin && rMax && t >= rMin && t <= rMax;
    const replaced = djpRealBaruDelete_(ws, r => { if (r[10] !== RB_LAPORAN && r[10] !== RB_DETAIL) return false;
      const t = djpNormTgl(r[0]), sf = U(r[1]); return keys[t + '|' + sf] || (sfSet[sf] && inRange(t)); });
    djpRealBaruAppend_(ws, out);
    try { CacheService.getScriptCache().removeAll(Object.keys(tglSet).map(t => 'djp_daily_' + t)); } catch (e) {}
    return { ok:true, baris: out.length, replaced, sf: Object.keys(keys).length, terplan, dikunjungi };
  } catch (e) {
    return { ok:false, error:'djpDetailSave: ' + e.message };
  } finally { lock.releaseLock(); }
}
