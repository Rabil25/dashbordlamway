// Konfigurasi, aturan wilayah, kolom data, helper periode & tier — dipisah dari index.html (Nov 2026).
// Script klasik (bukan module): konstanta & fungsi di sini global dan dipakai file berikutnya.
// Urutan muat: config.js → app.js → djp.js → Alpine (defer).
// ============================================================
//  CONFIG — ganti sesuai kebutuhan
// ============================================================
const DEFAULT_GAS_URL  = 'https://script.google.com/macros/s/AKfycbzBskwSjBBScoZPdwkZxKNJ9PcN5Yihb-8rVAwq26O2dodfraHxxClqlL379ryM04jupA/exec';
const DEFAULT_VERCEL_URL = 'https://dashbordlamway.vercel.app/api/data';

// Pilih URL terbaik: data endpoint pakai Vercel jika tersedia, auth tetap GAS
function getDataUrl(cfg) {
  const vercelUrl = (cfg && cfg.vercelUrl) || DEFAULT_VERCEL_URL;
  return vercelUrl ? vercelUrl : ((cfg && cfg.gasUrl) || DEFAULT_GAS_URL);
}

// Kolom index dalam array (CE=0, CF=1, ...)
const C = {
  STATUS_ORDER: 0,   // CE
  WOK:          1,   // CF
  STO:          2,   // CG
  PRODUK_KAT:   3,   // CH
  VAL_IO:       5,   // CJ
  BLN_IO:       6,   // CK
  TGL_IO:       7,   // CL
  VAL_RE:       8,   // CM
  BLN_RE:       9,   // CN
  TGL_RE:      10,   // CO
  VAL_PS:      11,   // CP
  BLN_PS:      12,   // CQ
  TGL_PS:      13,   // CR
  CODE_SF:     14,   // CS
  NAMA_SF:     15,   // CT
  CHANNEL2:    16,   // CU (Agency/Digital)
  KODE_IO:     17,   // CV
  VAL_CHANNEL: 18,   // CW
  STATUS_FNL:  19,   // CX
  KENDALA_DTL: 20,   // CY
  KENDALA_CAT: 21,   // CZ
  IO_TIMES:    23,   // DB
  TIME_RANGE:  24,   // DC
  DATE_IO:     29,   // DH (yyyymmdd)
  DATE_RE:     30,   // DI (yyyymmdd)
  DATE_PS:     31,   // DJ (yyyymmdd)
  ID_ORDER:    34,   // DM
  TGL_KENDALA: 35,   // DN
  HOMEPASS:    36,   // DO
  DEDUP:       37,   // DP ← wajib "Unik"
  VAL_KAB:     38,   // DQ ← cluster filter
  KOORDINAT:   39,   // DR
  SPEED:       41,   // DT
};

// ── odp.json dimuat SEKALI, dipakai bersama semua halaman (sebelumnya 5 kali fetch) ──
let _odpPromise = null;
function getOdpJson() {
  if (!_odpPromise) {
    _odpPromise = fetch('data/odp.json').then(r => { if (!r.ok) throw new Error('odp.json HTTP ' + r.status); return r.json(); })
      .catch(e => { _odpPromise = null; throw e; });
  }
  return _odpPromise;
}
// Bungkus agar kode lama yang memanggil res.json() tetap jalan
function fetchOdpJson() { return getOdpJson().then(d => ({ ok: true, json: () => Promise.resolve(d) })); }

// Escape teks dari sheet sebelum masuk HTML (pop-up peta) — cegah injeksi HTML
function esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ══ PERIODE = 'YYYY-MM' ══════════════════════════════════════════
// Kolom Bulan IO/RE/PS (CK/CN/CQ) di sheet hanya berisi angka bulan tanpa tahun.
// Tahun diambil dari kolom Tanggal yyyymmdd (DH/DI/DJ = C.DATE_IO/RE/PS).
// Saat data dimuat, normalizePeriods() mengganti isi kolom Bulan di setiap baris
// dengan 'YYYY-MM', sehingga semua perbandingan "bulan baris === bulan dipilih"
// di seluruh halaman otomatis memperhitungkan tahun.
const BULAN_NAMA = ['','Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
function currentPeriod() { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0'); }
function periodYear(p)  { return parseInt(String(p).slice(0,4)) || new Date().getFullYear(); }
function periodMonth(p) { return parseInt(String(p).slice(5,7)) || (new Date().getMonth()+1); }
function daysInPeriod(p) { return new Date(periodYear(p), periodMonth(p), 0).getDate(); }
function periodDate(p, day) { return new Date(periodYear(p), periodMonth(p)-1, day); }
function prevPeriod(p) { const y = periodYear(p), m = periodMonth(p); return m === 1 ? (y-1) + '-12' : y + '-' + String(m-1).padStart(2,'0'); }
function periodLabel(p) { return (BULAN_NAMA[periodMonth(p)] || '?') + ' ' + periodYear(p); }
// 'YYYYMMDD' (teks atau angka) → 'YYYY-MM'; kosong bila tidak valid
function ymdToPeriod(v) { const s = String(v ?? '').trim(); return /^20\d{2}(0[1-9]|1[0-2])\d{2}$/.test(s) ? s.slice(0,4) + '-' + s.slice(4,6) : ''; } // abaikan 18991230 (tanggal kosong Excel)
function normalizePeriods(rows) {
  for (const r of rows) {
    if (r._p) continue;
    r[C.BLN_IO] = ymdToPeriod(r[C.DATE_IO]);
    r[C.BLN_RE] = ymdToPeriod(r[C.DATE_RE]);
    r[C.BLN_PS] = ymdToPeriod(r[C.DATE_PS]);
    r._p = 1;
  }
  return rows;
}

// Cluster config
const CLUSTER = {
  'Lampung Barat': { kab: ['LAMPUNG BARAT','TANGGAMUS','PESISIR BARAT'] },
  'Way Kanan':     { kab: ['LAMPUNG UTARA','WAY KANAN'] },
};

// Cluster ditentukan oleh KABUPATEN lokasi ODP, bukan label STO/cluster di data ODP
// (mis. ODP-BKM di Kab. Lampung Barat → cluster Lampung Barat). Hasil: 'LAMPUNG BARAT' / 'WAY KANAN' / ''.
function clusterOfKab(kab) {
  const k = (kab || '').toString().trim().toUpperCase();
  const name = Object.keys(CLUSTER).find(c => CLUSTER[c].kab.includes(k));
  return name ? name.toUpperCase() : '';
}

// ── Kantor referensi untuk hitung jarak prioritas ──
const OFFICES = [
  { name: 'Tanggamus 1',    lat: -5.3664321, lng: 104.7803285, kab: 'TANGGAMUS' },
  { name: 'Tanggamus 2',    lat: -5.4933735, lng: 104.6186705, kab: 'TANGGAMUS' },
  { name: 'Pesisir Barat',  lat: -5.1962234, lng: 103.9309363, kab: 'PESISIR BARAT' },
  { name: 'Lampung Barat',  lat: -5.0294173, lng: 104.0743091, kab: 'LAMPUNG BARAT' },
  { name: 'Way Kanan',      lat: -4.7268253, lng: 104.5425645, kab: 'WAY KANAN' },
  { name: 'Lampung Utara 1',lat: -4.8308673, lng: 104.8889064, kab: 'LAMPUNG UTARA' },
  { name: 'Lampung Utara 2',lat: -4.8610996, lng: 104.5734521, kab: 'LAMPUNG UTARA' },
];

function haversineKm(lat1, lng1, lat2, lng2) {
  const R = 6371;
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat/2)**2 + Math.cos(lat1*Math.PI/180) * Math.cos(lat2*Math.PI/180) * Math.sin(dLng/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

// ── ECharts theme helpers ──────────────────────────────────────
function isLight() { return document.documentElement.getAttribute('data-theme') === 'light'; }

function getChartTheme() {
  const light = isLight();
  return {
    backgroundColor: 'transparent',
    textStyle: { fontFamily:"'Poppins',sans-serif", fontSize:11, color: light ? '#6b4f42' : '#8b949e' },
    color: light
      ? ['#c93c20','#27863a','#0e7a9e','#d97020','#9b2563','#6b5e00']
      : ['#00c2ff','#22d972','#ff4d4d','#fbbf24','#f472b6','#a855f7'],
    grid: { left:10, right:10, top:30, bottom:10, containLabel:true },
    legend: {
      textStyle: { color: light ? '#6b4f42' : '#8b949e', fontFamily:"'Poppins',sans-serif", fontSize:11 },
      itemWidth:12, itemHeight:8,
    },
    tooltip: {
      backgroundColor: light ? '#fff8f5' : '#1c2128',
      borderColor:     light ? '#e8d0c5' : '#30363d',
      borderWidth:1,
      textStyle:{ color: light ? '#1c0e08' : '#e6edf3', fontFamily:"'Poppins',sans-serif", fontSize:12 },
      padding:[8,12],
    },
    xAxis: {
      axisLine:{show:false}, axisTick:{show:false},
      axisLabel:{ color: light ? '#8a6050' : '#484f58', fontFamily:"'Poppins',sans-serif", fontSize:11 },
      splitLine:{show:false}
    },
    yAxis: {
      axisLine:{show:false}, axisTick:{show:false},
      axisLabel:{ color: light ? '#8a6050' : '#484f58', fontFamily:"'Poppins',sans-serif", fontSize:11 },
      splitLine:{ lineStyle:{ color: light ? 'rgba(180,120,90,.14)' : '#21262d' } }
    },
  };
}

// Static refs kept for compatibility — will be overridden per render
const ECHARTS_THEME = getChartTheme();
const PIE_THEME = (function(){const {xAxis,yAxis,grid,...t}=getChartTheme(); return t;})();

// ── SF TIER CONFIG ──────────────────────────────────────────
const TIER_CONFIG = [
  { name:'Black',    icon:'⬛', minPS:0,  maxPS:0,  target:5,  action:'Counseling', rangeLabel:'0'    },
  { name:'Bronze',   icon:'🥉', minPS:1,  maxPS:9,  target:15, action:'Couching',   rangeLabel:'1–9'  },
  { name:'Silver',   icon:'🥈', minPS:10, maxPS:15, target:20, action:'On Track',   rangeLabel:'10–15'},
  { name:'Gold',     icon:'🥇', minPS:16, maxPS:25, target:30, action:'On Track',   rangeLabel:'16–25'},
  { name:'Platinum', icon:'💎', minPS:26, maxPS:50, target:50, action:'On Track',   rangeLabel:'26–50'},
  { name:'Diamond',  icon:'💠', minPS:51, maxPS:Infinity, target:65, action:'On Track', rangeLabel:'>50'},
];

function getSFTier(prevPS) {
  for (const t of TIER_CONFIG) {
    if (prevPS >= t.minPS && prevPS <= t.maxPS) return t;
  }
  return TIER_CONFIG[0];
}

// Hitung action (On Track / Couching / Counseling) berdasarkan pace harian
// A = target/30 * hariIni (ekspektasi PS sampai hari ini)
// B = ps bulan ini (realisasi)
// B >= A → On Track; gap <= 20% → Couching; gap > 20% → Counseling
function getSFAction(ps, target, hariIni) {
  const A = (target / 30) * hariIni;
  if (A <= 0) return 'On Track';
  if (ps >= A) return 'On Track';
  const gap = (A - ps) / A * 100;
  return gap <= 20 ? 'Couching' : 'Counseling';
}


// ============================================================
//  FILTER CACHE — non-reactive closure variables
//  filteredRows dipanggil 15-20x per render; cache ini memastikan
//  filter 885 baris hanya dijalankan 1x per perubahan cluster/data
// ============================================================
let _frCache = null, _frKey = '';
let _loadSeq = 0; // naik setiap allRows diganti → cache filter ikut diperbarui
let _fprCache = null, _fprKey = '';

// ══ KOMPONEN BERSAMA: toast & dialog konfirmasi ═══════════════════
// Satu-satunya cara menampilkan notifikasi di aplikasi ini.
// type: 'ok' | 'err' | 'warn' | 'info' (alias lama: success→ok, error→err)
function showToast(msg, type = 'ok', ms) {
  type = { success: 'ok', error: 'err' }[type] || type;
  const wrap = document.getElementById('toast-wrap');
  if (!wrap) return;
  const el = document.createElement('div');
  el.className = 'toast toast-' + type;
  el.setAttribute('role', type === 'err' ? 'alert' : 'status');
  el.textContent = msg;
  wrap.appendChild(el);
  const dur = ms || (type === 'err' ? 6000 : 3500);   // error tampil lebih lama
  setTimeout(() => { el.style.opacity = '0'; setTimeout(() => el.remove(), 300); }, dur);
}

// Dialog konfirmasi pengganti confirm() bawaan browser.
// Pemakaian: if (!(await uiConfirm({ title, message, confirmText: 'Hapus', danger: true }))) return;
function uiConfirm({ title = 'Konfirmasi', message = '', confirmText = 'Ya', cancelText = 'Batal', danger = false } = {}) {
  return new Promise(resolve => {
    const prevFocus = document.activeElement;
    const bg = document.createElement('div');
    bg.className = 'modal-bg';
    bg.innerHTML =
      '<div class="modal ui-dialog" role="dialog" aria-modal="true" aria-labelledby="ui-dlg-t">' +
        '<div class="modal-title" id="ui-dlg-t"></div><div class="modal-sub"></div>' +
        '<div class="modal-actions">' +
          '<button type="button" class="btn" data-a="0"></button>' +
          '<button type="button" class="btn ' + (danger ? 'btn-danger btn-solid' : 'btn-primary') + '" data-a="1"></button>' +
        '</div></div>';
    bg.querySelector('.modal-title').textContent = title;   // textContent: aman dari injeksi
    bg.querySelector('.modal-sub').textContent = message;
    bg.querySelector('[data-a="0"]').textContent = cancelText;
    bg.querySelector('[data-a="1"]').textContent = confirmText;
    const done = ok => { document.removeEventListener('keydown', onKey); bg.remove(); if (prevFocus && prevFocus.focus) prevFocus.focus(); resolve(ok); };
    const onKey = e => { if (e.key === 'Escape') done(false); };
    bg.addEventListener('click', e => { if (e.target === bg) done(false); const a = e.target.closest('[data-a]'); if (a) done(a.dataset.a === '1'); });
    document.addEventListener('keydown', onKey);
    document.body.appendChild(bg);
    bg.querySelector('[data-a="0"]').focus();   // fokus default ke Batal (aman)
  });
}
