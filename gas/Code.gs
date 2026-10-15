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
const WILAYAH = {
  'Lampung Barat': {
    tl:        'ANDRES WELSEN',
    kabupaten: ['LAMPUNG BARAT', 'TANGGAMUS', 'PESISIR BARAT'],
    sto:       ['LIW', 'KRU', 'KOT', 'TGP'],
  },
  'Way Kanan': {
    tl:        'AGUS SETIAWAN',
    kabupaten: ['LAMPUNG UTARA', 'WAY KANAN'],
    sto:       ['BKM', 'KBI', 'BDU'],
  },
};

// Auto-build reverse lookup
const KAB_TO_CLUSTER = {};
const STO_TO_CLUSTER = {};
const TL_TO_CLUSTER  = {};
Object.entries(WILAYAH).forEach(([cluster, cfg]) => {
  cfg.kabupaten.forEach(k => { KAB_TO_CLUSTER[k] = cluster; });
  cfg.sto.forEach(s => { STO_TO_CLUSTER[s]       = cluster; });
  TL_TO_CLUSTER[cfg.tl]                           = cluster;
});

// ── Entry Point ──────────────────────────────────────────────
function doGet(e) {
  const params = e.parameter || {};
  if (params.token !== CONFIG.TOKEN) {
    return jsonOut({ ok: false, error: 'Unauthorized' });
  }
  try {
    const action = params.action || 'performance';
    let data;
    switch (action) {
      case 'performance': data = getPerformance();          break;
      case 'sf':          data = getSFList();               break;
      case 'absen':       data = getAbsen();                break;
      case 'kendala':     data = getKendala();              break;
      case 'wilayah':     data = { wilayah: WILAYAH };      break;
      default:            data = { error: 'Unknown action' };
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

// ── Hitung Hari Kerja (Senin-Jumat) ─────────────────────────
function calcHariKerja(now) {
  const tz   = CONFIG.TIMEZONE;
  const year = parseInt(Utilities.formatDate(now, tz, 'yyyy'));
  const mon  = parseInt(Utilities.formatDate(now, tz, 'MM'));
  const day  = parseInt(Utilities.formatDate(now, tz, 'dd'));

  // Hari kerja senin-jumat
  function isHK(d) { const w = d.getDay(); return w >= 1 && w <= 5; }

  const firstDay = new Date(year, mon - 1, 1);
  const lastDay  = new Date(year, mon, 0); // last day of month

  let total = 0, elapsed = 0;
  for (let d = new Date(firstDay); d <= lastDay; d.setDate(d.getDate() + 1)) {
    if (isHK(d)) {
      total++;
      if (d.getDate() <= day) elapsed++;
    }
  }
  return { total, elapsed, month: mon, year };
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
      cluster: TL_TO_CLUSTER[tlUpper] || STO_TO_CLUSTER[stoUpper] || '-',
      city:    r[C.city],
      status:  r[C.status],
      level:   r[C.level],
      agency:  r[C.agency],
    });
  }
  return { sfList: list };
}

// ── Performance Utama ────────────────────────────────────────
function getPerformance() {
  const now = new Date();
  const { sfList } = getSFList();
  const sfMap = {};
  sfList.forEach(sf => { sfMap[sf.kode.toLowerCase()] = sf; });

  // Apple-to-apple: bulan ini day 1–N vs bulan lalu day 1–N
  const rangeIni  = getDateRange(now,  0);
  const rangeLalu = getDateRange(now, -1);

  const ini  = parseBulan('Bulan_Ini', sfMap, rangeIni.start,  rangeIni.end);
  const lalu = parseBulan('Bulan_Ini', sfMap, rangeLalu.start, rangeLalu.end);

  const hariKerja = calcHariKerja(now);
  const wsa       = calcWSATSA(sfMap);

  const allKodes = new Set([
    ...Object.keys(ini.bySF),
    ...Object.keys(lalu.bySF),
    ...sfList.filter(s => s.status === 'ACTIVE').map(s => s.kode.toLowerCase()),
  ]);

  const sfPerf = [];
  allKodes.forEach(kode => {
    const info = sfMap[kode] || {};
    const i    = ini.bySF[kode]  || { io: 0, re: 0, ps: 0, kendala: 0, produk: {}, kab: {} };
    const l    = lalu.bySF[kode] || { io: 0, re: 0, ps: 0, kendala: 0, produk: {} };
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
      },
      bulanLalu: {
        io:  l.io,
        re:  l.re,
        ps:  l.ps,
        wsa: w.bulanLalu,
      },
    });
  });

  sfPerf.sort((a, b) => b.bulanIni.ps - a.bulanIni.ps);

  return {
    sfPerformance: sfPerf,
    rekapSTO:      rekapBySTO(sfPerf),
    rekapCluster:  rekapByCluster(sfPerf),
    rekapKab:      rekapByKabFull(ini, lalu),
    summary: {
      bulanIni:     ini.total,
      bulanLalu:    lalu.total,
      digitalIni:   ini.digital,
      digitalLalu:  lalu.digital,
      agencyIni:    ini.agency,
      agencyLalu:   lalu.agency,
    },
    hariKerja,
    wsaTrend:  wsa.trend,
    wilayah:   WILAYAH,
    updatedAt: new Date().toISOString(),
  };
}

// ── Parse Sheet dengan filter rentang tanggal (YYYYMMDD) ─────
// startStr & endStr: "YYYYMMDD" – apple-to-apple MoM comparison
function parseBulan(sheetName, sfMap, startStr, endStr) {
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
      const idx = hdr.findIndex(h => h.toString().toLowerCase() === n.toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const C = {
    status:    findCol(['Status Order']),
    io:        findCol(['Validasi IO']),
    tanggalIO: findCol(['Tanggal IO']),
    re:        findCol(['Validasi RE']),
    ps:        findCol(['Validasi PS']),
    codeSF:    findCol(['CODE SF', 'Validasi SF']),
    kendala:   findCol(['Kategori Kendala']),
    produk:    findCol(['Validasi Produk']),
    kab:       findCol(['VALIDASI KAB']),
    channel2:  findCol(['Channel 2']),
  };

  function mkStat() { return { io: 0, re: 0, ps: 0, byKab: {} }; }
  function addKab(stat, kab, field) {
    if (!kab) return;
    if (!stat.byKab[kab]) stat.byKab[kab] = { io: 0, re: 0, ps: 0 };
    stat.byKab[kab][field]++;
  }

  const bySF    = {};
  const total   = mkStat();
  const agency  = mkStat();
  const digital = { io: 0, re: 0, ps: 0 };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.status]) continue;

    // Filter rentang tanggal berdasarkan Tanggal IO (format YYYYMMDD)
    if (startStr && C.tanggalIO !== -1) {
      const tgl = r[C.tanggalIO] ? r[C.tanggalIO].toString().replace(/-/g, '').substring(0, 8) : '';
      if (!tgl || tgl < startStr || tgl > endStr) continue;
    }

    const ch       = (C.channel2 !== -1 ? r[C.channel2] : '').toString().trim().toLowerCase();
    const isAgency = ch === 'agency';
    const kabRaw   = C.kab !== -1 && r[C.kab] ? r[C.kab].toString().trim().toUpperCase() : '';

    const valIO = (C.io !== -1 ? r[C.io] : '').toString().toLowerCase();
    const valRE = (C.re !== -1 ? r[C.re] : '').toString().toLowerCase();
    const valPS = (C.ps !== -1 ? r[C.ps] : '').toString().toLowerCase();

    // ── Total semua channel ──
    if (valIO.includes('valid io')) {
      total.io++;
      addKab(total, kabRaw, 'io');
      if (!isAgency) digital.io++;
    }
    if (valRE.includes('valid re')) {
      total.re++;
      addKab(total, kabRaw, 're');
      if (!isAgency) digital.re++;
    }
    if (valPS.includes('put in service')) {
      total.ps++;
      addKab(total, kabRaw, 'ps');
      if (!isAgency) digital.ps++;
    }

    // ── Agency only ──
    if (isAgency) {
      if (valIO.includes('valid io')) { agency.io++; addKab(agency, kabRaw, 'io'); }
      if (valRE.includes('valid re')) { agency.re++; addKab(agency, kabRaw, 're'); }
      if (valPS.includes('put in service')) { agency.ps++; addKab(agency, kabRaw, 'ps'); }
    }

    // ── Per-SF (agency only) ──
    if (!isAgency) continue;
    const rawKode = (C.codeSF !== -1 ? r[C.codeSF] : '') || '';
    const kode    = rawKode.toString().trim().toLowerCase();
    if (!kode || kode === '#n/a') continue;

    if (!bySF[kode]) bySF[kode] = { io: 0, re: 0, ps: 0, kendala: 0, produk: {}, kab: {} };

    if (valIO.includes('valid io')) bySF[kode].io++;
    if (valRE.includes('valid re')) bySF[kode].re++;
    if (valPS.includes('put in service')) {
      bySF[kode].ps++;
      if (kabRaw) bySF[kode].kab[kabRaw] = (bySF[kode].kab[kabRaw] || 0) + 1;
    }
    if (C.kendala !== -1 && r[C.kendala]) bySF[kode].kendala++;
    if (C.produk  !== -1 && r[C.produk]) {
      const p = r[C.produk].toString().trim();
      if (p) bySF[kode].produk[p] = (bySF[kode].produk[p] || 0) + 1;
    }
  }

  return { bySF, total, agency, digital };
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

// ── Rekap per STO ────────────────────────────────────────────
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

  const now       = new Date();
  const thisMonth = now.getMonth();
  const thisYear  = now.getFullYear();
  const lastMonth = thisMonth === 0 ? 11 : thisMonth - 1;
  const lastYear  = thisMonth === 0 ? thisYear - 1 : thisYear;

  const bySF  = {};
  const trend = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;

    let d;
    try { d = new Date(r[0]); } catch(e) { continue; }
    if (isNaN(d)) continue;

    const yr = d.getFullYear(), mo = d.getMonth();
    const dateStr = fmtDate(d).slice(0, 7);

    const kode = (r[1] || '').toString().trim().toLowerCase();
    if (!kode) continue;

    const wsa = Number(r[2]) || 0;

    if (!bySF[kode]) bySF[kode] = { bulanIni: 0, bulanLalu: 0 };

    if (yr === thisYear  && mo === thisMonth) bySF[kode].bulanIni  += wsa;
    if (yr === lastYear  && mo === lastMonth) bySF[kode].bulanLalu += wsa;

    if (!trend[dateStr]) trend[dateStr] = 0;
    trend[dateStr] += wsa;
  }

  const trendArr = Object.entries(trend)
    .sort(([a], [b]) => a.localeCompare(b))
    .slice(-12)
    .map(([date, count]) => ({ date, count }));

  return { bySF, trend: trendArr };
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
        deal:      r[C.deal] ? 1 : 0,
      });
    }
  }

  const trendArr = Object.entries(trend)
    .sort(([a], [b]) => a.localeCompare(b))
    .slice(-30)
    .map(([date, count]) => ({ date, count }));

  return { today, trend: trendArr, count: today.length };
}
