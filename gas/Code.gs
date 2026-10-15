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
      case 'performance': data = getPerformance(params);     break;
      case 'sf':          data = getSFList();               break;
      case 'absen':       data = getAbsen();                break;
      case 'kendala':     data = getKendala();              break;
      case 'wilayah':     data = { wilayah: WILAYAH };      break;
      case 'orders':      data = getOrders(params);          break;
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

  // MoM: same day-range, previous month
  const startLalu = shiftMonthStr(startIni, -1);
  const endLalu   = shiftMonthStr(endIni,   -1);

  const rangeIni  = { start: startIni, end: endIni };
  const rangeLalu = { start: startLalu, end: endLalu };

  const produkFilter  = (params.produk  || '').toString().trim().toLowerCase();
  const channelFilter = (params.channel || '').toString().trim().toLowerCase();

  const stoFilter = (params.sto || '').toString().trim().toUpperCase();

  const ini  = parseBulan('Bulan_Ini',  sfMap, rangeIni.start,  rangeIni.end,  produkFilter, channelFilter, stoFilter);
  const lalu = parseBulan('Bulan_Lalu', sfMap, rangeLalu.start, rangeLalu.end, produkFilter, channelFilter, stoFilter);

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
    wsaTrend:   wsa.trend,
    wilayah:    WILAYAH,
    kendalaList: ini.kendalaList || [],
    updatedAt:  new Date().toISOString(),
    _debug: {
      rangeIni:  { start: startIni,  end: endIni  },
      rangeLalu: { start: startLalu, end: endLalu },
      totalIni:  ini.total,
      totalLalu: lalu.total,
      byKabIni:  ini.total.byKab,
      byKabLalu: lalu.total.byKab,
    },
  };
}

// ── Parse Sheet dengan filter rentang tanggal (YYYYMMDD) ─────
// startStr & endStr: "YYYYMMDD" – apple-to-apple MoM comparison
function parseBulan(sheetName, sfMap, startStr, endStr, produkFilter, channelFilter, stoFilter) {
  produkFilter  = (produkFilter  || '').toLowerCase().trim();
  channelFilter = (channelFilter || '').toLowerCase().trim();
  stoFilter     = (stoFilter     || '').toUpperCase().trim();
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

    // Tanggal masing-masing metric (format YYYYMMDD)
    const getTgl = (col) => col !== -1 && r[col] ? r[col].toString().replace(/-/g, '').substring(0, 8) : '';
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

    const valIO = (C.io !== -1 ? r[C.io] : '').toString().toLowerCase();
    const valRE = (C.re !== -1 ? r[C.re] : '').toString().toLowerCase();
    const valPS = (C.ps !== -1 ? r[C.ps] : '').toString().toLowerCase();

    // ── Total semua channel ──
    if (countIO && valIO.includes('valid io')) {
      total.io++;
      addKab(total, kabRaw, 'io');
      if (!isAgency) digital.io++;
    }
    if (countRE && valRE.includes('valid re')) {
      total.re++;
      addKab(total, kabRaw, 're');
      if (!isAgency) digital.re++;
    }
    if (countPS && valPS.includes('put in service')) {
      total.ps++;
      addKab(total, kabRaw, 'ps');
      if (!isAgency) digital.ps++;
    }

    // ── Agency only ──
    if (isAgency) {
      if (countIO && valIO.includes('valid io')) { agency.io++; addKab(agency, kabRaw, 'io'); }
      if (countRE && valRE.includes('valid re')) { agency.re++; addKab(agency, kabRaw, 're'); }
      if (countPS && valPS.includes('put in service')) { agency.ps++; addKab(agency, kabRaw, 'ps'); }
    }

    // ── Per-SF (agency only) ──
    if (!isAgency) continue;
    const rawKode = (C.codeSF !== -1 ? r[C.codeSF] : '') || '';
    const kode    = rawKode.toString().trim().toLowerCase();
    if (!kode || kode === '#n/a') continue;

    if (!bySF[kode]) bySF[kode] = { io: 0, re: 0, ps: 0, kendala: 0, produk: {}, kab: {} };

    if (countIO && valIO.includes('valid io')) bySF[kode].io++;
    if (countRE && valRE.includes('valid re')) bySF[kode].re++;
    if (countPS && valPS.includes('put in service')) {
      bySF[kode].ps++;
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

  // Kumpulkan unique kendala dari semua rows
  const kendalaSet = new Set();
  if (C.kendalaKat !== -1) {
    for (let i = 1; i < rows.length; i++) {
      const v = rows[i][C.kendalaKat];
      if (v) { const s = v.toString().trim(); if (s) kendalaSet.add(s); }
    }
  }

  return { bySF, total, agency, digital, kendalaList: [...kendalaSet].sort() };
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
      const idx = hdr.findIndex(h => h.toString().toLowerCase() === n.toLowerCase());
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
    const s = val.toString().replace(/-/g, '').substring(0, 8);
    if (s.length === 8) return s.substring(0,4)+'-'+s.substring(4,6)+'-'+s.substring(6,8);
    return s;
  }

  const orders = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.status]) continue;

    // Date filter on Tanggal IO
    const tglIO = C.tanggalIO !== -1 ? r[C.tanggalIO] : '';
    const tglStr = tglIO ? tglIO.toString().replace(/-/g, '').substring(0, 8) : '';
    if (startStr && tglStr && tglStr < startStr) continue;
    if (endStr   && tglStr && tglStr > endStr)   continue;

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
    const clusterBySTO = STO_TO_CLUSTER[stoRaw] || cluster;

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
      cluster:   clusterBySTO,
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
