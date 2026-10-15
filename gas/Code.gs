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
// Ubah di sini kalau ada reorganisasi wilayah
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
  const { sfList } = getSFList();
  const sfMap = {};
  sfList.forEach(sf => { sfMap[sf.kode.toLowerCase()] = sf; });

  const ini  = parseBulan('Bulan_Ini',  sfMap);
  const lalu = parseBulan('Bulan_Lalu', sfMap);
  const wsa  = calcWSATSA(sfMap);

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
        kab:     i.kab,      // PS per kabupaten (dari kolom VALIDASI KAB)
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
    rekapKab:      rekapByKab(sfPerf),
    summary: {
      bulanIni:    ini.total,
      bulanLalu:   lalu.total,
      digitalIni:  ini.digital,
      digitalLalu: lalu.digital,
    },
    wsaTrend:  wsa.trend,
    wilayah:   WILAYAH,
    updatedAt: new Date().toISOString(),
  };
}

// ── Parse Sheet Bulan_Ini / Bulan_Lalu ───────────────────────
function parseBulan(sheetName, sfMap) {
  const ws = ss().getSheetByName(sheetName);
  if (!ws) return {
    bySF:    {},
    total:   { io: 0, re: 0, ps: 0, kab: {} },
    digital: { io: 0, re: 0, ps: 0 },
  };

  const rows = ws.getDataRange().getValues();
  const hdr  = rows[0];

  // Case-insensitive column finder
  const findCol = (names) => {
    for (const n of names) {
      const idx = hdr.findIndex(h => h.toString().toLowerCase() === n.toLowerCase());
      if (idx !== -1) return idx;
    }
    return -1;
  };

  const C = {
    status:   findCol(['Status Order']),
    io:       findCol(['Validasi IO']),
    re:       findCol(['Validasi RE']),
    ps:       findCol(['Validasi PS']),
    codeSF:   findCol(['CODE SF', 'Validasi SF']),
    kendala:  findCol(['Kategori Kendala']),
    produk:   findCol(['Validasi Produk']),
    // Kolom CW di spreadsheet = header "VALIDASI KAB" (bukan Calendar Week)
    kab:      findCol(['VALIDASI KAB']),
    // Kolom CI = "Channel 2": nilai "Agency" = SF/Agency, selain itu = Digital
    channel2: findCol(['Channel 2']),
  };

  const bySF    = {};
  const total   = { io: 0, re: 0, ps: 0, kab: {} };
  const digital = { io: 0, re: 0, ps: 0 };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[C.status]) continue;

    // Tentukan channel: Agency (SF) vs Digital
    const ch        = (C.channel2 !== -1 ? r[C.channel2] : '').toString().trim().toLowerCase();
    const isAgency  = ch === 'agency';
    const isDigital = !isAgency;

    const valIO = (C.io !== -1 ? r[C.io] : '').toString().toLowerCase();
    const valRE = (C.re !== -1 ? r[C.re] : '').toString().toLowerCase();
    const valPS = (C.ps !== -1 ? r[C.ps] : '').toString().toLowerCase();

    // ── Hitung total semua baris (SF + Digital) ──
    if (valIO.includes('valid io')) { total.io++; if (isDigital) digital.io++; }
    if (valRE.includes('valid re')) { total.re++; if (isDigital) digital.re++; }
    if (valPS.includes('put in service')) {
      total.ps++;
      if (isDigital) digital.ps++;
      // Distribusi PS per kabupaten (semua channel)
      if (C.kab !== -1 && r[C.kab]) {
        const kabVal = r[C.kab].toString().trim().toUpperCase();
        if (kabVal) {
          total.kab[kabVal] = (total.kab[kabVal] || 0) + 1;
        }
      }
    }

    // ── Per-SF hanya untuk baris Agency ──
    if (!isAgency) continue;

    const rawKode = (C.codeSF !== -1 ? r[C.codeSF] : '') || '';
    const kode    = rawKode.toString().trim().toLowerCase();
    if (!kode || kode === '#n/a') continue;

    if (!bySF[kode]) {
      bySF[kode] = { io: 0, re: 0, ps: 0, kendala: 0, produk: {}, kab: {} };
    }

    if (valIO.includes('valid io')) bySF[kode].io++;
    if (valRE.includes('valid re')) bySF[kode].re++;
    if (valPS.includes('put in service')) {
      bySF[kode].ps++;
      if (C.kab !== -1 && r[C.kab]) {
        const kabVal = r[C.kab].toString().trim().toUpperCase();
        if (kabVal) {
          bySF[kode].kab[kabVal] = (bySF[kode].kab[kabVal] || 0) + 1;
        }
      }
    }

    if (C.kendala !== -1 && r[C.kendala]) bySF[kode].kendala++;

    if (C.produk !== -1 && r[C.produk]) {
      const p = r[C.produk].toString().trim();
      if (p) bySF[kode].produk[p] = (bySF[kode].produk[p] || 0) + 1;
    }
  }

  return { bySF, total, digital };
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

// ── Rekap per Kabupaten ──────────────────────────────────────
function rekapByKab(sfPerf) {
  const map = {};
  sfPerf.forEach(sf => {
    Object.entries(sf.bulanIni.kab || {}).forEach(([kab, ps]) => {
      const cluster = KAB_TO_CLUSTER[kab] || 'Tidak Dikenal';
      if (!map[kab]) map[kab] = { kabupaten: kab, cluster, bulanIni: { ps: 0 } };
      map[kab].bulanIni.ps += ps;
    });
  });
  return Object.values(map).sort((a, b) => b.bulanIni.ps - a.bulanIni.ps);
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
    const d = new Date(r[0]);
    if (isNaN(d)) continue;

    const kode    = (r[3] || '').toString().trim().toLowerCase();
    const dateStr = fmtDate(d);
    const m = d.getMonth(), y = d.getFullYear();

    if (!bySF[kode]) bySF[kode] = { bulanIni: 0, bulanLalu: 0 };

    if (m === thisMonth && y === thisYear) {
      bySF[kode].bulanIni++;
      trend[dateStr] = (trend[dateStr] || 0) + 1;
    } else if (m === lastMonth && y === lastYear) {
      bySF[kode].bulanLalu++;
    }
  }

  const trendArr = Object.entries(trend)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([date, count]) => ({ date, count }));

  return { bySF, trend: trendArr };
}

// ── Kendala ──────────────────────────────────────────────────
function getKendala() {
  const wsK  = ss().getSheetByName('Kendala');
  const wsFU = ss().getSheetByName('FUKendala');

  const kendala = [];
  if (wsK) {
    const rows = wsK.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      kendala.push({
        tanggal:    fmtDate(r[0]),
        teknisi:    (r[1] || '').toString().trim(),
        orderId:    (r[2] || '').toString().trim(),
        jenis:      (r[3] || '').toString().trim(),
        detail:     (r[4] || '').toString().trim(),
        koordinat:  (r[5] || '').toString().trim(),
        keterangan: (r[6] || '').toString().trim(),
      });
    }
    kendala.sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  }

  const fuKendala = [];
  if (wsFU) {
    const rows = wsFU.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      const stoUpper = (r[2] || '').toString().trim().toUpperCase();
      fuKendala.push({
        tanggal: fmtDate(r[0]),
        orderId: (r[1] || '').toString().trim(),
        sto:     stoUpper,
        cluster: STO_TO_CLUSTER[stoUpper] || '-',
        kendala: (r[3] || '').toString().trim(),
        solusi:  (r[4] || '').toString().trim(),
        wsaTsa:  (r[5] || '').toString().trim(),
      });
    }
    fuKendala.sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  }

  const byJenis = {};
  kendala.forEach(k => {
    const j = (k.jenis || 'Lainnya').replace(/^:\s*/, '').trim() || 'Lainnya';
    byJenis[j] = (byJenis[j] || 0) + 1;
  });

  return {
    kendala,
    fuKendala,
    byJenis,
    total:   kendala.length,
    totalFU: fuKendala.length,
  };
}

// ── Absensi ──────────────────────────────────────────────────
function getAbsen() {
  const ws   = ss().getSheetByName('Absen');
  const rows = ws.getDataRange().getValues();

  const today     = fmtDate(new Date());
  const todayList = [];
  const byDate    = {};

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const waktu = r[4] || r[0];
    const ds    = fmtDate(waktu);
    byDate[ds]  = (byDate[ds] || 0) + 1;
    if (ds === today) {
      todayList.push({
        userId:    r[1],
        nama:      r[2],
        username:  r[3],
        waktu,
        lat:       r[5],
        lng:       r[6],
        kelurahan: r[7],
        deal:      r[9],
      });
    }
  }

  const trend = [];
  for (let i = 13; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const ds = fmtDate(d);
    trend.push({ date: ds, count: byDate[ds] || 0 });
  }

  return { today: todayList, count: todayList.length, trend };
}
