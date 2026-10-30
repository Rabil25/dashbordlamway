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
};

// ── Index kolom dalam range CE2:DT (0-based dari CE=0) ──────────
// CE=0, CF=1, CG=2, CH=3, CI=4, CJ=5,
// CK=6(BLN_IO), CL=7(TGL_IO), CM=8, CN=9(BLN_RE), CO=10(TGL_RE),
// CP=11, CQ=12(BLN_PS), CR=13(TGL_PS), ..., DP=37(DEDUP), DQ=38(VAL_KAB)
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

// ── Entry Point GET ───────────────────────────────────────────
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  // Token auth — wajib
  if (params.token !== CONFIG.TOKEN) {
    return jsonOut({ ok: false, error: 'Unauthorized' }, 401);
  }

  try {
    const type  = (params.type  || '').toLowerCase();
    const bulan = (params.bulan || '').toString().trim();  // e.g. "10"
    let data;

    switch (type) {
      // ── Data Bulan Ini (kolom CE:DT), difilter per bulan ────
      case 'bulan_ini':
        data = getRawSheet('Bulan_Ini', 'CE2:DT', bulan);
        break;

      // ── Daftar bulan yang tersedia di sheet (ringan) ─────────
      case 'list_bulan':
        data = getListBulan();
        break;

      // ── Data Update SF ───────────────────────────────────────
      case 'update_sf':
        data = getRawSheet('Update SF', 'A2:AG', '');
        break;

      // ── Info wilayah & config ─────────────────────────────
      case 'wilayah':
        data = { ok: true, wilayah: WILAYAH };
        break;

      // ── Health check ──────────────────────────────────────
      case 'ping':
        data = { ok: true, message: 'pong', updatedAt: new Date().toISOString() };
        break;

      default:
        data = {
          ok: false,
          error: 'Unknown type. Use: bulan_ini | list_bulan | update_sf | wilayah | ping',
        };
    }

    return jsonOut(data);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ── Ambil daftar bulan unik dari kolom CK (BLN_IO) saja ──────
// Hanya baca 1 kolom → sangat cepat meski 7000+ baris
function getListBulan() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName('Bulan_Ini');
  if (!ws) return { ok: false, error: 'Sheet "Bulan_Ini" tidak ditemukan.' };

  // Baca CK:CQ (7 kolom: BLN_IO, TGL_IO, ?, BLN_RE, TGL_RE, ?, BLN_PS)
  const values = ws.getRange('CK2:CQ').getValues();
  const bulanSet = new Set();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[0] !== '' && row[0] !== null) bulanSet.add(row[0].toString().trim()); // BLN_IO
    if (row[3] !== '' && row[3] !== null) bulanSet.add(row[3].toString().trim()); // BLN_RE
    if (row[6] !== '' && row[6] !== null) bulanSet.add(row[6].toString().trim()); // BLN_PS
  }
  const bulanList = Array.from(bulanSet)
    .filter(function(b) { return /^\d+$/.test(b) && Number(b) >= 1 && Number(b) <= 12; })
    .map(Number)
    .sort(function(a, b) { return a - b; });

  return { ok: true, bulan: bulanList, updatedAt: new Date().toISOString() };
}

// ── Ambil data mentah dari sheet, filter per bulan ────────────
// filterBulan: string e.g. "8" — kosong = ambil semua
function getRawSheet(sheetName, range, filterBulan) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ws = ss.getSheetByName(sheetName);
  if (!ws) {
    return { ok: false, error: 'Sheet "' + sheetName + '" tidak ditemukan.' };
  }

  // Ambil data dari range yang ditentukan
  let values;
  try {
    values = ws.getRange(range).getValues();
  } catch (e) {
    return { ok: false, error: 'Gagal baca range ' + range + ': ' + e.message };
  }

  const isBulanIni = (sheetName === 'Bulan_Ini' && filterBulan !== '');

  // Bersihkan: konversi Date object → string "YYYY-MM-DD", filter bulan, buang baris kosong
  const cleaned = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    // Skip baris yang seluruhnya kosong
    const hasData = row.some(function(cell) {
      return cell !== '' && cell !== null && cell !== undefined;
    });
    if (!hasData) continue;

    // Filter per bulan (hanya untuk sheet Bulan_Ini)
    if (isBulanIni) {
      const blnIO = row[IDX_BLN_IO] !== undefined && row[IDX_BLN_IO] !== '' ? row[IDX_BLN_IO].toString().trim() : '';
      const blnRE = row[IDX_BLN_RE] !== undefined && row[IDX_BLN_RE] !== '' ? row[IDX_BLN_RE].toString().trim() : '';
      const blnPS = row[IDX_BLN_PS] !== undefined && row[IDX_BLN_PS] !== '' ? row[IDX_BLN_PS].toString().trim() : '';
      // Tampilkan baris yang relevan untuk bulan ini (IO, RE, atau PS di bulan yang dipilih)
      if (blnIO !== filterBulan && blnRE !== filterBulan && blnPS !== filterBulan) continue;
    }

    // Konversi setiap cell
    const cleanRow = row.map(function(cell) {
      if (cell instanceof Date) {
        // Pakai timezone Jakarta agar tanggal tidak geser
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
    bulan:     filterBulan || 'all',
    rows:      cleaned.length,
    data:      cleaned,
    updatedAt: new Date().toISOString(),
  };
}

// ── Helper: output JSON dengan CORS header ────────────────────
function jsonOut(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
