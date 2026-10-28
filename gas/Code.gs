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
    const type = (params.type || '').toLowerCase();
    let data;

    switch (type) {
      // ── Data Bulan Ini (kolom CE:DT) ────────────────────────
      case 'bulan_ini':
        data = getRawSheet('Bulan_Ini', 'CE2:DT');
        break;

      // ── Data Update SF ───────────────────────────────────────
      case 'update_sf':
        data = getRawSheet('Update SF', 'A2:AG');
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
          error: 'Unknown type. Use: bulan_ini | update_sf | wilayah | ping',
        };
    }

    return jsonOut(data);
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

// ── Ambil data mentah dari sheet, return sebagai array of arrays ──
// maxRows: batas baris (opsional, 0 = ambil semua)
function getRawSheet(sheetName, range, maxRows) {
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

  // Bersihkan: konversi Date object → string "YYYY-MM-DD", buang baris kosong di akhir
  const cleaned = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    // Skip baris yang seluruhnya kosong
    const hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (!hasData) continue;

    // Konversi setiap cell
    const cleanRow = row.map(cell => {
      if (cell instanceof Date) {
        // Pakai timezone Jakarta agar tanggal tidak geser
        return Utilities.formatDate(cell, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      }
      return cell;
    });

    cleaned.push(cleanRow);

    // Batas baris opsional
    if (maxRows && maxRows > 0 && cleaned.length >= maxRows) break;
  }

  return {
    ok:        true,
    sheet:     sheetName,
    range:     range,
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
