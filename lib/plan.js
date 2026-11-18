// ================================================================
//  Baca sheet DJP_Plan (Plan DJP) untuk menu Timeline Desa.
//  Read-only via Sheets API (API key) — sheet harus "Anyone with link can view".
//  Kolom: A TGL | B USERNAME_SF | C NAMA_SF | D ODP | E CLUSTER | F KAB | G KEC | H DESA
// ================================================================

const { google } = require('googleapis');

const PLAN_SHEET_ID = '1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ';
const TTL_MS = 60 * 1000;
let _rows = null, _ts = 0, _inflight = null;

// Terima 'yyyy-MM-dd', 'dd/MM/yyyy', 'd/M/yyyy', atau serial tanggal Sheets → 'yyyy-MM-dd'
function normTgl(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && v > 30000 && v < 80000) {
    const d = new Date(Math.round((v - 25569) * 86400e3));
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
  return '';
}

async function loadAll() {
  const sheets = google.sheets({ version: 'v4', auth: process.env.SHEETS_API_KEY });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: PLAN_SHEET_ID,
    range: 'DJP_Plan!A2:H',
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER',
  });
  const out = [];
  for (const r of res.data.values || []) {
    const tgl = normTgl(r[0]);
    const odp = (r[3] || '').toString().trim();
    if (!tgl || !odp) continue; // lewati baris kosong
    out.push({
      tgl,
      username_sf: (r[1] || '').toString().trim(),
      nama_sf:     (r[2] || '').toString().trim(),
      odp,
      kab:  (r[5] || '').toString().trim().toUpperCase(),
      kec:  (r[6] || '').toString().trim(),
      desa: (r[7] || '').toString().trim(),
    });
  }
  return out;
}

// periode 'YYYY-MM' → baris plan bulan itu
async function getPlanMonth(periode) {
  if (!_rows || Date.now() - _ts > TTL_MS) {
    if (!_inflight) {
      _inflight = loadAll()
        .then(rows => { _rows = rows; _ts = Date.now(); })
        .finally(() => { _inflight = null; });
    }
    await _inflight;
  }
  const p = String(periode || '');
  return { ok: true, periode: p, data: _rows.filter(r => r.tgl.slice(0, 7) === p) };
}

// ── Kunjungan nyata: tab DJP_Realisasi_Baru (1 baris = SF × desa × hari) ──
// A TGL | B KODE_SF | C NAMA_SF | D PLAN | E KAB | F KEC | G DESA | H JML_ODP | I DIKUNJUNGI | J TARGET_DJP | K SUMBER
// Hanya baris ber-DESA dan DIKUNJUNGI > 0 yang dihitung kunjungan.
let _real = null, _realTs = 0, _realInflight = null;
async function loadReal() {
  const sheets = google.sheets({ version: 'v4', auth: process.env.SHEETS_API_KEY });
  let vals = [];
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: PLAN_SHEET_ID, range: 'DJP_Realisasi_Baru!A2:K',
      valueRenderOption: 'UNFORMATTED_VALUE', dateTimeRenderOption: 'SERIAL_NUMBER',
    });
    vals = res.data.values || [];
  } catch (e) { vals = []; }   // tab belum dibuat
  const out = [];
  for (const r of vals) {
    const tgl = normTgl(r[0]), desa = (r[6] || '').toString().trim();
    if (!tgl || !desa || !(Number(r[8]) > 0)) continue;
    out.push({
      tgl, username_sf: (r[1] || '').toString().trim().toUpperCase(), nama_sf: (r[2] || '').toString().trim(),
      odp: '', kab: (r[4] || '').toString().trim().toUpperCase(), kec: (r[5] || '').toString().trim(), desa,
      jml_odp: Number(r[7]) || 0, dikunjungi: Number(r[8]) || 0, sumber: (r[10] || '').toString().trim(),
    });
  }
  return out;
}
async function getRealMonth(periode) {
  if (!_real || Date.now() - _realTs > TTL_MS) {
    if (!_realInflight) _realInflight = loadReal().then(rows => { _real = rows; _realTs = Date.now(); }).finally(() => { _realInflight = null; });
    await _realInflight;
  }
  const p = String(periode || '');
  return { ok: true, periode: p, data: _real.filter(r => r.tgl.slice(0, 7) === p) };
}

module.exports = { getPlanMonth, getRealMonth, normTgl };
