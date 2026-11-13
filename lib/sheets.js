// ================================================================
//  Google Sheets helper – baca data via API v4 dengan API Key
//  Spreadsheet harus di-set public "Anyone with link can view"
// ================================================================

const { google } = require('googleapis');

const SHEET_ID  = '1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM';
const API_KEY   = process.env.SHEETS_API_KEY;
const TIMEZONE  = 'Asia/Jakarta';

// Index kolom dalam range CE2:DT (0-based dari CE=0)
// Periode (tahun-bulan) diambil dari kolom Tanggal yyyymmdd, BUKAN dari kolom Bulan
// (CK/CN/CQ hanya berisi angka bulan tanpa tahun → tercampur antar tahun).
const IDX_DATE_IO = 29; // DH  yyyymmdd
const IDX_DATE_RE = 30; // DI
const IDX_DATE_PS = 31; // DJ

// 'YYYYMMDD' → 'YYYY-MM' ('' bila tidak valid)
function ymdToPeriod(v) {
  const s = String(v ?? '').trim();
  // Abaikan tanggal kosong ala Excel (18991230) dan nilai di luar 2000–2099
  return /^20\d{2}(0[1-9]|1[0-2])\d{2}$/.test(s) ? s.slice(0, 4) + '-' + s.slice(4, 6) : '';
}
// Terima 'YYYY-MM' atau angka bulan lama ('11') → 'YYYY-MM' (tahun berjalan WIB)
function normPeriod(b) {
  const s = String(b || '').trim();
  if (/^\d{4}-\d{2}$/.test(s)) return s;
  if (/^\d{1,2}$/.test(s)) {
    const y = new Date(Date.now() + 7 * 3600e3).getUTCFullYear();
    return y + '-' + s.padStart(2, '0');
  }
  return '';
}
function prevPeriod(p) {
  const y = parseInt(p.slice(0, 4)), m = parseInt(p.slice(5, 7));
  return m === 1 ? (y - 1) + '-12' : y + '-' + String(m - 1).padStart(2, '0');
}

const sheets = google.sheets({ version: 'v4', auth: API_KEY });

// ── In-memory cache sederhana (per instance Vercel, ~15 menit) ──
const _cache = new Map();
const CACHE_TTL_MS      = 10 * 60 * 1000;  // 10 menit data
const CACHE_TTL_BULAN_MS = 60 * 60 * 1000; // 1 jam list bulan

function cacheGet(key) {
  const entry = _cache.get(key);
  if (!entry) return null;
  if (Date.now() - entry.ts > entry.ttl) { _cache.delete(key); return null; }
  return entry.data;
}

function cacheSet(key, data, ttlMs) {
  _cache.set(key, { data, ts: Date.now(), ttl: ttlMs });
}

// ── Format tanggal ke yyyy-MM-dd (WIB) ──────────────────────────
function fmtDate(val) {
  if (!val) return val;
  // Sheets API mengembalikan tanggal sebagai string "2026-10-01" atau serial
  // Bila sudah string dd/mm/yyyy atau yyyy-mm-dd → kembalikan apa adanya
  if (typeof val === 'string') return val;
  // Bila Date object
  if (val instanceof Date) {
    const d = new Date(val.toLocaleString('en-US', { timeZone: TIMEZONE }));
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${dd}`;
  }
  return val;
}

// ── Baca satu range dari Google Sheets ──────────────────────────
async function readRange(sheetName, range) {
  const fullRange = `${sheetName}!${range}`;
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: fullRange,
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'FORMATTED_STRING',
  });
  return res.data.values || [];
}

// ── getRawSheet: baca sheet + filter bulan ────────────────────
async function getRawSheet(sheetName, range, filterBulan) {
  const values = await readRange(sheetName, range);

  const isBulanIni = (sheetName === 'Bulan_Ini' && filterBulan !== '');
  const cleaned = [];

  for (const row of values) {
    // Pad row ke panjang yg cukup
    while (row.length < 42) row.push('');

    const hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
    if (!hasData) continue;

    if (isBulanIni) {
      const pIO = ymdToPeriod(row[IDX_DATE_IO]);
      const pRE = ymdToPeriod(row[IDX_DATE_RE]);
      const pPS = ymdToPeriod(row[IDX_DATE_PS]);
      if (pIO !== filterBulan && pRE !== filterBulan && pPS !== filterBulan) continue;
    }

    const cleanRow = row.map(cell => fmtDate(cell));
    cleaned.push(cleanRow);
  }

  return {
    ok: true,
    sheet: sheetName,
    range,
    bulan: filterBulan || 'all',
    rows: cleaned.length,
    data: cleaned,
    updatedAt: new Date().toISOString(),
  };
}

// ── getRawSheetCached ─────────────────────────────────────────
async function getRawSheetCached(sheetName, range, filterBulan) {
  const key = `sheet_${sheetName.replace(/\s/g,'_')}_${filterBulan||'all'}`;
  const cached = cacheGet(key);
  if (cached) return { ...cached, _fromCache: true };

  const result = await getRawSheet(sheetName, range, filterBulan);
  if (result.ok) cacheSet(key, result, CACHE_TTL_MS);
  return result;
}

// ── getListBulan: daftar periode 'YYYY-MM' yang punya data ──────
async function getListBulan() {
  const key = 'list_bulan_v2';
  const cached = cacheGet(key);
  if (cached) return { ...cached, _fromCache: true };

  // DH:DJ = Tanggal IO/RE/PS (yyyymmdd)
  const values = await readRange('Bulan_Ini', 'DH2:DJ');
  const set = new Set();
  for (const row of values) for (const v of row) { const p = ymdToPeriod(v); if (p) set.add(p); }

  const result = { ok: true, bulan: [...set].sort(), updatedAt: new Date().toISOString() };
  cacheSet(key, result, CACHE_TTL_BULAN_MS);
  return result;
}

// ── getTargetData: baca sheet Target (Kabupaten | PS) ─────────
async function getTargetData() {
  const cacheKey = 'target_all';
  const cached = _cache.get(cacheKey);
  if (cached && Date.now() - cached.ts < 600000) return cached.data;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Target!A2:B',
  });
  const rows = res.data.values || [];
  const targets = {};
  for (const row of rows) {
    const kab = (row[0] || '').toString().trim().toUpperCase();
    const ps  = parseFloat(row[1]) || 0;
    if (kab) targets[kab] = { ps, re: Math.round(ps * 1.2), io: Math.round(ps * 1.44) };
  }
  const data = { ok: true, targets, updatedAt: new Date().toISOString() };
  _cache.set(cacheKey, { data, ts: Date.now() });
  return data;
}

// ── getAllData: combined endpoint ─────────────────────────────
async function getAllData(bulanParam) {
  const bulan     = normPeriod(bulanParam);
  const bulanLalu = bulan ? prevPeriod(bulan) : '';   // Jan 2027 → Des 2026

  const [bulanIni, bulanPrev, updateSf, listBulan, target] = await Promise.all([
    getRawSheetCached('Bulan_Ini', 'CE2:DT', bulan),
    bulanLalu ? getRawSheetCached('Bulan_Ini', 'CE2:DT', bulanLalu) : Promise.resolve({ ok:true, data:[], rows:0 }),
    getRawSheetCached('Update SF', 'A2:AG', ''),
    getListBulan(),
    getTargetData(),
  ]);

  return {
    ok: true,
    bulan,
    bulanLalu,
    bulanIni,
    bulanPrev,
    updateSf,
    listBulan,
    target,
    updatedAt: new Date().toISOString(),
  };
}

module.exports = { getRawSheetCached, getListBulan, getAllData, getTargetData, normPeriod };
