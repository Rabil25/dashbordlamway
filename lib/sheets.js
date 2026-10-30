// ================================================================
//  Google Sheets helper – baca data via API v4 dengan API Key
//  Spreadsheet harus di-set public "Anyone with link can view"
// ================================================================

const { google } = require('googleapis');

const SHEET_ID  = '1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM';
const API_KEY   = process.env.SHEETS_API_KEY;
const TIMEZONE  = 'Asia/Jakarta';

// Index kolom dalam range CE2:DT (0-based dari CE=0)
const IDX_BLN_IO = 6;  // CK
const IDX_BLN_RE = 9;  // CN
const IDX_BLN_PS = 12; // CQ

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
      const blnIO = (row[IDX_BLN_IO] ?? '').toString().trim();
      const blnRE = (row[IDX_BLN_RE] ?? '').toString().trim();
      const blnPS = (row[IDX_BLN_PS] ?? '').toString().trim();
      if (blnIO !== filterBulan && blnRE !== filterBulan && blnPS !== filterBulan) continue;
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

// ── getListBulan ──────────────────────────────────────────────
async function getListBulan() {
  const key = 'list_bulan';
  const cached = cacheGet(key);
  if (cached) return { ...cached, _fromCache: true };

  // Baca kolom CK, CN, CQ (index 6,9,12 dari CE) → kolom CK=86, CN=89, CQ=92
  // Range CK2:CQ = kolom 87–93 dari A → tapi kita baca seluruh 'CK2:CQ'
  const values = await readRange('Bulan_Ini', 'CK2:CQ');

  const bulanSet = new Set();
  for (const row of values) {
    // Dalam range CK:CQ, index 0=CK(blnIO), 3=CN(blnRE), 6=CQ(blnPS)
    [0, 3, 6].forEach(i => {
      const v = (row[i] ?? '').toString().trim();
      if (v) bulanSet.add(v);
    });
  }

  const bulanList = [...bulanSet]
    .filter(b => /^\d+$/.test(b) && Number(b) >= 1 && Number(b) <= 12)
    .map(Number)
    .sort((a, b) => a - b);

  const result = { ok: true, bulan: bulanList, updatedAt: new Date().toISOString() };
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
async function getAllData(bulan) {
  const bulanNum  = parseInt(bulan) || 0;
  const bulanLalu = bulanNum > 1 ? (bulanNum - 1).toString() : '';

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

module.exports = { getRawSheetCached, getListBulan, getAllData, getTargetData };
