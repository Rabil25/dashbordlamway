// ================================================================
//  Verifikasi session token user (SEC-04)
//  Session dibuat oleh Apps Script saat login (tab Users, kolom G=token, H=expiry).
//  Vercel membaca tab Users lewat Sheets API (API key) dan mencocokkan token
//  secara lokal → tanpa panggilan ke Apps Script (yang bisa 2–30 detik).
//  Peta token di-cache 60 detik per instance; token yang belum dikenal
//  memicu satu kali baca ulang (untuk user yang baru saja login).
// ================================================================

const { google } = require('googleapis');

const SHEET_ID = '1IhU91CyYKQRDLysUYgpPn258ieoADixO30Cl5YS3YeM';
const U = { USERNAME: 0, ROLE: 2, STATUS: 3, NAMA: 4, TOKEN: 6, EXPIRY: 7 };
const TTL_MS = 60 * 1000;

let _map = null, _ts = 0, _inflight = null;

async function loadTokens() {
  const sheets = google.sheets({ version: 'v4', auth: process.env.SHEETS_API_KEY });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: 'Users!A2:H',
    valueRenderOption: 'UNFORMATTED_VALUE',
  });
  const map = new Map();
  for (const r of res.data.values || []) {
    const token = (r[U.TOKEN] || '').toString().trim();
    if (!token) continue;
    map.set(token, {
      username: (r[U.USERNAME] || '').toString().trim(),
      role: (r[U.ROLE] || '').toString().trim(),
      status: (r[U.STATUS] || '').toString().trim(),
      expiry: parseInt(r[U.EXPIRY] || 0, 10) || 0,
    });
  }
  _map = map; _ts = Date.now();
  return map;
}

async function getMap(force) {
  if (!force && _map && Date.now() - _ts < TTL_MS) return _map;
  if (!_inflight) _inflight = loadTokens().finally(() => { _inflight = null; });
  return _inflight;
}

// Mengembalikan { username, role } bila token valid, atau null.
async function verifySession(token) {
  if (!token || token.length < 20) return null;
  let map = await getMap(false);
  let u = map.get(token);
  if (!u && Date.now() - _ts > 5000) { map = await getMap(true); u = map.get(token); }
  if (!u) return null;
  if (u.status !== 'active') return null;
  if (!u.expiry || Date.now() > u.expiry) return null;
  return { username: u.username, role: u.role };
}

module.exports = { verifySession };
