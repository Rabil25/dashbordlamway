// ================================================================
//  Users sheet helper – login, register, session via Google Sheets
//  Menulis ke sheet (login/register/logout) via Sheets API v4
//  Membutuhkan SERVICE ACCOUNT atau user auth untuk WRITE.
//  Untuk READ-ONLY auth verify: gunakan API Key cukup.
//  Untuk WRITE (login token, register, logout): gunakan Service Account.
//
//  CATATAN: Karena kita mau tetap gratis & simpel, strategi:
//  - READ (verify token, find user): pakai API Key (public viewer)
//  - WRITE (simpan token, register, dll): masih lewat GAS endpoint
//    karena GAS bisa write, Vercel + API Key tidak bisa write ke Sheets.
//
//  Jadi: data endpoint (all/bulan_ini/update_sf) → Vercel (cepat)
//        auth endpoint → masih GAS (lambat tapi jarang dipanggil)
// ================================================================

// Kolom Users sheet (0-based)
const U = {
  USERNAME: 0, HASH: 1, ROLE: 2, STATUS: 3,
  NAMA: 4, TGL_DAFTAR: 5, TOKEN: 6, EXPIRY: 7,
};

// Session cache in-memory (per Vercel instance)
const _sessCache = new Map();
const SESS_TTL_MS = 5 * 60 * 1000; // 5 menit

function sessCacheGet(token) {
  const entry = _sessCache.get(token);
  if (!entry) return null;
  if (Date.now() - entry.ts > SESS_TTL_MS) { _sessCache.delete(token); return null; }
  return entry.data;
}

function sessCacheSet(token, data) {
  _sessCache.set(token, { data, ts: Date.now() });
}

module.exports = { U, sessCacheGet, sessCacheSet };
