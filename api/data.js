// ================================================================
//  Vercel Serverless Function – Data endpoint
//  GET /api/data?type=all&bulan=10
//  Header: Authorization: Bearer <session token dari login>
//
//  Menggantikan GAS endpoint untuk READ data (bukan auth).
//  Auth masih diteruskan ke GAS_AUTH_URL.
// ================================================================

const { getAllData, getRawSheetCached, getListBulan, normPeriod } = require('../lib/sheets');
const { verifySession } = require('../lib/session');

// Token bersama lama HANYA berlaku bila env DATA_TOKEN diisi di Vercel
// (untuk integrasi non-browser, mis. bot). Tidak ada lagi nilai default di kode.
const DATA_TOKEN = process.env.DATA_TOKEN || '';

module.exports = async function handler(req, res) {
  // Hanya dipanggil dari situs ini sendiri (same-origin) → tidak perlu CORS terbuka
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization');
  if (req.method === 'OPTIONS') return res.status(204).end();

  const { token, type = '', bulan = '' } = req.query;

  // SEC-04: wajib session login yang valid (atau token integrasi bila diset di env)
  const auth = (req.headers.authorization || '').replace(/^Bearer\s+/i, '').trim();
  let allowed = false;
  if (type.toLowerCase() === 'ping') allowed = true;
  else if (DATA_TOKEN && token === DATA_TOKEN) allowed = true;
  else if (auth) {
    try { allowed = !!(await verifySession(auth)); }
    catch (e) { console.error('verifySession:', e.message); return res.status(503).json({ ok: false, error: 'Verifikasi sesi gagal, coba lagi.' }); }
  }
  if (!allowed) {
    res.setHeader('Cache-Control', 'no-store');
    return res.status(401).json({ ok: false, error: 'Sesi tidak valid. Silakan login ulang.' });
  }

  if (!process.env.SHEETS_API_KEY) {
    return res.status(500).json({ ok: false, error: 'SHEETS_API_KEY tidak dikonfigurasi.' });
  }

  try {
    const t = type.toLowerCase();
    let data;

    switch (t) {
      case 'all':
        data = await getAllData(bulan.toString());
        break;
      case 'bulan_ini':
        data = await getRawSheetCached('Bulan_Ini', 'CE2:DT', normPeriod(bulan));
        break;
      case 'update_sf':
        data = await getRawSheetCached('Update SF', 'A2:AG', '');
        break;
      case 'list_bulan':
        data = await getListBulan();
        break;
      case 'wilayah':
        data = {
          ok: true,
          wilayah: {
            'Lampung Barat': { kabupaten: ['LAMPUNG BARAT', 'TANGGAMUS', 'PESISIR BARAT'] },
            'Way Kanan':     { kabupaten: ['LAMPUNG UTARA', 'WAY KANAN'] },
          },
        };
        break;
      case 'ping':
        data = { ok: true, message: 'pong (Vercel)', updatedAt: new Date().toISOString() };
        break;
      default:
        data = { ok: false, error: 'Unknown type. Use: all | bulan_ini | update_sf | list_bulan | wilayah | ping' };
    }

    // Data hanya untuk user login → TIDAK boleh di-cache CDN publik.
    // (Vercel juga tidak meng-cache request ber-header Authorization.)
    // Cache memori server 10 menit (lib/sheets.js) tetap berlaku.
    res.setHeader('Cache-Control', 'private, no-store');
    return res.status(200).json(data);
  } catch (err) {
    console.error('Data handler error:', err.message);
    return res.status(500).json({ ok: false, error: 'Gagal membaca data sheet.' });
  }
};
