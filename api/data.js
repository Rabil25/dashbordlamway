// ================================================================
//  Vercel Serverless Function – Data endpoint
//  GET /api/data?token=lpgLampung2026&type=all&bulan=10
//
//  Menggantikan GAS endpoint untuk READ data (bukan auth).
//  Auth masih diteruskan ke GAS_AUTH_URL.
// ================================================================

const { getAllData, getRawSheetCached, getListBulan, getTargetData } = require('../lib/sheets');

const DATA_TOKEN = process.env.DATA_TOKEN || 'lpgLampung2026';

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(200).end();

  const { token, type = '', bulan = '' } = req.query;

  // Validasi token
  if (token !== DATA_TOKEN) {
    return res.status(401).json({ ok: false, error: 'Unauthorized' });
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
        data = await getRawSheetCached('Bulan_Ini', 'CE2:DT', bulan.toString());
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
      // sf_target: ringan – update_sf + target saja (bulan_ini kini dari GitHub static)
      case 'sf_target': {
        const [updateSf, target] = await Promise.all([
          getRawSheetCached('Update SF', 'A2:AG', ''),
          getTargetData(),
        ]);
        data = { ok: true, updateSf, target, updatedAt: new Date().toISOString() };
        break;
      }
      case 'ping':
        data = { ok: true, message: 'pong (Vercel)', updatedAt: new Date().toISOString() };
        break;
      default:
        data = { ok: false, error: 'Unknown type. Use: all | bulan_ini | update_sf | list_bulan | sf_target | wilayah | ping' };
    }

    return res.status(200).json(data);
  } catch (err) {
    console.error('Data handler error:', err.message);
    return res.status(500).json({ ok: false, error: err.message });
  }
};
