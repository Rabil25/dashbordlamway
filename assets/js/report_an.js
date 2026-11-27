// ================================================================
//  PPT Analisa (format standar semua PPT di web, Nov 2026)
//  7 slide compact: Ringkasan · Pencapaian & Tren · Funnel & Mix ·
//  DJP 3 bulan → PS (wilayah) · SF · Prioritas & Plan · Kesimpulan & Aksi.
//  Setiap slide ditutup analisa silogisme (premis mayor → minor → kesimpulan).
//
//  Aturan analisa DJP (disepakati dengan PJ):
//   - PS bulan M dibandingkan DJP bulan M-1, M-2, M-3 (keputusan berlangganan
//     butuh waktu; hasil kunjungan tampak di bulan ke-4). Jendela TIDAK digeser:
//     bulan yang datanya kosong ditandai "belum tersedia".
//   - Satuan kunjungan = 1 SF × 1 desa × 1 hari. ODP hanya alat check-in;
//     yang dikejar kehadiran SF di wilayah desa.
//
//  Mode (menu PPT):
//   mingguan : LB vs WK, PS 7 hari terakhir vs 7 hari sebelumnya
//   bulanan  : LB vs WK, PS MTD vs periode sama bulan lalu
//   cluster  : kabupaten di cluster aktif, MTD
//   all      : semua kabupaten kedua cluster, MTD
//  Bergantung pada config.js (C, CLUSTER, period helpers), report.js
//  (loadPptLib, sum) dan report_cmp.js (ymdNum, addDays, dateOfYmd, fmtTglPendek).
// ================================================================

const AN = {
  F: 'Poppins', RED: 'C00000', DARK: '1F1F1F', GREY: '6B6B6B', LG: 'BFBFBF', BG: 'F7F7F7',
  GRN: '00B050', AMB: 'D48A00', PINK: 'FFF7F7',
  ENT: ['C00000', '7F7F7F', 'D48A00', '2F5597', '548235'],
  KAT: ['prioritas', 'strategis', 'cadangan', 'skip'],
  KAT_L: ['Prioritas', 'Strategis', 'Cadangan', 'Skip'],
  KAT_C: ['C00000', 'D48A00', '8C8C8C', 'D9D9D9'],
  PROD: ['Tsel One', 'IndiHome Reguler', 'EZNet', 'HWA', 'Lainnya'],
  PROD_C: ['C00000', '404040', 'D48A00', 'BFBFBF', 'E7E6E6'],
  TL: { 'Lampung Barat': 'TL Andres', 'Way Kanan': 'TL Agus' },
};
const AN_BLN = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const anN = s => (s || '').toString().trim().toUpperCase();
const anK = (kec, desa) => anN(kec) + '|' + anN(desa);
const anTitle = s => String(s || '').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
const anD = (n, d = 1) => (n == null || !isFinite(n)) ? '–' : (Math.round(n * 10 ** d) / 10 ** d).toLocaleString('id-ID', { minimumFractionDigits: d, maximumFractionDigits: d });
const anI = n => (n == null || !isFinite(n)) ? '–' : Math.round(n).toLocaleString('id-ID');
const anP = (a, b) => b > 0 ? Math.round(a / b * 100) : null;
const anPs = p => p == null ? '–' : p + '%';
const anSg = p => p == null ? '–' : (p > 0 ? '+' : '') + p + '%';
const anCl = kab => Object.keys(CLUSTER).find(c => CLUSTER[c].kab.map(anN).includes(anN(kab))) || '';
const anShort = c => c === 'Lampung Barat' ? 'LB' : c === 'Way Kanan' ? 'WK' : anTitle(c);
const anSfName = n => String(n || '').replace(/\[.*?\]\s*/g, '').trim();
function anProd(v) {
  const s = (v || '').toString().toLowerCase();
  if (/tsel|one/.test(s)) return 'Tsel One';
  if (/indihome/.test(s)) return 'IndiHome Reguler';
  if (/eznet/.test(s)) return 'EZNet';
  if (/hwa/.test(s)) return 'HWA';
  return 'Lainnya';
}
const anMonthShort = p => (BULAN_NAMA[periodMonth(p)] || '?').slice(0, 3);
const anJoin = arr => arr.join(' / ');

// ================================================================
//  1. KUMPULKAN DATA
// ================================================================
async function anCollect(app, mode) {
  try { await app.loadArpu(); } catch (e) {}
  const bulan = app.cfg.bulan, days = daysInPeriod(bulan);
  const refDay = Math.max(1, Math.min(app.hariBerjalan || days, days));
  const end = +(bulan.replace('-', '') + String(refDay).padStart(2, '0'));
  const prev = prevPeriod(bulan);
  let W, W0, label, label0;
  if (mode === 'mingguan') {
    W = [addDays(end, -6), end]; W0 = [addDays(end, -13), addDays(end, -7)];
    label = `${fmtTglPendek(W[0])} – ${fmtTglPendek(W[1])}`; label0 = `${fmtTglPendek(W0[0])} – ${fmtTglPendek(W0[1])}`;
  } else {
    const dPrev = daysInPeriod(prev);
    W = [+(bulan.replace('-', '') + '01'), end];
    W0 = [+(prev.replace('-', '') + '01'), +(prev.replace('-', '') + String(Math.min(refDay, dPrev)).padStart(2, '0'))];
    label = refDay < days ? `1–${refDay} ${periodLabel(bulan)}` : periodLabel(bulan);
    label0 = `1–${Math.min(refDay, dPrev)} ${periodLabel(prev)}`;
  }
  const trendM = [bulan]; for (let i = 0; i < 4; i++) trendM.unshift(prevPeriod(trendM[0]));   // M-4..M
  const djpM = [prevPeriod(prevPeriod(prev)), prevPeriod(prev), prev];                      // M-3..M-1

  const auth = { headers: { Authorization: 'Bearer ' + app.sessionToken } };
  const getJSON = async url => { try { const r = await fetch(url, auth); const d = await r.json(); return d.ok ? d : null; } catch (e) { return null; } };
  const needRows = trendM.slice(0, 3);   // M-4..M-2 (M & M-1 sudah di allRows)
  const res = await Promise.all([
    ...needRows.map(m => getJSON(`/api/data?type=bulan_ini&bulan=${encodeURIComponent(m)}`)),
    ...djpM.map(m => getJSON(`/api/data?type=djp_real&bulan=${encodeURIComponent(m)}`)),
    getJSON(`/api/data?type=djp_plan&bulan=${encodeURIComponent(bulan)}`),
    getJSON(`/api/data?type=djp_real&bulan=${encodeURIComponent(bulan)}`),
  ]);
  const extraRows = res.slice(0, 3).flatMap(d => (d && d.data) || []);
  const planByM = djpM.map((m, i) => ((res[3 + i] && res[3 + i].data) || []));
  const planCur = (res[6] && res[6].data) || [];
  const realCur = (res[7] && res[7].data) || [];

  if (!app.prioritasOdpRaw.length) { try { await app.loadPrioritas(); } catch (e) {} }
  const stack = document.body._x_dataStack;
  const base = stack && stack[0] && window.Alpine && Alpine.raw ? Alpine.raw(stack[0]) : null;
  if (!base || !('cluster' in base)) throw new Error('Data dashboard belum siap, coba lagi');

  // baris unik semua bulan
  const seen = new Set(), all = [];
  [...app.allRows, ...extraRows].forEach(r => {
    if ((r[C.DEDUP] || '').trim() !== 'Unik') return;
    const id = (r[C.ID_ORDER] || '').toString().trim();
    if (id && seen.has(id)) return; if (id) seen.add(id);
    all.push(r);
  });
  const hasDesa = all.some(r => (r[C.DESA] || '').toString().trim());

  // klasifikasi desa kedua cluster (shadow view, dashboard tidak di-render ulang)
  const desaAll = [];
  if (app.prioritasOdpRaw.length) Object.keys(CLUSTER).forEach(name => {
    const v = Object.create(base); Object.defineProperty(v, 'cluster', { value: name, writable: true });
    desaAll.push(...v.klasifikasiDesa('', ''));
  });
  const katMap = {}; desaAll.forEach(d => { katMap[anK(d.kec, d.desa)] = d; });

  // kunjungan = SF × desa × hari
  const toVisits = rows => { const m = new Map(); rows.forEach(p => {
    const k = anN(p.username_sf) + '|' + anK(p.kec, p.desa) + '|' + p.tgl;
    if (!m.has(k)) m.set(k, { sf: anN(p.username_sf), nama: p.nama_sf, kab: anN(p.kab), dk: anK(p.kec, p.desa), tgl: p.tgl, m: String(p.tgl).slice(0, 7) });
  }); return [...m.values()]; };
  const visByM = planByM.map(toVisits);
  const availM = djpM.filter((m, i) => visByM[i].length > 0);
  const visits = visByM.flat();
  const visCur = toVisits(planCur);

  // entitas pembanding
  let ents;
  if (mode === 'mingguan' || mode === 'bulanan') ents = Object.keys(CLUSTER).map(c => ({ name: c, short: anShort(c), kabs: CLUSTER[c].kab.map(anN), cl: c }));
  else {
    const cls = mode === 'cluster' ? [app.cluster] : Object.keys(CLUSTER);
    ents = cls.flatMap(c => CLUSTER[c].kab.map(k => ({ name: anTitle(k), short: anTitle(k), kabs: [anN(k)], cl: c })));
  }

  const D = { mode, bulan, prev, days, refDay, sisa: Math.max(days - refDay, 0), W, W0, label, label0, trendM, djpM, availM, hasDesa,
    wDays: Math.round((dateOfYmd(W[1]) - dateOfYmd(W[0])) / 864e5) + 1, ents: [], katMap, allRows: all };
  ents.forEach((e, i) => D.ents.push(anEntity(app, D, e, i, all, visits, visCur, desaAll, katMap)));
  anDisiplin(D, planCur, realCur, katMap);
  return D;
}

// ── Disiplin plan vs realisasi DJP (periode laporan, s.d. tanggal realisasi terakhir) ──
// SF-hari: SF punya plan di hari itu & tercatat visit · desa-hari per kategori: desa plan didatangi
function anDisiplin(D, planCur, realCur, katMap) {
  const ymd = n => { const t = String(n); return t.slice(0, 4) + '-' + t.slice(4, 6) + '-' + t.slice(6, 8); };
  const last = realCur.reduce((m, r) => r.tgl > m ? r.tgl : m, '');
  const from = ymd(D.W[0]); let to = ymd(D.W[1]); if (last && last < to) to = last;
  D.discRange = last ? [from, to] : null;
  const KI = { prioritas: 0, strategis: 1, cadangan: 2, skip: 3 };
  const inW = t => t >= from && t <= to;
  D.ents.forEach(e => {
    const inK = kab => e.kabs.includes(anN(kab));
    const P = planCur.filter(p => p.desa && inW(p.tgl) && inK(p.kab)), R = realCur.filter(r => inW(r.tgl) && inK(r.kab));
    const rSF = new Set(R.map(r => anN(r.username_sf) + '|' + r.tgl)), rD = new Set(R.filter(r => r.desa).map(r => anK(r.kec, r.desa) + '|' + r.tgl));
    const pSF = new Set(P.map(p => anN(p.username_sf) + '|' + p.tgl));
    const pD = new Map(); P.forEach(p => pD.set(anK(p.kec, p.desa) + '|' + p.tgl, (katMap[anK(p.kec, p.desa)] || {}).kat || 'skip'));
    const kat = [0, 1, 2, 3].map(() => ({ plan: 0, hit: 0 }));
    pD.forEach((kt, key) => { const c = kat[KI[kt] ?? 3]; c.plan++; if (rD.has(key)) c.hit++; });
    const luar = [0, 0, 0, 0]; rD.forEach(key => { if (!pD.has(key)) { const kt = (katMap[key.split('|').slice(0, 2).join('|')] || {}).kat || 'skip'; luar[KI[kt] ?? 3]++; } });
    const sfHit = [...pSF].filter(x => rSF.has(x)).length;
    e.disc = { sfPlan: pSF.size, sfHit, sfPct: anP(sfHit, pSF.size), kat, luar,
      desaPlan: sum(kat.map(c => c.plan)), desaHit: sum(kat.map(c => c.hit)) };
  });
}

// ── metrik satu entitas ──
function anEntity(app, D, e, idx, all, visits, visCur, desaAll, katMap) {
  const inK = k => e.kabs.includes(anN(k));
  const rows = all.filter(r => inK(r[C.VAL_KAB]));
  const inW = (v, W) => { const n = ymdNum(v); return n >= W[0] && n <= W[1]; };
  const isPS = r => !!ymdNum(r[C.DATE_PS]) && !!(r[C.VAL_PS] || '').toString().trim();
  const val = (r, c) => !!(r[c] || '').toString().trim();
  const o = { ...e, color: AN.ENT[idx % AN.ENT.length], tl: AN.TL[e.cl] || 'TL' };

  const met = W => ({ io: rows.filter(r => val(r, C.VAL_IO) && inW(r[C.DATE_IO], W)).length,
    re: rows.filter(r => val(r, C.VAL_RE) && inW(r[C.DATE_RE], W)).length,
    ps: rows.filter(r => isPS(r) && inW(r[C.DATE_PS], W)).length });
  Object.assign(o, met(D.W)); o.ps0 = met(D.W0).ps;
  o.tgtM = sum(e.kabs.map(k => (app.cfg.targets[k] || {}).ps || 0));
  o.psre = o.re ? o.ps / o.re * 100 : null;
  o.monthly = D.trendM.map(m => { const p = m.replace('-', ''); return rows.filter(r => isPS(r) && String(ymdNum(r[C.DATE_PS])).slice(0, 6) === p).length; });
  if (D.mode === 'mingguan') {
    o.tgtW = o.tgtM * D.wDays / D.days; o.ach = o.tgtW ? o.ps / o.tgtW * 100 : null;
    o.drr = o.ps / D.wDays; o.mtd = o.monthly[4]; o.drrM = o.mtd / D.refDay;
  } else {
    o.drr = o.ps / D.refDay; o.mtd = o.ps; o.drrM = o.drr;
  }
  o.ol = Math.round(o.drrM * D.days); o.olAch = o.tgtM ? o.ol / o.tgtM * 100 : null;
  // IO (tuas yang dikontrol sales): target, MTD, outlook, IO di lokasi potensial (Prioritas+Strategis)
  o.ioTgt = sum(e.kabs.map(k => (app.cfg.targets[k] || {}).io || 0));
  const bp = D.bulan.replace('-', ''), ioM = rows.filter(r => val(r, C.VAL_IO) && String(ymdNum(r[C.DATE_IO])).slice(0, 6) === bp && ymdNum(r[C.DATE_IO]) <= +(bp + String(D.refDay).padStart(2, '0')));
  o.ioMtd = ioM.length; o.ioOl = Math.round(o.ioMtd / D.refDay * D.days); o.ioOlAch = o.ioTgt ? o.ioOl / o.ioTgt * 100 : null;
  o.ioDrr = o.ioMtd / D.refDay; o.ioNeed = D.sisa > 0 ? Math.max(0, (o.ioTgt - o.ioMtd) / D.sisa) : 0;
  const ioWin = rows.filter(r => val(r, C.VAL_IO) && inW(r[C.DATE_IO], D.W));
  o.ioW = ioWin.length; o.ioPotW = ioWin.filter(r => ['prioritas', 'strategis'].includes((katMap[anK(r[C.KEC], r[C.DESA])] || {}).kat)).length;
  o.ioPotShare = o.ioW ? o.ioPotW / o.ioW * 100 : null;
  o.need = D.sisa > 0 ? Math.max(0, (o.tgtM - o.mtd) / D.sisa) : 0;
  o.tgtToDate = o.tgtM * D.refDay / D.days; o.achToDate = o.tgtToDate ? o.mtd / o.tgtToDate * 100 : null;
  // harian
  o.daily = []; for (let n = D.W[0]; n <= D.W[1]; n = addDays(n, 1)) o.daily.push(rows.filter(r => isPS(r) && ymdNum(r[C.DATE_PS]) === n).length);
  // status dari IO jendela
  const coh = rows.filter(r => val(r, C.VAL_IO) && inW(r[C.DATE_IO], D.W));
  const cps = coh.filter(isPS).length, cc = coh.filter(r => !isPS(r) && /CANCEL|BATAL/.test(anN(r[C.STATUS_FNL]))).length;
  o.coh = [cps, cc, coh.length - cps - cc]; o.cancelPer100 = coh.length ? cc / coh.length * 100 : null; o.pendPct = coh.length ? (coh.length - cps - cc) / coh.length * 100 : null;
  // mix
  const psW = rows.filter(r => isPS(r) && inW(r[C.DATE_PS], D.W));
  o.prod = AN.PROD.map(p => psW.filter(r => anProd(r[C.PRODUK_KAT]) === p).length);
  o.agency = anP(psW.filter(r => /AGENCY/.test(anN(r[C.CHANNEL2]))).length, psW.length);
  o.sp20 = anP(psW.filter(r => String(r[C.SPEED] || '').trim() === '20').length, psW.length);
  o.ihReg = anP(o.prod[1], psW.length); o.tselOne = anP(o.prod[0], psW.length);

  // ARPU (rata-rata harga paket semua order) — tertimbang jumlah order per desa
  { let aS = 0, aN = 0; desaAll.filter(d => inK(d.kab)).forEach(d => { const a = app.arpuOf ? app.arpuOf(d.kec, d.desa) : null; if (a && a.arpu) { aS += a.arpu * a.n; aN += a.n; } });
    o.arpu = aN ? aS / aN : null; }
  // Kendala order (CZ kategori, CY alasan) dari IO periode ini: hasil akhir batal / PS / masih menunggu
  { const KC = ['Kendala Teknik', 'Kendala Pelanggan', 'Kendala Sistem'];
    const kk = KC.map(() => [0, 0, 0]), rs = {};
    coh.forEach(r => { const cat = String(r[C.KENDALA_CAT] || '').trim(); const ci = KC.findIndex(c => c.toUpperCase() === cat.toUpperCase()); if (ci < 0) return;
      const out = isPS(r) ? 1 : /CANCEL|BATAL/.test(anN(r[C.STATUS_FNL])) ? 0 : 2; kk[ci][out]++;
      if (out !== 1) { const dtl = String(r[C.KENDALA_DTL] || '').trim(); const al = (dtl.includes('|') ? dtl.split('|').slice(1).join('|') : '').trim() || (ci === 2 ? 'Error sistem' : 'Tanpa keterangan');
        const key = KC[ci].replace('Kendala ', '') + ': ' + anTitle(al.toLowerCase()).replace(/\bOdp\b/g, 'ODP').replace(/\bRna\b/g, 'RNA (tidak bisa dihubungi)'); rs[key] = (rs[key] || 0) + 1; } });
    o.kend = kk; o.kendTot = sum(kk.map(x => x[0] + x[2])); o.kendTop = Object.entries(rs).sort((a, b) => b[1] - a[1]).slice(0, 6); }

  // ── DJP 3 bulan (kunjungan SF×desa×hari) ──
  const vis = visits.filter(v => inK(v.kab));
  const nM = Math.max(1, D.availM.length);
  o.visByM = D.djpM.map(m => vis.filter(v => v.m === m).length);
  o.vis = vis.length;
  const sfV = {}, desaV = {}, pairV = {}, sfDesa = {}, sfNama = {};
  vis.forEach(v => { sfV[v.sf] = (sfV[v.sf] || 0) + 1; desaV[v.dk] = (desaV[v.dk] || 0) + 1; const pk = v.sf + '|' + v.dk; pairV[pk] = (pairV[pk] || 0) + 1;
    (sfDesa[v.sf] = sfDesa[v.sf] || new Set()).add(v.dk); sfNama[v.sf] = v.nama; });
  o.sfDJP = Object.keys(sfV).length; o.desaTouched = Object.keys(desaV).length;
  o.visPerSfM = o.sfDJP ? o.vis / o.sfDJP / nM : 0;
  o.desaPerSf = o.sfDJP ? sum(Object.values(sfDesa).map(s => s.size)) / o.sfDJP : 0;
  o.visPerPair = Object.keys(pairV).length ? o.vis / Object.keys(pairV).length : 0;
  o.katVis = AN.KAT.map(k => vis.filter(v => (katMap[v.dk] || {}).kat === k).length);
  o.prioVisShare = anP(o.katVis[0], o.vis); o.goodVisShare = anP(o.katVis[0] + o.katVis[1], o.vis);
  // PS jendela per desa
  const psD = {}; psW.forEach(r => { const k = anK(r[C.KEC], r[C.DESA]); psD[k] = (psD[k] || 0) + 1; });
  o.psFromVis = psW.filter(r => desaV[anK(r[C.KEC], r[C.DESA])]).length;
  o.psSamePair = psW.filter(r => pairV[anN(r[C.CODE_SF]) + '|' + anK(r[C.KEC], r[C.DESA])]).length;
  o.psRutin = psW.filter(r => (pairV[anN(r[C.CODE_SF]) + '|' + anK(r[C.KEC], r[C.DESA])] || 0) >= 3).length;
  // desa per intensitas (kunjungan/bulan rata-rata)
  const clDesa = desaAll.filter(d => inK(d.kab)); o.desaTotal = clDesa.length;
  const BK = [['0', 0, 0], ['≤2', 0.01, 2], ['2–5', 2.01, 5], ['5–12', 5.01, 12], ['>12', 12.01, 1e9]];
  o.bucketL = BK.map(b => b[0]);
  o.bucket = BK.map(([l, lo, hi]) => { const ks = clDesa.map(d => anK(d.kec, d.desa)).filter(k => { const v = (desaV[k] || 0) / nM; return l === '0' ? v === 0 : v >= lo && v <= hi; });
    return { l, n: ks.length, ps: sum(ks.map(k => psD[k] || 0)), avg: ks.length ? sum(ks.map(k => psD[k] || 0)) / ks.length : null }; });
  // SF per intensitas (kunjungan/bulan)
  const sfPS = {}, sfN = {}; psW.forEach(r => { const s = anN(r[C.CODE_SF]); if (!s) return; sfPS[s] = (sfPS[s] || 0) + 1; sfN[s] = r[C.NAMA_SF]; });
  const codes = new Set([...Object.keys(sfV), ...Object.keys(sfPS)]);
  const sfl = [...codes].map(c => ({ c, n: anSfName(sfN[c] || sfNama[c] || c), v: sfV[c] || 0, vm: (sfV[c] || 0) / nM, desa: sfDesa[c] ? sfDesa[c].size : 0, ps: sfPS[c] || 0 }));
  const SB = [['Tanpa DJP', 0, 0], ['1–10', 0.01, 10], ['11–25', 10.01, 25], ['>25', 25.01, 1e9]];
  o.sfBL = SB.map(b => b[0]);
  o.sfB = SB.map(([l, lo, hi]) => { const s = sfl.filter(x => l === 'Tanpa DJP' ? x.v === 0 : x.vm >= lo && x.vm <= hi); return { l, n: s.length, avg: s.length ? sum(s.map(x => x.ps)) / s.length : null }; });
  const xs = sfl.filter(x => x.v > 0);
  if (xs.length > 2) { const mx = sum(xs.map(x => x.v)) / xs.length, my = sum(xs.map(x => x.ps)) / xs.length; let a = 0, b = 0, c = 0;
    xs.forEach(x => { a += (x.v - mx) * (x.ps - my); b += (x.v - mx) ** 2; c += (x.ps - my) ** 2; }); o.corr = b && c ? a / Math.sqrt(b * c) : null; } else o.corr = null;
  o.sfAll = sfl.length; o.sfSell = sfl.filter(x => x.ps > 0).length;
  const prodMin = Math.max(1, Math.round(5 * D.wDays / 14));
  o.prodMin = prodMin; o.sfProd = sfl.filter(x => x.ps >= prodMin).length;
  o.topSF = sfl.filter(x => x.n).sort((a, b) => b.ps - a.ps || b.v - a.v).slice(0, 7);
  o.watch = sfl.filter(x => x.vm >= 15 && x.ps === 0).sort((a, b) => b.v - a.v).slice(0, 4);
  // prioritas & kategori
  o.desaKat = AN.KAT.map(k => clDesa.filter(d => d.kat === k).length);
  o.psKat = AN.KAT.map(k => sum(clDesa.filter(d => d.kat === k).map(d => psD[anK(d.kec, d.desa)] || 0)));
  const pri = clDesa.filter(d => d.kat === 'prioritas');
  o.prio = [pri.length, pri.filter(d => desaV[anK(d.kec, d.desa)]).length, pri.filter(d => psD[anK(d.kec, d.desa)]).length];
  o.prioUntouched = pri.filter(d => !desaV[anK(d.kec, d.desa)]).sort((a, b) => (b.skorPotensi || 0) - (a.skorPotensi || 0)).map(d => anTitle(d.desa));
  // desa jenuh (kunjungan tinggi, PS rendah) & desa produktif
  const dl = Object.keys({ ...desaV, ...psD }).map(k => ({ k, desa: anTitle(k.split('|')[1]), kat: (katMap[k] || {}).kat || '-', v: desaV[k] || 0, ps: psD[k] || 0 }));
  const vMed = (() => { const a = Object.values(desaV).sort((x, y) => x - y); return a.length ? a[Math.floor(a.length * 0.9)] : 0; })();
  o.saturated = dl.filter(x => x.v >= Math.max(vMed, 10) && x.ps <= 2).sort((a, b) => b.v - a.v).slice(0, 3);
  o.topDesa = dl.filter(x => x.ps > 0).sort((a, b) => b.ps - a.ps).slice(0, 4);
  o.skipProd = dl.filter(x => x.kat === 'skip' && x.ps >= 3).sort((a, b) => b.ps - a.ps).slice(0, 3);
  // penetrasi
  const pen = (app.penetrasiData || []).filter(p => inK(p.kab));
  o.penZero = pen.filter(p => +p.ps === 0 && +p.avai >= 20).length;
  // plan bulan berjalan
  const pc = visCur.filter(v => inK(v.kab));
  o.planCur = pc.length; o.planCurDesa = new Set(pc.map(v => v.dk)).size; o.planCurSF = new Set(pc.map(v => v.sf)).size;
  o.planCurGood = anP(pc.filter(v => ['prioritas', 'strategis'].includes((katMap[v.dk] || {}).kat)).length, pc.length);
  return o;
}

// ================================================================
//  2. ANALISA (silogisme) — teks dari angka, bisa ditelusuri
// ================================================================
function anAnalyze(D) {
  const per = (f) => D.ents.map(e => `${anTitle(e.name)} ${f(e)}`).join(', ');
  const nv = v => (v == null || !isFinite(v)) ? 'tidak ada SF' : anD(v);
  const E = D.ents, nm = e => anTitle(e.name);
  const by = (f, dir = -1) => [...E].filter(e => f(e) != null && isFinite(f(e))).sort((a, b) => dir * (f(a) - f(b)));
  const best = (f) => by(f)[0], worst = (f) => by(f, 1)[0];
  const S = {};
  const djpLbl = D.availM.length ? D.availM.map(m => AN_BLN[periodMonth(m)] || m).join(', ') + ' ' + D.availM[D.availM.length - 1].slice(0, 4) : 'tidak ada data';
  const missing = D.djpM.filter(m => !D.availM.includes(m)).map(periodLabel);
  S.djpLbl = djpLbl; S.missing = missing;
  const mg = D.mode === 'mingguan';

  // pace / momentum
  if (mg) {
    const down = E.filter(e => e.io < (e.io0 || 0));
    S.pace = [
      'Target tercapai bila jumlah IO cukup dan minimal 70% IO berasal dari desa Prioritas dan Strategis. IO adalah satu-satunya hasil yang sepenuhnya dikendalikan sales.',
      E.map(e => `${nm(e)}: IO minggu ini ${e.io}, rata-rata ${anD(e.ioDrr)} IO per hari, padahal yang dibutuhkan ${anD(e.ioNeed)} IO per hari; ${anPs(Math.round(e.ioPotShare || 0))} IO berasal dari desa Prioritas dan Strategis`).join('; ') + '.',
      (E.filter(e => e.ioDrr < e.ioNeed).length ? `Laju IO di bawah kebutuhan: ${E.filter(e => e.ioDrr < e.ioNeed).map(e => `${nm(e)} (kurang ${anD(e.ioNeed - e.ioDrr)} IO per hari)`).join(', ')}. ` : 'Jumlah IO per hari di semua wilayah sudah memenuhi kebutuhan. ') +
        (E.filter(e => (e.ioPotShare || 0) < 70).length ? `Kurang dari 70% IO berasal dari desa Prioritas dan Strategis di ${E.filter(e => (e.ioPotShare || 0) < 70).map(nm).join(', ')}.` : 'Minimal 70% IO sudah berasal dari desa Prioritas dan Strategis.'),
    ];
  } else {
    const wi = worst(e => e.ioOlAch);
    S.pace = [
      'Target tercapai bila jumlah IO cukup dan minimal 70% IO berasal dari desa Prioritas dan Strategis. IO adalah satu-satunya hasil yang sepenuhnya dikendalikan sales; PS mengikuti IO.',
      E.map(e => `${nm(e)}: IO sampai hari ini ${e.ioMtd}, perkiraan akhir bulan ${e.ioOl} dari target ${e.ioTgt} (${anPs(Math.round(e.ioOlAch || 0))}); ${anPs(Math.round(e.ioPotShare || 0))} IO berasal dari desa Prioritas dan Strategis`).join('; ') + '.',
      D.sisa === 0 ? `Capaian IO terendah: ${nm(wi)}.` : `${nm(wi)} paling perlu dorongan: butuh ${anD(wi.ioNeed)} IO per hari untuk mencapai target, sedangkan rata-rata saat ini ${anD(wi.ioDrr)} IO per hari. Utamakan desa Prioritas dan Strategis.`,
    ];
  }
  // funnel
  const wf = worst(e => e.psre), bf = best(e => e.psre);
  const gap = wf && bf && wf !== bf ? Math.round(wf.re * (bf.psre - wf.psre) / 100) : 0;
  S.funnel = [
    'Order yang sudah lolos registrasi (RE) tetapi belum menjadi PS masih bisa diselamatkan bersama teknisi. Sebagian penyebabnya di luar kendali sales: jarak tarikan kabel ke ODP, jumlah teknisi, dan cuaca.',
    E.map(e => `${nm(e)}: ${anPs(Math.round(e.psre || 0))} order RE sudah menjadi PS, ${e.coh[1]} order batal (${anD(e.cancelPer100)} per 100 IO), ${e.coh[2]} order masih menunggu`).join('; ') + `. Porsi PS produk Tsel One: ${per(e => anPs(e.tselOne))}; IndiHome Reguler: ${per(e => anPs(e.ihReg))}.`,
    gap >= 3 ? `Bila persentase RE yang menjadi PS di ${nm(wf)} sama dengan ${nm(bf)}, ${nm(wf)} akan mendapat sekitar ${gap} PS tambahan. Prioritas: selesaikan ${wf.coh[2]} order ${nm(wf)} yang masih menunggu, kurangi order batal, dan dorong penjualan IndiHome Reguler karena nilai tagihan bulanannya (ARPU) lebih tinggi.`
      : `Persentase RE yang menjadi PS di semua wilayah relatif setara. Fokus pada ${sum(E.map(e => e.coh[2]))} order yang masih menunggu agar menjadi PS, dan jaga porsi penjualan IndiHome Reguler.`,
  ];
  // kendala
  { const kt = e => e.kend[0][0] + e.kend[0][2], kp = e => e.kend[1][0] + e.kend[1][2], ks = e => e.kend[2][0] + e.kend[2][2];
    const T = sum(E.map(kt)), P = sum(E.map(kp));
    S.kendala = [
      'Kendala teknik berarti calon pelanggan ada, tetapi jaringan belum siap (misalnya ODP penuh, ODP jauh, atau tidak ada ODP). Kendala pelanggan berarti calon pelanggan batal, tidak bisa dihubungi, atau alamatnya tidak ditemukan. Cara menanganinya berbeda.',
      E.map(e => `${nm(e)}: ${kt(e)} order terhambat kendala teknik, ${kp(e)} kendala pelanggan, ${ks(e)} kendala sistem` + (e.kendTop[0] ? `; alasan terbanyak: ${e.kendTop[0][0].split(': ').slice(1).join(': ')} (${e.kendTop[0][1]} order)` : '')).join('. ') + '.',
      T > P ? `Kendala terbesar ada di jaringan. Kirim daftar lokasi ODP penuh dan ODP jauh ke tim jaringan sebagai usulan penambahan ODP, dan jangan dihitung sebagai kegagalan SF.`
        : `Kendala terbesar ada di sisi pelanggan. Pastikan data calon pelanggan lengkap dan benar sebelum order diinput, dan hubungi ulang pelanggan yang belum bisa dihubungi.`,
    ]; }
  // DJP wilayah
  const agg = E.map(e => e.bucket);
  const tot = e => e.bucket.map(b => b.avg);
  const peakTxt = E.map(e => { const a = tot(e); let pk = -1, pv = -1; a.forEach((v, i) => { if (v != null && v > pv) { pv = v; pk = i; } });
    return { e, pk, pv, last: a[a.length - 1] }; });
  const satur = peakTxt.filter(x => x.pk > 0 && x.pk < D.ents[0].bucketL.length - 1 && x.last != null && x.last < x.pv);
  const satList = E.flatMap(e => e.saturated.map(s => `${s.desa} (${s.v} kunjungan, ${s.ps} PS)`)).slice(0, 3);
  const totVis = sum(E.map(e => e.vis)), totPS = sum(E.map(e => e.ps));
  S.djp = [
    `Bila kunjungan DJP efektif, desa yang lebih sering dikunjungi SF dalam 3 bulan sebelumnya (${djpLbl}) akan menghasilkan PS lebih banyak, sampai batas tertentu ketika tambahan kunjungan tidak lagi menambah PS.`,
    `Rata-rata PS per desa bulan ini, dikelompokkan menurut rata-rata jumlah kunjungan per bulan (${D.ents[0].bucketL.join(', ')} kunjungan per bulan): ` + E.map(e => `${nm(e)} berturut-turut ${e.bucket.map(b => b.avg == null ? 'tidak ada desa' : anD(b.avg)).join(', ')} PS`).join('; ') + '.' + (satList.length ? ` Desa yang sering dikunjungi tetapi PS-nya sedikit: ${satList.join(', ')}.` : ''),
    `${anPs(anP(sum(E.map(e => e.psFromVis)), totPS))} PS bulan ini berasal dari desa yang dikunjungi DJP dalam 3 bulan sebelumnya` +
      (satur.length ? `. Di ${satur.map(x => nm(x.e)).join(', ')}, rata-rata PS per desa paling tinggi pada kelompok ${D.ents[0].bucketL[satur[0].pk]} kunjungan per bulan, lalu turun pada desa yang lebih sering dikunjungi. Kunjungan berlebih sebaiknya dialihkan ke desa Prioritas yang belum pernah dikunjungi.` : '. Makin sering desa dikunjungi, makin tinggi PS-nya, jadi frekuensi kunjungan di desa yang menghasilkan PS perlu dipertahankan.') +
      ' Catatan: desa yang sering dikunjungi biasanya memang desa besar, jadi angka ini menunjukkan hubungan, belum tentu sebab-akibat.',
  ];
  // SF
  const hi = E.map(e => e.sfB[3].avg), lo = E.map(e => e.sfB[1].avg);
  const lowN = E.map(e => e.sfB[1].n);
  S.sf = [
    'SF yang rutin mengunjungi desa wilayah kerjanya (kunjungan DJP) akan menghasilkan PS lebih banyak.',
    `Rata-rata PS per SF bulan ini: SF dengan lebih dari 25 kunjungan per bulan ${per(e => nv(e.sfB[3].avg))}; SF dengan 1–10 kunjungan per bulan ${per(e => nv(e.sfB[1].avg))}. Keeratan hubungan antara jumlah kunjungan dan PS per SF (skala 0 = tidak berhubungan sampai 1 = sangat berhubungan): ${per(e => anD(e.corr, 2))}. Rata-rata satu SF mengunjungi ${per(e => anD(e.desaPerSf) + ' desa')}.`,
    `Masih ada SF yang hanya 1–10 kali kunjungan per bulan (${E.map(e => `${nm(e)} ${e.sfB[1].n} SF`).join(', ')}); jumlah kunjungan mereka ke desa wilayah kerjanya perlu ditambah.` +
      (sum(E.map(e => e.watch.length)) ? ` Ada ${sum(E.map(e => e.watch.length))} SF yang sering kunjungan DJP tetapi belum menghasilkan PS; mereka perlu dibimbing soal cara menawarkan produk.` : ''),
  ];
  // prioritas
  const skipNames = E.flatMap(e => e.skipProd.map(s => `${s.desa} ${s.ps} PS`)).slice(0, 2);
  const noPlan = E.filter(e => e.planCur === 0);
  S.prio = [
    'Desa Prioritas dan Strategis (potensi tertinggi) seharusnya menjadi sumber IO dan PS terbesar, dan desa Skip yang terkecil.',
    `Porsi PS dari desa Prioritas dan Strategis: ${per(e => anPs(anP(e.psKat[0] + e.psKat[1], e.ps)))}; dari desa Skip: ${per(e => anPs(anP(e.psKat[3], e.ps)))}` + (skipNames.length ? ` (misalnya ${skipNames.join(', ')})` : '') +
      `. Desa Prioritas yang belum pernah dikunjungi DJP dalam 3 bulan terakhir: ${per(e => e.prioUntouched.length + ' desa')}. Jumlah kunjungan yang direncanakan di plan DJP ${periodLabel(D.bulan)}: ${per(e => e.planCur + ' kunjungan')}.`,
    (skipNames.length ? 'Desa Skip yang ternyata menghasilkan PS jangan ditinggalkan; kategorinya perlu ditinjau ulang. ' : '') + `Alihkan kunjungan dari desa yang sering dikunjungi tetapi PS-nya sedikit ke ${sum(E.map(e => e.prioUntouched.length))} desa Prioritas yang belum pernah dikunjungi` + (noPlan.length ? `, dan ${noPlan.map(nm).join(', ')} segera membuat plan DJP ${periodLabel(D.bulan)} di dashboard.` : '.'),
  ];
  // kesimpulan utama
  const wo = worst(e => mg ? e.drr - e.need : e.olAch);
  S.main = [];
  S.main.push([`IO: perkiraan akhir bulan dibanding target: ${per(e => anPs(Math.round(e.ioOlAch || 0)))}.`, ` Porsi IO dari desa Prioritas dan Strategis: ${per(e => anPs(Math.round(e.ioPotShare || 0)))} (standar minimal 70%). Paling perlu dorongan: ${nm(worst(e => e.ioOlAch))}, butuh ${anD(worst(e => e.ioOlAch).ioNeed)} IO per hari sampai akhir bulan.`]);
  if (mg) S.main.push([`Momentum minggu ini: ${E.map(e => `${nm(e)} ${anSg(anP(e.ps - e.ps0, e.ps0))}`).join(', ')}.`, ` Paling perlu dorongan: ${nm(wo)} (butuh ${anD(wo.need)} PS per hari sampai akhir bulan).`]);
  else S.main.push([`PS (hasil akhir), perkiraan akhir bulan dibanding target: ${per(e => anPs(Math.round(e.olAch)))}.`, D.sisa ? (E.every(e => (e.olAch || 0) >= 100) ? ' Semua wilayah diperkirakan melampaui target.' : ` Perkiraan terendah: ${nm(wo)}.`) : '']);
  S.main.push([gap >= 3 ? `Di ${nm(wf)}, banyak order tertahan setelah lolos registrasi.` : 'Order yang sudah lolos registrasi umumnya menjadi PS.', ` Persentase RE yang menjadi PS: ${per(e => anPs(Math.round(e.psre || 0)))}; ${sum(E.map(e => e.coh[2]))} order masih menunggu pemasangan (perlu bantuan teknisi).`]);
  S.main.push([`Kunjungan DJP ${djpLbl}:`, ` ${anPs(anP(sum(E.map(e => e.psFromVis)), totPS))} PS bulan ini berasal dari desa yang dikunjungi dalam 3 bulan tersebut${satur.length ? '; beberapa desa sudah sering dikunjungi tetapi PS-nya sedikit' : ''}.`]);
  S.main.push(['Rata-rata PS per SF menurut jumlah kunjungannya:', ` SF dengan lebih dari 25 kunjungan per bulan ${per(e => nv(e.sfB[3].avg))} PS; SF dengan 1–10 kunjungan per bulan ${per(e => nv(e.sfB[1].avg))} PS.`]);
  S.main.push([noPlan.length ? `Plan DJP ${periodLabel(D.bulan)} untuk ${noPlan.map(nm).join(', ')} masih kosong.` : `Plan DJP ${periodLabel(D.bulan)} sudah dibuat (${sum(E.map(e => e.planCur))} kunjungan direncanakan).`, ` Porsi PS dari desa kategori Skip: ${per(e => anPs(anP(e.psKat[3], e.ps)))}. Bila porsinya besar, desa tersebut sebenarnya produktif dan kategorinya perlu ditinjau.`]);
  if (missing.length) S.main.push([`Data DJP ${missing.join(', ')} belum tersedia.`, ' Analisa DJP hanya memakai bulan yang datanya ada.']);
  S.final = [
    'IO yang tepat sasaran didapat bila SF rutin mengunjungi desa Prioritas dan Strategis (desa yang jauh lewat jadwal khusus) dan kunjungannya berkualitas. PS mengikuti IO, tetapi juga dipengaruhi faktor teknis di luar kendali sales: jarak rumah ke ODP maksimal 250 meter, jumlah teknisi, dan cuaca.',
    E.map(e => `${nm(e)}: perkiraan IO akhir bulan ${anPs(Math.round(e.ioOlAch || 0))} dari target, ${anPs(Math.round(e.ioPotShare || 0))} IO dari desa Prioritas dan Strategis, rata-rata ${anD(e.visPerSfM, 0)} kunjungan per SF per bulan, ${anPs(Math.round(e.psre || 0))} RE menjadi PS, ${mg ? 'PS minggu ini ' + anSg(anP(e.ps - e.ps0, e.ps0)) + ' dibanding minggu lalu' : 'perkiraan PS akhir bulan ' + anPs(Math.round(e.olAch)) + ' dari target'}`).join('; ') + '.',
    (() => { const wi = worst(e => e.ioOlAch); return D.sisa === 0 ? `Bulan berikutnya: ${nm(wi)} menjadi fokus, dengan kunjungan rutin ke desa Prioritas dan Strategis sejak awal bulan.` : `Fokus sampai akhir bulan: ${nm(wi)} menaikkan IO menjadi sekitar ${anD(wi.ioNeed, 0)} IO per hari dari desa Prioritas dan Strategis (desa Strategis lewat jadwal khusus per kecamatan), dan order yang masih menunggu dikawal bersama teknisi.`; })(),
  ];
  // aksi
  const A = [];
  { const ioE = E.filter(e => e.ioNeed > e.ioDrr + 0.05 || (e.ioPotShare || 0) < 70);
    if (ioE.length) A.push(['Kejar IO dari desa Prioritas dan Strategis (minimal 70% dari seluruh IO); desa Strategis dijadwalkan per kecamatan', ioE.map(nm).join(', '), 'TL', ioE.map(e => `${nm(e)} minimal ${anD(e.ioNeed, 0)} IO per hari`).join('; '), 'Mulai besok']); }
  const pendE = E.filter(e => e.coh[2] >= 5 || (e.psre || 100) < 80);
  if (pendE.length) A.push(['Kawal order yang masih menunggu dan yang batal: buat daftar order yang lebih dari 3 hari belum PS, lalu eskalasi ke tim teknisi (TA/HCOT)', pendE.map(nm).join(', '), 'TL + PJ', `Order menunggu (${pendE.map(e => `${nm(e)} ${e.coh[2]}`).join(', ')}) turun minimal 60%; minimal 80% RE menjadi PS`, 'Minggu ini']);
  const lowE = E.filter(e => e.sfB[1].n > 0);
  if (lowE.length) A.push([`Tambah kunjungan DJP SF yang baru 1–10 kali per bulan, ke desa wilayah kerjanya (minimal 3 kali per desa per bulan)`, lowE.map(nm).join(', '), [...new Set(lowE.map(e => e.tl))].join(', '), `${sum(lowE.map(e => e.sfB[1].n))} SF naik menjadi lebih dari 10 kunjungan per bulan`, 'Mulai minggu depan']);
  const unt = E.flatMap(e => e.prioUntouched.slice(0, 2));
  if (satList.length || unt.length) A.push([`Alihkan kunjungan dari desa yang sering dikunjungi tetapi PS-nya sedikit${satList.length ? ' (' + E.flatMap(e => e.saturated.map(s => s.desa)).slice(0, 3).join(', ') + ')' : ''} ke desa Prioritas yang belum pernah dikunjungi${unt.length ? ' (' + unt.slice(0, 3).join(', ') + ')' : ''}`, E.map(nm).join(', '), 'TL + PJ', `${sum(E.map(e => e.prioUntouched.length))} desa Prioritas mulai dikunjungi`, 'Plan minggu depan']);
  if (noPlan.length) A.push([`Buat plan DJP ${periodLabel(D.bulan)} di dashboard (saat ini masih kosong)`, noPlan.map(nm).join(', '), [...new Set(noPlan.map(e => e.tl))].join(', '), 'Minimal 80% kunjungan ke desa Prioritas dan Strategis', 'Segera']);
  const needE = E.filter(e => e.need > e.drrM + 0.05);
  if (needE.length) A.push(['Tambah jumlah SF yang turun per hari, dan datangi desa yang masih punya minimal 20 port ODP kosong tetapi belum ada PS', needE.map(nm).join(', '), [...new Set(needE.map(e => e.tl))].join(', '), needE.map(e => `${nm(e)} minimal ${anD(e.need, 0)} PS per hari`).join('; '), 'Sisa bulan']);
  A.push(['Bimbing SF yang sering kunjungan tetapi belum ada PS; tinjau ulang desa Skip yang ternyata menghasilkan PS menjadi Strategis; dorong penjualan IndiHome Reguler', E.map(nm).join(', '), 'PJ', 'Tidak ada lagi SF yang sering kunjungan tanpa PS; porsi IndiHome Reguler naik', 'Akhir bulan']);
  const allN = E.map(nm).join(', ');
  A.forEach(a => { if (a[1] === allN && E.length > 2) a[1] = 'Semua'; });
  S.actions = A.slice(0, 6);
  return S;
}

// ================================================================
//  3. BANGUN PPT (7 slide)
// ================================================================
async function anBuild(D, S, B) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = 'Dashboard WOK LamWay'; pres.title = 'Analisa PS ' + D.label;
  const { F, RED, DARK, GREY, BG, GRN, AMB } = AN;
  let logo = null;
  try { const b = await (await fetch('assets/img/indihome-logo.png')).blob(); logo = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(b); }); } catch (e) {}
  const objs = [{ line: { x: 0.4, y: 0.95, w: 12.53, h: 0, line: { color: RED, width: 1.5 } } },
    { text: { text: 'Internal', options: { x: 12.1, y: 7.12, w: 0.85, h: 0.25, fontFace: F, fontSize: 8, color: GREY, align: 'right', margin: 0 } } }];
  if (logo) objs.unshift({ image: { data: logo, x: 11.55, y: 0.22, w: 1.42, h: 0.4 } });
  pres.defineSlideMaster({ title: 'AN', background: { color: 'FFFFFF' }, objects: objs, slideNumber: { x: 0.4, y: 7.12, w: 0.4, h: 0.25, fontFace: F, fontSize: 8, color: GREY } });

  const E = D.ents, k = E.length, names = E.map(e => e.short), cols = E.map(e => e.color);
  const mg = D.mode === 'mingguan';
  const tx = (s, t, o) => s.addText(t, { fontFace: F, margin: 0, isTextBox: true, ...o });
  const head = (s, t, sub) => { tx(s, t, { x: 0.4, y: 0.22, w: 10.9, h: 0.42, fontSize: 20, bold: true, color: DARK }); tx(s, sub, { x: 0.4, y: 0.62, w: 10.9, h: 0.28, fontSize: 10, color: GREY }); };
  const SRC = `Sumber: Monitoring HH (order unik) • kunjungan DJP dari DJP_Realisasi_Baru • plan dari DJP_Plan • kategori desa dari menu Prioritas • periode PS: ${D.label}${mg ? ' dibanding ' + D.label0 : ''} • periode kunjungan DJP: ${S.djpLbl}; 1 kunjungan = 1 SF datang ke 1 desa pada 1 tanggal` + (S.missing.length ? ` • data DJP ${S.missing.join(', ')} belum tersedia` : '');
  const foot = (s, extra) => tx(s, SRC + (extra || ''), { x: 0.85, y: 7.1, w: 11.1, h: 0.3, fontSize: 7.5, color: GREY });
  const card = (s, x, y, w, h, title, color) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: 'FFFFFF' }, line: { color: 'E3E3E3', width: 1 } });
    if (title) { s.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: y + 0.14, w: 0.06, h: 0.22, fill: { color: color || RED }, line: { color: color || RED } });
      tx(s, title, { x: x + 0.28, y: y + 0.08, w: w - 0.4, h: 0.34, fontSize: 11, bold: true, color: DARK, valign: 'middle' }); }
  };
  const syl = (s, x, y, w, h, m) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: AN.PINK }, line: { color: RED, width: 1.25 } });
    tx(s, 'ALUR ANALISA', { x: x + 0.2, y: y + 0.08, w: w - 0.4, h: 0.3, fontSize: 10.5, bold: true, color: RED });
    const rh = (h - 0.45) / 3;
    [['KENAPA PENTING', GREY], ['YANG TERJADI', AMB], ['ARTINYA', RED]].forEach(([l, c], i) => {
      const yy = y + 0.42 + i * rh;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.2, y: yy + 0.04, w: 1.3, h: 0.28, rectRadius: 0.05, fill: { color: c }, line: { color: c } });
      tx(s, l, { x: x + 0.2, y: yy + 0.04, w: 1.3, h: 0.28, fontSize: 7.5, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
      tx(s, m[i], { x: x + 1.6, y: yy, w: w - 1.8, h: rh - 0.04, fontSize: m[i].length > 300 ? 7.5 : 8.5, color: DARK, valign: 'top', bold: i === 2, fit: 'shrink' });
    });
  };
  const tile = (s, x, y, w, h, label, val, sub, color) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.06, fill: { color: BG }, line: { color: 'E8E8E8', width: 0.75 } });
    tx(s, label, { x: x + 0.12, y: y + 0.04, w: w - 0.24, h: 0.3, fontSize: 8.5, color: GREY, fit: 'shrink' });
    const L = String(val).length;
    tx(s, String(val), { x: x + 0.12, y: y + 0.3, w: w - 0.24, h: 0.45, fontSize: L > 18 ? 11 : L > 12 ? 13 : L > 8 ? 15 : 19, bold: true, color: color || DARK, valign: 'middle' });
    if (sub) tx(s, sub, { x: x + 0.12, y: y + h - 0.34, w: w - 0.24, h: 0.3, fontSize: 7.5, color: GREY, fit: 'shrink' });
  };
  const ax = { catAxisLabelColor: GREY, valAxisLabelColor: GREY, catAxisLabelFontFace: F, valAxisLabelFontFace: F, catAxisLabelFontSize: 9, valAxisLabelFontSize: 8,
    valGridLine: { color: 'EAEAEA', size: 0.5 }, catGridLine: { style: 'none' }, dataLabelFontFace: F, dataLabelFontSize: 8, legendFontFace: F, legendFontSize: 9,
    titleFontFace: F, titleFontSize: 10, titleColor: DARK, showTitle: true, showLegend: true, legendPos: 't' };
  const bar = (s, labels, series, o) => s.addChart(pres.charts.BAR, series.map(x => ({ name: x.name, labels, values: x.values.map(v => v == null ? 0 : v) })),
    { barDir: 'col', barGrouping: 'clustered', barGapWidthPct: 50, ...ax, showValue: true, dataLabelPosition: 'outEnd', dataLabelColor: DARK, ...o });
  const entSeries = f => E.map(e => ({ name: e.short, values: f(e) }));
  const pctBar = (s, cats, colors, fv, o) => s.addChart(pres.charts.BAR, cats.map((c, i) => ({ name: c, labels: names, values: E.map(e => fv(e)[i]) })),
    { barDir: 'bar', barGrouping: 'percentStacked', barGapWidthPct: 45, ...ax, chartColors: colors, valAxisLabelFormatCode: '0%', showValue: false, ...o });
  const donutPair = (s, x, y, w, h, title, cats, colors, fv) => {
    card(s, x, y, w, h);
    tx(s, title, { x: x + 0.2, y: y + 0.08, w: w - 0.4, h: 0.25, fontSize: 10, bold: true, color: DARK });
    if (k <= 2) {
      E.forEach((e, i) => {
        s.addChart(pres.charts.DOUGHNUT, [{ name: e.short, labels: cats, values: fv(e) }], { x: x + 0.05 + i * (w / 2), y: y + 0.33, w: w / 2 - 0.1, h: h - 1.0, holeSize: 55,
          chartColors: colors, showPercent: true, showValue: false, showLegend: false, dataLabelColor: DARK, dataLabelFontSize: 8, dataLabelFontFace: F, showTitle: false });
        tx(s, e.short, { x: x + i * (w / 2), y: y + h - 0.68, w: w / 2, h: 0.22, fontSize: 9, bold: true, color: e.color, align: 'center' });
      });
    } else {
      const tot = cats.map((c, i) => sum(E.map(e => fv(e)[i])));
      s.addChart(pres.charts.DOUGHNUT, [{ name: 'Total', labels: cats, values: tot }], { x: x + 0.05, y: y + 0.33, w: w * 0.42, h: h - 1.0, holeSize: 55,
        chartColors: colors, showPercent: true, showLegend: false, dataLabelColor: DARK, dataLabelFontSize: 8, dataLabelFontFace: F, showTitle: false });
      tx(s, 'Total', { x: x, y: y + h - 0.68, w: w * 0.45, h: 0.22, fontSize: 9, bold: true, color: DARK, align: 'center' });
      pctBar(s, cats, colors, fv, { x: x + w * 0.45, y: y + 0.3, w: w * 0.53, h: h - 0.75, showTitle: false, showLegend: false, catAxisLabelFontSize: 7 });
    }
    cats.forEach((c, i) => {
      const cw = (w - 0.3) / cats.length;
      s.addShape(pres.shapes.RECTANGLE, { x: x + 0.18 + i * cw, y: y + h - 0.33, w: 0.12, h: 0.12, fill: { color: colors[i] }, line: { color: colors[i] } });
      tx(s, c, { x: x + 0.33 + i * cw, y: y + h - 0.4, w: cw - 0.2, h: 0.26, fontSize: 7, color: GREY });
    });
  };
  const L2 = '  •  ';
  const modeTxt = { mingguan: 'Review Mingguan', bulanan: 'Review Bulanan', cluster: 'Detail Cluster ' + (E[0] && E[0].cl), all: 'Detail Semua Cluster' }[D.mode];

  // ── 1. RINGKASAN ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, `ANALISA PS ${mg ? 'MINGGUAN' : periodLabel(D.bulan).toUpperCase()} DAN KUNJUNGAN DJP 3 BULAN SEBELUMNYA`, `${modeTxt}${L2}PS periode ${D.label}${L2}kunjungan DJP ${S.djpLbl}${L2}${E.map(e => anTitle(e.name)).join(' dibanding ')}`);
    if (k === 2) {
      E.forEach((e, i) => {
        const x = 0.4 + i * 4.2;
        card(s, x, 1.15, 4.05, 5.8, e.name.toUpperCase(), e.color);
        tx(s, e.tl, { x: x + 0.28, y: 1.45, w: 3.6, h: 0.25, fontSize: 8.5, color: GREY });
        tx(s, [{ text: anI(e.ps), options: { fontSize: 40, bold: true, color: e.color } }, { text: mg ? `  PS minggu ini (minggu lalu ${e.ps0})` : `  PS dari target ${e.tgtM}`, options: { fontSize: 11, color: GREY } }],
          { x: x + 0.28, y: 1.75, w: 3.7, h: 0.75, valign: 'middle' });
        const t = [
          mg ? ['Dibanding minggu lalu', anSg(anP(e.ps - e.ps0, e.ps0)), `dari ${e.ps0} menjadi ${e.ps} PS`, e.ps >= e.ps0 ? GRN : RED]
            : ['Perkiraan PS akhir bulan', `${e.ol} (${Math.round(e.olAch)}%)`, `rata-rata ${anD(e.drrM)} PS per hari`, e.olAch >= 100 ? GRN : RED],
          mg ? ['Capaian target minggu ini', anPs(Math.round(e.ach)), `target minggu ini ${anD(e.tgtW, 0)} PS`, e.ach >= 100 ? GRN : RED]
            : ['Dibanding tanggal yang sama bulan lalu', anSg(anP(e.ps - e.ps0, e.ps0)), `dari ${e.ps0} menjadi ${e.ps} PS`, e.ps >= e.ps0 ? GRN : RED],
          ['RE yang menjadi PS', anPs(Math.round(e.psre || 0)), `IO ${e.io}, RE ${e.re}`, (e.psre || 0) >= 80 ? GRN : AMB],
          e.mtd >= e.tgtM ? ['PS per hari yang dibutuhkan', 'Target tercapai', `PS sudah ${e.mtd} dari target ${e.tgtM}`, GRN] : ['PS per hari yang dibutuhkan', anD(e.need), `untuk ${D.sisa} hari tersisa`, e.need > e.drrM ? RED : GRN],
          e.disc && e.disc.sfPlan ? ['Jadwal SF di plan yang terlaksana', anPs(e.disc.sfPct), `${e.disc.sfHit} dari ${e.disc.sfPlan} jadwal bulan ini`, e.disc.sfPct >= 80 ? GRN : e.disc.sfPct >= 50 ? AMB : RED]
            : [`Kunjungan DJP ${S.djpLbl}`, anI(e.vis), `oleh ${e.sfDJP} SF ke ${e.desaTouched} desa`, DARK],
          ['PS dari desa yang dikunjungi DJP', anPs(anP(e.psFromVis, e.ps)), `${e.psFromVis} dari ${e.ps} PS`, DARK],
        ];
        t.forEach((v, j) => tile(s, x + 0.2 + (j % 2) * 1.85, 2.65 + Math.floor(j / 2) * 1.4, 1.75, 1.28, ...v));
      });
    } else {
      card(s, 0.4, 1.15, 8.25, 5.8, 'KPI PER KABUPATEN', RED);
      const hdr = ['Indikator', ...names].map(t => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: RED }, align: 'center' } }));
      const rowsT = [
        ['PS ' + (mg ? 'minggu ini' : 'sejak tanggal 1'), ...E.map(e => anI(e.ps))], ['Target PS sebulan', ...E.map(e => anI(e.tgtM))],
        ['Perkiraan PS akhir bulan', ...E.map(e => `${e.ol} (${Math.round(e.olAch)}%)`)], ['Dibanding periode lalu', ...E.map(e => anSg(anP(e.ps - e.ps0, e.ps0)))],
        ['PS per hari: rata-rata / dibutuhkan', ...E.map(e => `${anD(e.drrM)} / ${anD(e.need)}`)], ['RE yang menjadi PS', ...E.map(e => anPs(Math.round(e.psre || 0)))],
        [`Kunjungan DJP ${S.djpLbl}`, ...E.map(e => anI(e.vis))], ['SF yang kunjungan / desa dikunjungi', ...E.map(e => `${e.sfDJP} SF / ${e.desaTouched} desa`)],
        ['PS dari desa yang dikunjungi DJP', ...E.map(e => anPs(anP(e.psFromVis, e.ps)))], ['Kunjungan di plan DJP bulan ini', ...E.map(e => anI(e.planCur))],
        ['Jadwal SF di plan yang terlaksana', ...E.map(e => e.disc && e.disc.sfPlan ? `${anPs(e.disc.sfPct)} (${e.disc.sfHit}/${e.disc.sfPlan})` : '–')],
      ];
      s.addTable([hdr, ...rowsT.map((r, i) => r.map((v, j) => ({ text: String(v), options: { fill: { color: i % 2 ? 'FFFFFF' : BG }, bold: j === 0, align: j ? 'center' : 'left' } })))],
        { x: 0.6, y: 1.6, w: 7.85, colW: [2.2, ...E.map(() => 5.65 / k)], rowH: 0.4, fontFace: F, fontSize: 9, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.05 });
    }
    card(s, 8.85, 1.15, 4.08, 5.8, 'KESIMPULAN UTAMA', RED);
    const pts = S.main.slice(0, 5), step = 5.2 / pts.length;
    pts.forEach((p, i) => {
      const y = 1.62 + i * step;
      s.addShape(pres.shapes.OVAL, { x: 9.05, y: y + 0.03, w: 0.36, h: 0.36, fill: { color: RED }, line: { color: RED } });
      tx(s, String(i + 1), { x: 9.05, y: y + 0.03, w: 0.36, h: 0.36, fontSize: 10, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
      tx(s, [{ text: p[0], options: { bold: true, color: DARK } }, { text: p[1], options: { color: '404040' } }], { x: 9.52, y, w: 3.28, h: step - 0.06, fontSize: 9, valign: 'top', fit: 'shrink' });
    });
    foot(s);
  }

  // ── 2. PENCAPAIAN & TREN ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'PENCAPAIAN & TREN PS', `Tren 5 bulan, capaian ${mg ? 'minggu ini dibanding minggu lalu' : 'dibanding target yang seharusnya tercapai sampai hari ini'}, dan PS harian ${D.label}`);
    card(s, 0.4, 1.12, 6.15, 2.95);
    const ml = D.trendM.map((m, i) => anMonthShort(m) + (i === 4 && D.refDay < D.days ? ' (perkiraan)' : ''));
    s.addChart(pres.charts.LINE, E.map(e => ({ name: e.short, labels: ml, values: [...e.monthly.slice(0, 4), D.refDay < D.days ? e.ol : e.monthly[4]] })),
      { x: 0.5, y: 1.18, w: 5.95, h: 2.85, ...ax, title: 'PS per Bulan' + (D.refDay < D.days ? ' (bulan ini: perkiraan sampai akhir bulan)' : ''), chartColors: cols, lineSize: 2.5, lineDataSymbolSize: 7,
        showValue: true, dataLabelPosition: 't', dataLabelColor: DARK });
    card(s, 6.78, 1.12, 6.15, 2.95);
    if (mg) bar(s, names, [{ name: 'Minggu lalu', values: E.map(e => e.ps0) }, { name: 'Minggu ini', values: E.map(e => e.ps) }, { name: 'Target 7 hari', values: E.map(e => Math.round(e.tgtW)) }],
      { x: 6.88, y: 1.18, w: 5.95, h: 2.85, title: 'PS Minggu Ini, Minggu Lalu, dan Target', chartColors: ['BFBFBF', RED, '595959'] });
    else bar(s, names, [{ name: 'PS ' + D.label, values: E.map(e => e.ps) }, { name: 'Target yang seharusnya tercapai sampai hari ini', values: E.map(e => Math.round(e.tgtToDate)) }],
      { x: 6.88, y: 1.18, w: 5.95, h: 2.85, title: 'PS Dibanding Target Sampai Hari Ini', chartColors: [RED, 'D9D9D9'] });
    card(s, 0.4, 4.22, 6.15, 2.78);
    const dl = []; for (let n = D.W[0]; n <= D.W[1]; n = addDays(n, 1)) dl.push(String(n % 100));
    bar(s, dl, entSeries(e => e.daily), { x: 0.5, y: 4.28, w: 5.95, h: 2.68, title: 'PS per Hari (data 1–2 hari terakhir mungkin belum lengkap)', chartColors: cols, showValue: dl.length <= 10, barGapWidthPct: 40 });
    syl(s, 6.78, 4.22, 6.15, 2.78, S.pace);
    foot(s);
  }

  // ── 3. FUNNEL, PRODUK & CHANNEL ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'ALUR ORDER, PRODUK & CHANNEL', `Perjalanan order dari IO ke RE ke PS, dan komposisi PS ${D.label}`);
    card(s, 0.4, 1.12, 4.1, 3.05);
    bar(s, names, [{ name: 'IO', values: E.map(e => e.io) }, { name: 'RE', values: E.map(e => e.re) }, { name: 'PS', values: E.map(e => e.ps) }],
      { x: 0.5, y: 1.18, w: 3.9, h: 2.95, title: 'Jumlah IO, RE, dan PS', chartColors: ['D9D9D9', '8C8C8C', RED], dataLabelFontSize: k > 3 ? 6 : 8 });
    card(s, 4.62, 1.12, 4.1, 3.05);
    pctBar(s, ['PS', 'Batal', 'Masih menunggu'], [GRN, RED, AMB], e => e.coh, { x: 4.72, y: 1.18, w: 3.9, h: 2.95, title: 'Status Order dari IO Periode Ini', showValue: true, dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF' });
    const pc = AN.PROD.filter((p, i) => sum(E.map(e => e.prod[i])) > 0), pi = pc.map(p => AN.PROD.indexOf(p));
    donutPair(s, 8.84, 1.12, 4.09, 3.05, 'Mix Produk PS', pc, pi.map(i => AN.PROD_C[i]), e => pi.map(i => e.prod[i]));
    const tl = [['Rata-rata harga paket (ARPU)', anJoin(E.map(e => e.arpu ? 'Rp' + Math.round(e.arpu / 1000) + ' ribu' : '–')), 'dari semua order Februari sampai sekarang', DARK], ['PS paket 20 Mbps', anJoin(E.map(e => anPs(e.sp20))), 'porsi dari seluruh PS', DARK],
      ['Order batal per 100 IO', anJoin(E.map(e => anD(e.cancelPer100))), 'jumlah batal: ' + anJoin(E.map(e => e.coh[1])), RED], ['PS IndiHome Reguler', anJoin(E.map(e => anPs(e.ihReg))), 'porsi dari seluruh PS; tagihan bulanan lebih tinggi', AMB]];
    tl.forEach((v, i) => tile(s, 0.4 + (i % 2) * 3.2, 4.32 + Math.floor(i / 2) * 1.36, 3.08, 1.26, ...v));
    tx(s, 'Urutan angka di kotak: ' + names.join(' / '), { x: 0.4, y: 6.98, w: 6.3, h: 0.14, fontSize: 7, color: GREY, italic: true });
    syl(s, 6.92, 4.32, 6.01, 2.68, S.funnel);
    foot(s);
  }

  // ── 3b. KENDALA ORDER ──
  if (E.some(e => e.kendTot || sum((e.kend || []).map(x => x[1])))) {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'KENDALA ORDER', `Order dari IO ${D.label} yang tercatat ada kendala, dikelompokkan menjadi kendala teknik (jaringan), kendala pelanggan, dan kendala sistem`);
    const KL = ['Teknik (jaringan)', 'Pelanggan', 'Sistem (aplikasi)'];
    card(s, 0.4, 1.12, 6.2, 3.3);
    bar(s, names, KL.map((l, i) => ({ name: l, values: E.map(e => e.kend[i][0] + e.kend[i][2]) })),
      { x: 0.5, y: 1.18, w: 6.0, h: 3.2, title: 'Order kendala yang batal atau belum selesai, per jenis kendala', chartColors: [RED, AMB, GREY], dataLabelFontSize: 8 });
    card(s, 6.72, 1.12, 6.21, 3.3, 'ALASAN KENDALA TERBANYAK', RED);
    const rowsK = [];
    E.forEach(e => e.kendTop.forEach(([al, n], i) => rowsK.push([i ? '' : anTitle(e.name), al, n])));
    const hdr = ['Wilayah', 'Alasan (order batal atau belum selesai)', 'Jumlah order'].map(t => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: RED } } }));
    s.addTable([hdr, ...rowsK.slice(0, k > 2 ? 10 : 12).map((r, i) => r.map((v, j) => ({ text: String(v), options: { fill: { color: i % 2 ? 'FFFFFF' : BG }, bold: j === 0, align: j === 2 ? 'center' : 'left' } })))],
      { x: 6.85, y: 1.55, w: 5.95, colW: [1.5, 3.35, 1.1], rowH: 0.22, fontFace: F, fontSize: 8, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.03 });
    const tl = E.map(e => { const t = e.kend.map(x => x[0] + x[2]); const fin = sum(e.kend.map(x => x[1]));
      return [anTitle(e.name), `${t[0]} / ${t[1]} / ${t[2]}`, `teknik / pelanggan / sistem · ${fin} order lain tetap jadi PS`, DARK]; });
    const tw = Math.min(3.0, (6.2 - 0.1 * (k - 1)) / k);
    tl.forEach((v, i) => tile(s, 0.4 + i * (tw + 0.1), 4.55, tw, 1.3, ...v));
    syl(s, 6.72, 4.55, 6.21, 2.4, S.kendala);
    foot(s);
  }

  // ── 4. DJP 3 BULAN → PS (WILAYAH) ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, `KUNJUNGAN DJP 3 BULAN SEBELUMNYA DAN PS PER DESA`, `Kunjungan SF ke desa ${S.djpLbl} dihubungkan dengan PS ${D.label} (1 kunjungan = 1 SF datang ke 1 desa pada 1 tanggal)`);
    card(s, 0.4, 1.12, 4.1, 2.95);
    bar(s, D.djpM.map(m => anMonthShort(m) + (D.availM.includes(m) ? '' : ' (tidak ada data)')), entSeries(e => e.visByM), { x: 0.5, y: 1.18, w: 3.9, h: 2.85, title: 'Jumlah Kunjungan Desa per Bulan', chartColors: cols, dataLabelFontSize: k > 3 ? 6 : 8 });
    card(s, 4.62, 1.12, 4.1, 2.95);
    pctBar(s, AN.KAT_L, AN.KAT_C, e => e.katVis, { x: 4.72, y: 1.18, w: 3.9, h: 2.85, title: `Kunjungan per Kategori Desa (ke desa Prioritas: ${anJoin(E.map(e => anPs(e.prioVisShare)))})` });
    card(s, 8.84, 1.12, 4.09, 2.95);
    bar(s, E[0].bucketL, entSeries(e => e.bucket.map(b => b.avg == null ? 0 : Math.round(b.avg * 10) / 10)), { x: 8.94, y: 1.18, w: 3.89, h: 2.85, title: 'Rata-rata PS per Desa menurut Jumlah Kunjungan', chartColors: cols, dataLabelFormatCode: '0.0',
      catAxisTitle: 'rata-rata kunjungan SF ke desa per bulan', showCatAxisTitle: true, catAxisTitleFontSize: 8, catAxisTitleColor: GREY, catAxisTitleFontFace: F, dataLabelFontSize: k > 3 ? 6 : 8 });
    const tl = [['PS dari desa yang dikunjungi DJP', anJoin(E.map(e => anPs(anP(e.psFromVis, e.ps)))), anJoin(E.map(e => e.psFromVis + ' PS')), DARK],
      ['PS oleh SF di desa yang ia kunjungi', anJoin(E.map(e => anPs(anP(e.psSamePair, e.ps)))), 'PS di desa yang dikunjungi minimal 3 kali: ' + anJoin(E.map(e => e.psRutin)), DARK],
      ['Kunjungan per SF per bulan', anJoin(E.map(e => anD(e.visPerSfM, 0))), 'desa per SF: ' + anJoin(E.map(e => anD(e.desaPerSf))), AMB],
      ['Desa yang pernah dikunjungi DJP', anJoin(E.map(e => e.desaTouched)), 'dari ' + anJoin(E.map(e => e.desaTotal)) + ' desa', DARK]];
    tl.forEach((v, i) => tile(s, 0.4 + (i % 2) * 2.05, 4.22 + Math.floor(i / 2) * 1.4, 1.95, 1.28, ...v));
    syl(s, 4.62, 4.22, 8.31, 2.6, S.djp);
    foot(s, L2 + 'urutan angka di kotak: ' + names.join(' / '));
  }

  // ── 5. SF ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'SF: KUNJUNGAN KE DESA WILAYAH KERJA DAN PS', `SF dikelompokkan menurut rata-rata jumlah kunjungan desa per bulan (${S.djpLbl})`);
    card(s, 0.4, 1.12, 4.1, 2.95);
    bar(s, E[0].sfBL, entSeries(e => e.sfB.map(b => b.avg == null ? 0 : Math.round(b.avg * 10) / 10)), { x: 0.5, y: 1.18, w: 3.9, h: 2.85, title: 'Rata-rata PS per SF menurut jumlah kunjungan per bulan', chartColors: cols, dataLabelFormatCode: '0.0', dataLabelFontSize: k > 3 ? 6 : 8 });
    card(s, 4.62, 1.12, 4.1, 2.95);
    bar(s, E[0].sfBL, entSeries(e => e.sfB.map(b => b.n)), { x: 4.72, y: 1.18, w: 3.9, h: 2.85, title: 'Jumlah SF menurut jumlah kunjungan per bulan', chartColors: cols, dataLabelFontSize: k > 3 ? 6 : 8 });
    card(s, 8.84, 1.12, 4.09, 2.95, 'TOP SF PS ' + (mg ? 'MINGGU INI' : 'PERIODE INI'), RED);
    const top = E.flatMap(e => e.topSF.map(x => ({ ...x, ent: e.short }))).sort((a, b) => b.ps - a.ps).slice(0, 7);
    const hdr = ['SF', k > 2 ? 'Kabupaten' : 'Cluster', 'Desa', 'Kunjungan', 'PS'].map(t => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: RED } } }));
    s.addTable([hdr, ...top.map((r, i) => [r.n.slice(0, 20), r.ent.slice(0, 10), r.desa, r.v, r.ps].map(v => ({ text: String(v), options: { fill: { color: i % 2 ? 'FFFFFF' : BG } } })))],
      { x: 9.0, y: 1.55, w: 3.77, colW: [1.35, 0.8, 0.45, 0.72, 0.45], rowH: 0.3, fontFace: F, fontSize: 8, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.03 });
    const nW = sum(E.map(e => e.watch.length));
    const tl = [['SF yang punya minimal 1 PS', anJoin(E.map(e => e.sfSell)), 'dari ' + anJoin(E.map(e => e.sfAll)) + ' SF', DARK],
      [`SF dengan minimal ${E[0].prodMin} PS`, anJoin(E.map(e => e.sfProd)), anJoin(E.map(e => anPs(anP(e.sfProd, e.sfSell)))) + ' dari SF yang punya PS', AMB],
      ['Hubungan kunjungan dan PS', anJoin(E.map(e => anD(e.corr, 2))), '0 = tidak berhubungan, 1 = sangat berhubungan', DARK], ['SF sering kunjungan tetapi 0 PS', anJoin(E.map(e => e.watch.length)), 'perlu dibimbing', RED]];
    tl.forEach((v, i) => tile(s, 0.4 + (i % 2) * 2.05, 4.22 + Math.floor(i / 2) * 1.4, 1.95, 1.28, ...v));
    if (nW) tx(s, 'SF yang sering kunjungan tetapi belum ada PS: ' + E.flatMap(e => e.watch.map(x => `${x.n} (${x.v} kunjungan, ${e.short})`)).slice(0, 5).join(' • '), { x: 0.4, y: 6.86, w: 12.5, h: 0.22, fontSize: 7.5, color: RED, italic: true });
    syl(s, 4.62, 4.22, 8.31, 2.6, S.sf);
    foot(s, L2 + 'urutan angka di kotak: ' + names.join(' / '));
  }

  // ── 6. PRIORITAS, PENETRASI & PLAN ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'KATEGORI DESA, PENETRASI & PLAN DJP', `Porsi PS per kategori desa, desa Prioritas yang sudah dikunjungi, dan status plan DJP ${periodLabel(D.bulan)}`);
    donutPair(s, 0.4, 1.12, 4.1, 2.95, 'PS per Kategori Desa', AN.KAT_L, AN.KAT_C, e => e.psKat);
    card(s, 4.62, 1.12, 4.1, 2.95);
    bar(s, names, [{ name: 'Jumlah desa Prioritas', values: E.map(e => e.prio[0]) }, { name: `Dikunjungi ${S.djpLbl}`, values: E.map(e => e.prio[1]) }, { name: 'Sudah ada PS', values: E.map(e => e.prio[2]) }],
      { x: 4.72, y: 1.18, w: 3.9, h: 2.85, title: 'Desa Prioritas: Dikunjungi dan Ada PS', chartColors: ['D9D9D9', AMB, RED], dataLabelFontSize: k > 3 ? 6 : 8 });
    card(s, 8.84, 1.12, 4.09, 2.95, 'DESA: PS DAN JUMLAH KUNJUNGAN DJP', RED);
    const dRows = [...E.flatMap(e => e.topDesa.slice(0, k > 2 ? 2 : 3)).sort((a, b) => b.ps - a.ps).slice(0, 4), ...E.flatMap(e => e.saturated).sort((a, b) => b.v - a.v).slice(0, 3)];
    const hdr = ['Desa', 'Kategori', 'Kunjungan', 'PS'].map(t => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: RED } } }));
    const nTop = Math.min(4, dRows.length);
    s.addTable([hdr, ...dRows.map((r, i) => [r.desa.slice(0, 20), anTitle(r.kat), r.v, r.ps].map(v => ({ text: String(v), options: { fill: { color: i >= nTop ? 'FDECEC' : (i % 2 ? 'FFFFFF' : BG) } } })))],
      { x: 9.0, y: 1.55, w: 3.77, colW: [1.42, 0.8, 0.85, 0.7], rowH: 0.26, fontFace: F, fontSize: 8, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.03 });
    tx(s, 'Baris merah muda = desa yang sering dikunjungi tetapi PS-nya sedikit', { x: 9.0, y: 3.83, w: 3.8, h: 0.18, fontSize: 7, color: GREY, italic: true });
    const tl = [[`Kunjungan di plan DJP ${periodLabel(D.bulan)}`, anJoin(E.map(e => e.planCur)), anJoin(E.map(e => e.planCur ? `${e.planCurSF} SF, ${anPs(e.planCurGood)} ke Prioritas dan Strategis` : 'belum ada plan')), E.some(e => !e.planCur) ? RED : DARK],
      ['Desa Prioritas belum pernah dikunjungi DJP', anJoin(E.map(e => e.prioUntouched.length)), E.flatMap(e => e.prioUntouched.slice(0, 1)).slice(0, 2).join(', ') || '—', DARK],
      ['Desa dengan minimal 20 port kosong, belum ada PS', anJoin(E.map(e => e.penZero)), 'peluang penjualan', AMB],
      ['Porsi PS dari desa Prioritas', anJoin(E.map(e => anPs(anP(e.psKat[0], e.ps)))), 'dari desa Skip: ' + anJoin(E.map(e => anPs(anP(e.psKat[3], e.ps)))), DARK]];
    tl.forEach((v, i) => tile(s, 0.4 + (i % 2) * 2.05, 4.22 + Math.floor(i / 2) * 1.4, 1.95, 1.28, ...v));
    syl(s, 4.62, 4.22, 8.31, 2.6, S.prio);
    foot(s, L2 + 'urutan angka di kotak: ' + names.join(' / '));
  }

  // ── 6b. DISIPLIN PLAN vs REALISASI DJP ──
  if (D.discRange) {
    const s = pres.addSlide({ masterName: 'AN' });
    const [f, t] = D.discRange, fl = x => +x.slice(8, 10) + ' ' + anMonthShort(x.slice(0, 7));
    head(s, 'DISIPLIN PLAN DAN REALISASI KUNJUNGAN DJP', `Jadwal kunjungan di plan (DJP_Plan) dibanding kunjungan yang benar-benar terjadi (DJP_Realisasi_Baru), ${fl(f)} – ${fl(t)}`);
    const tiles = [['Jadwal SF yang terlaksana', e => anPs(e.disc.sfPct), e => `${e.disc.sfHit} dari ${e.disc.sfPlan} jadwal: SF tercatat kunjungan pada tanggal plan`, e => e.disc.sfPct == null ? GREY : e.disc.sfPct >= 80 ? GRN : e.disc.sfPct >= 50 ? AMB : RED],
      ['Jadwal desa yang didatangi', e => anPs(anP(e.disc.desaHit, e.disc.desaPlan)), e => `${e.disc.desaHit} dari ${e.disc.desaPlan} jadwal: desa didatangi pada tanggal plan`, e => { const p = anP(e.disc.desaHit, e.disc.desaPlan); return p == null ? GREY : p >= 80 ? GRN : p >= 50 ? AMB : RED; }],
      ['Kunjungan di luar plan', e => anI(sum(e.disc.luar)), () => 'kunjungan ke desa yang tidak dijadwalkan pada tanggal itu', () => DARK]];
    const tw = Math.min(2.0, (4.1 - 0.1 * (k - 1)) / k);
    tiles.forEach(([lbl, fv, fs, fc], r) => {
      tx(s, lbl, { x: 0.4, y: 1.15 + r * 1.6, w: 4.1, h: 0.25, fontSize: 10, bold: true, color: DARK });
      E.forEach((e, i) => tile(s, 0.4 + i * (tw + 0.1), 1.42 + r * 1.6, tw, 1.2, anTitle(e.name), fv(e), fs(e), fc(e)));
    });
    card(s, 4.62, 1.12, 8.31, 3.05);
    bar(s, AN.KAT_L, E.map(e => ({ name: anTitle(e.name), values: e.disc.kat.map(c => c.plan ? Math.round(c.hit / c.plan * 100) : 0) })),
      { x: 4.72, y: 1.18, w: 8.1, h: 2.95, title: 'Persentase jadwal desa yang didatangi, per kategori desa', chartColors: cols, valAxisMaxVal: 100, valAxisLabelFormatCode: '0"%"', dataLabelFormatCode: '0"%"' });
    const hdr = ['Kategori', ...E.flatMap(e => [anTitle(e.name) + ': jadwal di plan', 'didatangi', 'di luar plan'])].map(x => ({ text: x, options: { bold: true, color: 'FFFFFF', fill: { color: RED }, align: 'center' } }));
    const rowsK = AN.KAT_L.map((l, ki) => [l, ...E.flatMap(e => { const c = e.disc.kat[ki]; return [c.plan, c.plan ? `${c.hit} (${anP(c.hit, c.plan)}%)` : '–', e.disc.luar[ki]]; })]);
    const nc = 1 + 3 * k, cw0 = 1.3, cwi = (8.31 - cw0) / (3 * k);
    s.addTable([hdr, ...rowsK.map((r, i) => r.map((v, j) => ({ text: String(v), options: { fill: { color: i % 2 ? 'FFFFFF' : BG }, bold: j === 0, align: j ? 'center' : 'left', color: j === 0 ? AN.KAT_C[i] === 'D9D9D9' ? GREY : AN.KAT_C[i] : DARK } })))],
      { x: 4.62, y: 4.32, w: 8.31, colW: [cw0, ...Array(nc - 1).fill(cwi)], rowH: 0.36, fontFace: F, fontSize: k > 3 ? 6.5 : 8.5, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.04 });
    tx(s, `Satu jadwal SF = satu SF yang dijadwalkan kunjungan pada satu tanggal. Satu jadwal desa = satu desa yang dijadwalkan dikunjungi pada satu tanggal. Dihitung sampai ${fl(t)}, tanggal terakhir data kunjungan yang sudah di-upload. Bila file DJP tanggal itu diunduh sebelum hari ditutup, angkanya masih bisa bertambah.`,
      { x: 4.62, y: 6.25, w: 8.31, h: 0.3, fontSize: 8, italic: true, color: GREY });
    foot(s);
  }

  // ── 7. DRIVER TREE & GAP IO ──
  if (B) {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'APA SAJA YANG MEMBENTUK IO, DAN BERAPA KEKURANGANNYA', 'Kotak berwarna merah adalah faktor yang paling menghambat. Perubahan RE menjadi PS juga dipengaruhi jarak ke ODP, teknisi, dan cuaca.');
    const n = B.drv.length, cw = Math.min(4.1, (12.53 - (n - 1) * 0.12) / n);
    B.drv.slice(0, 6).forEach((x, i) => {
      const xx = 0.4 + i * (cw + 0.12); card(s, xx, 1.12, cw, 3.25, x.name.toUpperCase(), E[i] ? E[i].color : RED);
      x.rows.forEach((r, j) => {
        const yy = 1.55 + j * 0.39, c = r.st === 'ok' ? GRN : r.st === 'warn' ? AMB : RED;
        s.addShape(pres.shapes.RECTANGLE, { x: xx + 0.15, y: yy, w: 0.06, h: 0.32, fill: { color: c }, line: { color: c } });
        tx(s, r.label.replace(' (di luar kendali penuh sales)', '*'), { x: xx + 0.27, y: yy - 0.02, w: cw - 1.2, h: 0.22, fontSize: n > 3 ? 6.5 : 8, color: DARK });
        tx(s, 'pembanding: ' + r.ref, { x: xx + 0.27, y: yy + 0.17, w: cw - 1.2, h: 0.17, fontSize: 6.5, color: GREY });
        tx(s, r.val, { x: xx + cw - 0.95, y: yy, w: 0.8, h: 0.32, fontSize: n > 3 ? 9 : 11, bold: true, color: c, align: 'right', valign: 'middle' });
      });
    });
    tx(s, '* Perubahan RE menjadi PS juga dipengaruhi hal di luar kendali sales, jadi tidak dipilih sebagai faktor penghambat.', { x: 0.4, y: 4.38, w: 12, h: 0.2, fontSize: 7, color: GREY, italic: true });
    const hdr = ['Wilayah', 'Target IO', 'IO sejauh ini', 'Perkiraan IO akhir bulan', 'Kekurangan IO', 'IO dari desa Prioritas dan Strategis', 'Tambahan IO yang dibutuhkan dari desa tersebut', 'PS / target PS', 'Penjelasan'].map(t => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: RED } } }));
    s.addTable([hdr, ...B.gap.map((g, i) => [g.name, g.tgt, g.mtd, g.ol, g.gap, `${g.pot} (${g.potShare}%)`, g.potGap ? '+' + g.potGap : '–', `${g.ps}/${g.psTgt}`, ((B.diag[i] || {}).lines || []).slice(1, 3).map(l => l[1]).join(' ')].map((v, j) => ({ text: String(v), options: { fill: { color: i % 2 ? 'FFFFFF' : BG }, bold: j === 4 && g.gap > 0, color: j === 4 && g.gap > 0 ? RED : DARK } })))],
      { x: 0.4, y: 4.62, w: 12.53, colW: [1.3, 0.75, 0.75, 0.8, 0.65, 1.15, 1.0, 0.75, 5.38], rowH: 0.36, fontFace: F, fontSize: n > 3 ? 6.5 : 7.5, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.04 });
    foot(s);
  }
  // ── 8. MATRIKS DESA & JADWAL STRATEGIS ──
  if (B) {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'DESA PRIORITAS & STRATEGIS: SISA POTENSI DAN JADWAL KHUSUS', 'Sisa port = port ODP yang masih kosong dikurangi PS. Tanda (Strategis) = desa berpotensi tinggi tetapi jauh dari kantor.');
    B.matrix.forEach((q, i) => {
      const xx = 0.4, yy = 1.12 + i * 1.95, col = q.c.replace('#', '');
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: xx, y: yy, w: 7.25, h: 1.85, rectRadius: 0.06, fill: { color: 'FFFFFF' }, line: { color: col, width: 2 } });
      tx(s, [{ text: q.t.replace(/^\S+\s/, ''), options: { bold: true, color: col, fontSize: 12 } }, { text: `   ${q.n} desa · ${q.s}`, options: { color: GREY, fontSize: 8.5 } }], { x: xx + 0.15, y: yy + 0.08, w: 7, h: 0.3 });
      tx(s, q.top.slice(0, 4).map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < Math.min(3, q.top.length - 1) } })), { x: xx + 0.15, y: yy + 0.4, w: 7, h: 1.05, fontSize: 7.5, color: DARK, valign: 'top', paraSpaceAfter: 2 });
      tx(s, 'Yang perlu dilakukan: ' + q.act, { x: xx + 0.15, y: yy + 1.48, w: 7, h: 0.3, fontSize: 8, bold: true, color: col, valign: 'top', fit: 'shrink' });
    });
    card(s, 7.82, 1.12, 5.11, 5.8, 'JADWAL KHUSUS DESA STRATEGIS YANG JARANG DIKUNJUNGI', RED);
    const lines = [];
    B.strat.forEach(t => {
      lines.push({ text: `${t.name}: ${t.kurang} dari ${t.total} desa Strategis dikunjungi kurang dari 2 kali per bulan`, options: { bold: true, color: DARK, breakLine: true, fontSize: 9 } });
      t.grup.slice(0, 3).forEach(gp => lines.push({ text: `Kecamatan ${gp.kec} (sekitar ${gp.jarak} km): ${gp.desa.slice(0, 4).join(', ')}`, options: { bullet: true, breakLine: true, fontSize: 8, color: '404040' } }));
    });
    if (!lines.length) lines.push({ text: 'Semua desa Strategis sudah dikunjungi rutin.', options: { fontSize: 9 } });
    tx(s, lines, { x: 8.0, y: 1.55, w: 4.8, h: 5.2, valign: 'top', paraSpaceAfter: 3, fit: 'shrink' });
    foot(s);
  }
  // ── 9. KESIMPULAN & AKSI ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'KESIMPULAN & RENCANA AKSI', `Kesimpulan dari analisa sebelumnya dan keputusan untuk ${D.sisa} hari tersisa di bulan ini`);
    card(s, 0.4, 1.12, 4.1, 2.95);
    bar(s, names, [{ name: 'Rata-rata IO per hari saat ini', values: E.map(e => Math.round(e.ioDrr * 10) / 10) }, { name: 'IO per hari yang dibutuhkan', values: E.map(e => Math.round(e.ioNeed * 10) / 10) }],
      { x: 0.5, y: 1.18, w: 3.9, h: 2.85, title: `IO per hari: saat ini dan yang dibutuhkan (${D.sisa} hari tersisa)`, chartColors: ['8C8C8C', RED], dataLabelFormatCode: '0.0', dataLabelFontSize: k > 3 ? 6 : 8 });
    syl(s, 4.62, 1.12, 8.31, 2.95, S.final);
    const hdr = ['#', 'Aksi', 'Untuk', 'Ukuran keberhasilan', 'Waktu'].map(t => ({ text: t, options: { bold: true, color: 'FFFFFF', fill: { color: RED } } }));
    s.addTable([hdr, ...S.actions.map((r, i) => [String(i + 1), r[0], r[1], r[3], r[4]].map((v, j) => ({ text: v, options: { fill: { color: i % 2 ? 'FFFFFF' : BG }, bold: j === 0, align: j === 0 ? 'center' : 'left' } })))],
      { x: 0.4, y: 4.22, w: 12.53, colW: [0.4, 5.83, 1.8, 3.1, 1.4], rowH: 0.4, fontFace: F, fontSize: 8, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.05 });
    foot(s);
  }
  // ── 10. KAMUS ISTILAH ──
  {
    const s = pres.addSlide({ masterName: 'AN' });
    head(s, 'KAMUS ISTILAH', 'Penjelasan istilah yang dipakai di laporan ini');
    const K = [['IO (Input Order)', 'Order yang diinput sales. Satu-satunya hasil yang sepenuhnya dikendalikan sales.'], ['RE (Registrasi)', 'IO yang sudah lolos registrasi dan validasi.'],
      ['PS (Put In Service)', 'Layanan sudah terpasang dan aktif. Dipengaruhi juga oleh faktor di luar sales: jarak rumah ke ODP maksimal 250 meter, jumlah teknisi, dan cuaca.'],
      ['Prioritas', 'Desa berpotensi tinggi dan dekat kantor; dikunjungi rutin.'], ['Strategis', 'Desa berpotensi tinggi tetapi jauh dari kantor; dikunjungi lewat jadwal khusus per kecamatan.'],
      ['Cadangan dan Skip', 'Desa dengan potensi lebih rendah; dikunjungi secukupnya.'], ['Kunjungan DJP', 'Satu SF datang ke satu desa pada satu tanggal. Diambil dari DJP_Realisasi_Baru untuk 3 bulan sebelum bulan laporan.'],
      ['Plan DJP', 'Jadwal kunjungan SF ke desa yang dibuat di menu Plan DJP dashboard (DJP_Plan).'],
      ['Perkiraan akhir bulan', 'Rata-rata per hari sejauh ini dikali jumlah hari dalam bulan.'], ['Target sampai hari ini', 'Target sebulan dibagi jumlah hari, dikali jumlah hari yang sudah berjalan.'],
      ['Faktor pembentuk IO', 'IO diuraikan menjadi jumlah SF, jumlah kunjungan, sasaran kunjungan, dan IO per kunjungan, supaya terlihat faktor mana yang menghambat.'],
      ['ARPU', 'Rata-rata harga paket per bulan dari order yang masuk.'], ['Kendala teknik, pelanggan, sistem', 'Teknik: jaringan belum siap (ODP penuh, jauh, tidak ada). Pelanggan: batal, tidak bisa dihubungi, alamat tidak ditemukan. Sistem: gangguan aplikasi order.'],
      ['Port kosong', 'Port ODP yang masih tersedia untuk pelanggan baru.'], ['Penetrasi', 'PS dibagi port ODP; menunjukkan seberapa banyak potensi yang sudah terambil.'],
      ['Sisa potensi desa', 'Belum tergarap: port kosong banyak, PS sedikit. Sudah tergarap: sebagian besar port sudah terisi. Perlu ditinjau: port hampir habis, atau sering dikunjungi tetapi tidak ada PS.'],
      ['Alur analisa', 'Kenapa hal itu penting, apa yang terjadi menurut data, dan artinya untuk kita.']];
    const half = Math.ceil(K.length / 2);
    [K.slice(0, half), K.slice(half)].forEach((col, ci) => {
      s.addTable(col.map((r, i) => [{ text: r[0], options: { bold: true, fill: { color: i % 2 ? 'FFFFFF' : BG } } }, { text: r[1], options: { fill: { color: i % 2 ? 'FFFFFF' : BG } } }]),
        { x: 0.4 + ci * 6.33, y: 1.15, w: 6.2, colW: [1.7, 4.5], rowH: 0.6, fontFace: F, fontSize: 8, color: DARK, border: { type: 'solid', pt: 0.5, color: 'E3E3E3' }, valign: 'middle', margin: 0.06 });
    });
    foot(s);
  }
  return pres;
}

async function exportAnalisaPPT(app, mode) {
  await loadPptLib();
  const D = await anCollect(app, mode);
  if (!D.ents.length) throw new Error('Tidak ada data cluster');
  const S = anAnalyze(D);
  let B = null; try { if (typeof wlBizAnalysis === 'function') B = await wlBizAnalysis(app, D, D.ents); } catch (e) { console.error(e); }
  const pres = await anBuild(D, S, B);
  const tgl = typeof localYMD === 'function' ? localYMD(new Date()) : new Date().toISOString().slice(0, 10);
  const tag = { mingguan: 'Mingguan-LB-vs-WK', bulanan: 'Bulanan-LB-vs-WK', cluster: 'Cluster-' + app.cluster.replace(/\s+/g, ''), all: 'Semua-Kabupaten' }[mode] || mode;
  const name = `Analisa-PS-${tag}-${tgl}.pptx`;
  await pres.writeFile({ fileName: name });
  return name;
}
