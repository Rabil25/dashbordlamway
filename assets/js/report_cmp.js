// ================================================================
//  PPT Perbandingan Cluster — Mingguan & Bulanan (MTD)
//  Satu deck, kedua cluster dibandingkan per slide: siapa unggul, selisihnya,
//  dan berapa PS yang bisa didapat bila cluster tertinggal menyamai yang unggul.
//  Bergantung pada report.js (loadPptLib, RP, R_FONT, fmt*, pct, sgn, sum).
//
//  Jendela waktu (pakai kolom tanggal absolut DH/DI/DJ = yyyymmdd):
//    mingguan : 7 hari s.d. hari acuan  vs  7 hari sebelumnya
//    bulanan  : tgl 1 s.d. hari acuan (MTD) vs  hari yang sama bulan lalu
//  Konversi = PS / (PS + batal) dari order yang IO-nya di jendela itu (order
//  yang masih berjalan tidak dihitung, agar minggu terakhir tidak tampak buruk).
// ================================================================

const CMP_NAMES = Object.keys(CLUSTER);   // ['Lampung Barat', 'Way Kanan']

// ── util tanggal yyyymmdd (angka) ──
const ymdNum = v => { const s = String(v || '').trim(); return /^20\d{6}$/.test(s) ? +s : 0; };
const ymdOf = d => d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
const dateOfYmd = n => new Date(Math.floor(n / 10000), Math.floor(n / 100) % 100 - 1, n % 100);
const addDays = (n, k) => { const d = dateOfYmd(n); d.setDate(d.getDate() + k); return ymdOf(d); };
const fmtTglPendek = n => dateOfYmd(n).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' });
const N = s => (s || '').toString().trim().toUpperCase();
const shortCl = c => c === 'Lampung Barat' ? 'LB' : c === 'Way Kanan' ? 'WK' : c;

// Kelompok alasan batal dari kolom Reason Kendala (CY) + Kategori (CZ)
function cancelGroup(r) {
  const d = N(r[C.KENDALA_DTL]), k = N(r[C.KENDALA_CAT]);
  const s = d + ' ' + k;
  if (/TAGGING|KOORDINAT|ALAMAT|SALAH LOKASI/.test(s)) return 'lokasi';
  if (/ODP|JALUR|RUTE|TARIKAN|TIANG|KABEL|PORT|TEKNI/.test(s)) return 'jaringan';
  if (/BATAL|GANTI PAKET|PENDING|PELANGGAN|TIDAK|RAGU|MAHAL|DOUBLE|DUPLIKAT/.test(s)) return 'pelanggan';
  return 'lain';
}
const CG_LABEL = { lokasi: 'Salah tagging/lokasi', jaringan: 'Jaringan (ODP penuh/jauh/jalur)', pelanggan: 'Pelanggan batal/ganti', lain: 'Tanpa keterangan/lain' };
const CG_KEYS = ['lokasi', 'jaringan', 'pelanggan', 'lain'];
const cgl = g => CG_LABEL[g].charAt(0).toLowerCase() + CG_LABEL[g].slice(1);   // huruf kecil di awal saja (ODP tetap)

// ================================================================
//  1. KUMPULKAN DATA (kedua cluster, 3 bulan untuk wilayah basis)
// ================================================================
async function cmpCollect(app, mode) {
  const bulan = app.cfg.bulan;
  const days = daysInPeriod(bulan);
  const refDay = Math.min(app.hariBerjalan, days);
  const end = +(bulan.replace('-', '') + String(refDay).padStart(2, '0'));
  let W, W0, label, label0;
  if (mode === 'mingguan') {
    W = [addDays(end, -6), end]; W0 = [addDays(end, -13), addDays(end, -7)];
    label = `${fmtTglPendek(W[0])} – ${fmtTglPendek(W[1])}`; label0 = `${fmtTglPendek(W0[0])} – ${fmtTglPendek(W0[1])}`;
  } else {
    const prev = prevPeriod(bulan), dPrev = daysInPeriod(prev);
    W = [+(bulan.replace('-', '') + '01'), end];
    W0 = [+(prev.replace('-', '') + '01'), +(prev.replace('-', '') + String(Math.min(refDay, dPrev)).padStart(2, '0'))];
    label = refDay < days ? `MTD 1–${refDay} ${periodLabel(bulan)}` : periodLabel(bulan);
    label0 = `1–${Math.min(refDay, dPrev)} ${periodLabel(prev)}`;
  }

  const auth = { headers: { Authorization: 'Bearer ' + app.sessionToken } };
  const getJSON = async url => { try { const r = await fetch(url, auth); const d = await r.json(); return d.ok ? d : null; } catch (e) { return null; } };

  // Bulan ke-3 (P-2) untuk wilayah basis; bulan ini & lalu sudah ada di allRows
  const p2 = prevPeriod(prevPeriod(bulan));
  const [d2, planCur, planPrev] = await Promise.all([
    getJSON(`/api/data?type=bulan_ini&bulan=${encodeURIComponent(p2)}`),
    getJSON(`/api/data?type=djp_plan&bulan=${encodeURIComponent(bulan)}`),
    getJSON(`/api/data?type=djp_plan&bulan=${encodeURIComponent(prevPeriod(bulan))}`),
  ]);
  const rowsP2 = (d2 && (d2.data || [])) || [];
  const plans = [...((planCur && planCur.data) || []), ...((planPrev && planPrev.data) || [])];

  if (!app.prioritasOdpRaw.length) { try { await app.loadPrioritas(); } catch (e) {} }
  const stack = document.body._x_dataStack;
  const base = stack && stack[0] && window.Alpine && Alpine.raw ? Alpine.raw(stack[0]) : null;
  if (!base || !('cluster' in base)) throw new Error('Data dashboard belum siap, coba lagi');

  // semua baris unik (bulan ini, lalu, P-2), tanpa duplikat order
  const seen = new Set(), all = [];
  [...app.allRows, ...rowsP2].forEach(r => {
    if ((r[C.DEDUP] || '').trim() !== 'Unik') return;
    const id = (r[C.ID_ORDER] || '').toString().trim();
    if (id && seen.has(id)) return; if (id) seen.add(id);
    all.push(r);
  });
  const hasDesa = all.some(r => (r[C.DESA] || '').toString().trim());

  const clusters = CMP_NAMES.map(name => {
    const kabs = CLUSTER[name].kab.map(k => k.toUpperCase());
    const view = Object.create(base);
    Object.defineProperty(view, 'cluster', { value: name, writable: true });
    const rows = all.filter(r => kabs.includes(N(r[C.VAL_KAB])));
    const sf = view.buildSFTable();
    const desa = app.prioritasOdpRaw.length ? view.klasifikasiDesa('', '') : [];
    const tgtPS = sum(kabs.map(k => (app.cfg.targets[k] || {}).ps || 0));
    const plan = plans.filter(p => kabs.includes(N(p.kab)));
    return { name, kabs, rows, sf, desa, tgtPS, plan, outlook: view.outlookData };
  });

  return { mode, bulan, days, refDay, W, W0, label, label0, p2, hasDesa, clusters, sisa: Math.max(days - refDay, 0) };
}

// ================================================================
//  2. METRIK per cluster untuk satu jendela
// ================================================================
function cmpMetrics(D, cl, W) {
  const inW = n => n >= W[0] && n <= W[1];
  const m = { io: 0, re: 0, ps: 0, coh: 0, cohPS: 0, cohCancel: 0, cohPend: 0, cg: { lokasi: 0, jaringan: 0, pelanggan: 0, lain: 0 },
    ch: {}, pr: {}, daily: {}, sf: {} };
  cl.rows.forEach(r => {
    const dIO = ymdNum(r[C.DATE_IO]), dRE = ymdNum(r[C.DATE_RE]), dPS = ymdNum(r[C.DATE_PS]);
    const vIO = (r[C.VAL_IO] || '').toString().trim(), vRE = (r[C.VAL_RE] || '').toString().trim(), vPS = (r[C.VAL_PS] || '').toString().trim();
    const code = (r[C.CODE_SF] || '').toString().trim();
    if (inW(dRE) && vRE) m.re++;
    if (inW(dPS) && vPS) { m.ps++; m.daily[dPS] = (m.daily[dPS] || 0) + 1; if (code) (m.sf[code] = m.sf[code] || { io: 0, ps: 0 }).ps++; }
    if (inW(dIO) && vIO) {
      m.io++;
      if (code) (m.sf[code] = m.sf[code] || { io: 0, ps: 0 }).io++;
      const isPS = /PUT IN SERVICE/.test(N(vPS)) || !!dPS;
      const isCancel = /CANCEL|BATAL/.test(N(r[C.STATUS_FNL]));
      const ch = (r[C.CHANNEL2] || 'Lainnya').toString().trim() || 'Lainnya', pr = (r[C.PRODUK_KAT] || 'Lainnya').toString().trim() || 'Lainnya';
      const c = m.ch[ch] || (m.ch[ch] = { io: 0, ps: 0, fin: 0 }), p = m.pr[pr] || (m.pr[pr] = { io: 0, ps: 0, fin: 0 });
      c.io++; p.io++;
      m.coh++;
      if (isPS) { m.cohPS++; c.ps++; p.ps++; c.fin++; p.fin++; }
      else if (isCancel) { m.cohCancel++; m.cg[cancelGroup(r)]++; c.fin++; p.fin++; }
      else m.cohPend++;
    }
  });
  const fin = m.cohPS + m.cohCancel;
  m.conv = fin ? m.cohPS / fin * 100 : null;          // konversi order selesai
  m.cancelPer100 = m.coh ? m.cohCancel / m.coh * 100 : null;
  // SF: roster cluster (Update SF) + penjual di jendela
  const roster = new Map(cl.sf.map(s => [s.code, s.nama]));
  const codes = new Set([...roster.keys(), ...Object.keys(m.sf)]);
  const dayLen = Math.round((dateOfYmd(W[1]) - dateOfYmd(W[0])) / 864e5) + 1;
  const prodMin = Math.max(1, Math.round(10 * dayLen / 30));     // ≥10 PS/bulan, diskalakan
  const per = [...codes].map(c => ({ code: c, nama: roster.get(c) || c, io: m.sf[c]?.io || 0, ps: m.sf[c]?.ps || 0 }));
  m.sfList = per;
  m.sfTotal = roster.size || per.length;
  m.sfAktif = per.filter(s => s.io > 0).length;
  m.sfPS = per.filter(s => s.ps > 0).length;
  m.sfProd = per.filter(s => s.ps >= prodMin).length;
  m.sf0 = per.filter(s => roster.has(s.code) && s.ps === 0);
  m.prodMin = prodMin;
  m.psPerSF = m.sfPS ? m.ps / m.sfPS : 0;
  const sorted = [...per].sort((a, b) => b.ps - a.ps), top = Math.max(1, Math.round(per.filter(s => s.ps > 0).length * 0.2));
  m.top20 = m.ps ? sum(sorted.slice(0, top).map(s => s.ps)) / m.ps * 100 : null;
  m.dayLen = dayLen;
  return m;
}

// ── DJP & wilayah basis (butuh desa order) ──
function cmpDJP(D, cl, W) {
  const key = (kec, desa) => N(kec) + '|' + N(desa);
  const katMap = {}; cl.desa.forEach(d => { katMap[key(d.kec, d.desa)] = d; });
  const toYmd = t => +String(t || '').replace(/-/g, '');
  const planW = cl.plan.filter(p => { const n = toYmd(p.tgl); return n >= W[0] && n <= W[1]; });
  const planMonth = cl.plan.filter(p => String(p.tgl).slice(0, 7) === D.bulan);
  const desaPlan = {}; planW.forEach(p => { const k = key(p.kec, p.desa); desaPlan[k] = (desaPlan[k] || 0) + 1; });
  const desaPlanMonth = new Set(planMonth.map(p => key(p.kec, p.desa)));
  const kat = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0, lain: 0 };
  Object.keys(desaPlan).forEach(k => kat[katMap[k]?.kat || 'lain']++);
  // IO jendela per desa
  const ioDesa = {};
  if (D.hasDesa) cl.rows.forEach(r => { const n = ymdNum(r[C.DATE_IO]); if (n >= W[0] && n <= W[1] && (r[C.VAL_IO] || '').toString().trim()) { const k = key(r[C.KEC], r[C.DESA]); ioDesa[k] = (ioDesa[k] || 0) + 1; } });
  const planned = Object.keys(desaPlan), others = Object.keys(katMap).filter(k => !desaPlan[k]);
  const avg = ks => ks.length ? sum(ks.map(k => ioDesa[k] || 0)) / ks.length : null;
  const prio = cl.desa.filter(d => d.kat === 'prioritas');
  const prioUntouched = prio.filter(d => !desaPlanMonth.has(key(d.kec, d.desa)));
  const sat = Object.entries(desaPlan).map(([k, n]) => ({ k, n, io: ioDesa[k] || 0, desa: katMap[k]?.desa || k.split('|')[1] }))
    .sort((a, b) => b.n - a.n).slice(0, 3);
  return {
    odp: planW.length, desa: planned.length, sfPlan: new Set(planW.map(p => p.username_sf)).size, kat,
    goodShare: planned.length ? (kat.prioritas + kat.strategis) / planned.length * 100 : null,
    ioPerPlanned: avg(planned), ioPerOther: avg(others),
    prioTotal: prio.length, prioUntouched: prioUntouched.length,
    prioUntouchedTop: prioUntouched.sort((a, b) => b.skorPotensi - a.skorPotensi).slice(0, 5),
    sat,
  };
}

function cmpBasis(D, cl, W) {
  if (!D.hasDesa) return null;
  // SF × desa → set periode IO; basis = ada IO di P-2, P-1 dan P (3 bulan berturut)
  const P = D.bulan, P1 = prevPeriod(P), P2 = D.p2;
  const per = n => { const s = String(n); return s.slice(0, 4) + '-' + s.slice(4, 6); };
  const sd = {};
  cl.rows.forEach(r => {
    const n = ymdNum(r[C.DATE_IO]); const code = (r[C.CODE_SF] || '').toString().trim(); const desa = N(r[C.DESA]);
    if (!n || !code || !desa || !(r[C.VAL_IO] || '').toString().trim()) return;
    const k = code + '|' + N(r[C.KEC]) + '|' + desa; (sd[k] = sd[k] || new Set()).add(per(n));
  });
  const basis = new Set(Object.keys(sd).filter(k => sd[k].has(P) && sd[k].has(P1) && sd[k].has(P2)));
  let io = 0, ioB = 0, finB = 0, psB = 0, finN = 0, psN = 0;
  cl.rows.forEach(r => {
    const n = ymdNum(r[C.DATE_IO]); if (!(n >= W[0] && n <= W[1]) || !(r[C.VAL_IO] || '').toString().trim()) return;
    io++;
    const k = (r[C.CODE_SF] || '').toString().trim() + '|' + N(r[C.KEC]) + '|' + N(r[C.DESA]);
    const isPS = !!ymdNum(r[C.DATE_PS]) || /PUT IN SERVICE/.test(N(r[C.VAL_PS])), isC = /CANCEL|BATAL/.test(N(r[C.STATUS_FNL]));
    if (basis.has(k)) { ioB++; if (isPS) { psB++; finB++; } else if (isC) finB++; }
    else { if (isPS) { psN++; finN++; } else if (isC) finN++; }
  });
  const sfB = new Set([...basis].map(k => k.split('|')[0]));
  return { pairs: basis.size, sfBasis: sfB.size, share: io ? ioB / io * 100 : null, convB: finB ? psB / finB * 100 : null, convN: finN ? psN / finN * 100 : null };
}

// ================================================================
//  3. ANALISA PERBANDINGAN (selisih + potensi PS)
// ================================================================
function cmpAnalyze(D) {
  const res = D.clusters.map(cl => {
    const cur = cmpMetrics(D, cl, D.W), prev = cmpMetrics(D, cl, D.W0);
    const djp = cmpDJP(D, cl, D.W), basis = cmpBasis(D, cl, D.W);
    const timePct = D.mode === 'bulanan' ? D.refDay / D.days * 100 : null;
    const tgtW = D.mode === 'bulanan' ? cl.tgtPS : cl.tgtPS / D.days * cur.dayLen;   // target jendela
    const ach = tgtW ? cur.ps / tgtW * 100 : null;
    const drr = cur.ps / cur.dayLen;
    const proj = D.mode === 'bulanan' ? cur.ps + drr * D.sisa : null;
    const projPct = proj != null && cl.tgtPS ? proj / cl.tgtPS * 100 : null;
    const need = D.mode === 'bulanan' && D.sisa > 0 && cl.tgtPS > cur.ps ? (cl.tgtPS - cur.ps) / D.sisa : 0;
    const chg = (a, b) => b ? Math.round((a - b) / b * 100) : null;
    return { name: cl.name, cl, cur, prev, djp, basis, tgtW, ach, timePct, drr, proj, projPct, need,
      dIO: chg(cur.io, prev.io), dPS: chg(cur.ps, prev.ps) };
  });
  const [a, b] = res;
  const r0 = x => x == null ? null : Math.round(x);

  // Pemenang per indikator (lebih tinggi lebih baik kecuali ditandai)
  const IND = [
    ['PS', x => x.cur.ps, fmtN],
    ['Target PS ' + (D.mode === 'bulanan' ? 'bulan' : 'minggu'), x => x.tgtW, fmtN, 'none'],
    ['Ach PS', x => x.ach, v => fmtPct(r0(v))],
    ...(D.mode === 'bulanan' ? [['Proyeksi akhir bulan', x => x.projPct, v => fmtPct(r0(v))]] : []),
    ['IO', x => x.cur.io, fmtN],
    [D.mode === 'bulanan' ? 'PS vs periode sama bln lalu' : 'PS vs minggu lalu', x => x.dPS, v => sgn(v)],
    ['Konversi order selesai', x => x.cur.conv, v => v == null ? '–' : fmt1(v) + '%'],
    ['Batal per 100 IO', x => x.cur.cancelPer100, v => v == null ? '–' : fmt1(v), 'low'],
    ['SF produktif', x => x.cur.sfProd, v => `${v}`],
    ['PS per SF yang PS', x => x.cur.psPerSF, fmt1],
    ['SF 0 PS', x => x.cur.sf0.length, v => `${v}`, 'low'],
    ['Desa DJP Prioritas/Strategis', x => x.djp.goodShare, v => v == null ? '–' : fmtPct(r0(v))],
    ...(D.hasDesa ? [['IO dari wilayah basis', x => x.basis?.share, v => v == null ? '–' : fmtPct(r0(v)), 'none']] : []),
  ];
  const table = IND.map(([label, f, fmt, dir]) => {
    const va = f(a), vb = f(b);
    let win = -1;
    if (dir !== 'none' && va != null && vb != null && Math.abs(va - vb) > 1e-9) win = dir === 'low' ? (va < vb ? 0 : 1) : (va > vb ? 0 : 1);
    return { label, vals: [fmt(va), fmt(vb)], win };
  });
  const wins = [0, 1].map(i => table.filter(t => t.win === i).length);

  // Potensi PS: bila cluster tertinggal menyamai yang unggul
  const gaps = [];
  const addGap = (label, f, potf) => {
    const va = f(a), vb = f(b); if (va == null || vb == null || va === vb) return;
    const lo = va < vb ? a : b, hi = lo === a ? b : a; const pot = Math.round(potf(lo, hi));
    if (pot > 0) gaps.push({ label, lo: lo.name, hi: hi.name, vlo: f(lo), vhi: f(hi), pot });
  };
  addGap('konversi order selesai', x => x.cur.conv, (lo, hi) => (hi.cur.conv - lo.cur.conv) / 100 * (lo.cur.cohPS + lo.cur.cohCancel));
  CG_KEYS.forEach(g => addGap(`batal karena ${cgl(g)}`, x => -(x.cur.coh ? x.cur.cg[g] / x.cur.coh : 0),
    (lo, hi) => (lo.cur.cg[g] / Math.max(lo.cur.coh, 1) - hi.cur.cg[g] / Math.max(hi.cur.coh, 1)) * lo.cur.coh * ((lo.cur.conv || 0) / 100)));
  gaps.sort((x, y) => y.pot - x.pot);

  // Kalimat
  const lead = (() => {
    if (!a || !b) return '';
    const kA = D.mode === 'bulanan' ? a.projPct : a.ach, kB = D.mode === 'bulanan' ? b.projPct : b.ach;
    const [hi, lo] = (kA ?? 0) >= (kB ?? 0) ? [a, b] : [b, a];
    const kHi = D.mode === 'bulanan' ? hi.projPct : hi.ach, kLo = D.mode === 'bulanan' ? lo.projPct : lo.ach;
    return `${hi.name} unggul ${D.mode === 'bulanan' ? 'proyeksi' : 'ach'} (${fmtPct(r0(kHi))} vs ${fmtPct(r0(kLo))}) dan menang di ${Math.max(...wins)} dari ${table.filter(t => t.win >= 0).length} indikator.`;
  })();
  const key = [];
  key.push(lead);
  gaps.slice(0, 2).forEach(g => key.push(`Jika ${g.lo} menyamai ${g.hi} di ${g.label}, tambahan ±${g.pot} PS dari order yang sudah ada.`));
  if (a && b) {
    const [hi, lo] = a.cur.psPerSF >= b.cur.psPerSF ? [a, b] : [b, a];
    if (lo.cur.psPerSF > 0 && hi.cur.psPerSF / lo.cur.psPerSF >= 1.15) key.push(`SF ${hi.name} ${fmt1(hi.cur.psPerSF / lo.cur.psPerSF)}× lebih produktif (${fmt1(hi.cur.psPerSF)} vs ${fmt1(lo.cur.psPerSF)} PS per SF); ${lo.name} punya ${lo.cur.sf0.length} SF belum PS.`);
  }
  res.forEach(x => {
    if (D.mode === 'bulanan' && x.need > x.drr) key.push(`${x.name} perlu ${fmt1(x.need)} PS/hari (sekarang ${fmt1(x.drr)}) untuk tembus target.`);
  });

  return { res, table, wins, gaps, key };
}

// ── Rekomendasi per cluster (aturan, dari angka) ──
function cmpActions(D, A, x) {
  const other = A.res.find(y => y !== x);
  const acts = [];
  const cg = x.cur.cg, coh = Math.max(x.cur.coh, 1);
  const topCg = CG_KEYS.filter(g => g !== 'lain').sort((g1, g2) => cg[g2] - cg[g1])[0];
  if (cg[topCg] >= 3) {
    const map = { lokasi: 'Validasi titik lokasi oleh TL sebelum input order', jaringan: 'Wajib cek port & jarak ODP sebelum input order', pelanggan: 'Perkuat penggalian kebutuhan & simulasi hemat vs kuota' };
    acts.push(`${map[topCg]} (${cg[topCg]} order batal, ${fmt1(cg[topCg] / coh * 100)} per 100 IO).`);
  }
  if (x.cur.sf0.length) acts.push(`Dampingi ${x.cur.sf0.length} SF yang belum PS: ${x.cur.sf0.slice(0, 3).map(s => shortName(s.nama)).join(', ')}${x.cur.sf0.length > 3 ? ', …' : ''}.`);
  if (x.djp.sat[0] && x.djp.sat[0].n >= 8 && x.djp.sat[0].io <= x.djp.sat[0].n / 4) acts.push(`Kurangi plan di ${x.djp.sat[0].desa} (${x.djp.sat[0].n} ODP, ${x.djp.sat[0].io} IO); alihkan ke desa Prioritas yang belum disentuh.`);
  if (x.djp.prioUntouched) acts.push(`Jadwalkan ${Math.min(x.djp.prioUntouched, 5)} dari ${x.djp.prioUntouched} desa Prioritas yang belum masuk plan bulan ini.`);
  if (other && x.cur.psPerSF < other.cur.psPerSF) acts.push(`Tiru pola SF ${other.name}: ${fmt1(other.cur.psPerSF)} PS per SF vs ${fmt1(x.cur.psPerSF)}.`);
  if (D.mode === 'bulanan' && x.need > x.drr) acts.push(`Naikkan run rate ke ${fmt1(x.need)} PS/hari (±${fmt1(x.need / Math.max(x.cur.sfPS, 1))} per SF aktif).`);
  return acts.slice(0, 4);
}

// ================================================================
//  4. SUSUN DECK (ringkas, satu pesan per slide)
// ================================================================
function cmpBuild(D, A) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  const judul = D.mode === 'bulanan' ? 'Review Bulanan' : 'Review Mingguan';
  pres.title = `${judul} WOK Lampung Barat – Way Kanan · ${D.label}`;
  pres.theme = { headFontFace: R_FONT, bodyFontFace: R_FONT };
  const dibuat = new Date().toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
  pres.defineSlideMaster({
    title: 'KONTEN', background: { color: RP.white },
    objects: [
      { text: { text: `WOK Lampung Barat – Way Kanan · ${judul} · ${D.label}`, options: { x: 0.5, y: 5.25, w: 7.5, h: 0.25, fontSize: 9, color: RP.muted, fontFace: R_FONT, margin: 0 } } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.28, w: 9, h: 0.55, fontSize: 22, bold: true, color: RP.ink, fontFace: R_FONT, margin: 0, valign: 'middle', align: 'left' }, text: '' } },
    ],
    slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.3, fontSize: 9, color: RP.muted, fontFace: R_FONT, align: 'right' },
  });
  const CL_COL = [RP.brand, RP.dark];                // LB = merah bata, WK = cokelat tua (konsisten)
  const [a, b] = A.res;
  const title = (s, t) => s.addText(t, { placeholder: 'title', align: 'left' });
  const sub = (s, t) => s.addText(t, { x: 0.5, y: 0.8, w: 9, h: 0.28, fontSize: 11, color: RP.muted, margin: 0, isTextBox: true });
  const card = (s, x, y, w, h, fill = RP.card) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 });
  const bullets = (s, items, x, y, w, h, size = 11) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 } })),
    { x, y, w, h, fontSize: size, color: RP.ink, valign: 'top', margin: 2, isTextBox: true, fit: 'shrink' });
  const insight = (s, items, x, y, w, h, head = 'Analisa perbandingan') => {
    card(s, x, y, w, h, RP.brandSoft);
    s.addText(head, { x: x + 0.18, y: y + 0.12, w: w - 0.36, h: 0.28, fontSize: 12, bold: true, color: RP.brand, margin: 0, isTextBox: true });
    bullets(s, (items.length ? items : ['Data belum cukup untuk dibandingkan.']).slice(0, 5), x + 0.12, y + 0.45, w - 0.24, h - 0.55);
  };
  const chartBase = (extra = {}) => Object.assign({
    catAxisLabelColor: RP.muted, valAxisLabelColor: RP.muted, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    catAxisLabelFontFace: R_FONT, valAxisLabelFontFace: R_FONT, dataLabelFontFace: R_FONT, titleFontFace: R_FONT,
    valGridLine: { color: 'EEE7E3', size: 0.5 }, catGridLine: { style: 'none' }, dataLabelFontSize: 9, dataLabelColor: RP.ink,
    titleFontSize: 11, titleColor: RP.ink, legendFontSize: 9, legendFontFace: R_FONT, legendColor: RP.muted,
  }, extra);
  const clTag = (s, x, y, i, text) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 0.16, h: 0.16, fill: { color: CL_COL[i] }, line: { color: CL_COL[i] }, rectRadius: 0.03 });
    s.addText(text, { x: x + 0.22, y: y - 0.06, w: 4, h: 0.28, fontSize: 12, bold: true, color: RP.ink, margin: 0, isTextBox: true });
  };
  const vs = (va, vb, f = fmtN) => `${f(va)} vs ${f(vb)}`;
  const r0 = v => v == null ? null : Math.round(v);
  const per100 = (x, g) => x.cur.coh ? x.cur.cg[g] / x.cur.coh * 100 : 0;

  // ── 1. Cover ──
  let s = pres.addSlide();
  s.background = { color: RP.dark };
  s.addShape(pres.shapes.OVAL, { x: 6.6, y: -1.2, w: 5, h: 5, fill: { color: RP.brand, transparency: 70 }, line: { color: RP.dark, transparency: 100 } });
  s.addText(judul, { x: 0.6, y: 1.3, w: 7, h: 0.5, fontSize: 18, color: 'F3C7BB', margin: 0, isTextBox: true });
  s.addText('Lampung Barat vs Way Kanan', { x: 0.6, y: 1.75, w: 8.8, h: 0.9, fontSize: 38, bold: true, color: RP.white, margin: 0, isTextBox: true, fit: 'shrink' });
  s.addText(`${D.label} · dibanding ${D.label0}`, { x: 0.6, y: 2.7, w: 8.8, h: 0.45, fontSize: 18, color: RP.white, margin: 0, isTextBox: true });
  s.addText(`WOK Lampung Barat – Way Kanan · dibuat ${dibuat}`, { x: 0.6, y: 4.6, w: 8, h: 0.35, fontSize: 12, color: 'D9B8AE', margin: 0, isTextBox: true });

  // ── 2. Scorecard perbandingan ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Siapa unggul, dan di mana');
  sub(s, `${D.label} · hijau = cluster yang lebih baik di indikator itu`);
  const hd = al => ({ bold: true, color: RP.muted, fontSize: 10, fill: { color: RP.card }, align: al || 'right' });
  const trows = [[{ text: 'Indikator', options: hd('left') }, { text: a.name, options: hd() }, { text: b.name, options: hd() }]];
  A.table.forEach(t => trows.push([{ text: t.label }, ...t.vals.map((v, i) => ({ text: v, options: { align: 'right', bold: t.win === i,
    color: t.win === i ? RP.green : RP.ink, fill: { color: t.win === i ? RP.greenSoft : RP.white } } }))]));
  const rowH = Math.min(0.29, 3.9 / trows.length);
  s.addTable(trows, { x: 0.5, y: 1.2, w: 5.3, colW: [2.7, 1.3, 1.3], fontSize: 10.5, fontFace: R_FONT, color: RP.ink, rowH,
    border: { type: 'solid', color: RP.line, pt: 0.5 }, margin: 0.04 });
  insight(s, A.key, 6.0, 1.2, 3.5, 3.9, 'Kesimpulan');

  // ── 3. Pencapaian & tren ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  if (D.mode === 'bulanan') {
    title(s, 'Laju PS menuju target');
    sub(s, 'PS kumulatif harian dibanding garis target (target bulan ÷ jumlah hari)');
    const L = []; for (let d = 1; d <= D.refDay; d++) L.push(String(d));
    const cum = x => { let c = 0; return L.map(d => { c += x.cur.daily[+(D.bulan.replace('-', '') + d.padStart(2, '0'))] || 0; return c; }); };
    const pace = x => L.map(d => Math.round(x.cl.tgtPS / D.days * +d));
    s.addChart(pres.charts.LINE, [
      { name: `${a.name}`, labels: L, values: cum(a) }, { name: `Target ${shortCl(a.name)}`, labels: L, values: pace(a) },
      { name: `${b.name}`, labels: L, values: cum(b) }, { name: `Target ${shortCl(b.name)}`, labels: L, values: pace(b) },
    ], chartBase({ x: 0.5, y: 1.15, w: 5.6, h: 3.95, chartColors: [CL_COL[0], 'E8A08F', CL_COL[1], 'B9A49C'], lineSize: 2,
      lineDataSymbol: 'none', showLegend: true, legendPos: 'b', showTitle: true, title: 'PS kumulatif vs pace target' }));
    insight(s, A.res.map(x => `${x.name}: ${fmtN(x.cur.ps)} PS = ${fmtPct(r0(x.ach))} target saat waktu berjalan ${fmtPct(r0(x.timePct))}; proyeksi ${fmtPct(r0(x.projPct))}` +
      (x.need > x.drr ? `, perlu ${fmt1(x.need)} PS/hari.` : '.')).concat(
      A.res.map(x => `${x.name} vs periode sama bulan lalu: PS ${sgn(x.dPS)}, IO ${sgn(x.dIO)}.`)), 6.3, 1.2, 3.2, 3.9);
  } else {
    title(s, 'Minggu ini vs minggu lalu');
    sub(s, `${D.label} dibanding ${D.label0}`);
    const cats = ['IO', 'PS'];
    s.addChart(pres.charts.BAR, [
      { name: `${shortCl(a.name)} minggu lalu`, labels: cats, values: [a.prev.io, a.prev.ps] },
      { name: `${shortCl(a.name)} minggu ini`, labels: cats, values: [a.cur.io, a.cur.ps] },
      { name: `${shortCl(b.name)} minggu lalu`, labels: cats, values: [b.prev.io, b.prev.ps] },
      { name: `${shortCl(b.name)} minggu ini`, labels: cats, values: [b.cur.io, b.cur.ps] },
    ], chartBase({ x: 0.5, y: 1.15, w: 5.6, h: 3.95, barDir: 'col', chartColors: ['F0C4B8', CL_COL[0], 'C9B8B1', CL_COL[1]], showValue: true,
      dataLabelPosition: 'outEnd', showLegend: true, legendPos: 'b', showTitle: true, title: 'IO dan PS per minggu', valAxisHidden: true, valGridLine: { style: 'none' } }));
    insight(s, A.res.map(x => `${x.name}: IO ${vs(x.cur.io, x.prev.io)} (${sgn(x.dIO)}), PS ${vs(x.cur.ps, x.prev.ps)} (${sgn(x.dPS)}); ach mingguan ${fmtPct(r0(x.ach))}.`)
      .concat([(() => { const up = [...A.res].sort((p, q) => (q.dPS ?? -999) - (p.dPS ?? -999)); return `Momentum lebih baik di ${up[0].name} (${sgn(up[0].dPS)} vs ${sgn(up[1].dPS)}).`; })()]), 6.3, 1.2, 3.2, 3.9);
  }

  // ── 4. Kebocoran funnel ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Di mana order bocor');
  sub(s, 'Order batal per 100 IO, menurut penyebab (order dengan IO di periode ini)');
  s.addChart(pres.charts.BAR, CG_KEYS.map(g => ({ name: CG_LABEL[g], labels: [a.name, b.name], values: [Math.round(per100(a, g) * 10) / 10, Math.round(per100(b, g) * 10) / 10] })),
    chartBase({ x: 0.5, y: 1.15, w: 5.6, h: 3.95, barDir: 'bar', barGrouping: 'stacked', chartColors: ['B45309', '475569', RP.brand, 'B9B2AE'],
      showValue: true, dataLabelPosition: 'ctr', dataLabelFormatCode: '0.0;;', dataLabelColor: RP.white, showLegend: true, legendPos: 'b',
      showTitle: true, title: 'Batal per 100 IO', valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin' }));
  const leak = [];
  leak.push(`Konversi order selesai: ${a.name} ${fmt1(a.cur.conv)}% vs ${b.name} ${fmt1(b.cur.conv)}%.`);
  CG_KEYS.filter(g => g !== 'lain').forEach(g => {
    const va = per100(a, g), vb = per100(b, g); if (Math.max(va, vb) < 1) return;
    const [hi, lo, vh, vl] = va >= vb ? [a, b, va, vb] : [b, a, vb, va];
    if (vh - vl >= 1) leak.push(`${CG_LABEL[g]}: ${hi.name} ${fmt1(vh)} vs ${lo.name} ${fmt1(vl)} per 100 IO (${hi.cur.cg[g]} order).`);
  });
  const topOf = x => CG_KEYS.filter(g => g !== 'lain').sort((g1, g2) => x.cur.cg[g2] - x.cur.cg[g1])[0];
  const ta = topOf(a), tb = topOf(b);
  if (a.cur.cohCancel && b.cur.cohCancel) leak.push(ta === tb
    ? `Penyebab terbesar sama di kedua cluster: ${cgl(ta)} (${fmt1(per100(a, ta))} dan ${fmt1(per100(b, tb))} per 100 IO), jadi bisa ditangani bersama.`
    : `Penyebab terbesar berbeda: ${a.name} ${cgl(ta)}, ${b.name} ${cgl(tb)}.`);
  A.gaps.filter(g => /batal/.test(g.label)).slice(0, 1).forEach(g => leak.push(`Menutup selisih terbesar = ±${g.pot} PS tambahan untuk ${g.lo}.`));
  insight(s, leak, 6.3, 1.2, 3.2, 3.9);

  // ── 5. Channel & produk ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Channel dan produk');
  sub(s, 'Konversi order selesai (%) per produk; porsi order lewat SF (Agency)');
  const prods = [...new Set([...Object.keys(a.cur.pr), ...Object.keys(b.cur.pr)])]
    .filter(p => (a.cur.pr[p]?.io || 0) + (b.cur.pr[p]?.io || 0) >= 5).sort((p, q) => ((b.cur.pr[q]?.io || 0) + (a.cur.pr[q]?.io || 0)) - ((b.cur.pr[p]?.io || 0) + (a.cur.pr[p]?.io || 0))).slice(0, 5);
  const cv = (x, p) => { const o = x.cur.pr[p]; return o && o.fin ? Math.round(o.ps / o.fin * 100) : 0; };
  if (prods.length) s.addChart(pres.charts.BAR, A.res.map((x, i) => ({ name: x.name, labels: prods, values: prods.map(p => cv(x, p)) })),
    chartBase({ x: 0.5, y: 1.15, w: 4.0, h: 3.95, barDir: 'bar', chartColors: CL_COL, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0"%"',
      showLegend: true, legendPos: 'b', showTitle: true, title: 'Konversi per produk', valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin', valAxisMaxVal: 100 }));
  A.res.forEach((x, i) => {
    const ag = x.cur.ch['Agency'] || { io: 0, ps: 0, fin: 0 }, dg = x.cur.ch['Digital'] || { io: 0, ps: 0, fin: 0 };
    const y = 1.2 + i * 1.3;
    card(s, 4.7, y, 1.45, 1.15);
    clTag(s, 4.82, y + 0.15, i, shortCl(x.name));
    s.addText(fmtPct(pct(ag.io, ag.io + dg.io)), { x: 4.82, y: y + 0.4, w: 1.2, h: 0.4, fontSize: 20, bold: true, color: RP.ink, margin: 0, isTextBox: true });
    s.addText('IO lewat SF', { x: 4.82, y: y + 0.8, w: 1.25, h: 0.25, fontSize: 9, color: RP.muted, margin: 0, isTextBox: true });
  });
  const chIns = A.res.map(x => { const ag = x.cur.ch['Agency'] || {}, dg = x.cur.ch['Digital'] || {};
    return `${x.name}: order lewat SF jadi PS ${ag.fin ? Math.round(ag.ps / ag.fin * 100) + '%' : '–'}` + (dg.fin ? `, order digital ${Math.round(dg.ps / dg.fin * 100)}% (${dg.io} order).` : '; tidak ada order digital yang selesai.'); });
  { const sh = x => { const ag = x.cur.ch['Agency']?.io || 0, dg = x.cur.ch['Digital']?.io || 0; return ag + dg ? ag / (ag + dg) * 100 : null; };
    const sa = sh(a), sb = sh(b); if (sa != null && sb != null && Math.abs(sa - sb) >= 3) { const [hi, lo] = sa > sb ? [a, b] : [b, a];
      chIns.push(`${hi.name} lebih bergantung pada SF (${Math.round(Math.max(sa, sb))}% vs ${Math.round(Math.min(sa, sb))}% IO lewat SF).`); } }
  prods.slice(0, 3).forEach(p => { const va = cv(a, p), vb = cv(b, p); if (Math.abs(va - vb) >= 5) chIns.push(`${p}: ${va > vb ? a.name : b.name} lebih baik (${Math.max(va, vb)}% vs ${Math.min(va, vb)}%).`); });
  insight(s, chIns, 6.3, 1.2, 3.2, 3.9);

  // ── 6. Produktivitas SF ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Produktivitas SF');
  const pm = a.cur.prodMin;
  sub(s, `Sebaran SF menurut PS di periode ini · produktif = ≥${pm} PS (setara ≥10 PS/bulan)`);
  const bk = [['0 PS', s0 => s0.ps === 0], [`1–${Math.max(pm - 1, 1)} PS`, s0 => s0.ps >= 1 && s0.ps < pm], [`≥${pm} PS`, s0 => s0.ps >= pm]];
  const rosterOnly = x => x.cur.sfList.filter(s0 => x.cl.sf.some(r => r.code === s0.code) || s0.ps > 0);
  s.addChart(pres.charts.BAR, A.res.map(x => ({ name: x.name, labels: bk.map(z => z[0]), values: bk.map(z => rosterOnly(x).filter(z[1]).length) })),
    chartBase({ x: 0.5, y: 1.15, w: 4.0, h: 3.95, barDir: 'col', chartColors: CL_COL, showValue: true, dataLabelPosition: 'outEnd', showLegend: true, legendPos: 'b',
      showTitle: true, title: 'Jumlah SF', valAxisHidden: true, valGridLine: { style: 'none' } }));
  A.res.forEach((x, i) => {
    const y = 1.2 + i * 1.3; card(s, 4.7, y, 1.45, 1.15); clTag(s, 4.82, y + 0.15, i, shortCl(x.name));
    s.addText(fmt1(x.cur.psPerSF), { x: 4.82, y: y + 0.4, w: 1.2, h: 0.4, fontSize: 20, bold: true, color: RP.ink, margin: 0, isTextBox: true });
    s.addText('PS per SF yang PS', { x: 4.82, y: y + 0.8, w: 1.3, h: 0.25, fontSize: 9, color: RP.muted, margin: 0, isTextBox: true });
  });
  const sfIns = A.res.map(x => `${x.name}: ${x.cur.sfProd} SF produktif, ${x.cur.sfPS} dari ${x.cur.sfTotal} SF sudah PS; 20% SF teratas = ${fmtPct(r0(x.cur.top20))} PS.`);
  { const [hi, lo] = a.cur.psPerSF >= b.cur.psPerSF ? [a, b] : [b, a];
    if (lo.cur.psPerSF > 0 && hi.cur.psPerSF / lo.cur.psPerSF >= 1.15) sfIns.push(`SF ${hi.name} rata-rata ${fmt1(hi.cur.psPerSF / lo.cur.psPerSF)}× lebih produktif (${fmt1(hi.cur.psPerSF)} vs ${fmt1(lo.cur.psPerSF)} PS per SF).`); }
  A.res.forEach(x => { if ((x.cur.top20 || 0) >= 60) sfIns.push(`${x.name} bergantung pada sedikit SF (${fmtPct(r0(x.cur.top20))} PS dari 20% SF).`); });
  insight(s, sfIns, 6.3, 1.2, 3.2, 3.9);

  // ── 7. Efektivitas DJP ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Efektivitas DJP');
  sub(s, 'Plan DJP di periode ini dibanding IO di desa yang diplan dan tidak diplan');
  const djpRows = [
    ['ODP diplan', x => fmtN(x.djp.odp)], ['Desa diplan', x => fmtN(x.djp.desa)], ['SF punya plan', x => `${x.djp.sfPlan}/${x.cur.sfTotal}`],
    ['Desa Prioritas/Strategis', x => fmtPct(r0(x.djp.goodShare))],
    ...(D.hasDesa ? [['IO per desa diplan', x => x.djp.ioPerPlanned == null ? '–' : fmt1(x.djp.ioPerPlanned)], ['IO per desa lain', x => x.djp.ioPerOther == null ? '–' : fmt1(x.djp.ioPerOther)]] : []),
    ['Prioritas belum diplan (bulan ini)', x => `${x.djp.prioUntouched}/${x.djp.prioTotal}`],
  ];
  const dr = [[{ text: '', options: hd('left') }, { text: a.name, options: hd() }, { text: b.name, options: hd() }]]
    .concat(djpRows.map(([l, f]) => [{ text: l }, { text: f(a), options: { align: 'right' } }, { text: f(b), options: { align: 'right' } }]));
  s.addTable(dr, { x: 0.5, y: 1.2, w: 5.6, colW: [2.8, 1.4, 1.4], fontSize: 11, fontFace: R_FONT, color: RP.ink, rowH: 0.32, border: { type: 'solid', color: RP.line, pt: 0.5 }, margin: 0.05 });
  const dj = [];
  A.res.forEach(x => {
    if (!x.djp.odp) { dj.push(`${x.name}: belum ada plan DJP tersimpan di periode ini.`); return; }
    if (x.djp.ioPerPlanned != null && x.djp.ioPerOther != null) dj.push(`${x.name}: desa diplan ${fmt1(x.djp.ioPerPlanned)} IO/desa vs desa lain ${fmt1(x.djp.ioPerOther)}.`);
    const t = x.djp.sat[0]; if (t && t.n >= 8) dj.push(`${x.name}: plan terpadat di ${t.desa} (${t.n} ODP, ${t.io} IO).`);
  });
  if (a.djp.goodShare != null && b.djp.goodShare != null) dj.push(`Ketepatan sasaran: ${a.name} ${fmtPct(r0(a.djp.goodShare))} vs ${b.name} ${fmtPct(r0(b.djp.goodShare))} desa Prioritas/Strategis.`);
  dj.push('Realisasi kunjungan belum terbaca; dampak dihitung dari plan.');
  insight(s, dj, 6.3, 1.2, 3.2, 3.9);

  // ── 8. Wilayah basis (SF × desa 3 bulan berturut) ──
  if (D.hasDesa) {
    s = pres.addSlide({ masterName: 'KONTEN' });
    title(s, 'Wilayah basis SF');
    sub(s, `Desa yang digarap SF yang sama 3 bulan berturut (${periodLabel(D.p2)} – ${periodLabel(D.bulan)})`);
    A.res.forEach((x, i) => {
      const bs = x.basis || {}, x0 = 0.5 + i * 2.95;
      card(s, x0, 1.2, 2.8, 3.9);
      clTag(s, x0 + 0.2, 1.4, i, x.name);
      const st = [[fmtPct(r0(bs.share)), 'IO periode ini dari wilayah basis'], [`${fmtN(bs.pairs)}`, `pasangan SF–desa basis (${bs.sfBasis || 0} SF)`],
        [`${bs.convB == null ? '–' : fmt1(bs.convB) + '%'}`, `konversi basis vs ${bs.convN == null ? '–' : fmt1(bs.convN) + '%'} lainnya`]];
      st.forEach(([v, l], j) => {
        s.addText(v, { x: x0 + 0.2, y: 1.8 + j * 1.05, w: 2.4, h: 0.5, fontSize: 24, bold: true, color: CL_COL[i], margin: 0, isTextBox: true });
        s.addText(l, { x: x0 + 0.2, y: 2.3 + j * 1.05, w: 2.45, h: 0.4, fontSize: 10, color: RP.muted, margin: 0, isTextBox: true, fit: 'shrink' });
      });
    });
    const bi = [];
    if (a.basis && b.basis) {
      const [hi, lo] = (a.basis.share ?? 0) >= (b.basis.share ?? 0) ? [a, b] : [b, a];
      bi.push(`${hi.name} lebih bertumpu pada wilayah basis (${fmtPct(r0(hi.basis.share))} vs ${fmtPct(r0(lo.basis.share))} IO).`);
      A.res.forEach(x => { if (x.basis.convB != null && x.basis.convN != null && x.basis.convB - x.basis.convN >= 3) bi.push(`${x.name}: order di wilayah basis lebih sering jadi PS (+${Math.round(x.basis.convB - x.basis.convN)} poin).`); });
      bi.push('Wilayah basis = sumber referal yang belum diformalkan: beri kode referal tiap pelanggan PS dan kunjungi ulang 2–4 minggu setelah pasang.');
    }
    insight(s, bi, 6.5, 1.2, 3.0, 3.9);
  }

  // ── 9. Aksi per cluster ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, D.mode === 'bulanan' ? 'Fokus sisa bulan' : 'Fokus minggu depan');
  sub(s, 'Aksi diturunkan dari angka di slide sebelumnya · target desa = Prioritas belum diplan, skor potensi tertinggi');
  A.res.forEach((x, i) => {
    const x0 = 0.5 + i * 4.6; card(s, x0, 1.2, 4.4, 3.9);
    clTag(s, x0 + 0.2, 1.4, i, x.name);
    bullets(s, cmpActions(D, A, x), x0 + 0.15, 1.7, 4.1, 2.0, 11);
    const td = x.djp.prioUntouchedTop;
    s.addText('Target desa', { x: x0 + 0.2, y: 3.75, w: 4, h: 0.25, fontSize: 10, bold: true, color: RP.muted, margin: 0, isTextBox: true });
    s.addText(td.length ? td.map(d => `${d.desa} (${d.kec})`).join(' · ') : 'Semua desa Prioritas sudah masuk plan bulan ini.',
      { x: x0 + 0.2, y: 4.0, w: 4.0, h: 1.0, fontSize: 10.5, color: RP.ink, valign: 'top', margin: 0, isTextBox: true, fit: 'shrink' });
  });

  return pres;
}

// ================================================================
//  5. ENTRY
// ================================================================
async function exportComparePPT(app, mode) {
  await loadPptLib();
  const D = await cmpCollect(app, mode);
  const A = cmpAnalyze(D);
  const pres = cmpBuild(D, A);
  const tgl = typeof localYMD === 'function' ? localYMD(new Date()) : new Date().toISOString().slice(0, 10);
  const name = `Review-${mode === 'bulanan' ? 'Bulanan' : 'Mingguan'}-LB-vs-WK-${tgl}.pptx`;
  await pres.writeFile({ fileName: name });
  return name;
}
