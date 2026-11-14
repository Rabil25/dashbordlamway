// ================================================================
//  Export PPT — Review Mingguan (dibuat di browser, pptxgenjs)
//  - Library dimuat hanya saat tombol diklik (assets/vendor/pptxgen.bundle.js)
//  - Semua angka diambil dari getter dashboard yang sama dengan tampilan web
//  - Analisa = aturan statistik sederhana yang bisa ditelusuri (bukan model ML):
//    pace vs waktu berjalan, DRR yang dibutuhkan, tren 7 hari, konsentrasi
//    (Pareto), pergerakan tier, cakupan desa per kategori.
// ================================================================

let _pptLibP = null;
function loadPptLib() {
  if (window.PptxGenJS) return Promise.resolve();
  if (!_pptLibP) {
    _pptLibP = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'assets/vendor/pptxgen.bundle.js';
      s.onload = res;
      s.onerror = () => { _pptLibP = null; rej(new Error('Gagal memuat pembuat PPT')); };
      document.head.appendChild(s);
    });
  }
  return _pptLibP;
}

// ── Palet (hex tanpa #) ──
const RP = {
  dark: '2B1410', brand: 'C93C20', brandSoft: 'FBE9E4', ink: '1F1A17', muted: '6B5E57',
  line: 'E7DED9', card: 'F7F3F1', white: 'FFFFFF', green: '15803D', greenSoft: 'DCFCE7',
  amber: 'B45309', amberSoft: 'FEF3C7',
  prioritas: 'EF4444', strategis: 'F97316', cadangan: 'EAB308', skip: '9CA3AF',
};
const R_FONT = 'Calibri';

// ── Util angka ──
const fmtN  = n => (n == null || isNaN(n)) ? '–' : Math.round(n).toLocaleString('id-ID');
const fmt1  = n => (n == null || isNaN(n)) ? '–' : (Math.round(n * 10) / 10).toLocaleString('id-ID');
const pct   = (a, b) => b > 0 ? Math.round(a / b * 100) : null;
const fmtPct = p => p == null ? '–' : p + '%';
const sgn   = p => p == null ? '–' : (p > 0 ? '+' : '') + p + '%';
const sum   = arr => arr.reduce((s, v) => s + (v || 0), 0);
// "M Lukman Hakim" → "M Lukman"; "LEO WALDI" → "LEO"
const shortName = n => { const w = String(n || '').replace(/\[.*?\]\s*/g, '').trim().split(/\s+/); return (w[0] && w[0].length <= 2 && w[1]) ? w[0] + ' ' + w[1] : (w[0] || ''); };

// ================================================================
//  1. KUMPULKAN DATA dari dashboard
// ================================================================
async function reportCollect(app) {
  const bulan = app.cfg.bulan;
  const days = daysInPeriod(bulan);
  // Hari acuan = akhir rentang tanggal yang dipakai KPI dashboard (bulan berjalan: hari ini)
  const hari = Math.min(parseInt(app.kpi.hariPS) || app.hariBerjalan, app.hariBerjalan, days);
  const sisa = Math.max(days - hari, 0);
  const kabs = (CLUSTER[app.cluster]?.kab || []).map(k => k.toUpperCase());

  // PS & IO harian bulan ini dan bulan lalu (hari ke-n)
  const psDay = Array(days + 1).fill(0), ioDay = Array(days + 1).fill(0);
  app.filteredRows.forEach(r => {
    if ((r[C.BLN_PS] || '').toString() === bulan && (r[C.VAL_PS] || '').toString().trim() !== '') {
      const d = parseInt(r[C.TGL_PS] || 0); if (d >= 1 && d <= days) psDay[d]++;
    }
    if ((r[C.BLN_IO] || '').toString() === bulan && (r[C.VAL_IO] || '').toString().trim() !== '') {
      const d = parseInt(r[C.TGL_IO] || 0); if (d >= 1 && d <= days) ioDay[d]++;
    }
  });

  // Mix bulan lalu per channel / produk (untuk perubahan porsi)
  const prev = app.prevBulan.toString();
  const prevMix = { ch: {}, pr: {} };
  app.filteredPrevRows.forEach(r => {
    if ((r[C.BLN_PS] || '').toString() !== prev) return;
    const ch = r[C.CHANNEL2] || 'Lainnya', pr = r[C.PRODUK_KAT] || 'Lainnya';
    prevMix.ch[ch] = (prevMix.ch[ch] || 0) + 1;
    prevMix.pr[pr] = (prevMix.pr[pr] || 0) + 1;
  });

  // Kendala (order IO bulan ini)
  const kendala = {};
  app.filteredRows.forEach(r => {
    if ((r[C.BLN_IO] || '').toString() !== bulan) return;
    const k = (r[C.KENDALA_CAT] || '').toString().trim();
    if (k) kendala[k] = (kendala[k] || 0) + 1;
  });

  // Desa (rumus Prioritas, per cluster)
  if (!app.prioritasOdpRaw.length) { try { await app.loadPrioritas(); } catch (e) {} }
  const desa = app.prioritasOdpRaw.length ? app.klasifikasiDesa('', '') : [];

  // Plan DJP bulan ini (Vercel, read-only)
  let plan = [];
  try {
    const r = await fetch(`/api/data?type=djp_plan&bulan=${encodeURIComponent(bulan)}`,
      { headers: { Authorization: 'Bearer ' + app.sessionToken } });
    const d = await r.json();
    if (d.ok) plan = (d.data || []).filter(p => kabs.includes((p.kab || '').toUpperCase()));
  } catch (e) {}

  // Tanggal acuan "minggu ini": hari terakhir data (bulan berjalan → hari ini)
  const refDay = hari;
  const refTgl = bulan + '-' + String(refDay).padStart(2, '0');
  const weekStart = bulan + '-' + String(Math.max(refDay - 6, 1)).padStart(2, '0');

  return {
    cluster: app.cluster, bulan, periode: periodLabel(bulan), days, hari, sisa, kabs,
    kpi: app.kpi, kab: app.kabBreakdown, outlook: app.outlookData,
    psDay, ioDay, prevMix,
    channel: app.deepdiveChannel, produk: app.deepdiveProduk, speed: app.deepdiveSpeed,
    sf: app.buildSFTable(), tier: app.sfperfTierDist,
    kendala: Object.entries(kendala).map(([name, n]) => ({ name, n })).sort((a, b) => b.n - a.n),
    desa, plan, refTgl, weekStart,
  };
}

// ================================================================
//  2. ANALISA OTOMATIS (aturan statistik yang bisa ditelusuri)
// ================================================================
function reportAnalyze(D) {
  const A = {};
  const k = D.kpi;
  const timePct = Math.round(D.hari / D.days * 100);
  const achPS = pct(k.ps, k.tgtPS);
  const drr = D.hari > 0 ? k.ps / D.hari : 0;
  const proj = Math.round(k.ps + drr * D.sisa);
  const projPct = pct(proj, k.tgtPS);
  const needDrr = D.sisa > 0 && k.tgtPS > k.ps ? (k.tgtPS - k.ps) / D.sisa : 0;
  const reIO = k.io > 0 ? k.re / k.io * 100 : 0, psRE = k.re > 0 ? k.ps / k.re * 100 : 0;
  A.core = { timePct, achPS, drr, proj, projPct, needDrr };

  // ── KPI ──
  A.kpi = [];
  if (k.tgtPS > 0) {
    const gap = achPS - timePct;
    A.kpi.push(`PS ${fmtN(k.ps)} dari target ${fmtN(k.tgtPS)} (${achPS}%) di hari ke-${D.hari} dari ${D.days} (${timePct}% waktu) → ` +
      (gap >= 0 ? `unggul ${gap} poin dari pace.` : `tertinggal ${-gap} poin dari pace.`));
    if (D.sisa > 0) A.kpi.push(`Proyeksi akhir bulan dengan DRR ${fmt1(drr)}/hari: ${fmtN(proj)} PS (${projPct}% target).`);
    else A.kpi.push(`Periode selesai: PS akhir ${fmtN(k.ps)} (${achPS}% target), DRR rata-rata ${fmt1(drr)}/hari.`);
    if (needDrr > 0) A.kpi.push(`Untuk tembus target perlu ${fmt1(needDrr)} PS/hari di ${D.sisa} hari tersisa ` +
      (drr > 0 ? `(${needDrr > drr ? 'naik ' + Math.round((needDrr / drr - 1) * 100) + '%' : 'sudah tercapai dengan DRR sekarang'}).` : '.'));
  }
  if (k.momPSvsLast != null) A.kpi.push(`Dibanding periode sama bulan lalu: PS ${sgn(k.momPSvsLast)}, IO ${sgn(k.momIOvsLast)}.`);
  if (k.io > 0) {
    const weak = reIO < psRE ? `IO→RE (${fmt1(reIO)}%)` : `RE→PS (${fmt1(psRE)}%)`;
    A.kpi.push(`Konversi terlemah di tahap ${weak}; fokus perbaikan di sini memberi dampak terbesar.`);
  }

  // ── Tren 7 hari ──
  const d = D.psDay, h = D.hari;
  const w0 = sum(d.slice(Math.max(h - 6, 1), h + 1)), w1 = sum(d.slice(Math.max(h - 13, 1), Math.max(h - 6, 1)));
  const dailyTgt = k.tgtPS > 0 ? k.tgtPS / D.days : 0;
  let best = 0, bestDay = 0, zero = 0, hit = 0;
  for (let i = 1; i <= h; i++) { if (d[i] > best) { best = d[i]; bestDay = i; } if (d[i] === 0) zero++; if (dailyTgt && d[i] >= dailyTgt) hit++; }
  // kemiringan tren (regresi linear PS/hari)
  let sx = 0, sy = 0, sxy = 0, sxx = 0, n = 0;
  for (let i = 1; i <= h; i++) { sx += i; sy += d[i]; sxy += i * d[i]; sxx += i * i; n++; }
  const slope = n > 1 ? (n * sxy - sx * sy) / (n * sxx - sx * sx) : 0;
  A.trend = [];
  A.trend.push(`7 hari terakhir ${fmtN(w0)} PS vs 7 hari sebelumnya ${fmtN(w1)} (${sgn(w1 > 0 ? Math.round((w0 - w1) / w1 * 100) : null)}).`);
  if (dailyTgt) A.trend.push(`Target harian ${fmt1(dailyTgt)} PS tercapai ${hit} dari ${h} hari; ${zero} hari tanpa PS.`);
  if (bestDay) A.trend.push(`Hari terbaik tanggal ${bestDay} (${best} PS). Arah tren: ${Math.abs(slope) < 0.05 ? 'datar' : slope > 0 ? 'naik ' + fmt1(slope) + ' PS/hari' : 'turun ' + fmt1(-slope) + ' PS/hari'}.`);
  const kabs = D.kab.filter(x => x.tPS > 0);
  if (kabs.length) {
    const low = [...kabs].sort((a, b) => (a.momPS ?? 0) - (b.momPS ?? 0))[0];
    const top = [...D.kab].sort((a, b) => b.ps - a.ps)[0];
    const kabTot = sum(D.kab.map(x => x.ps));
    A.trend.push(`${top.kab} penyumbang terbesar (${pct(top.ps, kabTot) ?? 0}% PS); ${low.kab} paling tertinggal (ach ${fmtPct(low.momPS)}).`);
  }
  A.w = { w0, w1, dailyTgt };

  // ── Deepdive ──
  A.deep = [];
  const totPS = sum(D.channel.map(x => x.ps)) || 0, prevTot = sum(Object.values(D.prevMix.ch));
  if (D.channel[0] && totPS) A.deep.push(`Channel ${D.channel[0].name} menyumbang ${pct(D.channel[0].ps, totPS)}% PS.`);
  const shift = (list, prevMap, prevTotal, label) => {
    if (!prevTotal || !totPS) return;
    const s = list.map(x => ({ name: x.name, d: Math.round((x.ps / totPS - (prevMap[x.name] || 0) / prevTotal) * 100) }))
      .sort((a, b) => Math.abs(b.d) - Math.abs(a.d))[0];
    if (s && Math.abs(s.d) >= 3) A.deep.push(`Porsi ${label} ${s.name} ${s.d > 0 ? 'naik' : 'turun'} ${Math.abs(s.d)} poin dibanding bulan lalu.`);
  };
  shift(D.channel, D.prevMix.ch, prevTot, 'channel');
  shift(D.produk, D.prevMix.pr, sum(Object.values(D.prevMix.pr)), 'produk');
  if (D.produk[0] && totPS) A.deep.push(`Produk dominan: ${D.produk[0].name} (${pct(D.produk[0].ps, totPS)}% PS).`);
  if (D.speed[0] && totPS) A.deep.push(`Speed terlaris: ${D.speed[0].name}${/^\d+$/.test(String(D.speed[0].name).trim()) ? ' Mbps' : ''} (${pct(D.speed[0].ps, totPS)}% PS).`);
  const cvt = D.channel.filter(x => x.re >= 5).map(x => ({ ...x, c: x.ps / x.re * 100 })).sort((a, b) => a.c - b.c);
  if (cvt.length > 1) A.deep.push(`Konversi RE→PS terendah di channel ${cvt[0].name} (${fmt1(cvt[0].c)}%).`);

  // ── SF ──
  const sf = D.sf, aktif = sf.filter(s => s.ps > 0);
  const sorted = [...sf].sort((a, b) => b.ps - a.ps);
  const top20 = Math.max(1, Math.round(sf.length * 0.2));
  const topShare = pct(sum(sorted.slice(0, top20).map(s => s.ps)), sum(sf.map(s => s.ps)));
  const order = TIER_CONFIG.map(t => t.name);
  const up = sf.filter(s => order.indexOf(s.tierNow) < order.indexOf(s.tier)).length;
  const down = sf.filter(s => order.indexOf(s.tierNow) > order.indexOf(s.tier)).length;
  const ps0 = sf.filter(s => s.ps === 0);
  A.sf = [];
  A.sf.push(`${aktif.length} dari ${sf.length} SF sudah PS (${pct(aktif.length, sf.length) ?? 0}%); ${ps0.length} SF masih 0 PS.`);
  if (topShare != null) A.sf.push(`${top20} SF teratas (20%) menyumbang ${topShare}% PS — ${topShare >= 60 ? 'sangat bergantung pada sedikit SF.' : 'sebaran cukup merata.'}`);
  A.sf.push(`Pergerakan tier vs bulan lalu: ${up} SF naik, ${down} SF turun.`);
  const behind = sf.filter(s => s.target > 0).map(s => ({ ...s, gap: s.ach - Math.round(D.hari / D.days * 100) }))
    .sort((a, b) => a.gap - b.gap).slice(0, 3);
  if (behind.length) A.sf.push(`Prioritas coaching: ${behind.map(s => shortName(s.nama) + ' (' + s.ach + '%)').join(', ')}.`);
  A.sfLists = { top: sorted.slice(0, 5), bottom: sf.filter(s => s.target > 0).sort((a, b) => a.ach - b.ach).slice(0, 5), ps0 };

  // ── DJP minggu ini ──
  const week = D.plan.filter(p => p.tgl >= D.weekStart && p.tgl <= D.refTgl);
  const sfPlan = {};
  week.forEach(p => { (sfPlan[p.nama_sf || p.username_sf] = sfPlan[p.nama_sf || p.username_sf] || new Set()).add(p.odp); });
  const katByDesa = {};
  D.desa.forEach(x => { katByDesa[(x.kec || '').toUpperCase() + '||' + (x.desa || '').toUpperCase()] = x; });
  const keyOf = p => (p.kec || '').toUpperCase() + '||' + (p.desa || '').toUpperCase();
  const weekDesa = new Set(week.map(keyOf));
  const weekKat = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0, lain: 0 };
  weekDesa.forEach(kk => { weekKat[katByDesa[kk]?.kat || 'lain']++; });
  const planNames = new Set(week.map(p => (p.nama_sf || '').trim().toUpperCase()));
  const planCodes = new Set(week.map(p => (p.username_sf || '').trim().toUpperCase()));
  const noPlan = sf.filter(s => !planCodes.has((s.code || '').trim().toUpperCase()) && !planNames.has((s.nama || '').trim().toUpperCase()));
  A.djp = [];
  if (!week.length) A.djp.push(`Belum ada plan DJP tersimpan pada ${D.weekStart.slice(8)}–${D.refTgl.slice(8)} ${D.periode}.`);
  else {
    A.djp.push(`${week.length} ODP diplan untuk ${Object.keys(sfPlan).length} SF di ${weekDesa.size} desa (7 hari terakhir).`);
    const good = weekKat.prioritas + weekKat.strategis;
    A.djp.push(`${pct(good, weekDesa.size) ?? 0}% desa yang dikunjungi berkategori Prioritas/Strategis; ${weekKat.skip} desa Skip.`);
    if (noPlan.length) A.djp.push(`${noPlan.length} SF belum punya plan minggu ini.`);
    A.djp.push('Realisasi kunjungan belum terbaca (menunggu CSV realisasi asli).');
  }
  A.djpData = { week, sfPlan, weekKat, weekDesa };

  // ── Timeline / cakupan bulan ini ──
  const monthDesa = new Set(D.plan.map(keyOf));
  const cov = { prioritas: [0, 0], strategis: [0, 0], cadangan: [0, 0], skip: [0, 0] };
  D.desa.forEach(x => { cov[x.kat][1]++; if (monthDesa.has((x.kec || '').toUpperCase() + '||' + (x.desa || '').toUpperCase())) cov[x.kat][0]++; });
  const planByKat = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0, lain: 0 };
  D.plan.forEach(p => { planByKat[katByDesa[keyOf(p)]?.kat || 'lain']++; });
  A.tl = [];
  A.tl.push(`${monthDesa.size} desa sudah diplan bulan ini (${D.plan.length} ODP).`);
  A.tl.push(`Cakupan Prioritas ${cov.prioritas[0]}/${cov.prioritas[1]} desa (${fmtPct(pct(cov.prioritas[0], cov.prioritas[1]))}), Strategis ${cov.strategis[0]}/${cov.strategis[1]}.`);
  const lowEff = planByKat.cadangan + planByKat.skip;
  if (D.plan.length) A.tl.push(`${pct(lowEff, D.plan.length)}% kunjungan ODP jatuh ke desa Cadangan/Skip` +
    (pct(lowEff, D.plan.length) > 30 ? ' — terlalu besar, alihkan ke desa Prioritas.' : '.'));
  A.cov = cov;

  // ── Sebaran desa Prioritas & Strategis ──
  const ps = D.desa.filter(x => x.kat === 'prioritas' || x.kat === 'strategis');
  const withPS = ps.filter(x => (x.psAktif || 0) > 0), noPS = ps.filter(x => !(x.psAktif > 0));
  A.map = [];
  A.map.push(`${withPS.length} dari ${ps.length} desa Prioritas/Strategis sudah ada PS bulan ini; ${noPS.length} belum.`);
  const pNo = noPS.filter(x => x.kat === 'prioritas');
  if (pNo.length) A.map.push(`${pNo.length} desa Prioritas belum ada PS dengan total ${fmtN(sum(pNo.map(x => x.totalAvail)))} port avail.`);
  if (withPS.length) {
    const topD = [...withPS].sort((a, b) => b.psAktif - a.psAktif)[0];
    A.map.push(`Desa dengan PS terbanyak: ${topD.desa} (${topD.psAktif} PS).`);
  }
  A.mapData = { ps, withPS, noPS };

  // ── Desa Prioritas belum digarap ──
  A.todo = D.desa.filter(x => x.kat === 'prioritas' && !(x.psAktif > 0) &&
    !monthDesa.has((x.kec || '').toUpperCase() + '||' + (x.desa || '').toUpperCase()))
    .sort((a, b) => b.skorPotensi - a.skorPotensi);

  // ── Ringkasan & rekomendasi ──
  A.summary = [];
  if (k.tgtPS > 0) {
    if (D.sisa === 0) A.summary.push(`Periode selesai dengan ${achPS}% target PS (${fmtN(k.ps)} dari ${fmtN(k.tgtPS)}).`);
    else A.summary.push(projPct >= 100 ? `On track: proyeksi ${projPct}% target PS.` : `Berisiko: proyeksi ${projPct}% target PS, butuh ${fmt1(needDrr)} PS/hari.`);
  }
  A.summary.push(A.trend[0]);
  if (ps0.length) A.summary.push(`${ps0.length} SF masih 0 PS — butuh pendampingan langsung minggu ini.`);
  A.summary.push(`Cakupan desa Prioritas bulan ini ${fmtPct(pct(cov.prioritas[0], cov.prioritas[1]))}.`);
  A.reco = [];
  if (needDrr > drr) A.reco.push(`Naikkan run rate ke ≥ ${fmt1(needDrr)} PS/hari: target per SF aktif ±${fmt1(needDrr / Math.max(aktif.length, 1))} PS/hari.`);
  if (A.todo.length) A.reco.push(`Jadwalkan DJP ke ${Math.min(A.todo.length, 10)} desa Prioritas teratas yang belum digarap (slide terakhir).`);
  if (lowEff && pct(lowEff, D.plan.length) > 30) A.reco.push('Kurangi kunjungan ke desa Cadangan/Skip, pindahkan ke Prioritas/Strategis.');
  if (behind.length) A.reco.push(`Coaching 1-on-1 untuk ${behind.map(s => shortName(s.nama)).join(', ')}.`);
  if (D.kendala[0]) A.reco.push(`Eskalasi kendala terbanyak: ${D.kendala[0].name} (${D.kendala[0].n} order).`);
  return A;
}

// ================================================================
//  3. SUSUN SLIDE
// ================================================================
function reportBuild(D, A) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';               // 10 × 5.625 in
  pres.author = 'WOK Dashboard';
  pres.title = `Review Mingguan ${D.cluster} – ${D.periode}`;
  pres.theme = { headFontFace: R_FONT, bodyFontFace: R_FONT };
  const now = new Date();
  const dibuat = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });

  pres.defineSlideMaster({
    title: 'KONTEN',
    background: { color: RP.white },
    objects: [
      { text: { text: `WOK Lampung Barat – Way Kanan · ${D.cluster} · ${D.periode}`,
        options: { x: 0.5, y: 5.25, w: 7, h: 0.25, fontSize: 9, color: RP.muted, fontFace: R_FONT, margin: 0 } } },
      { placeholder: { options: { name: 'title', type: 'title', x: 0.5, y: 0.3, w: 9, h: 0.55,
        fontSize: 24, bold: true, color: RP.ink, fontFace: R_FONT, margin: 0, valign: 'middle', align: 'left' }, text: '' } },
    ],
    slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.3, fontSize: 9, color: RP.muted, fontFace: R_FONT, align: 'right' },
  });

  const title = (s, t) => s.addText(t, { placeholder: 'title', align: 'left' });
  const sub = (s, t) => s.addText(t, { x: 0.5, y: 0.82, w: 9, h: 0.3, fontSize: 12, color: RP.muted, margin: 0, isTextBox: true });
  const card = (s, x, y, w, h, fill = RP.card) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 });
  const insight = (s, items, x = 6.35, y = 1.25, w = 3.15, h = 3.85, head = 'Analisa') => {
    card(s, x, y, w, h, RP.brandSoft);
    s.addText(head, { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: 0.3, fontSize: 13, bold: true, color: RP.brand, margin: 0, isTextBox: true });
    const list = (items.length ? items : ['Data belum cukup untuk dianalisa.']).slice(0, 5);
    s.addText(list.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < list.length - 1, paraSpaceAfter: 5 } })),
      { x: x + 0.15, y: y + 0.5, w: w - 0.3, h: h - 0.6, fontSize: 11, color: RP.ink, valign: 'top', margin: 2, isTextBox: true, fit: 'shrink' });
  };
  const stat = (s, x, y, w, val, label, color = RP.ink, fill = RP.card) => {
    card(s, x, y, w, 0.95, fill);
    s.addText(String(val), { x: x + 0.15, y: y + 0.08, w: w - 0.3, h: 0.5, fontSize: 24, bold: true, color, margin: 0, isTextBox: true });
    s.addText(label, { x: x + 0.15, y: y + 0.58, w: w - 0.3, h: 0.3, fontSize: 10, color: RP.muted, margin: 0, isTextBox: true });
  };
  const chartBase = (extra = {}) => Object.assign({
    catAxisLabelColor: RP.muted, valAxisLabelColor: RP.muted, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9,
    catAxisLabelFontFace: R_FONT, valAxisLabelFontFace: R_FONT, dataLabelFontFace: R_FONT, titleFontFace: R_FONT,
    valGridLine: { color: 'EEE7E3', size: 0.5 }, catGridLine: { style: 'none' },
    dataLabelFontSize: 9, dataLabelColor: RP.ink, titleFontSize: 11, titleColor: RP.ink,
  }, extra);
  const empty = (s, x, y, w, h, t) => {
    card(s, x, y, w, h);
    s.addText(t, { x, y, w, h, fontSize: 12, color: RP.muted, align: 'center', valign: 'middle', isTextBox: true });
  };

  const k = D.kpi, core = A.core;

  // ── 1. Cover ──
  let s = pres.addSlide();
  s.background = { color: RP.dark };
  s.addShape(pres.shapes.OVAL, { x: 6.6, y: -1.2, w: 5, h: 5, fill: { color: RP.brand, transparency: 70 }, line: { color: RP.dark, transparency: 100 } });
  s.addText('Review Mingguan', { x: 0.6, y: 1.35, w: 7, h: 0.5, fontSize: 18, color: 'F3C7BB', margin: 0, isTextBox: true });
  s.addText('Sales & DJP', { x: 0.6, y: 1.8, w: 8, h: 0.9, fontSize: 44, bold: true, color: RP.white, margin: 0, isTextBox: true });
  s.addText(`Cluster ${D.cluster} · ${D.periode}`, { x: 0.6, y: 2.75, w: 8, h: 0.45, fontSize: 20, color: RP.white, margin: 0, isTextBox: true });
  s.addText(`Data s.d. hari ke-${D.hari} · dibuat ${dibuat}`, { x: 0.6, y: 4.6, w: 8, h: 0.35, fontSize: 12, color: 'D9B8AE', margin: 0, isTextBox: true });

  // ── 2. Ringkasan analisa ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Ringkasan & rekomendasi');
  sub(s, 'Temuan utama minggu ini, dihitung otomatis dari data dashboard');
  const okProj = core.projPct != null && core.projPct >= 100;
  stat(s, 0.5, 1.25, 2.1, fmtN(k.ps), `PS · ach ${fmtPct(core.achPS)}`);
  stat(s, 2.75, 1.25, 2.1, fmtPct(core.projPct), D.sisa ? 'Proyeksi akhir bulan' : 'Ach akhir periode', okProj ? RP.green : RP.brand, okProj ? RP.greenSoft : RP.brandSoft);
  stat(s, 5.0, 1.25, 2.1, fmt1(core.drr), 'DRR sekarang (PS/hari)');
  stat(s, 7.25, 1.25, 2.25, core.needDrr ? fmt1(core.needDrr) : '–', 'DRR dibutuhkan', core.needDrr > core.drr ? RP.amber : RP.green);
  s.addText('Temuan', { x: 0.5, y: 2.45, w: 4.3, h: 0.3, fontSize: 14, bold: true, color: RP.ink, margin: 0, isTextBox: true });
  s.addText(A.summary.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < A.summary.length - 1, paraSpaceAfter: 6 } })),
    { x: 0.5, y: 2.8, w: 4.3, h: 2.3, fontSize: 12, color: RP.ink, valign: 'top', margin: 0, isTextBox: true, fit: 'shrink' });
  insight(s, A.reco, 5.0, 2.45, 4.5, 2.65, 'Rekomendasi minggu depan');

  // ── 3. KPI ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Ringkasan KPI');
  sub(s, `Hari ke-${D.hari} dari ${D.days} · ${D.sisa} hari tersisa`);
  const rowsK = [['', 'Aktual', 'Target', 'Ach', 'vs bln lalu'],
    ['IO', fmtN(k.io), fmtN(k.tgtIO), fmtPct(k.momIO), sgn(k.momIOvsLast)],
    ['RE', fmtN(k.re), fmtN(k.tgtRE), fmtPct(k.momRE), sgn(k.momREvsLast)],
    ['PS', fmtN(k.ps), fmtN(k.tgtPS), fmtPct(k.momPS), sgn(k.momPSvsLast)]];
  s.addTable(rowsK.map((r, i) => r.map((c, j) => ({ text: c, options: {
    bold: i === 0 || j === 0, color: i === 0 ? RP.muted : RP.ink, fill: { color: i === 0 ? RP.card : RP.white },
    align: j === 0 ? 'left' : 'right', fontSize: i === 0 ? 10 : 13 } }))),
    { x: 0.5, y: 1.25, w: 5.6, colW: [0.8, 1.2, 1.2, 1.1, 1.3], rowH: 0.36, fontFace: R_FONT, border: { type: 'solid', color: RP.line, pt: 0.5 }, margin: 0.06 });
  s.addChart(pres.charts.BAR, [{ name: 'Funnel', labels: ['IO', 'RE', 'PS'], values: [k.io, k.re, k.ps] }],
    chartBase({ x: 0.5, y: 2.85, w: 5.6, h: 2.25, barDir: 'bar', chartColors: ['E8A08F', 'D96A50', RP.brand], showValue: true,
      dataLabelPosition: 'outEnd', showTitle: true, title: `Funnel order · RE/IO ${k.reIO}% · PS/RE ${k.psRE}%`, showLegend: false,
      valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin' }));
  insight(s, A.kpi);

  // ── 4. Tren & D2D ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Tren harian & per kabupaten');
  sub(s, `7 hari terakhir ${fmtN(A.w.w0)} PS · sebelumnya ${fmtN(A.w.w1)} PS`);
  const dl = [], dv = [], dt = [];
  for (let i = 1; i <= D.hari; i++) { dl.push(String(i)); dv.push(D.psDay[i]); dt.push(Math.round(A.w.dailyTgt * 10) / 10); }
  if (dl.length) {
    s.addChart([
      { type: pres.charts.BAR, data: [{ name: 'PS harian', labels: dl, values: dv }], options: { chartColors: [RP.brand], barGapWidthPct: 40 } },
      { type: pres.charts.LINE, data: [{ name: 'Target harian', labels: dl, values: dt }], options: { chartColors: [RP.dark], lineSize: 1.5, lineDash: 'dash', lineDataSymbol: 'none' } },
    ], chartBase({ x: 0.5, y: 1.2, w: 5.6, h: 2.1, showLegend: true, legendPos: 'b', legendFontSize: 9, legendFontFace: R_FONT, legendColor: RP.muted,
      showTitle: true, title: 'PS per hari vs target harian' }));
  } else empty(s, 0.5, 1.2, 5.6, 2.1, 'Belum ada data harian');
  const kb = D.kab.filter(x => x.tPS > 0 || x.ps > 0);
  if (kb.length) {
    s.addChart(pres.charts.BAR, [{ name: 'Ach PS', labels: kb.map(x => x.kab), values: kb.map(x => x.momPS ?? 0) }],
      chartBase({ x: 0.5, y: 3.35, w: 5.6, h: 1.8, barDir: 'bar', chartColors: [RP.brand], showValue: true, dataLabelPosition: 'outEnd',
        dataLabelFormatCode: '0"%"', showTitle: true, title: `Ach PS per kabupaten (waktu berjalan ${core.timePct}%)`, showLegend: false,
        valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin' }));
  }
  insight(s, A.trend);

  // ── 5. Deepdive ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Deepdive channel, produk & speed');
  sub(s, 'Jumlah PS bulan berjalan');
  const ddc = (list, x, ttl, color) => {
    const L = list.filter(v => v.ps > 0).slice(0, 6);
    if (!L.length) return empty(s, x, 1.25, 1.85, 3.85, 'Tidak ada data');
    s.addChart(pres.charts.BAR, [{ name: ttl, labels: L.map(v => String(v.name).slice(0, 18)), values: L.map(v => v.ps) }],
      chartBase({ x, y: 1.25, w: 1.85, h: 3.85, barDir: 'bar', chartColors: [color], showValue: true, dataLabelPosition: 'outEnd',
        showTitle: true, title: ttl, showLegend: false, valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin' }));
  };
  ddc(D.channel, 0.5, 'Channel', RP.brand); ddc(D.produk, 2.45, 'Produk', 'D96A50');
  ddc(D.speed.map(v => ({ ...v, name: /^\d+$/.test(String(v.name).trim()) ? v.name + ' Mbps' : v.name })), 4.4, 'Speed', RP.dark);
  insight(s, A.deep);

  // ── 6. Performa SF ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Performa SF');
  const aktifN = D.sf.filter(x => x.ps > 0).length;
  sub(s, `${D.sf.length} SF aktif di cluster · ${aktifN} sudah PS`);
  const tiers = TIER_CONFIG.map(t => t.name);
  s.addChart(pres.charts.BAR, [
    { name: 'Bulan lalu', labels: tiers, values: tiers.map(t => D.tier.lm[t] || 0) },
    { name: 'Bulan ini', labels: tiers, values: tiers.map(t => D.tier.now[t] || 0) },
  ], chartBase({ x: 0.5, y: 1.2, w: 2.9, h: 3.9, barDir: 'col', chartColors: ['D9B8AE', RP.brand], showValue: true, dataLabelPosition: 'outEnd',
    showLegend: true, legendPos: 'b', legendFontSize: 9, legendFontFace: R_FONT, legendColor: RP.muted, showTitle: true, title: 'Sebaran tier SF',
    valAxisHidden: true, valGridLine: { style: 'none' } }));
  const sfRows = (list, head) => [[{ text: head, options: { bold: true, color: RP.muted, fontSize: 10, fill: { color: RP.card } } },
    { text: 'PS', options: { bold: true, color: RP.muted, fontSize: 10, align: 'right', fill: { color: RP.card } } },
    { text: 'Ach', options: { bold: true, color: RP.muted, fontSize: 10, align: 'right', fill: { color: RP.card } } }]]
    .concat(list.map(x => [{ text: x.nama.slice(0, 22) }, { text: fmtN(x.ps), options: { align: 'right' } }, { text: x.target ? x.ach + '%' : '–', options: { align: 'right' } }]));
  s.addTable(sfRows(A.sfLists.top, 'Top 5 SF'), { x: 3.55, y: 1.2, w: 2.6, colW: [1.6, 0.45, 0.55], fontSize: 10, fontFace: R_FONT, color: RP.ink, rowH: 0.27, border: { type: 'solid', color: RP.line, pt: 0.5 }, margin: 0.04 });
  s.addTable(sfRows(A.sfLists.bottom, 'Perlu coaching'), { x: 3.55, y: 3.05, w: 2.6, colW: [1.6, 0.45, 0.55], fontSize: 10, fontFace: R_FONT, color: RP.ink, rowH: 0.27, border: { type: 'solid', color: RP.line, pt: 0.5 }, margin: 0.04 });
  insight(s, A.sf);

  // ── 7. DJP minggu ini ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'DJP minggu ini');
  sub(s, `Plan ${D.weekStart.slice(8)}–${D.refTgl.slice(8)} ${D.periode}`);
  const sp = Object.entries(A.djpData.sfPlan).map(([n, set]) => ({ n, v: set.size })).sort((a, b) => b.v - a.v).slice(0, 12);
  if (sp.length) {
    s.addChart(pres.charts.BAR, [{ name: 'ODP diplan', labels: sp.map(x => x.n.slice(0, 18)), values: sp.map(x => x.v) }],
      chartBase({ x: 0.5, y: 1.2, w: 3.6, h: 3.9, barDir: 'bar', chartColors: [RP.brand], showValue: true, dataLabelPosition: 'outEnd',
        showTitle: true, title: 'ODP diplan per SF', showLegend: false, valAxisHidden: true, valGridLine: { style: 'none' }, catAxisOrientation: 'maxMin' }));
  } else empty(s, 0.5, 1.2, 3.6, 3.9, 'Belum ada plan tersimpan minggu ini');
  const wk = A.djpData.weekKat;
  stat(s, 4.3, 1.2, 1.85, fmtN(A.djpData.week.length), 'ODP diplan');
  stat(s, 4.3, 2.25, 1.85, fmtN(A.djpData.weekDesa.size), 'Desa dikunjungi');
  stat(s, 4.3, 3.3, 1.85, fmtN(wk.prioritas + wk.strategis), 'Desa Prioritas/Strategis', RP.brand);
  insight(s, A.djp);

  // ── 8. Timeline / cakupan desa ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Cakupan desa bulan ini');
  sub(s, 'Desa yang sudah masuk plan DJP dibanding total desa per kategori');
  const kats = ['prioritas', 'strategis', 'cadangan', 'skip'], kl = ['Prioritas', 'Strategis', 'Cadangan', 'Skip'];
  s.addChart(pres.charts.BAR, [
    { name: 'Sudah diplan', labels: kl, values: kats.map(x => A.cov[x][0]) },
    { name: 'Belum diplan', labels: kl, values: kats.map(x => A.cov[x][1] - A.cov[x][0]) },
  ], chartBase({ x: 0.5, y: 1.2, w: 5.6, h: 3.9, barDir: 'col', barGrouping: 'stacked', chartColors: [RP.brand, 'EADFDA'], showValue: true,
    dataLabelPosition: 'ctr', showLegend: true, legendPos: 'b', legendFontSize: 9, legendFontFace: R_FONT, legendColor: RP.muted,
    showTitle: true, title: 'Jumlah desa per kategori', valAxisHidden: true, valGridLine: { style: 'none' } }));
  insight(s, A.tl);

  // ── 9. Sebaran desa Prioritas & Strategis (peta titik) ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Sebaran desa Prioritas & Strategis');
  sub(s, 'Titik penuh = sudah ada PS bulan ini · titik kosong = belum ada PS');
  const pts = A.mapData.ps.filter(x => x.lat && x.lng);
  const MX = 0.5, MY = 1.2, MW = 3.9, MH = 3.9;
  card(s, MX, MY, MW, MH);
  if (pts.length) {
    const lats = pts.map(p => p.lat), lngs = pts.map(p => p.lng);
    let la0 = Math.min(...lats), la1 = Math.max(...lats), ln0 = Math.min(...lngs), ln1 = Math.max(...lngs);
    const pad = 0.25, iw = MW - 2 * pad, ih = MH - 2 * pad;
    const sc = Math.min(iw / Math.max(ln1 - ln0, 1e-6), ih / Math.max(la1 - la0, 1e-6));
    const ox = MX + pad + (iw - (ln1 - ln0) * sc) / 2, oy = MY + pad + (ih - (la1 - la0) * sc) / 2;
    const P = (la, ln) => ({ x: ox + (ln - ln0) * sc, y: oy + (la1 - la) * sc });
    pts.sort((a, b) => (a.kat === 'prioritas') - (b.kat === 'prioritas'));
    pts.forEach(p => {
      const c = p.kat === 'prioritas' ? RP.prioritas : RP.strategis, q = P(p.lat, p.lng), r = 0.12;
      s.addShape(pres.shapes.OVAL, { x: q.x - r / 2, y: q.y - r / 2, w: r, h: r,
        fill: { color: p.psAktif > 0 ? c : RP.white }, line: { color: c, width: 1.25 } });
    });
    (typeof OFFICES !== 'undefined' ? OFFICES : []).filter(o => D.kabs.includes(o.kab)).forEach(o => {
      if (o.lat < la0 - 0.05 || o.lat > la1 + 0.05 || o.lng < ln0 - 0.05 || o.lng > ln1 + 0.05) return;
      const q = P(o.lat, o.lng);
      s.addShape(pres.shapes.RECTANGLE, { x: q.x - 0.05, y: q.y - 0.05, w: 0.1, h: 0.1, fill: { color: RP.dark }, line: { color: RP.white, width: 0.75 } });
    });
    const leg = [[RP.prioritas, 'Prioritas'], [RP.strategis, 'Strategis']];
    leg.forEach(([c, t], i) => {
      s.addShape(pres.shapes.OVAL, { x: MX + 0.2 + i * 1.2, y: MY + MH - 0.3, w: 0.12, h: 0.12, fill: { color: c }, line: { color: c } });
      s.addText(t, { x: MX + 0.37 + i * 1.2, y: MY + MH - 0.36, w: 0.9, h: 0.24, fontSize: 9, color: RP.muted, margin: 0, isTextBox: true });
    });
    s.addShape(pres.shapes.RECTANGLE, { x: MX + 2.6, y: MY + MH - 0.29, w: 0.1, h: 0.1, fill: { color: RP.dark }, line: { color: RP.dark } });
    s.addText('Kantor', { x: MX + 2.75, y: MY + MH - 0.36, w: 0.8, h: 0.24, fontSize: 9, color: RP.muted, margin: 0, isTextBox: true });
  } else s.addText('Data desa belum termuat', { x: MX, y: MY, w: MW, h: MH, fontSize: 12, color: RP.muted, align: 'center', valign: 'middle', isTextBox: true });
  const mp = A.mapData;
  stat(s, 4.55, 1.2, 1.6, fmtN(mp.withPS.length), 'Desa sudah ada PS', RP.green, RP.greenSoft);
  stat(s, 4.55, 2.25, 1.6, fmtN(mp.noPS.length), 'Desa belum ada PS', RP.brand, RP.brandSoft);
  stat(s, 4.55, 3.3, 1.6, fmtPct(pct(mp.withPS.length, mp.ps.length)), 'Tingkat tergarap');
  insight(s, A.map);

  // ── 10. Desa Prioritas belum digarap ──
  s = pres.addSlide({ masterName: 'KONTEN' });
  title(s, 'Desa Prioritas belum digarap');
  sub(s, `Belum diplan DJP dan belum ada PS bulan ini · ${A.todo.length} desa · urut skor potensi`);
  const td = A.todo.slice(0, 10);
  if (td.length) {
    const head = ['#', 'Desa', 'Kecamatan', 'Kabupaten', 'Port avail', 'Jarak kantor', 'Skor'].map((t, j) => ({ text: t,
      options: { bold: true, color: RP.muted, fontSize: 10, fill: { color: RP.card }, align: j >= 4 || j === 0 ? 'right' : 'left' } }));
    const rows = td.map((x, i) => [String(i + 1), x.desa, x.kec, x.kab, fmtN(x.totalAvail),
      x.nearestDist < 9999 ? fmt1(x.nearestDist) + ' km' : '–', String(x.skorPotensi)]
      .map((t, j) => ({ text: t, options: { align: j >= 4 || j === 0 ? 'right' : 'left', bold: j === 1 } })));
    s.addTable([head, ...rows], { x: 0.5, y: 1.25, w: 9, colW: [0.4, 2.0, 1.7, 1.7, 1.1, 1.2, 0.9], fontSize: 11, fontFace: R_FONT,
      color: RP.ink, rowH: 0.33, border: { type: 'solid', color: RP.line, pt: 0.5 }, margin: 0.05 });
  } else empty(s, 0.5, 1.25, 9, 3.8, 'Semua desa Prioritas sudah diplan atau sudah ada PS');

  return pres;
}

// ================================================================
//  4. ENTRY: dipanggil tombol "PPT"
// ================================================================
async function exportWeeklyPPT(app) {
  await loadPptLib();
  const D = await reportCollect(app);
  const A = reportAnalyze(D);
  const pres = reportBuild(D, A);
  const tgl = localYMD ? localYMD(new Date()) : new Date().toISOString().slice(0, 10);
  const name = `Review-Mingguan-${D.cluster.replace(/\s+/g, '')}-${tgl}.pptx`;
  await pres.writeFile({ fileName: name });
  return name;
}
