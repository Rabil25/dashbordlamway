// ================================================================
//  Export PPT — ADO (Analisis Daerah Operasi), 3 slide per cluster aktif
//  1. Profil & Potensi Wilayah  2. Kinerja & Eksekusi DJP  3. Isu & Rencana Aksi
//  Memakai data yang sama dengan dashboard (reportCollect, Prioritas, ARPU,
//  DJP_Plan & DJP_Realisasi_Baru). Tidak mengubah data apa pun.
// ================================================================

async function adoCollect(app) {
  const D = await reportCollect(app);
  const U = v => String(v || '').trim().toUpperCase();
  // Realisasi DJP bulan ini
  let real = [];
  try {
    const r = await fetch(`/api/data?type=djp_real&bulan=${encodeURIComponent(D.bulan)}`, { headers: { Authorization: 'Bearer ' + app.sessionToken } });
    const d = await r.json(); real = (d && d.data) || [];
  } catch (e) {}
  real = real.filter(r => D.kabs.includes(U(r.kab)));
  try { await app.loadArpu(); } catch (e) {}

  // ── Profil wilayah ──
  const desa = D.desa;
  const kec = new Set(desa.map(d => d.kab + '|' + d.kec)).size;
  const odp = sum(desa.map(d => d.odpCount)), avail = sum(desa.map(d => d.totalAvail));
  let aSum = 0, aN = 0;
  desa.forEach(d => { const a = app.arpuOf(d.kec, d.desa); if (a && a.arpu) { aSum += a.arpu * a.n; aN += a.n; } });
  const kat = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0 };
  desa.forEach(d => { kat[d.kat] = (kat[d.kat] || 0) + 1; });
  const top = desa.filter(d => d.kat === 'prioritas' || d.kat === 'strategis')
    .sort((a, b) => b.skorPotensi - a.skorPotensi).slice(0, 5)
    .map(d => ({ desa: d.desa, kec: d.kec, kat: d.kat, skor: d.skorPotensi, avail: d.totalAvail, arpu: (app.arpuOf(d.kec, d.desa) || {}).arpu || null }));

  // ── Eksekusi DJP: s.d. tanggal realisasi terakhir (agar adil) ──
  const today = localYMD(new Date());
  const lastReal = real.reduce((m, r) => r.tgl > m ? r.tgl : m, '');
  const cut = lastReal && lastReal < today ? lastReal : today;
  const dk = (kec, ds) => U(kec) + '|' + U(ds);
  const planSF = new Set(), planDesa = new Set(), realSF = new Set(), realDesa = new Set(), planTouch = new Set(), realTouch = new Set();
  D.plan.forEach(p => { if (p.tgl > cut || !p.desa) return; planSF.add(U(p.username_sf) + '|' + p.tgl); planDesa.add(dk(p.kec, p.desa) + '|' + p.tgl); });
  D.plan.forEach(p => { if (p.desa) planTouch.add(dk(p.kec, p.desa)); });
  real.forEach(r => { realSF.add(U(r.username_sf) + '|' + r.tgl); if (r.desa) { realDesa.add(dk(r.kec, r.desa) + '|' + r.tgl); realTouch.add(dk(r.kec, r.desa)); } });
  const sfHit = [...planSF].filter(x => realSF.has(x)).length;
  const desaHit = [...planDesa].filter(x => realDesa.has(x)).length;
  const luarPlan = [...realDesa].filter(x => !planDesa.has(x)).length;
  const pot = desa.filter(d => d.kat === 'prioritas' || d.kat === 'strategis');
  const belum = pot.filter(d => !planTouch.has(dk(d.kec, d.desa)) && !realTouch.has(dk(d.kec, d.desa)))
    .sort((a, b) => b.skorPotensi - a.skorPotensi);
  const djp = {
    cut, planSF: planSF.size, sfHit, sfPct: pct(sfHit, planSF.size),
    planDesa: planDesa.size, desaHit, desaPct: pct(desaHit, planDesa.size), luarPlan,
    pot: pot.length, belum: belum.length, belumTop: belum.slice(0, 3).map(d => d.desa),
  };
  return { D, A: reportAnalyze(D), wil: { kab: D.kabs.length, kec, desa: desa.length, odp, avail, arpu: aN ? aSum / aN : null, kat, top }, djp };
}

function adoIssues(X) {
  const { D, A, djp, wil } = X, k = D.kpi, c = A.core;
  const I = [], T = [];
  const reIO = k.io > 0 ? k.re / k.io * 100 : 0, psRE = k.re > 0 ? k.ps / k.re * 100 : 0;
  if (c.achPS != null && c.achPS < c.timePct)
    I.push({ w: c.timePct - c.achPS, t: `PS ${c.achPS}% target, tertinggal ${c.timePct - c.achPS} poin dari waktu berjalan (${c.timePct}%).`,
      a: `Kejar ${fmt1(c.needDrr)} PS/hari di sisa ${D.sisa} hari; prioritaskan order RE yang belum PS.` });
  if (djp.sfPct != null && djp.sfPct < 80)
    I.push({ w: 100 - djp.sfPct, t: `Disiplin DJP rendah: hanya ${djp.sfPct}% SF-hari yang di-plan benar-benar visit.`,
      a: `Review harian plan vs realisasi di Timeline Desa; SF yang plan tanpa visit dibahas di briefing pagi.` });
  if (djp.belum > 0)
    I.push({ w: pct(djp.belum, djp.pot) || 0, t: `${djp.belum} dari ${djp.pot} desa Prioritas/Strategis belum tersentuh plan maupun visit bulan ini.`,
      a: `Masukkan desa belum tersentuh ke plan 2 minggu ke depan${djp.belumTop.length ? ' (mulai ' + djp.belumTop.join(', ') + ')' : ''}.` });
  if (k.io > 0 && Math.min(reIO, psRE) < 70)
    I.push({ w: 70 - Math.min(reIO, psRE), t: `Konversi lemah di tahap ${reIO < psRE ? 'IO→RE (' + fmt1(reIO) + '%)' : 'RE→PS (' + fmt1(psRE) + '%)'}.`,
      a: reIO < psRE ? 'Validasi inputan SF sebelum submit; pantau order IO yang belum RE > 3 hari.' : 'Eskalasi kendala PS (teknis/ODP) bersama tim teknisi tiap minggu.' });
  if (djp.planDesa && djp.luarPlan > djp.desaHit)
    I.push({ w: 20, t: `Kunjungan di luar plan (${djp.luarPlan} desa-hari) lebih banyak dari kunjungan sesuai plan (${djp.desaHit}).`,
      a: 'Plan dibuat H-1 sesuai rencana nyata SF; pakai saran desa di menu Plan DJP.' });
  I.sort((a, b) => b.w - a.w);
  if (!I.length) I.push({ t: 'Tidak ada isu menonjol dari data dashboard.', a: 'Pertahankan ritme plan dan visit; fokus desa Prioritas.' });
  return I.slice(0, 3);
}

function adoBuild(X) {
  const { D, A, wil, djp } = X, k = D.kpi, c = A.core;
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.author = 'WOK Dashboard';
  pres.title = `ADO ${D.cluster} – ${D.periode}`;
  pres.theme = { headFontFace: R_FONT, bodyFontFace: R_FONT };
  pres.defineSlideMaster({
    title: 'ADO', background: { color: RP.white },
    objects: [
      { rect: { x: 0, y: 0, w: 10, h: 0.08, fill: { color: RP.brand }, line: { color: RP.brand } } },
      { text: { text: `Analisis Daerah Operasi · ${D.cluster} · ${D.periode}`, options: { x: 0.5, y: 5.25, w: 7, h: 0.25, fontSize: 9, color: RP.muted, fontFace: R_FONT, margin: 0 } } },
    ],
    slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.3, fontSize: 9, color: RP.muted, fontFace: R_FONT, align: 'right' },
  });
  const T = (s, t, sub) => {
    s.addText(t, { x: 0.5, y: 0.25, w: 9, h: 0.5, fontSize: 22, bold: true, color: RP.ink, margin: 0, isTextBox: true });
    s.addText(sub, { x: 0.5, y: 0.72, w: 9, h: 0.3, fontSize: 11.5, color: RP.brand, italic: true, margin: 0, isTextBox: true });
  };
  const card = (s, x, y, w, h, fill = RP.card) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 });
  const stat = (s, x, y, w, val, label, color = RP.ink) => {
    card(s, x, y, w, 0.82);
    s.addText(String(val), { x: x + 0.12, y: y + 0.06, w: w - 0.24, h: 0.45, fontSize: 20, bold: true, color, margin: 0, isTextBox: true, fit: 'shrink' });
    s.addText(label, { x: x + 0.12, y: y + 0.5, w: w - 0.24, h: 0.26, fontSize: 9, color: RP.muted, margin: 0, isTextBox: true });
  };
  const head = (s, x, y, w, t) => s.addText(t, { x, y, w, h: 0.28, fontSize: 12, bold: true, color: RP.ink, margin: 0, isTextBox: true });
  const rp = v => v ? 'Rp' + Math.round(v / 1000).toLocaleString('id-ID') + 'rb' : '–';
  const KL = { prioritas: 'Prioritas', strategis: 'Strategis', cadangan: 'Cadangan', skip: 'Skip' };

  // ── Slide 1: Profil & Potensi ──
  let s = pres.addSlide({ masterName: 'ADO' });
  const potN = wil.kat.prioritas + wil.kat.strategis;
  T(s, 'Profil & Potensi Wilayah', `${potN} dari ${wil.desa} desa berpotensi tinggi (Prioritas + Strategis), ${fmtN(wil.avail)} port siap dijual.`);
  const st1 = [[wil.kab, 'Kabupaten'], [wil.kec, 'Kecamatan'], [wil.desa, 'Desa ber-ODP'], [fmtN(wil.odp), 'ODP'], [fmtN(wil.avail), 'Port tersedia'], [rp(wil.arpu), 'ARPU rata-rata']];
  st1.forEach(([v, l], i) => stat(s, 0.5 + (i % 2) * 1.55, 1.2 + Math.floor(i / 2) * 0.95, 1.45, v, l, i === 4 ? RP.brand : RP.ink));
  head(s, 3.75, 1.2, 2.4, 'Sebaran desa per kategori');
  s.addChart(pres.charts.BAR, [{ name: 'Desa', labels: Object.values(KL), values: Object.keys(KL).map(x => wil.kat[x] || 0) }], {
    x: 3.65, y: 1.45, w: 2.5, h: 2.6, barDir: 'bar', catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' },
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 10, dataLabelColor: RP.ink, catAxisLabelFontSize: 10, catAxisLabelColor: RP.muted,
    chartColors: [RP.prioritas, RP.strategis, RP.cadangan, RP.skip], showLegend: false, barGapWidthPct: 40,
  });
  head(s, 6.35, 1.2, 3.2, 'Top 5 desa potensial');
  const rows = [[{ text: 'Desa', options: { bold: true } }, { text: 'Skor', options: { bold: true, align: 'right' } }, { text: 'Port', options: { bold: true, align: 'right' } }, { text: 'ARPU', options: { bold: true, align: 'right' } }]];
  wil.top.forEach(d => rows.push([
    { text: [{ text: d.desa, options: { bold: true, color: d.kat === 'prioritas' ? RP.prioritas : RP.strategis } }, { text: '\n' + d.kec, options: { fontSize: 8, color: RP.muted } }] },
    { text: String(d.skor), options: { align: 'right' } }, { text: fmtN(d.avail), options: { align: 'right' } }, { text: rp(d.arpu), options: { align: 'right' } }]));
  s.addTable(rows, { x: 6.35, y: 1.5, w: 3.15, colW: [1.45, 0.45, 0.55, 0.7], fontSize: 9.5, color: RP.ink, fontFace: R_FONT,
    border: { type: 'solid', pt: 0.5, color: RP.line }, fill: { color: RP.white }, margin: 0.04 });
  s.addText('Skor potensi = √port tersedia × (1 − okupansi) × (1 − penetrasi PS); kategori dari skor dan jarak ke kantor.',
    { x: 0.5, y: 4.15, w: 5.6, h: 0.5, fontSize: 8.5, italic: true, color: RP.muted, margin: 0, isTextBox: true });

  // ── Slide 2: Kinerja & Eksekusi ──
  s = pres.addSlide({ masterName: 'ADO' });
  const gap = c.achPS != null ? c.achPS - c.timePct : null;
  T(s, 'Kinerja & Eksekusi',
    `PS ${c.achPS ?? '–'}% target di ${c.timePct}% waktu berjalan` + (djp.sfPct != null ? ` · ${djp.sfPct}% SF yang di-plan benar-benar visit.` : '.'));
  head(s, 0.5, 1.15, 4.4, `Pencapaian MTD (hari ke-${D.hari} dari ${D.days})`);
  [['IO', k.io, k.tgtIO], ['RE', k.re, k.tgtRE], ['PS', k.ps, k.tgtPS]].forEach(([l, v, t], i) => {
    const p = pct(v, t);
    stat(s, 0.5 + i * 1.5, 1.45, 1.4, fmtN(v), `${l} · target ${fmtN(t)} (${fmtPct(p)})`, p != null && p < c.timePct ? RP.amber : RP.green);
  });
  s.addChart(pres.charts.BAR, [{ name: 'Order', labels: ['IO', 'RE', 'PS'], values: [k.io, k.re, k.ps] }], {
    x: 0.4, y: 2.4, w: 4.5, h: 1.9, barDir: 'col', valAxisHidden: true, valGridLine: { style: 'none' }, showValue: true, dataLabelPosition: 'outEnd',
    dataLabelFontSize: 10, catAxisLabelFontSize: 10, catAxisLabelColor: RP.muted, chartColors: [RP.brand], showLegend: false,
    showTitle: true, title: `Funnel · RE/IO ${k.reIO}% · PS/RE ${k.psRE}%`, titleFontSize: 10, titleColor: RP.ink,
  });
  s.addText(gap == null ? '' : gap >= 0 ? `Unggul ${gap} poin dari pace waktu.` : `Tertinggal ${-gap} poin dari pace; butuh ${fmt1(c.needDrr)} PS/hari (DRR sekarang ${fmt1(c.drr)}).`,
    { x: 0.5, y: 4.35, w: 4.4, h: 0.4, fontSize: 10, bold: true, color: gap != null && gap < 0 ? RP.amber : RP.green, margin: 0, isTextBox: true });
  card(s, 5.2, 1.15, 4.3, 3.65, RP.brandSoft);
  s.addText('Eksekusi DJP', { x: 5.4, y: 1.25, w: 3.9, h: 0.3, fontSize: 12, bold: true, color: RP.brand, margin: 0, isTextBox: true });
  s.addText(`data realisasi s.d. ${djp.cut}`, { x: 5.4, y: 1.5, w: 3.9, h: 0.22, fontSize: 8.5, italic: true, color: RP.muted, margin: 0, isTextBox: true });
  const bar = (y, label, p, detail) => {
    s.addText(label, { x: 5.4, y, w: 2.9, h: 0.25, fontSize: 10, color: RP.ink, margin: 0, isTextBox: true });
    s.addText(fmtPct(p), { x: 8.3, y, w: 1.0, h: 0.25, fontSize: 12, bold: true, align: 'right', color: p == null ? RP.muted : p >= 80 ? RP.green : p >= 50 ? RP.amber : RP.prioritas, margin: 0, isTextBox: true });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: y + 0.3, w: 3.9, h: 0.12, fill: { color: RP.white }, line: { color: RP.white } });
    if (p) s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: y + 0.3, w: 3.9 * Math.min(p, 100) / 100, h: 0.12, fill: { color: p >= 80 ? RP.green : p >= 50 ? RP.amber : RP.prioritas }, line: { color: RP.white, transparency: 100 } });
    s.addText(detail, { x: 5.4, y: y + 0.45, w: 3.9, h: 0.22, fontSize: 8.5, color: RP.muted, margin: 0, isTextBox: true });
  };
  bar(1.85, 'SF yang di-plan benar-benar visit', djp.sfPct, `${djp.sfHit} dari ${djp.planSF} SF-hari`);
  bar(2.6, 'Desa plan yang didatangi', djp.desaPct, `${djp.desaHit} dari ${djp.planDesa} desa-hari · ${djp.luarPlan} desa-hari visit di luar plan`);
  const tp = djp.pot ? 100 - pct(djp.belum, djp.pot) : null;
  bar(3.35, 'Desa potensial tersentuh bulan ini', tp, `${djp.pot - djp.belum} dari ${djp.pot} desa Prioritas + Strategis`);

  // ── Slide 3: Isu & Rencana Aksi ──
  s = pres.addSlide({ masterName: 'ADO' });
  const I = adoIssues(X);
  T(s, 'Isu & Rencana Aksi', I[0].t);
  head(s, 0.5, 1.15, 4.3, 'Isu utama');
  head(s, 5.2, 1.15, 4.3, 'Rencana aksi');
  I.forEach((x, i) => {
    const y = 1.5 + i * 0.85;
    card(s, 0.5, y, 4.4, 0.75);
    s.addShape(pres.shapes.OVAL, { x: 0.62, y: y + 0.2, w: 0.34, h: 0.34, fill: { color: RP.brand }, line: { color: RP.brand } });
    s.addText(String(i + 1), { x: 0.62, y: y + 0.2, w: 0.34, h: 0.34, fontSize: 11, bold: true, color: RP.white, align: 'center', valign: 'middle', margin: 0, isTextBox: true });
    s.addText(x.t, { x: 1.08, y: y + 0.05, w: 3.75, h: 0.65, fontSize: 10, color: RP.ink, valign: 'middle', margin: 0, isTextBox: true, fit: 'shrink' });
    card(s, 5.2, y, 4.3, 0.75, RP.brandSoft);
    s.addText(x.a, { x: 5.35, y: y + 0.05, w: 4.05, h: 0.65, fontSize: 10, color: RP.ink, valign: 'middle', margin: 0, isTextBox: true, fit: 'shrink' });
  });
  const yk = 1.5 + I.length * 0.85 + 0.05;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: yk, w: 9, h: Math.max(4.95 - yk, 0.6), fill: { color: RP.white }, line: { color: RP.muted, dashType: 'dash', width: 1 }, rectRadius: 0.08 });
  s.addText([{ text: 'Kompetisi & catatan lapangan  ', options: { bold: true, color: RP.ink } }, { text: '(diisi manual: provider pesaing, harga, kondisi wilayah)', options: { italic: true, color: RP.muted } }],
    { x: 0.65, y: yk + 0.08, w: 8.7, h: 0.3, fontSize: 10, margin: 0, isTextBox: true });
  return pres;
}

async function exportAdoPPT(app) {
  await loadPptLib();
  const X = await adoCollect(app);
  const pres = adoBuild(X);
  const name = `ADO-${X.D.cluster.replace(/\s+/g, '')}-${localYMD(new Date())}.pptx`;
  await pres.writeFile({ fileName: name });
  return name;
}
