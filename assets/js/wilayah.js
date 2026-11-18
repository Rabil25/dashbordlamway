// ================================================================
//  PETA WILAYAH (Nov 2026) — satu peta: batas desa/kecamatan/kabupaten
//  (assets/geo/*.json, dari GeoJSON batas desa BIG, 5 kabupaten WOK),
//  kategori desa, kunjungan DJP 3 bulan (SF × desa × hari, bulan M-1..M-3),
//  order bulan berjalan, plan DJP per tanggal, kantor, dan sorotan.
//  Dipanggil dari app.renderPeta() setelah peta dasar & marker order dibuat.
//  Label: kabupaten (zoom jauh) → kecamatan → desa (zoom ≥ 12, anti-tabrakan).
// ================================================================
const wlRaw = o => (window.Alpine && Alpine.raw) ? Alpine.raw(o) : o;
const WL = {
  KAT_C: { prioritas: '#a8231c', strategis: '#d48a00', cadangan: '#8c8c8c', skip: '#cdb68a' },
  SF_C: ['#2f5f8a', '#7a3b6b', '#3f7d3a', '#c0622a', '#1f7a7a', '#8a6a1c', '#5a3e9a', '#a8231c'],
  geo: null, geoP: null, visits: {}, plans: {},
};
const wlN = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const wlK = (kec, desa) => wlN(kec) + '|' + wlN(desa);
const wlEsc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

function wlLoadGeo() {
  if (WL.geo) return Promise.resolve(WL.geo);
  if (!WL.geoP) WL.geoP = Promise.all(['desa', 'kec', 'kab'].map(n => fetch('assets/geo/' + n + '.json').then(r => r.json())))
    .then(([desa, kec, kab]) => (WL.geo = { desa, kec, kab }))
    .catch(e => { WL.geoP = null; throw e; });
  return WL.geoP;
}

// kunjungan DJP 3 bulan sebelum bulan aktif: Map desaKey -> {n, sf:{code:{n,nama,last}}, last}
async function wlLoadVisits(app) {
  const bulan = app.cfg.bulan;
  if (WL.visits[bulan]) return WL.visits[bulan];
  const ms = [prevPeriod(bulan)]; ms.push(prevPeriod(ms[0])); ms.push(prevPeriod(ms[1]));
  const auth = { headers: { Authorization: 'Bearer ' + app.sessionToken } };
  const res = await Promise.all(ms.map(m => fetch(`/api/data?type=djp_real&bulan=${encodeURIComponent(m)}`, auth).then(r => r.json()).catch(() => null)));
  const seen = new Set(), by = {}, avail = [];
  res.forEach((d, i) => {
    const rows = (d && d.data) || []; if (rows.length) avail.push(ms[i]);
    rows.forEach(p => {
      const dk = wlK(p.kec, p.desa), sf = wlN(p.username_sf), k = sf + '|' + dk + '|' + p.tgl;
      if (seen.has(k)) return; seen.add(k);
      const o = by[dk] = by[dk] || { n: 0, sf: {}, last: '' };
      o.n++; if (p.tgl > o.last) o.last = p.tgl;
      const s = o.sf[sf] = o.sf[sf] || { n: 0, nama: p.nama_sf || sf, last: '' };
      s.n++; if (p.tgl > s.last) s.last = p.tgl;
    });
  });
  return (WL.visits[bulan] = { by, months: ms.slice().reverse(), avail: avail.reverse() });
}

async function wlLoadPlan(app, tgl) {
  const bulan = String(tgl).slice(0, 7);
  if (!WL.plans[bulan]) {
    const auth = { headers: { Authorization: 'Bearer ' + app.sessionToken } };
    WL.plans[bulan] = fetch(`/api/data?type=djp_plan&bulan=${encodeURIComponent(bulan)}`, auth).then(r => r.json()).then(d => (d && d.data) || []).catch(() => []);
  }
  return (await WL.plans[bulan]).filter(p => p.tgl === tgl);
}

async function wlRender(app, map) {
  if (!map || !window.L) return;
  map = wlRaw(map);
  const st = app.wl;
  let geo;
  try { geo = await wlLoadGeo(); } catch (e) { showToast('Gagal memuat batas wilayah', 'error'); return; }
  if (wlRaw(app._leafletMap) !== wlRaw(map)) return;   // peta sudah diganti saat menunggu
  const kabs = (CLUSTER[app.cluster] || { kab: [] }).kab.map(k => k.toUpperCase());
  const inCl = f => kabs.includes(f.properties.kab) && (!st.kab || f.properties.kab === st.kab);

  // data desa: kategori (klasifikasi dashboard), PS bulan aktif, kunjungan 3 bln
  const desaInfo = {};
  (app.prioritasOdpRaw.length ? app.klasifikasiDesa('', '') : []).forEach(d => { desaInfo[wlK(d.kec, d.desa)] = d; });
  const ps = {}, ord = {};
  app.filteredRows.forEach(r => {
    const k = wlK(r[C.KEC], r[C.DESA]); if (!k || k === '|') return;
    if ((r[C.BLN_PS] || '').toString() === app.cfg.bulan && (r[C.VAL_PS] || '').toString().trim()) ps[k] = (ps[k] || 0) + 1;
    ord[k] = (ord[k] || 0) + 1;
  });
  let V = { by: {}, months: [], avail: [] };
  if (st.visit || st.sf) { try { V = await wlLoadVisits(app); } catch (e) {} }
  if (wlRaw(app._leafletMap) !== wlRaw(map)) return;
  const visN = k => { const v = V.by[k]; if (!v) return 0; return st.sf ? ((v.sf[st.sf] || {}).n || 0) : v.n; };
  const nM = Math.max(1, V.avail.length);

  // sorotan
  const flag = k => {
    const d = desaInfo[k], v = visN(k) / nM, p = ps[k] || 0;
    if (d && d.kat === 'prioritas' && !visN(k)) return 'untouched';
    if (v >= 5 && p <= 1) return 'jenuh';
    if (d && d.kat === 'skip' && p >= 2) return 'skipps';
    return '';
  };

  // bersihkan layer lama
  (app._wlLayers || []).forEach(l => map.removeLayer(l)); app._wlLayers = [];
  const add = l => { l.addTo(map); app._wlLayers.push(l); return l; };
  const pane = (n, z) => { if (!map.getPane(n)) { map.createPane(n); map.getPane(n).style.zIndex = z; } return n; };
  pane('wlArea', 350); pane('wlLine', 360); pane('wlLabel', 640);

  // mode batas: otomatis mengikuti zoom, atau manual
  const mode = () => st.mode === 'none' ? 'none' : st.mode !== 'auto' ? st.mode : (map.getZoom() >= 11 ? 'desa' : 'kec');

  const desaF = geo.desa.features.filter(f => f.geometry && inCl(f));
  const kecF = geo.kec.features.filter(f => f.geometry && inCl(f) && String(f.properties.kec).trim());
  const kabF = geo.kab.features.filter(f => f.geometry && inCl(f));

  // ringkasan per kecamatan
  const kecSum = {};
  desaF.forEach(f => {
    const k = wlK(f.properties.kec, f.properties.desa), kk = wlN(f.properties.kec);
    const s = kecSum[kk] = kecSum[kk] || { desa: 0, prio: 0, ps: 0, vis: 0 };
    s.desa++; s.ps += ps[k] || 0; s.vis += visN(k); if ((desaInfo[k] || {}).kat === 'prioritas') s.prio++;
  });
  const maxVis = Math.max(1, ...desaF.map(f => visN(wlK(f.properties.kec, f.properties.desa))));

  const desaStyle = f => {
    const k = wlK(f.properties.kec, f.properties.desa), d = desaInfo[k], v = visN(k), fl = st.hl ? flag(k) : '';
    const c = st.kat && d ? WL.KAT_C[d.kat] || '#cdb68a' : '#cdb68a';
    return { pane: 'wlArea', color: fl === 'untouched' ? '#a8231c' : fl === 'jenuh' ? '#c0622a' : fl === 'skipps' ? '#3f7d3a' : '#6b4423',
      weight: fl ? 1.6 : 0.5, dashArray: fl === 'untouched' ? '4 4' : null, fillColor: c,
      fillOpacity: (st.kat && d ? 0.22 : 0.08) + (st.visit ? 0.45 * Math.sqrt(v / maxVis) : 0) };
  };

  const card = (f) => {
    const p = f.properties, k = wlK(p.kec, p.desa), d = desaInfo[k] || {}, v = V.by[k];
    const top = v ? Object.entries(v.sf).sort((a, b) => b[1].n - a[1].n).slice(0, 3) : [];
    const sfPS = {}; app.filteredRows.forEach(r => { if (wlK(r[C.KEC], r[C.DESA]) === k && (r[C.BLN_PS] || '').toString() === app.cfg.bulan && (r[C.VAL_PS] || '').toString().trim()) { const s = wlN(r[C.CODE_SF]); sfPS[s] = (sfPS[s] || 0) + 1; } });
    const fl = flag(k), FL = { untouched: ['Prioritas belum disentuh', '#a8231c'], jenuh: ['Jenuh', '#c0622a'], skipps: ['Skip tapi ada PS', '#3f7d3a'] }[fl];
    const kat = d.kat ? `<span class="wl-tag" style="color:${WL.KAT_C[d.kat]}">${d.kat[0].toUpperCase() + d.kat.slice(1)}</span>` : '<span class="wl-tag" style="color:#8a735c">Tanpa data ODP</span>';
    return `<div class="wl-card"><div class="wl-card-h">${wlEsc(p.desa)}</div><div class="wl-card-s">${wlEsc(p.kec)} · ${wlEsc(p.kab)}</div>
      ${kat}${FL ? `<span class="wl-tag" style="color:${FL[1]}">${FL[0]}</span>` : ''}
      <div class="wl-kv"><div>Avai<b>${d.totalAvail != null ? d.totalAvail : '–'}</b></div><div>PS bln ini<b>${ps[k] || 0}</b></div>
      <div>Kunj. 3 bln<b>${v ? v.n : 0}</b></div><div>Terakhir<b class="sm">${v && v.last ? fmtTglId(v.last) : '–'}</b></div></div>
      ${top.length ? `<table class="wl-mini"><tr><th colspan="3">SF paling sering</th></tr>${top.map(([c, s]) => `<tr><td>${wlEsc(String(s.nama).replace(/\[.*?\]\s*/g, ''))}</td><td class="n">${s.n} kunj.</td><td class="n">${sfPS[c] || 0} PS</td></tr>`).join('')}</table>` : '<div class="wl-none">Belum pernah dikunjungi DJP 3 bulan terakhir.</div>'}
      <div class="wl-note">DJP: ${V.avail.length ? V.avail.map(m => periodLabel(m)).join(', ') : 'data belum tersedia'}</div></div>`;
  };

  const m = mode();
  if (m === 'none') {
    // tanpa batas wilayah: hanya order, plan, kantor
  } else if (m === 'desa') {
    add(L.geoJSON({ type: 'FeatureCollection', features: desaF }, {
      style: desaStyle,
      onEachFeature: (f, l) => {
        l.bindTooltip(wlEsc(f.properties.desa), { sticky: true, className: 'wl-tip', direction: 'top' });
        l.bindPopup(() => card(f), { maxWidth: 240, className: 'wl-pop', autoPanPadding: [20, 20] });
        l.on('mouseover', () => l.setStyle({ weight: 1.8 })).on('mouseout', () => l.setStyle(desaStyle(f)));
      },
    }));
    add(L.geoJSON({ type: 'FeatureCollection', features: kecF }, { style: { pane: 'wlLine', color: '#3b2a1e', weight: 1, dashArray: '6 4', fill: false, opacity: .55 }, interactive: false }));
  } else {
    const kecStyle = f => {
      const s = kecSum[wlN(f.properties.kec)] || { desa: 1, prio: 0, vis: 0 }, pr = s.prio / Math.max(1, s.desa);
      const c = m === 'kab' ? '#cdb68a' : pr > .4 ? '#a8231c' : pr > .2 ? '#d48a00' : '#8c8c8c';
      return { pane: 'wlArea', color: '#4a2c14', weight: m === 'kab' ? 0.4 : 0.8, dashArray: m === 'kab' ? '2 4' : null, fillColor: c, fillOpacity: m === 'kab' ? .12 : .18 + pr * .35 };
    };
    add(L.geoJSON({ type: 'FeatureCollection', features: kecF }, {
      style: kecStyle,
      onEachFeature: (f, l) => {
        const s = kecSum[wlN(f.properties.kec)] || { desa: 0, prio: 0, ps: 0, vis: 0 };
        l.bindTooltip(`<b>${wlEsc(f.properties.kec)}</b><br>${s.ps} PS · ${s.desa} desa · ${s.prio} prioritas · ${s.vis} kunjungan`, { sticky: true, className: 'wl-tip' });
        l.on('click', () => { st.mode = 'desa'; map.fitBounds(l.getBounds(), { padding: [20, 20] }); });
      },
    }));
  }
  if (m !== 'none') add(L.geoJSON({ type: 'FeatureCollection', features: kabF }, { style: { pane: 'wlLine', color: '#8a2a1c', weight: 1.4, fill: false, opacity: .75 }, interactive: false }));

  // ── label (anti-tabrakan) ──
  const lblLayer = add(L.layerGroup());
  const centroid = l => l.getBounds().getCenter();
  const drawLabels = () => {
    lblLayer.clearLayers();
    const z = map.getZoom(), mm = mode(), boxes = [], vb = map.getBounds().pad(0.05);
    if (mm === 'none') return;
    const place = (ll, html, cls, w, h, prio) => ({ ll, html, cls, w, h, prio });
    const cand = [];
    if (z < 10 || mm === 'kab') kabF.forEach(f => cand.push(place(L.geoJSON(f).getBounds().getCenter(), wlEsc(f.properties.kab), 'wl-lbl-kab', f.properties.kab.length * 11, 22, 3)));
    if (mm === 'kec' && z >= 9) kecF.forEach(f => { const s = kecSum[wlN(f.properties.kec)] || { ps: 0, desa: 0 }; cand.push(place(L.geoJSON(f).getBounds().getCenter(), `${wlEsc(f.properties.kec)}<small>${s.ps} PS · ${s.desa} desa</small>`, 'wl-lbl-kec', Math.max(f.properties.kec.length * 7.5, 80), 32, 2)); });
    if (mm === 'desa' && st.labels) desaF.forEach(f => {
      const k = wlK(f.properties.kec, f.properties.desa), d = desaInfo[k];
      const pr = (flag(k) ? 2 : 0) + (d && d.kat === 'prioritas' ? 1 : 0) + (ps[k] ? .5 : 0);
      cand.push(place(L.geoJSON(f).getBounds().getCenter(), wlEsc(f.properties.desa), 'wl-lbl-desa', f.properties.desa.length * 6.6, 16, pr));
    });
    cand.sort((a, b) => b.prio - a.prio).forEach(c => {
      if (!vb.contains(c.ll)) return;
      const p = map.latLngToContainerPoint(c.ll), b = [p.x - c.w / 2, p.y - c.h / 2, p.x + c.w / 2, p.y + c.h / 2];
      if (!st.labels && boxes.some(o => !(b[2] < o[0] || b[0] > o[2] || b[3] < o[1] || b[1] > o[3]))) return;
      boxes.push(b);
      lblLayer.addLayer(L.marker(c.ll, { pane: 'wlLabel', interactive: false, icon: L.divIcon({ className: 'wl-lbl ' + c.cls, html: c.html, iconSize: [c.w, c.h], iconAnchor: [c.w / 2, c.h / 2] }) }));
    });
  };
  drawLabels();
  if (app._wlZoomFn) map.off('zoomend moveend', app._wlZoomFn);
  let lastMode = m;
  app._wlZoomFn = () => { (app._wlLayers || []).forEach(l => { if (l instanceof L.Marker) l.update(); });
    if (mode() !== lastMode) { lastMode = mode(); wlRender(app, map); } else drawLabels(); };
  map.on('zoomend moveend', app._wlZoomFn);

  // ── kantor ──
  if (st.kantor) OFFICES.filter(o => !kabs.length || kabs.includes(String(o.kab).toUpperCase())).forEach(of => add(L.marker([of.lat, of.lng], {
    icon: L.divIcon({ className: '', html: '<div class="wl-office">🏰</div>', iconSize: [26, 26], iconAnchor: [13, 13] }),
  }).bindTooltip('Kantor ' + wlEsc(of.name), { className: 'wl-tip' })));

  // ── plan DJP tanggal terpilih: rute bernomor per SF ──
  if (st.plan && st.planTgl) {
    const rows = (await wlLoadPlan(app, st.planTgl)).filter(p => kabs.includes(String(p.kab).toUpperCase()) && (!st.sf || wlN(p.username_sf) === st.sf));
    if (wlRaw(app._leafletMap) !== wlRaw(map)) return;
    const odpLL = {}; (app.prioritasOdpRaw || []).forEach(o => { const lat = parseFloat(o.lat || o.LATITUDE), lng = parseFloat(o.lng || o.LONGITUDE); if (lat && lng) odpLL[String(o.odp || o.ODP).toUpperCase()] = [lat, lng]; });
    const bySF = {}; rows.forEach(p => (bySF[p.username_sf] = bySF[p.username_sf] || []).push(p));
    Object.entries(bySF).forEach(([sf, ps2], i) => {
      const col = WL.SF_C[i % WL.SF_C.length], pts = ps2.map(p => odpLL[String(p.odp).toUpperCase()] || (p.lat && p.lng ? [+p.lat, +p.lng] : null)).filter(Boolean);
      if (pts.length > 1) add(L.polyline(pts, { color: col, weight: 3, dashArray: '8 6', opacity: .9 }));
      ps2.forEach((p, j) => { const ll = odpLL[String(p.odp).toUpperCase()]; if (!ll) return;
        add(L.marker(ll, { icon: L.divIcon({ className: '', html: `<div class="wl-plan" style="background:${col}">${j + 1}</div>`, iconSize: [22, 22], iconAnchor: [11, 11] }) })
          .bindTooltip(`${wlEsc(p.nama_sf)} · ${wlEsc(p.odp)}<br>${wlEsc(p.desa)}`, { className: 'wl-tip' })); });
    });
    app.wlPlanInfo = rows.length ? `${rows.length} ODP · ${Object.keys(bySF).length} SF` : 'tidak ada plan';
  } else app.wlPlanInfo = '';

  // ringkasan sorotan & kecocokan nama
  let u = 0, j = 0, s = 0, matched = 0;
  desaF.forEach(f => { const k = wlK(f.properties.kec, f.properties.desa); if (desaInfo[k]) matched++; const fl = flag(k); if (fl === 'untouched') u++; else if (fl === 'jenuh') j++; else if (fl === 'skipps') s++; });
  app.wlStats = { untouched: u, jenuh: j, skipps: s, desa: desaF.length, matched, djp: V.avail.map(m2 => periodLabel(m2)).join(', ') };
}

function fmtTglId(t) {
  const d = new Date(String(t).slice(0, 10) + 'T00:00:00'); if (isNaN(d)) return t;
  return d.getDate() + ' ' + ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'][d.getMonth()];
}

// ================================================================
//  ANALISA WILAYAH (di bawah peta) — cluster aktif, per kabupaten.
//  Memakai mesin analisa PPT (report_an.js: anCollect mode 'cluster' + anAnalyze):
//  PS bulan aktif vs DJP 3 bulan sebelumnya (kunjungan = SF × desa × hari),
//  kategori desa, SF, funnel; setiap bagian ditutup silogisme.
// ================================================================
async function wlAnalisa(app, force) {
  if (typeof anCollect !== 'function') return;
  const key = app.cluster + '|' + app.cfg.bulan;
  if (!force && app.wlAn && app.wlAn.key === key) return;
  if (app.wlAnBusy) return;
  app.wlAnBusy = true;
  try {
    const D = await anCollect(app, 'cluster');
    const S = anAnalyze(D);
    const E = D.ents, sum2 = f => E.reduce((a, e) => a + (f(e) || 0), 0);
    const pct = (a, b) => b ? Math.round(a / b * 100) + '%' : '–';
    const d1 = v => v == null || !isFinite(v) ? '–' : (Math.round(v * 10) / 10).toLocaleString('id-ID');
    const tot = { ps: sum2(e => e.ps), tgt: sum2(e => e.tgtM), re: sum2(e => e.re), io: sum2(e => e.io), vis: sum2(e => e.vis), ol: sum2(e => e.ol),
      fromVis: sum2(e => e.psFromVis), prio: sum2(e => e.prio[0]), prioT: sum2(e => e.prio[1]), untouched: sum2(e => e.prioUntouched.length) };
    const obj = {
      key, label: D.label, djp: S.djpLbl, missing: S.missing, sisa: D.sisa,
      tiles: [
        ['PS ' + D.label, tot.ps, 'target ' + tot.tgt],
        ['Outlook akhir bulan', tot.ol + ' (' + pct(tot.ol, tot.tgt) + ')', D.sisa + ' hari tersisa'],
        ['PS / RE', pct(tot.ps, tot.re), 'IO ' + tot.io + ' · RE ' + tot.re],
        ['Kunjungan DJP ' + S.djpLbl, tot.vis.toLocaleString('id-ID'), 'SF × desa × hari'],
        ['PS dari desa ber-DJP', pct(tot.fromVis, tot.ps), tot.fromVis + ' dari ' + tot.ps + ' PS'],
        ['Desa prioritas tersentuh', tot.prioT + ' / ' + tot.prio, tot.untouched + ' belum disentuh'],
      ],
      kab: E.map(e => ({
        name: e.name, ps: e.ps, tgt: e.tgtM, ol: e.ol, olp: Math.round(e.olAch || 0), drr: d1(e.drrM), need: d1(e.need), psre: Math.round(e.psre || 0),
        batal: e.coh[1], pend: e.coh[2], vis: e.vis, visM: e.visByM, sfDJP: e.sfDJP, visSf: d1(e.visPerSfM), desaSf: d1(e.desaPerSf), visDesa: d1(e.visPerPair),
        prioVis: e.prioVisShare, fromVis: pct(e.psFromVis, e.ps), samePair: pct(e.psSamePair, e.ps), rutin: e.psRutin,
        bucket: e.bucket.map(b => b.avg == null ? '–' : d1(b.avg)), bucketN: e.bucket.map(b => b.n),
        sfB: e.sfB.map(b => b.avg == null ? '–' : d1(b.avg)), sfBN: e.sfB.map(b => b.n), corr: e.corr == null ? '–' : (Math.round(e.corr * 100) / 100).toLocaleString('id-ID'),
        psKat: e.psKat.map(v => pct(v, e.ps)), prio: e.prio, untouched: e.prioUntouched.slice(0, 4).join(', '), untouchedN: e.prioUntouched.length,
        jenuh: e.saturated.map(s => `${s.desa} (${s.v} kunj., ${s.ps} PS)`).join(', '), skipPS: e.skipProd.map(s => `${s.desa} (${s.ps} PS)`).join(', '),
        watch: e.watch.map(w => `${w.n} (${w.v} kunj.)`).join(', '), plan: e.planCur,
      })),
      bucketL: (E[0] || {}).bucketL || [], sfBL: (E[0] || {}).sfBL || [], months: D.djpM.map(m => periodLabel(m) + (D.availM.includes(m) ? '' : ' (n/a)')),
      syl: [['Pencapaian', S.pace], ['Funnel order', S.funnel], ['DJP 3 bulan → PS (wilayah desa)', S.djp], ['SF & teritori', S.sf], ['Kategori desa & plan', S.prio], ['Kesimpulan', S.final]],
      main: S.main.map(m => m.join('')), actions: S.actions,
    };
    Object.assign(obj, await wlBizAnalysis(app, D, E));
    app.wlAn = obj;
  } catch (e) { console.error('Analisa wilayah:', e); app.wlAn = { key, error: e.message }; }
  finally { app.wlAnBusy = false; }
}


// ── Analisa bisnis (IO-sentris): Driver Tree → Gap IO → Matriks Potensi × Penetrasi → Diagnosis & Aksi ──
// Prinsip (PJ): yang dikontrol penuh sales adalah IO; IO harus di lokasi potensial (desa Prioritas + Strategis).
// PS dipengaruhi faktor di luar sales (tarikan ≤250 m ke ODP, teknisi, cuaca) → ditampilkan sebagai hasil, bukan tuas utama.
// Desa Strategis = potensi tinggi tapi jauh → butuh jadwal khusus (dikelompokkan per kecamatan).
async function wlBizAnalysis(app, D, E) {
  const d1 = v => v == null || !isFinite(v) ? '–' : (Math.round(v * 10) / 10).toLocaleString('id-ID');
  const nM = Math.max(1, D.availM.length), scale = D.days / Math.max(1, D.refDay);
  const katOf = {}; (D.katMap ? Object.values(D.katMap) : (app.prioritasOdpRaw.length ? app.klasifikasiDesa('', '') : [])).forEach(d => { katOf[wlK(d.kec, d.desa)] = d; });
  // IO periode ini per entitas (dari anEntity): total & di lokasi potensial
  const io = E.map(e => ({ all: e.ioW, pot: e.ioPotW, potShare: e.ioPotShare || 0, tgt: e.ioTgt, mtd: e.ioMtd, ol: e.ioOl,
    olPot: Math.round(e.ioPotW * (e.ioOl / Math.max(1, e.ioW || e.ioMtd || 1))) }));
  // 1) DRIVER TREE (IO-sentris)
  const drv = E.map((e, i) => ({
    name: e.name,
    f: {
      sf: e.sfDJP, kunj: e.visPerSfM, sasaran: e.goodVisShare || 0,
      ioHasil: e.vis ? io[i].mtd / D.refDay * D.days / (e.vis / nM) * 10 : 0,
      ioPot: io[i].potShare,
      ioRe: e.io ? e.re / e.io * 100 : 0, rePs: e.psre || 0,
    },
  }));
  const FL = { sf: ['SF aktif DJP', '', 1], kunj: ['Kunjungan / SF / bulan', '', 1], sasaran: ['Kunjungan ke desa Prioritas + Strategis', '%', 1],
    ioHasil: ['IO per 10 kunjungan', '', 1], ioPot: ['IO di lokasi potensial (Prioritas + Strategis)', '%', 1],
    ioRe: ['Konversi IO → RE', '%', 1], rePs: ['Konversi RE → PS (di luar kendali penuh sales)', '%', 0] };
  const avg = k => drv.reduce((a, x) => a + x.f[k], 0) / Math.max(1, drv.length);
  const BENCH = { sasaran: 80, ioPot: 70, ioRe: 95, rePs: 80 };
  drv.forEach(x => {
    x.rows = Object.keys(FL).map(k => {
      const ref = BENCH[k] != null ? BENCH[k] : avg(k), r = ref ? x.f[k] / ref : 1;
      return { k, label: FL[k][0], ctl: FL[k][2], val: d1(x.f[k]) + FL[k][1], ref: d1(ref) + FL[k][1] + (BENCH[k] != null ? ' (standar)' : ' (rata² cluster)'), r, st: r >= 1 ? 'ok' : r >= .85 ? 'warn' : 'bad' };
    });
    const worst = x.rows.filter(r => r.ctl && r.k !== 'sf').sort((a, b) => a.r - b.r)[0];   // hanya faktor yang dikontrol sales
    x.weak = worst && worst.r < 1 ? worst : null;
  });

  // 2) GAP IO: target IO – outlook IO, dan berapa yang di lokasi potensial
  const gap = E.map((e, i) => {
    const o = io[i], g = Math.max(0, o.tgt - o.ol), potNeed = Math.round(o.tgt * BENCH.ioPot / 100), potGap = Math.max(0, potNeed - o.olPot);
    const ioPerVis = e.vis ? o.ol / (e.vis / nM) : 0, visNeed = ioPerVis > 0 ? Math.ceil(g / ioPerVis) : null;
    return { name: e.name, tgt: o.tgt, mtd: o.mtd, ol: o.ol, gap: g, pot: o.pot, potShare: Math.round(o.potShare), potNeed, olPot: o.olPot, potGap, visNeed, ps: e.ps, psTgt: e.tgtM, psOl: e.ol,
      txt: (g === 0 ? `Outlook IO ${o.ol} sudah ≥ target ${o.tgt}.` : `Kurang ${g} IO (outlook ${o.ol} dari target ${o.tgt}); ±${visNeed != null ? visNeed : '?'} kunjungan tambahan per bulan pada hasil sekarang.`) +
        ` IO di lokasi potensial ${Math.round(o.potShare)}%` + (o.potShare >= 70 ? ' (sudah ≥70%)' : ' (di bawah acuan 70%)') + (potGap ? `; untuk target, butuh +${potGap} IO lagi dari desa Prioritas/Strategis.` : '.') };
  });

  // 3) MATRIKS POTENSI (port avai) × PENETRASI (PS/avai)
  let V = { by: {} }; try { V = await wlLoadVisits(app); } catch (e) {}
  const kabs = E.flatMap(e => e.kabs);
  // potensi & PS per desa: dari klasifikasi desa (port avai) + order bulan aktif (tidak bergantung data Deepdive yang dimuat terpisah)
  const psDesa = {}, bp = D.bulan.replace('-', '');
  (D.allRows || app.filteredRows).forEach(r => { if (!(r[C.VAL_PS] || '').toString().trim()) return; if (String(ymdNum(r[C.DATE_PS])).slice(0, 6) !== bp) return; const k = wlK(r[C.KEC], r[C.DESA]); psDesa[k] = (psDesa[k] || 0) + 1; });
  const pen = Object.values(katOf).filter(d => kabs.includes(String(d.kab).toUpperCase()) && +d.totalAvail > 0)
    .map(d => ({ kab: d.kab, kec: d.kec, desa: d.desa, avai: +d.totalAvail, ps: psDesa[wlK(d.kec, d.desa)] || 0 }));
  const med = a => { const b = a.slice().sort((x, y) => x - y); return b.length ? b[Math.floor(b.length / 2)] : 0; };
  const mA = med(pen.map(p => +p.avai)), rates = pen.map(p => +p.ps / +p.avai);
  let mP = med(rates); if (!mP) mP = rates.reduce((a, b) => a + b, 0) / Math.max(1, rates.length);
  // Hanya desa Prioritas + Strategis, diberi label sisa potensi (menggantikan matriks 4 kuadran)
  const tc = s2 => String(s2).toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
  const Q = { belum: [], sudah: [], tinjau: [], kurangi: [] };
  pen.forEach(p => {
    const k = wlK(p.kec, p.desa), d = katOf[k] || {}, v = (V.by[k] || {}).n || 0, avai = +p.avai, ps = +p.ps, rate = ps / avai;
    const o = { desa: tc(p.desa), kec: tc(p.kec), kab: String(p.kab).toUpperCase(), avai, ps, pen: Math.round(rate * 100), vis: v, kat: d.kat || '-', jarak: d.nearestDist };
    if (d.kat === 'prioritas' || d.kat === 'strategis') {
      if (avai < mA * 0.5 || (v >= 10 && ps === 0)) Q.tinjau.push(o);
      else if (rate > mP) Q.sudah.push(o); else Q.belum.push(o);
    } else if (v >= 3 && ps === 0) Q.kurangi.push(o);
  });
  Q.belum.sort((a, b) => (b.avai - b.ps) - (a.avai - a.ps)); Q.sudah.sort((a, b) => b.ps - a.ps); Q.tinjau.sort((a, b) => b.vis - a.vis); Q.kurangi.sort((a, b) => b.vis - a.vis);
  const lbl = x => `${x.desa}${x.kat === 'strategis' ? ' ★' : ''} — sisa ${Math.max(0, x.avai - x.ps)} port, PS ${x.ps}, ${x.vis} kunj.`;
  const matrix = [
    { k: 'belum', t: '🔴 Belum tergarap', s: 'Port masih banyak, PS sedikit', c: '#a8231c', act: 'Kejar IO di sini (desa ★ lewat jadwal khusus).' },
    { k: 'sudah', t: '🟢 Sudah tergarap', s: 'Sebagian besar potensi sudah terambil', c: '#3f7d3a', act: 'Cukup dijaga; dorong referral pelanggan.' },
    { k: 'tinjau', t: '⚠️ Perlu ditinjau', s: 'Port hampir habis, atau sering dikunjungi tapi PS 0', c: '#c0622a', act: 'Cek ulang kategori desa & ketersediaan ODP.' },
  ].map(q => ({ ...q, n: Q[q.k].length, top: Q[q.k].slice(0, 6).map(lbl) }));
  const mx = { avai: mA, pen: d1(mP * 100) };

  // 3b) JADWAL DESA STRATEGIS: kurang dikunjungi → kelompok per kecamatan (satu hari kunjungan per kelompok)
  const strat = E.map(e => {
    const ds = Object.values(katOf).filter(d => e.kabs.includes(String(d.kab).toUpperCase()) && d.kat === 'strategis')
      .map(d => ({ desa: d.desa, kec: d.kec, avai: d.totalAvail, jarak: d.nearestDist, vis: ((V.by[wlK(d.kec, d.desa)] || {}).n || 0) / nM }));
    const kurang = ds.filter(d => d.vis < 2);
    const byKec = {}; kurang.forEach(d => (byKec[d.kec] = byKec[d.kec] || []).push(d));
    const grup = Object.entries(byKec).map(([kec, xs]) => ({ kec, desa: xs.sort((a, b) => b.avai - a.avai).map(x => x.desa), avai: xs.reduce((a, x) => a + (x.avai || 0), 0),
      jarak: Math.round(xs.reduce((a, x) => a + (x.jarak || 0), 0) / xs.length) })).sort((a, b) => b.avai - a.avai);
    return { name: e.name, total: ds.length, kurang: kurang.length, grup: grup.slice(0, 4) };
  });

  // 4) DIAGNOSIS → AKSI per kabupaten — bahasa sederhana, satu poin per baris
  const PLAIN = {
    kunj: x => `SF jarang turun ke desa: ${x.val} kunjungan per SF per bulan (rata-rata cluster ${x.ref.split(' ')[0]}).`,
    sasaran: x => `Kunjungan belum ke desa yang tepat: baru ${x.val} ke desa Prioritas/Strategis (acuan 80%).`,
    ioHasil: x => `SF sudah datang, tapi jarang dapat order: tiap 10 kunjungan dapat ${x.val} IO (rata-rata cluster ${x.ref.split(' ')[0]}).`,
    ioPot: x => `Banyak IO datang dari desa Cadangan/Skip: baru ${x.val} dari desa Prioritas/Strategis (acuan 70%).`,
    ioRe: x => `Banyak IO tidak lolos registrasi: IO→RE ${x.val} (acuan 95%).`,
  };
  const diag = E.map((e, i) => {
    const x = drv[i], g = gap[i], st2 = strat[i];
    const garapKab = Q.belum.filter(q => e.kabs.includes(q.kab)).slice(0, 3);
    const kurangKab = Q.kurangi.filter(q => e.kabs.includes(q.kab)).slice(0, 2);
    const perHari = Math.ceil(g.gap / Math.max(1, D.sisa));
    const lines = [
      ['Status IO', g.gap ? `Akhir bulan diperkirakan ${g.ol} IO dari target ${g.tgt} — kurang ${g.gap}.` : `Akhir bulan diperkirakan ${g.ol} IO, sudah ≥ target ${g.tgt}.`],
      ['Masalah utama', (g.gap === 0 && g.potShare >= 70)
        ? 'Tidak ada masalah besar — target IO tercapai & lokasi sudah tepat.' + (x.weak && PLAIN[x.weak.k] ? ' Yang masih bisa ditingkatkan — ' + PLAIN[x.weak.k](x.weak) : '')
        : (x.weak && PLAIN[x.weak.k] ? PLAIN[x.weak.k](x.weak) : 'Tidak ada faktor yang jauh di bawah acuan.')],
      ['Lokasi IO', `${g.potShare}% IO dari desa Prioritas/Strategis` + (g.potShare >= 70 ? ' — sudah sesuai acuan (≥70%).' : ' — masih di bawah acuan 70%.')],
      ['PS (hasil)', `${g.ps} sekarang, diperkirakan ${g.psOl} dari target ${g.psTgt}.`],
    ];
    const acts = [];
    if (g.gap) acts.push(`Cari ${g.gap} IO lagi sampai akhir bulan (±${perHari} per hari), utamakan desa Prioritas/Strategis.`);
    if (x.weak && x.weak.k === 'kunj') acts.push(`Tambah kunjungan SF ke ±${x.weak.ref.split(' ')[0]} per SF per bulan.`);
    if (x.weak && x.weak.k === 'sasaran') acts.push('Arahkan kunjungan ke desa Prioritas/Strategis (minimal 80%).');
    if (x.weak && x.weak.k === 'ioHasil') acts.push(g.gap === 0 ? 'Pertahankan ritme; coaching ringan agar hasil per kunjungan naik.' : 'Coaching cara menawarkan ke pelanggan — kunjungan sudah ada, hasilnya yang rendah.');
    if (x.weak && x.weak.k === 'ioPot' || g.potShare < 70) {
      const sk = {}; (D.allRows || []).forEach(r => { if (!(r[C.VAL_IO] || '').toString().trim()) return; const n = ymdNum(r[C.DATE_IO]); if (n < D.W[0] || n > D.W[1]) return;
        if (!e.kabs.includes(String(r[C.VAL_KAB] || '').toUpperCase())) return; const d = katOf[wlK(r[C.KEC], r[C.DESA])]; if (d && (d.kat === 'skip' || d.kat === 'cadangan')) sk[tc(d.desa)] = (sk[tc(d.desa)] || 0) + 1; });
      const topSk = Object.entries(sk).sort((a, b) => b[1] - a[1]).slice(0, 3);
      acts.push('Tambah IO dari desa Prioritas/Strategis sampai ≥70%' + (topSk.length ? `. IO dari desa Cadangan/Skip tetap boleh — ${topSk.map(([n2, c]) => n2 + ' (' + c + ' IO)').join(', ')} ternyata produktif, jadi tinjau ulang kategorinya, jangan ditinggalkan.` : '.'));
    }
    if (x.weak && x.weak.k === 'ioRe') acts.push('Cek kelengkapan data & tagging sebelum input order.');
    if (st2.grup.length) acts.push(`Jadwalkan hari khusus ke desa Strategis: ` + st2.grup.slice(0, 2).map(gp => `Kec. ${gp.kec} (${gp.desa.slice(0, 3).join(', ')})`).join(', lalu ') + '.');
    if (garapKab.length) acts.push(`Garap desa yang masih banyak sisa port: ${garapKab.map(q => q.desa + (q.kat === 'strategis' ? ' ★' : '')).join(', ')}.`);
    if (kurangKab.length) acts.push(`Kurangi kunjungan ke ${kurangKab.map(q => q.desa).join(' dan ')} — sudah ${kurangKab.length > 1 ? 'masing-masing ' + kurangKab.map(q => q.vis).join(' dan ') : kurangKab[0].vis} kali dikunjungi, PS-nya 0.`);
    if (e.planCur === 0) acts.push(`Publish plan DJP ${periodLabel(D.bulan)} (masih kosong).`);
    if ((e.psre || 0) < 80 && e.coh[2]) acts.push(`(Bantuan teknis) Kawal ${e.coh[2]} order pending bersama teknisi agar jadi PS.`);
    return {
      name: e.name, lines, acts,
      diag: lines.map(l => l[0] + ': ' + l[1]).join(' '),
      status: g.gap === 0 && g.potShare >= 70 ? 'ok' : g.gap <= g.tgt * .1 ? 'warn' : 'bad',
    };
  });
  return { drv, gap, matrix, mx, diag, strat };
}
