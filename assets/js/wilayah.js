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
  const res = await Promise.all(ms.map(m => fetch(`/api/data?type=djp_plan&bulan=${encodeURIComponent(m)}`, auth).then(r => r.json()).catch(() => null)));
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
      weight: fl ? 2.6 : 0.9, dashArray: fl === 'untouched' ? '4 4' : null, fillColor: c,
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
        l.on('mouseover', () => l.setStyle({ weight: 2.6 })).on('mouseout', () => l.setStyle(desaStyle(f)));
      },
    }));
    add(L.geoJSON({ type: 'FeatureCollection', features: kecF }, { style: { pane: 'wlLine', color: '#3b2a1e', weight: 2, dashArray: '8 5', fill: false, opacity: .6 }, interactive: false }));
  } else {
    const kecStyle = f => {
      const s = kecSum[wlN(f.properties.kec)] || { desa: 1, prio: 0, vis: 0 }, pr = s.prio / Math.max(1, s.desa);
      const c = m === 'kab' ? '#cdb68a' : pr > .4 ? '#a8231c' : pr > .2 ? '#d48a00' : '#8c8c8c';
      return { pane: 'wlArea', color: '#4a2c14', weight: m === 'kab' ? 0.6 : 1.6, dashArray: m === 'kab' ? '2 4' : null, fillColor: c, fillOpacity: m === 'kab' ? .12 : .18 + pr * .35 };
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
  if (m !== 'none') add(L.geoJSON({ type: 'FeatureCollection', features: kabF }, { style: { pane: 'wlLine', color: '#8a2a1c', weight: 3, fill: false, opacity: .8 }, interactive: false }));

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
