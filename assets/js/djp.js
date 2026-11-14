// Modul DJP: Plan, Realisasi, Daily, List ODP — dipisah dari index.html (Nov 2026).
// Script klasik (bukan module): konstanta & fungsi di sini global dan dipakai file berikutnya.
// Urutan muat: config.js → app.js → djp.js → Alpine (defer).
// ════════════════════════════════════════════════════════
// DJP – Plan DJP (v2 — multi-desa + multi-SF generate)
// ════════════════════════════════════════════════════════
// Tanggal lokal perangkat (WIB) — toISOString() memakai UTC, jadi sebelum jam 07.00 jatuh ke kemarin
function localYMD(d = new Date()) {
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

// Toast untuk halaman DJP — warna mengikuti tema (bukan biru/indigo)
// Toast halaman DJP memakai komponen bersama (config.js)
function djpToast(msg, type='info', ms) { showToast(msg, type, ms); }

// Kategori desa — label & warna sama dengan menu Prioritas / Timeline
const DJP_KAT = {
  prioritas: { label: 'Prioritas', color: '#ef4444', order: 1 },
  strategis: { label: 'Strategis', color: '#f97316', order: 2 },
  cadangan:  { label: 'Cadangan',  color: '#eab308', order: 3 },
  skip:      { label: 'Skip',      color: '#9ca3af', order: 4 },
};

// Warna penanda SF di peta Review (urut SF dalam draft)
const DJP_SF_COLORS = ['#c93c20', '#2563eb', '#16a34a', '#9333ea', '#d97706', '#0891b2', '#db2777', '#4d7c0f', '#7c3aed', '#0f766e'];

function djpPlanData() {
  // Cache hasil hitung berat (di luar reaktivitas Alpine): dihitung ulang hanya bila kunci berubah
  const _memo = {};
  const memo = (name, key, fn) => {
    const m = _memo[name];
    if (m && m.key === key) return m.val;
    const val = fn();
    _memo[name] = { key, val };
    return val;
  };
  const _today = localYMD();
  const _year  = new Date().getFullYear().toString();
  return {
    // ── Setup state ──
    _odpAll: [],
    _clusterKabs: [],
    sfList: [],
    form: { tgl: _today, kab: '', kantor: '' },
    desaSelected: [],
    sfSelected:   [],
    desaSearch:   '',
    desaSort:     'dekat',
    desaKat:      '',        // filter kategori desa ('' = semua)   // 'dekat' = urut jarak dari desa terpilih · 'az' = abjad
    sfSearch:     '',
    // ── Generate ──
    draftPlan:   [],
    planHistory: {},
    generating:  false,
    // ── Plan tersimpan ──
    viewTgl:     _today,
    viewTahun:   _year,
    planData:    [],
    planLoading: false,
    saving:      false,
    // ── Wizard (tampilan) ──
    step:        1,      // 1 Setup · 2 Generate · 3 Review · 4 Publish
    genStage:    0,      // 0..4 tahap proses generate (untuk progress)
    revSf:       0,      // index SF yang dibuka di Review
    revOdp:      null,   // ODP yang dipilih (detail)
    revView:     'peta', // tab di HP: 'peta' | 'tabel'
    savedOk:     false,
    savedInfo:   null,

    // init() sengaja kosong: data DJP baru dimuat saat menu Plan DJP dibuka (event page-djp_plan).
    // Sebelumnya djp_users (~14 dtk) + odp.json 1,2 MB ikut dimuat di halaman dashboard.
    init() {},

    async initPlan() {
      this._initCluster();
      // Kategori desa (rumus menu Prioritas) dimuat di latar belakang — daftar desa tidak menunggu
      try { if (!this.prioritasOdpRaw.length && !this.prioritasLoading) this.loadPrioritas(); } catch(e) {}
      try {
        await Promise.all([this._loadODP(), this._loadSF()]);
        // Restore draft dari localStorage kalau masih ada (misal: tab di-reload)
        try {
          const savedDraft = localStorage.getItem('djp_draft_' + this.form.tgl);
          if (savedDraft) {
            const { draft, ts } = JSON.parse(savedDraft);
            const ageMin = (Date.now() - ts) / 60000;
            if (Array.isArray(draft) && ageMin < 120) { // draft valid max 2 jam
              this.draftPlan = draft;
              this.step = 3;
              this.$nextTick(() => this.initReviewMap());
              this._toast('📋 Draft sebelumnya dipulihkan', 'info');
            }
          }
        } catch(e) {}
        await this.loadPlan();
      } catch(e) {}
    },

    _initCluster() {
      try {
        const rootEl = document.querySelector('[x-data="app()"]');
        const cl = rootEl?._x_dataStack?.[0]?.cluster;
        this._clusterKabs = (cl && CLUSTER[cl]) ? [...CLUSTER[cl].kab] : [];
      } catch(e) { this._clusterKabs = []; }
    },

    // ── Load ODP (all, no avai filter — filter at generate time) ──
    async _loadODP() {
      if (this._odpAll.length) return;
      const r = await fetchOdpJson();
      const raw = await r.json();
      const data = Array.isArray(raw) ? raw : (raw.data || []);
      this._odpAll = data
        .filter(o => o.odp && o.kab)
        .map(o => ({ ...o, kab: (o.kab||'').toUpperCase() }));
    },

    // ── Load SF (semua aktif di cluster) ──
    async _loadSF() {
      const kabs = this._clusterKabs;
      const SF_CACHE_KEY = 'djp_sf_' + kabs.join('_');
      const SF_TTL_MS = 30 * 60 * 1000; // 30 menit

      // 1. Cek cache dulu — kalau masih fresh, pakai langsung
      try {
        const cached = localStorage.getItem(SF_CACHE_KEY);
        if (cached) {
          const { data, ts } = JSON.parse(cached);
          if (Array.isArray(data) && data.length && Date.now() - ts < SF_TTL_MS) {
            this.sfList = data;
            return; // ← skip GAS, cache masih valid
          }
        }
      } catch(e) {}

      const gasUrl = this._gasUrl();
      const st = localStorage.getItem('wok_token')||'';
      const kabsParam = encodeURIComponent(JSON.stringify(kabs));
      const parseSFArr = (arr, clusterKabs) => {
        const seen = new Set();
        return arr
          .filter(u => {
            if (!Array.isArray(u)) return true;
            if ((u[16]||'').toString().toUpperCase() !== 'ACTIVE') return false;
            if (clusterKabs.length > 0) return clusterKabs.includes((u[20]||'').toString().toUpperCase().trim());
            return true;
          })
          .map(u => Array.isArray(u)
            ? { username: u[2].toString().trim(), nama: u[1].toString().trim(), kab:(u[20]||'').toString().trim() }
            : { username:(u.username||'').toString().trim(), nama:(u.nama||'').toString().trim(), kab:(u.kab||'').toString().trim() })
          .filter(u => { if(!u.username||seen.has(u.username)) return false; seen.add(u.username); return true; });
      };
      try {
        const r = await fetch(`${gasUrl}?type=djp_users&session_token=${encodeURIComponent(st)}&kabs=${kabsParam}`);
        const d = await r.json();
        if (d.ok && d.users && d.users.length) {
          const seen = new Set();
          this.sfList = d.users.filter(u => u.username && !seen.has(u.username) && seen.add(u.username));
          try { localStorage.setItem(SF_CACHE_KEY, JSON.stringify({ data: this.sfList, ts: Date.now() })); } catch(e) {}
          return;
        }
        throw new Error('no users');
      } catch(e) {
        try {
          const r2 = await fetch('data/update_sf.json');
          const sf2 = await r2.json();
          this.sfList = parseSFArr(sf2.data||[], kabs.map(k=>k.toUpperCase()));
        } catch(e2) { this.sfList = []; }
      }
    },

    _gasUrl() {
      const cfg = (typeof Alpine!=='undefined' && Alpine.store) ? Alpine.store('cfg') : null;
      return (cfg && cfg.gasUrl) || DEFAULT_GAS_URL;
    },

    _toast(msg, type='info', ms) { djpToast(msg, type, ms); },

    // ── Getters: kabupaten list (filtered by cluster) ──
    get kabList() {
      const cl = this._clusterKabs;
      const all = [...new Set(this._odpAll.map(o=>o.kab).filter(Boolean))].sort();
      return cl.length ? all.filter(k => cl.includes(k)) : all;
    },

    // ── Getters: kantor referensi (filtered by kab if set) ──
    get kantorList() {
      if (!this.form.kab) return [];
      return (typeof OFFICES !== 'undefined' ? OFFICES : []).filter(o => o.kab === this.form.kab);
    },

    // ── Getters: statistik avail per desa ──
    get desaStats() {
      return memo('stats', this.form.kab + '|' + this._odpAll.length, () => this._desaStats());
    },
    _desaStats() {
      const kab = this.form.kab;
      const stats = {};
      this._odpAll
        .filter(o => (!kab || o.kab === kab) && (o.avai||0) > 0)
        .forEach(o => {
          const d = o.desa||'';
          if (!stats[d]) stats[d] = { avail: 0 };
          stats[d].avail += 1;
        });
      return stats;
    },

    // ── Getters: desa list dengan search ──
    get desaListFiltered() {
      const kab = this.form.kab;
      if (!kab) return [];
      let all = this.desaAllKab;
      const q = this.desaSearch.trim().toLowerCase();
      if (q) all = all.filter(d => d.toLowerCase().includes(q));
      if (this.desaKat) { const km = this.desaKatMap; all = all.filter(d => (km[d]?.kat || '') === this.desaKat); }
      // Mode "Terdekat": desa terpilih di atas, sisanya urut jarak dari titik tengah desa terpilih
      const dist = this.desaDistMap;
      if (this.desaSort === 'dekat' && dist) {
        const sel = this.desaSelected.filter(d => all.includes(d));
        const rest = all.filter(d => !this.desaSelected.includes(d))
          .sort((a, b) => (dist[a] ?? 1e9) - (dist[b] ?? 1e9));
        return [...sel, ...rest];
      }
      return all;
    },

    // Kategori desa (Prioritas/Strategis/Cadangan/Skip) — rumus menu Prioritas, per cluster
    get desaKatMap() {
      const raw = this.prioritasOdpRaw || [], pen = this.penetrasiData || [];
      const kab = this.form.kab;
      return memo('kat', this.cluster + '|' + kab + '|' + raw.length + '|' + pen.length, () => {
        const out = {};
        if (!raw.length || !kab) return out;
        this.klasifikasiDesa('', '').forEach(d => {
          if ((d.kab || '').toUpperCase() !== kab) return;
          const prev = out[d.desa];
          if (!prev || DJP_KAT[d.kat].order < DJP_KAT[prev.kat].order) out[d.desa] = { kat: d.kat, skor: d.skorPotensi };
        });
        return out;
      });
    },
    // Semua desa di kab terpilih yang punya ODP avai > 0 (urut abjad)
    get desaAllKab() {
      const kab = this.form.kab;
      return memo('desaAll', kab + '|' + this._odpAll.length, () => [...new Set(
        this._odpAll.filter(o => o.kab === kab && (o.avai||0) > 0 && o.desa).map(o => o.desa)
      )].sort());
    },
    get desaKatReady() { return Object.keys(this.desaKatMap).length > 0; },
    get desaKatCounts() {
      const km = this.desaKatMap, all = this.desaAllKab;
      const c = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0 };
      all.forEach(d => { const k = km[d]?.kat; if (k) c[k]++; });
      return c;
    },

    // Titik tengah ODP per desa (kab terpilih)
    get desaCentroid() {
      return memo('centroid', this.form.kab + '|' + this._odpAll.length, () => this._desaCentroid());
    },
    _desaCentroid() {
      const kab = this.form.kab, acc = {};
      if (!kab) return acc;
      this._odpAll.forEach(o => {
        if (o.kab !== kab || !o.desa) return;
        const la = parseFloat(o.lat), ln = parseFloat(o.lng);
        if (!la || !ln) return;
        const a = acc[o.desa] || (acc[o.desa] = { la: 0, ln: 0, n: 0 });
        a.la += la; a.ln += ln; a.n++;
      });
      const out = {};
      Object.entries(acc).forEach(([d, a]) => { out[d] = { lat: a.la / a.n, lng: a.ln / a.n }; });
      return out;
    },

    // Jarak (km) tiap desa ke titik tengah desa-desa terpilih; null bila belum ada yang dipilih
    get desaDistMap() {
      return memo('dist', this.form.kab + '|' + this._odpAll.length + '|' + this.desaSelected.join('\u0001'), () => this._desaDistMap());
    },
    _desaDistMap() {
      const c = this.desaCentroid;
      const pts = this.desaSelected.map(d => c[d]).filter(Boolean);
      if (!pts.length) return null;
      const lat = pts.reduce((s, p) => s + p.lat, 0) / pts.length;
      const lng = pts.reduce((s, p) => s + p.lng, 0) / pts.length;
      const out = {};
      Object.entries(c).forEach(([d, p]) => { out[d] = haversineKm(lat, lng, p.lat, p.lng); });
      return out;
    },
    get desaJarakMap() {
      return memo('label', this.form.kab + '|' + this._odpAll.length + '|' + this.desaSelected.join('\u0001'), () => {
        const m = this.desaDistMap, out = {};
        if (!m) return out;
        const sel = new Set(this.desaSelected);
        Object.entries(m).forEach(([d, km]) => {
          if (!sel.has(d)) out[d] = (km < 10 ? km.toFixed(1).replace('.', ',') : Math.round(km)) + ' km';
        });
        return out;
      });
    },
    desaJarakLabel(d) { return this.desaJarakMap[d] || ''; },

    // ── Getters: SF list dengan search ──
    get sfListFiltered() {
      const q = this.sfSearch.trim().toLowerCase();
      return q ? this.sfList.filter(u => u.nama.toLowerCase().includes(q) || u.username.toLowerCase().includes(q))
               : this.sfList;
    },

    // ── Getters: canGenerate ──
    get canGenerate() {
      return !this.generating && this.form.tgl && this.form.kab && this.desaSelected.length > 0 && this.sfSelected.length > 0;
    },

    // ── Toggle / select helpers ──
    toggleDesa(d) {
      const i = this.desaSelected.indexOf(d);
      if (i >= 0) this.desaSelected.splice(i, 1);
      else this.desaSelected.push(d);
    },
    toggleSF(u) {
      const i = this.sfSelected.indexOf(u);
      if (i >= 0) this.sfSelected.splice(i, 1);
      else this.sfSelected.push(u);
    },
    selectAllDesa() { this.desaSelected = [...this.desaListFiltered]; },
    clearDesa()     { this.desaSelected = []; },
    selectAllSF()   { this.sfSelected = this.sfListFiltered.map(u => u.username); },
    clearSF()       { this.sfSelected = []; },

    // ── Remove ODP from draft ──
    removeDraftOdp(sfIdx, odpIdx) {
      this.draftPlan[sfIdx].odps.splice(odpIdx, 1);
    },

    // ── Fetch history (last plan date per ODP) ──
    async _fetchHistory() {
      try {
        const gasUrl = this._gasUrl();
        const st = localStorage.getItem('wok_token')||'';
        const desas = JSON.stringify(this.desaSelected);
        const url = `${gasUrl}?type=djp_plan_history&session_token=${encodeURIComponent(st)}&kab=${encodeURIComponent(this.form.kab)}&desas=${encodeURIComponent(desas)}`;
        const r = await fetch(url);
        const d = await r.json();
        if (d.ok && d.history) return d.history;
        throw new Error(d.error || 'tidak ada data');
      } catch(e) {
        this._toast('⚠️ Riwayat rotasi ODP gagal dimuat (' + (e.message||'koneksi') + '). Draft dibuat tanpa prioritas rotasi.', 'error', 7000);
        return {};
      }
    },

    // ── Main generate logic ──
    async generateDraft() {
      if (!this.canGenerate) return;
      this.generating = true;
      this.draftPlan = [];
      this.genStage = 0;
      try {
        // 1. Build ODP pool: hanya desa yang dipilih, avail > 0
        const pool = this._odpAll.filter(o =>
          o.kab === this.form.kab &&
          this.desaSelected.includes(o.desa||'') &&
          (o.avai||0) > 0
        );

        if (pool.length === 0) {
          this._toast('Tidak ada ODP tersedia di desa yang dipilih', 'error');
          this.generating = false;
          return;
        }

        this.genStage = 1;
        // 2. Fetch rotation history
        const history = await this._fetchHistory();
        this.planHistory = history;
        this.genStage = 2;

        // 3. Sort pool: (a) never planned first, (b) oldest last-plan asc, (c) avai desc, (d) kapasitas desc
        pool.sort((a, b) => {
          const ha = history[a.odp] || '';
          const hb = history[b.odp] || '';
          // never planned = '' → goes first
          if (!ha && hb) return -1;
          if (ha && !hb) return 1;
          if (ha !== hb) return ha < hb ? -1 : 1;  // oldest first
          if ((b.avai||0) !== (a.avai||0)) return (b.avai||0) - (a.avai||0);
          return (b.kapasitas||0) - (a.kapasitas||0);
        });

        // 4. Take top N = sfSelected.length * 4
        const maxOdp = this.sfSelected.length * 4;
        const chosen = pool.slice(0, maxOdp);

        // 5. Urutan geografis ODP terpilih:
        //    - Kantor Referensi dipilih → urut jarak dari kantor (terdekat dulu), sesuai aturan
        //      "sort by jarak ke kantor → block-assign 4 ODP/SF"
        //    - Tanpa kantor → urut lintang lalu bujur (perilaku lama)
        const kantor = (typeof OFFICES !== 'undefined' ? OFFICES : []).find(o => o.name === this.form.kantor);
        if (kantor) {
          const dist = o => haversineKm(kantor.lat, kantor.lng, parseFloat(o.lat)||0, parseFloat(o.lng)||0);
          chosen.forEach(o => { o._jarakKm = Math.round(dist(o) * 10) / 10; });
          chosen.sort((a, b) => a._jarakKm - b._jarakKm);
        } else {
          chosen.sort((a, b) => {
            const latA = parseFloat(a.lat)||0, latB = parseFloat(b.lat)||0;
            if (Math.abs(latA - latB) > 0.0001) return latA - latB;
            return (parseFloat(a.lng)||0) - (parseFloat(b.lng)||0);
          });
        }

        this.genStage = 3;
        // 6. Block-assign to SFs
        const sfObjs = this.sfSelected.map(u => this.sfList.find(s => s.username === u)).filter(Boolean);
        const draft = sfObjs.map((sf, i) => ({
          username_sf: sf.username,
          nama_sf: sf.nama,
          odps: chosen.slice(i * 4, (i + 1) * 4).map(o => ({ ...o })),
        }));

        this.draftPlan = draft;
        this.genStage = 4;
        // Simpan draft ke localStorage — persist kalau reload
        try {
          localStorage.setItem('djp_draft_' + this.form.tgl, JSON.stringify({
            tgl: this.form.tgl, kab: this.form.kab, draft, ts: Date.now()
          }));
        } catch(e) {}
        const total = draft.reduce((s, sf) => s + sf.odps.length, 0);
        const kurang = draft.filter(sf => sf.odps.length < 4).map(sf => sf.nama_sf);
        if (kurang.length) {
          this._toast(`⚠️ ODP tersedia hanya ${pool.length}. ${kurang.length} SF dapat kurang dari 4 ODP: ${kurang.slice(0,3).join(', ')}${kurang.length>3?'…':''}. Tambah desa bila perlu.`, 'error', 7000);
        } else {
          this._toast(`✅ Draft dibuat: ${total} ODP untuk ${sfObjs.length} SF`, 'success');
        }
      } catch(e) {
        this._toast('❌ ' + e.message, 'error');
      }
      this.generating = false;
    },

    // ════════ WIZARD ════════
    get draftTotal()   { return this.draftPlan.reduce((s, sf) => s + sf.odps.length, 0); },
    get draftActive()  { return this.draftPlan.filter(sf => sf.odps.length > 0); },
    get draftBalanced(){ return this.draftPlan.length > 0 && this.draftPlan.every(sf => sf.odps.length === 4); },
    get draftDesa()    { return [...new Set(this.draftPlan.flatMap(sf => sf.odps.map(o => o.desa)).filter(Boolean))]; },
    get revSfObj()     { return this.draftPlan[this.revSf] || null; },
    get selectedSfObjs(){ return this.sfSelected.map(u => this.sfList.find(s => s.username === u)).filter(Boolean); },
    sfColor(i) { return DJP_SF_COLORS[i % DJP_SF_COLORS.length]; },
    fmtTglPanjang(t) {
      if (!t) return '—';
      return new Date(t + 'T00:00:00').toLocaleDateString('id-ID', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
    },
    odpAlasan(o) {
      const h = this.planHistory[o.odp];
      return h ? 'Terakhir diplan ' + h + ' — paling lama tidak dikunjungi di pool' : 'Belum pernah masuk rencana kunjungan';
    },

    goStep(n) {
      if (n >= 3 && !this.draftPlan.length) return;
      this.step = n;
      window.scrollTo({ top: 0, behavior: 'smooth' });
      if (n === 3) this.$nextTick(() => this.initReviewMap());
    },
    async startGenerate() {
      if (!this.canGenerate) return;
      this.step = 2;
      this.savedOk = false;
      const t0 = Date.now();
      await this.generateDraft();
      if (!this.draftTotal) { this.step = 1; return; }
      // jeda singkat agar progress terbaca, lalu ke Review
      await new Promise(r => setTimeout(r, Math.max(0, 700 - (Date.now() - t0))));
      this.revSf = 0; this.revOdp = null;
      this.goStep(3);
    },
    async publish() {
      await this.saveAllDraft();
      if (this.savedOk) window.scrollTo({ top: 0, behavior: 'smooth' });
    },
    newPlan() {
      this.savedOk = false; this.savedInfo = null;
      this.draftPlan = []; this.desaSelected = []; this.sfSelected = [];
      this.step = 1;
    },
    removeOdpReview(odpIdx) {
      const sf = this.revSfObj; if (!sf) return;
      const o = sf.odps[odpIdx];
      sf.odps.splice(odpIdx, 1);
      if (this.revOdp && o && this.revOdp.odp === o.odp) this.revOdp = null;
      try {
        localStorage.setItem('djp_draft_' + this.form.tgl, JSON.stringify({ tgl: this.form.tgl, kab: this.form.kab, draft: this.draftPlan, ts: Date.now() }));
      } catch(e) {}
      this.renderReviewMap();
    },
    pickSf(i) { this.revSf = i; this.revOdp = null; this.renderReviewMap(true); },
    pickOdp(o) {
      this.revOdp = o;
      const ctx = window._djpRev; if (!ctx || !ctx.map) return;
      const la = parseFloat(o.lat), ln = parseFloat(o.lng);
      if (la && ln) ctx.map.setView([la, ln], Math.max(ctx.map.getZoom(), 15), { animate: true });
    },

    // Peta Review — objek Leaflet disimpan di luar proxy Alpine
    initReviewMap() {
      const el = document.getElementById('djp-rev-map');
      if (!el || typeof L === 'undefined') return;
      const ctx = window._djpRev || (window._djpRev = {});
      if (ctx.map && ctx.el !== el) { try { ctx.map.remove(); } catch(e) {} ctx.map = null; }
      if (!ctx.map) {
        ctx.el = el;
        ctx.map = L.map(el, { zoomControl: true, attributionControl: false });
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19 }).addTo(ctx.map);
        ctx.layer = L.layerGroup().addTo(ctx.map);
      }
      setTimeout(() => { ctx.map.invalidateSize(); this.renderReviewMap(true); }, 60);
    },
    renderReviewMap(fit = false) {
      const ctx = window._djpRev; if (!ctx || !ctx.map) return;
      ctx.layer.clearLayers();
      const pts = [], focus = [];
      this.draftPlan.forEach((sf, i) => {
        const on = i === this.revSf, c = this.sfColor(i);
        sf.odps.forEach((o, j) => {
          const la = parseFloat(o.lat), ln = parseFloat(o.lng);
          if (!la || !ln) return;
          pts.push([la, ln]); if (on) focus.push([la, ln]);
          const m = L.marker([la, ln], {
            icon: L.divIcon({
              className: 'djp-pin-wrap',
              html: `<span class="djp-pin${on ? ' on' : ''}" style="--c:${c}">${on ? j + 1 : ''}</span>`,
              iconSize: on ? [26, 26] : [12, 12], iconAnchor: on ? [13, 13] : [6, 6],
            }),
            zIndexOffset: on ? 1000 : 0,
          }).addTo(ctx.layer);
          m.bindTooltip(`<b>${esc(o.odp)}</b><br>${esc(sf.nama_sf)} · ${esc(o.desa || '')}`, { direction: 'top' });
          m.on('click', () => { if (i !== this.revSf) this.pickSf(i); this.revOdp = this.draftPlan[i].odps[j]; });
        });
      });
      const k = (typeof OFFICES !== 'undefined' ? OFFICES : []).find(o => o.name === this.form.kantor);
      if (k) {
        L.marker([k.lat, k.lng], { icon: L.divIcon({ className: 'djp-pin-wrap', html: '<span class="djp-kantor">🏢</span>', iconSize: [26, 26], iconAnchor: [13, 13] }) })
          .bindTooltip(esc(k.name), { direction: 'top' }).addTo(ctx.layer);
      }
      if (fit) {
        const b = focus.length ? focus : pts;
        if (b.length) ctx.map.fitBounds(b, { padding: [36, 36], maxZoom: 15 });
      }
    },

    // ── Save all draft (per SF) ──
    async saveAllDraft() {
      if (this.saving) return;
      const activeSFs = this.draftPlan.filter(sf => sf.odps.length > 0);
      if (!activeSFs.length) { this._toast('Tidak ada ODP untuk disimpan', 'error'); return; }
      this.saving = true;
      const gasUrl = this._gasUrl();
      const st = localStorage.getItem('wok_token')||'';
      const tahun = (this.form.tgl||'').slice(0,4) || this.viewTahun;

      const allRows = [];
      for (const sf of activeSFs) {
        for (const o of sf.odps) {
          allRows.push({
            tgl: this.form.tgl,
            username_sf: sf.username_sf,
            nama_sf: sf.nama_sf,
            odp: o.odp,
            cluster: clusterOfKab(o.kab) || o.cluster || '',
            kab: o.kab||'',
            kec: o.kec||'',
            desa: o.desa||'',
            lat: o.lat||'',
            lng: o.lng||'',
            avai: o.avai||0,
            tahun,
          });
        }
      }

      // ── SIMPAN JUJUR: UI baru bilang "tersimpan" setelah server mengonfirmasi ──
      // Draft TIDAK dihapus sampai GAS menjawab ok, jadi simpan yang gagal tidak menghilangkan plan.
      const tglSimpan = this.form.tgl;
      this._toast('⏳ Menyimpan ' + allRows.length + ' ODP ke server…', 'info');
      try {
        const r = await fetch(gasUrl, {
          method: 'POST',
          body: JSON.stringify({ type:'djp_plan_save', session_token:st, tgl:tglSimpan, rows: allRows }),
        });
        const d = await r.json();
        if (!d.ok) throw new Error(d.error || 'Server menolak simpan');

        // Sukses → baru hapus draft, lalu muat ulang isi sheet yang sebenarnya
        try { localStorage.removeItem('djp_draft_' + tglSimpan); } catch(e) {}
        try { localStorage.removeItem('djp_plandata_' + tglSimpan); } catch(e) {}
        this.draftPlan = [];
        // Server sudah mengganti plan SF-SF ini pada tanggal ini → susun ulang tampilan secara lokal
        const sfSaved = new Set(allRows.map(r => r.username_sf));
        const sebelumnya = (this.viewTgl === tglSimpan ? (this.planData || []) : [])
          .filter(r => !sfSaved.has(r.username_sf || r[1]));
        this.viewTgl = tglSimpan;
        this.planData = [...sebelumnya, ...allRows.map(r => ({ ...r }))];
        try { localStorage.setItem('djp_plandata_' + tglSimpan, JSON.stringify({ data: this.planData, ts: Date.now() })); } catch(e) {}
        this.savedOk = true;
        this.savedInfo = { tgl: tglSimpan, odp: d.saved || allRows.length, sf: activeSFs.length };
        this._toast(`✅ ${d.saved || allRows.length} ODP tersimpan untuk ${activeSFs.length} SF`, 'success');
      } catch(e) {
        // Gagal → draft tetap ada di layar dan di localStorage, user bisa klik Simpan lagi
        this._toast('❌ Gagal menyimpan: ' + (e.message || 'koneksi terputus') + '. Draft masih ada, coba Simpan lagi.', 'error', 8000);
      } finally {
        this.saving = false;
      }
    },

    // ── Load plan tersimpan ──
    async loadPlan(forceRefresh = false) {
      const cacheKey = 'djp_plandata_' + this.viewTgl;
      const PLAN_TTL_MS = 5 * 60 * 1000; // 5 menit — cukup untuk navigasi bolak-balik
      let cacheAge = Infinity;

      // 1. Cek cache localStorage — tampil instan kalau ada
      try {
        const cached = localStorage.getItem(cacheKey);
        if (cached) {
          const { data, ts } = JSON.parse(cached);
          if (Array.isArray(data)) {
            this.planData = data;
            this.planLoading = false;
            cacheAge = Date.now() - ts;
          }
        }
      } catch(e) {}

      // 2. Skip GAS kalau cache masih fresh (< 5 menit) dan tidak force-refresh
      if (!forceRefresh && cacheAge < PLAN_TTL_MS) return;

      // 3. Fetch GAS di background — update cache & UI kalau ada perubahan
      try {
        this.planLoading = this.planData.length === 0; // spinner hanya kalau belum ada data
        const gasUrl = this._gasUrl();
        const st = localStorage.getItem('wok_token')||'';
        const r = await fetch(`${gasUrl}?type=djp_plan_get&session_token=${encodeURIComponent(st)}&tgl=${this.viewTgl}&tahun=${this.viewTahun}`);
        const d = await r.json();
        if (!d.ok) throw new Error(d.error||'Gagal load plan');
        this.planData = d.data||[];
        // Update cache dengan data terbaru dari GAS
        try {
          localStorage.setItem(cacheKey, JSON.stringify({ data: this.planData, ts: Date.now() }));
        } catch(e) {}
      } catch(e) {
        if (this.planData.length === 0) this.planData = [];
      }
      this.planLoading = false;
    },

    async deletePlanSF(username_sf) {
      const sfNama = (this.planBySF.find(x => x.username_sf === username_sf) || {}).nama_sf || username_sf;
      const ok = await uiConfirm({
        title: 'Hapus plan?',
        message: `Semua ODP plan ${sfNama} pada ${this.viewTgl} akan dihapus dari sheet DJP_Plan. Tindakan ini tidak bisa dibatalkan dari aplikasi.`,
        confirmText: 'Hapus plan', danger: true,
      });
      if (!ok) return;
      const gasUrl = this._gasUrl();
      const st = localStorage.getItem('wok_token')||'';
      try {
        const r = await fetch(gasUrl, { method:'POST', body: JSON.stringify({ type:'djp_plan_delete', session_token:st, tgl:this.viewTgl, username_sf }) });
        const d = await r.json();
        if (!d.ok) throw new Error(d.error||'Gagal hapus');
        // Hapus dari tampilan lokal (server sudah menghapus) — tanpa muat ulang dari GAS
        this.planData = (this.planData || []).filter(r => (r.username_sf || r[1]) !== username_sf);
        try { localStorage.setItem('djp_plandata_' + this.viewTgl, JSON.stringify({ data: this.planData, ts: Date.now() })); } catch(e) {}
        this._toast(`🗑 Plan ${username_sf} dihapus (${d.deleted} baris)`, 'info');
      } catch(e) { this._toast('❌ ' + e.message, 'error', 6000); }
    },

    get planBySF() {
      const map = {};
      (this.planData||[]).forEach(row => {
        const sf = row.username_sf || row[1] || '';
        if (!sf) return;
        if (!map[sf]) map[sf] = { username_sf: sf, nama_sf: row.nama_sf || row[2] || sf, odps: [] };
        map[sf].odps.push({
          odp:     row.odp     || row[3] || '',
          cluster: row.cluster || row[4] || '',
          kab:     row.kab     || row[5] || '',
          kec:     row.kec     || row[6] || '',
          desa:    row.desa    || row[7] || '',
          avai:    row.avai    || 0,
          visit:   row.visit   || false,
        });
      });
      return Object.values(map);
    },
  };
}

// ════════════════════════════════════════════════════════
// DJP – Realisasi DJP (Upload CSV)
// ════════════════════════════════════════════════════════
function djpRealData() {
  return {
    csvRows: [],
    csvPreview: [],
    fileName: '',
    uploading: false,
    msg: '',
    SHEET_URL: 'https://docs.google.com/spreadsheets/d/1ioPvgGbRZIHb_zP-ushATvz434b3p1LQvGcpn1OxBIQ/edit',

    init() {
      window.addEventListener('page-djp_real', () => { this.csvRows=[]; this.csvPreview=[]; this.msg=''; this.fileName=''; });
    },

    onDrop(e) {
      e.preventDefault();
      const f = e.dataTransfer.files[0];
      if (f) this.readFile(f);
    },

    onFileChange(e) {
      const f = e.target.files[0];
      if (f) this.readFile(f);
    },

    readFile(file) {
      this.fileName = file.name;
      this.msg = '';
      const reader = new FileReader();
      reader.onload = (ev) => {
        const text = ev.target.result;
        const lines = text.split(/\r?\n/).filter(l=>l.trim());
        if (!lines.length) { this.msg='❌ File kosong'; return; }
        const sep = lines[0].includes('\t') ? '\t' : ',';
        const headers = lines[0].split(sep).map(h=>h.trim().replace(/^"|"$/g,''));
        const rows = lines.slice(1).map(l => {
          const cols = l.split(sep).map(c=>c.trim().replace(/^"|"$/g,''));
          const obj = {};
          headers.forEach((h,i) => obj[h] = cols[i]||'');
          return obj;
        }).filter(r => Object.values(r).some(v=>v));
        this.csvRows = rows;
        this.csvPreview = rows.slice(0,5);
        this.msg = `📋 ${rows.length} baris terdeteksi dari ${file.name}`;
      };
      reader.readAsText(file, 'UTF-8');
    },

    get csvHeaders() {
      if (!this.csvRows.length) return [];
      return Object.keys(this.csvRows[0]);
    },

    async upload() {
      if (!this.csvRows.length) { this.msg='❌ Belum ada file CSV'; return; }
      const cfg = (typeof Alpine!=='undefined' && Alpine.store) ? Alpine.store('cfg') : null;
      const gasUrl = (cfg && cfg.gasUrl) || DEFAULT_GAS_URL;
      const st = localStorage.getItem('wok_token')||'';
      const uploader = (JSON.parse(localStorage.getItem('wok_user')||'null')||{}).username||'';
      this.uploading = true;
      this.msg = 'Mengupload...';
      try {
        const BATCH = 500;
        let inserted = 0, skipped = 0;
        for (let i = 0; i < this.csvRows.length; i += BATCH) {
          const batch = this.csvRows.slice(i, i + BATCH);
          const r = await fetch(gasUrl, { method:'POST', body: JSON.stringify({ type:'djp_realisasi_upload', session_token:st, uploaded_by:uploader, rows:batch }) });
          const d = await r.json();
          if (!d.ok) throw new Error(d.error||'Upload gagal');
          inserted += d.saved||d.inserted||0;
          skipped  += d.skipped||0;
          this.msg = `Mengupload... ${Math.min(i+BATCH, this.csvRows.length)}/${this.csvRows.length}`;
        }
        this.msg = `✅ Upload selesai — ${inserted} baris baru, ${skipped} duplikat dilewati`;
        this.csvRows = []; this.csvPreview = []; this.fileName = '';
      } catch(e) { this.msg = '❌ ' + e.message; }
      this.uploading = false;
    },
  };
}

// ════════════════════════════════════════════════════════
// DJP – DJP Daily (Monitor Harian)
// ════════════════════════════════════════════════════════
function djpDailyData() {
  return {
    dailyTgl: localYMD(),
    dailyData: null,   // { total_sf, total_plan, total_real, summary: [{username_sf,nama_sf,odp_list:[{odp,kec,desa,status,checkin,odp_condition,prospect}]}] }
    loadingDaily: false,

    init() {}, // dimuat lewat @page-djp_daily saat menu dibuka

    async loadDaily() {
      const cfg = (typeof Alpine!=='undefined' && Alpine.store) ? Alpine.store('cfg') : null;
      const gasUrl = (cfg && cfg.gasUrl) || DEFAULT_GAS_URL;
      const st = localStorage.getItem('wok_token')||'';
      this.loadingDaily = true;
      this.dailyData = null;
      try {
        const url = `${gasUrl}?type=djp_daily&session_token=${encodeURIComponent(st)}&tgl=${this.dailyTgl}`;
        const r = await fetch(url);
        const d = await r.json();
        if (!d.ok) throw new Error(d.error||'Gagal load');
        // GAS returns { ok, total_sf, total_plan, total_real, rows:[{nama_sf,username_sf,odp,checkin_jam,realisasi}], summary }
        // Fallback: jika rows tidak ada (GAS belum redeploy), build dari summary.odp_list
        let rows = d.rows || [];
        if (!rows.length && (d.summary || []).length) {
          rows = (d.summary || []).flatMap(sf =>
            (sf.odp_list || sf.plans || []).map(p => ({
              nama_sf:     sf.nama_sf     || sf.username_sf || '',
              username_sf: sf.username_sf || '',
              odp:         p.odp         || '',
              checkin_jam: p.checkin_jam  || '',
              realisasi:   (p.realisasi === 'Done' || p.status === 'Done') ? 'Done' : '',
            }))
          );
        }
        this.dailyData = {
          total_sf:   d.total_sf   || 0,
          total_plan: d.total_plan || 0,
          total_real: d.total_real || 0,
          rows,
          summary:    d.summary    || [],
        };
      } catch(e) {
        this.dailyData = { total_sf:0, total_plan:0, total_real:0, rows:[], summary:[], error: e.message || 'Gagal memuat' };
        if (typeof djpToast === 'function') djpToast('❌ DJP Daily gagal dimuat: ' + (e.message||'koneksi'), 'error', 6000);
      }
      this.loadingDaily = false;
    },
  };
}

// ════════════════════════════════════════════════════════
// DJP – List ODP (Fix #5)
// ════════════════════════════════════════════════════════
function djpOdpListData() {
  const todayYM = () => localYMD().slice(0,7); // yyyy-MM (lokal)
  return {
    bulan: todayYM(),
    loading: false,
    odpData: [],   // [{odp, occ, avail, kab, kec, desa, assign, visit}]

    init() {}, // dimuat lewat @page-djp_odp_list saat menu dibuka

    async loadOdpList() {
      const cfg = (typeof Alpine!=='undefined' && Alpine.store) ? Alpine.store('cfg') : null;
      const gasUrl = (cfg && cfg.gasUrl) || DEFAULT_GAS_URL;
      const st = localStorage.getItem('wok_token')||'';

      // ambil kabs dari cluster user untuk filter lokal
      const rootEl = document.querySelector('[x-data="app()"]');
      const rootCluster = rootEl?._x_dataStack?.[0]?.cluster;
      const clusterKabs = (rootCluster && typeof CLUSTER!=='undefined' && CLUSTER[rootCluster])
        ? CLUSTER[rootCluster].kab.map(k => k.toUpperCase()) : [];

      this.loading = true;
      this.odpData = [];
      try {
        // 1. ODP base data dari data/odp.json (sudah di-export script, cepat)
        const [jsonRes, gasRes] = await Promise.all([
          fetchOdpJson(),
          fetch(`${gasUrl}?type=djp_odp_list&session_token=${encodeURIComponent(st)}&bulan=${this.bulan}`),
        ]);
        const rawJson = await jsonRes.json();
        // odp.json bisa { updated, total, data:[...] } atau langsung array
        const allOdp = Array.isArray(rawJson) ? rawJson : (rawJson.data || []);

        let assignSet = new Set(), visitSet = new Set();
        try {
          const gasData = await gasRes.json();   // { ok, assign:[], visit:[] }
          assignSet = new Set((gasData.assign || []).map(s => s.trim()));
          visitSet  = new Set((gasData.visit  || []).map(s => s.trim()));
        } catch(e) {}

        // 2. Filter by cluster kabs + merge assign/visit
        const rows = allOdp
          .filter(o => {
            if (!clusterKabs.length) return true;
            const kab = (o.KAB || o.kab || '').toString().toUpperCase().trim();
            return clusterKabs.some(k => kab.includes(k) || k.includes(kab));
          })
          .map(o => {
            const name = (o.ODP || o.odp || '').toString().trim();
            return {
              odp:       name,
              occ:       o['OCC 2'] ?? o.occ ?? '',
              avail:     o.AVAI  ?? o.avai  ?? '',
              kab:       o.KAB   || o.kab   || '',
              kec:       o.KEC   || o.kec   || '',
              desa:      o.DESA  || o.desa  || '',
              tahun_odp: o.tahun_odp || '',
              assign:    assignSet.has(name),
              visit:     visitSet.has(name),
            };
          });

        this.odpData = rows;
      } catch(e) { this.odpData = []; console.error('djpOdpList:', e); if (typeof djpToast === 'function') djpToast('❌ List ODP gagal dimuat: ' + (e.message||'koneksi'), 'error', 6000); }
      this.loading = false;
    },
  };
}
