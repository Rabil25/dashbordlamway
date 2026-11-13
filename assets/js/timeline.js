// ================================================================
//  Timeline Desa — arah plan DJP per desa sepanjang bulan berjalan.
//  Sumber plan  : /api/data?type=djp_plan (sheet DJP_Plan, read-only)
//  Kategori desa: klasifikasiDesa() milik menu Prioritas, dihitung per
//                 CLUSTER tanpa filter kabupaten agar warna desa konsisten.
//  Komponen nested Alpine: properti induk (cluster, prioritasOdpRaw,
//  klasifikasiDesa, loadPrioritas, sessionToken) diakses lewat scope Alpine.
// ================================================================

const TL_KAT = {
  prioritas: { label: 'Prioritas', color: '#ef4444', order: 1 },
  strategis: { label: 'Strategis', color: '#f97316', order: 2 },
  cadangan:  { label: 'Cadangan',  color: '#eab308', order: 3 },
  skip:      { label: 'Skip',      color: '#9ca3af', order: 4 },
  lain:      { label: 'Tanpa data', color: '#cbd5e1', order: 5 },
};
const TL_HARI = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];

function tlKey(kec, desa) {
  return (kec || '').trim().toUpperCase() + '||' + (desa || '').trim().toUpperCase();
}

function timelineData() {
  return {
    tlPeriode: currentPeriod(),
    tlRows: [],
    tlLoading: false,
    tlError: '',
    tlLoaded: false,
    tlSel: '',          // 'yyyy-mm-dd' tanggal terpilih
    tlKab: '',
    tlSf: '',
    tlKat: '',          // filter kategori
    tlOpen: {},         // desa yang dibuka daftar ODP-nya

    async tlLoad(force = false) {
      if (this.tlLoading) return;
      if (this.tlLoaded && !force) return;
      this.tlLoading = true;
      this.tlError = '';
      try {
        const tasks = [
          fetch(`/api/data?type=djp_plan&bulan=${encodeURIComponent(this.tlPeriode)}` + (force ? `&_=${Date.now()}` : ''),
                { headers: { Authorization: 'Bearer ' + this.sessionToken } })
            .then(async r => {
              const d = await r.json().catch(() => ({}));
              if (!r.ok || !d.ok) throw new Error(d.error || ('HTTP ' + r.status));
              return d.data || [];
            }),
        ];
        // Kategori desa butuh data ODP + PS (sama dengan menu Prioritas)
        if (!this.prioritasOdpRaw.length) tasks.push(this.loadPrioritas());
        const [rows] = await Promise.all(tasks);
        this.tlRows = rows;
        this.tlLoaded = true;
        this.tlPickDefault();
      } catch (e) {
        this.tlError = e.message || String(e);
      } finally {
        this.tlLoading = false;
      }
    },

    // Tanggal awal: hari ini bila ada plan, jika tidak tanggal plan terakhir
    tlPickDefault() {
      const today = localYMD(new Date());
      const days = this.tlDays.filter(d => d.desas.length);
      if (days.some(d => d.tgl === today)) this.tlSel = today;
      else if (days.length) this.tlSel = days[days.length - 1].tgl;
      else this.tlSel = today;
      this.$nextTick(() => {
        const el = document.querySelector('.tl-day.is-sel');
        if (el) el.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
      });
    },

    // Peta kategori per desa (per cluster, tanpa filter kab/kategori)
    get tlKatMap() {
      const m = {};
      if (!this.prioritasOdpRaw.length) return m;
      this.klasifikasiDesa('', '').forEach(d => {
        m[tlKey(d.kec, d.desa)] = { kat: d.kat, skor: d.skorPotensi, kab: d.kab, kec: d.kec, desa: d.desa, avail: d.totalAvail };
      });
      return m;
    },

    get tlClusterRows() {
      const kabs = (CLUSTER[this.cluster]?.kab || []).map(k => k.toUpperCase());
      return this.tlRows.filter(r =>
        kabs.includes((r.kab || '').toUpperCase()) &&
        (!this.tlKab || r.kab === this.tlKab.toUpperCase()) &&
        (!this.tlSf || r.username_sf === this.tlSf));
    },

    get tlSfList() {
      const kabs = (CLUSTER[this.cluster]?.kab || []).map(k => k.toUpperCase());
      const m = {};
      this.tlRows.forEach(r => { if (kabs.includes(r.kab)) m[r.username_sf] = r.nama_sf || r.username_sf; });
      return Object.entries(m).map(([u, n]) => ({ u, n })).sort((a, b) => a.n.localeCompare(b.n));
    },

    // Satu entri per hari di bulan berjalan, berisi desa yang diplan
    get tlDays() {
      const p = this.tlPeriode, n = daysInPeriod(p), km = this.tlKatMap;
      const today = localYMD(new Date());
      const byTgl = {};
      this.tlClusterRows.forEach(r => {
        const k = tlKey(r.kec, r.desa);
        const day = byTgl[r.tgl] || (byTgl[r.tgl] = {});
        const d = day[k] || (day[k] = {
          key: k, desa: r.desa, kec: r.kec, kab: r.kab,
          kat: km[k]?.kat || 'lain', skor: km[k]?.skor ?? null, odps: [], sf: {},
        });
        d.odps.push({ odp: r.odp, sf: r.nama_sf || r.username_sf });
        d.sf[r.username_sf] = r.nama_sf || r.username_sf;
      });
      const out = [];
      for (let i = 1; i <= n; i++) {
        const tgl = p + '-' + String(i).padStart(2, '0');
        const dt = new Date(tgl + 'T00:00:00');
        let desas = Object.values(byTgl[tgl] || {})
          .map(d => ({ ...d, sfNames: Object.values(d.sf) }))
          .sort((a, b) => TL_KAT[a.kat].order - TL_KAT[b.kat].order || (b.skor ?? -1) - (a.skor ?? -1));
        const odpCount = desas.reduce((s, d) => s + d.odps.length, 0);
        if (this.tlKat) desas = desas.filter(d => d.kat === this.tlKat);
        out.push({
          tgl, day: i, hari: TL_HARI[dt.getDay()], minggu: dt.getDay() === 0,
          isToday: tgl === today, isFuture: tgl > today,
          desas, odpCount,
          sfCount: new Set(desas.flatMap(d => Object.keys(d.sf))).size,
        });
      }
      return out;
    },

    get tlSelDay() {
      return this.tlDays.find(d => d.tgl === this.tlSel) || null;
    },

    // Ringkasan bulan: desa unik per kategori yang sudah diplan vs total desa kategori itu
    get tlSummary() {
      const km = this.tlKatMap;
      const planned = {};
      this.tlClusterRows.forEach(r => {
        const k = tlKey(r.kec, r.desa);
        (planned[k] = planned[k] || { n: 0 }).n++;
      });
      const total = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0 };
      Object.values(km).forEach(v => {
        if (this.tlKab && v.kab !== this.tlKab.toUpperCase()) return;
        total[v.kat]++;
      });
      const done = { prioritas: 0, strategis: 0, cadangan: 0, skip: 0, lain: 0 };
      Object.keys(planned).forEach(k => { done[km[k]?.kat || 'lain']++; });
      const desaTotal = Object.keys(planned).length;
      return ['prioritas', 'strategis', 'cadangan', 'skip'].map(id => ({
        id, ...TL_KAT[id], done: done[id], total: total[id],
        pct: total[id] ? Math.round(done[id] / total[id] * 100) : 0,
        share: desaTotal ? Math.round(done[id] / desaTotal * 100) : 0,
      })).concat(done.lain ? [{ id: 'lain', ...TL_KAT.lain, done: done.lain, total: null, pct: 0, share: desaTotal ? Math.round(done.lain / desaTotal * 100) : 0 }] : []);
    },

    get tlTotals() {
      const rows = this.tlClusterRows;
      return {
        desa: new Set(rows.map(r => tlKey(r.kec, r.desa))).size,
        odp: rows.length,
        hari: new Set(rows.map(r => r.tgl)).size,
        sf: new Set(rows.map(r => r.username_sf)).size,
      };
    },

    // Desa Prioritas yang belum diplan sama sekali bulan ini (urut skor)
    get tlBelumDisentuh() {
      if (this.tlSf) return [];
      const planned = new Set(this.tlClusterRows.map(r => tlKey(r.kec, r.desa)));
      return Object.entries(this.tlKatMap)
        .filter(([k, v]) => v.kat === 'prioritas' && !planned.has(k) &&
                            (!this.tlKab || v.kab === this.tlKab.toUpperCase()))
        .map(([k, v]) => ({ key: k, ...v }))
        .sort((a, b) => b.skor - a.skor);
    },

    tlFmtTgl(tgl) {
      if (!tgl) return '';
      const dt = new Date(tgl + 'T00:00:00');
      return dt.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    },

    tlStep(dir) {
      const days = this.tlDays;
      const i = days.findIndex(d => d.tgl === this.tlSel);
      const j = Math.min(Math.max(i + dir, 0), days.length - 1);
      this.tlSel = days[j].tgl;
      this.$nextTick(() => {
        const el = document.querySelector('.tl-day.is-sel');
        if (el) el.scrollIntoView({ behavior: 'smooth', inline: 'center', block: 'nearest' });
      });
    },
  };
}
