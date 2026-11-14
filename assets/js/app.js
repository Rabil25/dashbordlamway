// Komponen utama dashboard (Alpine app()) — dipisah dari index.html (Nov 2026).
// Script klasik (bukan module): konstanta & fungsi di sini global dan dipakai file berikutnya.
// Urutan muat: config.js → app.js → djp.js → Alpine (defer).

// ============================================================
//  ALPINE APP
// ============================================================
// Satu pemuatan data Prioritas yang sedang berjalan dipakai bersama
// (prefetch latar belakang, menu Prioritas, Timeline, Plan DJP).
let _prioInflight = null;

function app() {
  return {
    // ── AUTH STATE ──────────────────────────────────────────
    currentUser:       JSON.parse(localStorage.getItem('wok_user') || 'null'),
    sessionToken:      localStorage.getItem('wok_token') || '',
    authView:          'login',   // 'login' | 'register'
    authLoading:       false,
    authError:         '',
    authSuccess:       '',
    authForm:          { username:'', password:'', password2:'', nama:'' },

    // Admin panel
    showAdminPanel:    false,
    adminUsers:        [],
    adminLoading:      false,
    adminFilter:       'pending',

    // Change password
    showChangePassword: false,
    chpwdForm:         { old:'', new1:'', new2:'' },
    chpwdError:        '',
    chpwdSuccess:      '',
    chpwdLoading:      false,

    // State
    page: 'dashboard',
    cluster: 'Lampung Barat',
    loading: false,
    filterLoading: false,
    showConfig: false,
    themeMode: localStorage.getItem('dashTheme') || 'light',
    allRows: [],      // baris untuk bulan aktif (GAS filter per bulan di server)
    _availBulanList: null,  // daftar bulan tersedia dari GAS list_bulan
    petaFilter: { io: true, re: true, ps: true },  // checkbox filter peta
    mapLayer: 'osm',  // 'osm' | 'satellite'
    petaFilterSTO: [],     // [] = semua STO, array untuk multi-select
    petaFilterProduk: '',  // '' = semua produk
    showODP: false,        // toggle ODP markers, default off
    _odpData: null,        // cache ODP JSON
    _odpLayer: null,       // Leaflet layer group ODP
    _odpCanvas: null,      // Canvas renderer (1 canvas untuk 4000+ titik, bukan per-DOM)
    lastUpdate: '',
    TIER_CONFIG,      // expose to template

    // Config
    cfg: {
      gasUrl:    localStorage.getItem('wok_gas_url')    || DEFAULT_GAS_URL,
      vercelUrl: localStorage.getItem('wok_vercel_url') || DEFAULT_VERCEL_URL,
      bulan: currentPeriod(),   // default: periode berjalan 'YYYY-MM'
      tglAwal: 1,                                  // default: tanggal 1
      tglAkhir: new Date().getDate(),              // default: hari ini (MTD)
      targets: {},  // diisi dari sheet Target saat loadData
    },

    // Filters
    sfSort: 'ps',
    sfSearch: '', sfKabFilter: '', sfProdukFilter: '',
    orderSearch: '', orderStatus: '', orderKab: '', orderProduk: '',
    orderPage: 0,
    orderSort: { key: 'tglIO', dir: -1 }, // default: tglIO terbaru dulu
    kendalaSearch: '', kendalaCat: '',

    // Prioritas Kunjungan state
    prioritasOdpRaw: [],
    prioritasLoading: false,
    prioritasError: '',
    prioritasKabFilter: '',
    prioritasKategoriFilter: '',
    prioritasSelectedDesa: null,
    prioritasMap: null,
    prioritasMarkers: [],
    prioritasMarkerMap: {},   // key: "kec||desa" → {marker, color, radius}
    prioritasSelectedMarker: null,

    // Penetrasi ODP state
    penetrasiData: [],
    penetrasiLoading: false,
    penetrasiError: '',
    penetrasiKabFilter: '',
    penetrasiKecFilter: '',
    penetrasiSort: { key: 'penetrasi', dir: 1 },

    // New page tabs
    ddTab: 'channel',
    olTab: 'ringkasan',
    sfpTab: 'scorecard',
    simDrr: 0,

    // Charts refs
    _charts: {},

    // UI state
    menuOpen: false,
    glossaryOpen: false,
    lastSync: localStorage.getItem('wok_last_sync') || '',

    // ── Init ──
    async init() {
      // Store bersama agar komponen DJP memakai URL Apps Script dari Pengaturan
      try { Alpine.store('cfg', this.cfg); } catch(e) {}
      // Expose deferred render to window for nested x-data components (date picker)
      window._deferredRender = () => this._deferredRender();
      window._deferredUpdateTgl = (a, b) => this._deferredUpdateTgl(a, b);

      // Apply saved theme
      document.documentElement.setAttribute('data-theme', this.themeMode);

      // Verifikasi session token jika ada — PARALEL dengan muat data.
      // Sebelumnya data baru diminta setelah auth_verify selesai (~10 detik ke Apps Script).
      if (this.sessionToken && this.currentUser) {
        this.loading = true;
        if (this.cfg.gasUrl) this.loadData();
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        fetch(gasBase + '?type=auth_verify&session_token=' + encodeURIComponent(this.sessionToken))
          .then(r => r.json())
          .then(json => {
            if (!json.ok) { this.doLogout(); this.toast('Sesi habis, silakan login ulang.', 'err'); return; }
            this.currentUser = json.user;
            localStorage.setItem('wok_user', JSON.stringify(json.user));
          })
          .catch(() => { /* jaringan error: tetap pakai sesi lokal */ });
      } else {
        this.currentUser = null;
      }

      // Resize hanya grafik yang sedang terlihat (halaman lain display:none → resize error),
      // di-debounce agar tidak berulang saat address bar HP muncul/hilang.
      let _rzT;
      window.addEventListener('resize', () => {
        clearTimeout(_rzT);
        _rzT = setTimeout(() => {
          Object.values(this._charts).forEach(c => {
            try {
              if (!c || !c.resize || (c.isDisposed && c.isDisposed())) return;
              const el = c.getDom && c.getDom();
              if (el && el.offsetParent === null) return;
              c.resize();
            } catch (e) { /* abaikan grafik yang belum siap */ }
          });
        }, 150);
      });
    },

    toggleTheme() {
      this.themeMode = this.themeMode === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', this.themeMode);
      localStorage.setItem('dashTheme', this.themeMode);
      // Dispose all charts so they reinit with new theme colors on next render
      setTimeout(() => {
        this.disposeCharts();
        this.renderCharts();
      }, 80);
    },

    // ════════════════════════════════════════════════════════
    //  AUTH FUNCTIONS
    // ════════════════════════════════════════════════════════

    // SHA-256 → hex (Web Crypto API)
    async sha256(str) {
      const buf  = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
      return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2,'0')).join('');
    },

    async doLogin() {
      this.authError = '';
      const { username, password } = this.authForm;
      if (!username || !password) { this.authError = 'Username dan password wajib diisi.'; return; }
      this.authLoading = true;
      try {
        const hash = await this.sha256(password);
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        if (!gasBase) { this.authError = 'GAS URL belum dikonfigurasi. Hubungi admin.'; return; }
        const url = gasBase + '?type=auth_login&username=' + encodeURIComponent(username) + '&hash=' + hash;
        const res  = await fetch(url);
        const json = await res.json();
        if (!json.ok) { this.authError = json.error; return; }
        // Simpan session
        this.currentUser  = json.user;
        this.sessionToken = json.token;
        localStorage.setItem('wok_user',  JSON.stringify(json.user));
        localStorage.setItem('wok_token', json.token);
        this.authForm = { username:'', password:'', password2:'', nama:'' };
        // Muat data dashboard
        if (this.cfg.gasUrl) this.loadData();
      } catch(e) {
        this.authError = 'Gagal terhubung ke server. Periksa koneksi internet.';
      } finally {
        this.authLoading = false;
      }
    },

    async doRegister() {
      this.authError = ''; this.authSuccess = '';
      const { username, password, password2, nama } = this.authForm;
      if (!nama) { this.authError = 'Nama lengkap wajib diisi.'; return; }
      if (!username) { this.authError = 'Username wajib diisi.'; return; }
      if (!/^[a-z0-9_]{3,20}$/.test(username)) { this.authError = 'Username: 3-20 karakter, huruf kecil/angka/underscore.'; return; }
      if (!password || password.length < 6) { this.authError = 'Password minimal 6 karakter.'; return; }
      if (password !== password2) { this.authError = 'Konfirmasi password tidak cocok.'; return; }
      this.authLoading = true;
      try {
        const hash = await this.sha256(password);
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        if (!gasBase) { this.authError = 'GAS URL belum dikonfigurasi. Hubungi admin.'; return; }
        const url = gasBase + '?type=auth_register&username=' + encodeURIComponent(username)
                  + '&hash=' + hash + '&nama=' + encodeURIComponent(nama);
        const res  = await fetch(url);
        const json = await res.json();
        if (!json.ok) { this.authError = json.error; return; }
        this.authSuccess = json.message || 'Pendaftaran berhasil!';
        this.authForm.password = ''; this.authForm.password2 = '';
      } catch(e) {
        this.authError = 'Gagal terhubung ke server.';
      } finally {
        this.authLoading = false;
      }
    },

    async doLogout() {
      try {
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        if (gasBase && this.sessionToken) {
          fetch(gasBase + '?type=auth_logout&session_token=' + encodeURIComponent(this.sessionToken));
        }
      } catch(e) {}
      this.currentUser  = null;
      this.sessionToken = '';
      localStorage.removeItem('wok_user');
      localStorage.removeItem('wok_token');
      // Reset data
      this.allRows = [];
      this.authView = 'login';
      this.authError = '';
      this.authSuccess = '';
    },

    async loadPendingUsers() {
      if (!this.sessionToken) return;
      this.adminLoading = true;
      this.adminUsers = [];
      try {
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        const url = gasBase + '?type=auth_pending_users&session_token=' + encodeURIComponent(this.sessionToken)
                  + '&filter=' + this.adminFilter;
        const res  = await fetch(url);
        const json = await res.json();
        if (json.ok) this.adminUsers = json.users || [];
        else this.toast(json.error || 'Gagal memuat data user.', 'err');
      } catch(e) {
        this.toast('Gagal terhubung ke server.', 'err');
      } finally {
        this.adminLoading = false;
      }
    },

    async approveUser(username, action) {
      try {
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        const url = gasBase + '?type=auth_approve&session_token=' + encodeURIComponent(this.sessionToken)
                  + '&target_username=' + encodeURIComponent(username) + '&action=' + action;
        const res  = await fetch(url);
        const json = await res.json();
        if (json.ok) {
          this.toast(json.message, 'ok');
          this.loadPendingUsers();
        } else {
          this.toast(json.error || 'Gagal.', 'err');
        }
      } catch(e) {
        this.toast('Gagal terhubung ke server.', 'err');
      }
    },

    async doChangePassword() {
      this.chpwdError = ''; this.chpwdSuccess = '';
      const { old: oldPwd, new1, new2 } = this.chpwdForm;
      if (!oldPwd) { this.chpwdError = 'Password lama wajib diisi.'; return; }
      if (!new1 || new1.length < 6) { this.chpwdError = 'Password baru minimal 6 karakter.'; return; }
      if (new1 !== new2) { this.chpwdError = 'Konfirmasi password baru tidak cocok.'; return; }
      this.chpwdLoading = true;
      try {
        const oldHash = await this.sha256(oldPwd);
        const newHash = await this.sha256(new1);
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        const url = gasBase + '?type=auth_change_password&session_token=' + encodeURIComponent(this.sessionToken)
                  + '&old_hash=' + oldHash + '&new_hash=' + newHash;
        const res  = await fetch(url);
        const json = await res.json();
        if (json.ok) {
          this.chpwdSuccess = json.message || 'Password berhasil diubah!';
          this.chpwdForm = { old:'', new1:'', new2:'' };
          setTimeout(() => {
            this.showChangePassword = false; this.chpwdSuccess = '';
            if (json.relogin) this.doLogout();   // sesi lama sudah dicabut server
          }, 2000);
        } else {
          this.chpwdError = json.error;
        }
      } catch(e) {
        this.chpwdError = 'Gagal terhubung ke server.';
      } finally {
        this.chpwdLoading = false;
      }
    },

    // ── Nav ──
    nav(p) {
      this.page = p;
      this.orderPage = 0;
      this.$nextTick(() => {
        this.renderCharts();
        // dispatch event untuk DJP pages supaya auto-load
        if (['djp_plan','djp_real','djp_daily','djp_odp_list','timeline'].includes(p)) {
          window.dispatchEvent(new CustomEvent('page-' + p));
        }
        // Prioritas memuat sendiri saat dibuka (tanpa klik "Muat Data")
        if (p === 'prioritas') {
          if (!this.prioritasOdpRaw.length) this.loadPrioritas().catch(() => {});
          else this.initPrioritasMap();
        }
      });
    },

    get pageTitle() {
      const map = {
        dashboard:'Sales Dashboard', d2d:'D2D Sales', deepdive:'Deepdive Sales',
        outlook:'Outlook Sales', sfperf:'SF Performance',
        sf:'Performa SF', order:'Detail Order', kendala:'Kendala', peta:'Peta', prioritas:'Prioritas Kunjungan',
        djp_plan:'Plan DJP', timeline:'Timeline Desa', djp_real:'Realisasi DJP', djp_daily:'DJP Daily', djp_odp_list:'List ODP',
      };
      return map[this.page] || '';
    },

    // ── Cluster ──
    setCluster(c) {
      // Show loading bar FIRST, then change data in next frame so browser can repaint
      this.filterLoading = true;
      clearTimeout(this._renderTimer);
      this._renderTimer = setTimeout(() => {
        this.cluster = c;
        this.sfKabFilter = '';
        this.orderKab = '';
        this.orderPage = 0;
        this.penetrasiKabFilter = '';
        this.penetrasiKecFilter = '';
        this.penetrasiData = [];
        this.penetrasiSort = { key: 'penetrasi', dir: 1 };
        this.filterLoading = false;
        this.$nextTick(() => {
          this.renderCharts();
          // Trigger re-init komponen DJP Plan ketika cluster berubah
          window.dispatchEvent(new CustomEvent('page-djp_plan'));
        });
      }, 30);
    },

    _deferredRender() {
      this.filterLoading = true;
      clearTimeout(this._renderTimer);
      this._renderTimer = setTimeout(() => {
        this.filterLoading = false;
        this.$nextTick(() => this.renderCharts());
      }, 30);
    },

    _deferredUpdateTgl(a, b) {
      // Show loading bar FIRST, then update date range in next frame so browser can repaint
      this.filterLoading = true;
      clearTimeout(this._renderTimer);
      this._renderTimer = setTimeout(() => {
        this.cfg.tglAwal = a;
        this.cfg.tglAkhir = b;
        this.filterLoading = false;
        this.$nextTick(() => this.renderCharts());
      }, 30);
    },

    get clusterKab() {
      return CLUSTER[this.cluster]?.kab || [];
    },

    // ── Config ──
    saveConfig() {
      localStorage.setItem('wok_gas_url',    this.cfg.gasUrl);
      localStorage.setItem('wok_vercel_url', this.cfg.vercelUrl || '');
      this.showConfig = false;
      this.loadData();
    },

    get dpCells() {
      const yr = periodYear(this.cfg.bulan);
      const mo = periodMonth(this.cfg.bulan) - 1;
      const first = new Date(yr, mo, 1).getDay();
      const days  = new Date(yr, mo + 1, 0).getDate();
      const cells = [];
      for (let i = 0; i < first; i++) cells.push({ d: 0, other: true });
      for (let d = 1; d <= days; d++) cells.push({ d, other: false });
      while (cells.length % 7 !== 0) cells.push({ d: 0, other: true });
      return cells;
    },

    // Periode sebelumnya ('YYYY-MM'); Jan 2027 → Des 2026
    get prevBulan() { return prevPeriod(this.cfg.bulan); },

    // Daftar periode: dari list_bulan server, fallback dari allRows, fallback periode berjalan
    get availableMonths() {
      let list = (this._availBulanList || []).map(String).filter(v => /^\d{4}-\d{2}$/.test(v));
      if (!list.length) {
        const set = new Set();
        for (const r of this.allRows) [C.BLN_IO, C.BLN_RE, C.BLN_PS].forEach(c => { const v = r[c]; if (v && /^\d{4}-\d{2}$/.test(v)) set.add(v); });
        list = [...set];
      }
      if (!list.length) list = [currentPeriod()];
      return list.sort().map(v => ({ val: v, label: periodLabel(v) }));
    },

    get prevBulanLabel() { return periodLabel(this.prevBulan); },

    // ── Load Data via GAS (filter per bulan di server) ──
    async loadData(forceRefresh = false) {
      if (!this.currentUser) return;  // belum login
      if (!this.cfg.gasUrl) {
        if (this.currentUser.role === 'admin') this.showConfig = true;
        return;
      }
      this.loading = true;
      const CACHE_TTL_MS    = 15 * 60 * 1000; // 15 menit (in-memory)
      const CACHE_BULAN_MS  = 60 * 60 * 1000; // list_bulan cache 1 jam

      try {
        // Cek apakah semua data sudah ada di in-memory cache
        if (!this._memCache) this._memCache = new Map();
        const bulanStr     = this.cfg.bulan;
        const prevBulanStr = String(this.prevBulan);
        const cacheKeyAll  = `all_b${bulanStr}`;
        const cacheKeyBulan = 'list_bulan';

        const cachedAll   = !forceRefresh && this._memCache.get(cacheKeyAll);
        const cachedBulan = !forceRefresh && this._memCache.get(cacheKeyBulan);
        const nowMs = Date.now();

        if (cachedAll   && (nowMs - cachedAll.ts)   < CACHE_TTL_MS &&
            cachedBulan && (nowMs - cachedBulan.ts) < CACHE_BULAN_MS) {
          // Semua data dari cache – tidak perlu request GAS sama sekali
          const d = cachedAll.data;
          { const seen=new Set(); _loadSeq++; this.allRows=normalizePeriods([...(d.bulanIni||[]),...(d.bulanPrev||[])]).filter(r=>{ const id=(r[C.ID_ORDER]||'').toString().trim()||JSON.stringify(r); if(seen.has(id))return false; seen.add(id);return true; }); }
          this.sfMaster = d.updateSf || [];
          const nowStr  = new Date().toLocaleTimeString('id-ID', { hour:'2-digit', minute:'2-digit' });
          this.lastUpdate = nowStr;
          this.lastSync   = nowStr;
          this.$nextTick(() => this.renderCharts());
          this.loading = false;
          return;
        }

        // Belum di cache atau forceRefresh → 1 request ke endpoint type=all
        const base = getDataUrl(this.cfg);
        // SEC-04: API Vercel hanya melayani user login → kirim session token di header
        // (bukan di URL). Fallback ke Apps Script tetap memakai token lama.
        const viaGas = base.includes('script.google.com');
        const url  = viaGas ? `${base}?session_token=${encodeURIComponent(this.sessionToken)}&type=all&bulan=${bulanStr}`
                            : `${base}?type=all&bulan=${bulanStr}` + (forceRefresh ? `&_=${Date.now()}` : '');
        const res  = await fetch(url, viaGas ? {} : { headers: { Authorization: 'Bearer ' + this.sessionToken } });
        if (res.status === 401) {
          this.loading = false;
          this.doLogout();
          this.toast('Sesi berakhir. Silakan login ulang.', 'err');
          return;
        }
        if (!res.ok) throw new Error('HTTP ' + res.status);
        const json = await res.json();
        if (!json.ok) throw new Error(json.error || 'GAS error (all)');

        // Simpan ke in-memory cache
        const allPayload = {
          bulanIni:  json.bulanIni?.data  || [],
          bulanPrev: json.bulanPrev?.data || [],
          updateSf:  json.updateSf?.data  || [],
        };
        this._memCache.set(cacheKeyAll, { data: allPayload, ts: nowMs });

        // Simpan list_bulan juga dari respons yang sama
        if (json.listBulan?.ok && json.listBulan.bulan) {
          this._memCache.set(cacheKeyBulan, { data: json.listBulan.bulan, ts: nowMs });
          const listBulanRaw = json.listBulan.bulan;
          if (listBulanRaw.length > 0) {
            const bulanAvail = listBulanRaw.map(String);
            this._availBulanList = listBulanRaw;
            if (!bulanAvail.includes(bulanStr)) {
              // Bulan terpilih belum punya data (mis. tanggal 1) → pindah ke periode terakhir yang ada, muat ulang
              this.cfg.bulan = [...bulanAvail].sort().pop();
              this.loading = false;
              return this.loadData(forceRefresh);
            }
          }
        }

        // Terapkan ke state (dedup by ID_ORDER agar baris lintas-bulan tidak double-count)
        { const seen=new Set(); _loadSeq++; this.allRows=normalizePeriods([...allPayload.bulanIni,...allPayload.bulanPrev]).filter(r=>{ const id=(r[C.ID_ORDER]||'').toString().trim()||JSON.stringify(r); if(seen.has(id))return false; seen.add(id);return true; }); }
        this.sfMaster = allPayload.updateSf;

        // Ambil target dari sheet (override localStorage)
        if (json.target?.ok && json.target.targets) {
          this.cfg.targets = json.target.targets;
        }

        const nowStr = new Date().toLocaleTimeString('id-ID', { hour:'2-digit', minute:'2-digit' });
        this.lastUpdate = nowStr;
        this.lastSync   = nowStr;

        this.$nextTick(() => this.renderCharts());
        this.prefetchPrioritas();

        this.toast(
          allPayload.bulanIni.length + ' baris bln ' + bulanStr +
          ' + ' + allPayload.bulanPrev.length + ' baris bln ' + prevBulanStr + ' dimuat',
          'ok'
        );

      } catch(e) {
        this.toast('Gagal memuat: ' + e.message, 'err');
      }
      this.loading = false;
    },

    // ── Fetch daftar bulan dari GAS (ringan) ──
    // Hari ke-berapa dalam bulan terpilih: bulan berjalan → hari ini; bulan lampau → hari terakhir bulan itu
    get hariBerjalan() {
      if (this.cfg.bulan === currentPeriod()) return new Date().getDate();
      return daysInPeriod(this.cfg.bulan);
    },

    inRange(val) {
      if (val === null || val === undefined || val === '') return false;
      const a = parseInt(this.cfg.tglAwal)  || 1;
      const b = parseInt(this.cfg.tglAkhir) || new Date().getDate();
      // Angka integer langsung = hari
      if (typeof val === 'number') return val >= a && val <= b;
      // String angka (e.g. "5", "21")
      const n = parseInt(val);
      if (!isNaN(n) && String(n) === String(val).trim()) return n >= a && n <= b;
      // String tanggal "yyyy-MM-dd" atau "yyyy-MM-dd HH:mm:ss"
      const d = new Date(val);
      if (isNaN(d.getTime())) return false;
      return d.getDate() >= a && d.getDate() <= b;
    },

    // ── Filtered rows (cluster + dedup) — dari allRows, filtering bulan dilakukan di getter KPI ──
    // Cached via closure variable _frCache: filter 885 baris hanya 1x per perubahan cluster/data
    get filteredRows() {
      const key = this.cluster + '|' + _loadSeq + '|' + this.allRows.length;
      if (_frKey === key && _frCache) return _frCache;
      const kab = CLUSTER[this.cluster]?.kab || [];
      _frCache = this.allRows.filter(r => {
        if ((r[C.DEDUP]||'').trim() !== 'Unik') return false;
        if (!kab.includes((r[C.VAL_KAB]||'').trim().toUpperCase())) return false;
        return true;
      });
      _frKey = key;
      return _frCache;
    },

    // ── Filtered rows bulan sebelumnya (untuk MoM) — sama source, filter bulan di getter ──
    get filteredPrevRows() {
      const key = this.cluster + '|' + _loadSeq + '|' + this.allRows.length;
      if (_fprKey === key && _fprCache) return _fprCache;
      const kab = CLUSTER[this.cluster]?.kab || [];
      _fprCache = this.allRows.filter(r => {
        if ((r[C.DEDUP]||'').trim() !== 'Unik') return false;
        if (!kab.includes((r[C.VAL_KAB]||'').trim().toUpperCase())) return false;
        return true;
      });
      _fprKey = key;
      return _fprCache;
    },

    // ── KPI ──
    get kpi() {
      const rows = this.filteredRows;
      let io=0, re=0, ps=0, cancel=0;
      const bulan = this.cfg.bulan;
      const today = new Date();
      const hariPS = parseInt(this.cfg.tglAkhir) || today.getDate();

      // Weekly buckets for WoW (minggu ini = last 7 days, minggu lalu = 8-14 days ago)
      const now = today.getTime();
      const d7  = now - 7*86400000;
      const d14 = now - 14*86400000;
      let psWk0=0, psWk1=0;

      rows.forEach(r => {
        const blnIO = (r[C.BLN_IO]||'').toString();
        const blnRE = (r[C.BLN_RE]||'').toString();
        const blnPS = (r[C.BLN_PS]||'').toString();
        const valIO = (r[C.VAL_IO]||'').toString().trim();
        const valRE = (r[C.VAL_RE]||'').toString().trim();
        const valPS = (r[C.VAL_PS]||'').toString().trim();
        // IO: masuk rentang tanggal IO + status IO tidak kosong
        if (blnIO === bulan && this.inRange(r[C.TGL_IO]) && valIO !== '') io++;
        // RE: masuk rentang tanggal RE + status RE tidak kosong
        if (blnRE === bulan && this.inRange(r[C.TGL_RE]) && valRE !== '') re++;
        // PS: masuk rentang tanggal PS + status PS tidak kosong
        if (blnPS === bulan && this.inRange(r[C.TGL_PS]) && valPS !== '') {
          ps++;
          // weekly PS
          const tgl = r[C.DATE_PS] || r[C.DATE_IO] || '';
          const tglFmt = (tgl && typeof tgl === 'string' && tgl.length === 8)
            ? tgl.slice(0,4)+'-'+tgl.slice(4,6)+'-'+tgl.slice(6) : tgl;
          const ts  = tglFmt ? new Date(tglFmt).getTime() : 0;
          if (ts > d7)       psWk0++;
          else if (ts > d14) psWk1++;
        }
        const fx = (r[C.STATUS_FNL]||'').toUpperCase();
        if (blnIO === bulan && this.inRange(r[C.TGL_IO]) && (fx.includes('CANCEL') || fx.includes('BATAL'))) cancel++;
      });

      // MoM: hitung dari prevRows dengan rentang tanggal yang SAMA
      const prevBulan = this.prevBulan.toString();
      let ioPrevTotal=0, rePrevTotal=0, psPrevTotal=0;
      this.filteredPrevRows.forEach(r => {
        const blnIO = (r[C.BLN_IO]||'').toString();
        const blnRE = (r[C.BLN_RE]||'').toString();
        const blnPS = (r[C.BLN_PS]||'').toString();
        const valIO = (r[C.VAL_IO]||'').toString().trim();
        const valRE = (r[C.VAL_RE]||'').toString().trim();
        const valPS = (r[C.VAL_PS]||'').toString().trim();
        if (blnPS === prevBulan && this.inRange(r[C.TGL_PS]) && valPS !== '') psPrevTotal++;
        if (blnIO === prevBulan && this.inRange(r[C.TGL_IO]) && valIO !== '') ioPrevTotal++;
        if (blnRE === prevBulan && this.inRange(r[C.TGL_RE]) && valRE !== '') rePrevTotal++;
      });

      // Total targets for cluster
      const kabs = CLUSTER[this.cluster]?.kab || [];
      let tgtIO=0, tgtRE=0, tgtPS=0;
      kabs.forEach(k => {
        const t = this.cfg.targets[k] || {};
        tgtPS += t.ps||0; tgtRE += t.re||0; tgtIO += t.io||0;
      });

      const momPS = tgtPS > 0 ? Math.round(ps/tgtPS*100) : null;
      const momRE = tgtRE > 0 ? Math.round(re/tgtRE*100) : null;
      const momIO = tgtIO > 0 ? Math.round(io/tgtIO*100) : null;
      const wowPS = psWk1 > 0 ? Math.round((psWk0-psWk1)/psWk1*100) : null;
      const momPSvsLast = psPrevTotal > 0 ? Math.round((ps-psPrevTotal)/psPrevTotal*100) : null;

      return {
        io, re, ps, cancel,
        hariPS,
        drr: hariPS > 0 ? (ps / hariPS).toFixed(2) : '0',
        psRE: re > 0 ? (ps/re*100).toFixed(1) : '0',
        reIO: io > 0 ? (re/io*100).toFixed(1) : '0',
        tgtPS, tgtRE, tgtIO,
        momPS, momRE, momIO,
        wowPS, momPSvsLast,
        psPrev: psPrevTotal,
        ioPrev: ioPrevTotal,
        rePrev: rePrevTotal,
        momIOvsLast: ioPrevTotal > 0 ? Math.round((io - ioPrevTotal) / ioPrevTotal * 100) : null,
        momREvsLast: rePrevTotal > 0 ? Math.round((re - rePrevTotal) / rePrevTotal * 100) : null,
      };
    },

    // ── Breakdown per Kabupaten ──
    get kabBreakdown() {
      const rows = this.filteredRows;
      const bulan = this.cfg.bulan;
      const prevBulan = this.prevBulan.toString();
      const kabs = CLUSTER[this.cluster]?.kab || [];
      const map = {};
      kabs.forEach(k => { map[k] = { kab:k, io:0, re:0, ps:0, prevPS:0 }; });

      rows.forEach(r => {
        const kab = (r[C.VAL_KAB]||'').trim().toUpperCase();
        if (!map[kab]) return;
        const blnIO = (r[C.BLN_IO]||'').toString();
        const blnRE = (r[C.BLN_RE]||'').toString();
        const blnPS = (r[C.BLN_PS]||'').toString();
        if (blnIO === bulan) map[kab].io++;
        if (blnRE === bulan) map[kab].re++;
        if (blnPS === bulan) map[kab].ps++;
        if (blnPS === prevBulan && this.inRange(r[C.TGL_PS])) map[kab].prevPS++;
      });

      return kabs.map(k => {
        const d = map[k];
        const t = this.cfg.targets[k] || {};
        const tPS = t.ps||0, tRE = t.re||0, tIO = t.io||0;
        const momPS = tPS > 0 ? Math.round(d.ps/tPS*100) : null;
        const momRE = tRE > 0 ? Math.round(d.re/tRE*100) : null;
        const psRE  = d.re > 0 ? Math.round(d.ps/d.re*100) : 0;
        const momDelta = d.prevPS > 0 ? Math.round((d.ps - d.prevPS)/d.prevPS*100) : null;
        return { ...d, tPS, tRE, tIO, momPS, momRE, psRE, momDelta };
      });
    },

    // ── Top SF ──
    get topSF() {
      return this.buildSFTable().slice(0, 10);
    },

    buildSFTable() {
      const map = {};
      const bulan = this.cfg.bulan;
      const prevBulan = this.prevBulan.toString();
      const produkFilter = this.sfProdukFilter || '';
      const currentCluster = this.cluster;

      // ── Seed map dari Update SF: semua SF ACTIVE di cluster ini ──
      // Col 0-based: 1=Nama, 2=Kode, 16=Status, 20=Kab
      // Filter cluster berdasarkan kab (kolom AI di sheet kosong)
      const USF_NAMA = 1, USF_KODE = 2, USF_STATUS = 16, USF_KAB = 20;
      const clusterKabs = (CLUSTER[currentCluster]?.kab || []).map(k => k.toUpperCase());
      const validCodes = new Set(); // SF ACTIVE yang kab-nya masuk cluster ini
      (this.sfMaster || []).forEach(row => {
        const status = (row[USF_STATUS] || '').toString().trim().toUpperCase();
        if (status !== 'ACTIVE' && status !== 'AKTIF') return;
        const kab = (row[USF_KAB] || '').toString().trim().toUpperCase();
        if (clusterKabs.length > 0 && !clusterKabs.includes(kab)) return;
        const code = (row[USF_KODE] || '').toString().trim();
        if (!code) return;
        validCodes.add(code);
        if (!map[code]) map[code] = {
          code, nama: (row[USF_NAMA]||'').toString().trim(), channel: '',
          kab,
          io:0, re:0, ps:0, cancel:0, prevPS:0, prevPSFull:0, produk: '',
        };
      });

      // Hitung PS bulan ini dari filteredRows — filter produk masuk ke sini
      this.filteredRows.forEach(r => {
        const code = (r[C.CODE_SF]||'').trim();
        if (!code) return;
        // Jika sfMaster sudah ter-load dengan cluster data, hanya proses SF yang
        // terdaftar di cluster ini. Ini mencegah SF lintas-cluster masuk tabel.
        // Guard validCodes.size > 0 → fallback ke perilaku lama jika sfMaster belum ada cluster.
        if (validCodes.size > 0 && !validCodes.has(code)) return;
        const produk = (r[C.PRODUK_KAT]||'');
        // Inisialisasi entry SF (pakai semua row untuk nama/kab/channel)
        if (!map[code]) map[code] = {
          code, nama: r[C.NAMA_SF]||'', channel: r[C.CHANNEL2]||'',
          kab: r[C.VAL_KAB]||'', io:0, re:0, ps:0, cancel:0, prevPS:0, prevPSFull:0,
          produk,
        };
        // Update channel & produk dari data order (lebih detail dari Update SF)
        const m = map[code];
        if (!m.channel && r[C.CHANNEL2]) m.channel = r[C.CHANNEL2];
        if (!m.produk  && produk)        m.produk  = produk;
        // Hitung angka HANYA untuk produk yang dipilih (atau semua jika tidak ada filter)
        if (produkFilter && produk !== produkFilter) return;
        const blnIOsf = (r[C.BLN_IO]||'').toString();
        const blnPS   = (r[C.BLN_PS]||'').toString();
        const blnRE   = (r[C.BLN_RE]||'').toString();
        if (blnIOsf===bulan) m.io++;
        if (blnRE===bulan) m.re++;
        if (blnPS===bulan) m.ps++;
        const fx = (r[C.STATUS_FNL]||'').toUpperCase();
        if (blnIOsf===bulan && (fx.includes('CANCEL') || fx.includes('BATAL'))) m.cancel++;
      });

      // Hitung prevPS (ikut rentang tanggal — untuk perbandingan progress)
      // dan prevPSFull (full bulan lalu — untuk Tier LM & Target, tidak terpengaruh filter)
      this.filteredPrevRows.forEach(r => {
        const code = (r[C.CODE_SF]||'').trim();
        if (!code) return;
        if (validCodes.size > 0 && !validCodes.has(code)) return;
        if (!map[code]) map[code] = {
          code, nama: r[C.NAMA_SF]||'', channel: r[C.CHANNEL2]||'',
          kab: r[C.VAL_KAB]||'', io:0, re:0, ps:0, cancel:0, prevPS:0, prevPSFull:0,
          produk: r[C.PRODUK_KAT]||'',
        };
        if ((r[C.BLN_PS]||'').toString()===prevBulan) {
          if (this.inRange(r[C.TGL_PS])) map[code].prevPS++;   // ikut filter rentang
          map[code].prevPSFull++;                               // selalu full bulan
        }
      });

      const arr = Object.values(map);
      const today = new Date();
      const hariIni = today.getDate();
      const isCurrentMonth = bulan === currentPeriod();

      arr.forEach(sf => {
        // PS/RE conversion (menggantikan LOS yang salah)
        sf.psRE = sf.re > 0 ? Math.round((sf.ps / sf.re) * 100) : 0;
        // Tier LM & Target dari prevPSFull (full bulan lalu, tidak ikut filter rentang)
        // agar plan operasional tidak kacau saat filter diganti
        const tierPrev = getSFTier(sf.prevPSFull);
        sf.tier     = tierPrev.name;
        sf.tierIcon = tierPrev.icon;
        sf.target   = tierPrev.target;
        sf.action   = getSFAction(sf.ps, sf.target, hariIni);
        sf.ach      = sf.target > 0 ? Math.round((sf.ps / sf.target) * 100) : 0;
        // Tier bulan ini (berdasarkan PS bulan ini) — untuk tabel gabungan
        const tierNow = getSFTier(sf.ps);
        sf.tierNow     = tierNow.name;
        sf.tierNowIcon = tierNow.icon;
        // Flag SF ACTIVE dengan PS=0 setelah tanggal 15 bulan berjalan
        // (io > 0 tidak disyaratkan lagi karena semua SF ACTIVE sudah masuk dari Update SF)
        sf.alertPS0 = isCurrentMonth && hariIni > 15 && sf.ps === 0;
      });

      return arr.sort((a,b) => b.ps - a.ps);
    },

    get sfTable() { return this.buildSFTable(); },

    get filteredSF() {
      let arr = this.sfTable;
      const q = this.sfSearch.toLowerCase();
      if (q) arr = arr.filter(sf => sf.code.toLowerCase().includes(q) || sf.nama.toLowerCase().includes(q));
      if (this.sfKabFilter) arr = arr.filter(sf => sf.kab.toUpperCase() === this.sfKabFilter.toUpperCase());
      // Filter produk: buildSFTable sudah recount angka per produk,
      // tapi kita tetap sembunyikan SF yang sama sekali tidak punya transaksi produk tersebut
      if (this.sfProdukFilter) arr = arr.filter(sf => sf.io > 0 || sf.re > 0 || sf.ps > 0);
      // Sort
      const tierOrder = ['Diamond','Platinum','Gold','Silver','Bronze','Black'];
      if (this.sfSort === 'ach')  arr = [...arr].sort((a,b) => b.ach - a.ach);
      else if (this.sfSort === 'tier') arr = [...arr].sort((a,b) => tierOrder.indexOf(a.tier) - tierOrder.indexOf(b.tier));
      else if (this.sfSort === 'prev') arr = [...arr].sort((a,b) => b.prevPS - a.prevPS);
      // default 'ps' is already sorted in buildSFTable
      return arr;
    },

    get sfProdukList() {
      const s = new Set();
      this.sfTable.forEach(sf => { if (sf.produk) s.add(sf.produk); });
      return ['', ...Array.from(s).sort()];
    },

    get avgAch() {
      const arr = this.filteredSF;
      if (!arr.length) return 0;
      return Math.round(arr.reduce((s,sf) => s + sf.ach, 0) / arr.length);
    },

    // ── D2D Sales ──
    get d2dKpi() {
      const rows = this.filteredRows;
      const bulan = this.cfg.bulan;
      const today = new Date();
      const hariPS = this.hariBerjalan;
      let ps = 0;
      const dayMap = {}; // day → count
      rows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString() === bulan) {
          ps++;
          const d = parseInt(r[C.TGL_PS]||0);
          if (d >= 1) dayMap[d] = (dayMap[d]||0) + 1;
        }
      });
      const drr = hariPS > 0 ? ps / hariPS : 0;
      // Hari capai target: days where PS count >= daily target (DRR)
      // Use daily target = totalTgtPS / daysInMonth
      const kabs = CLUSTER[this.cluster]?.kab || [];
      let tgtPS = 0;
      kabs.forEach(k => { tgtPS += (this.cfg.targets[k]?.ps||0); });
      const daysInMonth = daysInPeriod(bulan);
      const dailyTgt = daysInMonth > 0 ? tgtPS / daysInMonth : 0;
      let hariCapai = 0;
      for (let d = 1; d <= hariPS; d++) {
        if ((dayMap[d]||0) >= dailyTgt && dailyTgt > 0) hariCapai++;
      }
      return { ps, drr: drr.toFixed(2), hariCapai, hariPS, tgtPS, daysInMonth };
    },

    get d2dHeatmap() {
      const rows = this.filteredRows;
      const bulan = this.cfg.bulan;
      const kabs = CLUSTER[this.cluster]?.kab || [];
      // dayMap[kab][day] = count
      const map = {};
      kabs.forEach(k => { map[k] = {}; });
      let maxVal = 0;
      rows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString() !== bulan) return;
        const kab = (r[C.VAL_KAB]||'').trim().toUpperCase();
        if (!map[kab]) return;
        const d = parseInt(r[C.TGL_PS]||0);
        if (d < 1 || d > 31) return;
        map[kab][d] = (map[kab][d]||0) + 1;
        if (map[kab][d] > maxVal) maxVal = map[kab][d];
      });
      const daysInMonth = daysInPeriod(bulan);
      return { map, kabs, daysInMonth, maxVal };
    },

    hmClass(val, max) {
      if (!val || val === 0) return 'hm-0';
      const pct = max > 0 ? val / max : 0;
      if (pct < 0.2) return 'hm-lo';
      if (pct < 0.45) return 'hm-md';
      if (pct < 0.75) return 'hm-hi';
      return 'hm-vhi';
    },

    // ── Deepdive ──
    get deepdiveChannel() {
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        const ch = r[C.CHANNEL2] || 'Lainnya';
        if (!map[ch]) map[ch] = {io:0,re:0,ps:0};
        if ((r[C.BLN_IO]||'').toString()===bulan) map[ch].io++;
        if ((r[C.BLN_RE]||'').toString()===bulan) map[ch].re++;
        if ((r[C.BLN_PS]||'').toString()===bulan) map[ch].ps++;
      });
      return Object.entries(map).map(([name,v])=>({name,...v})).sort((a,b)=>b.ps-a.ps);
    },

    get deepdiveProduk() {
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        const p = r[C.PRODUK_KAT] || 'Lainnya';
        if (!map[p]) map[p] = {io:0,re:0,ps:0};
        if ((r[C.BLN_IO]||'').toString()===bulan) map[p].io++;
        if ((r[C.BLN_RE]||'').toString()===bulan) map[p].re++;
        if ((r[C.BLN_PS]||'').toString()===bulan) map[p].ps++;
      });
      return Object.entries(map).map(([name,v])=>({name,...v})).sort((a,b)=>b.ps-a.ps);
    },

    get deepdiveSpeed() {
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        const sp = r[C.SPEED] || 'Lainnya';
        if (!map[sp]) map[sp] = {io:0,re:0,ps:0};
        if ((r[C.BLN_IO]||'').toString()===bulan) map[sp].io++;
        if ((r[C.BLN_RE]||'').toString()===bulan) map[sp].re++;
        if ((r[C.BLN_PS]||'').toString()===bulan) map[sp].ps++;
      });
      return Object.entries(map).map(([name,v])=>({name,...v})).sort((a,b)=>b.ps-a.ps);
    },

    get ddActiveData() {
      if (this.ddTab === 'channel') return this.deepdiveChannel;
      if (this.ddTab === 'produk')  return this.deepdiveProduk;
      return this.deepdiveSpeed;
    },

    // ── Penetrasi ODP ──
    get penetrasiKabList() {
      const clusterKabs = (CLUSTER[this.cluster]?.kab || []).map(k => k.toUpperCase());
      const set = new Set();
      this.penetrasiData.forEach(r => {
        const kab = (r.kab||'').trim().toUpperCase();
        if (clusterKabs.includes(kab)) set.add(kab);
      });
      return [...set].sort();
    },

    get penetrasiKecList() {
      const kabF = this.penetrasiKabFilter.toUpperCase();
      const set = new Set();
      this.penetrasiData.forEach(r => {
        if (kabF && (r.kab||'').trim().toUpperCase() !== kabF) return;
        const kec = (r.kec||'').trim();
        if (kec) set.add(kec);
      });
      return [...set].sort();
    },

    get penetrasiFiltered() {
      const clusterKabs = (CLUSTER[this.cluster]?.kab || []).map(k => k.toUpperCase());
      const kabF = this.penetrasiKabFilter.toUpperCase();
      const kecF = this.penetrasiKecFilter.toUpperCase();
      const { key, dir } = this.penetrasiSort;
      return this.penetrasiData.filter(r => {
        const kab = (r.kab||'').trim().toUpperCase();
        if (!clusterKabs.includes(kab)) return false;
        if (kabF && kab !== kabF) return false;
        if (kecF && (r.kec||'').trim().toUpperCase() !== kecF) return false;
        return true;
      }).sort((a,b) => {
        const av = typeof a[key]==='string' ? a[key].localeCompare(b[key]) : (a[key]||0)-(b[key]||0);
        return av * dir;
      });
    },

    async loadPenetrasiOdp() {
      if (this.penetrasiLoading) return;
      this.penetrasiLoading = true;
      this.penetrasiError = '';
      try {
        // Fetch GAS (PS per kec/desa) + odp.json (port available per kec/desa) in parallel
        const [gasRes, odpRes] = await Promise.all([
          fetch((this.cfg.gasUrl || DEFAULT_GAS_URL) + '?session_token=' + encodeURIComponent(this.sessionToken) + '&type=deepdive_odp&periode=' + encodeURIComponent(this.cfg.bulan) + '&bulan=' + periodMonth(this.cfg.bulan)),
          fetchOdpJson()
        ]);
        const gasText = await gasRes.text();
        const odpRaw  = await odpRes.json();

        let gasJson;
        try { gasJson = JSON.parse(gasText); }
        catch(e) {
          throw new Error('GAS belum di-redeploy. Buka Apps Script → Deploy ulang terlebih dahulu.');
        }
        if (!gasJson.ok) throw new Error(gasJson.error || 'GAS error');

        const allOdp = Array.isArray(odpRaw) ? odpRaw : (odpRaw.data || []);

        // Build avai map: key = "KEC||DESA" → {kab, avai_sum}
        const availMap = {};
        allOdp.forEach(o => {
          const kab  = (o.kab  || '').trim().toUpperCase();
          const kec  = (o.kec  || '').trim().toUpperCase();
          const desa = (o.desa || '').trim().toUpperCase();
          const key  = kec + '||' + desa;
          if (!availMap[key]) availMap[key] = { kab, avai: 0 };
          availMap[key].avai += parseInt(o.avai || 0, 10);
        });

        // Build PS map from GAS
        const psMap = {};
        (gasJson.data || []).forEach(r => {
          const kec  = (r.kec  || '').trim().toUpperCase();
          const desa = (r.desa || '').trim().toUpperCase();
          const key  = kec + '||' + desa;
          psMap[key] = (psMap[key] || 0) + (r.ps || 0);
        });

        // Merge
        const merged = {};
        // Add all avai entries
        Object.entries(availMap).forEach(([key, v]) => {
          if (!merged[key]) merged[key] = { kab: v.kab, kec: key.split('||')[0], desa: key.split('||')[1], avai: 0, ps: 0 };
          merged[key].avai += v.avai;
        });
        // Add PS
        Object.entries(psMap).forEach(([key, ps]) => {
          if (!merged[key]) {
            const [kec, desa] = key.split('||');
            // Try to find kab from availMap
            const kabFromAvail = availMap[key]?.kab || '';
            merged[key] = { kab: kabFromAvail, kec, desa, avai: 0, ps: 0 };
          }
          merged[key].ps += ps;
        });

        this.penetrasiData = Object.values(merged)
          .filter(r => r.kec || r.desa)
          .map(r => ({
            ...r,
            kab:  r.kab  || '',
            kec:  this._titleCase(r.kec),
            desa: this._titleCase(r.desa),
            penetrasi: r.avai > 0 ? Math.round(r.ps / r.avai * 100) : 0,
          }));
      } catch(e) {
        this.penetrasiError = e.message;
      } finally {
        this.penetrasiLoading = false;
      }
    },

    penetrasiSortBy(key) {
      if (this.penetrasiSort.key === key) {
        this.penetrasiSort.dir *= -1;
      } else {
        // Default dir: angka → desc (besar dulu), string → asc (A dulu)
        const numKeys = ['avai','ps','penetrasi'];
        this.penetrasiSort = { key, dir: numKeys.includes(key) ? -1 : 1 };
      }
    },

    downloadOrderCSV() {
      const rows = this.filteredOrders;
      if (!rows.length) return;
      const header = ['ID Order','STO','Kabupaten','Nama SF','ID SF','Jenis Produk','Speed','Channel','Tgl IO','Tgl PS','Status IO','Status PS','Kendala','Detail Kendala'];
      const lines = [header.join(',')];
      rows.forEach(r => {
        lines.push([
          '"'+(r.idOrder||'')+'"',
          '"'+(r.sto||'')+'"',
          '"'+(r.kab||'')+'"',
          '"'+(r.namaSF||'')+'"',
          '"'+(r.codeSF||'')+'"',
          '"'+(r.jenisProduk||'')+'"',
          '"'+(r.speed||'')+'"',
          '"'+(r.channel||'')+'"',
          '"'+(r.tglIO||'')+'"',
          '"'+(r.tglPS||'')+'"',
          '"'+(r.statusIO||'')+'"',
          '"'+(r.statusPS||'Proses')+'"',
          '"'+(r.kendala||'')+'"',
          '"'+(r.kendalaDetail||'')+'"',
        ].join(','));
      });
      const blob = new Blob(['﻿'+lines.join('\n')], {type:'text/csv;charset=utf-8;'});
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      const stat = this.orderStatus || 'Semua';
      const kab  = this.orderKab    || 'Semua';
      a.href     = url;
      a.download = `DetailOrder_${this.cluster}_${stat}_${kab}.csv`.replace(/\s+/g,'_');
      a.click();
      URL.revokeObjectURL(url);
    },

    downloadPenetrasiCSV() {
      const rows = this.penetrasiFiltered;
      if (!rows.length) return;
      const header = ['Kabupaten','Kecamatan','Desa','Port Avail','PS','Penetrasi (%)'];
      const lines = [header.join(',')];
      rows.forEach(r => {
        lines.push([
          '"'+(r.kab||'')+'"',
          '"'+(r.kec||'')+'"',
          '"'+(r.desa||'')+'"',
          r.avai||0,
          r.ps||0,
          r.penetrasi||0
        ].join(','));
      });
      // Footer total
      const totAvai = rows.reduce((s,r)=>s+r.avai,0);
      const totPs   = rows.reduce((s,r)=>s+r.ps,0);
      const totPct  = totAvai>0 ? Math.round(totPs/totAvai*100) : 0;
      lines.push(['"TOTAL","","",'+totAvai+','+totPs+','+totPct].join(''));

      const blob = new Blob(['﻿'+lines.join('\n')], {type:'text/csv;charset=utf-8;'});
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      const kab  = this.penetrasiKabFilter || 'Semua';
      const kec  = this.penetrasiKecFilter || 'Semua';
      a.href     = url;
      a.download = `PenetrasiODP_${this.cluster}_${kab}_${kec}.csv`.replace(/\s+/g,'_');
      a.click();
      URL.revokeObjectURL(url);
    },

    // ── Prioritas Kunjungan ──────────────────────────────────────
    loadPrioritas() {
      if (_prioInflight) return _prioInflight;
      _prioInflight = this._loadPrioritas().finally(() => { _prioInflight = null; });
      return _prioInflight;
    },
    // Ambil di latar belakang setelah dashboard tampil, agar Prioritas/Timeline/Plan DJP langsung siap
    prefetchPrioritas() {
      if (this.prioritasOdpRaw.length || _prioInflight) return;
      const go = () => { if (this.currentUser && !this.prioritasOdpRaw.length) this.loadPrioritas().catch(() => {}); };
      if ('requestIdleCallback' in window) requestIdleCallback(go, { timeout: 3000 }); else setTimeout(go, 1500);
    },

    async _loadPrioritas() {
      this.prioritasLoading = true;
      this.prioritasError = '';
      try {
        const gasBase = this.cfg.gasUrl || DEFAULT_GAS_URL;
        const bulan   = this.cfg.bulan || '';
        // Fetch odp.json + GAS deepdive_odp (PS per desa) secara paralel
        const [odpRes, gasRes] = await Promise.all([
          fetchOdpJson(),
          fetch(`${gasBase}?session_token=${encodeURIComponent(this.sessionToken)}&type=deepdive_odp&periode=${encodeURIComponent(bulan)}&bulan=${periodMonth(bulan)}`)
        ]);
        const odpRaw  = await odpRes.json();
        const gasText = await gasRes.text();

        this.prioritasOdpRaw = Array.isArray(odpRaw) ? odpRaw : (odpRaw.data || []);

        // Parse PS per desa dari GAS → simpan ke penetrasiData jika belum ada
        try {
          const gasJson = JSON.parse(gasText);
          if (gasJson.ok && !this.penetrasiData.length) {
            const allOdp = this.prioritasOdpRaw;
            const availMap = {};
            allOdp.forEach(o => {
              const kab  = (o.kab||'').trim().toUpperCase();
              const kec  = (o.kec||'').trim().toUpperCase();
              const desa = (o.desa||'').trim().toUpperCase();
              const key  = kec + '||' + desa;
              if (!availMap[key]) availMap[key] = { kab, avai: 0 };
              availMap[key].avai += parseInt(o.avai||0, 10);
            });
            const psMap = {};
            (gasJson.data || []).forEach(r => {
              const key = (r.kec||'').trim().toUpperCase() + '||' + (r.desa||'').trim().toUpperCase();
              psMap[key] = (psMap[key]||0) + (r.ps||0);
            });
            const merged = {};
            Object.entries(availMap).forEach(([key, v]) => {
              merged[key] = { kab: v.kab, kec: key.split('||')[0], desa: key.split('||')[1], avai: v.avai, ps: psMap[key]||0 };
            });
            Object.entries(psMap).forEach(([key, ps]) => {
              if (!merged[key]) merged[key] = { kab: availMap[key]?.kab||'', kec: key.split('||')[0], desa: key.split('||')[1], avai: 0, ps };
            });
            this.penetrasiData = Object.values(merged).filter(r => r.kec||r.desa).map(r => ({
              ...r, penetrasi: r.avai > 0 ? Math.round(r.ps/r.avai*100) : 0
            }));
          }
        } catch(_) { /* GAS parse error — skor tetap jalan tanpa PS */ }

      } catch(e) {
        this.prioritasError = e.message;
      } finally {
        this.prioritasLoading = false;
        if (this.page === 'prioritas') this.$nextTick(() => this.initPrioritasMap());
      }
    },

    get prioritasDesaList() {
      return this.klasifikasiDesa(this.prioritasKabFilter, this.prioritasKategoriFilter);
    },

    // Klasifikasi kuadran desa (Prioritas/Strategis/Cadangan/Skip).
    // Dipakai menu Prioritas (dengan filter kab) dan Timeline Desa (kabFilter='' → per cluster).
    klasifikasiDesa(kabFilter = '', katFilter = '') {
      const clusterKabs = (CLUSTER[this.cluster]?.kab || []).map(k => k.toUpperCase());
      const kabF = (kabFilter || '').toUpperCase();
      const data = this.prioritasOdpRaw;
      if (!data.length) return [];

      // Build PS-per-desa map dari penetrasiData (kec||desa → ps)
      const psMap = {};
      (this.penetrasiData || []).forEach(p => {
        const key = (p.kec||'').trim().toUpperCase() + '||' + (p.desa||'').trim().toUpperCase();
        psMap[key] = (psMap[key] || 0) + (p.ps || 0);
      });

      // Aggregate per desa dari odp.json
      const map = {};
      data.forEach(o => {
        const kab  = (o.kab  ||'').trim().toUpperCase();
        const kec  = (o.kec  ||'').trim();
        const desa = (o.desa ||'').trim();
        if (!clusterKabs.includes(kab)) return;
        if (kabF && kab !== kabF) return;
        const key = kab + '||' + kec + '||' + desa;
        if (!map[key]) map[key] = { kab, kec, desa, totalAvail:0, totalUsed:0, lats:[], lngs:[], odps:[] };
        map[key].totalAvail += (o.avai||0);
        map[key].totalUsed  += (o.used||0);
        if (o.lat && o.lng) { map[key].lats.push(o.lat); map[key].lngs.push(o.lng); }
        map[key].odps.push(o);
      });

      const desas = Object.values(map).map(d => {
        const lat = d.lats.length ? d.lats.reduce((a,b)=>a+b,0)/d.lats.length : null;
        const lng = d.lngs.length ? d.lngs.reduce((a,b)=>a+b,0)/d.lngs.length : null;
        // Jarak ke kantor terdekat — filter by kab desa (Cluster → Kab → Kantor)
        let nearestOffice = '—', nearestDist = 9999;
        if (lat && lng) {
          const kabUpper = d.kab.toUpperCase();
          const kabOffices = OFFICES.filter(of => of.kab.toUpperCase() === kabUpper);
          const officePool = kabOffices.length ? kabOffices : OFFICES; // fallback ke semua jika kab tidak ada kantornya
          officePool.forEach(of => {
            const dist = haversineKm(lat, lng, of.lat, of.lng);
            if (dist < nearestDist) { nearestDist = dist; nearestOffice = of.name; }
          });
        }
        const total = d.totalAvail + d.totalUsed;
        const occ_rate = total > 0 ? d.totalUsed / total : 0; // tingkat okupansi port
        // PS aktif per desa (dari penetrasiData)
        const psKey = d.kec.trim().toUpperCase() + '||' + d.desa.trim().toUpperCase();
        const psAktif = psMap[psKey] || 0;
        // sales_rate: seberapa aktif area ini sudah digarap (PS vs total port)
        const sales_rate = total > 0 ? Math.min(psAktif / total, 1) : 0;
        return { ...d, lat, lng, nearestOffice, nearestDist, occ_rate, psAktif, sales_rate, odpCount: d.odps.length };
      });

      // ── Skor Potensi (murni nilai bisnis, bebas dari jarak) ─────────────
      // Opportunity = √avail × (1 - occ_rate) × (1 - sales_rate)
      //   • √avail   : diminishing returns
      //   • (1-occ)  : port makin kosong makin potensial
      //   • (1-sales): area belum digarap makin tinggi prioritas
      // ───────────────────────────────────────────────────────────────────
      desas.forEach(d => {
        d._opp = Math.sqrt(d.totalAvail) * (1 - d.occ_rate) * (1 - d.sales_rate);
      });
      const maxOpp = Math.max(...desas.map(d => d._opp), 1e-9);
      desas.forEach(d => { d.skorPotensi = Math.round(d._opp / maxOpp * 100); });

      // ── Quadrant Matrix ─────────────────────────────────────────────────
      // Threshold: median skorPotensi & median jarak (self-adapting per cluster)
      const sortedOpp  = [...desas].map(d => d.skorPotensi).sort((a,b) => a-b);
      const sortedDist = [...desas].filter(d => d.nearestDist < 9999).map(d => d.nearestDist).sort((a,b) => a-b);
      const medOpp  = sortedOpp[Math.floor(sortedOpp.length / 2)] || 50;
      const medDist = sortedDist[Math.floor(sortedDist.length / 2)] || 30;

      const KUADRAN = {
        prioritas: { label:'Prioritas',  color:'#ef4444', order:1 },
        strategis: { label:'Strategis',  color:'#f97316', order:2 },
        cadangan:  { label:'Cadangan',   color:'#eab308', order:3 },
        skip:      { label:'Skip',       color:'#9ca3af', order:4 },
      };

      const katF = katFilter;

      return desas.map(d => {
        const tinggi = d.skorPotensi >= medOpp;
        const dekat  = d.nearestDist <= medDist;
        const kat    = tinggi && dekat  ? 'prioritas'
                     : tinggi && !dekat ? 'strategis'
                     : !tinggi && dekat ? 'cadangan'
                     :                   'skip';
        return { ...d, kat, katLabel: KUADRAN[kat].label, katColor: KUADRAN[kat].color, katOrder: KUADRAN[kat].order };
      })
      .filter(d => !katF || d.kat === katF)
      .sort((a,b) => a.katOrder - b.katOrder || b.skorPotensi - a.skorPotensi);
    },

    initPrioritasMap() {
      if (!this.prioritasOdpRaw.length) return;
      const el = document.getElementById('prioritas-map');
      if (!el) return;
      const _oldMap = window._prioritasCtx?.map || this.prioritasMap;
      if (_oldMap) { try { _oldMap.remove(); } catch(e) {} }
      this.prioritasMap = null;
      this.prioritasMarkerMap = {};
      this.prioritasSelectedMarker = null;
      // Simpan di luar Alpine reactive proxy agar Leaflet method tidak ter-wrap
      window._prioritasCtx = { markerMap: {}, selectedMarker: null, map: null };

      const map = L.map('prioritas-map', { zoomControl: true });
      L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '© OpenStreetMap', maxZoom: 18
      }).addTo(map);
      this.prioritasMap = map;
      window._prioritasCtx.map = map; // bypass Alpine proxy

      const desas = this.prioritasDesaList;
      if (!desas.length) return;

      const bounds = [];

      // Ambil semua desas tanpa filter kategori untuk peta lengkap
      const allDesas = this.prioritasDesaList.length ? this.prioritasDesaList : [];

      // Render semua desa (peta tidak difilter kuadran, biar konteks penuh)
      const mapData = (() => {
        // Re-compute tanpa filter kategori untuk peta
        const tmp = this.prioritasKategoriFilter;
        return allDesas; // sudah terfilter, tapi peta tetap tampilkan semuanya
      })();

      desas.forEach(d => {
        if (!d.lat || !d.lng) return;
        const color  = d.katColor || '#9ca3af';
        const radius = 6 + Math.round(Math.sqrt(d.totalAvail));

        const marker = L.circleMarker([d.lat, d.lng], {
          radius: Math.min(Math.max(radius, 6), 20),
          fillColor: color, color: '#fff',
          weight: 1.5, opacity: 1, fillOpacity: 0.85
        }).addTo(map);

        marker.bindPopup(`
          <b>${esc(d.desa)}</b> <span style="display:inline-block;padding:1px 6px;border-radius:10px;font-size:10px;font-weight:700;color:#fff;background:${color}">${d.katLabel}</span><br>
          ${esc(d.kec)} · ${esc(d.kab)}<br>
          <hr style="margin:4px 0">
          Avail: <b>${d.totalAvail}</b> port · Occ: <b>${Math.round(d.occ_rate*100)}%</b><br>
          PS Aktif: <b>${d.psAktif||0}</b> · Jarak: <b>${d.nearestDist.toFixed(1)} km</b> dari ${esc(d.nearestOffice)}<br>
          Skor Potensi: <b>${d.skorPotensi}</b>
        `);

        // Simpan referensi marker di luar Alpine proxy
        const r = Math.min(Math.max(radius, 6), 20);
        const entry = { marker, color, radius: r };
        this.prioritasMarkerMap[d.kec+'||'+d.desa] = entry;
        if (window._prioritasCtx) window._prioritasCtx.markerMap[d.kec+'||'+d.desa] = entry;

        // Tambah kantor sebagai ikon bintang
        bounds.push([d.lat, d.lng]);
      });

      // Tambah marker kantor
      OFFICES.forEach(of => {
        L.marker([of.lat, of.lng], {
          icon: L.divIcon({ className:'', html:`<div style="background:#4f8ef7;color:#fff;border-radius:50%;width:28px;height:28px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,0.4)">K</div>`, iconSize:[28,28], iconAnchor:[14,14] })
        }).addTo(map).bindPopup(`<b>Kantor</b><br>${esc(of.name)}`);
      });

      if (bounds.length) map.fitBounds(bounds, { padding: [30, 30] });
      this.prioritasMarkers = desas;
    },

    prioritasFlyTo(desa) {
      const ctx = window._prioritasCtx || { markerMap: {}, selectedMarker: null };

      // Reset marker sebelumnya ke style normal
      const prev = ctx.selectedMarker;
      if (prev) {
        prev.marker.setStyle({
          radius: prev.radius,
          fillColor: prev.color,
          color: '#fff',
          weight: 1.5,
          fillOpacity: 0.85
        });
        prev.marker.closePopup();
      }

      this.prioritasSelectedDesa = desa;

      // Highlight marker yang dipilih
      const key = desa.kec + '||' + desa.desa;
      const entry = ctx.markerMap[key];
      if (entry) {
        entry.marker.setStyle({
          radius: entry.radius + 6,
          fillColor: entry.color,
          color: '#fff',
          weight: 3,
          fillOpacity: 1
        });
        entry.marker.bringToFront();
        entry.marker.openPopup();
        ctx.selectedMarker = entry;
      }

      // Alpine.raw() unwrap proxy agar method Leaflet bisa dipanggil langsung
      const lmap = ctx.map || (this.prioritasMap ? Alpine.raw(this.prioritasMap) : null);
      if (lmap && desa.lat && desa.lng) {
        if (window.innerWidth < 768) {
          const mapEl = document.getElementById('prioritas-map');
          if (mapEl) {
            mapEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
            setTimeout(() => {
              lmap.invalidateSize();
              lmap.flyTo([desa.lat, desa.lng], 14, { duration: 1 });
            }, 400);
          }
        } else {
          lmap.flyTo([desa.lat, desa.lng], 14, { duration: 1 });
        }
      }
    },

    _titleCase(str) {
      if (!str) return '';
      return str.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    },

    // ── Outlook ──
    get outlookData() {
      const rows = this.filteredRows;
      const bulan = this.cfg.bulan;
      const prevBulan = this.prevBulan.toString();
      const today = new Date();
      const hariPS = this.hariBerjalan;
      const daysInMonth = daysInPeriod(bulan);
      const sisaHari = Math.max(daysInMonth - hariPS, 0);
      const kabs = CLUSTER[this.cluster]?.kab || [];
      const map = {};
      kabs.forEach(k => { map[k] = {mtd:0, prev:0, drr:0}; });

      rows.forEach(r => {
        const kab = (r[C.VAL_KAB]||'').trim().toUpperCase();
        if (!map[kab]) return;
        if ((r[C.BLN_PS]||'').toString()===bulan) map[kab].mtd++;
        if ((r[C.BLN_PS]||'').toString()===prevBulan) map[kab].prev++;
      });

      return kabs.map(k => {
        const d = map[k];
        const drr = hariPS > 0 ? d.mtd / hariPS : 0;
        const simDrr = parseFloat(this.simDrr)||0;
        const effDrr = simDrr > 0 ? simDrr : drr;
        const olEom = Math.round(d.mtd + effDrr * sisaHari);
        const tgtPS = this.cfg.targets[k]?.ps||0;
        const vsTarget = tgtPS > 0 ? Math.round(olEom/tgtPS*100) : null;
        const vsLM = d.prev > 0 ? Math.round((olEom-d.prev)/d.prev*100) : null;
        return { kab:k, mtd:d.mtd, drr:drr.toFixed(2), olEom, tgtPS, vsTarget, vsLM, prev:d.prev };
      });
    },

    get outlookMeta() {
      const today = new Date();
      const bulan = this.cfg.bulan;
      const daysInMonth = daysInPeriod(bulan);
      const hariPS = this.hariBerjalan;
      const sisaHari = Math.max(daysInMonth - hariPS, 0);
      return { daysInMonth, hariPS, sisaHari };
    },

    // ── SF Performance (New Page) ──
    get sfperfKpi() {
      // Gunakan sfTable yang sudah fix prevPS dari filteredPrevRows
      const all = this.sfTable;
      const total = all.length;
      const aktif = all.filter(s=>s.ps>=1).length;
      const produktif = all.filter(s=>s.ps>=10).length;
      const baru = all.filter(s=>s.prevPS===0 && s.ps>=1).length;
      return { total, aktif, produktif, baru };
    },

    get sfperfScorecard() {
      const kabs = CLUSTER[this.cluster]?.kab || [];
      const kabSF = {};
      kabs.forEach(k => { kabSF[k] = []; });
      this.sfTable.forEach(sf => {
        const k = (sf.kab||'').trim().toUpperCase();
        if (kabSF[k]) kabSF[k].push(sf);
      });
      return kabs.map(k => {
        const sfs = kabSF[k] || [];
        const total = sfs.length;
        const aktif = sfs.filter(s=>s.ps>=1).length;
        const produktif = sfs.filter(s=>s.ps>=10).length;
        const baru = sfs.filter(s=>s.prevPS===0&&s.ps>=1).length;
        return { kab:k, total, aktif, produktif, baru };
      });
    },

    get sfperfTierDist() {
      // Tier LM (berdasarkan prevPS dari filteredPrevRows — via sfTable)
      const distLM = {};
      // Tier sekarang (berdasarkan PS bulan ini)
      const distNow = {};
      TIER_CONFIG.forEach(t => { distLM[t.name] = 0; distNow[t.name] = 0; });
      this.sfTable.forEach(sf => {
        distLM[sf.tier]++;
        distNow[sf.tierNow]++;
      });
      return { lm: distLM, now: distNow };
    },

    // ── Orders ──
    get filteredOrders() {
      let rows = this.filteredRows;
      const q = this.orderSearch.toLowerCase();
      const bulan = this.cfg.bulan;

      rows = rows.map(r => ({
        idOrder:   r[C.ID_ORDER]||'',
        sto:       r[C.STO]||'',
        kab:       r[C.VAL_KAB]||'',
        namaSF:    r[C.NAMA_SF]||'',
        codeSF:    r[C.CODE_SF]||'',
        jenisProduk: r[C.PRODUK_KAT]||'',
        speed:     (r[C.SPEED]||'').toString(),
        produk:    r[C.SPEED]||r[C.PRODUK_KAT]||'',
        produkKey: ((r[C.PRODUK_KAT]||'').trim()) + '|' + ((r[C.SPEED]||'').toString().trim()),
        channel:   r[C.CHANNEL2]||'',
        statusIO:  r[C.VAL_IO]||'',
        statusPS:  r[C.VAL_PS]||'',
        statusFnl: r[C.STATUS_FNL]||'',
        tglIO:     (function(d){ if(!d||typeof d!=='string'||d.length!==8) return d||''; return d.slice(6)+'/'+d.slice(4,6)+'/'+d.slice(0,4); })(r[C.DATE_IO]),
        tglPS:     (function(d){ if(!d||typeof d!=='string'||d.length!==8) return d||''; return d.slice(6)+'/'+d.slice(4,6)+'/'+d.slice(0,4); })(r[C.DATE_PS]),
        _dateIO:   parseInt((r[C.DATE_IO]||'0').toString().replace(/\D/g,''))||0,
        _datePS:   parseInt((r[C.DATE_PS]||'0').toString().replace(/\D/g,''))||0,
        kendala:   r[C.KENDALA_CAT]||'',
        kendalaDetail: r[C.KENDALA_DTL]||'',
        blnIO:     (r[C.BLN_IO]||'').toString(),
      }));

      if (q) rows = rows.filter(r =>
        r.idOrder.toLowerCase().includes(q) ||
        r.namaSF.toLowerCase().includes(q) ||
        r.kendala.toLowerCase().includes(q)
      );
      if (this.orderKab) rows = rows.filter(r => r.kab.toUpperCase()===this.orderKab.toUpperCase());
      if (this.orderProduk) rows = rows.filter(r => r.produkKey===this.orderProduk);
      if (this.orderStatus === 'Put In Service') rows = rows.filter(r => r.statusPS==='Put In Service');
      else if (this.orderStatus === 'Valid IO')  rows = rows.filter(r => r.statusIO==='Valid IO');
      else if (this.orderStatus === 'Masih IO')  rows = rows.filter(r => !r.statusIO || r.statusIO==='Masih IO');
      else if (this.orderStatus === 'CANCEL')     rows = rows.filter(r => (r.statusFnl||'').includes('CANCEL'));

      const { key, dir } = this.orderSort;
      const dateKeys = { tglIO: '_dateIO', tglPS: '_datePS' };
      rows.sort((a, b) => {
        const sk = dateKeys[key] || key;
        const av = typeof a[sk] === 'string' ? a[sk].localeCompare(b[sk]) : (a[sk]||0)-(b[sk]||0);
        return av * dir;
      });

      return rows;
    },

    orderSortBy(key) {
      if (this.orderSort.key === key) {
        this.orderSort.dir *= -1;
      } else {
        this.orderSort = { key, dir: 1 };
      }
      this.orderPage = 0;
    },

    get paginatedOrders() {
      const start = this.orderPage * 50;
      return this.filteredOrders.slice(start, start + 50);
    },

    get produkList() {
      // Gabung PRODUK_KAT (CH) + SPEED (DT) jadi key unik, hanya dari bulan aktif
      const bulan = this.cfg.bulan;
      const map = new Map(); // key → label
      this.filteredRows.forEach(r => {
        const blnIO = (r[C.BLN_IO]||'').toString();
        const blnRE = (r[C.BLN_RE]||'').toString();
        const blnPS = (r[C.BLN_PS]||'').toString();
        if (blnIO !== bulan && blnRE !== bulan && blnPS !== bulan) return;
        const kat   = (r[C.PRODUK_KAT]||'').trim();
        const speed = (r[C.SPEED]||'').toString().trim();
        if (!kat && !speed) return;
        const key   = kat + '|' + speed;
        if (!map.has(key)) {
          const label = kat && speed ? kat + ' · ' + speed + ' Mbps'
                      : kat ? kat
                      : speed + ' Mbps';
          map.set(key, label);
        }
      });
      return [...map.entries()].sort((a,b)=>a[1].localeCompare(b[1])).map(([key,label])=>({key,label}));
    },

    get stoList() {
      const s = new Set();
      this.filteredRows.forEach(r => { const v=(r[C.STO]||'').trim(); if(v) s.add(v); });
      return [...s].sort();
    },

    // ── Kendala ──
    get filteredKendala() {
      let rows = this.filteredRows.filter(r => r[C.KENDALA_CAT] || r[C.KENDALA_DTL]);
      rows = rows.map(r => ({
        idOrder:      r[C.ID_ORDER]||'',
        namaSF:       r[C.NAMA_SF]||'',
        tglKendala:   r[C.TGL_KENDALA]||'',
        kategori:     r[C.KENDALA_CAT]||'',
        kendalaDetail:r[C.KENDALA_DTL]||'',
      }));
      const q = this.kendalaSearch.toLowerCase();
      if (q) rows = rows.filter(r => r.kategori.toLowerCase().includes(q) || r.kendalaDetail.toLowerCase().includes(q));
      if (this.kendalaCat) rows = rows.filter(r => r.kategori===this.kendalaCat);
      return rows;
    },

    get kendalaCategories() {
      const s = new Set();
      this.filteredRows.forEach(r => { const v=r[C.KENDALA_CAT]; if(v) s.add(v); });
      return [...s].sort();
    },

    // ── Koordinat rows ──
    get koordinatRows() {
      const bulan = this.cfg.bulan;
      const tglA  = parseInt(this.cfg.tglAwal)  || 1;
      const tglB  = parseInt(this.cfg.tglAkhir) || new Date().getDate();
      const fSTO    = this.petaFilterSTO;
      const fProduk = this.petaFilterProduk;
      return this.filteredRows.filter(r => {
        // Hanya tampilkan order yang IO-nya di bulan aktif
        const blnIO = (r[C.BLN_IO]||'').toString();
        if (blnIO !== bulan) return false;
        // Filter tanggal berdasarkan tgl IO
        const tgl = parseInt(r[C.TGL_IO]||0);
        if (tgl < tglA || tgl > tglB) return false;
        // Filter STO (multi)
        if (fSTO.length > 0 && !fSTO.includes((r[C.STO]||'').trim())) return false;
        // Filter Produk (key = "PRODUK_KAT|SPEED")
        if (fProduk) {
          const key = (r[C.PRODUK_KAT]||'').trim() + '|' + (r[C.SPEED]||'').toString().trim();
          if (key !== fProduk) return false;
        }
        // Harus punya koordinat valid
        const dr = r[C.KOORDINAT]||'';
        const parts = dr.split(',');
        return parts.length===2 && parseFloat(parts[0]) && parseFloat(parts[1]);
      }).map(r => {
        const [lat, lng] = (r[C.KOORDINAT]||'').split(',').map(Number);
        // Status peta berdasarkan CP (VAL_PS) dan CZ (KENDALA_CAT)
        const valPS   = (r[C.VAL_PS]||'').toString().trim();
        const kendala = (r[C.KENDALA_CAT]||'').toString().trim();
        const sudahPS = valPS.toLowerCase().includes('put in service');
        // type: 'ps' = hijau, 'kendala' = merah, 'proses' = kuning
        const type = sudahPS ? 'ps' : (kendala ? 'kendala' : 'proses');
        const statusLabel = sudahPS ? 'Put In Service' : (kendala ? ('Kendala: ' + kendala) : 'Belum PS');
        return { lat, lng, status: statusLabel, type, kendala, sf: r[C.NAMA_SF]||'', id: r[C.ID_ORDER]||'' };
      });
    },

    // ── CHARTS ──
    renderCharts() {
      if (this.page === 'dashboard') {
        this.renderTren();
        this.renderProduk();
        this.renderChannel();
        this.renderTime();
        this.renderHP();
      } else if (this.page === 'd2d') {
        this.renderD2dTren();
      } else if (this.page === 'deepdive') {
        this.renderDdChart();
      } else if (this.page === 'outlook') {
        this.renderOlTren();
      } else if (this.page === 'sfperf') {
        if (this.sfpTab === 'tier') this.renderSfpTier();
      } else if (this.page === 'kendala') {
        this.renderKendalaCat();
        setTimeout(() => this.renderKendalaTren(), 0);
      } else if (this.page === 'peta') {
        this.renderPeta();
      }
    },

    getOrInit(id, forPie) {
      const el = document.getElementById(id);
      if (!el) return null;
      if (!this._charts[id]) {
        this._charts[id] = echarts.init(el, null, { renderer: 'svg' });
      }
      return this._charts[id];
    },

    disposeCharts() {
      Object.values(this._charts).forEach(c => { try { c && c.dispose && c.dispose(); } catch(e){} });
      this._charts = {};
    },

    renderTren() {
      const chart = this.getOrInit('chart-tren');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const dayCount = daysInPeriod(bulan);
      const days = Array.from({length: dayCount}, (_,i) => i+1);
      const io = new Array(dayCount).fill(0);
      const re = new Array(dayCount).fill(0);
      const ps = new Array(dayCount).fill(0);
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_IO]||'').toString()===bulan) { const d=parseInt(r[C.TGL_IO]||0)-1; if(d>=0&&d<dayCount) io[d]++; }
        if ((r[C.BLN_RE]||'').toString()===bulan) { const d=parseInt(r[C.TGL_RE]||0)-1; if(d>=0&&d<dayCount) re[d]++; }
        if ((r[C.BLN_PS]||'').toString()===bulan) { const d=parseInt(r[C.TGL_PS]||0)-1; if(d>=0&&d<dayCount) ps[d]++; }
      });
      // 7-day moving average helper
      const ma7 = (arr) => arr.map((_, i) => {
        const slice = arr.slice(Math.max(0, i-6), i+1);
        return parseFloat((slice.reduce((s,v)=>s+v,0)/slice.length).toFixed(1));
      });
      chart.setOption(Object.assign({}, getChartTheme(), {
        legend: { data:['IO','RE','PS','PS MA7'], top: 0, ...getChartTheme().legend },
        xAxis: { ...getChartTheme().xAxis, type:'category', data: days.map(d=>d) },
        yAxis: { ...getChartTheme().yAxis, type:'value' },
        series: [
          { name:'IO', type:'line', data:io, smooth:true, lineStyle:{width:2}, color:'#fbbf24', symbol:'none' },
          { name:'RE', type:'line', data:re, smooth:true, lineStyle:{width:2}, color:'#58a6ff', symbol:'none' },
          { name:'PS', type:'line', data:ps, smooth:true, lineStyle:{width:2,color:'#22d972'}, areaStyle:{color:'rgba(34,217,114,.08)'}, color:'#22d972', symbol:'none' },
          { name:'PS MA7', type:'line', data:ma7(ps), smooth:true, lineStyle:{width:2,type:'dashed',color:'#a78bfa'}, color:'#a78bfa', symbol:'none', tooltip:{show:true} },
        ],
        tooltip: { ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'cross'} },
      }));
    },

    renderProduk() {
      const chart = this.getOrInit('chart-produk');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString() !== bulan) return;
        if (!(r[C.VAL_PS]||'').toString().trim()) return;
        const v = r[C.SPEED] || r[C.PRODUK_KAT] || 'Lainnya';
        map[v] = (map[v]||0) + 1;
      });
      // Limit to top 7 + group Lainnya to avoid label collision
      const sorted = Object.entries(map).sort((a,b)=>b[1]-a[1]);
      const TOP = 7;
      let data = sorted.slice(0, TOP).map(([n,v])=>({name:n, value:v}));
      const rest = sorted.slice(TOP).reduce((s,[,v])=>s+v, 0);
      if (rest > 0) data.push({name:'Lainnya', value:rest});
      chart.setOption(Object.assign({}, (function(){const {xAxis,yAxis,grid,...t}=getChartTheme();return t;}()), {
        legend: { type:'scroll', orient:'vertical', right:4, top:'middle', textStyle:{color:'#8b949e',fontSize:10}, pageTextStyle:{color:'#8b949e'} },
        tooltip: { ...(function(){const {xAxis,yAxis,grid,...t}=getChartTheme();return t;}()).tooltip, trigger:'item', formatter:'{b}: {c} ({d}%)' },
        series: [{
          type:'pie', radius:['45%','70%'], center:['40%','52%'],
          data,
          label:{ show:false },
          labelLine:{ show:false },
          itemStyle: { borderColor:'#0d1117', borderWidth:2 },
          emphasis:{ label:{ show:true, fontSize:11, formatter:'{b}\n{c} ({d}%)', color:'#e6edf3' } },
        }],
      }));
    },

    renderChannel() {
      const chart = this.getOrInit('chart-channel');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString() !== bulan) return;
        if (!(r[C.VAL_PS]||'').toString().trim()) return;
        const v=r[C.CHANNEL2]||'Lainnya'; map[v]=(map[v]||0)+1;
      });
      const entries = Object.entries(map).sort((a,b)=>b[1]-a[1]);
      chart.setOption(Object.assign({}, getChartTheme(), {
        tooltip: { ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'shadow'} },
        xAxis: { ...getChartTheme().xAxis, type:'value' },
        yAxis: { ...getChartTheme().yAxis, type:'category', data: entries.map(e=>e[0]) },
        series: [{ type:'bar', data: entries.map(e=>e[1]), barMaxWidth:20, itemStyle:{borderRadius:[0,4,4,0], color:'#00c2ff'} }],
        grid: { left:80, right:20, top:10, bottom:10 },
      }));
    },

    renderTime() {
      const chart = this.getOrInit('chart-time');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_IO]||'').toString() !== bulan) return;
        if (!(r[C.VAL_IO]||'').toString().trim()) return;
        const v=r[C.TIME_RANGE]||'Lainnya'; map[v]=(map[v]||0)+1;
      });
      const order = ['07:00-09:59','10:00-11:59','12:00-13:59','14:00-15:59','16:00-17:59','18:00-20:59','21:00-23:59'];
      const cats = order.filter(k=>map[k]).concat(Object.keys(map).filter(k=>!order.includes(k)));
      chart.setOption(Object.assign({}, getChartTheme(), {
        tooltip: { ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'shadow'} },
        xAxis: { ...getChartTheme().xAxis, type:'category', data: cats },
        yAxis: { ...getChartTheme().yAxis, type:'value' },
        series: [{ type:'bar', data: cats.map(c=>map[c]||0), barMaxWidth:24, itemStyle:{borderRadius:[4,4,0,0], color:'#a855f7'} }],
        grid: { left:30, right:10, top:10, bottom:30, containLabel:true },
      }));
    },

    renderHP() {
      const chart = this.getOrInit('chart-hp');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      let hp=0, nonHp=0;
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString() !== bulan) return;
        if (!(r[C.VAL_PS]||'').toString().trim()) return;
        (r[C.HOMEPASS]==='Homepass' ? hp++ : nonHp++);
      });
      chart.setOption(Object.assign({}, (function(){const {xAxis,yAxis,grid,...t}=getChartTheme();return t;}()), {
        tooltip: { ...(function(){const {xAxis,yAxis,grid,...t}=getChartTheme();return t;}()).tooltip, trigger:'item' },
        series: [{
          type:'pie', radius:['45%','70%'], center:['50%','50%'],
          data:[{name:'Homepass',value:hp},{name:'Non-HP',value:nonHp}],
          color:['#58a6ff','#30363d'],
          label:{ show:true, fontSize:10, color:'#8b949e', formatter:'{b}\n{c} ({d}%)' },
          labelLine:{ lineStyle:{color:'#484f58'} },
          itemStyle:{ borderColor:'#0d1117', borderWidth:2 },
        }],
        legend:{ bottom:0, ...(function(){const {xAxis,yAxis,grid,...t}=getChartTheme();return t;}()).legend },
      }));
    },

    renderKendalaCat() {
      const chart = this.getOrInit('chart-kendala-cat');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        const blnIO = (r[C.BLN_IO]||'').toString();
        const blnRE = (r[C.BLN_RE]||'').toString();
        const blnPS = (r[C.BLN_PS]||'').toString();
        if (blnIO !== bulan && blnRE !== bulan && blnPS !== bulan) return;
        const v=r[C.KENDALA_CAT]; if(v) map[v]=(map[v]||0)+1;
      });
      const entries = Object.entries(map).sort((a,b)=>b[1]-a[1]);
      chart.setOption(Object.assign({}, getChartTheme(), {
        tooltip: { ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'shadow'} },
        xAxis: { ...getChartTheme().xAxis, type:'value' },
        yAxis: { ...getChartTheme().yAxis, type:'category', data:entries.map(e=>e[0]) },
        series:[{ type:'bar', data:entries.map(e=>e[1]), barMaxWidth:20, itemStyle:{borderRadius:[0,4,4,0], color:'#ff4d4d'} }],
        grid:{ left:130, right:20, top:10, bottom:10 },
      }));
    },

    renderKendalaTren() {
      const chart = this.getOrInit('chart-kendala-tren');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const map = {};
      this.filteredRows.forEach(r => {
        const blnIO = (r[C.BLN_IO]||'').toString();
        const blnRE = (r[C.BLN_RE]||'').toString();
        const blnPS = (r[C.BLN_PS]||'').toString();
        if (blnIO !== bulan && blnRE !== bulan && blnPS !== bulan) return;
        const d = r[C.TGL_KENDALA]; const cat = r[C.KENDALA_CAT];
        if (d && cat) { const key=d.toString().slice(0,8); map[key]=(map[key]||0)+1; }
      });
      const keys = Object.keys(map).sort();
      chart.setOption(Object.assign({}, getChartTheme(), {
        tooltip:{ ...getChartTheme().tooltip, trigger:'axis' },
        xAxis:{ ...getChartTheme().xAxis, type:'category', data:keys.map(k=>k.slice(6)) },
        yAxis:{ ...getChartTheme().yAxis, type:'value' },
        series:[{ type:'bar', data:keys.map(k=>map[k]), barMaxWidth:20, itemStyle:{borderRadius:[4,4,0,0], color:'#ff4d4d'} }],
        grid:{ left:30, right:10, top:10, bottom:30, containLabel:true },
      }));
    },

    renderPeta() {
      // Use Leaflet.js + OpenStreetMap
      const rows = this.koordinatRows;
      const mapEl = document.getElementById('leaflet-map');
      if (!mapEl) return;
      if (!window.L) { mapEl.innerHTML = '<p style="padding:20px;color:var(--text-3)">Leaflet belum tersedia</p>'; return; }

      // Destroy existing map instance to avoid "already initialized" error
      if (this._leafletMap) {
        this._leafletMap.remove();
        this._leafletMap = null;
        this._leafletMarkerLayer = null;
        this._odpLayer = null;  // reset karena map dihancurkan
        this._odpCanvas = null; // canvas renderer juga ikut map lama
      }

      // Default center: Lampung Barat region
      const DEFAULT_LAT = -5.0, DEFAULT_LNG = 104.2;
      const map = L.map('leaflet-map').setView([DEFAULT_LAT, DEFAULT_LNG], 9);
      this._leafletMap = map;

      const isSatellite = this.mapLayer === 'satellite';
      mapEl.classList.toggle('satellite-mode', isSatellite);
      if (isSatellite) {
        L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
          attribution: '© <a href="https://www.esri.com/">Esri</a> World Imagery',
          maxZoom: 19,
        }).addTo(map);
        // Label overlay agar nama jalan/desa tetap terbaca
        L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}', {
          attribution: '',
          maxZoom: 19,
          opacity: 0.8,
        }).addTo(map);
      } else {
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
          attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
          maxZoom: 18,
        }).addTo(map);
      }

      if (rows.length === 0) {
        map.setView([DEFAULT_LAT, DEFAULT_LNG], 9);
        return;
      }

      // Warna marker: ps=hijau, kendala=merah, proses=kuning
      const typeColor = { ps: '#22c55e', kendala: '#ef4444', proses: '#fbbf24' };

      // Marker layer group
      const markers = L.layerGroup().addTo(map);
      this._leafletMarkerLayer = markers;
      const pf = this.petaFilter;

      const bounds = [];
      rows.forEach(r => {
        if (!r.lat || !r.lng) return;
        // Filter by checkbox
        if (r.type === 'ps'      && !pf.ps) return;
        if (r.type === 'kendala' && !pf.re) return;  // RE slot = kendala
        if (r.type === 'proses'  && !pf.io) return;  // IO slot = proses

        const color = typeColor[r.type] || '#fbbf24';
        const circle = L.circleMarker([r.lat, r.lng], {
          radius: r.type === 'ps' ? 8 : 6,
          fillColor: color,
          color: '#fff',
          weight: 1,
          opacity: 0.9,
          fillOpacity: 0.85,
        });
        circle.bindPopup(`<div style="font-family:'Poppins',sans-serif;font-size:12px;line-height:1.7">
          <b>${esc(r.id)}</b><br>
          SF: <b>${esc(r.sf)}</b><br>
          Status: <b style="color:${color}">${esc(r.status)}</b><br>
          <small>${r.lat.toFixed(5)}, ${r.lng.toFixed(5)}</small>
        </div>`);
        markers.addLayer(circle);
        bounds.push([r.lat, r.lng]);
      });

      if (bounds.length > 0) {
        map.fitBounds(L.latLngBounds(bounds), {padding: [20,20]});
      }

      // Render ODP layer jika toggle aktif
      if (this.showODP) this.renderODPLayer();
    },

    async toggleODPLayer() {
      if (!this._leafletMap) return;
      if (!this.showODP) {
        // Sembunyikan layer
        if (this._odpLayer) { this._odpLayer.clearLayers(); }
        return;
      }
      await this.renderODPLayer();
    },

    async renderODPLayer() {
      const map = this._leafletMap;
      if (!map) return;

      // Canvas renderer — 1 canvas element untuk semua ODP, bukan 4000 DOM div
      if (!this._odpCanvas || !map.hasLayer(this._odpCanvas)) {
        this._odpCanvas = L.canvas({ padding: 0.5 });
      }
      // Selalu buat ulang layer group (map mungkin sudah diganti)
      if (this._odpLayer) {
        this._odpLayer.clearLayers();
        if (!map.hasLayer(this._odpLayer)) {
          this._odpLayer.addTo(map);
        }
      } else {
        this._odpLayer = L.layerGroup().addTo(map);
      }

      // Fetch JSON (cache)
      if (!this._odpData) {
        try {
          // Dari situs sendiri (cache bersama), bukan raw.githubusercontent.com
          this._odpData = await getOdpJson();
        } catch(e) {
          console.error('Gagal load ODP JSON:', e);
          return;
        }
      }

      const colorMap = {
        'GREEN':  '#22c55e',
        'YELLOW': '#eab308',
        'ORANGE': '#f97316',
        'RED':    '#ef4444',
        'BLACK':  '#6b7280',
      };

      (this._odpData.data || []).forEach(o => {
        if (!o.lat || !o.lng) return;
        const clr = colorMap[(o.occ||'').toUpperCase()] || '#6b7280';
        // circleMarker dgn canvas renderer — ringan, 1 canvas untuk 4000+ titik
        // radius 4 (lebih kecil dari order radius 8) + stroke berwarna beda dari order
        const circle = L.circleMarker([o.lat, o.lng], {
          renderer: this._odpCanvas,
          radius: 2,
          fillColor: clr,
          color: '#222',    // stroke hitam tipis — beda dari order yg stroke putih
          weight: 1,
          fillOpacity: 0.8,
          opacity: 0.9,
        });
        circle.bindPopup(`<div style="font-family:'Poppins',sans-serif;font-size:12px;line-height:1.8">
          <b>${esc(o.odp || '-')}</b><br>
          Avail: <b>${esc(o.avai ?? '-')}</b> &nbsp;|&nbsp; Total Port: <b>${esc(o.occ ?? '-')}</b><br>
          <small style="color:#888">${esc(o.kec || '')}, ${esc(o.kab || '')}</small>
        </div>`);
        this._odpLayer.addLayer(circle);
      });
    },

    // ── D2D Trend Chart ──
    renderD2dTren() {
      const chart = this.getOrInit('chart-d2d-tren');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const dayCount = daysInPeriod(bulan);
      const days = Array.from({length: dayCount}, (_,i) => i+1);
      const ps = new Array(dayCount).fill(0);
      const cum = new Array(dayCount).fill(0);
      let running = 0;
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString()===bulan) {
          const d = parseInt(r[C.TGL_PS]||0)-1;
          if (d>=0&&d<dayCount) ps[d]++;
        }
      });
      for (let i=0;i<dayCount;i++) { running+=ps[i]; cum[i]=running; }
      chart.setOption(Object.assign({}, getChartTheme(), {
        legend:{ data:['PS Harian','Kumulatif'], top:0, ...getChartTheme().legend },
        xAxis:{ ...getChartTheme().xAxis, type:'category', data:days },
        yAxis:{ ...getChartTheme().yAxis, type:'value' },
        tooltip:{ ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'cross'} },
        series:[
          { name:'PS Harian', type:'bar', data:ps, barMaxWidth:16, itemStyle:{borderRadius:[3,3,0,0], color:'rgba(34,217,114,.6)'} },
          { name:'Kumulatif', type:'line', data:cum, smooth:true, lineStyle:{width:2,color:'#22d972'}, symbol:'none', yAxisIndex:0, color:'#22d972' },
        ],
      }));
    },

    // ── Deepdive Charts ──
    renderDdChart() {
      const chart = this.getOrInit('chart-dd-main');
      if (!chart) return;
      const data = this.ddActiveData;
      const names = data.map(d=>d.name);
      chart.setOption(Object.assign({}, getChartTheme(), {
        legend:{ data:['IO','RE','PS'], top:0, ...getChartTheme().legend },
        tooltip:{ ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'shadow'} },
        grid:{ left:90, right:20, top:30, bottom:10, containLabel:false },
        xAxis:{ ...getChartTheme().xAxis, type:'value' },
        yAxis:{ ...getChartTheme().yAxis, type:'category', data:names.slice().reverse() },
        series:[
          { name:'IO', type:'bar', data:names.slice().reverse().map(n=>{const d=data.find(x=>x.name===n);return d?d.io:0;}), barMaxWidth:12, itemStyle:{borderRadius:[0,4,4,0], color:'#fbbf24'} },
          { name:'RE', type:'bar', data:names.slice().reverse().map(n=>{const d=data.find(x=>x.name===n);return d?d.re:0;}), barMaxWidth:12, itemStyle:{borderRadius:[0,4,4,0], color:'#58a6ff'} },
          { name:'PS', type:'bar', data:names.slice().reverse().map(n=>{const d=data.find(x=>x.name===n);return d?d.ps:0;}), barMaxWidth:12, itemStyle:{borderRadius:[0,4,4,0], color:'#22d972'} },
        ],
      }));
    },

    // ── Outlook Trend ──
    renderOlTren() {
      const chart = this.getOrInit('chart-ol-tren');
      if (!chart) return;
      const bulan = this.cfg.bulan;
      const today = new Date();
      const daysInMonth = daysInPeriod(bulan);
      const hariPS = this.hariBerjalan;
      const days = Array.from({length: daysInMonth}, (_,i)=>i+1);
      const ps = new Array(daysInMonth).fill(null);
      const cum = new Array(daysInMonth).fill(null);
      const proj = new Array(daysInMonth).fill(null);
      let running = 0;
      const dayMap = {};
      this.filteredRows.forEach(r => {
        if ((r[C.BLN_PS]||'').toString()===bulan) {
          const d = parseInt(r[C.TGL_PS]||0)-1;
          if (d>=0&&d<daysInMonth) dayMap[d]=(dayMap[d]||0)+1;
        }
      });
      for (let i=0;i<daysInMonth;i++) {
        if (i < hariPS) { const v=dayMap[i]||0; ps[i]=v; running+=v; cum[i]=running; }
      }
      const drr = hariPS > 0 ? running / hariPS : 0;

      // Weighted DRR: Senin–Jumat bobot lebih tinggi (ratio weekday vs weekend PS)
      // Bobot: Senin=1.2, Selasa=1.2, Rabu=1.15, Kamis=1.1, Jumat=1.1, Sabtu=0.8, Minggu=0.7
      const dayWeights = [0.7, 1.2, 1.2, 1.15, 1.1, 1.1, 0.8]; // Sun=0, Mon=1, ..., Sat=6
      let weightedSum = 0, weightedCount = 0;
      for (let i = 0; i < hariPS; i++) {
        const dow = periodDate(bulan, i+1).getDay();
        const w = dayWeights[dow];
        weightedSum += (dayMap[i]||0) / w;
        weightedCount += 1;
      }
      const drrWeighted = weightedCount > 0 ? weightedSum / weightedCount : drr;

      // Projection line from current MTD forward (both simple and weighted)
      const projWeighted = new Array(daysInMonth).fill(null);
      for (let i = hariPS-1; i < daysInMonth; i++) {
        proj[i] = Math.round(running + drr * (i - hariPS + 1));
        // Weighted: multiply DRR by today's day weight for remaining days
        let wEOM = running;
        for (let j = hariPS; j <= i; j++) {
          const dow = periodDate(bulan, j+1).getDay();
          wEOM += drrWeighted * dayWeights[dow];
        }
        projWeighted[i] = Math.round(wEOM);
      }
      chart.setOption(Object.assign({}, getChartTheme(), {
        legend:{ data:['PS Kumulatif','Proyeksi EOM','Proyeksi Weighted'], top:0, ...getChartTheme().legend },
        xAxis:{ ...getChartTheme().xAxis, type:'category', data:days },
        yAxis:{ ...getChartTheme().yAxis, type:'value' },
        tooltip:{ ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'cross'} },
        series:[
          { name:'PS Kumulatif', type:'line', data:cum, smooth:false, lineStyle:{width:2,color:'#22d972'}, symbol:'none', color:'#22d972', areaStyle:{color:'rgba(34,217,114,.08)'} },
          { name:'Proyeksi EOM', type:'line', data:proj, smooth:false, lineStyle:{width:2,color:'#58a6ff',type:'dashed'}, symbol:'none', color:'#58a6ff' },
          { name:'Proyeksi Weighted', type:'line', data:projWeighted, smooth:false, lineStyle:{width:2,color:'#f59e0b',type:'dotted'}, symbol:'none', color:'#f59e0b' },
        ],
      }));
    },

    // ── SF Perf Tier Distribution ──
    renderSfpTier() {
      const chart = this.getOrInit('chart-sfp-tier');
      if (!chart) return;
      const dist = this.sfperfTierDist;
      const tiers = TIER_CONFIG.map(t=>t.name);
      const colors = { Black:'#6b7280', Bronze:'#f0a060', Silver:'#cdd6e8', Gold:'#fcd34d', Platinum:'#00d4ff', Diamond:'#c084fc' };
      chart.setOption(Object.assign({}, getChartTheme(), {
        tooltip:{ ...getChartTheme().tooltip, trigger:'axis', axisPointer:{type:'shadow'} },
        legend:{ data:['Bulan Lalu','Bulan Ini'], top:0, textStyle:{color:'#8b949e',fontSize:11} },
        xAxis:{ ...getChartTheme().xAxis, type:'category', data:tiers },
        yAxis:{ ...getChartTheme().yAxis, type:'value' },
        series:[
          {
            name:'Bulan Lalu', type:'bar', barMaxWidth:30,
            data: tiers.map(t=>({ value: dist.lm[t]||0, itemStyle:{ color:colors[t]||'#8b949e', opacity:0.4, borderRadius:[4,4,0,0] } })),
            label:{ show:true, position:'top', color:'#8b949e', fontSize:10, fontFamily:"'Poppins',sans-serif" },
          },
          {
            name:'Bulan Ini', type:'bar', barMaxWidth:30,
            data: tiers.map(t=>({ value: dist.now[t]||0, itemStyle:{ color:colors[t]||'#00c2ff', borderRadius:[4,4,0,0] } })),
            label:{ show:true, position:'top', color:'#8b949e', fontSize:10, fontFamily:"'Poppins',sans-serif" },
          },
        ],
        grid:{ left:30, right:10, top:40, bottom:10, containLabel:true },
      }));
    },

    // ── TOAST ──
    toast(msg, type='ok', ms) { showToast(msg, type, ms); },   // komponen bersama (config.js)
  };
}

