import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// ── DEFAULT DATA SETS DARI DOKUMEN SOP KEAMANAN & ASET PERUSAHAAN ────────────────
export const DEFAULT_IK_SECURITY = [
  {
    id: 1,
    title: 'Pengamanan Area & Aset Perusahaan',
    category: 'Pengamanan Fisik & Patroli',
    icon: '🛡️',
    color: '#2563EB',
    bg: '#EFF6FF',
    target: 'Seluruh area dalam & luar Head Quarter dan Klinik Cabang Audy Dental',
    tools: [
      'Google Form Checklist Monitoring Klinik: https://forms.gle/1EAvkqDJVdctXEGn7',
      'Senter Patroli',
      'Komputer dengan Akses Monitor CCTV (Live View)',
      'Alat Komunikasi / Handphone / HT'
    ],
    steps: [
      'Lakukan Safety Patrol secara rutin mengelilingi seluruh area sesuai jadwal yang telah ditentukan (Patrol 1, 2, dan 3).',
      'Catat hasil temuan dan kondisi keamanan secara berkala ke dalam Google Form Safety Patrol resmi.',
      'Awasi kendaraan yang terparkir di area luar klinik; pastikan khusus sepeda motor dalam keadaan terkunci ganda.',
      'Pantau situasi lingkungan dan area sensitif secara berkala melalui layar monitor CCTV saat berada di pos security.',
      'Segera lakukan tindakan persuasif ramah apabila ditemukan potensi gangguan keamanan, dan segera hubungi pihak Kepolisian terdekat bila situasi darurat tidak dapat diatasi.',
      'Bila ditemukan permasalahan fisik, kerusakan fasilitas, atau kendala di area klinik, segera dokumentasikan (foto) dan laporkan langsung ke SPV Klinik dan Leader Facility Care.',
      'Disaat malam hari atau saat klinik telah tutup operasional, pastikan seluruh AC dan lampu ruangan dalam keadaan mati, kecuali lampu keamanan tertentu yang memang harus dibiarkan menyala terang.'
    ],
    notes: 'Prioritaskan keselamatan diri dan tindakan persuasif. Dilarang meninggalkan area pengawasan tanpa personil pengganti yang jelas.'
  },
  {
    id: 2,
    title: 'Penertiban & Pengaturan Parkir Kendaraan',
    category: 'Operasional Parkir & Luar',
    icon: '🅿️',
    color: '#D97706',
    bg: '#FFFBEB',
    target: 'Area parkir mobil, motor, akses keluar masuk, dan teras depan klinik',
    tools: [
      'Lampu Lalu Lintas (Lampu Lalin)',
      'Cone / Kerucut Pembatas Parkir',
      'Peluit Parkir'
    ],
    steps: [
      'Berdiri di posisi strategis dan bantu mengarahkan setiap kendaraan pasien atau tamu yang datang secara tepat dan rapi ke slot parkir yang tersedia.',
      'Berikan informasi dan solusi dengan sikap ramah serta santun apabila seluruh slot parkir klinik dalam kondisi penuh.',
      'Dalam kondisi tertentu (bila ada mobil yang terhalang mobil lain dan harus dipindahkan, sementara pasien sedang dalam tindakan perawatan): bantu memindahkan kendaraan HANYA JIKA petugas security memiliki kemampuan mengemudi dan memiliki SIM aktif.',
      'Apabila petugas security tidak memiliki kemampuan mengemudikan mobil: sampaikan permohonan maaf dengan tulus kepada pasien/tamu, minta untuk menunggu sejenak, dan segera minta bantuan kepada staf Resepsionis untuk menghubungi pemilik mobil yang menghalangi.',
      'Bantu pasien atau tamu yang akan keluar parkir dengan memberikan aba-aba dan arahan lalu lintas jalan raya yang aman dan tepat.'
    ],
    notes: 'Dilarang memaksakan memindahkan mobil pasien jika tidak memiliki lisensi mengemudi demi menghindari risiko gesekan atau kerusakan kendaraan.'
  },
  {
    id: 3,
    title: 'Penerimaan Tamu & Pasien Klinik',
    category: 'Layanan & Hospitalitas',
    icon: '🤝',
    color: '#059669',
    bg: '#ECFDF5',
    target: 'Pintu masuk utama klinik, lobby, area resepsionis, dan pos penerimaan tamu',
    tools: [
      'Alat Tulis Kantor (Pulpen)',
      'Buku Tamu / Log Book Keamanan',
      'Alat Komunikasi / Handphone'
    ],
    steps: [
      'Berdiri dengan tegap dan sigap, bukakan pintu untuk setiap pasien/tamu yang datang dengan menerapkan budaya 5S (Senyum, Salam, Sapa, Sopan, Santun) dan menyapa: "Selamat pagi / siang / sore, ada yang bisa kami bantu?".',
      'Setelah pasien atau tamu memberitahukan tujuannya, persilakan dan arahkan pasien langsung menuju meja Resepsionis.',
      'Apabila yang datang adalah tamu instansi, vendor, atau pihak Manajemen/BOD: persilakan mengisi Buku Tamu (Log Book) terlebih dahulu dan segera hubungi staf/karyawan yang ingin ditemui.',
      'Apabila pasien langsung melangkah ke arah tujuannya, berikan gestur tubuh hormat menyilangkan kedua tangan di depan, dan jika pasien/tamu mengucapkan terima kasih, jawab dengan ramah: "Dengan Senang Hati".',
      'Dalam kondisi khusus (pasien lansia, ibu hamil, atau pengguna kursi roda): petugas security wajib mendampingi, mengawal, dan membantu mendorong kursi roda sampai ke depan meja resepsionis.',
      'Buka kembali pintu untuk setiap pasien yang telah selesai melakukan perawatan saat hendak keluar klinik sambil mengucapkan ucapan hangat: "Terima Kasih, semoga lekas sehat kembali".'
    ],
    notes: 'Jaga senyum tulus dan keramahan sebagai garda terdepan representasi citra Audy Dental.'
  }
];

export const DEFAULT_SECURITY_ASSETS = [
  { no: 1, type: 'Aset Tetap', item: 'Rumah & Bangunan', description: 'Gedung klinik, pos satpam, pagar, kanopi, teras beton', pic: 'Building Maintenance & Security' },
  { no: 2, type: 'Aset Tetap', item: 'Genset & Lift', description: 'Mesin generator listrik cadangan dan elevator penumpang tetap', pic: 'Teknisi & Security' },
  { no: 3, type: 'Aset Tetap', item: 'AC Central & Split Tetap', description: 'Outdoor dan indoor AC yang terpasang permanen di dinding/plafon', pic: 'Teknisi AC & Security' },
  { no: 4, type: 'Aset Tetap', item: 'Dental Unit Permanen', description: 'Kursi dental unit terpasang instalasi air dan kompresor', pic: 'Tim Klinik & Security' },
  { no: 5, type: 'Aset Tetap', item: 'Instalasi Listrik, Pipa & Jaringan IT Permanen', description: 'Panel MCB, kabel fiber optik, pipa drainase, server rack', pic: 'IT & Teknisi' },
  { no: 6, type: 'Aset Tetap', item: 'Furniture Melekat (Built-in)', description: 'Meja resepsionis cor/melekat, partisi kaca mati, kitchen set pantry', pic: 'General Support' },
  { no: 7, type: 'Aset Bergerak', item: 'Komputer, Laptop & Printer', description: 'PC kasir/resepsionis, laptop staf, printer rontgen/dokumen', pic: 'Staf Terkait & IT' },
  { no: 8, type: 'Aset Bergerak', item: 'Kendaraan Operasional', description: 'Mobil operasional dan sepeda motor inventaris perusahaan', pic: 'Security & Driver' },
  { no: 9, type: 'Aset Bergerak', item: 'Furniture Lepasan', description: 'Meja kerja, kursi resepsionis, kursi tunggu pasien, lemari file', pic: 'Facility Care & Staf' },
  { no: 10, type: 'Aset Bergerak', item: 'Hand Tools, Alat Kerja & Alat Dental Portabel', description: 'Handpiece dental, alat rontgen portable, toolbox teknisi', pic: 'Dokter/Perawat & Teknisi' },
  { no: 11, type: 'Low Value Asset', item: 'Peralatan Kantor (Telepon, ATK)', description: 'Pesawat telepon, kalkulator, stapler, perlengkapan meja', pic: 'Resepsionis' },
  { no: 12, type: 'Low Value Asset', item: 'Alat-alat Kebersihan', description: 'Lobby duster, ember wringer pel, sapu, tangga, botol spray', pic: 'Facility Care' },
  { no: 13, type: 'Low Value Asset', item: 'Perlengkapan Security', description: 'Lampu lalin, cone parkir, peluit, senter patroli, jas hujan', pic: 'Security' }
];

export const ZONASI_SECURITY = {
  cabang: [
    { zone: 'Zona Hijau', level: 'Akses Publik Bebas', color: '#059669', bg: '#ECFDF5', rooms: ['Area Parkir Kendaraan', 'Lobby Lantai 1', 'Teras Luar Klinik'], desc: 'Area dengan pengamanan umum yang bebas diakses pasien dan tamu setelah pengarahan security.' },
    { zone: 'Zona Kuning', level: 'Akses Terbatas / Pendampingan', color: '#D97706', bg: '#FFFBEB', rooms: ['Koridor Ruang Tindakan', 'Tangga & Lift Klinik', 'Ruang Tunggu Lantai 2 & 3', 'Kids Playground', 'Toilet Lantai 1 & 2', 'Mushola Staf & Pasien'], desc: 'Area terbatas, tamu wajib didampingi petugas atau staf klinik terkait.' },
    { zone: 'Zona Merah', level: 'Area Terlarang / Restricted', color: '#DC2626', bg: '#FEF2F2', rooms: ['Area Meja Receptionist / Kasir', 'Seluruh Ruang Dental Unit (DU 101, 102, 201, 202, 301, 302)', 'Ruang Sterilisasi Alat', 'Ruang Rontgen Radiologi', 'Ruang Suction & Kompresor', 'Ruang Tim Staf & Gudang Material Medis'], desc: 'Area sangat sensitif, hanya staf medis bertugas dan atasan pemilik area. Tamu wajib izin Manajemen/BOD dan didampingi penuh.' }
  ],
  hq: [
    { floor: 'Lantai 1', hijau: ['Area Parkir Mobil/Motor', 'Lobby Resepsionis Lantai 1'], kuning: ['Seluruh Ruang Meeting Lantai 1', 'Ruang Training Lantai 1', 'Koridor Lift & Toilet'], merah: ['Ruang Customer Services', 'Gudang Utama Lantai 1', 'Gudang Facility Care', 'Gudang Teknisi & Ruang Teknisi', 'Pos Satpam & Ruang Genset'] },
    { floor: 'Lantai 2', hijau: ['-'], kuning: ['Ruang Meeting Lantai 2', 'Koridor Lift & Toilet', 'Pantry'], merah: ['Ruang CEO', 'Ruang Sekretaris CEO', 'Ruang Human Capital', 'Ruang Dentist Relation', 'Ruang General Support', 'Ruang Eks. Patrel'] },
    { floor: 'Lantai 3', hijau: ['-'], kuning: ['Koridor Lift & Toilet', 'Pantry', 'Mushola & Tempat Wudhu', 'Ruang Santai Karyawan'], merah: ['Ruang Marketing', 'Studio Konten Marketing', 'Gudang Marketing', 'Ruang Patient Relation', 'Ruang OPS', 'Ruang Training Lantai 3', 'Gudang Lantai 3'] },
    { floor: 'Lantai 4', hijau: ['-'], kuning: ['Koridor Lift & Toilet', 'Pantry'], merah: ['Ruang Finance & Accounting', 'Ruang VP Finance & Accounting', 'Ruang Audit Internal', 'Ruang Training Lantai 4', 'Gudang Lantai 4'] },
    { floor: 'Lantai 5', hijau: ['-'], kuning: ['Lobby Lantai 5', 'Pantry Lantai 5', 'Toilet Lantai 5', 'Ruang Training Besar Lantai 5'], merah: ['Area Dak Toren Air & Rooftop'] }
  ]
};

// ── STORAGE REPOSITORY ──────────────────────────────────────────────────────────
const KEY_IK_SEC = 'fcms_sop_sec_ik_v1';
const KEY_ASSETS_SEC = 'fcms_sop_sec_assets_v1';

function getStoredIK() {
  try {
    const raw = localStorage.getItem(KEY_IK_SEC);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.error('Error load IK Sec', e); }
  return JSON.parse(JSON.stringify(DEFAULT_IK_SECURITY));
}

function saveStoredIK(list) {
  localStorage.setItem(KEY_IK_SEC, JSON.stringify(list));
}

function getStoredAssets() {
  try {
    const raw = localStorage.getItem(KEY_ASSETS_SEC);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.error('Error load assets sec', e); }
  return JSON.parse(JSON.stringify(DEFAULT_SECURITY_ASSETS));
}

function saveStoredAssets(list) {
  localStorage.setItem(KEY_ASSETS_SEC, JSON.stringify(list));
}

// ── MAIN RENDER FUNCTION ────────────────────────────────────────────────────────
export async function renderSOPSecurity(container) {
  let activeTab = 'ik'; // 'ik' | 'zonasi' | 'assets' | 'policy' | 'links'
  let currentSearch = '';
  let currentCategory = 'all';
  let assetSearch = '';
  let zonasiView = 'cabang'; // 'cabang' | 'hq'

  let listIK = getStoredIK();
  let listAssets = getStoredAssets();

  function renderPage() {
    container.innerHTML = `
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
            <span>🛡️</span> SOP Pengelolaan Keamanan & Aset Perusahaan
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
            Standar Operasional Prosedur Pengawasan Keamanan Area, Fasilitas, dan Perlindungan Aset PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            🔄 Reset ke Standar
          </button>
          <button id="btn-print-sec" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            🖨️ Cetak Ringkasan SOP
          </button>
          <button id="btn-export-sec" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            📥 Download Data Excel
          </button>
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto;">
        <button class="sec-tab-btn ${activeTab === 'ik' ? 'active' : ''}" data-tab="ik" style="${getTabStyle(activeTab === 'ik')}">
          🛡️ Instruksi Kerja (IK) <span class="badge" style="${getBadgeStyle(activeTab === 'ik')}">${listIK.length}</span>
        </button>
        <button class="sec-tab-btn ${activeTab === 'zonasi' ? 'active' : ''}" data-tab="zonasi" style="${getTabStyle(activeTab === 'zonasi')}">
          🗺️ Zonasi Akses & Denah
        </button>
        <button class="sec-tab-btn ${activeTab === 'assets' ? 'active' : ''}" data-tab="assets" style="${getTabStyle(activeTab === 'assets')}">
          📦 Klasifikasi Aset <span class="badge" style="${getBadgeStyle(activeTab === 'assets')}">${listAssets.length}</span>
        </button>
        <button class="sec-tab-btn ${activeTab === 'policy' ? 'active' : ''}" data-tab="policy" style="${getTabStyle(activeTab === 'policy')}">
          📋 Shift, Patrol & Insiden
        </button>
        <button class="sec-tab-btn ${activeTab === 'links' ? 'active' : ''}" data-tab="links" style="${getTabStyle(activeTab === 'links')}">
          🔗 Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sec-content-area"></div>
    `;

    bindTabEvents();
    renderActiveTabContent();
  }

  function getTabStyle(isActive) {
    if (isActive) {
      return 'border:none; background:#1E40AF; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;';
    }
    return 'border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;';
  }

  function getBadgeStyle(isActive) {
    if (isActive) {
      return 'background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;';
    }
    return 'background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;';
  }

  function bindTabEvents() {
    container.querySelectorAll('.sec-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        activeTab = btn.dataset.tab;
        renderPage();
      });
    });

    const btnReset = container.querySelector('#btn-reset-sec-default');
    if (btnReset) {
      btnReset.addEventListener('click', () => {
        confirmDialog('Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja dan Klasifikasi Aset Security ke standar dokumen asli?', () => {
          localStorage.removeItem(KEY_IK_SEC);
          localStorage.removeItem(KEY_ASSETS_SEC);
          listIK = getStoredIK();
          listAssets = getStoredAssets();
          toastSuccess('Data SOP Security berhasil di-reset ke standar dokumen');
          renderPage();
        }, 'Reset ke Standar Dokumen');
      });
    }

    const btnPrint = container.querySelector('#btn-print-sec');
    if (btnPrint) {
      btnPrint.addEventListener('click', () => {
        window.print();
      });
    }

    const btnExport = container.querySelector('#btn-export-sec');
    if (btnExport) {
      btnExport.addEventListener('click', () => {
        exportAllData();
      });
    }
  }

  function renderActiveTabContent() {
    const content = container.querySelector('#sec-content-area');
    if (!content) return;

    if (activeTab === 'ik') renderIKTab(content);
    else if (activeTab === 'zonasi') renderZonasiTab(content);
    else if (activeTab === 'assets') renderAssetsTab(content);
    else if (activeTab === 'policy') renderPolicyTab(content);
    else if (activeTab === 'links') renderLinksTab(content);
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 1: INSTRUKSI KERJA SECURITY (CRUD)
  // ─────────────────────────────────────────────────────────────────────────────
  function renderIKTab(parent) {
    parent.innerHTML = `
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${currentSearch}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">🔍</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${currentCategory === 'all' ? 'active' : ''}" data-cat="all">Semua (${listIK.length})</button>
              <button class="sec-cat-pill ${currentCategory === 'Pengamanan Fisik & Patroli' ? 'active' : ''}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${listIK.filter(x => x.category === 'Pengamanan Fisik & Patroli').length})</button>
              <button class="sec-cat-pill ${currentCategory === 'Operasional Parkir & Luar' ? 'active' : ''}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${listIK.filter(x => x.category === 'Operasional Parkir & Luar').length})</button>
              <button class="sec-cat-pill ${currentCategory === 'Layanan & Hospitalitas' ? 'active' : ''}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${listIK.filter(x => x.category === 'Layanan & Hospitalitas').length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${listIK.length} Instruksi Kerja
            </div>
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
              ➕ Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="sec-ik-cards-grid-container"></div>
    `;

    // Inject Cat pill styles
    const styleTag = document.createElement('style');
    styleTag.innerHTML = `
      .sec-cat-pill {
        border: 1px solid var(--border);
        background: #F8FAFC;
        color: var(--text-2);
        font-size: 0.8rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
      }
      .sec-cat-pill:hover { background: #E2E8F0; }
      .sec-cat-pill.active {
        background: #1E40AF;
        color: #fff;
        border-color: #1E40AF;
      }
      .sec-ik-card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.08) !important;
        transform: translateY(-2px);
      }
    `;
    parent.appendChild(styleTag);

    const gridContainer = parent.querySelector('#sec-ik-cards-grid-container');
    const countLabel = parent.querySelector('#sec-ik-count-label');

    function updateIKGrid() {
      const filteredIK = listIK.filter(item => {
        const matchCat = currentCategory === 'all' || item.category === currentCategory;
        const q = currentSearch.toLowerCase().trim();
        const matchQuery = !q ||
          item.title.toLowerCase().includes(q) ||
          item.target.toLowerCase().includes(q) ||
          item.tools.some(t => t.toLowerCase().includes(q)) ||
          item.steps.some(s => s.toLowerCase().includes(q));
        return matchCat && matchQuery;
      });

      if (countLabel) {
        countLabel.textContent = `Menampilkan ${filteredIK.length} dari ${listIK.length} Instruksi Kerja`;
      }

      if (filteredIK.length === 0) {
        gridContainer.innerHTML = `
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">🔍</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;
        return;
      }

      gridContainer.innerHTML = `
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${filteredIK.map(ik => `
            <div class="card sec-ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${ik.bg || '#EFF6FF'}; border-radius:10px;">${ik.icon || '🛡️'}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        IK-${String(ik.id).padStart(2, '0')}. ${ik.title}
                      </h3>
                      <span class="badge" style="background:${ik.bg || '#EFF6FF'}; color:${ik.color || '#1E40AF'}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${ik.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-sec-ik" data-id="${ik.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      ✏️
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-sec-ik" data-id="${ik.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      🗑️
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${ik.color || '#1E40AF'};">
                  <strong>Area Sasaran:</strong> ${ik.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">🧰 Alat & Perlengkapan:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${ik.tools.slice(0, 3).map(t => `<li style="margin-bottom:2px;">${escapeHtml(t)}</li>`).join('')}
                    ${ik.tools.length > 3 ? `<li style="color:var(--text-3); font-style:italic;">+ ${ik.tools.length - 3} perlengkapan lainnya...</li>` : ''}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${ik.steps.length} Langkah Prosedur
                </span>
                <button class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${ik.id}" style="padding:6px 14px; font-weight:600; border-radius:6px; background:#1E40AF; border-color:#1E40AF;">
                  Lihat Panduan Lengkap ➔
                </button>
              </div>
            </div>
          `).join('')}
        </div>
      `;

      // Event detail modal
      gridContainer.querySelectorAll('.btn-view-sec-ik').forEach(btn => {
        btn.addEventListener('click', () => {
          const ikId = parseInt(btn.dataset.id, 10);
          const ik = listIK.find(x => x.id === ikId);
          if (ik) openIKDetailModal(ik);
        });
      });

      // Event edit modal
      gridContainer.querySelectorAll('.btn-edit-sec-ik').forEach(btn => {
        btn.addEventListener('click', () => {
          const ikId = parseInt(btn.dataset.id, 10);
          const ik = listIK.find(x => x.id === ikId);
          if (ik) openIKFormModal(ik);
        });
      });

      // Event delete
      gridContainer.querySelectorAll('.btn-delete-sec-ik').forEach(btn => {
        btn.addEventListener('click', () => {
          const ikId = parseInt(btn.dataset.id, 10);
          const ik = listIK.find(x => x.id === ikId);
          if (!ik) return;
          confirmDialog(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${ik.title}"?`, () => {
            listIK = listIK.filter(x => x.id !== ikId);
            saveStoredIK(listIK);
            toastSuccess('Instruksi Kerja Security berhasil dihapus');
            renderPage();
          }, 'Hapus Instruksi Kerja');
        });
      });
    }

    updateIKGrid();

    // Event search
    const searchInput = parent.querySelector('#sec-ik-search');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        currentSearch = e.target.value;
        updateIKGrid();
      });
    }

    // Event cat pills
    parent.querySelectorAll('.sec-cat-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        parent.querySelectorAll('.sec-cat-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        currentCategory = pill.dataset.cat;
        updateIKGrid();
      });
    });

    // Event Add New IK
    const btnAddIK = parent.querySelector('#btn-add-sec-ik');
    if (btnAddIK) {
      btnAddIK.addEventListener('click', () => {
        openIKFormModal();
      });
    }
  }

  // MODAL DETAIL IK
  function openIKDetailModal(ik) {
    createModal({
      title: `${ik.icon || '🛡️'} IK-${String(ik.id).padStart(2, '0')}: ${ik.title}`,
      size: 'lg',
      content: `
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${ik.bg || '#EFF6FF'}; color:${ik.color || '#1E40AF'}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${ik.category}</span>
            <span>Kode: IK-SEC-${String(ik.id).padStart(2, '0')}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${ik.color || '#1E40AF'};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">📍 Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${ik.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">🧰 Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${ik.tools.map(t => {
                  if (t.includes('http')) {
                    const parts = t.split(': ');
                    return `<li>${parts[0]}: <a href="${parts.slice(1).join(': ')}" target="_blank" rel="noopener" style="color:#2563EB; font-weight:600; text-decoration:underline;">Buka Tautan Form ➔</a></li>`;
                  }
                  return `<li>${escapeHtml(t)}</li>`;
                }).join('')}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">📝 Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${ik.steps.map(s => `<li style="margin-bottom:6px;">${escapeHtml(s)}</li>`).join('')}
              </ol>
            </div>
          </div>

          ${ik.notes ? `
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">⚠️ Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${escapeHtml(ik.notes)}</span>
            </div>
          ` : ''}
        </div>
      `,
      confirmText: 'Tutup Panduan',
      onConfirm: (overlay, close) => {
        close();
      }
    });
  }

  // MODAL FORM TAMBAH & EDIT IK
  function openIKFormModal(existingIK = null) {
    const isEdit = !!existingIK;
    const initialTools = existingIK ? existingIK.tools.join('\n') : '';
    const initialSteps = existingIK ? existingIK.steps.join('\n') : '';

    const contentHtml = `
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${existingIK ? escapeHtml(existingIK.title) : ''}" placeholder="Contoh: Pengamanan Area & Aset Perusahaan" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${existingIK ? existingIK.icon : '🛡️'}" placeholder="🛡️" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
          <select id="form-sec-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Pengamanan Fisik & Patroli" ${existingIK?.category === 'Pengamanan Fisik & Patroli' ? 'selected' : ''}>Pengamanan Fisik & Patroli</option>
            <option value="Operasional Parkir & Luar" ${existingIK?.category === 'Operasional Parkir & Luar' ? 'selected' : ''}>Operasional Parkir & Luar</option>
            <option value="Layanan & Hospitalitas" ${existingIK?.category === 'Layanan & Hospitalitas' ? 'selected' : ''}>Layanan & Hospitalitas</option>
          </select>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${existingIK ? escapeHtml(existingIK.target) : ''}" placeholder="Contoh: Seluruh area luar klinik dan akses parkir" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" placeholder="Senter Patroli&#10;Lampu Lalin&#10;Peluit" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${escapeHtml(initialTools)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" placeholder="1. Lakukan Safety Patrol sesuai jadwal.&#10;2. Arahkan kendaraan pasien dengan ramah..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${escapeHtml(initialSteps)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" placeholder="Peringatan keselamatan atau pedoman khusus..." style="width:100%; border-radius:6px; font-size:0.85rem;">${existingIK ? escapeHtml(existingIK.notes || '') : ''}</textarea>
        </div>
      </form>
    `;

    createModal({
      title: isEdit ? `✏️ Edit IK-${existingIK.id}: ${existingIK.title}` : `➕ Tambah Instruksi Kerja Security Baru`,
      content: contentHtml,
      size: 'lg',
      confirmText: 'Simpan Perubahan',
      onConfirm: (overlay, close) => {
        const title = overlay.querySelector('#form-sec-ik-title').value.trim();
        const icon = overlay.querySelector('#form-sec-ik-icon').value.trim() || '🛡️';
        const category = overlay.querySelector('#form-sec-ik-cat').value;
        const target = overlay.querySelector('#form-sec-ik-target').value.trim();
        const toolsRaw = overlay.querySelector('#form-sec-ik-tools').value.trim();
        const stepsRaw = overlay.querySelector('#form-sec-ik-steps').value.trim();
        const notes = overlay.querySelector('#form-sec-ik-notes').value.trim();

        if (!title || !target || !toolsRaw || !stepsRaw) {
          toastError('Harap lengkapi semua kolom bertanda bintang (*)');
          return;
        }

        const tools = toolsRaw.split('\n').map(s => s.trim().replace(/^[-•*]\s*/, '')).filter(Boolean);
        const steps = stepsRaw.split('\n').map(s => s.trim().replace(/^\d+[\.\)]\s*/, '')).filter(Boolean);

        const categoryColors = {
          'Pengamanan Fisik & Patroli': { color: '#2563EB', bg: '#EFF6FF' },
          'Operasional Parkir & Luar': { color: '#D97706', bg: '#FFFBEB' },
          'Layanan & Hospitalitas': { color: '#059669', bg: '#ECFDF5' }
        };
        const styling = categoryColors[category] || { color: '#1E40AF', bg: '#EFF6FF' };

        if (isEdit) {
          existingIK.title = title;
          existingIK.icon = icon;
          existingIK.category = category;
          existingIK.target = target;
          existingIK.tools = tools;
          existingIK.steps = steps;
          existingIK.notes = notes;
          existingIK.color = styling.color;
          existingIK.bg = styling.bg;
          saveStoredIK(listIK);
          toastSuccess(`Instruksi Kerja "${title}" berhasil diperbarui`);
        } else {
          const nextId = listIK.length > 0 ? Math.max(...listIK.map(x => x.id)) + 1 : 1;
          const newIK = {
            id: nextId,
            title,
            category,
            icon,
            color: styling.color,
            bg: styling.bg,
            target,
            tools,
            steps,
            notes
          };
          listIK.push(newIK);
          saveStoredIK(listIK);
          toastSuccess(`Instruksi Kerja baru IK-${nextId} berhasil ditambahkan`);
        }

        close();
        renderPage();
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 2: ZONASI AKES & DENAH (LAMPIRAN 8.1)
  // ─────────────────────────────────────────────────────────────────────────────
  function renderZonasiTab(parent) {
    parent.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:20px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-cabang-btn" class="btn btn-sm ${zonasiView === 'cabang' ? 'btn-primary' : 'btn-ghost'}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            🏥 Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${zonasiView === 'hq' ? 'btn-primary' : 'btn-ghost'}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            🏢 Zonasi Head Quarter (5 Lantai)
          </button>
        </div>
        <div style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Ketentuan Pembagian Akses Ruangan (Lampiran 8.1 SOP)
        </div>
      </div>

      <!-- Definisi 3 Zona Header Cards -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px; margin-bottom:22px;">
        <div class="card" style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#DC2626; display:inline-block;"></span>
            <strong style="color:#991B1B; font-size:0.95rem;">ZONA MERAH (Restricted Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#7F1D1D; line-height:1.5;">
            Area sangat terbatas. Hanya personal yang bertugas spesifik di area tersebut dan BOD. Tamu WAJIB mendapat izin tertulis Manajemen dan didampingi penuh.
          </p>
        </div>

        <div class="card" style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#D97706; display:inline-block;"></span>
            <strong style="color:#92400E; font-size:0.95rem;">ZONA KUNING (Limited Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#78350F; line-height:1.5;">
            Area dengan pengamanan terbatas. Karyawan umum atau tamu yang memasuki area ini perlu mendapatkan arahan dan didampingi staf area terkait.
          </p>
        </div>

        <div class="card" style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:10px; padding:14px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:14px; height:14px; border-radius:50%; background:#059669; display:inline-block;"></span>
            <strong style="color:#065F46; font-size:0.95rem;">ZONA HIJAU (Public Area)</strong>
          </div>
          <p style="margin:0; font-size:0.82rem; color:#064E3B; line-height:1.5;">
            Area umum dengan pengawasan security. Diperbolehkan untuk seluruh tamu dan pasien setelah diarahkan oleh petugas security.
          </p>
        </div>
      </div>

      <!-- Zonasi Content Area -->
      <div id="zonasi-content-container"></div>
    `;

    const subContent = parent.querySelector('#zonasi-content-container');

    if (zonasiView === 'cabang') {
      subContent.innerHTML = `
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
          ${ZONASI_SECURITY.cabang.map(z => `
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; border-bottom:1px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.05rem; font-weight:700; color:${z.color}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${z.color}; display:inline-block;"></span>
                  ${z.zone}
                </h3>
                <span class="badge" style="background:${z.bg}; color:${z.color}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:6px;">
                  ${z.level}
                </span>
              </div>
              <p style="font-size:0.82rem; color:var(--text-3); margin-bottom:14px; line-height:1.5;">${z.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:var(--text-2); display:block; margin-bottom:6px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 16px; padding:0; font-size:0.85rem; color:var(--text-1); line-height:1.7;">
                  ${z.rooms.map(r => `<li>${r}</li>`).join('')}
                </ul>
              </div>
            </div>
          `).join('')}
        </div>
      `;
    } else {
      subContent.innerHTML = `
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${ZONASI_SECURITY.hq.map(fl => `
            <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:2px solid #F1F5F9; padding-bottom:8px;">
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
                  <span>🏢</span> Head Quarter — ${fl.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:#ECFDF5; border:1px solid #A7F3D0; border-radius:8px; padding:12px;">
                  <strong style="color:#065F46; font-size:0.85rem; display:block; margin-bottom:6px;">🟢 ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#064E3B; line-height:1.6;">
                    ${fl.hijau.map(item => `<li>${item}</li>`).join('')}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:8px; padding:12px;">
                  <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:6px;">🟡 ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#78350F; line-height:1.6;">
                    ${fl.kuning.map(item => `<li>${item}</li>`).join('')}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:8px; padding:12px;">
                  <strong style="color:#991B1B; font-size:0.85rem; display:block; margin-bottom:6px;">🔴 ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.82rem; color:#7F1D1D; line-height:1.6;">
                    ${fl.merah.map(item => `<li>${item}</li>`).join('')}
                  </ul>
                </div>
              </div>
            </div>
          `).join('')}
        </div>
      `;
    }

    parent.querySelector('#view-cabang-btn').addEventListener('click', () => {
      zonasiView = 'cabang';
      renderZonasiTab(parent);
    });

    parent.querySelector('#view-hq-btn').addEventListener('click', () => {
      zonasiView = 'hq';
      renderZonasiTab(parent);
    });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 3: KLASIFIKASI ASET (CRUD)
  // ─────────────────────────────────────────────────────────────────────────────
  function renderAssetsTab(parent) {
    parent.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div>
          <h3 style="margin:0 0 4px 0; font-size:1.1rem; font-weight:700; color:var(--text-1);">
            Daftar Klasifikasi Aset yang Dipantau Security (Lampiran 8.3)
          </h3>
          <p style="margin:0; font-size:0.85rem; color:var(--text-3);">
            Pengawasan dan perlindungan sarana prasarana fisik, elektronik, dan peralatan klinik
          </p>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${assetSearch}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem; background:#1E40AF; border-color:#1E40AF;">
            ➕ Tambah Aset Baru
          </button>
        </div>
      </div>

      <div id="sec-asset-table-container"></div>
    `;

    const subContent = parent.querySelector('#sec-asset-table-container');
    const searchInput = parent.querySelector('#sec-asset-search');

    function renderAssetTable() {
      const q = assetSearch.toLowerCase().trim();
      const filtered = listAssets.filter(x => !q || x.item.toLowerCase().includes(q) || x.type.toLowerCase().includes(q) || x.description.toLowerCase().includes(q) || (x.pic && x.pic.toLowerCase().includes(q)));

      subContent.innerHTML = `
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:var(--text-1); font-size:0.95rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge badge-primary">${filtered.length} Aset</span>
          </div>
          <div style="overflow-x:auto;">
            <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
              <thead>
                <tr style="background:#F1F5F9; border-bottom:2px solid #E2E8F0; text-align:left;">
                  <th style="padding:10px 14px; width:60px;">No</th>
                  <th style="padding:10px 14px; width:140px;">Tipe Aset</th>
                  <th style="padding:10px 14px; width:220px;">Nama Aset / Kelompok</th>
                  <th style="padding:10px 14px;">Deskripsi / Contoh Item</th>
                  <th style="padding:10px 14px; width:180px;">PIC Pengawas</th>
                  <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${filtered.map((item, idx) => `
                  <tr style="border-bottom:1px solid #F1F5F9; background:${idx % 2 === 0 ? '#fff' : '#FAFAFA'};">
                    <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${item.no || idx + 1}</td>
                    <td style="padding:10px 14px;">
                      <span class="badge" style="${getTypeBadgeStyle(item.type)}">${escapeHtml(item.type)}</span>
                    </td>
                    <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${escapeHtml(item.item)}</td>
                    <td style="padding:10px 14px; color:var(--text-2);">${escapeHtml(item.description)}</td>
                    <td style="padding:10px 14px; font-weight:600; color:#1E40AF;">${escapeHtml(item.pic || '-')}</td>
                    <td style="padding:10px 14px; text-align:center;">
                      <button class="btn btn-sm btn-ghost btn-edit-asset" data-no="${item.no}" title="Edit Aset" style="padding:4px 6px; cursor:pointer;">✏️</button>
                      <button class="btn btn-sm btn-ghost btn-delete-asset" data-no="${item.no}" title="Hapus Aset" style="padding:4px 6px; color:#DC2626; cursor:pointer;">🗑️</button>
                    </td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>
      `;

      // Event Edit Asset
      subContent.querySelectorAll('.btn-edit-asset').forEach(btn => {
        btn.addEventListener('click', () => {
          const no = parseInt(btn.dataset.no, 10);
          const asset = listAssets.find(x => x.no === no);
          if (asset) openAssetFormModal(asset);
        });
      });

      // Event Delete Asset
      subContent.querySelectorAll('.btn-delete-asset').forEach(btn => {
        btn.addEventListener('click', () => {
          const no = parseInt(btn.dataset.no, 10);
          const asset = listAssets.find(x => x.no === no);
          if (!asset) return;
          confirmDialog(`Apakah Anda yakin ingin menghapus aset: "${asset.item}"?`, () => {
            listAssets = listAssets.filter(x => x.no !== no);
            listAssets.forEach((a, i) => { a.no = i + 1; });
            saveStoredAssets(listAssets);
            toastSuccess('Aset berhasil dihapus');
            renderPage();
          }, 'Hapus Aset');
        });
      });
    }

    function getTypeBadgeStyle(type) {
      if (type === 'Aset Tetap') return 'background:#EFF6FF; color:#1E40AF; font-weight:700; padding:3px 8px; border-radius:6px;';
      if (type === 'Aset Bergerak') return 'background:#ECFDF5; color:#047857; font-weight:700; padding:3px 8px; border-radius:6px;';
      return 'background:#FEF3C7; color:#92400E; font-weight:700; padding:3px 8px; border-radius:6px;';
    }

    renderAssetTable();

    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        assetSearch = e.target.value;
        renderAssetTable();
      });
    }

    parent.querySelector('#btn-add-sec-asset').addEventListener('click', () => {
      openAssetFormModal();
    });
  }

  // MODAL FORM TAMBAH & EDIT ASET
  function openAssetFormModal(existingAsset = null) {
    const isEdit = !!existingAsset;
    const contentHtml = `
      <form id="asset-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Tipe / Klasifikasi Aset *</label>
          <select id="form-asset-type" class="form-control" style="width:100%; border-radius:6px; height:36px;">
            <option value="Aset Tetap" ${existingAsset?.type === 'Aset Tetap' ? 'selected' : ''}>Aset Tetap</option>
            <option value="Aset Bergerak" ${existingAsset?.type === 'Aset Bergerak' ? 'selected' : ''}>Aset Bergerak</option>
            <option value="Low Value Asset" ${existingAsset?.type === 'Low Value Asset' ? 'selected' : ''}>Low Value Asset</option>
          </select>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Aset / Kelompok *</label>
          <input type="text" id="form-asset-item" class="form-control" required value="${existingAsset ? escapeHtml(existingAsset.item) : ''}" placeholder="Contoh: Genset & Elevator" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required placeholder="Contoh: Generator daya darurat dan panel listrik utama" style="width:100%; border-radius:6px; font-size:0.85rem;">${existingAsset ? escapeHtml(existingAsset.description) : ''}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${existingAsset ? escapeHtml(existingAsset.pic) : ''}" placeholder="Contoh: Teknisi & Security" style="width:100%; border-radius:6px; height:36px;">
        </div>
      </form>
    `;

    createModal({
      title: isEdit ? `✏️ Edit Aset: ${existingAsset.item}` : `➕ Tambah Klasifikasi Aset Baru`,
      content: contentHtml,
      confirmText: 'Simpan',
      onConfirm: (overlay, close) => {
        const type = overlay.querySelector('#form-asset-type').value;
        const item = overlay.querySelector('#form-asset-item').value.trim();
        const description = overlay.querySelector('#form-asset-desc').value.trim();
        const pic = overlay.querySelector('#form-asset-pic').value.trim();

        if (!item || !description || !pic) {
          toastError('Harap lengkapi semua kolom');
          return;
        }

        if (isEdit) {
          existingAsset.type = type;
          existingAsset.item = item;
          existingAsset.description = description;
          existingAsset.pic = pic;
          saveStoredAssets(listAssets);
          toastSuccess(`Aset "${item}" berhasil diperbarui`);
        } else {
          const nextNo = listAssets.length > 0 ? Math.max(...listAssets.map(x => x.no)) + 1 : 1;
          listAssets.push({ no: nextNo, type, item, description, pic });
          saveStoredAssets(listAssets);
          toastSuccess(`Aset baru "${item}" berhasil ditambahkan`);
        }

        close();
        renderPage();
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 4: PEDOMAN SHIFT, PATROLI & PENANGANAN PELANGGARAN
  // ─────────────────────────────────────────────────────────────────────────────
  function renderPolicyTab(parent) {
    parent.innerHTML = `
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(340px, 1fr)); gap:20px;">
        <!-- Card 1: Shift Kerja -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#EFF6FF; padding:8px; border-radius:8px;">⏰</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Sistem Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Ketentuan 12 Jam Kerja dengan 2 Hari Libur per Minggu</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:10px; border-left:3px solid #2563EB;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">☀️ Shift 1 (Pagi - Siang):</strong>
              <span>Pukul <strong>07.00 s/d 19.00 WIB</strong></span>
            </div>
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:12px; border-left:3px solid #1E293B;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">🌙 Shift 2 (Malam):</strong>
              <span>Pukul <strong>19.00 s/d 07.00 WIB</strong></span>
            </div>
            <p style="margin:0; font-size:0.82rem; color:var(--text-3);">
              <strong>Ketentuan Off:</strong> Berlaku 2 hari off dalam setiap pekannya sesuai jadwal roster yang ditetapkan oleh Facility Care Supervisor / General Support.
            </p>
          </div>
        </div>

        <!-- Card 2: Jadwal Safety Patrol -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF3C7; padding:8px; border-radius:8px;">🚶‍♂️</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Jadwal Safety Patrol Harian (6.2 B)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Keliling Pengecekan Area & Form Checklist Digital</span>
            </div>
          </div>
          <div style="display:flex; flex-direction:column; gap:8px; font-size:0.86rem;">
            <div style="background:#F8FAFC; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 1 (Pagi):</strong>
              <span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:700;">07.30 - 08.30 WIB</span>
            </div>
            <div style="background:#F8FAFC; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 2 (Malam):</strong>
              <span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:700;">19.30 - 20.30 WIB</span>
            </div>
            <div style="background:#F8FAFC; padding:8px 12px; border-radius:6px; display:flex; justify-content:space-between; align-items:center;">
              <strong>Patroli 3 (Dini Hari):</strong>
              <span class="badge" style="background:#F3E8FF; color:#7E22CE; font-weight:700;">04.00 - 05.00 WIB</span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); margin-top:4px;">
              * <em>Patroli 3 khusus berlaku untuk area Head Quarter atau lokasi yang memiliki lebih dari 1 personil security.</em>
            </div>
          </div>
        </div>

        <!-- Card 3: Penanganan Pelanggaran Disiplin -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#ECFDF5; padding:8px; border-radius:8px;">⚠️</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Penanganan Pelanggaran Disiplin (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Pelanggaran SK Tata Tertib Perusahaan</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.85rem; color:var(--text-2); line-height:1.7;">
            <li>Petugas security membuat <strong>Berita Acara Umum (BA)</strong> kejadian secara lengkap.</li>
            <li>Melaporkan kronologis tertulis kepada atasan (General Support Manager / FC Supervisor).</li>
            <li>Melaporkan dan meneruskan berkas kepada atasan langsung dari karyawan yang bersangkutan untuk penanganan HR/HC.</li>
          </ol>
        </div>

        <!-- Card 4: Penanganan Pelanggaran Kriminal -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF2F2; padding:8px; border-radius:8px;">🚨</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:#DC2626;">Penanganan Pelanggaran Kriminal (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Tindak Pidana, Pencurian, & Pelanggaran Hukum</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.84rem; color:var(--text-2); line-height:1.7;">
            <li>Buat <strong>Berita Acara Umum</strong> secara resmi.</li>
            <li>Periksa dan amankan rekaman bukti visual melalui <strong>CCTV</strong>.</li>
            <li>Laporkan ke atasan dari karyawan terkait dengan rincian kronologis (bila pelaku karyawan).</li>
            <li>Minta dan catat detail identitas <strong>KTP / ID Card</strong> dari pihak yang bersangkutan.</li>
            <li>Amankan dan netralisir <strong>Tempat Kejadian Perkara (TKP)</strong> dari gangguan orang lain.</li>
            <li>Laporkan ke <strong>Kepolisian terdekat</strong> untuk proses hukum tindak lanjut.</li>
          </ol>
        </div>
      </div>
    `;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 5: FORM & TAUTAN
  // ─────────────────────────────────────────────────────────────────────────────
  function renderLinksTab(parent) {
    parent.innerHTML = `
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Google Form Safety Patrol -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">📑</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Safety Patrol Audy</h3>
                <span class="badge" style="background:#E0F2FE; color:#0369A1; font-size:0.75rem; font-weight:700;">Google Forms Resmi</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi keamanan yang wajib diisi oleh tim security saat melakukan Safety Patrol rutin di klinik dan kantor pusat Audy Dental.
            </p>
          </div>
          <a href="https://forms.gle/1EAvkqDJVdctXEGn7" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#1E40AF; border-color:#1E40AF;">
            📋 Buka Form Safety Patrol ➔
          </a>
        </div>

        <!-- Card: Dasar Hukum & Ruang Lingkup Perusahaan -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">⚖️</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Dasar Hukum & Ruang Lingkup</h3>
                <span class="badge" style="background:#F1F5F9; color:var(--text-2); font-size:0.75rem; font-weight:700;">Legalitas</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:12px;">
              SOP ini diterapkan sebagai tata cara Pengelolaan Keamanan Aset Perusahaan di seluruh lokasi Audy Dental Klinik di bawah naungan:
            </p>
            <ul style="margin:0 0 0 16px; padding:0; font-size:0.84rem; color:var(--text-2); line-height:1.6;">
              <li><strong>PT. Audy Mandiri Indonesia</strong></li>
              <li><strong>PT. Kastara Ananta</strong></li>
              <li><strong>PT. Nara Adidaya</strong></li>
            </ul>
          </div>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#fff; border:1px solid var(--border); border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
          <span>📜</span> Validasi & Pengesahan Dokumen SOP Security
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3px solid #2563EB; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3px solid #059669; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.8rem; color:var(--text-2);">General Support Manager</div>
          </div>
          <div style="border-left:3px solid #7C3AED; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `;
  }

  // EXPORT EXCEL DATA
  async function exportAllData() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const ikSheet = listIK.map(ik => ({
        'Kode IK': `IK-SEC-${String(ik.id).padStart(2, '0')}`,
        'Judul Prosedur': ik.title,
        'Kategori': ik.category,
        'Sasaran Target': ik.target,
        'Alat & Perlengkapan': ik.tools.join(', '),
        'Langkah Pelaksanaan': ik.steps.map((s, i) => `${i + 1}. ${s}`).join(' | '),
        'Catatan Khusus': ik.notes || '-'
      }));

      const dateStr = new Date().toISOString().slice(0, 10);
      downloadExcel(ikSheet, `SOP_Security_Instruksi_Kerja_${dateStr}`);
      toastSuccess('Data Excel SOP Security berhasil diunduh');
    } catch (e) {
      console.error(e);
      toastError('Gagal mendownload data Excel: ' + e.message);
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  renderPage();
}
