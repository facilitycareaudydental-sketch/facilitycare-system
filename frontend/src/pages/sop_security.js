import { getUser, hasPermission, apiFetch } from '../config.js';
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
  apiFetch('/api/sop/bundle/' + KEY_IK_SEC, {
    method: 'PUT',
    body: JSON.stringify(list)
  }).catch(() => {});
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
  apiFetch('/api/sop/bundle/' + KEY_ASSETS_SEC, {
    method: 'PUT',
    body: JSON.stringify(list)
  }).catch(() => {});
}

// ── MAIN RENDER FUNCTION ────────────────────────────────────────────────────────
export async function renderSOPSecurity(container) {
  const currentUser = getUser();
  const canEdit = hasPermission(currentUser, 'sop', 'write');
  const canDelete = hasPermission(currentUser, 'sop', 'delete');
  let activeTab = 'ik'; // 'ik' | 'zonasi' | 'assets' | 'policy' | 'links'
  let currentSearch = '';
  let currentCategory = 'all';
  let assetSearch = '';
  let zonasiView = 'cabang'; // 'cabang' | 'hq'

  try {
    const [remoteIK, remoteAssets] = await Promise.all([
      apiFetch('/api/sop/bundle/' + KEY_IK_SEC),
      apiFetch('/api/sop/bundle/' + KEY_ASSETS_SEC),
    ]);
    if (remoteIK.ok && remoteIK.data?.data) {
      localStorage.setItem(KEY_IK_SEC, JSON.stringify(remoteIK.data.data));
    }
    if (remoteAssets.ok && remoteAssets.data?.data) {
      localStorage.setItem(KEY_ASSETS_SEC, JSON.stringify(remoteAssets.data.data));
    }
  } catch (e) {
    console.warn('Could not load remote Security SOP bundle:', e);
  }

  let listIK = getStoredIK();
  let listAssets = getStoredAssets();

  const SEC_CATEGORY_THEMES = {
    'Pengamanan Fisik & Patroli': { dot: '#60A5FA', text: '#93C5FD', accent: '#3B82F6' },
    'Operasional Parkir & Luar': { dot: '#FBBF24', text: '#FDE68A', accent: '#F59E0B' },
    'Layanan & Hospitalitas': { dot: '#34D399', text: '#A7F3D0', accent: '#10B981' },
  };
  function getSecIKTheme(ik) {
    return SEC_CATEGORY_THEMES[ik?.category] || SEC_CATEGORY_THEMES['Pengamanan Fisik & Patroli'];
  }
  function renderSecHandbookIcon() {
    return `
      <span style="width:40px; height:40px; flex-shrink:0; display:inline-flex; align-items:center; justify-content:center; background:#0E182B; border:1px solid #223456; border-radius:9px; color:#93C5FD;" aria-hidden="true">
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">
          <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"></path>
        </svg>
      </span>
    `;
  }
  function renderSecHandbookActions(ik) {
    if (!canEdit && !canDelete) return '';
    const titleSafe = escapeHtml(ik.title || 'Instruksi Kerja');
    return `
      <div style="display:flex; gap:5px; flex-shrink:0;">
        ${canEdit ? `
        <button type="button" class="sec-icon-btn btn-edit-sec-ik" data-id="${ik.id}" title="Edit ${titleSafe}" aria-label="Edit SOP ${titleSafe}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
        </button>
        ` : ''}
        ${canDelete ? `
        <button type="button" class="sec-icon-btn sec-icon-del btn-delete-sec-ik" data-id="${ik.id}" title="Hapus ${titleSafe}" aria-label="Hapus SOP ${titleSafe}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
        ` : ''}
      </div>
    `;
  }

  const SEC_ZONE_THEMES = {
    'Zona Hijau': { title: '#6EE7B7', dot: '#34D399', bg: 'rgba(16, 185, 129, 0.14)', border: 'rgba(52, 211, 153, 0.36)', body: '#F8FAFC' },
    'Zona Kuning': { title: '#FCD34D', dot: '#FBBF24', bg: 'rgba(245, 158, 11, 0.14)', border: 'rgba(251, 191, 36, 0.36)', body: '#F8FAFC' },
    'Zona Merah': { title: '#FCA5A5', dot: '#F87171', bg: 'rgba(239, 68, 68, 0.14)', border: 'rgba(248, 113, 113, 0.36)', body: '#F8FAFC' },
  };

  function renderPage() {
    container.innerHTML = `
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.65rem; font-weight:700; color:#FFFFFF;">
            <span>🛡️</span> SOP Pengelolaan Keamanan & Aset Perusahaan
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.92rem; color:#CBD5E1; line-height:1.5;">
            Standar Operasional Prosedur Pengawasan Keamanan Area, Fasilitas, dan Perlindungan Aset PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          ${canDelete ? `
          <button id="btn-reset-sec-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#131F36; border:1px solid #223456; color:#E2E8F0; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem;">
            🔄 Reset ke Standar
          </button>
          ` : ''}
          <button id="btn-print-sec" class="btn btn-outline" style="background:#131F36; border:1px solid #223456; color:#E2E8F0; font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem;">
            🖨️ Cetak Ringkasan SOP
          </button>
          ${hasPermission(currentUser, "export") ? `<button id="btn-export-sec" class="btn btn-primary" style="font-weight:700; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; color:#FFFFFF;">
            📥 Download Data Excel
          </button>` : ''}
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div class="fc-nav-tabs-strip" data-fc-scroll-tabs style="display:flex; gap:8px; margin-bottom:20px; border-bottom:1px solid #223456; padding-bottom:10px; overflow-x:auto; flex-wrap:nowrap; -webkit-overflow-scrolling:touch; scrollbar-width:none;">
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
      return 'border:1px solid #3B82F6; background:#2563EB; color:#FFFFFF; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:8px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;';
    }
    return 'background:#172033; color:#E2E8F0; border:1px solid #334155; font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:8px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;';
  }

  function getBadgeStyle(isActive) {
    if (isActive) {
      return 'background:#0F172A; color:#FFFFFF; font-size:0.75rem; padding:2px 8px; border-radius:10px; font-weight:700;';
    }
    return 'background:#22314A; color:#E2E8F0; border:1px solid #334155; font-size:0.75rem; padding:2px 8px; border-radius:10px; font-weight:700;';
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
          Promise.all([
            apiFetch('/api/sop/bundle/' + KEY_IK_SEC, { method: 'DELETE' }),
            apiFetch('/api/sop/bundle/' + KEY_ASSETS_SEC, { method: 'DELETE' }),
          ]).catch(() => {});
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
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:260px;">
            <div style="position:relative; flex:1; min-width:200px; max-width:340px;">
              <input type="text" id="sec-ik-search" class="form-control" placeholder="Cari instruksi kerja security..." value="${currentSearch}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%; background:#0E182B; border-color:#223456;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:#94A3B8; font-size:0.9rem;">🔍</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="sec-cat-pills-wrap">
              <button class="sec-cat-pill ${currentCategory === 'all' ? 'active' : ''}" data-cat="all">Semua (${listIK.length})</button>
              <button class="sec-cat-pill ${currentCategory === 'Pengamanan Fisik & Patroli' ? 'active' : ''}" data-cat="Pengamanan Fisik & Patroli">Patroli & Fisik (${listIK.filter(x => x.category === 'Pengamanan Fisik & Patroli').length})</button>
              <button class="sec-cat-pill ${currentCategory === 'Operasional Parkir & Luar' ? 'active' : ''}" data-cat="Operasional Parkir & Luar">Parkir & Luar (${listIK.filter(x => x.category === 'Operasional Parkir & Luar').length})</button>
              <button class="sec-cat-pill ${currentCategory === 'Layanan & Hospitalitas' ? 'active' : ''}" data-cat="Layanan & Hospitalitas">Tamu & Pasien (${listIK.filter(x => x.category === 'Layanan & Hospitalitas').length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="sec-ik-count-label" style="font-size:0.86rem; color:#CBD5E1; font-weight:600;">
              Menampilkan ${listIK.length} Instruksi Kerja
            </div>
            ${canEdit ? `
            <button id="btn-add-sec-ik" class="btn btn-primary" style="font-weight:700; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
              ➕ Tambah IK Baru
            </button>
            ` : ''}
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
        border: 1px solid #223456;
        background: #0E182B;
        color: #CBD5E1;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 13px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
        flex-shrink: 0;
        white-space: nowrap;
      }
      .sec-cat-pill:hover { background: #1A2A48; color:#FFFFFF; border-color:#3B82F6; }
      .sec-cat-pill.active {
        background: #2563EB;
        color: #FFFFFF;
        border-color: #3B82F6;
      }
      .sec-ik-card:hover {
        box-shadow: 0 8px 20px rgba(0,0,0,0.28) !important;
        border-color: #3B82F6 !important;
      }
      .sec-icon-btn {
        background: transparent;
        border: 1px solid rgba(51, 65, 85, 0.7);
        border-radius: 7px;
        width: 28px;
        height: 28px;
        padding: 0;
        color: #94A3B8;
        cursor: pointer;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        transition: all 0.15s;
      }
      .sec-icon-btn:hover {
        background: rgba(37, 99, 235, 0.18);
        border-color: rgba(96, 165, 250, 0.55);
        color: #93C5FD;
      }
      .sec-icon-del:hover {
        background: rgba(239, 68, 68, 0.16);
        border-color: rgba(248, 113, 113, 0.55);
        color: #FCA5A5;
      }
      .sec-icon-btn:focus-visible {
        outline: 2px solid #60A5FA;
        outline-offset: 2px;
      }
      @media (max-width: 768px) {
        #sec-cat-pills-wrap {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          scrollbar-width: none !important;
          padding-bottom: 4px !important;
          width: 100% !important;
        }
        #sec-cat-pills-wrap::-webkit-scrollbar {
          display: none !important;
        }
        .sec-ik-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
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
          <div style="text-align:center; padding:48px 20px; background:#111827; border-radius:12px; border:1px solid #29364B;">
            <div style="font-size:2.5rem; margin-bottom:8px;">🔍</div>
            <h3 style="margin:0 0 6px; font-size:1.12rem; color:#FFFFFF;">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.88rem; color:#CBD5E1;">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;
        return;
      }

      gridContainer.innerHTML = `
        <div class="sec-ik-grid" style="display:grid; grid-template-columns:repeat(auto-fill, minmax(340px, 1fr)); gap:18px;">
          ${filteredIK.map(ik => {
            const theme = getSecIKTheme(ik);
            return `
            <div class="card sec-ik-card" style="background:#131F36; border:1px solid #223456; border-radius:14px; padding:20px; display:flex; flex-direction:column; justify-content:space-between; transition:border-color 0.18s ease, box-shadow 0.18s ease;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:10px; margin-bottom:14px;">
                  <div style="display:flex; align-items:flex-start; gap:12px; min-width:0;">
                    ${renderSecHandbookIcon()}
                    <div style="min-width:0;">
                      <h3 style="margin:0 0 6px 0; font-size:1.06rem; font-weight:700; color:#FFFFFF; line-height:1.35;">
                        IK-${String(ik.id).padStart(2, '0')}. ${ik.title}
                      </h3>
                      <span class="badge" style="background:#0E182B; color:${theme.text}; border:1px solid #24385A; font-size:0.74rem; font-weight:600; border-radius:5px; padding:2px 8px; display:inline-flex; align-items:center; gap:6px;">
                        <span style="width:6px; height:6px; border-radius:50%; background:${theme.dot}; display:inline-block;"></span>
                        ${ik.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus (non-dominant SVG) -->
                  ${renderSecHandbookActions(ik)}
                </div>

                <div style="font-size:0.88rem; color:#F1F5F9; line-height:1.55; margin-bottom:14px; background:#0E182B; padding:10px 13px; border-radius:8px; border:1px solid #1E3050; border-left:3px solid ${theme.accent};">
                  <strong style="color:#FFFFFF;">Area Sasaran:</strong> ${ik.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.75rem; font-weight:700; color:#93C5FD; text-transform:uppercase; letter-spacing:0.05em; margin-bottom:7px;">Alat &amp; Perlengkapan</div>
                  <ul style="margin:0; padding:0; list-style:none; font-size:0.86rem; color:#E2E8F0; line-height:1.5;">
                    ${ik.tools.slice(0, 3).map(t => `
                      <li style="display:flex; align-items:flex-start; gap:8px; padding:4px 0; border-bottom:1px solid rgba(34, 52, 86, 0.45); word-break:break-word;">
                        <span style="color:#60A5FA; font-weight:700; line-height:1.4;">•</span>
                        <span>${escapeHtml(t)}</span>
                      </li>
                    `).join('')}
                    ${ik.tools.length > 3 ? `<li style="color:#93C5FD; font-size:0.82rem; font-weight:600; padding-top:5px;">+ ${ik.tools.length - 3} perlengkapan lainnya dalam panduan</li>` : ''}
                  </ul>
                </div>
              </div>

              <div style="border-top:1px solid #1E2E4A; padding-top:13px; margin-top:6px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.84rem; color:#CBD5E1; font-weight:600;">
                  ${ik.steps.length} Langkah Prosedur
                </span>
                <button type="button" class="btn btn-sm btn-primary btn-view-sec-ik" data-id="${ik.id}" style="padding:7px 15px; font-weight:700; font-size:0.85rem; border-radius:8px; background:#2563EB; border:1px solid #3B82F6; color:#FFFFFF;">
                  Lihat Panduan Lengkap →
                </button>
              </div>
            </div>
            `;
          }).join('')}
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
    const theme = getSecIKTheme(ik);
    createModal({
      title: `${ik.icon || '🛡️'} IK-${String(ik.id).padStart(2, '0')}: ${ik.title}`,
      size: 'lg',
      content: `
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${theme.bg}; color:${theme.text}; border:1px solid ${theme.border}; padding:11px 15px; border-radius:9px; font-weight:700; font-size:0.92rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${ik.category}</span>
            <span style="color:#FFFFFF;">Kode: IK-SEC-${String(ik.id).padStart(2, '0')}</span>
          </div>

          <div style="background:#0F172A; padding:12px 15px; border-radius:9px; border:1px solid #29364B; border-left:4px solid ${theme.accent};">
            <strong style="color:#FFFFFF; font-size:0.88rem; display:block; margin-bottom:4px;">📍 Area & Sasaran Target:</strong>
            <span style="font-size:0.92rem; color:#F8FAFC; line-height:1.55;">${ik.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.98rem; font-weight:700; color:#FFFFFF;">🧰 Alat & Perlengkapan yang Disiapkan:</h4>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:9px; padding:12px 16px;">
              <ul style="margin:0 0 0 18px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.75;">
                ${ik.tools.map(t => {
                  if (t.includes('http')) {
                    const parts = t.split(': ');
                    return `<li style="margin-bottom:4px;">${parts[0]}: <a href="${parts.slice(1).join(': ')}" target="_blank" rel="noopener" style="color:#60A5FA; font-weight:700; text-decoration:underline;">Buka Tautan Form ➔</a></li>`;
                  }
                  return `<li style="margin-bottom:4px;">${escapeHtml(t)}</li>`;
                }).join('')}
              </ul>
            </div>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.98rem; font-weight:700; color:#FFFFFF;">📝 Langkah & Prosedur Pelaksanaan:</h4>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:9px; padding:13px 18px;">
              <ol style="margin:0 0 0 20px; padding:0; font-size:0.9rem; color:#F8FAFC; line-height:1.8;">
                ${ik.steps.map(s => `<li style="margin-bottom:7px;">${escapeHtml(s)}</li>`).join('')}
              </ol>
            </div>
          </div>

          ${ik.notes ? `
            <div style="background:rgba(59, 130, 246, 0.14); border:1px solid rgba(96, 165, 250, 0.36); border-radius:9px; padding:12px 15px;">
              <strong style="color:#93C5FD; font-size:0.88rem; display:block; margin-bottom:3px;">⚠️ Catatan Khusus & Ketentuan Penting:</strong>
              <span style="color:#F8FAFC; font-size:0.88rem; line-height:1.6;">${escapeHtml(ik.notes)}</span>
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
    if (!canEdit) return;
    const isEdit = !!existingIK;
    const initialTools = existingIK ? existingIK.tools.join('\n') : '';
    const initialSteps = existingIK ? existingIK.steps.join('\n') : '';

    const contentHtml = `
      <form id="sec-ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-sec-ik-title" class="form-control" required value="${existingIK ? escapeHtml(existingIK.title) : ''}" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-sec-ik-icon" class="form-control" value="${existingIK ? existingIK.icon : '🛡️'}" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
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
          <input type="text" id="form-sec-ik-target" class="form-control" required value="${existingIK ? escapeHtml(existingIK.target) : ''}" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Perlengkapan (1 per baris) *</label>
          <textarea id="form-sec-ik-tools" class="form-control" rows="3" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${escapeHtml(initialTools)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah & Prosedur Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-sec-ik-steps" class="form-control" rows="5" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${escapeHtml(initialSteps)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Khusus / Ketentuan Penting</label>
          <textarea id="form-sec-ik-notes" class="form-control" rows="2" style="width:100%; border-radius:6px; font-size:0.85rem;">${existingIK ? escapeHtml(existingIK.notes || '') : ''}</textarea>
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
        <div class="btn-group" style="background:#111827; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-cabang-btn" class="btn btn-sm ${zonasiView === 'cabang' ? 'btn-primary' : 'btn-ghost'}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            🏥 Zonasi Klinik / Cabang
          </button>
          <button id="view-hq-btn" class="btn btn-sm ${zonasiView === 'hq' ? 'btn-primary' : 'btn-ghost'}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            🏢 Zonasi Head Quarter (5 Lantai)
          </button>
        </div>
        <div style="font-size:0.86rem; color:#CBD5E1; font-weight:600;">
          Ketentuan Pembagian Akses Ruangan (Lampiran 8.1 SOP)
        </div>
      </div>

      <!-- Definisi 3 Zona Header Cards -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px; margin-bottom:22px;">
        <div class="card" style="background:rgba(239, 68, 68, 0.14); border:1px solid rgba(248, 113, 113, 0.36); border-radius:11px; padding:15px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:13px; height:13px; border-radius:50%; background:#F87171; display:inline-block;"></span>
            <strong style="color:#FCA5A5; font-size:0.96rem;">ZONA MERAH (Restricted Area)</strong>
          </div>
          <p style="margin:0; font-size:0.86rem; color:#F8FAFC; line-height:1.55;">
            Area sangat terbatas. Hanya personal yang bertugas spesifik di area tersebut dan BOD. Tamu WAJIB mendapat izin tertulis Manajemen dan didampingi penuh.
          </p>
        </div>

        <div class="card" style="background:rgba(245, 158, 11, 0.14); border:1px solid rgba(251, 191, 36, 0.36); border-radius:11px; padding:15px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:13px; height:13px; border-radius:50%; background:#FBBF24; display:inline-block;"></span>
            <strong style="color:#FCD34D; font-size:0.96rem;">ZONA KUNING (Limited Area)</strong>
          </div>
          <p style="margin:0; font-size:0.86rem; color:#F8FAFC; line-height:1.55;">
            Area dengan pengamanan terbatas. Karyawan umum atau tamu yang memasuki area ini perlu mendapatkan arahan dan didampingi staf area terkait.
          </p>
        </div>

        <div class="card" style="background:rgba(16, 185, 129, 0.14); border:1px solid rgba(52, 211, 153, 0.36); border-radius:11px; padding:15px;">
          <div style="display:flex; align-items:center; gap:8px; margin-bottom:6px;">
            <span style="width:13px; height:13px; border-radius:50%; background:#34D399; display:inline-block;"></span>
            <strong style="color:#6EE7B7; font-size:0.96rem;">ZONA HIJAU (Public Area)</strong>
          </div>
          <p style="margin:0; font-size:0.86rem; color:#F8FAFC; line-height:1.55;">
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
          ${ZONASI_SECURITY.cabang.map(z => {
            const zt = SEC_ZONE_THEMES[z.zone] || SEC_ZONE_THEMES['Zona Hijau'];
            return `
            <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
              <div style="display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:12px; border-bottom:1px solid #29364B; padding-bottom:10px;">
                <h3 style="margin:0; font-size:1.08rem; font-weight:700; color:${zt.title}; display:flex; align-items:center; gap:8px;">
                  <span style="width:12px; height:12px; border-radius:50%; background:${zt.dot}; display:inline-block;"></span>
                  ${z.zone}
                </h3>
                <span class="badge" style="background:${zt.bg}; color:${zt.title}; border:1px solid ${zt.border}; font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">
                  ${z.level}
                </span>
              </div>
              <p style="font-size:0.86rem; color:#CBD5E1; margin-bottom:14px; line-height:1.55;">${z.desc}</p>
              <div>
                <strong style="font-size:0.8rem; text-transform:uppercase; color:#E2E8F0; letter-spacing:0.03em; display:block; margin-bottom:7px;">Daftar Ruangan Terkait:</strong>
                <ul style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.75;">
                  ${z.rooms.map(r => `<li style="margin-bottom:2px;">${r}</li>`).join('')}
                </ul>
              </div>
            </div>
            `;
          }).join('')}
        </div>
      `;
    } else {
      subContent.innerHTML = `
        <div style="display:flex; flex-direction:column; gap:16px;">
          ${ZONASI_SECURITY.hq.map(fl => `
            <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:18px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:10px;">
                <h3 style="margin:0; font-size:1.12rem; font-weight:700; color:#FFFFFF; display:flex; align-items:center; gap:8px;">
                  <span>🏢</span> Head Quarter — ${fl.floor}
                </h3>
              </div>
              <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:14px;">
                <!-- Zona Hijau -->
                <div style="background:rgba(16, 185, 129, 0.14); border:1px solid rgba(52, 211, 153, 0.35); border-radius:9px; padding:13px;">
                  <strong style="color:#6EE7B7; font-size:0.88rem; display:block; margin-bottom:7px;">🟢 ZONA HIJAU (Publik):</strong>
                  <ul style="margin:0 0 0 18px; padding:0; font-size:0.86rem; color:#F8FAFC; line-height:1.65;">
                    ${fl.hijau.map(item => `<li>${item}</li>`).join('')}
                  </ul>
                </div>
                <!-- Zona Kuning -->
                <div style="background:rgba(245, 158, 11, 0.14); border:1px solid rgba(251, 191, 36, 0.35); border-radius:9px; padding:13px;">
                  <strong style="color:#FCD34D; font-size:0.88rem; display:block; margin-bottom:7px;">🟡 ZONA KUNING (Terbatas):</strong>
                  <ul style="margin:0 0 0 18px; padding:0; font-size:0.86rem; color:#F8FAFC; line-height:1.65;">
                    ${fl.kuning.map(item => `<li>${item}</li>`).join('')}
                  </ul>
                </div>
                <!-- Zona Merah -->
                <div style="background:rgba(239, 68, 68, 0.14); border:1px solid rgba(248, 113, 113, 0.35); border-radius:9px; padding:13px;">
                  <strong style="color:#FCA5A5; font-size:0.88rem; display:block; margin-bottom:7px;">🔴 ZONA MERAH (Restricted):</strong>
                  <ul style="margin:0 0 0 18px; padding:0; font-size:0.86rem; color:#F8FAFC; line-height:1.65;">
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
          <h3 style="margin:0 0 4px 0; font-size:1.12rem; font-weight:700; color:#FFFFFF;">
            Daftar Klasifikasi Aset yang Dipantau Security (Lampiran 8.3)
          </h3>
          <p style="margin:0; font-size:0.88rem; color:#CBD5E1;">
            Pengawasan dan perlindungan sarana prasarana fisik, elektronik, dan peralatan klinik
          </p>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="sec-asset-search" class="form-control" placeholder="Cari nama aset / tipe / PIC..." value="${assetSearch}" style="height:36px; border-radius:8px; font-size:0.88rem; width:240px;">
            ${canEdit ? `
          <button id="btn-add-sec-asset" class="btn btn-primary" style="font-weight:700; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.86rem; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
            ➕ Tambah Aset Baru
          </button>
            ` : ''}
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
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; overflow:hidden; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="padding:14px 18px; background:#111827; border-bottom:1px solid #29364B; display:flex; justify-content:space-between; align-items:center;">
            <strong style="color:#FFFFFF; font-size:0.96rem;">Rekap Klasifikasi Aset Perusahaan</strong>
            <span class="badge" style="background:rgba(59,130,246,0.2); color:#93C5FD; border:1px solid rgba(96,165,250,0.4); font-weight:700; padding:3px 9px; border-radius:6px;">${filtered.length} Aset</span>
          </div>
          <div style="overflow-x:auto;">
            <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.9rem;">
              <thead>
                <tr style="background:#111827; border-bottom:1px solid #29364B; text-align:left;">
                  <th style="padding:11px 14px; width:60px; color:#CBD5E1;">No</th>
                  <th style="padding:11px 14px; width:145px; color:#CBD5E1;">Tipe Aset</th>
                  <th style="padding:11px 14px; width:220px; color:#CBD5E1;">Nama Aset / Kelompok</th>
                  <th style="padding:11px 14px; color:#CBD5E1;">Deskripsi / Contoh Item</th>
                  <th style="padding:11px 14px; width:190px; color:#CBD5E1;">PIC Pengawas</th>
                  <th style="padding:11px 14px; width:110px; text-align:center; color:#CBD5E1;">Aksi</th>
                </tr>
              </thead>
              <tbody>
                ${filtered.map((item, idx) => `
                  <tr style="border-bottom:1px solid #1E2E4A; background:${idx % 2 === 0 ? '#131F36' : '#0E182B'}; color:#F8FAFC;">
                    <td style="padding:11px 14px; font-weight:600; color:#CBD5E1;">${item.no || idx + 1}</td>
                    <td style="padding:11px 14px;">
                      <span class="badge" style="${getTypeBadgeStyle(item.type)}">${escapeHtml(item.type)}</span>
                    </td>
                    <td style="padding:11px 14px; font-weight:700; color:#FFFFFF;">${escapeHtml(item.item)}</td>
                    <td style="padding:11px 14px; color:#CBD5E1; line-height:1.5;">${escapeHtml(item.description)}</td>
                    <td style="padding:11px 14px; font-weight:700; color:#93C5FD;">${escapeHtml(item.pic || '-')}</td>
                    <td style="padding:11px 14px; text-align:center;">
                      <div style="display:inline-flex; gap:5px;">
                    ${canEdit ? `
                      <button type="button" class="fc-action-btn fc-action-edit btn-edit-asset" data-no="${item.no}" title="Edit Aset" aria-label="Edit aset ${escapeHtml(item.item || '')}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>
                    ` : ''}
                    ${canDelete ? `
                      <button type="button" class="fc-action-btn fc-action-del btn-delete-asset" data-no="${item.no}" title="Hapus Aset" aria-label="Hapus aset ${escapeHtml(item.item || '')}"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>
                    ` : ''}
                      </div>
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
      if (type === 'Aset Tetap') return 'background:rgba(59,130,246,0.18); color:#93C5FD; border:1px solid rgba(96,165,250,0.38); font-weight:700; padding:3px 9px; border-radius:6px;';
      if (type === 'Aset Bergerak') return 'background:rgba(16,185,129,0.18); color:#6EE7B7; border:1px solid rgba(52,211,153,0.38); font-weight:700; padding:3px 9px; border-radius:6px;';
      return 'background:rgba(245,158,11,0.18); color:#FCD34D; border:1px solid rgba(251,191,36,0.38); font-weight:700; padding:3px 9px; border-radius:6px;';
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
    if (!canEdit) return;
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
          <input type="text" id="form-asset-item" class="form-control" required value="${existingAsset ? escapeHtml(existingAsset.item) : ''}" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Deskripsi / Contoh Item *</label>
          <textarea id="form-asset-desc" class="form-control" rows="3" required style="width:100%; border-radius:6px; font-size:0.85rem;">${existingAsset ? escapeHtml(existingAsset.description) : ''}</textarea>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">PIC Pengawas / Penanggung Jawab *</label>
          <input type="text" id="form-asset-pic" class="form-control" required value="${existingAsset ? escapeHtml(existingAsset.pic) : ''}" style="width:100%; border-radius:6px; height:36px;">
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
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(59,130,246,0.16); border:1px solid rgba(96,165,250,0.35); padding:8px; border-radius:9px;">⏰</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Sistem Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Ketentuan 12 Jam Kerja dengan 2 Hari Libur per Minggu</span>
            </div>
          </div>
          <div style="font-size:0.9rem; color:#F8FAFC; line-height:1.7;">
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:8px; padding:11px 14px; margin-bottom:10px; border-left:3.5px solid #60A5FA;">
              <strong style="color:#FFFFFF; display:block; margin-bottom:2px;">☀️ Shift 1 (Pagi - Siang):</strong>
              <span style="color:#E2E8F0;">Pukul <strong style="color:#FFFFFF;">07.00 s/d 19.00 WIB</strong></span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; border-radius:8px; padding:11px 14px; margin-bottom:12px; border-left:3.5px solid #A78BFA;">
              <strong style="color:#FFFFFF; display:block; margin-bottom:2px;">🌙 Shift 2 (Malam):</strong>
              <span style="color:#E2E8F0;">Pukul <strong style="color:#FFFFFF;">19.00 s/d 07.00 WIB</strong></span>
            </div>
            <p style="margin:0; font-size:0.85rem; color:#CBD5E1; line-height:1.6;">
              <strong style="color:#FFFFFF;">Ketentuan Off:</strong> Berlaku 2 hari off dalam setiap pekannya sesuai jadwal roster yang ditetapkan oleh Facility Care Supervisor / General Support.
            </p>
          </div>
        </div>

        <!-- Card 2: Jadwal Safety Patrol -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(245,158,11,0.16); border:1px solid rgba(251,191,36,0.35); padding:8px; border-radius:9px;">🚶‍♂️</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Jadwal Safety Patrol Harian (6.2 B)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Keliling Pengecekan Area & Form Checklist Digital</span>
            </div>
          </div>
          <div style="display:flex; flex-direction:column; gap:9px; font-size:0.88rem;">
            <div style="background:#0F172A; border:1px solid #29364B; padding:9px 13px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#F8FAFC;">Patroli 1 (Pagi):</strong>
              <span class="badge" style="background:rgba(56,189,248,0.16); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); font-weight:700; padding:3px 9px; border-radius:6px;">07.30 - 08.30 WIB</span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; padding:9px 13px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#F8FAFC;">Patroli 2 (Malam):</strong>
              <span class="badge" style="background:rgba(245,158,11,0.16); color:#FCD34D; border:1px solid rgba(251,191,36,0.35); font-weight:700; padding:3px 9px; border-radius:6px;">19.30 - 20.30 WIB</span>
            </div>
            <div style="background:#0F172A; border:1px solid #29364B; padding:9px 13px; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:#F8FAFC;">Patroli 3 (Dini Hari):</strong>
              <span class="badge" style="background:rgba(139,92,246,0.18); color:#DDD6FE; border:1px solid rgba(167,139,250,0.38); font-weight:700; padding:3px 9px; border-radius:6px;">04.00 - 05.00 WIB</span>
            </div>
            <div style="font-size:0.84rem; color:#CBD5E1; margin-top:4px; line-height:1.5;">
              * <em>Patroli 3 khusus berlaku untuk area Head Quarter atau lokasi yang memiliki lebih dari 1 personil security.</em>
            </div>
          </div>
        </div>

        <!-- Card 3: Penanganan Pelanggaran Disiplin -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(16,185,129,0.16); border:1px solid rgba(52,211,153,0.35); padding:8px; border-radius:9px;">⚠️</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Penanganan Pelanggaran Disiplin (6.4)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Pelanggaran SK Tata Tertib Perusahaan</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.75;">
            <li style="margin-bottom:4px;">Petugas security membuat <strong style="color:#FFFFFF;">Berita Acara Umum (BA)</strong> kejadian secara lengkap.</li>
            <li style="margin-bottom:4px;">Melaporkan kronologis tertulis kepada atasan (General Support Manager / FC Supervisor).</li>
            <li>Melaporkan dan meneruskan berkas kepada atasan langsung dari karyawan yang bersangkutan untuk penanganan HR/HC.</li>
          </ol>
        </div>

        <!-- Card 4: Penanganan Pelanggaran Kriminal -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2);">
          <div style="display:flex; align-items:center; gap:12px; margin-bottom:14px; border-bottom:1px solid #29364B; padding-bottom:12px;">
            <span style="font-size:1.5rem; background:rgba(239,68,68,0.16); border:1px solid rgba(248,113,113,0.35); padding:8px; border-radius:9px;">🚨</span>
            <div>
              <h3 style="margin:0 0 2px; font-size:1.12rem; font-weight:700; color:#FCA5A5;">Penanganan Pelanggaran Kriminal (6.4)</h3>
              <span style="font-size:0.84rem; color:#CBD5E1;">Tindak Pidana, Pencurian, & Pelanggaran Hukum</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.75;">
            <li style="margin-bottom:3px;">Buat <strong style="color:#FFFFFF;">Berita Acara Umum</strong> secara resmi.</li>
            <li style="margin-bottom:3px;">Periksa dan amankan rekaman bukti visual melalui <strong style="color:#FFFFFF;">CCTV</strong>.</li>
            <li style="margin-bottom:3px;">Laporkan ke atasan dari karyawan terkait dengan rincian kronologis (bila pelaku karyawan).</li>
            <li style="margin-bottom:3px;">Minta dan catat detail identitas <strong style="color:#FFFFFF;">KTP / ID Card</strong> dari pihak yang bersangkutan.</li>
            <li style="margin-bottom:3px;">Amankan dan netralisir <strong style="color:#FFFFFF;">Tempat Kejadian Perkara (TKP)</strong> dari gangguan orang lain.</li>
            <li>Laporkan ke <strong style="color:#FFFFFF;">Kepolisian terdekat</strong> untuk proses hukum tindak lanjut.</li>
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
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">📑</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Form Safety Patrol Audy</h3>
                <span class="badge" style="background:rgba(56,189,248,0.16); color:#7DD3FC; border:1px solid rgba(56,189,248,0.35); font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Google Forms Resmi</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi keamanan yang wajib diisi oleh tim security saat melakukan Safety Patrol rutin di klinik dan kantor pusat Audy Dental.
            </p>
          </div>
          <a href="https://forms.gle/1EAvkqDJVdctXEGn7" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:700; padding:10px; border-radius:8px; background:#2563EB; border-color:#3B82F6; color:#FFFFFF;">
            📋 Buka Form Safety Patrol ➔
          </a>
        </div>

        <!-- Card: Dasar Hukum & Ruang Lingkup Perusahaan -->
        <div class="card" style="background:#172033; border:1px solid #29364B; border-radius:12px; padding:20px; box-shadow:0 4px 12px rgba(0,0,0,0.2); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
              <span style="font-size:2rem;">⚖️</span>
              <div>
                <h3 style="margin:0 0 4px; font-size:1.12rem; font-weight:700; color:#FFFFFF;">Dasar Hukum & Ruang Lingkup</h3>
                <span class="badge" style="background:#22314A; color:#E2E8F0; border:1px solid #334155; font-size:0.78rem; font-weight:700; padding:3px 9px; border-radius:6px;">Legalitas</span>
              </div>
            </div>
            <p style="font-size:0.88rem; color:#CBD5E1; line-height:1.6; margin-bottom:12px;">
              SOP ini diterapkan sebagai tata cara Pengelolaan Keamanan Aset Perusahaan di seluruh lokasi Audy Dental Klinik di bawah naungan:
            </p>
            <ul style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:#F8FAFC; line-height:1.7;">
              <li><strong style="color:#FFFFFF;">PT. Audy Mandiri Indonesia</strong></li>
              <li><strong style="color:#FFFFFF;">PT. Kastara Ananta</strong></li>
              <li><strong style="color:#FFFFFF;">PT. Nara Adidaya</strong></li>
            </ul>
          </div>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#172033; border:1px solid #29364B; border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1.05rem; font-weight:700; color:#FFFFFF; display:flex; align-items:center; gap:8px;">
          <span>📜</span> Validasi & Pengesahan Dokumen SOP Security
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3.5px solid #60A5FA; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3.5px solid #34D399; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">General Support Manager</div>
          </div>
          <div style="border-left:3.5px solid #A78BFA; padding-left:12px;">
            <div style="font-size:0.78rem; color:#CBD5E1; font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.98rem; font-weight:700; color:#FFFFFF; margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.84rem; color:#CBD5E1;">Chief Executive Officer</div>
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
