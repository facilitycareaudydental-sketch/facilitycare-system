import { apiFetch } from '../config.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// ── INITIAL PRE-SEEDED DOCUMENTS (Sesuai Master Dokumen Operasional FCMS) ────
const INITIAL_DOCUMENTS = [
  {
    id: 1,
    name: 'Berita Acara Kerusakan Alat',
    category: 'Berita Acara',
    document_link: 'https://docs.google.com/document/d/1example_ba_kerusakan',
    description: 'Format berita acara resmi pencatatan dan eskalasi kerusakan peralatan klinik dan aset fasilitas.'
  },
  {
    id: 2,
    name: 'Checklist Apar',
    category: 'Apar & K3',
    document_link: 'https://docs.google.com/spreadsheets/d/1example_apar',
    description: 'Lembar inspeksi berkala tabung pemadam api ringan (APAR), tekanan manometer, dan masa kadaluarsa.'
  },
  {
    id: 3,
    name: 'Checklist Deep Cleaning',
    category: 'Kegiatan Deep Cleaning',
    document_link: 'https://docs.google.com/spreadsheets/d/1example_deep_cleaning',
    description: 'Panduan dan verifikasi checklist pembersihan mendalam (deep cleaning) area dental unit dan klinik.'
  },
  {
    id: 4,
    name: 'Checklist Fulldraine',
    category: 'Drain Cleaning',
    document_link: 'https://docs.google.com/spreadsheets/d/1example_fulldraine',
    description: 'Checklist monitoring pembersihan saluran pembuangan air kotor, floor drain, dan sedotan suction.'
  },
  {
    id: 5,
    name: 'Checklist General Cleaning',
    category: 'Kegiatan General Cleaning',
    document_link: 'https://docs.google.com/spreadsheets/d/1example_general_cleaning',
    description: 'Standar checklist kebersihan harian/mingguan seluruh ruangan, lobby, toilet, dan koridor.'
  },
  {
    id: 6,
    name: 'Form Aktivity Reliefer',
    category: 'Form Operasional',
    document_link: 'https://forms.gle/1example_reliefer',
    description: 'Formulir pencatatan absensi dan penugasan harian reliefer Facility Care di cabang pengganti.'
  },
  {
    id: 7,
    name: 'Form Aktivity Security HQ',
    category: 'Form Operasional',
    document_link: 'https://forms.gle/1example_security_hq',
    description: 'Formulir monitoring shift dan pengawasan keamanan harian di Kantor Pusat (Head Quarter).'
  }
];

// Helper get category badge style
function getCategoryMeta(category = '') {
  const cat = (category || '').toLowerCase();
  if (cat.includes('berita acara')) {
    return { label: category, bg: '#F3E8FF', color: '#7E22CE', border: '#D8B4FE', icon: '⚖️' };
  }
  if (cat.includes('apar') || cat.includes('k3') || cat.includes('safety')) {
    return { label: category, bg: '#FEE2E2', color: '#B91C1C', border: '#FECACA', icon: '🧯' };
  }
  if (cat.includes('deep cleaning')) {
    return { label: category, bg: '#DCFCE7', color: '#15803D', border: '#BBF7D0', icon: '✨' };
  }
  if (cat.includes('drain')) {
    return { label: category, bg: '#FEF3C7', color: '#B45309', border: '#FDE68A', icon: '🚿' };
  }
  if (cat.includes('general cleaning') || cat.includes('kebersihan')) {
    return { label: category, bg: '#E0F2FE', color: '#0369A1', border: '#BAE6FD', icon: '🧹' };
  }
  if (cat.includes('form')) {
    return { label: category, bg: '#E0E7FF', color: '#4338CA', border: '#C7D2FE', icon: '📝' };
  }
  return { label: category || 'Dokumen', bg: '#F1F5F9', color: '#475569', border: '#CBD5E1', icon: '📁' };
}

// Helper detect link platform
function getPlatformMeta(url = '') {
  if (!url) return { type: 'Tidak Ada Tautan', icon: '📎', badgeClass: 'bg-slate-100 text-slate-500' };
  const lower = url.toLowerCase();
  if (lower.includes('forms.gle') || lower.includes('google.com/forms')) {
    return { type: 'Google Form', icon: '🟣', badgeClass: 'badge-form', color: '#673AB7' };
  }
  if (lower.includes('docs.google.com/spreadsheets') || lower.includes('sheet')) {
    return { type: 'Google Sheets', icon: '🟢', badgeClass: 'badge-sheet', color: '#0F9D58' };
  }
  if (lower.includes('docs.google.com/document') || lower.includes('doc')) {
    return { type: 'Google Docs', icon: '🔵', badgeClass: 'badge-doc', color: '#4285F4' };
  }
  if (lower.includes('drive.google.com')) {
    return { type: 'Google Drive', icon: '🟡', badgeClass: 'badge-drive', color: '#F4B400' };
  }
  if (lower.endsWith('.pdf')) {
    return { type: 'PDF Document', icon: '🔴', badgeClass: 'badge-pdf', color: '#EA4335' };
  }
  return { type: 'Web Link', icon: '🌐', badgeClass: 'badge-link', color: '#2563EB' };
}

export async function renderChecklist(container) {
  let documents = [];
  let activeCategory = 'all';
  let searchQuery = '';
  let currentView = localStorage.getItem('fcms_doc_view') || 'grid'; // 'grid' | 'table'

  // Inject Custom Styles
  const styleId = 'fcms-master-doc-styles';
  if (!document.getElementById(styleId)) {
    const styleEl = document.createElement('style');
    styleEl.id = styleId;
    styleEl.textContent = `
      .doc-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .doc-hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: #ffffff;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .doc-hero-card::after {
        content: '';
        position: absolute;
        right: -60px;
        top: -60px;
        width: 240px;
        height: 240px;
        background: radial-gradient(circle, rgba(59, 130, 246, 0.18) 0%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
      }
      .doc-hero-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
      }
      .doc-hero-title {
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
        letter-spacing: -0.02em;
      }
      .doc-hero-subtitle {
        color: #94A3B8;
        font-size: 0.92rem;
        margin: 0;
        max-width: 720px;
        line-height: 1.5;
      }
      .doc-stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin-top: 24px;
      }
      .doc-stat-box {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .doc-stat-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        background: rgba(59, 130, 246, 0.18);
        color: #60A5FA;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        flex-shrink: 0;
      }
      .doc-stat-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.2;
      }
      .doc-stat-label {
        font-size: 0.78rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
      }
      .doc-controls-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .doc-controls-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 14px;
        flex-wrap: wrap;
      }
      .doc-search-wrapper {
        position: relative;
        flex: 1;
        min-width: 260px;
      }
      .doc-search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94A3B8;
        font-size: 0.95rem;
        pointer-events: none;
      }
      .doc-search-input {
        width: 100%;
        padding: 10px 14px 10px 40px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.9rem;
        background: #F8FAFC;
        transition: all 0.2s ease;
      }
      .doc-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #FFFFFF;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .doc-view-toggle {
        display: flex;
        background: #F1F5F9;
        padding: 4px;
        border-radius: 10px;
        gap: 4px;
      }
      .doc-view-btn {
        border: none;
        background: transparent;
        padding: 6px 14px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748B;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 0.15s ease;
      }
      .doc-view-btn.active {
        background: #FFFFFF;
        color: #1E293B;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
      }
      .doc-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .doc-pill {
        border: 1px solid #E2E8F0;
        background: #F8FAFC;
        color: #475569;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 30px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .doc-pill:hover {
        background: #EEF2F6;
        border-color: #CBD5E1;
      }
      .doc-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .doc-pill-count {
        background: rgba(0, 0, 0, 0.08);
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 0.75rem;
      }
      .doc-pill.active .doc-pill-count {
        background: rgba(255, 255, 255, 0.25);
        color: #FFFFFF;
      }
      /* Grid View */
      .doc-cards-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(340px, 1fr));
        gap: 20px;
      }
      .doc-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
        overflow: hidden;
      }
      .doc-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.1);
        border-color: #CBD5E1;
      }
      .doc-card-top {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 12px;
        margin-bottom: 14px;
      }
      .doc-category-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        text-transform: uppercase;
      }
      .doc-platform-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        font-size: 0.74rem;
        font-weight: 600;
        color: #64748B;
        background: #F1F5F9;
        padding: 3px 8px;
        border-radius: 6px;
      }
      .doc-card-title {
        font-size: 1.08rem;
        font-weight: 700;
        color: #0F172A;
        margin: 0 0 8px 0;
        line-height: 1.4;
      }
      .doc-card-desc {
        color: #64748B;
        font-size: 0.85rem;
        line-height: 1.5;
        margin: 0 0 18px 0;
        flex-grow: 1;
        display: -webkit-box;
        -webkit-line-clamp: 3;
        -webkit-box-orient: vertical;
        overflow: hidden;
      }
      .doc-card-footer {
        padding-top: 14px;
        border-top: 1px solid #F1F5F9;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 8px;
      }
      .doc-btn-open {
        background: #2563EB;
        color: #FFFFFF;
        border: none;
        border-radius: 8px;
        padding: 8px 14px;
        font-size: 0.84rem;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        text-decoration: none;
        cursor: pointer;
        transition: background 0.15s ease;
      }
      .doc-btn-open:hover {
        background: #1D4ED8;
        color: #FFFFFF;
      }
      .doc-btn-open.disabled {
        background: #E2E8F0;
        color: #94A3B8;
        pointer-events: none;
      }
      .doc-card-actions {
        display: flex;
        align-items: center;
        gap: 4px;
      }
      .doc-icon-btn {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        color: #64748B;
        width: 34px;
        height: 34px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.15s ease;
      }
      .doc-icon-btn:hover {
        background: #EDF2F7;
        color: #1E293B;
        border-color: #CBD5E1;
      }
      .doc-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      /* Table View */
      .doc-table-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        overflow: hidden;
      }
      .doc-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
      }
      .doc-table th {
        background: #F8FAFC;
        padding: 14px 18px;
        font-size: 0.78rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        border-bottom: 1.5px solid #E2E8F0;
      }
      .doc-table td {
        padding: 16px 18px;
        font-size: 0.88rem;
        color: #1E293B;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: middle;
      }
      .doc-table tr:hover td {
        background: #F8FAFC;
      }
      .doc-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #FFFFFF;
        border-radius: 16px;
        border: 1px dashed #CBD5E1;
      }

      /* ── TAMPILAN RESPONSIVE HP / MOBILE (≤ 768px) ─────────────────── */
      @media (max-width: 768px) {
        .doc-page-container {
          padding: 12px 14px !important;
        }
        .doc-hero-card {
          padding: 18px 16px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
        }
        .doc-hero-title {
          font-size: 1.25rem !important;
          gap: 8px !important;
        }
        .doc-hero-subtitle {
          font-size: 0.82rem !important;
          line-height: 1.4 !important;
        }
        .doc-hero-header {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 12px !important;
        }
        .doc-hero-header > div:last-child {
          display: flex !important;
          width: 100% !important;
          gap: 8px !important;
        }
        #doc-btn-add {
          flex: 2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 12px !important;
        }
        #doc-btn-excel {
          flex: 1.2 !important;
          justify-content: center !important;
          font-size: 0.82rem !important;
          padding: 9px 10px !important;
        }
        #doc-btn-refresh {
          flex: 0 0 38px !important;
          justify-content: center !important;
          padding: 9px 0 !important;
        }
        /* 2x2 compact stats grid di layar HP */
        .doc-stats-grid {
          grid-template-columns: repeat(2, 1fr) !important;
          gap: 8px !important;
          margin-top: 14px !important;
        }
        .doc-stat-box {
          padding: 10px 12px !important;
          gap: 10px !important;
          border-radius: 10px !important;
        }
        .doc-stat-icon {
          width: 36px !important;
          height: 36px !important;
          font-size: 1.1rem !important;
          border-radius: 8px !important;
        }
        .doc-stat-value {
          font-size: 1.15rem !important;
        }
        .doc-stat-label {
          font-size: 0.68rem !important;
        }
        .doc-controls-panel {
          padding: 14px 12px !important;
          border-radius: 14px !important;
          margin-bottom: 16px !important;
          gap: 12px !important;
        }
        .doc-controls-top {
          flex-direction: column !important;
          align-items: stretch !important;
          gap: 10px !important;
        }
        .doc-search-wrapper {
          min-width: 100% !important;
          width: 100% !important;
        }
        .doc-view-toggle {
          width: 100% !important;
        }
        .doc-view-btn {
          flex: 1 !important;
          justify-content: center !important;
          padding: 7px 10px !important;
        }
        /* PITA FILTER SWIPEABLE HORIZONTAL DI HP (TIDAK LAGI NUMPUK VERTIKAL) */
        .doc-filter-pills {
          display: flex !important;
          flex-wrap: nowrap !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
          padding-bottom: 6px !important;
          padding-top: 2px !important;
          scrollbar-width: none !important;
          -ms-overflow-style: none !important;
        }
        .doc-filter-pills::-webkit-scrollbar {
          display: none !important;
        }
        .doc-pill {
          flex-shrink: 0 !important;
          white-space: nowrap !important;
          padding: 6px 12px !important;
          font-size: 0.8rem !important;
        }
        .doc-cards-grid {
          grid-template-columns: 1fr !important;
          gap: 12px !important;
        }
        .doc-card {
          padding: 16px !important;
          border-radius: 14px !important;
        }
        .doc-card-title {
          font-size: 1rem !important;
        }
        .doc-card-desc {
          font-size: 0.82rem !important;
        }
        .doc-table-card {
          border-radius: 14px !important;
          overflow-x: auto !important;
          -webkit-overflow-scrolling: touch !important;
        }
        .doc-table {
          min-width: 650px !important;
        }
        .doc-table th, .doc-table td {
          padding: 10px 12px !important;
          font-size: 0.82rem !important;
        }
      }
    `;
    document.head.appendChild(styleEl);
  }

  // Fetch Documents from API
  async function loadDocuments() {
    try {
      const res = await apiFetch('/api/checklist?limit=1000');
      if (res.ok && res.data?.data && Array.isArray(res.data.data) && res.data.data.length > 0) {
        documents = res.data.data;
        localStorage.setItem('fcms_master_doc_cache', JSON.stringify(documents));
      } else {
        const cached = localStorage.getItem('fcms_master_doc_cache');
        if (cached) {
          documents = JSON.parse(cached);
        } else {
          documents = [...INITIAL_DOCUMENTS];
          localStorage.setItem('fcms_master_doc_cache', JSON.stringify(documents));
        }
      }
    } catch (err) {
      console.warn('Fallback loading documents from cache/seed:', err);
      const cached = localStorage.getItem('fcms_master_doc_cache');
      documents = cached ? JSON.parse(cached) : [...INITIAL_DOCUMENTS];
    }
    renderUI();
  }

  // Filter & Search Logic
  function getFilteredDocuments() {
    return documents.filter(doc => {
      const matchCat = activeCategory === 'all' || doc.category === activeCategory;
      const q = searchQuery.toLowerCase().trim();
      const matchSearch = !q || 
        (doc.name || '').toLowerCase().includes(q) ||
        (doc.category || '').toLowerCase().includes(q) ||
        (doc.description || '').toLowerCase().includes(q);
      return matchCat && matchSearch;
    });
  }

  // Collect unique categories with counts
  function getCategoryStats() {
    const stats = { all: documents.length };
    documents.forEach(d => {
      const cat = d.category || 'Lainnya';
      stats[cat] = (stats[cat] || 0) + 1;
    });
    return stats;
  }

  // Add / Edit Modal Form
  function openDocumentModal(item = null) {
    const isEdit = !!item;
    const existingCats = Array.from(new Set(documents.map(d => d.category).filter(Boolean)));
    
    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Nama Dokumen / Checklist <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Checklist Pembersihan Wastafel & Toilet" value="${item?.name ? String(item.name).replace(/"/g, '&quot;') : ''}" required />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Kategori Dokumen <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="doc-modal-category" list="doc-categories-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Pilih atau ketik kategori baru..." value="${item?.category ? String(item.category).replace(/"/g, '&quot;') : ''}" required />
          <datalist id="doc-categories-list">
            ${existingCats.map(c => `<option value="${c}">${c}</option>`).join('')}
            <option value="Berita Acara">Berita Acara</option>
            <option value="Apar & K3">Apar & K3</option>
            <option value="Kegiatan Deep Cleaning">Kegiatan Deep Cleaning</option>
            <option value="Kegiatan General Cleaning">Kegiatan General Cleaning</option>
            <option value="Drain Cleaning">Drain Cleaning</option>
            <option value="Form Operasional">Form Operasional</option>
          </datalist>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Tautan / Link Dokumen (Google Drive, Docs, Sheets, atau Forms)
          </label>
          <input type="url" id="doc-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://docs.google.com/..." value="${item?.document_link ? String(item.document_link).replace(/"/g, '&quot;') : ''}" />
          <span style="font-size:0.75rem; color:#64748B; margin-top:4px; display:block;">
            Bisa berupa tautan Google Forms untuk inspeksi harian, atau link Google Sheets / Docs sebagai acuan.
          </span>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:6px;">
            Deskripsi / Petunjuk Penggunaan
          </label>
          <textarea id="doc-modal-desc" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan keterangan singkat mengenai tujuan dan instruksi pengisian dokumen ini...">${item?.description || ''}</textarea>
        </div>
      </div>
    `;

    createModal({
      title: isEdit ? '✏️ Edit Master Dokumen' : '➕ Tambah Master Dokumen Baru',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Tambahkan Dokumen',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const nameInput = content.querySelector('#doc-modal-name');
        const catInput = content.querySelector('#doc-modal-category');
        const linkInput = content.querySelector('#doc-modal-link');
        const descInput = content.querySelector('#doc-modal-desc');

        const name = (nameInput.value || '').trim();
        const category = (catInput.value || '').trim();
        const document_link = (linkInput.value || '').trim();
        const description = (descInput.value || '').trim();

        if (!name) {
          toastError('Nama Dokumen wajib diisi');
          nameInput.focus();
          return;
        }
        if (!category) {
          toastError('Kategori Dokumen wajib diisi');
          catInput.focus();
          return;
        }

        const payload = { name, category, document_link, description };

        try {
          if (isEdit) {
            const res = await apiFetch(`/api/checklist/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              const idx = documents.findIndex(d => d.id === item.id);
              if (idx !== -1) documents[idx] = { ...documents[idx], ...payload };
              toastSuccess('Dokumen berhasil diperbarui');
            } else {
              // Local fallback
              const idx = documents.findIndex(d => d.id === item.id);
              if (idx !== -1) documents[idx] = { ...documents[idx], ...payload };
              toastSuccess('Dokumen disimpan secara lokal');
            }
          } else {
            const res = await apiFetch('/api/checklist', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
            if (res.ok && res.data?.id) {
              documents.unshift({ id: res.data.id, ...payload });
              toastSuccess('Dokumen baru berhasil ditambahkan');
            } else {
              // Local fallback
              const newId = Date.now();
              documents.unshift({ id: newId, ...payload });
              toastSuccess('Dokumen baru ditambahkan secara lokal');
            }
          }
          localStorage.setItem('fcms_master_doc_cache', JSON.stringify(documents));
          close();
          renderUI();
        } catch (err) {
          console.error('Error saving document:', err);
          toastError('Gagal menyimpan dokumen: ' + err.message);
        }
      }
    });
  }

  // Delete Document
  function deleteDocument(doc) {
    confirmDialog(
      `Apakah Anda yakin ingin menghapus dokumen <strong>"${doc.name}"</strong>?`,
      async () => {
        try {
          const res = await apiFetch(`/api/checklist/${doc.id}`, { method: 'DELETE' });
          if (!res.ok) console.warn('API delete warning:', res);
          documents = documents.filter(d => d.id !== doc.id);
          localStorage.setItem('fcms_master_doc_cache', JSON.stringify(documents));
          toastSuccess(`Dokumen "${doc.name}" berhasil dihapus`);
          renderUI();
        } catch (err) {
          documents = documents.filter(d => d.id !== doc.id);
          localStorage.setItem('fcms_master_doc_cache', JSON.stringify(documents));
          toastSuccess(`Dokumen "${doc.name}" berhasil dihapus`);
          renderUI();
        }
      },
      'Hapus Dokumen'
    );
  }

  // Copy Link to Clipboard
  function copyLink(link, title) {
    if (!link) {
      toastInfo('Dokumen belum memiliki tautan digital.');
      return;
    }
    navigator.clipboard.writeText(link).then(() => {
      toastSuccess(`Tautan "${title}" disalin ke clipboard!`);
    }).catch(() => {
      toastInfo(`Tautan: ${link}`);
    });
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = documents.map((d, index) => ({
        'No': index + 1,
        'Nama Dokumen': d.name || '',
        'Kategori': d.category || '',
        'Platform / Format': getPlatformMeta(d.document_link).type,
        'Link Dokumen': d.document_link || '',
        'Deskripsi / Keterangan': d.description || ''
      }));
      downloadExcel(exportData, `Master_Dokumen_FCMS_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('Berhasil mengunduh data dokumen ke Excel.');
    } catch (err) {
      toastError('Gagal mengekspor data: ' + err.message);
    }
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredDocuments();
    const stats = getCategoryStats();
    const uniqueCats = Object.keys(stats).filter(k => k !== 'all');
    const linkedCount = documents.filter(d => !!d.document_link).length;

    container.innerHTML = `
      <div class="doc-page-container">
        
        <!-- Hero Header -->
        <div class="doc-hero-card">
          <div class="doc-hero-header">
            <div>
              <h1 class="doc-hero-title">
                <span>📋</span> Master Document & Template Standar
              </h1>
              <p class="doc-hero-subtitle">
                Pusat repositori formulir digital, checklist inspeksi kebersihan, berita acara, dan dokumen kerja standar PT Audy Mandiri Indonesia.
              </p>
            </div>
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <button id="doc-btn-add" class="btn btn-primary" style="background:#2563EB; border:none; border-radius:10px; padding:10px 18px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>➕</span> Tambah Dokumen Baru
              </button>
              <button id="doc-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 16px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>📥</span> Ekspor Excel
              </button>
              <button id="doc-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 14px; font-weight:600;">
                <span>🔄</span>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="doc-stats-grid">
            <div class="doc-stat-box">
              <div class="doc-stat-icon">📄</div>
              <div>
                <div class="doc-stat-value">${documents.length}</div>
                <div class="doc-stat-label">Total Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon">🗂️</div>
              <div>
                <div class="doc-stat-value">${uniqueCats.length}</div>
                <div class="doc-stat-label">Kategori Dokumen</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(16,185,129,0.18); color:#34D399;">🌐</div>
              <div>
                <div class="doc-stat-value">${linkedCount} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">/ ${documents.length}</span></div>
                <div class="doc-stat-label">Tautan Online</div>
              </div>
            </div>
            <div class="doc-stat-box">
              <div class="doc-stat-icon" style="background:rgba(168,85,247,0.18); color:#C084FC;">⚡</div>
              <div>
                <div class="doc-stat-value">Aktif</div>
                <div class="doc-stat-label">Status Repositori</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls: Search, View Switch, Category Pills -->
        <div class="doc-controls-panel">
          <div class="doc-controls-top">
            <div class="doc-search-wrapper">
              <span class="doc-search-icon">🔍</span>
              <input type="text" id="doc-search-input" class="doc-search-input" placeholder="Cari nama dokumen, kategori, atau deskripsi..." value="${searchQuery ? searchQuery.replace(/"/g, '&quot;') : ''}" />
            </div>

            <!-- View Mode Switch -->
            <div class="doc-view-toggle">
              <button id="doc-view-grid" class="doc-view-btn ${currentView === 'grid' ? 'active' : ''}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
                <span>Grid Kartu</span>
              </button>
              <button id="doc-view-table" class="doc-view-btn ${currentView === 'table' ? 'active' : ''}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel Data</span>
              </button>
            </div>
          </div>

          <!-- Category Pills -->
          <div class="doc-filter-pills">
            <button class="doc-pill ${activeCategory === 'all' ? 'active' : ''}" data-cat="all">
              <span>Semua Dokumen</span>
              <span class="doc-pill-count">${stats.all || 0}</span>
            </button>
            ${uniqueCats.map(cat => {
              const meta = getCategoryMeta(cat);
              const isActive = activeCategory === cat;
              return `
                <button class="doc-pill ${isActive ? 'active' : ''}" data-cat="${cat}">
                  <span>${meta.icon}</span>
                  <span>${cat}</span>
                  <span class="doc-pill-count">${stats[cat] || 0}</span>
                </button>
              `;
            }).join('')}
          </div>
        </div>

        <!-- Documents View Container -->
        <div id="doc-list-view">
          ${filtered.length === 0 ? `
            <div class="doc-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Dokumen Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Coba sesuaikan kata kunci pencarian atau pilih kategori lain.</p>
              <button id="doc-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Tampilkan Semua Dokumen
              </button>
            </div>
          ` : currentView === 'grid' ? `
            <!-- Grid Card View -->
            <div class="doc-cards-grid">
              ${filtered.map(doc => {
                const catMeta = getCategoryMeta(doc.category);
                const platMeta = getPlatformMeta(doc.document_link);
                return `
                  <div class="doc-card" data-id="${doc.id}">
                    <div>
                      <div class="doc-card-top">
                        <span class="doc-category-badge" style="background:${catMeta.bg}; color:${catMeta.color}; border:1px solid ${catMeta.border};">
                          <span>${catMeta.icon}</span>
                          <span>${catMeta.label}</span>
                        </span>
                        <span class="doc-platform-badge">
                          <span>${platMeta.icon}</span>
                          <span>${platMeta.type}</span>
                        </span>
                      </div>
                      
                      <h3 class="doc-card-title">${doc.name}</h3>
                      <p class="doc-card-desc">${doc.description || 'Formulir acuan standar operasional dan administrasi fasilitas.'}</p>
                    </div>

                    <div class="doc-card-footer">
                      ${doc.document_link ? `
                        <a href="${doc.document_link}" target="_blank" rel="noopener noreferrer" class="doc-btn-open">
                          <span>Buka Dokumen</span>
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                        </a>
                      ` : `
                        <button class="doc-btn-open disabled" title="Belum ada tautan yang dikaitkan">
                          <span>Tidak Ada Link</span>
                        </button>
                      `}

                      <div class="doc-card-actions">
                        ${doc.document_link ? `
                          <button class="doc-icon-btn copy-link-btn" title="Salin Tautan Dokumen" data-link="${doc.document_link}" data-title="${doc.name}">
                            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                          </button>
                        ` : ''}
                        <button class="doc-icon-btn edit-doc-btn" title="Edit Dokumen" data-id="${doc.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="doc-icon-btn delete delete-doc-btn" title="Hapus Dokumen" data-id="${doc.id}">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                      </div>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          ` : `
            <!-- Table View -->
            <div class="doc-table-card">
              <table class="doc-table">
                <thead>
                  <tr>
                    <th style="width:50px; text-align:center;">No</th>
                    <th>Nama Dokumen</th>
                    <th>Kategori</th>
                    <th>Format</th>
                    <th>Tautan Dokumen</th>
                    <th>Deskripsi</th>
                    <th style="width:100px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${filtered.map((doc, idx) => {
                    const catMeta = getCategoryMeta(doc.category);
                    const platMeta = getPlatformMeta(doc.document_link);
                    return `
                      <tr>
                        <td style="text-align:center; font-weight:600; color:#94A3B8;">${idx + 1}</td>
                        <td>
                          <div style="font-weight:700; color:#0F172A;">${doc.name}</div>
                        </td>
                        <td>
                          <span class="doc-category-badge" style="background:${catMeta.bg}; color:${catMeta.color}; border:1px solid ${catMeta.border};">
                            <span>${catMeta.icon}</span>
                            <span>${catMeta.label}</span>
                          </span>
                        </td>
                        <td>
                          <span class="doc-platform-badge">
                            <span>${platMeta.icon}</span>
                            <span>${platMeta.type}</span>
                          </span>
                        </td>
                        <td>
                          ${doc.document_link ? `
                            <div style="display:flex; align-items:center; gap:6px;">
                              <a href="${doc.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-primary" style="font-size:0.78rem; padding:5px 10px; border-radius:6px; display:inline-flex; align-items:center; gap:4px; text-decoration:none;">
                                <span>Buka Dokumen</span>
                                <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                              </a>
                              <button class="doc-icon-btn copy-link-btn" title="Salin Link" style="width:28px; height:28px;" data-link="${doc.document_link}" data-title="${doc.name}">
                                <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>
                              </button>
                            </div>
                          ` : '<span style="color:#94A3B8; font-size:0.82rem;">-</span>'}
                        </td>
                        <td style="max-width:280px; color:#64748B; font-size:0.84rem;">
                          ${doc.description || '-'}
                        </td>
                        <td style="text-align:center;">
                          <div style="display:flex; justify-content:center; gap:6px;">
                            <button class="doc-icon-btn edit-doc-btn" style="width:30px; height:30px;" title="Edit" data-id="${doc.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                            <button class="doc-icon-btn delete delete-doc-btn" style="width:30px; height:30px;" title="Hapus" data-id="${doc.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          </div>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          `}
        </div>

      </div>
    `;

    // Event Listeners Binding
    bindEvents();
  }

  // Bind Listeners
  function bindEvents() {
    // Top Buttons
    const btnAdd = container.querySelector('#doc-btn-add');
    if (btnAdd) btnAdd.onclick = () => openDocumentModal();

    const btnExcel = container.querySelector('#doc-btn-excel');
    if (btnExcel) btnExcel.onclick = () => exportToExcel();

    const btnRefresh = container.querySelector('#doc-btn-refresh');
    if (btnRefresh) btnRefresh.onclick = () => loadDocuments();

    const btnResetFilter = container.querySelector('#doc-btn-reset-filter');
    if (btnResetFilter) {
      btnResetFilter.onclick = () => {
        searchQuery = '';
        activeCategory = 'all';
        renderUI();
      };
    }

    // Search Input
    const searchInput = container.querySelector('#doc-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        const filtered = getFilteredDocuments();
        renderUI();
        const newSearch = container.querySelector('#doc-search-input');
        if (newSearch) {
          newSearch.focus();
          newSearch.setSelectionRange(newSearch.value.length, newSearch.value.length);
        }
      };
    }

    // View Toggle
    const btnGrid = container.querySelector('#doc-view-grid');
    const btnTable = container.querySelector('#doc-view-table');
    if (btnGrid) {
      btnGrid.onclick = () => {
        currentView = 'grid';
        localStorage.setItem('fcms_doc_view', 'grid');
        renderUI();
      };
    }
    if (btnTable) {
      btnTable.onclick = () => {
        currentView = 'table';
        localStorage.setItem('fcms_doc_view', 'table');
        renderUI();
      };
    }

    // Category Pills
    container.querySelectorAll('.doc-pill').forEach(pill => {
      pill.onclick = () => {
        activeCategory = pill.getAttribute('data-cat') || 'all';
        renderUI();
      };
    });

    // Copy Link Buttons
    container.querySelectorAll('.copy-link-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const link = btn.getAttribute('data-link');
        const title = btn.getAttribute('data-title');
        copyLink(link, title);
      };
    });

    // Edit Buttons
    container.querySelectorAll('.edit-doc-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const doc = documents.find(d => d.id === id);
        if (doc) openDocumentModal(doc);
      };
    });

    // Delete Buttons
    container.querySelectorAll('.delete-doc-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const doc = documents.find(d => d.id === id);
        if (doc) deleteDocument(doc);
      };
    });
  }

  // Initial Load
  await loadDocuments();
}

