import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches, fetchExportRows } from '../utils/dataCache.js';
import { statusBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';
import { submitJournaledImport } from '../utils/import_client.js';
import { escapeHtml, sanitizeHtmlFragment } from '../utils/sanitize.js';

function renderInspectionMobileCard(row, handlers = {}) {
  const card = document.createElement('article');
  card.className = 'fc-mobile-card insp-mobile-card';

  const branchName = row.branch_name || '-';
  const periodHtml = row.period ? sanitizeHtmlFragment(periodBadge(row.period)) : '';

  const rawStatus = row.status !== null && row.status !== undefined ? String(row.status).trim() : '';
  let rightStatusHtml = '';
  if (rawStatus && rawStatus !== '-') {
    if (/^\d+(\.\d+)?$/.test(rawStatus)) {
      rightStatusHtml = `<span class="badge insp-total-badge" title="Nilai Total Inspeksi" aria-label="Nilai Total ${escapeHtml(rawStatus)}">${escapeHtml(rawStatus)}</span>`;
    } else {
      rightStatusHtml = sanitizeHtmlFragment(statusBadge(rawStatus));
    }
  }

  const formattedDate = window.formatDate ? (window.formatDate(row.inspection_date) || '-') : (row.inspection_date || '-');
  const hasFc = row.fc_score !== null && row.fc_score !== undefined && String(row.fc_score).trim() !== '';
  const fcNum = hasFc ? Number(row.fc_score) : NaN;
  const fcColorCls = !hasFc || Number.isNaN(fcNum)
    ? ''
    : (fcNum >= 80 ? 'insp-score-good' : fcNum >= 60 ? 'insp-score-warn' : 'insp-score-low');
  const fcDisplay = hasFc ? String(row.fc_score) : '-';

  const hasSpv = row.spv_score !== null && row.spv_score !== undefined && String(row.spv_score).trim() !== '';
  const spvDisplay = hasSpv ? String(row.spv_score) : '-';

  const notesStr = row.notes !== null && row.notes !== undefined ? String(row.notes).trim() : '';
  const hasNotes = Boolean(notesStr && notesStr !== '-');
  const rawDoc = typeof row.document_link === 'string' ? row.document_link.trim() : '';
  const safeDocHref = /^(https?:\/\/|\/)/i.test(rawDoc) ? rawDoc : '';

  const extraBlocks = [];
  if (hasNotes) {
    extraBlocks.push(`
      <div class="insp-extra-notes">
        <span class="insp-extra-label">Catatan</span>
        <div class="insp-extra-text">${escapeHtml(notesStr)}</div>
      </div>
    `);
  }
  if (safeDocHref) {
    extraBlocks.push(`
      <a href="${escapeHtml(safeDocHref)}" target="_blank" rel="noopener" class="insp-doc-link">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
        <span>Buka Dokumen</span>
      </a>
    `);
  }

  card.innerHTML = `
    <div class="insp-card-header">
      <div class="insp-card-branch">
        <span class="insp-branch-icon" aria-hidden="true">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <path d="M3 21h18"/>
            <path d="M5 21V7l8-4v18"/>
            <path d="M19 21V11l-6-4"/>
            <path d="M9 9v.01"/>
            <path d="M9 12v.01"/>
            <path d="M9 15v.01"/>
            <path d="M9 18v.01"/>
          </svg>
        </span>
        <span class="insp-branch-name">${escapeHtml(branchName)}</span>
      </div>
      ${(periodHtml || rightStatusHtml) ? `
        <div class="insp-card-header-badges">
          ${periodHtml}
          ${rightStatusHtml}
        </div>
      ` : ''}
    </div>

    <div class="insp-metrics-box">
      <div class="insp-metric-col insp-metric-col--date">
        <span class="insp-metric-label">Tanggal</span>
        <span class="insp-metric-value insp-metric-value--date">${escapeHtml(formattedDate)}</span>
      </div>
      <div class="insp-metric-col insp-metric-col--score">
        <span class="insp-metric-label">Point FC</span>
        <span class="insp-metric-value insp-metric-value--num ${fcColorCls}">${escapeHtml(fcDisplay)}</span>
      </div>
      <div class="insp-metric-col insp-metric-col--score">
        <span class="insp-metric-label">Point SPV</span>
        <span class="insp-metric-value insp-metric-value--num">${escapeHtml(spvDisplay)}</span>
      </div>
    </div>

    ${extraBlocks.length > 0 ? `<div class="insp-card-extra">${extraBlocks.join('')}</div>` : ''}
  `;

  const { onView, onEdit, onDelete, actions = [] } = handlers || {};
  if (onView || onEdit || onDelete || actions.length > 0) {
    const footer = document.createElement('div');
    footer.className = 'insp-card-footer';
    const actionsWrap = document.createElement('div');
    actionsWrap.className = 'insp-card-actions';

    if (onView) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'insp-action-btn insp-action-view fc-action-view';
      btn.setAttribute('aria-label', `Lihat detail inspeksi ${branchName}`);
      btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg><span>Detail</span>`;
      btn.addEventListener('click', () => onView(row));
      actionsWrap.appendChild(btn);
    }
    if (onEdit) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'insp-action-btn insp-action-edit fc-action-edit';
      btn.setAttribute('aria-label', `Edit inspeksi ${branchName}`);
      btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg><span>Edit</span>`;
      btn.addEventListener('click', () => onEdit(row));
      actionsWrap.appendChild(btn);
    }
    actions.forEach(action => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `insp-action-btn ${action.class || ''}`;
      btn.setAttribute('aria-label', action.label || 'Aksi');
      btn.innerHTML = sanitizeHtmlFragment(action.icon ? `${action.icon}<span>${escapeHtml(action.label || '')}</span>` : escapeHtml(action.label || 'Aksi'));
      btn.addEventListener('click', () => action.handler(row));
      actionsWrap.appendChild(btn);
    });
    if (onDelete) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'insp-action-btn insp-action-del fc-action-del';
      btn.setAttribute('aria-label', `Hapus inspeksi ${branchName}`);
      btn.innerHTML = `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg><span>Hapus</span>`;
      btn.addEventListener('click', () => onDelete(row));
      actionsWrap.appendChild(btn);
    }

    footer.appendChild(actionsWrap);
    card.appendChild(footer);
  }

  return card;
}

export async function renderInspectionReports(container) {
  const branchOptions = await getCachedBranches();
  const years = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Laporan Inspeksi Hygiene',
    icon: '🔍',
    apiPath: '/api/reports/inspection',
    enableMobileFilterSheet: true,
    itemLabel: 'Laporan Inspeksi',
    bulkDelete: false,
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'inspection_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'fc_score', label: 'Point FC', render: v => v !== null && v !== undefined ? `<strong class="${v >= 80 ? 'text-success' : v >= 60 ? 'text-warning' : 'text-danger'}">${v}</strong>` : '-' },
      { key: 'spv_score', label: 'Point SPV', render: v => v !== null && v !== undefined ? `<strong>${v}</strong>` : '-' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    mobileCardRenderer: (row, handlers) => renderInspectionMobileCard(row, handlers),
    filterFields: [
      { type: 'search', placeholder: 'Cari cabang...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'month', label: 'Bulan', options: [
          { value: '01', label: 'Jan' },
          { value: '02', label: 'Feb' },
          { value: '03', label: 'Mar' },
          { value: '04', label: 'Apr' },
          { value: '05', label: 'Mei' },
          { value: '06', label: 'Jun' },
          { value: '07', label: 'Jul' },
          { value: '08', label: 'Agu' },
          { value: '09', label: 'Sep' },
          { value: '10', label: 'Okt' },
          { value: '11', label: 'Nov' },
          { value: '12', label: 'Des' }
      ]},
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'inspection_date', label: 'Tanggal Inspeksi', type: 'date', required: true, value: data?.inspection_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done'], value: data?.status || '' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'fc_score', label: 'Point FC', type: 'number', step: '0.1', min: '0', max: '100', value: data?.fc_score },
          { name: 'spv_score', label: 'Point SPV', type: 'number', step: '0.1', min: '0', max: '100', value: data?.spv_score },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', value: data?.document_link },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'inspection_reports',
      onExport: async (filters) => {
        const rows = await fetchExportRows('/api/reports/inspection/export', filters);
        const data = rows.map(d => ({
          'Cabang': d.branch_name || '',
          'Periode': d.period || '',
          'Tanggal': d.inspection_date || '',
          'Point FC': d.fc_score !== null && d.fc_score !== undefined ? d.fc_score : '',
          'Point SPV': d.spv_score !== null && d.spv_score !== undefined ? d.spv_score : '',
          'Status': d.status || '',
          'Link Dokumen': d.document_link || ''
        }));
        downloadExcel(data, `Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`);
      },
      onTemplate: () => {
        const template = [
          { 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1', 'Tanggal': '2026-01-08', 'Point FC': 85, 'Point SPV': 90, 'Status': 'Done', 'Link Dokumen': 'https://drive.google.com/...', 'Catatan': '' }
        ];
        downloadExcel(template, 'Template_Import_Inspeksi');
      },
      onImport: async (json, importCtx = {}) => {
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').toLowerCase().trim();
          const codeMatch = s.match(/^(\d{3})/);
          if (codeMatch) {
            const byCode = branchOptions.find(r => String(r.code || '').trim() === codeMatch[1] || String(r.label || '').toLowerCase().startsWith(codeMatch[1] + '.'));
            if (byCode) return byCode.value;
          }
          const b = branchOptions.find(r => String(r.label || '').toLowerCase() === s || String(r.code || '').toLowerCase() === s || String(r.name || '').toLowerCase() === s);
          return b ? b.value : null;
        };
        
        const parseDate = (v) => {
          if (v === undefined || v === null || v === '') return '';
          if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0, 10);
          const s = String(v).trim();
          if (s === '' || s === '0') return '';
          if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
          if (/^\d{4,5}$/.test(s)) {
            const n = Number(s);
            if (n > 20000 && n < 99999) {
              const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
              return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
            }
          }
          const parts = s.split(/[\/\-\.]/);
          if (parts.length === 3) {
            const [a, b, c] = parts.map(p => p.trim());
            if (a.length === 4 && b.length <= 2 && c.length <= 2) return `${a}-${b.padStart(2, '0')}-${c.padStart(2, '0')}`;
            if (c.length === 4 && b.length <= 2 && a.length <= 2) return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`;
          }
          return s; // Fallback
        };

        const payload = json.map(row => ({
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          period: String(row['Periode'] || '').trim(),
          inspection_date: parseDate(row['Tanggal']),
          fc_score: row['Point FC'] !== undefined && row['Point FC'] !== '' ? Number(row['Point FC']) : null,
          spv_score: row['Point SPV'] !== undefined && row['Point SPV'] !== '' ? Number(row['Point SPV']) : null,
          status: String(row['Status'] || '').trim(),
          document_link: String(row['Link Dokumen'] || '').trim(),
          notes: String(row['Catatan'] || row['Keterangan'] || '').trim(),
        })).filter(row => row.branch_id && row.period && row.inspection_date);
        
        return await submitJournaledImport({
          module: 'inspection_reports',
          label: 'Laporan Inspeksi',
          endpoint: '/api/import/inspection',
          body: { rows: payload, onDuplicate: 'update' },
          idempotencyKey: importCtx?.idempotencyKey,
          recover: Boolean(importCtx?.recover),
        });
      }
    }
  });
}
