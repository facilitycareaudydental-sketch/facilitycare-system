// Badge/chip helpers
export function statusBadge(status) {
  const map = {
    'Done': 'badge-success',
    'Aktif': 'badge-success',
    'Open': 'badge-warning',
    'In Progress': 'badge-info',
    'Pending': 'badge-warning',
    'Diproses': 'badge-info',
    'Selesai': 'badge-success',
    'Tidak Aktif': 'badge-neutral',
    'Resign': 'badge-neutral',
    'Cut': 'badge-danger',
    'Tidak Datang': 'badge-danger',
  };
  if (!status || status === '-' || String(status).trim() === '') return '';
  const cls = map[status] || 'badge-neutral';
  return `<span class="badge ${cls}">${status}</span>`;
}

export function daysRemainingBadge(days) {
  if (days === null || days === undefined) return '<span class="badge badge-neutral">-</span>';
  if (days < 0) return `<span class="badge badge-danger">Expired (${Math.abs(days)}h)</span>`;
  if (days <= 14) return `<span class="badge badge-danger">${days} hari</span>`;
  if (days <= 30) return `<span class="badge badge-warning">${days} hari</span>`;
  return `<span class="badge badge-success">${days} hari</span>`;
}

export function divisionBadge(div) {
  const map = { 'FACILITY CARE': 'badge-info', 'SECURITY': 'badge-secondary' };
  return `<span class="badge ${map[div] || 'badge-neutral'}">${div || '-'}</span>`;
}

export function activityTypeBadge(type) {
  const map = {
    'Inspeksi Hygiene & Aset Bangunan': 'badge-info',
    'General Cleaning': 'badge-success',
    'Deep Cleaning': 'badge-purple',
    'Fogging': 'badge-warning',
  };
  return `<span class="badge ${map[type] || 'badge-neutral'}">${type || '-'}</span>`;
}

export function periodBadge(period) {
  const map = { Q1: 'badge-info', Q2: 'badge-success', Q3: 'badge-warning', Q4: 'badge-danger' };
  return `<span class="badge ${map[period] || 'badge-neutral'}">${period || '-'}</span>`;
}
