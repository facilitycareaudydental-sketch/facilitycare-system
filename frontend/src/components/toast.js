// Toast notification system
let container;

function getContainer() {
  if (!container) {
    container = document.createElement('div');
    container.id = 'toast-container';
    document.body.appendChild(container);
  }
  return container;
}

export function toast(message, type = 'info', duration = 3500) {
  const c = getContainer();
  const safeType = ['success', 'error', 'warning', 'info'].includes(type) ? type : 'info';
  const el = document.createElement('div');
  el.className = `toast toast-${safeType}`;
  const icons = { success: '✓', error: '✕', warning: '⚠', info: 'ℹ' };

  const iconSpan = document.createElement('span');
  iconSpan.className = 'toast-icon';
  iconSpan.textContent = icons[safeType] || 'ℹ';

  const msgSpan = document.createElement('span');
  msgSpan.className = 'toast-msg';
  msgSpan.textContent = message === null || message === undefined ? '' : String(message);

  el.appendChild(iconSpan);
  el.appendChild(msgSpan);
  c.appendChild(el);
  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => {
    el.classList.remove('show');
    setTimeout(() => el.remove(), 350);
  }, duration);
}

export const toastSuccess = (msg) => toast(msg, 'success');
export const toastError = (msg) => toast(msg, 'error');
export const toastWarning = (msg) => toast(msg, 'warning');
export const toastInfo = (msg) => toast(msg, 'info');
