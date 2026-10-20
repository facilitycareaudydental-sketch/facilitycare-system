// Modal component
export function createModal({ title, content, onConfirm, onCancel, confirmText = 'Simpan', cancelText = 'Batal', size = 'md', confirmClass = 'btn-primary' }) {
  const sizes = { sm: '400px', md: '560px', lg: '720px', xl: '900px' };
  const overlay = document.createElement('div');
  overlay.className = 'modal-overlay';
  overlay.innerHTML = `
    <div class="modal" style="max-width:${sizes[size] || sizes.md}">
      <div class="modal-header">
        <h3 class="modal-title">${title}</h3>
        <button class="modal-close" aria-label="Close">&times;</button>
      </div>
      <div class="modal-body">${typeof content === 'string' ? content : ''}</div>
      <div class="modal-footer">
        <button class="btn btn-ghost modal-cancel">${cancelText}</button>
        ${onConfirm ? `<button class="btn ${confirmClass} modal-confirm">${confirmText}</button>` : ''}
      </div>
    </div>
  `;

  if (content instanceof HTMLElement) {
    overlay.querySelector('.modal-body').appendChild(content);
  }

  const close = () => {
    overlay.classList.remove('show');
    setTimeout(() => overlay.remove(), 250);
  };

  overlay.querySelector('.modal-close').addEventListener('click', () => { if (onCancel) onCancel(); close(); });
  overlay.querySelector('.modal-cancel').addEventListener('click', () => { if (onCancel) onCancel(); close(); });
  if (onConfirm) {
    overlay.querySelector('.modal-confirm').addEventListener('click', () => onConfirm(overlay, close));
  }
  overlay.addEventListener('click', e => { if (e.target === overlay) { if (onCancel) onCancel(); close(); } });

  document.body.appendChild(overlay);
  requestAnimationFrame(() => overlay.classList.add('show'));
  return { overlay, close };
}

export function confirmDialog(message, onConfirm, title = 'Konfirmasi') {
  // Support both object options: confirmDialog({ title, message, onConfirm, confirmText, confirmClass, cancelText })
  // and positional arguments: confirmDialog(message, onConfirm, title)
  if (typeof message === 'object' && message !== null) {
    const opts = message;
    return createModal({
      title: opts.title || 'Konfirmasi',
      content: typeof opts.message === 'string' ? `<p>${opts.message}</p>` : opts.message,
      onConfirm: async (_, close) => {
        if (opts.onConfirm) await opts.onConfirm();
        close();
      },
      onCancel: opts.onCancel,
      confirmText: opts.confirmText || 'Ya, Lanjutkan',
      cancelText: opts.cancelText || 'Batal',
      confirmClass: opts.confirmClass || 'btn-danger',
    });
  }

  return createModal({
    title,
    content: typeof message === 'string' ? `<p>${message}</p>` : message,
    onConfirm: async (_, close) => {
      if (onConfirm) await onConfirm();
      close();
    },
    confirmText: 'Ya, Lanjutkan',
    confirmClass: 'btn-danger',
  });
}
