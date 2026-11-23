// Form builder helper
import { escapeHtml } from '../utils/sanitize.js';

export function buildFormHTML(fields) {
  return fields.map(field => {
    if (field.type === 'hidden') return `<input type="hidden" name="${escapeHtml(field.name)}" value="${escapeHtml(field.value ?? '')}">`;
    if (field.type === 'html') return field.html || '';
    if (field.type === 'row') return `<div class="form-row">${buildFormHTML(field.fields)}</div>`;

    const required = field.required ? 'required' : '';
    const label = field.label ? `<label class="form-label">${escapeHtml(field.label)}${field.required ? ' <span class="required">*</span>' : ''}</label>` : '';
    let input = '';

    switch (field.type) {
      case 'textarea':
        input = `<textarea name="${escapeHtml(field.name)}" class="form-control" placeholder="${escapeHtml(field.placeholder || '')}" ${required} rows="${field.rows || 3}">${escapeHtml(field.value ?? '')}</textarea>`;
        break;
      case 'select':
        const opts = (field.options || []).map(o => {
          const val = typeof o === 'object' ? o.value : o;
          const lbl = typeof o === 'object' ? o.label : o;
          const selected = field.value == val ? 'selected' : '';
          return `<option value="${escapeHtml(val ?? '')}" ${selected}>${escapeHtml(lbl ?? '')}</option>`;
        }).join('');
        input = `<select name="${escapeHtml(field.name)}" class="form-control" ${required}><option value="">-- Pilih ${escapeHtml(field.label || '')} --</option>${opts}</select>`;
        break;
      case 'combobox':
        const dlId = `dl-${escapeHtml(field.name)}-${Math.random().toString(36).substring(7)}`;
        const cbOpts = (field.options || []).map(o => {
          let lbl = typeof o === 'object' ? (o.label || o.value || '') : (o || '');
          if (lbl === 'undefined' || lbl === '[object Object]' || lbl === 'null') lbl = '';
          if (!lbl) return '';
          return `<option value="${escapeHtml(lbl)}"></option>`;
        }).join('');
        let displayVal = field.value || '';
        if (field.value) {
            const found = (field.options || []).find(o => (typeof o === 'object' ? o.value : o) == field.value);
            if (found) {
               let foundLbl = typeof found === 'object' ? (found.label || found.value || '') : (found || '');
               if (foundLbl && foundLbl !== 'undefined' && foundLbl !== '[object Object]' && foundLbl !== 'null') {
                  displayVal = foundLbl;
               }
            }
        }
        input = `
          <input type="text" name="${escapeHtml(field.name)}" list="${dlId}" class="form-control" value="${escapeHtml(displayVal)}" ${required} autocomplete="off">
          <datalist id="${dlId}">${cbOpts}</datalist>
        `;
        break;
      case 'employee-combobox': {
        input = `<div class="fc-crud-emp-mount" data-emp-mount="${escapeHtml(field.name)}" data-init-name="${escapeHtml(field.value || '')}"></div>`;
        break;
      }
      case 'checkbox':
        input = `<label class="checkbox-label"><input type="checkbox" name="${escapeHtml(field.name)}" value="1" ${field.value ? 'checked' : ''}> ${escapeHtml(field.checkLabel || field.label)}</label>`;
        break;
      case 'date':
        const safeVal = (window.parseFlexibleDate && field.value) ? window.parseFlexibleDate(field.value) : (field.value || '');
        input = `<input type="date" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(safeVal)}" ${required}>`;
        break;
      case 'month':
        const safeMonth = field.value ? String(field.value).slice(0, 7) : '';
        input = `<input type="month" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(safeMonth)}" ${required}>`;
        break;
      case 'number':
        const roNum = field.readonly ? 'readonly' : '';
        input = `<input type="number" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(field.value ?? '')}" placeholder="${escapeHtml(field.placeholder || '')}" min="${escapeHtml(field.min ?? '')}" max="${escapeHtml(field.max ?? '')}" step="${escapeHtml(field.step || '1')}" ${required} ${roNum}>`;
        break;
      case 'email':
        input = `<input type="email" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(field.value ?? '')}" placeholder="${escapeHtml(field.placeholder || '')}" ${required} autocomplete="${escapeHtml(field.autocomplete || 'off')}">`;
        break;
      case 'password':
        input = `<input type="password" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(field.value ?? '')}" placeholder="${escapeHtml(field.placeholder || '')}" ${required} autocomplete="${escapeHtml(field.autocomplete || 'new-password')}">`;
        break;
      case 'url':
        input = `<input type="url" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(field.value ?? '')}" placeholder="${escapeHtml(field.placeholder || '')}" ${required}>`;
        break;
      default:
        input = `<input type="${escapeHtml(field.type || 'text')}" name="${escapeHtml(field.name)}" class="form-control" value="${escapeHtml(field.value ?? '')}" placeholder="${escapeHtml(field.placeholder || '')}" ${required} autocomplete="${escapeHtml(field.autocomplete || 'off')}">`;
    }

    const hint = field.hint ? `<div class="form-hint">${escapeHtml(field.hint)}</div>` : '';
    return `<div class="form-group ${escapeHtml(field.class || '')}">${label}${input}${hint}</div>`;
  }).join('');
}

export function getFormData(form) {
  const data = {};
  const formData = new FormData(form);
  for (const [key, val] of formData.entries()) {
    data[key] = val === '' ? null : val;
  }
  // Handle unchecked checkboxes
  form.querySelectorAll('input[type=checkbox]').forEach(cb => {
    if (!cb.checked) data[cb.name] = null;
  });
  return data;
}

export function setFormErrors(form, errors) {
  form.querySelectorAll('.form-error').forEach(el => el.remove());
  if (typeof errors === 'string') {
    const div = document.createElement('div');
    div.className = 'form-error-global';
    div.textContent = errors;
    form.prepend(div);
  }
}

export function populateForm(form, data) {
  if (!data) return;
  Object.entries(data).forEach(([key, val]) => {
    const el = form.querySelector(`[name="${key}"]`);
    if (!el) return;
    if (el.hasAttribute('list') || el.classList.contains('fc-emp-input')) return; // Do not overwrite combobox display labels with raw IDs
    if (el.type === 'checkbox') el.checked = !!val;
    else if (el.type === 'date' && val && window.parseFlexibleDate) el.value = window.parseFlexibleDate(val);
    else if (el.type === 'month' && val) el.value = String(val).slice(0, 7);
    else el.value = val !== null && val !== undefined ? val : '';
  });
}
