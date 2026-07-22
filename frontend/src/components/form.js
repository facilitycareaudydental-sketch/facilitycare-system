// Form builder helper
export function buildFormHTML(fields) {
  return fields.map(field => {
    if (field.type === 'hidden') return `<input type="hidden" name="${field.name}" value="${field.value || ''}">`;
    if (field.type === 'row') return `<div class="form-row">${buildFormHTML(field.fields)}</div>`;

    const required = field.required ? 'required' : '';
    const label = field.label ? `<label class="form-label">${field.label}${field.required ? ' <span class="required">*</span>' : ''}</label>` : '';
    let input = '';

    switch (field.type) {
      case 'textarea':
        input = `<textarea name="${field.name}" class="form-control" placeholder="${field.placeholder || ''}" ${required} rows="${field.rows || 3}">${field.value || ''}</textarea>`;
        break;
      case 'select':
        const opts = (field.options || []).map(o => {
          const val = typeof o === 'object' ? o.value : o;
          const lbl = typeof o === 'object' ? o.label : o;
          const selected = field.value == val ? 'selected' : '';
          return `<option value="${val}" ${selected}>${lbl}</option>`;
        }).join('');
        input = `<select name="${field.name}" class="form-control" ${required}><option value="">-- Pilih ${field.label || ''} --</option>${opts}</select>`;
        break;
      case 'combobox':
        const dlId = `dl-${field.name}-${Math.random().toString(36).substring(7)}`;
        const cbOpts = (field.options || []).map(o => {
          const val = typeof o === 'object' ? o.value : o;
          const lbl = typeof o === 'object' ? o.label : o;
          return `<option value="${lbl}"></option>`;
        }).join('');
        let displayVal = field.value || '';
        if (field.value) {
            const found = (field.options || []).find(o => (typeof o === 'object' ? o.value : o) == field.value);
            if (found) displayVal = typeof found === 'object' ? found.label : found;
        }
        input = `
          <input type="text" name="${field.name}" list="${dlId}" class="form-control" value="${displayVal}" placeholder="Pilih atau ketik baru..." ${required} autocomplete="off">
          <datalist id="${dlId}">${cbOpts}</datalist>
        `;
        break;
      case 'checkbox':
        input = `<label class="checkbox-label"><input type="checkbox" name="${field.name}" value="1" ${field.value ? 'checked' : ''}> ${field.checkLabel || field.label}</label>`;
        break;
      case 'date':
        input = `<input type="date" name="${field.name}" class="form-control" value="${field.value || ''}" ${required}>`;
        break;
      case 'number':
        input = `<input type="number" name="${field.name}" class="form-control" value="${field.value || ''}" placeholder="${field.placeholder || ''}" min="${field.min || ''}" max="${field.max || ''}" step="${field.step || '1'}" ${required}>`;
        break;
      case 'email':
        input = `<input type="email" name="${field.name}" class="form-control" value="${field.value || ''}" placeholder="${field.placeholder || ''}" ${required}>`;
        break;
      case 'url':
        input = `<input type="url" name="${field.name}" class="form-control" value="${field.value || ''}" placeholder="${field.placeholder || 'https://...'}" ${required}>`;
        break;
      default:
        input = `<input type="${field.type || 'text'}" name="${field.name}" class="form-control" value="${field.value || ''}" placeholder="${field.placeholder || ''}" ${required} autocomplete="off">`;
    }

    const hint = field.hint ? `<div class="form-hint">${field.hint}</div>` : '';
    return `<div class="form-group ${field.class || ''}">${label}${input}${hint}</div>`;
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
    if (el.type === 'checkbox') el.checked = !!val;
    else el.value = val !== null && val !== undefined ? val : '';
  });
}
