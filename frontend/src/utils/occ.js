import { toastSuccess, toastError } from '../components/toast.js';

/**
 * Strictly parse an OCC row_version as a positive integer (>= 1).
 * Never coerces missing, empty, boolean, or invalid values into a fallback version.
 *
 * @param {unknown} raw
 * @returns {number | null}
 */
export function parseValidRowVersion(raw) {
  if (raw === null || raw === undefined || raw === '' || typeof raw === 'boolean') {
    return null;
  }
  const num = Number(raw);
  return Number.isInteger(num) && num >= 1 ? num : null;
}

/**
 * Resolve row_version from the edit modal hidden field when mounted,
 * falling back to the record's existing row_version only if no field is present.
 *
 * @param {ParentNode | null} rootEl
 * @param {string} selector
 * @param {unknown} fallbackVersion
 * @returns {number | null}
 */
export function resolveModalRowVersion(rootEl, selector, fallbackVersion) {
  const input = rootEl && typeof rootEl.querySelector === 'function'
    ? rootEl.querySelector(selector)
    : null;
  if (input) {
    return parseValidRowVersion(input.value);
  }
  return parseValidRowVersion(fallbackVersion);
}

/**
 * Extract the incremented row_version returned by backend OCC update endpoints.
 *
 * @param {{ data?: { row_version?: unknown, data?: { row_version?: unknown } } } | null | undefined} res
 * @returns {number | null}
 */
export function extractUpdatedRowVersion(res) {
  return parseValidRowVersion(res?.data?.data?.row_version ?? res?.data?.row_version);
}

/**
 * Handle API save responses uniformly for OCC-protected modules:
 * - 2xx: sync local item row_version, toast success, close modal, reload fresh list.
 * - 409 Conflict: toast conflict with reload guidance, close stale modal, reload fresh list.
 * - Other errors: display backend error while keeping modal open for correction.
 *
 * @param {object} opts
 * @param {{ ok: boolean, status?: number, data?: { error?: string, message?: string, data?: { row_version?: unknown }, row_version?: unknown } }} opts.res
 * @param {boolean} opts.isEdit
 * @param {Record<string, any> | null} opts.item
 * @param {() => void} opts.close
 * @param {() => Promise<void>} opts.reloadData
 * @param {string} opts.editSuccessMessage
 * @param {string} opts.createSuccessMessage
 * @param {string} opts.defaultErrorMessage
 * @returns {Promise<boolean>}
 */
export async function handleOccSaveResponse({
  res,
  isEdit,
  item,
  close,
  reloadData,
  editSuccessMessage,
  createSuccessMessage,
  defaultErrorMessage,
}) {
  if (res.ok) {
    const nextVersion = extractUpdatedRowVersion(res);
    if (isEdit && nextVersion !== null && item && typeof item === 'object') {
      item.row_version = nextVersion;
    }
    toastSuccess(isEdit ? editSuccessMessage : createSuccessMessage);
    close();
    await reloadData();
    return true;
  }

  if (res.status === 409) {
    const backendMessage = res.data?.error || res.data?.message || 'Data telah diubah oleh pengguna atau sesi lain.';
    toastError(`Konflik data (HTTP 409): ${backendMessage} Silakan muat ulang / periksa data terbaru sebelum mengedit kembali.`);
    close();
    await reloadData();
    return false;
  }

  toastError(res.data?.error || res.data?.message || defaultErrorMessage);
  return false;
}
