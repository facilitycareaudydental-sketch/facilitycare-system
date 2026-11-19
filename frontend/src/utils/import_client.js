/**
 * import_client.js — Resumable, Idempotent Frontend Import Client (Phase 5.66.5)
 * ----------------------------------------------------------------------------
 * Manages:
 *  - Deterministic payload fingerprinting + persistent Idempotency-Key reuse across
 *    retries, network drops, and browser refreshes.
 *  - Double-click & in-flight concurrency protection.
 *  - Explicit recovery header (`X-Import-Recover: true`) ONLY when triggered for
 *    stale `IN_PROGRESS` recovery — never sent on initial submit.
 *  - Multi-sheet batch session recovery state across page refreshes (`#/import`).
 */
import { apiFetch } from '../config.js';

const SESSIONS_STORAGE_KEY = 'fcms_import_sessions_v1';
const BATCH_STORAGE_KEY = 'fcms_import_batch_v1';
const MAX_STORED_SESSIONS = 15;
const MAX_INLINE_BODY_CHARS = 450000;

const ACTIVE_INFLIGHT_FINGERPRINTS = new Set();

export class ImportRequestError extends Error {
  constructor(message, details = {}) {
    super(message || 'Gagal mengimport data');
    this.name = 'ImportRequestError';
    this.status = Number(details.status ?? 0);
    this.code = details.code || 'IMPORT_ERROR';
    this.importJobId = details.importJobId || null;
    this.committedChunks = Number(details.committedChunks || 0);
    this.totalChunks = Number(details.totalChunks || 0);
    this.committedRowIndex = Number(details.committedRowIndex ?? -1);
    this.recoverable = Boolean(details.recoverable);
    this.idempotencyKey = details.idempotencyKey || null;
    this.fingerprint = details.fingerprint || null;
    this.isNetworkError = Boolean(details.isNetworkError);
    this.serverStatus = details.serverStatus || null;
  }
}

export function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function canonicalizeValue(val) {
  if (val === null || val === undefined) return null;
  if (Array.isArray(val)) return val.map(canonicalizeValue);
  if (typeof val === 'object') {
    const out = {};
    for (const k of Object.keys(val).sort()) {
      out[k] = canonicalizeValue(val[k]);
    }
    return out;
  }
  return val;
}

function fnv1aHex(str) {
  let h1 = 0x811c9dc5;
  let h2 = 0x9e3779b9;
  for (let i = 0; i < str.length; i++) {
    const ch = str.charCodeAt(i);
    h1 ^= ch;
    h1 = Math.imul(h1, 0x01000193) >>> 0;
    h2 ^= ch + (i & 0xff);
    h2 = Math.imul(h2, 0x85ebca6b) >>> 0;
  }
  return `${h1.toString(16).padStart(8, '0')}${h2.toString(16).padStart(8, '0')}`;
}

export async function computeImportFingerprint(endpoint, body) {
  const canonical = JSON.stringify({
    endpoint: String(endpoint || ''),
    body: canonicalizeValue(body),
  });
  try {
    const subtle = globalThis?.crypto?.subtle;
    if (subtle && typeof TextEncoder !== 'undefined') {
      const encoded = new TextEncoder().encode(canonical);
      const digest = await subtle.digest('SHA-256', encoded);
      return Array.from(new Uint8Array(digest))
        .map((b) => b.toString(16).padStart(2, '0'))
        .join('');
    }
  } catch {
    // Fallback to deterministic FNV hash below when WebCrypto subtle is unavailable
  }
  return `fnv_${fnv1aHex(canonical)}_${canonical.length}`;
}

function generateIdempotencyKey(moduleName) {
  const prefix = String(moduleName || 'import').replace(/[^a-z0-9_-]/gi, '').slice(0, 20) || 'import';
  try {
    if (globalThis?.crypto?.randomUUID) {
      return `imp_${prefix}_${globalThis.crypto.randomUUID()}`;
    }
  } catch { /* fallback below */ }
  return `imp_${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

function readStorageJson(key, fallback) {
  try {
    if (typeof localStorage === 'undefined') return fallback;
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function writeStorageJson(key, value) {
  try {
    if (typeof localStorage === 'undefined') return false;
    localStorage.setItem(key, JSON.stringify(value));
    return true;
  } catch {
    return false;
  }
}

function saveSessionMap(map) {
  const entries = Object.values(map || {})
    .sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')))
    .slice(0, MAX_STORED_SESSIONS);
  const pruned = {};
  for (const entry of entries) {
    if (entry?.fingerprint) pruned[entry.fingerprint] = entry;
  }
  if (!writeStorageJson(SESSIONS_STORAGE_KEY, pruned)) {
    // Quota fallback: strip payloads from older sessions first, then current if needed
    for (const entry of entries) {
      if (entry && 'body' in entry) {
        entry.body = undefined;
        entry.payloadStored = false;
      }
    }
    writeStorageJson(SESSIONS_STORAGE_KEY, pruned);
  }
}

export function getImportSessions() {
  return readStorageJson(SESSIONS_STORAGE_KEY, {});
}

export function getImportSession(fingerprint) {
  if (!fingerprint) return null;
  const map = getImportSessions();
  return map[fingerprint] || null;
}

export function getIncompleteImportSessions(moduleFilter = null) {
  const map = getImportSessions();
  const list = Object.values(map).filter((s) => {
    if (!s || s.status === 'COMPLETED') return false;
    if (moduleFilter && s.module !== moduleFilter) return false;
    return true;
  });
  return list.sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
}

function upsertImportSession(session) {
  if (!session?.fingerprint) return;
  const map = getImportSessions();
  const serializedBody = session.body !== undefined ? JSON.stringify(session.body) : '';
  const canStoreBody = serializedBody.length > 0 && serializedBody.length <= MAX_INLINE_BODY_CHARS;
  map[session.fingerprint] = {
    fingerprint: session.fingerprint,
    idempotencyKey: session.idempotencyKey,
    module: session.module || 'import',
    label: session.label || session.module || 'Import',
    endpoint: session.endpoint,
    status: session.status,
    importJobId: session.importJobId || null,
    committedChunks: Number(session.committedChunks || 0),
    totalChunks: Number(session.totalChunks || 0),
    committedRowIndex: Number(session.committedRowIndex ?? -1),
    recoverable: Boolean(session.recoverable),
    errorCode: session.errorCode || null,
    errorMessage: session.errorMessage || null,
    idempotentReplay: Boolean(session.idempotentReplay),
    result: session.result || null,
    payloadStored: session.status === 'COMPLETED' ? false : canStoreBody,
    body: session.status === 'COMPLETED' ? undefined : (canStoreBody ? session.body : undefined),
    updatedAt: new Date().toISOString(),
  };
  saveSessionMap(map);
}

export function clearImportSession(fingerprint) {
  if (!fingerprint) return;
  const map = getImportSessions();
  if (map[fingerprint]) {
    delete map[fingerprint];
    saveSessionMap(map);
  }
}

export function saveBatchRecoveryState(batchState) {
  if (!batchState || typeof batchState !== 'object') return false;
  return writeStorageJson(BATCH_STORAGE_KEY, {
    ...batchState,
    updatedAt: new Date().toISOString(),
  });
}

export function getBatchRecoveryState() {
  return readStorageJson(BATCH_STORAGE_KEY, null);
}

export function clearBatchRecoveryState() {
  try {
    if (typeof localStorage !== 'undefined') localStorage.removeItem(BATCH_STORAGE_KEY);
  } catch { /* ignore */ }
}

export function isImportInFlight(fingerprint = null) {
  if (fingerprint) return ACTIVE_INFLIGHT_FINGERPRINTS.has(fingerprint);
  return ACTIVE_INFLIGHT_FINGERPRINTS.size > 0;
}

function unwrapResponseData(rawData) {
  if (rawData && typeof rawData === 'object' && !Array.isArray(rawData) && rawData.data && typeof rawData.data === 'object') {
    return rawData.data;
  }
  return rawData && typeof rawData === 'object' ? rawData : {};
}

export async function submitJournaledImport({
  module: moduleName,
  label = null,
  endpoint,
  body,
  idempotencyKey = null,
  recover = false,
  forceNewKey = false,
  configOverride = null,
  fetchImpl = null,
}) {
  const fingerprint = await computeImportFingerprint(endpoint, body);

  if (ACTIVE_INFLIGHT_FINGERPRINTS.has(fingerprint)) {
    const existing = getImportSession(fingerprint);
    throw new ImportRequestError('Proses import untuk data ini sedang berjalan. Mohon tunggu hingga selesai.', {
      status: 409,
      code: 'CLIENT_IMPORT_IN_PROGRESS',
      importJobId: existing?.importJobId || null,
      committedChunks: existing?.committedChunks || 0,
      totalChunks: existing?.totalChunks || 0,
      recoverable: false,
      idempotencyKey: existing?.idempotencyKey || idempotencyKey || null,
      fingerprint,
    });
  }

  const existing = getImportSession(fingerprint);
  let activeKey = idempotencyKey ? String(idempotencyKey).trim() : null;
  if (!activeKey && !forceNewKey && existing && existing.status !== 'COMPLETED' && existing.idempotencyKey) {
    activeKey = existing.idempotencyKey;
  }
  if (!activeKey) {
    activeKey = generateIdempotencyKey(moduleName);
  }

  ACTIVE_INFLIGHT_FINGERPRINTS.add(fingerprint);

  upsertImportSession({
    fingerprint,
    idempotencyKey: activeKey,
    module: moduleName,
    label: label || moduleName,
    endpoint,
    status: 'IN_PROGRESS',
    importJobId: existing?.importJobId || null,
    committedChunks: existing?.committedChunks || 0,
    totalChunks: existing?.totalChunks || 0,
    committedRowIndex: existing?.committedRowIndex ?? -1,
    recoverable: true,
    body,
  });

  const headers = {
    'Idempotency-Key': activeKey,
    ...(recover ? { 'X-Import-Recover': 'true' } : {}),
  };

  try {
    const res = await apiFetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
      ...(configOverride ? { configOverride } : {}),
      ...(fetchImpl ? { fetchImpl } : {}),
    });

    if (res.ok) {
      const payloadData = unwrapResponseData(res.data);
      const importJobId = payloadData.importJobId || existing?.importJobId || null;
      const idempotentReplay = Boolean(payloadData.idempotentReplay);
      upsertImportSession({
        fingerprint,
        idempotencyKey: activeKey,
        module: moduleName,
        label: label || moduleName,
        endpoint,
        status: 'COMPLETED',
        importJobId,
        committedChunks: Number(existing?.totalChunks || existing?.committedChunks || 1),
        totalChunks: Number(existing?.totalChunks || existing?.committedChunks || 1),
        committedRowIndex: existing?.committedRowIndex ?? -1,
        recoverable: false,
        idempotentReplay,
        result: payloadData,
      });
      return {
        ok: true,
        status: res.status,
        data: payloadData,
        raw: res.data,
        idempotencyKey: activeKey,
        fingerprint,
        importJobId,
        idempotentReplay,
      };
    }

    const errData = (res.data && typeof res.data === 'object') ? res.data : {};
    const status = Number(res.status || 0);
    const code = errData.code || (status === 0 ? 'NETWORK_ERROR' : status === 409 ? 'IMPORT_CONFLICT' : 'IMPORT_FAILED');
    const committedChunks = Number(errData.committedChunks ?? existing?.committedChunks ?? 0);
    const totalChunks = Number(errData.totalChunks ?? existing?.totalChunks ?? 0);
    const committedRowIndex = Number(errData.committedRowIndex ?? existing?.committedRowIndex ?? -1);
    const importJobId = errData.importJobId || existing?.importJobId || null;
    const isNetworkError = status === 0;
    const recoverable = isNetworkError
      || (status === 409 && code === 'IMPORT_IN_PROGRESS')
      || Boolean(errData.recoverable)
      || (status >= 500);

    let nextStatus = 'FAILED';
    if (isNetworkError) nextStatus = 'NETWORK_ERROR';
    else if (status === 409 && code === 'IMPORT_IN_PROGRESS') nextStatus = 'CONFLICT_IN_PROGRESS';
    else if (code === 'IMPORT_PARTIAL_FAILED' || status >= 500) nextStatus = 'PARTIAL_FAILED';

    let message = errData.error || `Gagal mengimport data (HTTP ${status})`;
    if (isNetworkError) {
      message = errData.error || 'Koneksi terputus saat mengimport data. Anda dapat mencoba ulang secara aman dengan key yang sama.';
    } else if (status === 409 && code === 'IMPORT_IN_PROGRESS') {
      message = errData.error || 'Import identik sedang berlangsung atau terkunci (HTTP 409).';
    } else if (status === 409 && code === 'IDEMPOTENCY_KEY_MISMATCH') {
      message = errData.error || 'Idempotency-Key bertabrakan dengan payload berbeda (HTTP 409).';
    } else if (code === 'IMPORT_PARTIAL_FAILED') {
      message = `${errData.error || 'Import gagal di tengah chunk'} (${committedChunks}/${totalChunks || '?'} chunk tersimpan).`;
    }

    upsertImportSession({
      fingerprint,
      idempotencyKey: activeKey,
      module: moduleName,
      label: label || moduleName,
      endpoint,
      status: nextStatus,
      importJobId,
      committedChunks,
      totalChunks,
      committedRowIndex,
      recoverable,
      errorCode: code,
      errorMessage: message,
      body,
    });

    throw new ImportRequestError(message, {
      status,
      code,
      importJobId,
      committedChunks,
      totalChunks,
      committedRowIndex,
      recoverable,
      idempotencyKey: activeKey,
      fingerprint,
      isNetworkError,
      serverStatus: errData.status || null,
    });
  } finally {
    ACTIVE_INFLIGHT_FINGERPRINTS.delete(fingerprint);
  }
}

export async function resumeImportSession(fingerprint, { recover = false, configOverride = null, fetchImpl = null } = {}) {
  const session = getImportSession(fingerprint);
  if (!session) {
    throw new ImportRequestError('Sesi import tidak ditemukan atau sudah dihapus.', {
      status: 404,
      code: 'IMPORT_SESSION_NOT_FOUND',
      recoverable: false,
      fingerprint,
    });
  }
  if (!session.payloadStored || session.body === undefined) {
    throw new ImportRequestError('Payload terlalu besar untuk disimpan di browser. Silakan pilih ulang file Excel yang sama — Idempotency-Key akan tetap dipertahankan.', {
      status: 400,
      code: 'IMPORT_PAYLOAD_REUPLOAD_REQUIRED',
      recoverable: false,
      idempotencyKey: session.idempotencyKey,
      fingerprint,
    });
  }
  return await submitJournaledImport({
    module: session.module,
    label: session.label,
    endpoint: session.endpoint,
    body: session.body,
    idempotencyKey: session.idempotencyKey,
    recover,
    configOverride,
    fetchImpl,
  });
}

export function renderImportStatusBannerHTML(state) {
  if (!state) return '';
  const {
    variant = 'info', // 'progress' | 'success' | 'warning' | 'error' | 'recover'
    title = '',
    message = '',
    idempotencyKey = '',
    importJobId = '',
    committedChunks = 0,
    totalChunks = 0,
    errorCode = '',
    fingerprint = '',
    canRetry = false,
    canRecover = false,
    canReplay = false,
    canDismiss = true,
  } = state;

  const palette = {
    progress: { bg: '#EFF6FF', border: '#93C5FD', color: '#1E3A8A' },
    success:  { bg: '#ECFDF5', border: '#6EE7B7', color: '#065F46' },
    warning:  { bg: '#FFFBEB', border: '#FCD34D', color: '#92400E' },
    recover:  { bg: '#FFF7ED', border: '#FDBA74', color: '#9A3412' },
    error:    { bg: '#FEF2F2', border: '#FCA5A5', color: '#991B1B' },
    info:     { bg: '#F8FAFC', border: '#CBD5E1', color: '#334155' },
  }[variant] || { bg: '#F8FAFC', border: '#CBD5E1', color: '#334155' };

  const chunkBadge = totalChunks > 0
    ? `<span class="badge badge-info" style="margin-right:6px;">Chunk ${committedChunks}/${totalChunks}</span>`
    : '';
  const codeBadge = errorCode
    ? `<span class="badge badge-danger" style="margin-right:6px;">${escapeHtml(errorCode)}</span>`
    : '';
  const keyPreview = idempotencyKey
    ? `<code title="${escapeHtml(idempotencyKey)}" style="font-size:0.75rem;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:4px;">Key: ${escapeHtml(idempotencyKey.slice(0, 26))}${idempotencyKey.length > 26 ? '…' : ''}</code>`
    : '';
  const jobPreview = importJobId
    ? `<code title="${escapeHtml(importJobId)}" style="font-size:0.75rem;background:rgba(255,255,255,0.7);padding:2px 6px;border-radius:4px;margin-left:6px;">Job: ${escapeHtml(importJobId.slice(0, 12))}…</code>`
    : '';

  return `
    <div class="import-status-banner" data-variant="${escapeHtml(variant)}" data-fingerprint="${escapeHtml(fingerprint)}" style="background:${palette.bg};border:1px solid ${palette.border};color:${palette.color};border-radius:10px;padding:12px 16px;">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;">
        <div style="flex:1;min-width:240px;">
          <div style="font-weight:700;font-size:0.95rem;margin-bottom:4px;">${escapeHtml(title)}</div>
          <div style="font-size:0.88rem;line-height:1.45;">${escapeHtml(message)}</div>
          <div style="margin-top:8px;display:flex;flex-wrap:wrap;align-items:center;gap:4px;">
            ${codeBadge}${chunkBadge}${keyPreview}${jobPreview}
          </div>
        </div>
        <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
          ${canRetry ? `<button type="button" class="btn btn-primary btn-sm" data-import-action="retry" data-fingerprint="${escapeHtml(fingerprint)}">🔄 Coba Lagi / Lanjutkan</button>` : ''}
          ${canRecover ? `<button type="button" class="btn btn-warning btn-sm" data-import-action="recover" data-fingerprint="${escapeHtml(fingerprint)}">🛠️ Pulihkan Lock (Recover)</button>` : ''}
          ${canReplay ? `<button type="button" class="btn btn-secondary btn-sm" data-import-action="replay" data-fingerprint="${escapeHtml(fingerprint)}">♻️ Replay Key Sama</button>` : ''}
          ${canDismiss ? `<button type="button" class="btn btn-ghost btn-sm" data-import-action="dismiss" data-fingerprint="${escapeHtml(fingerprint)}">✕ Tutup</button>` : ''}
        </div>
      </div>
    </div>
  `;
}
