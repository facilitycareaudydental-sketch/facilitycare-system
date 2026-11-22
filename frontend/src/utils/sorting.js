// Standardized client-side chronological date parsing & sorting utilities
// Matches backend sorting rules: chronological DESC, NULL/empty/invalid last,
// tie-breaker: updated_at DESC, created_at DESC, id DESC.

const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/;
const DMY_DATE_RE = /^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/;
const YM_DATE_RE = /^(\d{4})-(\d{2})$/;
const EXCEL_SERIAL_RE = /^\d{5}$/;

function validUtcDateMs(year, month, day) {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const dt = new Date(Date.UTC(year, month - 1, day));
  if (
    dt.getUTCFullYear() !== year ||
    dt.getUTCMonth() !== month - 1 ||
    dt.getUTCDate() !== day
  ) {
    return null;
  }
  return dt.getTime();
}

/**
 * Parses heterogeneous date strings into UTC midnight milliseconds, or null if empty/invalid.
 * Supports: YYYY-MM-DD, YYYY-MM-DDTHH:MM:SS, DD-MM-YYYY, DD/MM/YYYY, YYYY-MM, Excel 5-digit serial.
 */
export function parseSortableDateMs(value, { allowYearMonth = true } = {}) {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  if (!s || s === '-') return null;

  let m = s.match(ISO_DATE_RE);
  if (m) {
    return validUtcDateMs(Number(m[1]), Number(m[2]), Number(m[3]));
  }

  m = s.match(DMY_DATE_RE);
  if (m) {
    return validUtcDateMs(Number(m[3]), Number(m[2]), Number(m[1]));
  }

  if (EXCEL_SERIAL_RE.test(s)) {
    const serial = Number(s);
    if (serial >= 10000) {
      return Date.UTC(1899, 11, 30) + serial * 86400000;
    }
    return null;
  }

  if (allowYearMonth) {
    m = s.match(YM_DATE_RE);
    if (m) {
      return validUtcDateMs(Number(m[1]), Number(m[2]), 1);
    }
  }

  return null;
}

/**
 * Parses ISO/SQLite timestamp strings into UTC epoch ms, or null if invalid.
 */
export function parseSortableTimestampMs(value) {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  if (!s || s === '-') return null;
  const normalized = s.includes(' ') && !s.includes('T') ? `${s.replace(' ', 'T')}Z` : s;
  const ms = Date.parse(normalized);
  return Number.isFinite(ms) ? ms : parseSortableDateMs(s);
}

/**
 * Compares two records newest-to-oldest by `dateField`, placing empty/null/invalid dates last,
 * with standardized tie-breakers: updated_at DESC -> created_at DESC -> id DESC.
 */
export function compareRowsByDateDesc(a, b, dateField) {
  const da = parseSortableDateMs(a?.[dateField]);
  const db = parseSortableDateMs(b?.[dateField]);

  if (da === null && db !== null) return 1;
  if (da !== null && db === null) return -1;
  if (da !== null && db !== null && db !== da) return db - da;

  const ua = parseSortableTimestampMs(a?.updated_at);
  const ub = parseSortableTimestampMs(b?.updated_at);
  if (ua === null && ub !== null) return 1;
  if (ua !== null && ub === null) return -1;
  if (ua !== null && ub !== null && ub !== ua) return ub - ua;

  const ca = parseSortableTimestampMs(a?.created_at);
  const cb = parseSortableTimestampMs(b?.created_at);
  if (ca === null && cb !== null) return 1;
  if (ca !== null && cb === null) return -1;
  if (ca !== null && cb !== null && cb !== ca) return cb - ca;

  return Number(b?.id || 0) - Number(a?.id || 0);
}

/**
 * Sorts an array of YYYY-MM strings newest-to-oldest (empty/invalid last).
 */
export function sortYearMonthDesc(months) {
  return [...months].sort((a, b) => {
    const ma = parseSortableDateMs(a, { allowYearMonth: true });
    const mb = parseSortableDateMs(b, { allowYearMonth: true });
    if (ma === null && mb !== null) return 1;
    if (ma !== null && mb === null) return -1;
    if (ma !== null && mb !== null && mb !== ma) return mb - ma;
    return String(b || '').localeCompare(String(a || ''));
  });
}
