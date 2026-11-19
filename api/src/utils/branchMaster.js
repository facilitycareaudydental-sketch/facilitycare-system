/**
 * Canonical Master Cabang Audy Dental — 70 Cabang Operasional (001–070)
 * + 2 Kode Non-Operasional Khusus:
 *   - 000. Management (HQ / Pusat)
 *   - A01. Aesthica (cabang khusus bergabung ke area 002. Kemang)
 */

export const OFFICIAL_OPERATIONAL_BRANCHES = Object.freeze([
  ['001', 'Pondok Bambu'],
  ['002', 'Kemang'],
  ['003', 'Depok'],
  ['004', 'Cipete'],
  ['005', 'Greenville'],
  ['006', 'Kelapa Gading MOI'],
  ['007', 'Bekasi'],
  ['008', 'PIK'],
  ['009', 'Bintaro'],
  ['010', 'Bogor'],
  ['011', 'Kuningan'],
  ['012', 'BSD'],
  ['013', 'Bandung Gatsu'],
  ['014', 'Sunter'],
  ['015', 'Gading Serpong'],
  ['016', 'Cibubur'],
  ['017', 'Bandung Setrasari'],
  ['018', 'Harapan Indah'],
  ['019', 'Cinere'],
  ['020', 'Surabaya'],
  ['021', 'Kelapa Gading Boulevard'],
  ['022', 'Cikarang'],
  ['023', 'Semarang'],
  ['024', 'Grand Wisata'],
  ['025', 'Bogor Baru'],
  ['026', 'Tangerang City'],
  ['027', 'Rawamangun'],
  ['028', 'Taman Palem'],
  ['029', 'Karawang'],
  ['030', 'Ciputat'],
  ['031', 'Sawangan'],
  ['032', 'Cawang'],
  ['033', 'Taman Mini'],
  ['034', 'Puri Indah'],
  ['035', 'Ciledug'],
  ['036', 'Jagakarsa'],
  ['037', 'Bali'],
  ['038', 'Greenlake'],
  ['039', 'Bandung Mekarwangi'],
  ['040', 'Bekasi Timur'],
  ['041', 'Karawaci'],
  ['042', 'Muara Karang'],
  ['043', 'Tebet'],
  ['044', 'Depok GDC'],
  ['045', 'FX Mall'],
  ['046', 'Kaliurang Jogja'],
  ['047', 'Bogor Tajur'],
  ['048', 'PIK 2'],
  ['049', 'Surabaya Merr'],
  ['050', 'Cibinong'],
  ['051', 'Bintaro Sektor 9'],
  ['052', 'BSD Rawa Buntu'],
  ['053', 'Alam Sutera'],
  ['054', 'Bogor Gunung Batu'],
  ['055', 'Gajah Mada'],
  ['056', 'Citra Garden 2'],
  ['057', 'Cikupa'],
  ['058', 'Solo Baru'],
  ['059', 'Malang Suhat'],
  ['060', 'Palmerah'],
  ['061', 'Cirebon'],
  ['062', 'Summarecon Bekasi'],
  ['063', 'Pasar Minggu'],
  ['064', 'Pondok Indah'],
  ['065', 'Kebon Sirih'],
  ['066', 'Bandung KBP'],
  ['067', 'Agora Mall'],
  ['068', 'Kota Wisata'],
  ['069', 'Depok Pancoran Mas'],
  ['070', 'Bekasi Kota Bintang'],
].map(([code, name]) => Object.freeze({
  code,
  name,
  full_name: `${code}. ${name}`,
  is_operational: true,
  branch_scope: 'operational',
  parent_branch_code: null,
  scope_label: 'Cabang Operasional',
})));

export const SPECIAL_NON_OPERATIONAL_BRANCHES = Object.freeze([
  Object.freeze({
    code: '000',
    name: 'Management',
    full_name: '000. Management',
    is_operational: false,
    branch_scope: 'hq',
    parent_branch_code: null,
    scope_label: 'HQ / Management (Non-Operasional)',
  }),
  Object.freeze({
    code: 'A01',
    name: 'Aesthica',
    full_name: 'A01. Aesthica',
    is_operational: false,
    branch_scope: 'joined_kemang',
    parent_branch_code: '002',
    scope_label: 'Khusus — Gabung Area 002. Kemang',
  }),
]);

export const OFFICIAL_MANAGEMENT_BRANCH = SPECIAL_NON_OPERATIONAL_BRANCHES[0];

export const OFFICIAL_MASTER_BRANCH_MAP = new Map(
  [...OFFICIAL_OPERATIONAL_BRANCHES, ...SPECIAL_NON_OPERATIONAL_BRANCHES].map(entry => [entry.code, entry])
);

const LEGACY_MISMATCHED_NAMES = {
  '069': 'Bekasi Kota Bintang',
  '070': 'J7',
};

/**
 * Normalizes a branch code string (trimmed).
 */
export function normalizeBranchCode(code) {
  return String(code ?? '').trim();
}

/**
 * Returns true strictly for operational Audy Dental branch codes 001..070.
 */
export function isOperationalBranchCode(code) {
  const c = normalizeBranchCode(code);
  return /^(00[1-9]|0[1-6][0-9]|070)$/.test(c);
}

/**
 * Returns true for special non-operational codes (000. Management HQ, A01. Aesthica).
 */
export function isSpecialBranchCode(code) {
  const c = normalizeBranchCode(code);
  return c === '000' || c === 'A01';
}

/**
 * Returns true for any code in the official 70 + 2 (000, A01) master.
 */
export function isOfficialMasterBranchCode(code) {
  return OFFICIAL_MASTER_BRANCH_MAP.has(normalizeBranchCode(code));
}

/**
 * Returns true for synthetic/automated test branch codes that must not pollute UI selectors.
 */
export function isSyntheticTestBranch(row) {
  const code = normalizeBranchCode(row?.code);
  const full = String(row?.full_name || row?.name || '').trim();
  return /^T\d+[A-Z0-9-]*$/i.test(code) || /^TEST_/i.test(full) || /^U558D/i.test(code) || /\[UAT-/i.test(full);
}

/**
 * Returns classification metadata for any branch code.
 */
export function getBranchClassification(code) {
  const c = normalizeBranchCode(code);
  const official = OFFICIAL_MASTER_BRANCH_MAP.get(c);
  if (official) return official;
  return {
    code: c,
    name: null,
    full_name: null,
    is_operational: false,
    branch_scope: 'other',
    parent_branch_code: null,
    scope_label: 'Non-Standar / Tambahan',
  };
}

/**
 * Enriches a DB branch row with canonical scope metadata and repairs stale 069/070 display labels
 * without altering historical relational branch_id values.
 */
export function enrichBranchMasterRow(row, { normalizeLegacyMismatch = true } = {}) {
  if (!row) return row;
  const code = normalizeBranchCode(row.code);
  const meta = getBranchClassification(code);
  const legacyMismatch = LEGACY_MISMATCHED_NAMES[code];
  const isStaleMismatch = Boolean(
    legacyMismatch &&
    String(row.name || '').trim().toLowerCase() === legacyMismatch.toLowerCase()
  );

  const enriched = {
    ...row,
    is_operational: meta.is_operational,
    is_official_master: OFFICIAL_MASTER_BRANCH_MAP.has(code),
    branch_scope: meta.branch_scope,
    parent_branch_code: meta.parent_branch_code,
    scope_label: meta.scope_label,
  };

  if (normalizeLegacyMismatch && isStaleMismatch && meta.name && meta.full_name) {
    enriched.legacy_db_name = row.name;
    enriched.legacy_db_full_name = row.full_name;
    enriched.master_sync_pending = true;
    enriched.name = meta.name;
    enriched.full_name = meta.full_name;
  }

  return enriched;
}

/**
 * Extracts official code sort order: 000 first, 001..070 next, A01 after, others at end.
 */
export function compareBranchesMasterOrder(a, b) {
  const ca = normalizeBranchCode(a?.code);
  const cb = normalizeBranchCode(b?.code);
  const rank = (c) => {
    if (c === '000') return 0;
    if (isOperationalBranchCode(c)) return parseInt(c, 10);
    if (c === 'A01') return 900;
    return 1000;
  };
  const ra = rank(ca);
  const rb = rank(cb);
  if (ra !== rb) return ra - rb;
  return ca.localeCompare(cb) || ((a?.id || 0) - (b?.id || 0));
}

/**
 * Extracts branch code from full_name ("069. Depok Pancoran Mas" -> "069") when b.code is unavailable.
 */
export function extractBranchCodeFromLabel(label) {
  const m = String(label ?? '').trim().match(/^([A-Za-z]?\d{2,3})\s*\./);
  return m ? m[1] : '';
}

/**
 * SQL fragment predicate identifying active official operational branches (001..070).
 * Keeps the exact `is_active=1 AND deleted_at IS NULL` prefix relied on by regression monitors.
 */
export const OPERATIONAL_BRANCH_SQL_WHERE =
  "is_active=1 AND deleted_at IS NULL AND code BETWEEN '001' AND '070' AND LENGTH(TRIM(code)) = 3";

/**
 * Returns SQL fragment predicate identifying the official Management HQ branch (code '000').
 * Strictly rejects wildcard '%management%' matches so operational branches cannot accidentally match.
 * @param {string} [alias='b'] - Table alias for branches table
 * @returns {string} SQL condition string
 */
export function getManagementBranchSqlPredicate(alias = 'b') {
  const p = alias ? `${alias}.` : '';
  return `(TRIM(COALESCE(${p}code, '')) = '000' OR (LOWER(TRIM(COALESCE(${p}name, ''))) = 'management' AND (TRIM(COALESCE(${p}code, '')) = '' OR TRIM(${p}code) = '000')) OR LOWER(TRIM(COALESCE(${p}full_name, ''))) IN ('000. management', 'cabang 000. management'))`;
}

/**
 * Returns a SQL subquery string that resolves active Management branch IDs.
 * @param {string} [col='branch_id'] - Column name to match against
 * @returns {string} SQL condition string
 */
export function getManagementBranchSubquery(col = 'branch_id') {
  return `${col} IN (SELECT id FROM branches WHERE deleted_at IS NULL AND COALESCE(is_active, 1) = 1 AND ${getManagementBranchSqlPredicate('')})`;
}

/**
 * Checks whether a branch object or branch code represents the canonical Management area.
 * @param {Object|string} branchOrCode
 * @returns {boolean}
 */
export function isManagementBranch(branchOrCode) {
  if (!branchOrCode) return false;
  if (typeof branchOrCode === 'string') {
    const c = normalizeBranchCode(branchOrCode);
    const lower = c.toLowerCase();
    return c === '000' || lower === 'management' || lower === '000. management' || lower === 'cabang 000. management';
  }
  const code = normalizeBranchCode(branchOrCode.code);
  const name = String(branchOrCode.name || '').trim().toLowerCase();
  const full = String(branchOrCode.full_name || '').trim().toLowerCase();
  return (
    code === '000' ||
    (name === 'management' && (!code || code === '000')) ||
    full === '000. management' ||
    full === 'cabang 000. management'
  );
}

/**
 * Resolves the active canonical Management branch ID set fail-closed.
 * Throws if zero or multiple conflicting Management branches exist.
 * @param {Object} db
 * @returns {Promise<Set<number>>}
 */
export async function resolveManagementBranchIds(db) {
  let res;
  try {
    res = await db.prepare(
      `SELECT id, code, name, full_name FROM branches WHERE deleted_at IS NULL AND COALESCE(is_active, 1) = 1 AND ${getManagementBranchSqlPredicate('')}`
    ).all();
  } catch (dbErr) {
    console.error('resolveManagementBranchIds DB error:', dbErr?.message || dbErr);
    const err = new Error('Failed to verify Management branch (000. Management) scope');
    err.code = 'MANAGEMENT_BRANCH_DB_ERROR';
    throw err;
  }
  const rows = (res?.results || []).filter(r => isManagementBranch(r));
  const ids = new Set(rows.map(r => Number(r.id)).filter(n => Number.isInteger(n) && n > 0));
  if (ids.size !== 1) {
    const err = new Error(
      ids.size === 0
        ? 'Management branch (000. Management) not found or inactive; operation blocked fail-closed'
        : 'Ambiguous Management branch definitions detected; operation blocked fail-closed'
    );
    err.code = ids.size === 0 ? 'MANAGEMENT_BRANCH_MISSING' : 'MANAGEMENT_BRANCH_AMBIGUOUS';
    throw err;
  }
  return ids;
}
