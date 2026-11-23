/**
 * Frontend Canonical Master Cabang Audy Dental — 70 Cabang Operasional (001–070)
 * + 2 Kode Non-Operasional Khusus: 000. Management (HQ) & A01. Aesthica (gabung 002. Kemang).
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
})));

export const SPECIAL_NON_OPERATIONAL_BRANCHES = Object.freeze([
  Object.freeze({
    code: '000',
    name: 'Management',
    full_name: '000. Management',
    is_operational: false,
    branch_scope: 'hq',
    parent_branch_code: null,
  }),
  Object.freeze({
    code: 'A01',
    name: 'Aesthica',
    full_name: 'A01. Aesthica',
    is_operational: false,
    branch_scope: 'joined_kemang',
    parent_branch_code: '002',
  }),
]);

export const OFFICIAL_MASTER_BRANCH_MAP = new Map(
  [...OFFICIAL_OPERATIONAL_BRANCHES, ...SPECIAL_NON_OPERATIONAL_BRANCHES].map(entry => [entry.code, entry])
);

const LEGACY_MISMATCHED_NAMES = {
  '069': 'Bekasi Kota Bintang',
  '070': 'J7',
};

export function normalizeBranchCode(code) {
  return String(code ?? '').trim();
}

export function isOperationalBranchCode(code) {
  const c = normalizeBranchCode(code);
  return /^(00[1-9]|0[1-6][0-9]|070)$/.test(c);
}

export function isSpecialBranchCode(code) {
  const c = normalizeBranchCode(code);
  return c === '000' || c === 'A01';
}

export function isOfficialMasterBranchCode(code) {
  return OFFICIAL_MASTER_BRANCH_MAP.has(normalizeBranchCode(code));
}

export function isSyntheticTestBranch(row) {
  const code = normalizeBranchCode(row?.code);
  const full = String(row?.full_name || row?.name || '').trim();
  return /^T\d+[A-Z0-9-]*$/i.test(code) || /^TEST_/i.test(full) || /^U558D/i.test(code) || /\[UAT-/i.test(full);
}

/**
 * Formats a branch label for dropdown options, displaying canonical 069/070 labels
 * if a legacy DB row still carries an un-reconciled placeholder name, and tagging 000/A01.
 */
export function formatBranchDropdownLabel(branch) {
  const code = normalizeBranchCode(branch?.code);
  const canonical = OFFICIAL_MASTER_BRANCH_MAP.get(code);
  const rawName = String(branch?.name || '').trim();
  const rawFull = String(branch?.full_name || rawName || `Cabang #${branch?.id}`).trim();
  const mismatch = LEGACY_MISMATCHED_NAMES[code];

  const baseLabel = (canonical && mismatch && rawName.toLowerCase() === mismatch.toLowerCase())
    ? canonical.full_name
    : rawFull;

  return baseLabel;
}

/**
 * Normalizes and sorts raw branch rows returned by /api/branches?all=1 into clean dropdown options.
 */
export function normalizeBranchDropdownOptions(rawRows = []) {
  const rows = Array.isArray(rawRows) ? rawRows.filter(b => !b.deleted_at && !isSyntheticTestBranch(b)) : [];
  const officialCount = rows.filter(b => isOfficialMasterBranchCode(b.code)).length;
  const filtered = officialCount >= 65
    ? rows.filter(b => isOfficialMasterBranchCode(b.code))
    : rows;

  const rank = (code) => {
    const c = normalizeBranchCode(code);
    if (c === '000') return 0;
    if (isOperationalBranchCode(c)) return parseInt(c, 10);
    if (c === 'A01') return 900;
    return 1000;
  };

  return [...filtered]
    .sort((a, b) => rank(a.code) - rank(b.code) || normalizeBranchCode(a.code).localeCompare(normalizeBranchCode(b.code)) || ((a.id || 0) - (b.id || 0)))
    .map(b => ({
      value: b.id,
      label: formatBranchDropdownLabel(b),
      code: normalizeBranchCode(b.code),
      is_operational: isOperationalBranchCode(b.code),
      branch_scope: b.code === '000' ? 'hq' : b.code === 'A01' ? 'joined_kemang' : isOperationalBranchCode(b.code) ? 'operational' : 'other',
    }));
}
