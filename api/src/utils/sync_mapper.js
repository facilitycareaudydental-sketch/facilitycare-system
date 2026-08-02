/**
 * FCMS Google Sheets Bidirectional Sync - Dictionary Mapper
 * 
 * Maps Excel Header Names to D1 Database Column Names
 */

export const SYNC_MAPPER = {
  // Phase 2: Master Karyawan
  'Master Karyawan': {
    table: 'employees',
    columns: {
      'FCMS_ID': 'id', // Primary Key
      'Nama Lengkap': 'full_name',
      'No Karyawan': 'employee_id',
      'Area Penempatan': 'branch_id', // Note: Needs branch resolution logic
      'Status Karyawan': 'employee_status',
      'No KTP': 'id_card',
      'No Telpon / HP': 'phone_number',
      'Nama Rekening': 'bank_account_name',
      'No Rekening': 'bank_account_number',
      'Tempat Lahir': 'birth_place',
      'Tanggal Lahir': 'birth_date',
      'Alamat Sesuai KTP': 'address',
      'Nama Ibu Kandung': 'mother_name',
      'Kontak Darurat (Keluarga/Kerabat)': 'emergency_contact',
      'No HP Darurat': 'emergency_phone',
      'Divisi (Wajib Diisi)': 'division',
      'Golongan Darah': 'blood_type',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version' // Will be managed internally by OCC
    }
  },
  
  // Example for Branch
  'Data Cabang': {
    table: 'branches',
    columns: {
      'FCMS_ID': 'id',
      'ID Cabang (Kode)': 'branch_code',
      'Nama Cabang': 'branch_name',
      'LAST_SYNC_SOURCE': 'last_sync_source'
    }
  }
};

/**
 * Helper to convert a payload from Google Sheets into a D1-compatible object
 */
export function mapPayloadToDB(sheetName, excelPayload) {
  const mapDefinition = SYNC_MAPPER[sheetName];
  if (!mapDefinition) throw new Error(`MappingError: No mapping defined for sheet: ${sheetName}`);
  
  const dbObject = {};
  for (const [excelCol, val] of Object.entries(excelPayload)) {
    if (!excelCol || excelCol.trim() === '') continue; // Ignore empty headers
    
    const dbCol = mapDefinition.columns[excelCol];
    if (!dbCol) {
      throw new Error(`MappingError: Unknown column header '${excelCol}' in sheet '${sheetName}'. Update SYNC_MAPPER or delete the column in Google Sheets.`);
    }
    dbObject[dbCol] = val;
  }
  return {
    table: mapDefinition.table,
    data: dbObject
  };
}

/**
 * Helper to convert a D1 database record into a Google Sheets payload
 */
export function mapDBToPayload(sheetName, dbRecord) {
  const mapDefinition = SYNC_MAPPER[sheetName];
  if (!mapDefinition) throw new Error(`No mapping defined for sheet: ${sheetName}`);
  
  const payload = {};
  for (const [excelCol, dbCol] of Object.entries(mapDefinition.columns)) {
    // We provide a fallback for null/undefined to empty string so Sheets doesn't break
    payload[excelCol] = dbRecord[dbCol] !== undefined && dbRecord[dbCol] !== null ? dbRecord[dbCol] : '';
  }
  return payload;
}
