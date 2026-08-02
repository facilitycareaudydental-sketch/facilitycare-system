/**
 * FCMS Google Sheets Bidirectional Sync - Dictionary Mapper
 * 
 * Maps Excel Header Names to D1 Database Column Names
 */

export const SYNC_MAPPER = {
  // Phase 2: Master Karyawan
  'Master Karyawan': {
    table: 'employees',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Nama Lengkap': 'full_name',
      'No Karyawan': 'employee_id',
      'Area Penempatan': 'branch_id',
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
      'ROW_VERSION': 'row_version'
    }
  },
  
  // Phase 3: Data Cabang
  'Data Cabang': {
    table: 'branches',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'ID Cabang (Kode)': 'branch_code',
      'Nama Cabang': 'branch_name',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  // Phase 3: Master Kontrak
  'Master Kontrak': {
    table: 'contracts',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Employee ID': 'employee_id',
      'Start Date': 'start_date',
      'End Date': 'end_date',
      'Contract Type': 'contract_type',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  // Phase 3: Master SOP
  'Master SOP': {
    table: 'sop',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Title': 'title',
      'Description': 'description',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  // Phase 3: Master Checklist
  'Master Checklist': {
    table: 'master_checklist',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Name': 'name',
      'Category': 'category',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  // Phase 3: Master Form
  'Master Form': {
    table: 'master_forms',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Form Name': 'name',
      'Form Type': 'type',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  // =====================================
  // Phase 4: Modul Operasional
  // =====================================

  'Time Line': {
    table: 'activity_schedule',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Branch ID': 'branch_id',
      'Activity Type': 'activity_type',
      'Period': 'period',
      'PIC': 'pic',
      'Opening Date': 'opening_date',
      'Target Date': 'target_date',
      'Completion Date': 'completion_date',
      'Status': 'status',
      'Notes': 'notes',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Jadwal Reliefer': {
    table: 'relievers',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Branch ID': 'branch_id',
      'Original FC Name': 'original_fc_name',
      'Period': 'period',
      'Reliever Name': 'reliever_name',
      'Backup Date': 'backup_date',
      'Completion Date': 'completion_date',
      'Reason': 'reason',
      'Shift': 'shift',
      'Status': 'status',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Training': {
    table: 'training',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Training Date': 'training_date',
      'Batch': 'batch',
      'Subject': 'subject',
      'Participants': 'participants',
      'Branch ID': 'branch_id',
      'Trainer': 'trainer',
      'Score': 'score',
      'Notes': 'notes',
      'Document Link': 'document_link',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'One on One': {
    table: 'one_on_one',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Meeting Date': 'meeting_date',
      'Branch ID': 'branch_id',
      'Employee Name': 'employee_name',
      'PIC': 'pic',
      'Problem': 'problem',
      'Solution': 'solution',
      'Status': 'status',
      'Completion Date': 'completion_date',
      'Day Count': 'day_count',
      'Document Link': 'document_link',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Permasalahan': {
    table: 'issues',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Report Date': 'report_date',
      'Branch ID': 'branch_id',
      'Category': 'category',
      'Source': 'source',
      'Complaint': 'complaint',
      'Employee Name': 'employee_name',
      'FC Specialist': 'fc_specialist',
      'Solution': 'solution',
      'Status': 'status',
      'Completion Date': 'completion_date',
      'Day Count': 'day_count',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Hygiene': {
    table: 'cleaning_reports',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Branch ID': 'branch_id',
      'Activity Type': 'activity_type',
      'Period': 'period',
      'Activity Date': 'activity_date',
      'Status': 'status',
      'Document Link': 'document_link',
      'Notes': 'notes',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'GC-DC': {
    table: 'inspection_reports',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Branch ID': 'branch_id',
      'Period': 'period',
      'Inspection Date': 'inspection_date',
      'Status': 'status',
      'FC Score': 'fc_score',
      'SPV Score': 'spv_score',
      'Document Link': 'document_link',
      'Notes': 'notes',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Fogging': {
    table: 'fogging_reports',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Branch ID': 'branch_id',
      'Activity Type': 'activity_type',
      'Period': 'period',
      'Activity Date': 'activity_date',
      'Status': 'status',
      'Document Link': 'document_link',
      'Notes': 'notes',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Basecamp': {
    table: 'basecamp_reports',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Branch ID': 'branch_id',
      'Problem': 'problem',
      'PIC': 'pic',
      'Info Date': 'info_date',
      'Done Date': 'done_date',
      'Status': 'status',
      'Notes': 'notes',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  },

  'Chemical': {
    table: 'supply_requests',
    enabled: true,
    direction: 'BIDIRECTIONAL',
    delete_strategy: 'SOFT',
    soft_delete_col: 'deleted_at',
    schema_version: '1.0',
    columns: {
      'FCMS_ID': 'id',
      'Submitted At': 'submitted_at',
      'Submitter Name': 'submitter_name',
      'Branch ID': 'branch_id',
      'Branch Name': 'branch_name',
      'Tools Items': 'tools_items',
      'Tools Quantity': 'tools_quantity',
      'Chemical Items': 'chemical_items',
      'Chemical Quantity': 'chemical_quantity',
      'Additional Notes': 'additional_notes',
      'Status': 'status',
      'Processed By': 'processed_by',
      'Processed At': 'processed_at',
      'LAST_SYNC_SOURCE': 'last_sync_source',
      'ROW_VERSION': 'row_version'
    }
  }
};

/**
 * Helper to convert a payload from Google Sheets into a D1-compatible object
 */
export function mapPayloadToDB(sheetName, excelPayload, payloadVersion) {
  const mapDefinition = SYNC_MAPPER[sheetName];
  if (!mapDefinition) throw new Error(`MappingError: No mapping defined for sheet: ${sheetName}`);
  
  if (!mapDefinition.enabled) {
    throw new Error(`SyncConfigError: Sync is disabled for sheet ${sheetName}`);
  }
  if (mapDefinition.direction === 'OUTBOUND_ONLY') {
    throw new Error(`SyncConfigError: Sheet ${sheetName} only allows outbound sync. Webhook rejected.`);
  }
  if (payloadVersion && mapDefinition.schema_version !== payloadVersion) {
    console.warn(`Version mismatch for ${sheetName}: expected ${mapDefinition.schema_version}, got ${payloadVersion}`);
  }
  
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
    delete_strategy: mapDefinition.delete_strategy,
    soft_delete_col: mapDefinition.soft_delete_col,
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
