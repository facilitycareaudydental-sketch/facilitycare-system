import Papa from 'papaparse';
import { resolveSyncMode } from './sync_mode.js';
import { resolveMaintenanceFreezeState } from './maintenance.js';
import { getManagementBranchSubquery, resolveManagementBranchIds } from './branchMaster.js';
import {
  buildAuditStatement,
  buildCreateAuditStatement,
  buildBulkHardDeleteAuditStatements,
  executeAtomicImportBatch
} from './audit.js';

export const PRODUCTION_SPREADSHEET_ID = '1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o';

/**
 * Resolves the Google Spreadsheet ID with strict environment isolation:
 * - In staging: Requires explicit env.GOOGLE_SPREADSHEET_ID; rejects Production Spreadsheet ID.
 * - In production: Uses env.GOOGLE_SPREADSHEET_ID or falls back to PRODUCTION_SPREADSHEET_ID.
 */
export function resolveSpreadsheetId(env = {}) {
  const isStaging = (env.ENVIRONMENT === 'staging') || (env.ENV === 'staging');
  const configuredId = (env.GOOGLE_SPREADSHEET_ID || '').trim();

  if (isStaging) {
    if (!configuredId) {
      const err = new Error('SYNC_CONFIG_ERROR: Staging environment requires an explicit staging GOOGLE_SPREADSHEET_ID.');
      err.code = 'SYNC_CONFIG_ERROR';
      throw err;
    }
    if (configuredId === PRODUCTION_SPREADSHEET_ID) {
      const err = new Error('SECURITY_VIOLATION: Staging environment cannot use Production Google Spreadsheet ID.');
      err.code = 'SECURITY_VIOLATION';
      throw err;
    }
    return configuredId;
  }

  return configuredId || PRODUCTION_SPREADSHEET_ID;
}

export async function fetchCSV(sheetName, spreadsheetId = PRODUCTION_SPREADSHEET_ID, fetcher = fetch) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const res = await fetcher(url);
  if (!res.ok) throw new Error(`Failed to fetch sheet ${sheetName}`);
  const text = await res.text();
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return parsed.data;
}

export async function syncGoogleSheets(env = {}, fetcher = fetch, actor = null) {
  const auditUser = actor || { id: null, username: 'system_google_sync', full_name: 'Google Sheets Sync', role: 'system' };
  try {
    const freezeState = await resolveMaintenanceFreezeState(env);
    if (freezeState.freezeActive) {
      return { success: false, blocked: true, error: 'MAINTENANCE_FREEZE_ACTIVE', source: freezeState.source };
    }

    console.log('Starting Google Sheets Sync...');

    // 1. Resolve spreadsheet ID with isolation checks & fail-closed SYNC_MODE
    const spreadsheetId = resolveSpreadsheetId(env);
    const syncState = resolveSyncMode(env);
    const isDryRun = !syncState.isLive;

    // 2. Fetch Data
    const empData = await fetchCSV('Master Karyawan', spreadsheetId, fetcher);
    const valData = await fetchCSV('Validasi', spreadsheetId, fetcher);

    // 3. Process Branches
    const branches = new Set();
    empData.forEach(r => { if (r['Cabang']) branches.add(r['Cabang'].trim()); });
    valData.forEach(r => { if (r['CABANG']) branches.add(r['CABANG'].trim()); });

    const currentBranches = (await env.DB.prepare('SELECT id, full_name FROM branches').all()).results;
    const branchesToInsert = [];
    for (const b of branches) {
       if (!b) continue;
       const existing = currentBranches.find(cb => cb.full_name === b);
       if (!existing) {
         branchesToInsert.push(b);
         if (!isDryRun) {
           const code = b.split('.')[0] || '';
           await env.DB.batch([
             env.DB.prepare('INSERT INTO branches (code, name, full_name, is_active) VALUES (?, ?, ?, 1)').bind(code, b, b),
             buildCreateAuditStatement(env, auditUser, 'branches', 'new', { code, name: b, full_name: b, is_active: 1 }, { source: 'google_sync' })
           ]);
         }
       }
    }
    const updatedBranches = (await env.DB.prepare('SELECT id, full_name FROM branches').all()).results;
    const branchMap = {};
    updatedBranches.forEach(b => branchMap[b.full_name] = b.id);

    // 4. Process Employees (Phase 5.66.19: Strictly restricted to Management 000 + FACILITY CARE/SECURITY)
    const mgmtBranchIds = await resolveManagementBranchIds(env.DB);
    const empPairStmts = [];
    let empUpdateCount = 0;
    const currentEmp = (await env.DB.prepare(
      `SELECT id, full_name, status, branch_id, division FROM employees
       WHERE deleted_at IS NULL AND division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('branch_id')}`
    ).all()).results;
    const sheetEmpNames = new Set();

    for (const row of empData) {
       const name = (row['Nama Lengkap'] || '').trim();
       if (!name) continue;
       sheetEmpNames.add(name);
       const bName = (row['Cabang'] || '').trim();
       const bId = branchMap[bName] || null;
       if (!bId || !mgmtBranchIds.has(Number(bId))) continue;
       const rowDiv = String(row['Div / Bagian'] || row['Divisi (Wajib Diisi)'] || '').trim().toUpperCase();
       if (rowDiv && rowDiv !== 'FACILITY CARE' && rowDiv !== 'SECURITY') continue;
       const existing = currentEmp.find(e => e.full_name === name);
       if (existing) {
          const isBranchChanged = String(existing.branch_id ?? '') !== String(bId ?? '');

          // Status protection: do not change non-active to 'Aktif'
          if (existing.status === 'Tidak Aktif' || existing.status === 'Resign' || existing.status === 'Cut') {
            if (isBranchChanged) {
              const afterEmp = { ...existing, branch_id: bId };
              empPairStmts.push(
                env.DB.prepare("UPDATE employees SET branch_id=?, updated_at=datetime('now') WHERE id=?").bind(bId, existing.id),
                buildAuditStatement(env, auditUser, 'UPDATE', 'employees', existing.id, existing, afterEmp, { source: 'google_sync' }, 'SUCCESS', true)
              );
              empUpdateCount++;
            }
          } else {
            const isStatusChanged = existing.status !== 'Aktif';
            if (isBranchChanged || isStatusChanged) {
              const afterEmp = { ...existing, branch_id: bId, status: 'Aktif' };
              empPairStmts.push(
                env.DB.prepare("UPDATE employees SET branch_id=?, status=?, updated_at=datetime('now') WHERE id=?").bind(bId, 'Aktif', existing.id),
                buildAuditStatement(env, auditUser, 'UPDATE', 'employees', existing.id, existing, afterEmp, { source: 'google_sync' }, 'SUCCESS', true)
              );
              empUpdateCount++;
            }
          }
       }
    }

    valData.forEach(row => {
       const name = (row['NAMA KARYAWAN'] || '').trim();
       if (name && !sheetEmpNames.has(name)) {
          sheetEmpNames.add(name);
          const existing = currentEmp.find(e => e.full_name === name);
          if (existing && existing.status !== 'Tidak Aktif' && existing.status !== 'Resign' && existing.status !== 'Cut') {
             if (existing.status !== 'Aktif') {
               const afterEmp = { ...existing, status: 'Aktif' };
               empPairStmts.push(
                 env.DB.prepare("UPDATE employees SET status=?, updated_at=datetime('now') WHERE id=?").bind('Aktif', existing.id),
                 buildAuditStatement(env, auditUser, 'UPDATE', 'employees', existing.id, existing, afterEmp, { source: 'google_sync' }, 'SUCCESS', true)
               );
               empUpdateCount++;
             }
          }
       }
    });

    // 5. Process PICs
    const picPairStmts = [];
    let picCount = 0;
    const currentPics = [];

    for (const row of valData) {
       const pic = (row['PIC'] || '').trim();
       if (pic) {
          const existing = currentPics.find(p => p.name === pic);
          if (!existing) {
             picPairStmts.push(
               env.DB.prepare('INSERT OR REPLACE INTO pic_list (name, role) VALUES (?, ?)').bind(pic, 'FC Spesialis'),
               buildAuditStatement(env, auditUser, 'CREATE', 'pic', pic, null, { name: pic, role: 'FC Spesialis' }, { source: 'google_sync' }, 'SUCCESS', true)
             );
             currentPics.push({ name: pic, role: 'FC Spesialis' });
             picCount++;
          }
       }

       const picPelapor = (row['PIC PELAPOR'] || '').trim();
       if (picPelapor) {
          const existingPelapor = currentPics.find(p => p.name === picPelapor);
          if (!existingPelapor) {
             picPairStmts.push(
               env.DB.prepare('INSERT OR REPLACE INTO pic_list (name, role) VALUES (?, ?)').bind(picPelapor, 'Pelapor'),
               buildAuditStatement(env, auditUser, 'CREATE', 'pic', picPelapor, null, { name: picPelapor, role: 'Pelapor' }, { source: 'google_sync' }, 'SUCCESS', true)
             );
             currentPics.push({ name: picPelapor, role: 'Pelapor' });
             picCount++;
          }
       }
    }

    // 6. Enforce DRY_RUN Check: Skip physical database writes
    if (isDryRun) {
      console.log(`[${syncState.mode}] Google Sheets sync dry-run/fail-closed complete (${syncState.reason || 'DRY_RUN'}). Identified: ${branchesToInsert.length} new branches, ${empUpdateCount} employee updates, ${picCount} PIC records. Zero database writes executed.`);
      return {
        success: true,
        dryRun: true,
        syncMode: syncState.mode,
        failClosed: !syncState.valid,
        ...(syncState.reason ? { reason: syncState.reason } : {}),
        message: `[DRY_RUN] Berhasil simulasi sinkronisasi Google Sheets. (Simulasi: ${empUpdateCount} Karyawan, ${picCount} PIC, ${branchesToInsert.length} Cabang Baru). Database tidak diubah.`,
        branchesToInsertCount: branchesToInsert.length,
        employeesToUpdateCount: empUpdateCount,
        picListCount: picCount
      };
    }

    // 7. Execute actual atomic writes if NOT dry run
    if (empPairStmts.length > 0) {
      await executeAtomicImportBatch(env, empPairStmts);
    }

    // Clear old pic list + audit atomically before inserting updated PIC entries
    const oldPics = (await env.DB.prepare('SELECT name, role FROM pic_list').all()).results || [];
    const picWipeGroup = [env.DB.prepare('DELETE FROM pic_list')];
    if (oldPics.length > 0) {
      picWipeGroup.push(...buildBulkHardDeleteAuditStatements(env, auditUser, 'pic', oldPics, 'name', { source: 'google_sync_wipe' }));
    }

    if (picWipeGroup.length + picPairStmts.length <= 80) {
      await env.DB.batch([...picWipeGroup, ...picPairStmts]);
    } else if (picWipeGroup.length <= 70 && picPairStmts.length > 0) {
      const headPairsCount = Math.floor((80 - picWipeGroup.length) / 2) * 2;
      await env.DB.batch([...picWipeGroup, ...picPairStmts.slice(0, headPairsCount)]);
      if (picPairStmts.length > headPairsCount) {
        await executeAtomicImportBatch(env, picPairStmts.slice(headPairsCount));
      }
    } else {
      await env.DB.batch(picWipeGroup);
      if (picPairStmts.length > 0) {
        await executeAtomicImportBatch(env, picPairStmts);
      }
    }

    console.log('Google Sheets Sync process finished successfully!');
    return {
      success: true,
      dryRun: false,
      message: `Berhasil sinkronisasi Google Sheets. (Ditambahkan/Diupdate: ${empUpdateCount} Karyawan, ${picCount} PIC). Data FCMS lama tetap aman.`
    };
  } catch (err) {
    console.error('Google Sheets Sync Error:', err);
    return {
      success: false,
      error: err.message,
      code: err.code || 'SYNC_ERROR'
    };
  }
}
