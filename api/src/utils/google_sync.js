import Papa from 'papaparse';

const SPREADSHEET_ID = '1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o';

async function fetchCSV(sheetName) {
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch sheet ${sheetName}`);
  const text = await res.text();
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return parsed.data;
}

export async function syncGoogleSheets(env) {
  try {
    console.log('Starting Google Sheets Sync...');
    
    // 1. Fetch Data
    const empData = await fetchCSV('Master Karyawan');
    const valData = await fetchCSV('Validasi');

    // 2. Process Branches
    const branches = new Set();
    empData.forEach(r => { if (r['Cabang']) branches.add(r['Cabang'].trim()); });
    valData.forEach(r => { if (r['CABANG']) branches.add(r['CABANG'].trim()); });

    const currentBranches = (await env.DB.prepare('SELECT id, full_name FROM branches').all()).results;
    for (const b of branches) {
       if (!b) continue;
       const existing = currentBranches.find(cb => cb.full_name === b);
       if (!existing) {
         await env.DB.prepare('INSERT INTO branches (code, name, full_name, is_active) VALUES (?, ?, ?, 1)')
           .bind(b.split('.')[0] || '', b, b).run();
       }
    }
    const updatedBranches = (await env.DB.prepare('SELECT id, full_name FROM branches').all()).results;
    const branchMap = {};
    updatedBranches.forEach(b => branchMap[b.full_name] = b.id);

    // 3. Process Employees
    const empStmts = [];
    const currentEmp = (await env.DB.prepare('SELECT id, full_name FROM employees').all()).results;
    const sheetEmpNames = new Set();

    for (const row of empData) {
       const name = (row['Nama Lengkap'] || '').trim();
       if (!name) continue;
       sheetEmpNames.add(name);
       const bName = (row['Cabang'] || '').trim();
       const bId = branchMap[bName] || null;
       const existing = currentEmp.find(e => e.full_name === name);
       if (existing) {
          empStmts.push(env.DB.prepare('UPDATE employees SET branch_id=?, status=?, updated_at=datetime("now") WHERE id=?').bind(bId, 'Aktif', existing.id));
       } else {
          empStmts.push(env.DB.prepare('INSERT INTO employees (full_name, branch_id, status) VALUES (?, ?, ?)').bind(name, bId, 'Aktif'));
       }
    }

    // Add missing employees from Validasi sheet
    valData.forEach(row => {
       const name = (row['NAMA KARYAWAN'] || '').trim();
       if (name && !sheetEmpNames.has(name)) {
          sheetEmpNames.add(name);
          const existing = currentEmp.find(e => e.full_name === name);
          if (!existing) {
             empStmts.push(env.DB.prepare('INSERT INTO employees (full_name, branch_id, status) VALUES (?, ?, ?)').bind(name, branchMap[(row['CABANG']||'').trim()] || null, 'Aktif'));
          } else {
             empStmts.push(env.DB.prepare('UPDATE employees SET status=?, updated_at=datetime("now") WHERE id=?').bind('Aktif', existing.id));
          }
       }
    });

    // Deactivate employees not in sheets
    for (const ce of currentEmp) {
       if (!sheetEmpNames.has(ce.full_name)) {
          empStmts.push(env.DB.prepare('UPDATE employees SET status="Tidak Aktif" WHERE id=?').bind(ce.id));
       }
    }

    for (let i = 0; i < empStmts.length; i += 100) {
      await env.DB.batch(empStmts.slice(i, i + 100));
    }

    // 4. Process PICs
    const picStmts = [];
    picStmts.push(env.DB.prepare('DELETE FROM pic_list')); // Wipe clean before inserting
    
    for (const row of valData) {
       const bId = branchMap[(row['CABANG'] || '').trim()] || null;
       const pic = (row['PIC'] || '').trim();
       const kegiatan = (row['KEGIATAN'] || '').trim();
       if (bId && pic && kegiatan) {
          picStmts.push(env.DB.prepare('INSERT INTO pic_list (branch_id, name, pic_type) VALUES (?, ?, ?)').bind(bId, pic, kegiatan));
       }
    }
    
    for (let i = 0; i < picStmts.length; i += 100) {
      await env.DB.batch(picStmts.slice(i, i + 100));
    }

    console.log('Google Sheets Sync Complete!');
    return { success: true, message: `Synced ${sheetEmpNames.size} employees and ${picStmts.length - 1} PICs.` };
  } catch (err) {
    console.error('Google Sheets Sync Error:', err);
    return { success: false, error: err.message };
  }
}
