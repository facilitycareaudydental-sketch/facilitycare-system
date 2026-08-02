#!/usr/bin/env node
// D1 Production Audit Script
const { execSync } = require('child_process');

const DB = 'fm-operations-db';
const tables = [
  'employees','contracts','branches','activity_schedule','issues',
  'one_on_one','training','relievers','inspection_reports','cleaning_reports',
  'fogging_reports','basecamp_reports','sp_data','mutasi_data','users',
  'sop','supply_requests','roles','sessions','audit_logs','test_audit',
  'timeline_events','rbac_roles','rbac_permissions','rbac_role_permissions',
  'master_checklist','master_forms','pic_list'
];

function query(sql) {
  try {
    const result = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "${sql.replace(/"/g, "'")}" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname }
    );
    const lines = result.split('\n').filter(l => l.trim().startsWith('[') || l.trim().startsWith('{'));
    const json = lines.join('\n');
    return JSON.parse(json);
  } catch(e) {
    return null;
  }
}

async function main() {
  console.log('='.repeat(60));
  console.log('AUDIT D1 PRODUCTION — fm-operations-db');
  console.log('='.repeat(60));
  console.log('');

  // 1. Row counts per table
  console.log('## 1. JUMLAH DATA PER TABEL');
  console.log('-'.repeat(40));
  for (const t of tables) {
    try {
      const res = execSync(
        `npx wrangler d1 execute ${DB} --remote --command "SELECT COUNT(*) as cnt FROM ${t}" --json 2>nul`,
        { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
      );
      const match = res.match(/"cnt":\s*(\d+)/);
      const cnt = match ? match[1] : '?';
      console.log(`  ${t.padEnd(30)} : ${cnt}`);
    } catch(e) {
      console.log(`  ${t.padEnd(30)} : ERROR`);
    }
  }

  console.log('');
  console.log('## 2. SAMPLE DATA USERS (tanpa password)');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT id, username, full_name, role, is_active, created_at FROM users LIMIT 20" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/\[[\s\S]*?\]/);
    if (match) {
      const data = JSON.parse(match[0]);
      if (data[0] && data[0].results) {
        data[0].results.forEach(u => {
          console.log(`  [${u.id}] ${u.username} | ${u.full_name} | role:${u.role} | active:${u.is_active}`);
        });
      }
    }
  } catch(e) { console.log('  ERROR:', e.message.slice(0,100)); }

  console.log('');
  console.log('## 3. EMPLOYEES STATUS DISTRIBUTION');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT status, COUNT(*) as cnt FROM employees GROUP BY status ORDER BY cnt DESC" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/\[[\s\S]*?\]/);
    if (match) {
      const data = JSON.parse(match[0]);
      if (data[0] && data[0].results) {
        data[0].results.forEach(r => console.log(`  ${(r.status||'(kosong)').padEnd(20)} : ${r.cnt}`));
      }
    }
  } catch(e) { console.log('  ERROR'); }

  console.log('');
  console.log('## 4. CONTRACTS STATUS & EXPIRY');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT status, COUNT(*) as cnt FROM contracts GROUP BY status" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/\[[\s\S]*?\]/);
    if (match) {
      const data = JSON.parse(match[0]);
      if (data[0] && data[0].results) {
        data[0].results.forEach(r => console.log(`  ${(r.status||'(kosong)').padEnd(20)} : ${r.cnt}`));
      }
    }
  } catch(e) { console.log('  ERROR'); }

  try {
    const res2 = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT COUNT(*) as cnt FROM contracts WHERE end_date BETWEEN date('now') AND date('now','+30 days') AND status='Aktif'" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match2 = res2.match(/"cnt":\s*(\d+)/);
    console.log(`  Kontrak habis 30 hari       : ${match2 ? match2[1] : '?'}`);
  } catch(e) {}

  console.log('');
  console.log('## 5. ISSUES STATUS');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT status, COUNT(*) as cnt FROM issues GROUP BY status" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/\[[\s\S]*?\]/);
    if (match) {
      const data = JSON.parse(match[0]);
      if (data[0] && data[0].results) {
        data[0].results.forEach(r => console.log(`  ${(r.status||'(kosong)').padEnd(20)} : ${r.cnt}`));
      }
    }
  } catch(e) { console.log('  ERROR'); }

  console.log('');
  console.log('## 6. EMPLOYEES TANPA BRANCH (orphan)');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT COUNT(*) as cnt FROM employees WHERE branch_id IS NULL" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/"cnt":\s*(\d+)/);
    console.log(`  Employees tanpa branch_id   : ${match ? match[1] : '?'}`);
  } catch(e) { console.log('  ERROR'); }

  console.log('');
  console.log('## 7. TABEL ZOMBIE / TIDAK TERPAKAI');
  console.log('-'.repeat(40));
  const zombieTables = ['roles','sessions','audit_logs','test_audit','timeline_events','rbac_roles','rbac_permissions','rbac_role_permissions'];
  for (const t of zombieTables) {
    try {
      const res = execSync(
        `npx wrangler d1 execute ${DB} --remote --command "SELECT COUNT(*) as cnt FROM ${t}" --json 2>nul`,
        { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
      );
      const match = res.match(/"cnt":\s*(\d+)/);
      const cnt = match ? match[1] : '?';
      const status = cnt === '0' ? '⚠️ KOSONG' : `${cnt} baris`;
      console.log(`  ${t.padEnd(30)} : ${status}`);
    } catch(e) {
      console.log(`  ${t.padEnd(30)} : ERROR`);
    }
  }

  console.log('');
  console.log('## 8. INDEXES YANG ADA');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT name, tbl_name FROM sqlite_master WHERE type='index' AND name NOT LIKE 'sqlite_%' ORDER BY tbl_name" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/\[[\s\S]*?\]/);
    if (match) {
      const data = JSON.parse(match[0]);
      if (data[0] && data[0].results) {
        data[0].results.forEach(r => console.log(`  [${r.tbl_name}] ${r.name}`));
      }
    }
  } catch(e) { console.log('  ERROR'); }

  console.log('');
  console.log('## 9. SESSIONS AKTIF');
  console.log('-'.repeat(40));
  try {
    const res = execSync(
      `npx wrangler d1 execute ${DB} --remote --command "SELECT COUNT(*) as cnt FROM sessions WHERE expires_at > datetime('now')" --json 2>nul`,
      { encoding: 'utf8', cwd: __dirname, timeout: 15000 }
    );
    const match = res.match(/"cnt":\s*(\d+)/);
    console.log(`  Session aktif               : ${match ? match[1] : '?'}`);
  } catch(e) { console.log('  ERROR'); }

  console.log('');
  console.log('## 10. DATABASE SIZE');
  console.log('-'.repeat(40));
  console.log('  File size (dari d1 list)    : 364,544 bytes (~356 KB)');

  console.log('');
  console.log('='.repeat(60));
  console.log('AUDIT SELESAI');
  console.log('='.repeat(60));
}

main().catch(console.error);
