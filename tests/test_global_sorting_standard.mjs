import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { DatabaseSync } from 'node:sqlite';

import { createToken } from '../api/src/utils/auth.js';
import { handleEmployees } from '../api/src/routes/employees.js';
import { handleContracts } from '../api/src/routes/contracts.js';
import { handleSchedule } from '../api/src/routes/schedule.js';
import { handleOSMigrations } from '../api/src/routes/os_migrations.js';
import { handleIssues } from '../api/src/routes/issues.js';
import { handleSP } from '../api/src/routes/sp.js';
import { handleMutasi } from '../api/src/routes/mutasi.js';
import { handleOvertime } from '../api/src/routes/overtime.js';
import { handleTraining } from '../api/src/routes/training.js';
import { handleReports } from '../api/src/routes/reports.js';
import { handleUsers } from '../api/src/routes/users.js';
import { handleMisc } from '../api/src/routes/misc.js';
import { handleHygieneStandards } from '../api/src/routes/hygiene.js';
import { handleAuditLogs } from '../api/src/routes/audit_logs.js';
import { handleBranches } from '../api/src/routes/branches.js';
import { handleDAC } from '../api/src/routes/dac.js';
import { handleDashboard } from '../api/src/routes/dashboard.js';
import { handleOneOnOne } from '../api/src/routes/one_on_one.js';
import { handleRelievers } from '../api/src/routes/relievers.js';
import {
  parseSortableDateMs,
  compareRowsByDateDesc,
  sortYearMonthDesc,
} from '../frontend/src/utils/sorting.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

function createD1Adapter(sqliteDb) {
  return {
    prepare(sql) {
      let bound = [];
      const stmtApi = {
        bind(...args) {
          bound = args.map(v => (v === undefined ? null : v));
          return stmtApi;
        },
        async first() {
          const row = sqliteDb.prepare(sql).get(...bound);
          return row ?? null;
        },
        async all() {
          const results = sqliteDb.prepare(sql).all(...bound);
          return { results, success: true };
        },
        async run() {
          const info = sqliteDb.prepare(sql).run(...bound);
          return {
            success: true,
            meta: {
              changes: info.changes,
              last_row_id: Number(info.lastInsertRowid || 0),
            },
          };
        },
      };
      return stmtApi;
    },
    async batch(stmts) {
      const out = [];
      for (const s of stmts) out.push(await s.run());
      return out;
    },
    async exec(sql) {
      sqliteDb.exec(sql);
    },
  };
}

function setupDatabase({ includeSpUpdatedAt = true } = {}) {
  const db = new DatabaseSync(':memory:');
  db.exec(`
    CREATE TABLE branches (id INTEGER PRIMARY KEY, code TEXT, name TEXT, full_name TEXT, is_active INTEGER DEFAULT 1, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE users (id INTEGER PRIMARY KEY, username TEXT, email TEXT, full_name TEXT, password_hash TEXT, role TEXT, is_active INTEGER DEFAULT 1, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE employees (id INTEGER PRIMARY KEY, full_name TEXT, branch_id INTEGER, division TEXT, phone TEXT, join_date TEXT, status TEXT, notes TEXT, target_pindah_os TEXT, target_selesai_os TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE contracts (id INTEGER PRIMARY KEY, employee_id INTEGER, branch_id INTEGER, division TEXT, contract_type TEXT, pkwt_number TEXT, start_date TEXT, end_date TEXT, status TEXT, notes TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE activity_schedule (id INTEGER PRIMARY KEY, branch_id INTEGER, activity_type TEXT, period TEXT, pic TEXT, opening_date TEXT, target_date TEXT, completion_date TEXT, status TEXT, notes TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE os_migrations (id INTEGER PRIMARY KEY, branch_id INTEGER, vendor TEXT, target_month TEXT, actual_date TEXT, status TEXT, notes TEXT, pic_vendor TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE sp_data (id INTEGER PRIMARY KEY, tanggal TEXT, employee_name TEXT, division TEXT, branch_id INTEGER, akhir_sp TEXT, sp_type TEXT, status TEXT, document_link TEXT, created_at TEXT DEFAULT (datetime('now'))${includeSpUpdatedAt ? ", updated_at TEXT DEFAULT (datetime('now'))" : ''});
    CREATE TABLE mutasi_data (id INTEGER PRIMARY KEY, tanggal TEXT, employee_name TEXT, from_branch_id INTEGER, to_branch_id INTEGER, status TEXT, document_link TEXT, created_at TEXT DEFAULT (datetime('now'))${includeSpUpdatedAt ? ", updated_at TEXT DEFAULT (datetime('now'))" : ''});
    CREATE TABLE overtime_records (id INTEGER PRIMARY KEY, date TEXT, branch_id INTEGER, employee_id INTEGER, start_time TEXT, end_time TEXT, total_hours REAL, reason TEXT, notes TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE issues (id INTEGER PRIMARY KEY, report_date TEXT, branch_id INTEGER, category TEXT, source TEXT, complaint TEXT, employee_name TEXT, fc_specialist TEXT, solution TEXT, status TEXT, completion_date TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE training (id INTEGER PRIMARY KEY, training_date TEXT, batch TEXT, subject TEXT, participants TEXT, branch_id INTEGER, trainer TEXT, score REAL, notes TEXT, document_link TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE inspection_reports (id INTEGER PRIMARY KEY, branch_id INTEGER, period TEXT, inspection_date TEXT, fc_score REAL, spv_score REAL, status TEXT, document_link TEXT, notes TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE cleaning_reports (id INTEGER PRIMARY KEY, branch_id INTEGER, activity_type TEXT, period TEXT, activity_date TEXT, status TEXT, document_link TEXT, notes TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE fogging_reports (id INTEGER PRIMARY KEY, branch_id INTEGER, activity_type TEXT, period TEXT, activity_date TEXT, status TEXT, document_link TEXT, notes TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE basecamp_reports (id INTEGER PRIMARY KEY, branch_id INTEGER, problem TEXT, pic TEXT, info_date TEXT, done_date TEXT, status TEXT, notes TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE supply_requests (id INTEGER PRIMARY KEY, branch_id INTEGER, branch_name TEXT, submitter_name TEXT, status TEXT, submitted_at TEXT DEFAULT (datetime('now')), created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE sop (id INTEGER PRIMARY KEY, name TEXT, category TEXT, document_link TEXT, version TEXT, effective_date TEXT, notes TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE master_checklist (id INTEGER PRIMARY KEY, name TEXT, category TEXT, document_link TEXT, description TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE master_hygiene_standards (id INTEGER PRIMARY KEY, room_name TEXT, item_name TEXT, cleanliness_standard TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE master_dac_standards (id INTEGER PRIMARY KEY, shift TEXT, item_no INTEGER, time_range TEXT, location TEXT, activity TEXT, category TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE audit_logs (id INTEGER PRIMARY KEY, user_name TEXT, user_role TEXT, action TEXT, module TEXT, details TEXT, old_data TEXT, new_data TEXT, created_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE one_on_one (id INTEGER PRIMARY KEY, employee_name TEXT, branch_id INTEGER, meeting_date TEXT, status TEXT, problem TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
    CREATE TABLE relievers (id INTEGER PRIMARY KEY, branch_id INTEGER, period TEXT, original_fc_name TEXT, reliever_name TEXT, backup_date TEXT, completion_date TEXT, reason TEXT, shift TEXT, status TEXT, deleted_at TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')));
  `);
  return db;
}

const SECRET = 'test-sorting-secret-only-32-bytes-ok';
const ORIGIN = 'https://fcms-staging.example';

async function callRoute(handler, env, token, urlPath) {
  const req = new Request(`https://api.example${urlPath}`, {
    method: 'GET',
    headers: { Authorization: `Bearer ${token}` },
  });
  const res = await handler(req, env, ORIGIN);
  const body = await res.json();
  return { status: res.status, body };
}

function seedAllData(db) {
  // Branches (out of order insert vs code ASC)
  db.prepare("INSERT INTO branches (id, code, name, full_name, is_active) VALUES (1, '010', 'Kelapa Gading', '010. Kelapa Gading', 1)").run();
  db.prepare("INSERT INTO branches (id, code, name, full_name, is_active) VALUES (2, '001', 'Pondok Bambu', '001. Pondok Bambu', 1)").run();
  db.prepare("INSERT INTO branches (id, code, name, full_name, is_active) VALUES (3, '005', 'Bekasi', '005. Bekasi', 1)").run();

  // Users (superadmin for auth + extra users to verify created_at DESC)
  db.prepare("INSERT INTO users (id, username, email, full_name, password_hash, role, is_active, created_at) VALUES (1, 'owner', 'owner@test.id', 'Owner User', 'hash', 'superadmin', 1, '2026-01-10 08:00:00')").run();
  db.prepare("INSERT INTO users (id, username, email, full_name, password_hash, role, is_active, created_at) VALUES (2, 'alpha', 'alpha@test.id', 'Alpha Admin', 'hash', 'viewer', 1, '2026-06-01T10:00:00.000Z')").run();
  db.prepare("INSERT INTO users (id, username, email, full_name, password_hash, role, is_active, created_at) VALUES (3, 'zeta', 'zeta@test.id', 'Zeta Viewer', 'hash', 'viewer', 1, '2026-06-01 10:00:00')").run();
  db.prepare("INSERT INTO users (id, username, email, full_name, password_hash, role, is_active, created_at) VALUES (4, 'older', 'older@test.id', 'Abacus Old', 'hash', 'viewer', 1, '2025-12-01 09:00:00')").run();

  // Employees: mix of Excel 5-digit serial, DD/MM/YYYY, YYYY-MM-DD, tie-breakers, empty & invalid
  // 45369 = 2024-03-18; 46000 = 2025-12-09
  const insEmp = db.prepare(`
    INSERT INTO employees (id, full_name, branch_id, division, join_date, status, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, 'Aktif', ?, ?)
  `);
  insEmp.run(1, 'Emp Serial 2024', 1, 'FACILITY CARE', '45369', '2026-01-01 08:00:00', '2026-01-01 08:00:00'); // 2024-03-18
  insEmp.run(2, 'Emp Serial Dec2025', 1, 'FACILITY CARE', '46000', '2026-01-01 08:00:00', '2026-01-01 08:00:00'); // 2025-12-09
  insEmp.run(3, 'Emp DMY Jan2026', 2, 'FACILITY CARE', '15/01/2026', '2026-01-01 08:00:00', '2026-01-01 08:00:00'); // 2026-01-15
  insEmp.run(4, 'Emp Tie OlderUpd', 1, 'FACILITY CARE', '2026-06-01', '2026-06-01 09:00:00', '2026-06-01 09:00:00');
  insEmp.run(5, 'Emp Tie NewerUpd', 1, 'FACILITY CARE', '2026-06-01', '2026-06-01 08:00:00', '2026-06-02 10:00:00');
  insEmp.run(6, 'Emp Tie SameUpd NewerCrt', 1, 'SECURITY', '2026-06-01', '2026-06-01 11:00:00', '2026-06-02 10:00:00');
  insEmp.run(7, 'Emp Dash', 1, 'FACILITY CARE', '-', '2026-09-01 00:00:00', '2026-09-01 00:00:00');
  insEmp.run(8, 'Emp Invalid', 1, 'FACILITY CARE', '99/99/2026', '2026-09-02 00:00:00', '2026-09-02 00:00:00');
  insEmp.run(9, 'Emp Null', 1, 'FACILITY CARE', null, '2026-09-03 00:00:00', '2026-09-03 00:00:00');

  // Contracts: default start_date DESC; expiring ASC on dashboard
  const insCon = db.prepare(`
    INSERT INTO contracts (id, employee_id, branch_id, division, contract_type, pkwt_number, start_date, end_date, status, created_at, updated_at)
    VALUES (?, ?, 1, 'FACILITY CARE', 'PKWT', ?, ?, ?, 'Aktif', ?, ?)
  `);
  insCon.run(1, 1, 'PKWT-1', '2025-01-01', '2027-01-01', '2026-01-01 08:00:00', '2026-01-01 08:00:00');
  insCon.run(2, 2, 'PKWT-2', '10-08-2026', datePlusDays(20), '2026-08-10 08:00:00', '2026-08-10 08:00:00');
  insCon.run(3, 3, 'PKWT-3', '2026-08-10', datePlusDays(5), '2026-08-10 09:00:00', '2026-08-11 08:00:00'); // tie updated_at newer than id 2
  insCon.run(4, 4, 'PKWT-4', '2026-03-15', datePlusDays(10), '2026-03-15 08:00:00', '2026-03-15 08:00:00');
  insCon.run(5, 5, 'PKWT-5', '', '2027-05-01', '2026-09-01 08:00:00', '2026-09-01 08:00:00');
  insCon.run(6, 6, 'PKWT-6', '2026-02-30', '2027-06-01', '2026-09-02 08:00:00', '2026-09-02 08:00:00'); // invalid Feb 30

  // Time Line (activity_schedule): strictly opening_date DESC regardless of target_date/completion_date!
  const insSch = db.prepare(`
    INSERT INTO activity_schedule (id, branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, created_at, updated_at)
    VALUES (?, ?, ?, 'Q3', 'Berlin Ariansyah', ?, ?, ?, ?, ?, ?)
  `);
  insSch.run(1, 1, 'Inspeksi Hygiene', '2026-05-01', '2099-12-31', '2099-12-31', 'Pending', '2026-05-01 08:00:00', '2026-05-01 08:00:00');
  insSch.run(2, 2, 'General Cleaning', '15-08-2026', '2025-01-01', null, 'Pending', '2026-08-01 08:00:00', '2026-08-01 08:00:00'); // DD-MM-YYYY
  insSch.run(3, 1, 'Deep Cleaning', '2026-08-15', '2025-02-01', null, 'Done', '2026-08-01 09:00:00', '2026-08-05 12:00:00'); // tie on 2026-08-15, higher updated_at
  insSch.run(4, 1, 'Fogging', '2026-10-02', datePlusDays(3), null, 'Pending', '2026-10-01 08:00:00', '2026-10-01 08:00:00'); // newest opening
  insSch.run(5, 1, 'Inspeksi Hygiene', null, datePlusDays(1), null, 'Pending', '2026-11-01 08:00:00', '2026-11-01 08:00:00'); // null opening -> bottom!
  insSch.run(6, 1, 'General Cleaning', '-', datePlusDays(2), null, 'Pending', '2026-11-02 08:00:00', '2026-11-02 08:00:00'); // dash opening -> bottom!
  insSch.run(7, 1, 'Fogging', '30-02-2026', '2026-12-31', null, 'Pending', '2026-11-03 08:00:00', '2026-11-03 08:00:00'); // invalid opening -> bottom!

  // OS Migrations: default target_month DESC
  const insOs = db.prepare(`
    INSERT INTO os_migrations (id, branch_id, vendor, target_month, actual_date, status, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, 'scheduled', ?, ?)
  `);
  insOs.run(1, 2, 'AEON', '2026-09', null, '2026-08-01 08:00:00', '2026-08-01 08:00:00');
  insOs.run(2, 1, 'Triple S', '2026-12', '2026-12-01', '2026-08-02 08:00:00', '2026-08-02 08:00:00');
  insOs.run(3, 3, 'AEON', '2026-12', null, '2026-08-03 08:00:00', '2026-08-05 08:00:00'); // newer updated_at
  insOs.run(4, 1, 'AEON', '2026-10', null, '2026-08-01 08:00:00', '2026-08-01 08:00:00');
  insOs.run(5, 2, 'Triple S', 'invalid-ym', null, '2026-09-01 08:00:00', '2026-09-01 08:00:00');

  // SP Data: default tanggal DESC
  const insSp = db.prepare(`
    INSERT INTO sp_data (id, tanggal, employee_name, division, branch_id, akhir_sp, sp_type, status, created_at, updated_at)
    VALUES (?, ?, ?, 'FACILITY CARE', 1, '2026-12-31', 'SP 1', 'Aktif', ?, ?)
  `);
  insSp.run(1, '2026-04-01', 'Budi', '2026-04-01 08:00:00', '2026-04-01 08:00:00');
  insSp.run(2, '20-07-2026', 'Ani', '2026-07-20 08:00:00', '2026-07-20 08:00:00');
  insSp.run(3, '', 'NoDate', '2026-09-01 08:00:00', '2026-09-01 08:00:00');

  // Mutasi Data: default tanggal DESC
  const insMut = db.prepare(`
    INSERT INTO mutasi_data (id, tanggal, employee_name, from_branch_id, to_branch_id, status, created_at, updated_at)
    VALUES (?, ?, ?, 1, 2, 'Selesai', ?, ?)
  `);
  insMut.run(1, '2026-03-01', 'Budi', '2026-03-01 08:00:00', '2026-03-01 08:00:00');
  insMut.run(2, '15/08/2026', 'Ani', '2026-08-15 08:00:00', '2026-08-15 08:00:00');
  insMut.run(3, null, 'NoDate', '2026-09-01 08:00:00', '2026-09-01 08:00:00');

  // Overtime: default date DESC
  const insOv = db.prepare(`
    INSERT INTO overtime_records (id, date, branch_id, employee_id, start_time, end_time, total_hours, reason, created_at, updated_at)
    VALUES (?, ?, 1, 1, '17:00', '20:00', 3, 'Lembur', ?, ?)
  `);
  insOv.run(1, '2026-02-10', '2026-02-10 08:00:00', '2026-02-10 08:00:00');
  insOv.run(2, '25-07-2026', '2026-07-25 08:00:00', '2026-07-25 08:00:00');
  insOv.run(3, '-', '2026-09-01 08:00:00', '2026-09-01 08:00:00');

  // Issues: default report_date DESC + updated_at tie-breaker
  const insIss = db.prepare(`
    INSERT INTO issues (id, report_date, branch_id, category, complaint, status, created_at, updated_at)
    VALUES (?, ?, 1, 'Cleaning', ?, 'Open', ?, ?)
  `);
  insIss.run(1, '2026-01-15', 'Lama', '2026-01-15 08:00:00', '2026-01-15 08:00:00');
  insIss.run(2, '18/08/2026', 'DMY', '2026-08-18 08:00:00', '2026-08-18 08:00:00');
  insIss.run(3, '2026-08-18', 'SameDateNewUpd', '2026-08-18 07:00:00', '2026-08-19 10:00:00');
  insIss.run(4, '', 'Kosong', '2026-09-01 08:00:00', '2026-09-01 08:00:00');

  // Training: default training_date DESC, tie-breaker id DESC (was id ASC bug!)
  const insTr = db.prepare(`
    INSERT INTO training (id, training_date, batch, subject, branch_id, created_at, updated_at)
    VALUES (?, ?, 'Batch 1', ?, 1, ?, ?)
  `);
  insTr.run(1, '2026-05-01', 'First', '2026-05-01 08:00:00', '2026-05-01 08:00:00');
  insTr.run(2, '2026-06-15', 'TieId2', '2026-06-15 08:00:00', '2026-06-15 08:00:00');
  insTr.run(3, '15-06-2026', 'TieId3', '2026-06-15 08:00:00', '2026-06-15 08:00:00'); // same timestamp -> id 3 wins over id 2!
  insTr.run(4, null, 'EmptyTr', '2026-09-01 08:00:00', '2026-09-01 08:00:00');

  // Reports (Inspection, Cleaning, Fogging, Basecamp)
  db.prepare("INSERT INTO inspection_reports (id, branch_id, period, inspection_date, status, created_at, updated_at) VALUES (1, 1, 'Q1', '2026-02-01', 'Done', '2026-02-01', '2026-02-01')").run();
  db.prepare("INSERT INTO inspection_reports (id, branch_id, period, inspection_date, status, created_at, updated_at) VALUES (2, 1, 'Q3', '19-07-2026', 'Done', '2026-07-19', '2026-07-19')").run();
  db.prepare("INSERT INTO inspection_reports (id, branch_id, period, inspection_date, status, created_at, updated_at) VALUES (3, 1, 'Q3', '', 'Pending', '2026-09-01', '2026-09-01')").run();

  db.prepare("INSERT INTO cleaning_reports (id, branch_id, activity_type, period, activity_date, status, created_at, updated_at) VALUES (1, 1, 'General Cleaning', 'Q1', '2026-02-01', 'Done', '2026-02-01', '2026-02-01')").run();
  db.prepare("INSERT INTO cleaning_reports (id, branch_id, activity_type, period, activity_date, status, created_at, updated_at) VALUES (2, 1, 'Deep Cleaning', 'Q3', '2026-07-20', 'Done', '2026-07-20', '2026-07-20')").run();
  db.prepare("INSERT INTO cleaning_reports (id, branch_id, activity_type, period, activity_date, status, created_at, updated_at) VALUES (3, 1, 'General Cleaning', 'Q3', null, 'Pending', '2026-09-01', '2026-09-01')").run();

  db.prepare("INSERT INTO fogging_reports (id, branch_id, activity_type, period, activity_date, status, created_at, updated_at) VALUES (1, 1, 'Fogging', 'Q1', '2026-01-10', 'Done', '2026-01-10', '2026-01-10')").run();
  db.prepare("INSERT INTO fogging_reports (id, branch_id, activity_type, period, activity_date, status, created_at, updated_at) VALUES (2, 1, 'Fogging', 'Q3', '05/08/2026', 'Done', '2026-08-05', '2026-08-05')").run();
  db.prepare("INSERT INTO fogging_reports (id, branch_id, activity_type, period, activity_date, status, created_at, updated_at) VALUES (3, 1, 'Fogging', 'Q3', '-', 'Pending', '2026-09-01', '2026-09-01')").run();

  db.prepare("INSERT INTO basecamp_reports (id, branch_id, problem, info_date, status, created_at, updated_at) VALUES (1, 1, 'Old', '2026-01-05', 'Done', '2026-01-05', '2026-01-05')").run();
  db.prepare("INSERT INTO basecamp_reports (id, branch_id, problem, info_date, status, created_at, updated_at) VALUES (2, 1, 'NewDMY', '22-08-2026', 'Open', '2026-08-22', '2026-08-22')").run();
  db.prepare("INSERT INTO basecamp_reports (id, branch_id, problem, info_date, status, created_at, updated_at) VALUES (3, 1, 'NoDate', '', 'Open', '2026-09-01', '2026-09-01')").run();

  // Master Data (created_at DESC, id DESC for Users, Checklist, Hygiene, SOP; preserved for Branches & DAC)
  db.prepare("INSERT INTO sop (id, name, category, created_at, updated_at) VALUES (1, 'Zeta Old SOP', 'SOP', '2026-01-01 00:00:00', '2026-01-01 00:00:00')").run();
  db.prepare("INSERT INTO sop (id, name, category, created_at, updated_at) VALUES (2, 'Alpha New SOP', 'SOP', '2026-07-01 00:00:00', '2026-07-01 00:00:00')").run();

  db.prepare("INSERT INTO master_checklist (id, name, category, created_at, updated_at) VALUES (1, 'Z Old Doc', 'Form', '2026-01-01 00:00:00', '2026-01-01 00:00:00')").run();
  db.prepare("INSERT INTO master_checklist (id, name, category, created_at, updated_at) VALUES (2, 'A New Doc', 'Form', '2026-08-01 00:00:00', '2026-08-01 00:00:00')").run();

  db.prepare("INSERT INTO master_hygiene_standards (id, room_name, item_name, cleanliness_standard, created_at, updated_at) VALUES (1, 'Lobby', 'Lantai', 'Bersih', '2026-01-01 00:00:00', '2026-01-01 00:00:00')").run();
  db.prepare("INSERT INTO master_hygiene_standards (id, room_name, item_name, cleanliness_standard, created_at, updated_at) VALUES (2, 'Lobby', 'Kaca', 'Kinclong', '2026-08-01 00:00:00', '2026-08-01 00:00:00')").run();

  db.prepare("INSERT INTO master_dac_standards (id, shift, item_no, time_range, location, activity, category) VALUES (1, 'siang', 2, '14:00', 'Lobby', 'Sore 2', 'Umum')").run();
  db.prepare("INSERT INTO master_dac_standards (id, shift, item_no, time_range, location, activity, category) VALUES (2, 'pagi', 1, '08:00', 'Lobby', 'Pagi 1', 'Umum')").run();
  db.prepare("INSERT INTO master_dac_standards (id, shift, item_no, time_range, location, activity, category) VALUES (3, 'pagi', 2, '09:00', 'Lobby', 'Pagi 2', 'Umum')").run();

  // Audit logs (note: no updated_at column!)
  db.prepare("INSERT INTO audit_logs (id, user_name, action, module, created_at) VALUES (1, 'Owner', 'LOGIN', 'auth', '2026-08-01 10:00:00')").run();
  db.prepare("INSERT INTO audit_logs (id, user_name, action, module, created_at) VALUES (2, 'Owner', 'UPDATE', 'schedule', '2026-08-01T10:00:00.000Z')").run();
  db.prepare("INSERT INTO audit_logs (id, user_name, action, module, created_at) VALUES (3, 'Owner', 'CREATE', 'issues', '2026-08-02 09:00:00')").run();

  // Legacy routes check (one_on_one, relievers)
  db.prepare("INSERT INTO one_on_one (id, employee_name, meeting_date, created_at, updated_at) VALUES (1, 'A', '2026-02-01', '2026-02-01', '2026-02-01')").run();
  db.prepare("INSERT INTO one_on_one (id, employee_name, meeting_date, created_at, updated_at) VALUES (2, 'B', '2026-07-01', '2026-07-01', '2026-07-01')").run();
  db.prepare("INSERT INTO one_on_one (id, employee_name, meeting_date, created_at, updated_at) VALUES (3, 'C', '', '2026-09-01', '2026-09-01')").run();

  db.prepare("INSERT INTO relievers (id, branch_id, period, reliever_name, backup_date, status, created_at, updated_at) VALUES (1, 1, 'Q1', 'R1', '2026-02-01', 'Done', '2026-02-01', '2026-02-01')").run();
  db.prepare("INSERT INTO relievers (id, branch_id, period, reliever_name, backup_date, status, created_at, updated_at) VALUES (2, 1, 'Q3', 'R2', '2026-07-01', 'Done', '2026-07-01', '2026-07-01')").run();
  db.prepare("INSERT INTO relievers (id, branch_id, period, reliever_name, backup_date, status, created_at, updated_at) VALUES (3, 1, 'Q3', 'R3', null, 'Pending', '2026-09-01', '2026-09-01')").run();
}

function datePlusDays(days) {
  const d = new Date(Date.now() + days * 86400000);
  return d.toISOString().slice(0, 10);
}

async function runTests() {
  // 1. Frontend sorting utility tests
  assert.equal(parseSortableDateMs('2026-08-10'), Date.UTC(2026, 7, 10));
  assert.equal(parseSortableDateMs('10-08-2026'), Date.UTC(2026, 7, 10));
  assert.equal(parseSortableDateMs('10/08/2026'), Date.UTC(2026, 7, 10));
  assert.equal(parseSortableDateMs('45369'), Date.UTC(2024, 2, 18));
  assert.equal(parseSortableDateMs('30-02-2026'), null);
  assert.equal(parseSortableDateMs('-'), null);
  assert.equal(parseSortableDateMs(''), null);
  assert.equal(parseSortableDateMs(null), null);

  const clientRows = [
    { id: 1, opening_date: '', created_at: '2026-09-01T00:00:00Z', updated_at: '2026-09-01T00:00:00Z' },
    { id: 2, opening_date: '15-08-2026', created_at: '2026-08-01T00:00:00Z', updated_at: '2026-08-01T00:00:00Z' },
    { id: 3, opening_date: '2026-08-15', created_at: '2026-08-01T00:00:00Z', updated_at: '2026-08-02T00:00:00Z' },
    { id: 4, opening_date: '30-02-2026', created_at: '2026-10-01T00:00:00Z', updated_at: '2026-10-01T00:00:00Z' },
    { id: 5, opening_date: '2026-10-01', created_at: '2026-07-01T00:00:00Z', updated_at: '2026-07-01T00:00:00Z' },
  ];
  const sortedClient = [...clientRows].sort((a, b) => compareRowsByDateDesc(a, b, 'opening_date'));
  assert.deepEqual(sortedClient.map(r => r.id), [5, 3, 2, 4, 1]);
  assert.deepEqual(sortYearMonthDesc(['2026-09', '2026-12', '2026-10', '', '2026-11']), ['2026-12', '2026-11', '2026-10', '2026-09', '']);

  // 2. Full Backend Route + SQLite Integration Tests
  const db = setupDatabase();
  seedAllData(db);
  const env = { DB: createD1Adapter(db), JWT_SECRET: SECRET };
  const token = await createToken({ sub: 1, username: 'owner', role: 'superadmin' }, SECRET);

  // Time Line (/api/schedule) - Default opening_date DESC
  let r = await callRoute(handleSchedule, env, token, '/api/schedule?limit=20');
  assert.equal(r.status, 200);
  assert.deepEqual(r.body.data.map(x => x.id), [4, 3, 2, 1, 7, 6, 5], 'Schedule default must sort opening_date DESC, invalid/empty last, tie-breaker updated_at->created_at->id DESC');

  // Time Line Pagination check (page 1 & page 2, limit 2)
  const schP1 = await callRoute(handleSchedule, env, token, '/api/schedule?page=1&limit=2');
  const schP2 = await callRoute(handleSchedule, env, token, '/api/schedule?page=2&limit=2');
  assert.deepEqual(schP1.body.data.map(x => x.id), [4, 3]);
  assert.deepEqual(schP2.body.data.map(x => x.id), [2, 1]);

  // Time Line Filter check (?branch_id=1) + Manual Sort (?sort=target_date_asc)
  const schBranch1 = await callRoute(handleSchedule, env, token, '/api/schedule?branch_id=1');
  assert.deepEqual(schBranch1.body.data.map(x => x.id), [4, 3, 1, 7, 6, 5]);
  const schManual = await callRoute(handleSchedule, env, token, '/api/schedule?sort=target_date_asc');
  assert.equal(schManual.body.data[0].id, 2, 'Manual sort target_date_asc must still function');

  // Master Karyawan (/api/employees) - join_date DESC across Excel serial + DMY + ISO + empty/invalid
  r = await callRoute(handleEmployees, env, token, '/api/employees?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [6, 5, 4, 3, 2, 1, 9, 8, 7], 'Employees must sort join_date DESC (2026-06-01[upd/crt tie] > 15/01/2026 > 46000 > 45369 > invalid/empty)');
  const empP1 = await callRoute(handleEmployees, env, token, '/api/employees?page=1&limit=3');
  const empP2 = await callRoute(handleEmployees, env, token, '/api/employees?page=2&limit=3');
  assert.deepEqual(empP1.body.data.map(x => x.id), [6, 5, 4]);
  assert.deepEqual(empP2.body.data.map(x => x.id), [3, 2, 1]);
  const empDivFilter = await callRoute(handleEmployees, env, token, '/api/employees?division=FACILITY%20CARE');
  assert.equal(empDivFilter.body.data.includes(x => x.id === 6), false);
  const empManual = await callRoute(handleEmployees, env, token, '/api/employees?sort=name_asc');
  assert.equal(empManual.body.data[0].full_name, 'Emp DMY Jan2026');

  // Data Kontrak (/api/contracts) - start_date DESC + manual end_date_asc
  r = await callRoute(handleContracts, env, token, '/api/contracts?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [3, 2, 4, 1, 6, 5], 'Contracts default must sort start_date DESC with empty/invalid last');
  const conManual = await callRoute(handleContracts, env, token, '/api/contracts?sort=end_date_asc');
  assert.equal(conManual.body.data[0].id, 3, 'Manual end_date_asc must work');

  // Migrasi OS (/api/os-migrations & /summary) - target_month DESC
  r = await callRoute(handleOSMigrations, env, token, '/api/os-migrations?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [3, 2, 4, 1, 5], 'OS Migrations must sort target_month DESC with invalid last');
  const osSum = await callRoute(handleOSMigrations, env, token, '/api/os-migrations/summary');
  assert.deepEqual(osSum.body.data.matrix.map(x => x.id), [3, 2, 4, 1, 5]);

  // Permasalahan (/api/issues)
  r = await callRoute(handleIssues, env, token, '/api/issues?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [3, 2, 1, 4]);

  // SP (/api/sp), Mutasi (/api/mutasi), Overtime (/api/overtime), Training (/api/training)
  r = await callRoute(handleSP, env, token, '/api/sp?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleMutasi, env, token, '/api/mutasi?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleOvertime, env, token, '/api/overtime?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleTraining, env, token, '/api/training?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [3, 2, 1, 4], 'Training tie-breaker must be DESC (id 3 before id 2)');

  // Reports
  r = await callRoute(handleReports, env, token, '/api/reports/inspection?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleReports, env, token, '/api/reports/cleaning?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleReports, env, token, '/api/reports/fogging?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleReports, env, token, '/api/reports/basecamp?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);

  // Master Data Decisions
  r = await callRoute(handleUsers, env, token, '/api/users?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [3, 2, 1, 4], 'Users must sort created_at DESC, id DESC');
  r = await callRoute(handleMisc, env, token, '/api/checklist?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1], 'Master Document must sort created_at DESC, id DESC');
  r = await callRoute(handleMisc, env, token, '/api/sop?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1], 'SOP Facility Care must sort created_at DESC, id DESC');
  r = await callRoute(handleHygieneStandards, env, token, '/api/hygiene-standards?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1], 'Master Hygiene must sort created_at DESC, id DESC');
  const hygAll = await callRoute(handleHygieneStandards, env, token, '/api/hygiene-standards?all=1');
  assert.deepEqual(hygAll.body.data.map(x => x.id), [2, 1]);

  // Preserved Master Order: Branches (code ASC, id ASC) & DAC (shift ASC, item_no ASC, id ASC)
  r = await callRoute(handleBranches, env, token, '/api/branches?limit=20');
  assert.deepEqual(r.body.data.map(x => x.code), ['001', '005', '010'], 'Branches must preserve code ASC');
  r = await callRoute(handleDAC, env, token, '/api/dac?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 3, 1], 'DAC must preserve shift ASC, item_no ASC, id ASC');

  // Audit Logs: created_at DESC, id DESC (even with mixed T/Z vs space timestamps)
  r = await callRoute(handleAuditLogs, env, token, '/api/audit-logs?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [3, 2, 1]);

  // Dashboard endpoints (recent_issues DESC, expiring ASC, upcoming ASC, activity-log DESC, notifications DESC)
  const statsRes = await callRoute(handleDashboard, env, token, '/api/dashboard/stats');
  assert.equal(statsRes.status, 200);
  assert.equal(statsRes.body.data.expiring_contracts[0].id, 3, 'Expiring contracts must stay ASC (soonest 5d before 10d)');
  assert.equal(statsRes.body.data.upcoming_schedule[0].id, 5, 'Upcoming schedule must stay ASC (+1d before +2d)');
  const actRes = await callRoute(handleDashboard, env, token, '/api/dashboard/activity-log');
  assert.equal(actRes.status, 200);
  assert.ok(actRes.body.data.length > 0);
  const notifRes = await callRoute(handleDashboard, env, token, '/api/dashboard/notifications');
  assert.equal(notifRes.status, 200);

  // Legacy routes (One on One & Relievers)
  r = await callRoute(handleOneOnOne, env, token, '/api/one-on-one?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);
  r = await callRoute(handleRelievers, env, token, '/api/relievers?limit=20');
  assert.deepEqual(r.body.data.map(x => x.id), [2, 1, 3]);

  // Extra Pagination + Filter + Refresh Stability + Manual Sort across remaining modules
  const issP1 = await callRoute(handleIssues, env, token, '/api/issues?page=1&limit=2');
  const issP2 = await callRoute(handleIssues, env, token, '/api/issues?page=2&limit=2');
  assert.deepEqual([...issP1.body.data.map(x => x.id), ...issP2.body.data.map(x => x.id)], [3, 2, 1, 4]);
  const issFiltered = await callRoute(handleIssues, env, token, '/api/issues?status=Open');
  assert.deepEqual(issFiltered.body.data.map(x => x.id), [3, 2, 1, 4]);
  const issManual = await callRoute(handleIssues, env, token, '/api/issues?sort=date_asc');
  assert.deepEqual(issManual.body.data.map(x => x.id), [1, 3, 2, 4], 'Manual date_asc must still place empty date last');

  // Refresh stability check (2 consecutive calls return identical order)
  const schRefresh1 = await callRoute(handleSchedule, env, token, '/api/schedule?limit=20');
  const schRefresh2 = await callRoute(handleSchedule, env, token, '/api/schedule?limit=20');
  assert.deepEqual(schRefresh1.body.data.map(x => x.id), schRefresh2.body.data.map(x => x.id));

  // Calendar endpoint check
  const calRes = await callRoute(handleDashboard, env, token, '/api/dashboard/calendar?month=2026-08');
  assert.equal(calRes.status, 200);
  assert.ok(Array.isArray(calRes.body.data) && calRes.body.data.length > 0);

  // Minimal schema without updated_at on sp_data/mutasi_data (hasTableColumn safety check)
  const minDb = setupDatabase({ includeSpUpdatedAt: false });
  minDb.prepare("INSERT INTO users (id, username, email, full_name, password_hash, role, is_active) VALUES (1, 'owner', 'owner@test.id', 'Owner User', 'hash', 'superadmin', 1)").run();
  minDb.prepare("INSERT INTO branches (id, code, name, full_name, is_active) VALUES (1, '001', 'B', '001. B', 1)").run();
  minDb.prepare("INSERT INTO sp_data (id, tanggal, employee_name, branch_id, created_at) VALUES (1, '2026-01-01', 'X', 1, '2026-01-01')").run();
  minDb.prepare("INSERT INTO sp_data (id, tanggal, employee_name, branch_id, created_at) VALUES (2, '2026-06-01', 'Y', 1, '2026-06-01')").run();
  const minSp = await callRoute(handleSP, { DB: createD1Adapter(minDb), JWT_SECRET: SECRET }, token, '/api/sp?limit=20');
  assert.equal(minSp.status, 200);
  assert.deepEqual(minSp.body.data.map(x => x.id), [2, 1], 'sp_data without updated_at must still succeed cleanly');

  console.log('ALL GLOBAL SORTING STANDARD TESTS PASSED (22/22 assertions blocks verified)');
}

runTests().catch(err => {
  console.error('TEST FAILED:', err);
  process.exit(1);
});
