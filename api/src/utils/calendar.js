// Calendar Sync Engine Utility for FCMS SSOT Architecture

export function getSyncStatements(db, module, referenceId, eventData) {
  const stmts = [];
  stmts.push(db.prepare('DELETE FROM calendar_events WHERE module = ? AND reference_id = ?').bind(module, referenceId));

  if (!eventData) return stmts; // Delete only (e.g. on record deletion)

  // Map module to calendar event mapping
  switch (module) {
    case 'contracts': {
      const { empName, branchId, endDate, status } = eventData;
      if (status !== 'Aktif' || !endDate) break;
      const intervals = [90, 60, 30, 14, 7, 3, 1, 0];
      const end = new Date(endDate + 'T00:00:00');
      if (!isNaN(end.getTime())) {
        intervals.forEach(days => {
          const remDate = new Date(end.getTime());
          remDate.setDate(end.getDate() - days);
          const remStr = remDate.toISOString().slice(0, 10);
          const title = days === 0
            ? `Hari H: Kontrak ${empName} Berakhir`
            : `Reminder H-${days}: Kontrak ${empName} Berakhir`;
          const priority = days <= 7 ? 'critical' : days <= 30 ? 'high' : 'medium';
          const color = days <= 7 ? 'red' : days <= 30 ? 'orange' : 'yellow';

          stmts.push(
            db.prepare(
              `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
            ).bind(module, referenceId, 'expiry_reminder', title, `Kontrak berakhir pada ${endDate}`, branchId, remStr, status, priority, color)
          );
        });
      }
      break;
    }

    case 'schedule': {
      const { activity_type, target_date, status, pic, branch_id, notes } = eventData;
      if (!target_date) break;
      const color = activity_type === 'Inspeksi Hygiene & Aset Bangunan' ? 'blue'
                  : activity_type === 'General Cleaning' ? 'green'
                  : activity_type === 'Deep Cleaning' ? 'purple'
                  : activity_type === 'Fogging' ? 'orange' : 'gray';
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'schedule_task', activity_type, `PIC: ${pic || '-'}. Catatan: ${notes || ''}`, branch_id, target_date, status, 'medium', color)
      );
      break;
    }

    case 'relievers': {
      const { reliever_name, backup_date, status, branch_id, original_fc_name, reason, shift } = eventData;
      if (!backup_date) break;
      const desc = `Menggantikan: ${original_fc_name || '-'}. Shift: ${shift || '-'}. Alasan: ${reason || '-'}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'reliever_shift', `Reliever: ${reliever_name}`, desc, branch_id, backup_date, status, 'medium', 'teal')
      );
      break;
    }

    case 'issues': {
      const { category, report_date, target_date, completion_date, status, branch_id, complaint, employee_name, fc_specialist } = eventData;
      const desc = `Keluhan: ${complaint || ''}. Karyawan: ${employee_name || ''}. Specialist: ${fc_specialist || ''}`;
      
      // 1. Report Date Event
      if (report_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, 'issue_report', `Laporan Permasalahan: ${category}`, desc, branch_id, report_date, status, 'high', 'red')
        );
      }
      // 2. Target/Due Date Event (if exists)
      if (target_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, 'issue_due', `Due Date Permasalahan: ${category}`, desc, branch_id, target_date, status, 'critical', 'red')
        );
      }
      // 3. Closed Date Event (if done)
      if (status === 'Done' && completion_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, 'issue_closed', `Permasalahan Selesai: ${category}`, desc, branch_id, completion_date, status, 'low', 'green')
        );
      }
      break;
    }

    case 'one_on_one': {
      const { employee_name, meeting_date, status, branch_id, pic, problem, solution } = eventData;
      if (!meeting_date) break;
      const desc = `PIC: ${pic || '-'}. Masalah: ${problem || ''}. Solusi: ${solution || ''}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'one_on_one_meeting', `One on One: ${employee_name}`, desc, branch_id, meeting_date, status, 'medium', 'pink')
      );
      break;
    }

    case 'training': {
      const { subject, training_date, branch_id, trainer, notes, score } = eventData;
      if (!training_date) break;
      const desc = `Trainer: ${trainer || '-'}. Score: ${score || '-'}. Catatan: ${notes || ''}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'training_class', `Training: ${subject}`, desc, branch_id, training_date, 'Done', 'medium', 'indigo')
      );
      break;
    }

    case 'cleaning': {
      const { activity_type, activity_date, status, branch_id, notes, period } = eventData;
      if (!activity_date) break;
      const desc = `Periode: ${period || '-'}. Catatan: ${notes || ''}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'cleaning_report', `${activity_type}`, desc, branch_id, activity_date, status, 'medium', 'green')
      );
      break;
    }

    case 'inspection': {
      const { inspection_date, status, branch_id, notes, period, fc_score, spv_score } = eventData;
      if (!inspection_date) break;
      const desc = `Score FC: ${fc_score || '-'}. Score SPV: ${spv_score || '-'}. Catatan: ${notes || ''}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'inspection_report', `Hygiene Inspection (Periode: ${period})`, desc, branch_id, inspection_date, status, 'medium', 'blue')
      );
      break;
    }

    case 'fogging': {
      const { activity_date, status, branch_id, notes, period } = eventData;
      if (!activity_date) break;
      const desc = `Periode: ${period || '-'}. Catatan: ${notes || ''}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'fogging_report', `Fogging`, desc, branch_id, activity_date, status, 'medium', 'orange')
      );
      break;
    }

    case 'basecamp': {
      const { problem, info_date, done_date, status, branch_id, pic, notes } = eventData;
      const desc = `PIC: ${pic || '-'}. Catatan: ${notes || ''}`;
      if (info_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, 'basecamp_info', `Laporan Basecamp: ${problem}`, desc, branch_id, info_date, status, 'medium', 'purple')
        );
      }
      if (status === 'Done' && done_date) {
        stmts.push(
          db.prepare(
            `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(module, referenceId, 'basecamp_done', `Basecamp Selesai: ${problem}`, desc, branch_id, done_date, status, 'low', 'green')
        );
      }
      break;
    }

    case 'supply': {
      const { submitter_name, submitted_at, status, branch_id, tools_items, chemical_items } = eventData;
      if (!submitted_at) break;
      const items = [tools_items, chemical_items].filter(Boolean).join(', ');
      const desc = `Pengirim: ${submitter_name}. Barang: ${items || '-'}`;
      stmts.push(
        db.prepare(
          `INSERT INTO calendar_events (module, reference_id, event_type, title, description, branch_id, start_date, status, priority, color)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(module, referenceId, 'supply_request', `Permintaan Barang`, desc, branch_id, submitted_at.slice(0, 10), status, 'low', 'brown')
      );
      break;
    }
  }

  return stmts;
}

export async function runSync(db, module, referenceId, eventData) {
  const stmts = getSyncStatements(db, module, referenceId, eventData);
  if (stmts.length > 0) {
    await db.batch(stmts);
  }
}
