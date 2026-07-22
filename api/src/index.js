/**
 * FM Operations Management System - Cloudflare Worker API
 * Routes: /api/auth, /api/users, /api/branches, /api/employees, /api/contracts,
 *         /api/schedule, /api/issues, /api/one-on-one, /api/training,
 *         /api/relievers, /api/reports/*, /api/sop, /api/checklist,
 *         /api/forms, /api/pic, /api/dashboard, /api/sp, /api/mutasi
 */

import { handleAuth } from './routes/auth.js';
import { handleUsers } from './routes/users.js';
import { handleBranches } from './routes/branches.js';
import { handleEmployees } from './routes/employees.js';
import { handleContracts } from './routes/contracts.js';
import { handleSchedule } from './routes/schedule.js';
import { handleIssues } from './routes/issues.js';
import { handleOneOnOne } from './routes/one_on_one.js';
import { handleTraining } from './routes/training.js';
import { handleRelievers } from './routes/relievers.js';
import { handleReports } from './routes/reports.js';
import { handleMisc } from './routes/misc.js';
import { handleDashboard } from './routes/dashboard.js';
import { handleImport } from './routes/import.js';
import { handleSP } from './routes/sp.js';
import { handleMutasi } from './routes/mutasi.js';
import { syncGoogleSheets } from './utils/google_sync.js';
import { options, error, ok, forbidden } from './utils/response.js';
import { authenticate, hasPermission } from './utils/auth.js';

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const origin = request.headers.get('Origin') || env.CORS_ORIGIN || '*';

    // Handle CORS preflight
    if (request.method === 'OPTIONS') {
      return options(origin);
    }

    // Only handle /api/* routes
    if (!url.pathname.startsWith('/api/')) {
      return new Response('FM Operations API', {
        status: 200,
        headers: { 'Content-Type': 'text/plain' }
      });
    }

    if (url.pathname === '/api/migrate-sp') {
      try { await env.DB.prepare('ALTER TABLE sp_data ADD COLUMN division TEXT').run(); } catch(e){}
      try { await env.DB.prepare('ALTER TABLE sp_data ADD COLUMN akhir_sp TEXT').run(); } catch(e){}
      return ok({ message: 'Migration applied!' }, 200, origin);
    }

    try {
      const path = url.pathname;

      // Generic Bulk Delete Handler
      if (request.method === 'DELETE' && path.endsWith('/bulk')) {
        const user = await authenticate(request, env);
        if (!user) return new Response('Unauthorized', { status: 401, headers: { 'Access-Control-Allow-Origin': origin } });
        
        const moduleMap = {
          '/api/branches/bulk': { table: 'branches', perm: 'branches' },
          '/api/employees/bulk': { table: 'employees', perm: 'employees' },
          '/api/issues/bulk': { table: 'issues', perm: 'issues', calModule: 'issues' },
          '/api/schedule/bulk': { table: 'activity_schedule', perm: 'schedule', calModule: 'schedule' },
          '/api/one-on-one/bulk': { table: 'one_on_one', perm: 'one_on_one', calModule: 'one_on_one' },
          '/api/training/bulk': { table: 'training', perm: 'training', calModule: 'training' },
          '/api/relievers/bulk': { table: 'relievers', perm: 'relievers', calModule: 'relievers' },
          '/api/sp/bulk': { table: 'sp_data', perm: 'sp' },
          '/api/mutasi/bulk': { table: 'mutasi_data', perm: 'mutasi' },
          '/api/sop/bulk': { table: 'sop', perm: 'sop' },
          '/api/checklist/bulk': { table: 'master_checklist', perm: 'checklist' },
          '/api/forms/bulk': { table: 'master_forms', perm: 'forms' },
          '/api/pic/bulk': { table: 'pic_list', perm: 'pic' },
          '/api/users/bulk': { table: 'users', perm: 'users' },
          '/api/contracts/bulk': { table: 'contracts', perm: 'contracts', calModule: 'contracts' },
          '/api/reports/cleaning/bulk': { table: 'cleaning_reports', perm: 'reports' },
          '/api/reports/fogging/bulk': { table: 'fogging_reports', perm: 'reports' },
          '/api/reports/basecamp/bulk': { table: 'basecamp_reports', perm: 'reports' },
          '/api/reports/supply/bulk': { table: 'supply_requests', perm: 'reports' },
          '/api/reports/inspection/bulk': { table: 'inspection_reports', perm: 'reports' }
        };

        const config = moduleMap[path];
        if (!config) return error('Invalid bulk delete module', 400, origin);
        if (!hasPermission(user, config.perm, 'delete') && !hasPermission(user, config.perm, 'write')) return forbidden(origin);

        try {
          const { ids } = await request.json();
          if (!Array.isArray(ids) || ids.length === 0) return error('No IDs provided', 400, origin);
          
          const chunkSize = 50;
          for (let i = 0; i < ids.length; i += chunkSize) {
            const chunk = ids.slice(i, i + chunkSize);
            const placeholders = chunk.map(() => '?').join(',');
            await env.DB.prepare(`DELETE FROM ${config.table} WHERE id IN (${placeholders})`).bind(...chunk).run();
            
            // Delete associated calendar events if applicable
            if (config.calModule) {
              try {
                await env.DB.prepare(`DELETE FROM calendar_events WHERE module = ? AND reference_id IN (${placeholders})`)
                  .bind(config.calModule, ...chunk).run();
              } catch (calError) {
                // Ignore if calendar_events table does not exist
                console.warn('Could not delete calendar events, table might not exist:', calError.message);
              }
            }
          }
          return ok({ message: `Deleted ${ids.length} items` }, 200, origin);
        } catch (e) {
          return error('Server Error: ' + e.message, 500, origin);
        }
      }

      if (path.startsWith('/api/auth')) return handleAuth(request, env, origin);
      if (path.startsWith('/api/users')) return handleUsers(request, env, origin);
      if (path.startsWith('/api/branches')) return handleBranches(request, env, origin);
      if (path.startsWith('/api/employees')) return handleEmployees(request, env, origin);
      if (path.startsWith('/api/contracts')) return handleContracts(request, env, origin);
      if (path.startsWith('/api/schedule')) return handleSchedule(request, env, origin);
      if (path.startsWith('/api/issues')) return handleIssues(request, env, origin);
      if (path.startsWith('/api/one-on-one')) return handleOneOnOne(request, env, origin);
      if (path.startsWith('/api/training')) return handleTraining(request, env, origin);
      if (path.startsWith('/api/relievers')) return handleRelievers(request, env, origin);
      if (path.startsWith('/api/reports')) return handleReports(request, env, origin);
      if (path.startsWith('/api/dashboard')) return handleDashboard(request, env, origin);
      if (path.startsWith('/api/import')) return handleImport(request, env, origin);
      if (path.startsWith('/api/sp')) return handleSP(request, env, origin);
      if (path.startsWith('/api/mutasi')) return handleMutasi(request, env, origin);
      
      // Manual sync trigger
      if (path === '/api/emergency-fix-dates' && request.method === 'GET') {
        const schedules = await env.DB.prepare("SELECT id, opening_date FROM activity_schedule").all();
        let fixed = 0;
        for (const row of (schedules.results || [])) {
          let oDate = row.opening_date;
          if (oDate && !isNaN(Number(oDate)) && Number(oDate) > 40000) {
             const d = new Date((Number(oDate) - 25569) * 86400 * 1000);
             oDate = d.toISOString().slice(0, 10);
             await env.DB.prepare("UPDATE activity_schedule SET opening_date = ? WHERE id = ?").bind(oDate, row.id).run();
             fixed++;
          }
        }
        return new Response('Fixed opening dates: ' + fixed, { status: 200, headers: { 'Access-Control-Allow-Origin': '*' } });
      }

      if (path === '/api/sync/google-sheets' && request.method === 'POST') {
        const result = await syncGoogleSheets(env);
        return new Response(JSON.stringify(result), {
          status: result.success ? 200 : 500,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': origin }
        });
      }

      if (path.startsWith('/api/sop') || path.startsWith('/api/checklist') ||
        path.startsWith('/api/forms') || path.startsWith('/api/pic') || 
        path.startsWith('/api/options') || path.startsWith('/api/audit-emp') || path.startsWith('/api/audit-clean-up-5') || path.startsWith('/api/audit-clean-up-4') || path.startsWith('/api/audit-clean-up-3') || path.startsWith('/api/audit-emp-clean') || path.startsWith('/api/audit-emp-fixdates') || path.startsWith('/api/audit-emp-dupes') || path.startsWith('/api/audit-duplicates-2')) {
        return handleMisc(request, env, origin);
      }

      return error('API endpoint not found', 404, origin);
    } catch (err) {
      console.error('Worker error:', err);
      return error('Internal server error: ' + err.message, 500, origin);
    }
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(syncGoogleSheets(env));
  }
};
