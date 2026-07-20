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
import { options, error } from './utils/response.js';

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

    try {
      const path = url.pathname;
      const { authenticate, hasPermission, forbidden, badRequest, ok, serverError } = await import('./utils/auth.js');

      // Generic Bulk Delete Handler
      if (request.method === 'DELETE' && path.endsWith('/bulk')) {
        const user = await authenticate(request, env);
        if (!user) return new Response('Unauthorized', { status: 401, headers: { 'Access-Control-Allow-Origin': origin } });
        
        const moduleMap = {
          '/api/employees/bulk': { table: 'employees', perm: 'employees' },
          '/api/issues/bulk': { table: 'issues', perm: 'issues' },
          '/api/schedule/bulk': { table: 'schedule', perm: 'schedule' },
          '/api/one-on-one/bulk': { table: 'one_on_one', perm: 'one_on_one' },
          '/api/training/bulk': { table: 'training', perm: 'training' },
          '/api/relievers/bulk': { table: 'relievers', perm: 'relievers' },
          '/api/sp/bulk': { table: 'sp_data', perm: 'sp' },
          '/api/mutasi/bulk': { table: 'mutasi_data', perm: 'mutasi' },
          '/api/sop/bulk': { table: 'sops', perm: 'sop' },
          '/api/checklist/bulk': { table: 'master_checklists', perm: 'checklist' },
          '/api/forms/bulk': { table: 'master_forms', perm: 'forms' },
          '/api/pic/bulk': { table: 'pics', perm: 'pic' },
          '/api/users/bulk': { table: 'users', perm: 'users' },
          '/api/contracts/bulk': { table: 'contracts', perm: 'contracts' }
        };

        const config = moduleMap[path];
        if (!config) return badRequest('Invalid bulk delete module', origin);
        if (!hasPermission(user, config.perm, 'delete') && !hasPermission(user, config.perm, 'write')) return forbidden(origin);

        try {
          const { ids } = await request.json();
          if (!Array.isArray(ids) || ids.length === 0) return badRequest('No IDs provided', origin);
          const placeholders = ids.map(() => '?').join(',');
          await env.DB.prepare(`DELETE FROM ${config.table} WHERE id IN (${placeholders})`).bind(...ids).run();
          return ok({ message: `Deleted ${ids.length} items` }, 200, origin);
        } catch (e) {
          return serverError(e, origin);
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
      if (path === '/api/sync/google-sheets' && request.method === 'POST') {
        const result = await syncGoogleSheets(env);
        return new Response(JSON.stringify(result), {
          status: result.success ? 200 : 500,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': origin }
        });
      }

      if (path.startsWith('/api/sop') || path.startsWith('/api/checklist') ||
          path.startsWith('/api/forms') || path.startsWith('/api/pic')) {
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
