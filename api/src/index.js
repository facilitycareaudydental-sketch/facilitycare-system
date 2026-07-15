/**
 * FM Operations Management System - Cloudflare Worker API
 * Routes: /api/auth, /api/users, /api/branches, /api/employees, /api/contracts,
 *         /api/schedule, /api/issues, /api/one-on-one, /api/training,
 *         /api/relievers, /api/reports/*, /api/sop, /api/checklist,
 *         /api/forms, /api/pic, /api/dashboard
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
      if (path.startsWith('/api/sop') || path.startsWith('/api/checklist') ||
          path.startsWith('/api/forms') || path.startsWith('/api/pic')) {
        return handleMisc(request, env, origin);
      }

      return error('API endpoint not found', 404, origin);
    } catch (err) {
      console.error('Worker error:', err);
      return error('Internal server error: ' + err.message, 500, origin);
    }
  }
};
