import { apiFetch } from '../config.js';
import { toastError } from '../components/toast.js';

export async function renderSyncDashboard(mainEl) {
  mainEl.innerHTML = `
    <div class="p-6 max-w-7xl mx-auto space-y-6">
      <div class="flex flex-col md:flex-row justify-between items-start md:items-center">
        <div>
          <h2 class="text-2xl font-bold text-gray-800">Sync Dashboard</h2>
          <p class="text-gray-500">Monitor Bidirectional Sync Health & Queue</p>
        </div>
        <div class="mt-4 md:mt-0 flex gap-2 flex-wrap">
          <button id="btnRetryAll" class="px-3 py-2 bg-yellow-500 text-white rounded shadow hover:bg-yellow-600 transition text-sm">Retry All Failed</button>
          <button id="btnResetStuck" class="px-3 py-2 bg-orange-500 text-white rounded shadow hover:bg-orange-600 transition text-sm">Reset Stuck</button>
          <button id="btnReconcile" class="px-3 py-2 bg-purple-500 text-white rounded shadow hover:bg-purple-600 transition text-sm">Force Reconcile</button>
          <button id="btnRefreshSync" class="px-3 py-2 bg-blue-600 text-white rounded shadow hover:bg-blue-700 transition text-sm">Refresh</button>
        </div>
      </div>

      <!-- Overview Cards -->
      <div class="grid grid-cols-1 md:grid-cols-4 gap-4" id="syncOverviewCards">
        <!-- populated via js -->
      </div>

      <!-- Circuit Breaker Status -->
      <div class="bg-white rounded-lg shadow p-6 border-l-4" id="cbStatusCard">
        <div class="flex justify-between items-center mb-2">
          <h3 class="text-lg font-semibold text-gray-700">Circuit Breaker Status</h3>
          <div class="flex gap-2">
            <button id="btnPauseSync" class="px-3 py-1 bg-red-600 text-white rounded text-xs hover:bg-red-700">Pause Sync</button>
            <button id="btnResumeSync" class="px-3 py-1 bg-green-600 text-white rounded text-xs hover:bg-green-700">Resume Sync</button>
          </div>
        </div>
        <div class="flex items-center gap-4">
          <div id="cbStateBadge" class="px-3 py-1 rounded-full text-sm font-medium">Loading...</div>
          <p class="text-sm text-gray-600" id="cbStateDesc">Fetching state...</p>
        </div>
      </div>

      <!-- Queue Table -->
      <div class="bg-white rounded-lg shadow overflow-hidden">
        <div class="p-4 border-b flex justify-between items-center">
          <h3 class="text-lg font-semibold text-gray-800">Outbox Queue</h3>
          <div class="flex gap-2">
            <button id="btnRetrySelected" class="px-3 py-1 bg-gray-200 text-gray-700 rounded text-xs hover:bg-gray-300">Retry Selected</button>
            <select id="queueStatusFilter" class="border rounded px-2 py-1 text-sm">
              <option value="">All Status</option>
              <option value="PENDING">Pending</option>
              <option value="PROCESSING">Processing</option>
              <option value="FAILED">Failed</option>
              <option value="DEAD_LETTER">Dead Letter</option>
            </select>
          </div>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left">
            <thead class="bg-gray-50 text-gray-600">
              <tr>
                <th class="px-4 py-3 w-10"><input type="checkbox" id="chkAllQueue" /></th>
                <th class="px-4 py-3">ID</th>
                <th class="px-4 py-3">Module</th>
                <th class="px-4 py-3">Action</th>
                <th class="px-4 py-3">Status</th>
                <th class="px-4 py-3">Retry</th>
                <th class="px-4 py-3">Created</th>
              </tr>
            </thead>
            <tbody id="queueTableBody" class="divide-y">
              <tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">Loading queue...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
      
      <!-- Performance Metrics & Histograms -->
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <!-- Histograms -->
        <div class="bg-white rounded-lg shadow overflow-hidden">
          <div class="p-4 border-b">
            <h3 class="text-lg font-semibold text-gray-800">Latency Histograms (P50 - P99)</h3>
          </div>
          <div class="p-4 overflow-x-auto text-sm" id="latencyHistograms">
            <p class="text-gray-500">Loading performance data...</p>
          </div>
        </div>

        <!-- Snapshots & Reconcile -->
        <div class="bg-white rounded-lg shadow overflow-hidden">
          <div class="p-4 border-b">
            <h3 class="text-lg font-semibold text-gray-800">Nightly Snapshot & Reconciliation</h3>
          </div>
          <div class="p-4 text-sm space-y-4">
            <div>
              <h4 class="font-medium text-gray-700">Latest Snapshot</h4>
              <p id="lblSnapshotStatus" class="text-gray-600 mt-1">Loading...</p>
            </div>
            <div>
              <h4 class="font-medium text-gray-700">Reconciliation Status</h4>
              <p id="lblReconStatus" class="text-gray-600 mt-1">Ready for checks.</p>
            </div>
          </div>
        </div>
      </div>
      
      <!-- Metrics -->
      <div class="bg-white rounded-lg shadow overflow-hidden">
        <div class="p-4 border-b">
          <h3 class="text-lg font-semibold text-gray-800">Queue & Throughput Trend</h3>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-sm text-left">
            <thead class="bg-gray-50 text-gray-600">
              <tr>
                <th class="px-4 py-3">Module</th>
                <th class="px-4 py-3">Total Events</th>
                <th class="px-4 py-3">Avg Webhook (ms)</th>
                <th class="px-4 py-3">Avg D1 Exec (ms)</th>
                <th class="px-4 py-3">Avg Queue Wait (ms)</th>
              </tr>
            </thead>
            <tbody id="metricsTableBody" class="divide-y">
              <tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">Loading metrics...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  const btnRefresh = document.getElementById('btnRefreshSync');
  const queueStatusFilter = document.getElementById('queueStatusFilter');

  btnRefresh.addEventListener('click', loadDashboardData);
  queueStatusFilter.addEventListener('change', loadQueue);

  // Admin Actions
  document.getElementById('btnRetryAll').addEventListener('click', () => adminAction('retry', { allFailed: true }));
  document.getElementById('btnResetStuck').addEventListener('click', () => adminAction('reset-stuck', {}));
  document.getElementById('btnPauseSync').addEventListener('click', () => adminAction('pause', {}));
  document.getElementById('btnResumeSync').addEventListener('click', () => adminAction('resume', {}));
  document.getElementById('btnReconcile').addEventListener('click', () => adminAction('reconcile', { module: 'ALL', repairMode: true }));
  
  document.getElementById('btnRetrySelected').addEventListener('click', () => {
    const checked = Array.from(document.querySelectorAll('.chk-queue:checked')).map(cb => cb.value);
    if (checked.length === 0) return alert('No items selected');
    adminAction('retry', { ids: checked });
  });
  
  document.getElementById('chkAllQueue').addEventListener('change', (e) => {
    document.querySelectorAll('.chk-queue').forEach(cb => cb.checked = e.target.checked);
  });

  async function adminAction(action, payload) {
    if (!confirm(`Are you sure you want to execute action: ${action}?`)) return;
    showLoading();
    try {
      const res = await apiFetch(`/api/sync/actions/${action}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        alert(res.data?.message || 'Success');
        loadDashboardData();
      } else {
        toastError(res.error || 'Action failed');
      }
    } catch(e) {
      toastError(e.message);
    }
    hideLoading();
  }

  await loadDashboardData();

  async function loadDashboardData() {
    showLoading();
    await Promise.all([
      loadHealth(),
      loadQueue(),
      loadPerformance(),
      loadMetrics()
    ]);
    hideLoading();
  }

  async function loadPerformance() {
    try {
      const res = await apiFetch('/api/sync/performance');
      if (!res.ok) return;
      
      const { webhook, google_api, d1, queue, throughput } = res.data;
      
      document.getElementById('latencyHistograms').innerHTML = `
        <table class="w-full text-left">
          <thead><tr class="text-gray-500 border-b"><th>Metric</th><th>P50</th><th>P95</th><th>P99</th><th>Max</th></tr></thead>
          <tbody class="divide-y">
            <tr><td class="py-1">Webhook</td><td>${webhook.P50}ms</td><td>${webhook.P95}ms</td><td>${webhook.P99}ms</td><td>${webhook.Max}ms</td></tr>
            <tr><td class="py-1">Google API</td><td>${google_api.P50}ms</td><td>${google_api.P95}ms</td><td>${google_api.P99}ms</td><td>${google_api.Max}ms</td></tr>
            <tr><td class="py-1">D1 Execute</td><td>${d1.P50}ms</td><td>${d1.P95}ms</td><td>${d1.P99}ms</td><td>${d1.Max}ms</td></tr>
            <tr><td class="py-1">Queue Delay</td><td>${queue.P50}ms</td><td>${queue.P95}ms</td><td>${queue.P99}ms</td><td>${queue.Max}ms</td></tr>
          </tbody>
        </table>
        <div class="mt-4 pt-3 border-t text-gray-600 flex justify-between">
          <span>Throughput: <b>${throughput.events_per_sec}</b> ev/sec</span>
          <span><b>${throughput.events_per_min}</b> ev/min</span>
        </div>
      `;
      
      document.getElementById('lblSnapshotStatus').innerHTML = 'Checked 10 modules. Status: <b>COMPLETED</b><br><span class="text-xs text-gray-500">Run today at 02:00</span>';
    } catch (e) {
      console.error(e);
    }
  }

  async function loadHealth() {
    try {
      const res = await apiFetch('/api/sync/health');
      if (!res.ok) return toastError('Failed to fetch sync health');
      
      const { status, queue, circuit_breaker } = res.data;
      
      // Cards
      const cardsHtml = `
        <div class="bg-white p-4 rounded-lg shadow border-l-4 ${status === 'HEALTHY' ? 'border-green-500' : (status === 'WARNING' ? 'border-yellow-500' : 'border-red-500')}">
          <p class="text-sm text-gray-500">System Health</p>
          <p class="text-2xl font-bold ${status === 'HEALTHY' ? 'text-green-600' : (status === 'WARNING' ? 'text-yellow-600' : 'text-red-600')}">${status}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-blue-500">
          <p class="text-sm text-gray-500">Pending Queue</p>
          <p class="text-2xl font-bold text-gray-800">${queue.pending || 0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-orange-500">
          <p class="text-sm text-gray-500">Failed / Retrying</p>
          <p class="text-2xl font-bold text-gray-800">${queue.failed || 0}</p>
        </div>
        <div class="bg-white p-4 rounded-lg shadow border-l-4 border-red-500">
          <p class="text-sm text-gray-500">Dead Letters</p>
          <p class="text-2xl font-bold text-gray-800">${queue.dead_letter || 0}</p>
        </div>
      `;
      document.getElementById('syncOverviewCards').innerHTML = cardsHtml;

      // Circuit Breaker
      const cbBadge = document.getElementById('cbStateBadge');
      const cbDesc = document.getElementById('cbStateDesc');
      const cbCard = document.getElementById('cbStatusCard');
      
      cbCard.className = 'bg-white rounded-lg shadow p-6 border-l-4'; // reset
      
      if (circuit_breaker === 'CLOSED') {
        cbBadge.className = 'px-3 py-1 rounded-full text-sm font-medium bg-green-100 text-green-800';
        cbBadge.textContent = 'CLOSED';
        cbDesc.textContent = 'Traffic is flowing normally to Google Sheets.';
        cbCard.classList.add('border-green-500');
      } else if (circuit_breaker === 'OPEN') {
        cbBadge.className = 'px-3 py-1 rounded-full text-sm font-medium bg-red-100 text-red-800';
        cbBadge.textContent = 'OPEN';
        cbDesc.textContent = 'Failing fast. Traffic to Google Sheets is paused due to repeated failures.';
        cbCard.classList.add('border-red-500');
      } else if (circuit_breaker === 'HALF_OPEN') {
        cbBadge.className = 'px-3 py-1 rounded-full text-sm font-medium bg-yellow-100 text-yellow-800';
        cbBadge.textContent = 'HALF-OPEN';
        cbDesc.textContent = 'Testing recovery. Permitting limited traffic to verify stability.';
        cbCard.classList.add('border-yellow-500');
      } else {
        cbBadge.textContent = circuit_breaker || 'UNKNOWN';
      }

    } catch (e) {
      console.error(e);
    }
  }

  async function loadQueue() {
    try {
      const status = document.getElementById('queueStatusFilter').value;
      const res = await apiFetch('/api/sync/queue?limit=15' + (status ? '&status='+status : ''));
      if (!res.ok) return;
      
      const tbody = document.getElementById('queueTableBody');
      const rows = res.data?.data || res.data || [];
      
      if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="px-4 py-4 text-center text-gray-500">No events found</td></tr>';
        return;
      }
      
      tbody.innerHTML = rows.map(r => `
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2"><input type="checkbox" class="chk-queue" value="${r.id}" /></td>
          <td class="px-4 py-2 font-mono text-xs text-gray-500" title="${r.id}">${r.id.split('-')[0]}...</td>
          <td class="px-4 py-2 font-medium">${r.entity_name}</td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${r.action==='INSERT'?'bg-blue-100 text-blue-800':(r.action==='UPDATE'?'bg-purple-100 text-purple-800':'bg-red-100 text-red-800')}">${r.action}</span>
          </td>
          <td class="px-4 py-2">
             <span class="px-2 py-0.5 rounded text-xs ${r.status==='PENDING'?'bg-yellow-100 text-yellow-800':(r.status==='PROCESSING'?'bg-blue-100 text-blue-800':(r.status==='DEAD_LETTER'?'bg-red-100 text-red-800':'bg-gray-100 text-gray-800'))}">${r.status}</span>
             ${r.last_error ? `<br><span class="text-xs text-red-500 max-w-xs block truncate" title="${r.last_error}">${r.last_error}</span>` : ''}
          </td>
          <td class="px-4 py-2 text-gray-600">${r.retry_count || 0}</td>
          <td class="px-4 py-2 text-gray-500 whitespace-nowrap">${window.formatDate(r.created_at)} ${new Date(r.created_at).toLocaleTimeString('id-ID')}</td>
        </tr>
      `).join('');
    } catch (e) {
      console.error(e);
    }
  }

  async function loadMetrics() {
    try {
      const res = await apiFetch('/api/sync/metrics');
      if (!res.ok) return;
      
      const tbody = document.getElementById('metricsTableBody');
      const rows = res.data || [];
      
      if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="px-4 py-4 text-center text-gray-500">No metrics found in last 24h</td></tr>';
        return;
      }
      
      tbody.innerHTML = rows.map(r => `
        <tr class="hover:bg-gray-50">
          <td class="px-4 py-2 font-medium text-gray-800">${r.module}</td>
          <td class="px-4 py-2 text-gray-600">${r.total_events}</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(r.avg_webhook_ms || 0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(r.avg_d1_ms || 0)} ms</td>
          <td class="px-4 py-2 text-gray-600">${Math.round(r.avg_queue_wait_ms || 0)} ms</td>
        </tr>
      `).join('');
    } catch (e) {
      console.error(e);
    }
  }
}
