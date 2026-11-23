// Hash-based router with module error isolation and navigation race-condition protection
const routes = {};
let currentCleanup = null;
let activeNavToken = 0;

export function registerRoute(path, handler) {
  routes[path] = handler;
}

export function navigate(path) {
  window.location.hash = path;
}

export function initRouter() {
  async function handleRoute() {
    const navToken = ++activeNavToken;
    const hash = window.location.hash.replace('#', '') || '/dashboard';
    const [base, ...rest] = hash.split('?');

    // Find matching route (exact or prefix)
    let handler = routes[base];
    if (!handler) {
      for (const [pattern, h] of Object.entries(routes)) {
        if (pattern.endsWith('/*') && base.startsWith(pattern.slice(0, -2))) {
          handler = h;
          break;
        }
      }
    }

    if (currentCleanup) {
      try {
        currentCleanup();
      } catch (cleanupErr) {
        console.warn('[Router] Previous page cleanup error:', cleanupErr?.message || cleanupErr);
      } finally {
        currentCleanup = null;
      }
    }

    const main = document.getElementById('main-content');

    if (main) {
      main.innerHTML = '<div class="loading-spinner"><div class="spinner"></div></div>';
    }

    if (handler) {
      try {
        const params = new URLSearchParams(rest.join('?'));
        const segments = base.split('/').filter(Boolean);
        const cleanup = await handler({ path: base, params, segments, main });
        if (navToken !== activeNavToken) {
          if (typeof cleanup === 'function') {
            try { cleanup(); } catch { /* ignore stale cleanup */ }
          }
          return;
        }
        if (typeof cleanup === 'function') currentCleanup = cleanup;
      } catch (routeErr) {
        if (navToken !== activeNavToken) return;
        console.error(`[Router] Module render error on '${base}':`, routeErr?.message || routeErr);
        const target = main || document.getElementById('app');
        if (target) {
          target.innerHTML = `
            <div class="empty-state" style="padding:32px;text-align:center;">
              <h2 style="margin-bottom:8px;">Terjadi kendala saat memuat halaman</h2>
              <p class="text-muted" style="margin-bottom:16px;">Modul ini mengalami kesalahan sementara dan telah diisolasi agar tidak mengganggu modul lainnya.</p>
              <div style="display:flex;gap:8px;justify-content:center;">
                <button type="button" class="btn btn-primary" id="btn-router-retry">Muat Ulang Halaman</button>
                <a href="#/dashboard" class="btn btn-secondary" style="text-decoration:none;">Kembali ke Dashboard</a>
              </div>
            </div>
          `;
          target.querySelector('#btn-router-retry')?.addEventListener('click', () => handleRoute());
        }
      }
    } else {
      const target = main || document.getElementById('app');
      if (target) target.innerHTML = '<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>';
    }
  }

  window.addEventListener('hashchange', handleRoute);
  handleRoute();
}
