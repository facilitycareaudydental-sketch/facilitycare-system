// Hash-based router
const routes = {};
let currentCleanup = null;

export function registerRoute(path, handler) {
  routes[path] = handler;
}

export function navigate(path) {
  window.location.hash = path;
}

export function initRouter() {
  async function handleRoute() {
    const hash = window.location.hash.replace('#', '') || '/dashboard';
    const [base, ...rest] = hash.split('?');

    // Find matching route (exact or prefix)
    let handler = routes[base];
    if (!handler) {
      // Try prefix match for dynamic segments
      for (const [pattern, h] of Object.entries(routes)) {
        if (pattern.endsWith('/*') && base.startsWith(pattern.slice(0, -2))) {
          handler = h;
          break;
        }
      }
    }

    if (currentCleanup) { currentCleanup(); currentCleanup = null; }

    const main = document.getElementById('main-content');
    if (!main) return;
    main.innerHTML = '<div class="loading-spinner"><div class="spinner"></div></div>';

    if (handler) {
      const params = new URLSearchParams(rest.join('?'));
      const segments = base.split('/').filter(Boolean);
      const cleanup = await handler({ path: base, params, segments, main });
      if (cleanup) currentCleanup = cleanup;
    } else {
      main.innerHTML = '<div class="empty-state"><h2>404 - Halaman tidak ditemukan</h2></div>';
    }
  }

  window.addEventListener('hashchange', handleRoute);
  handleRoute();
}
