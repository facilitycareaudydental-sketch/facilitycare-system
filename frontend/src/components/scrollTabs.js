/**
 * scrollTabs.js — Global Clean Navy Horizontal Scrollable Tabs & Pills Controller
 * Enhances any horizontal tab/filter-pill strip with accessible left/right scroll buttons,
 * smooth mouse-wheel / touchpad / touch scrolling, edge fade hints, and ResizeObserver support.
 */

const TAB_STRIP_SELECTORS = [
  '[data-fc-scroll-tabs]',
  '#hygiene-zone-tabs',
  '#doc-filter-pills',
  '.emp-filter-pills',
  '.con-filter-pills',
  '.sp-filter-pills',
  '.mutasi-filter-pills',
  '.lembur-filter-pills',
  '.issues-filter-pills',
  '.training-filter-pills',
  '.fc-nav-tabs-strip',
  '#cat-pills-wrap',
  '#sec-cat-pills-wrap',
].join(', ');

const CHEVRON_LEFT = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg>`;
const CHEVRON_RIGHT = `<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="9 18 15 12 9 6"/></svg>`;

export function enhanceScrollTabs(trackEl) {
  if (!trackEl || trackEl.dataset.fcScrollEnhanced === '1') return;
  if (!trackEl.parentElement) return;

  // Avoid wrapping if already inside a shell
  if (trackEl.parentElement.classList.contains('fc-scroll-tabs-shell')) {
    trackEl.dataset.fcScrollEnhanced = '1';
    return;
  }

  trackEl.dataset.fcScrollEnhanced = '1';
  trackEl.classList.add('fc-scroll-tabs-track');

  const shell = document.createElement('div');
  shell.className = 'fc-scroll-tabs-shell';

  const prevBtn = document.createElement('button');
  prevBtn.type = 'button';
  prevBtn.className = 'fc-scroll-tabs-btn fc-scroll-tabs-btn--prev';
  prevBtn.setAttribute('aria-label', 'Geser daftar tab ke kiri');
  prevBtn.title = 'Geser ke kiri';
  prevBtn.innerHTML = CHEVRON_LEFT;

  const nextBtn = document.createElement('button');
  nextBtn.type = 'button';
  nextBtn.className = 'fc-scroll-tabs-btn fc-scroll-tabs-btn--next';
  nextBtn.setAttribute('aria-label', 'Geser daftar tab ke kanan');
  nextBtn.title = 'Geser ke kanan';
  nextBtn.innerHTML = CHEVRON_RIGHT;

  const parent = trackEl.parentElement;
  parent.insertBefore(shell, trackEl);
  shell.appendChild(prevBtn);
  shell.appendChild(trackEl);
  shell.appendChild(nextBtn);

  const updateControls = () => {
    if (!trackEl.isConnected) return;
    const maxScroll = Math.max(0, trackEl.scrollWidth - trackEl.clientWidth);
    const hasOverflow = maxScroll > 4;
    const left = trackEl.scrollLeft;

    shell.classList.toggle('is-overflowing', hasOverflow);
    const canLeft = hasOverflow && left > 4;
    const canRight = hasOverflow && left < maxScroll - 4;

    shell.classList.toggle('can-scroll-left', canLeft);
    shell.classList.toggle('can-scroll-right', canRight);

    prevBtn.disabled = !canLeft;
    nextBtn.disabled = !canRight;
  };

  const scrollStep = (dir) => {
    const step = Math.max(180, Math.round(trackEl.clientWidth * 0.65));
    trackEl.scrollBy({ left: dir * step, behavior: 'smooth' });
  };

  prevBtn.addEventListener('click', (e) => {
    e.preventDefault();
    scrollStep(-1);
  });
  nextBtn.addEventListener('click', (e) => {
    e.preventDefault();
    scrollStep(1);
  });

  // Convert vertical wheel to horizontal scroll when hovering overflowing strip
  trackEl.addEventListener('wheel', (e) => {
    if (trackEl.scrollWidth <= trackEl.clientWidth + 4) return;
    if (Math.abs(e.deltaY) > Math.abs(e.deltaX)) {
      const prevLeft = trackEl.scrollLeft;
      trackEl.scrollLeft += e.deltaY;
      if (trackEl.scrollLeft !== prevLeft) {
        e.preventDefault();
      }
    }
  }, { passive: false });

  trackEl.addEventListener('scroll', updateControls, { passive: true });

  // Recalculate on resize or child mutation (e.g. dynamic pills load or sidebar collapse)
  if (typeof ResizeObserver !== 'undefined') {
    const ro = new ResizeObserver(() => requestAnimationFrame(updateControls));
    ro.observe(shell);
    ro.observe(trackEl);
  }
  if (typeof MutationObserver !== 'undefined') {
    const mo = new MutationObserver(() => requestAnimationFrame(updateControls));
    mo.observe(trackEl, { childList: true, subtree: false });
  }
  window.addEventListener('resize', updateControls, { passive: true });

  // Initial & deferred measurements after layout/fonts settle
  requestAnimationFrame(updateControls);
  setTimeout(updateControls, 250);
  setTimeout(updateControls, 700);
}

export function scanAndEnhanceScrollTabs(root = document) {
  if (!root || !root.querySelectorAll) return;
  root.querySelectorAll(TAB_STRIP_SELECTORS).forEach((el) => {
    enhanceScrollTabs(el);
  });
}

let _globalObserver = null;
export function initGlobalScrollTabsObserver(rootEl) {
  if (!rootEl) return;
  if (_globalObserver) {
    _globalObserver.disconnect();
    _globalObserver = null;
  }
  scanAndEnhanceScrollTabs(rootEl);

  let scheduled = false;
  _globalObserver = new MutationObserver(() => {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      scanAndEnhanceScrollTabs(rootEl);
    });
  });
  _globalObserver.observe(rootEl, { childList: true, subtree: true });
}
