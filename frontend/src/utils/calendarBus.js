/**
 * calendarBus.js — Global Event Bus for Real-Time Calendar Sync
 *
 * Architecture: Lightweight pub/sub singleton.
 * Any module that mutates date-sensitive data emits an event here.
 * The calendar page subscribes and auto-refreshes immediately.
 *
 * Usage:
 *   import { calendarBus } from '../utils/calendarBus.js';
 *   calendarBus.emit('data:changed', { module: 'schedule' }); // from _crud.js
 *   calendarBus.on('data:changed', handler);                  // from calendar.js
 *   calendarBus.off('data:changed', handler);                 // on page unmount
 */

const _listeners = {};

export const calendarBus = {
  /**
   * Subscribe to an event.
   * @param {string} event
   * @param {Function} handler
   */
  on(event, handler) {
    if (!_listeners[event]) _listeners[event] = new Set();
    _listeners[event].add(handler);
  },

  /**
   * Unsubscribe from an event. Call this when the calendar page unmounts
   * to prevent memory leaks and phantom re-renders.
   * @param {string} event
   * @param {Function} handler
   */
  off(event, handler) {
    if (_listeners[event]) _listeners[event].delete(handler);
  },

  /**
   * Emit an event with optional payload. All subscribers are called synchronously.
   * @param {string} event
   * @param {*} payload
   */
  emit(event, payload) {
    if (_listeners[event]) {
      _listeners[event].forEach(fn => {
        try { fn(payload); } catch (e) { console.warn('[calendarBus] Handler error:', e); }
      });
    }
  },

  /**
   * Remove all listeners — useful for testing or full app reset.
   */
  clear() {
    Object.keys(_listeners).forEach(k => delete _listeners[k]);
  }
};

/**
 * Convenience: modules call this after any mutation (add/edit/delete/import/bulk).
 * Emits 'data:changed' with the module name so calendar can decide whether to refresh.
 *
 * Calendar-relevant modules (trigger re-render):
 *   schedule, cleaning, inspection, fogging, reliever, contract, issue,
 *   training, one_on_one, sp, mutasi, basecamp, supply
 *
 * Non-calendar modules (ignored by calendar):
 *   sop, checklist, forms, employees, branches, users
 */
const CALENDAR_MODULES = new Set([
  'schedule', 'cleaning', 'cleaning_reports',
  'inspection', 'inspection_reports',
  'fogging', 'fogging_reports',
  'reliever', 'relievers',
  'contract', 'contracts',
  'issue', 'issues',
  'training',
  'one_on_one',
  'sp', 'sp_data',
  'mutasi',
  'basecamp', 'basecamp_reports',
  'supply',
]);

export function notifyCalendar(module) {
  if (!module) { calendarBus.emit('data:changed', { module: 'unknown' }); return; }
  const normalized = String(module).toLowerCase().replace(/^\/api\//, '').replace(/^reports\//, '');
  calendarBus.emit('data:changed', { module: normalized, relevant: CALENDAR_MODULES.has(normalized) });
}
