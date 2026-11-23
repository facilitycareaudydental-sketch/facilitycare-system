/**
 * Centralized Frontend XSS Sanitization & Safe URL Utilities (Phase 5.66.19)
 *
 * Works across both Browser DOM and Node.js unit test environments.
 */

const DANGEROUS_SCHEME_REGEX = /^\s*(?:javascript|vbscript|data|file)\s*:/i;
const CONTROL_WHITESPACE_REGEX = /[\u0000-\u001f\u007f\s]+/g;

/**
 * Escapes all HTML-significant characters for safe text/attribute interpolation.
 */
export function escapeHtml(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/`/g, '&#96;');
}

/**
 * Validates URLs and strips dangerous URI schemes (javascript:, vbscript:, data:, file:),
 * including control-character/whitespace obfuscated variants.
 */
export function sanitizeUrl(rawUrl, fallback = '#') {
  if (rawUrl === null || rawUrl === undefined) return fallback;
  const trimmed = String(rawUrl).trim();
  if (!trimmed) return fallback;

  const collapsed = trimmed.replace(CONTROL_WHITESPACE_REGEX, '');
  if (DANGEROUS_SCHEME_REGEX.test(collapsed)) {
    return fallback;
  }

  // Permit relative paths/fragments and explicit http/https/mailto/tel
  if (/^(?:#|\/(?!\/)|\.{1,2}\/|https?:\/\/|mailto:|tel:)/i.test(trimmed)) {
    return trimmed;
  }

  return fallback;
}

/**
 * Safely sets plain text content on a DOM node without HTML parsing.
 */
export function setSafeText(el, value) {
  if (!el) return;
  el.textContent = value === null || value === undefined ? '' : String(value);
}

/**
 * Sanitizes an HTML string fragment by removing active/executable tags,
 * inline event handlers, and unsafe URL attributes.
 */
export function sanitizeHtmlFragment(rawHtml) {
  if (rawHtml === null || rawHtml === undefined) return '';
  let str = String(rawHtml);

  // Strip dangerous containers completely (including contents for script/style/svg/iframe/object)
  str = str.replace(/<\s*(script|style|iframe|object|embed|applet|svg|math|meta|link|base)\b[^>]*>[\s\S]*?<\s*\/\s*\1\s*>/gi, '');
  // Strip self-closing or unclosed dangerous tags
  str = str.replace(/<\s*\/?\s*(script|style|iframe|object|embed|applet|svg|math|meta|link|base)\b[^>]*\/?>/gi, '');

  // Remove any inline event handlers (onmouseover=..., onclick=..., onerror=...)
  str = str.replace(/\s+on[a-z0-9_-]+\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '');
  // Remove srcdoc and formaction
  str = str.replace(/\s+(?:srcdoc|formaction)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '');

  // Neutralize dangerous schemes inside href / src / action / xlink:href / poster
  str = str.replace(
    /(\b(?:href|src|action|xlink:href|poster)\s*=\s*)(?:"([^"]*)"|'([^']*)'|([^\s>]+))/gi,
    (_, prefix, dq, sq, bare) => {
      const val = dq ?? sq ?? bare ?? '';
      const safe = sanitizeUrl(val, '#');
      return `${prefix}"${escapeHtml(safe)}"`;
    }
  );

  return str;
}
