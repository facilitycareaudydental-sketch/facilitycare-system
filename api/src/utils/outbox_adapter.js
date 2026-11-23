/**
 * Outbox Delivery Adapters (Local Fixture & Dry-Run Simulation Only)
 *
 * Provides clean adapter boundaries decoupled from business mutations.
 * STRICT LOCAL-ONLY: Never sends outbound HTTP requests to production endpoints.
 */

import { validateOutboundPayload } from './sync_mapper.js';
import { sanitizeOutboxError } from './outbox_retry.js';

export async function signOutboundBody(secret, rawBody) {
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw',
    encoder.encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, encoder.encode(rawBody));
  return [...new Uint8Array(sig)].map((b) => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Builds the canonical outbound envelope from a sync_outbox row.
 */
export function buildOutboundEnvelope(outboxRow) {
  let parsedData;
  try {
    parsedData = typeof outboxRow.payload === 'string'
      ? JSON.parse(outboxRow.payload)
      : outboxRow.payload;
  } catch {
    return { valid: false, errorCode: 'INVALID_PAYLOAD', errorMessage: 'Malformed JSON in outbox payload' };
  }

  const rowVersion = parsedData && parsedData.ROW_VERSION !== undefined && parsedData.ROW_VERSION !== ''
    ? Number(parsedData.ROW_VERSION)
    : 1;

  const envelope = {
    payload_version: '1.0',
    event_id: outboxRow.id,
    entity_name: outboxRow.entity_name,
    entity_id: String(outboxRow.entity_id),
    action: outboxRow.action,
    source: 'FCMS',
    row_version: Number.isInteger(rowVersion) && rowVersion >= 0 ? rowVersion : 1,
    timestamp: new Date().toISOString(),
    correlation_id: outboxRow.correlation_id || null,
    data: parsedData
  };

  const validation = validateOutboundPayload(outboxRow.entity_name, parsedData, envelope);
  if (!validation.valid) {
    return {
      valid: false,
      errorCode: validation.code || 'SCHEMA_MISMATCH',
      errorMessage: validation.error || 'Payload validation failed'
    };
  }

  return { valid: true, envelope };
}

/**
 * Creates a Local Google Sheets Mock/HTTP-Fixture Adapter.
 * Requires an injected `transport(requestDescriptor)` function and never calls
 * production Google Apps Script endpoints.
 */
export function createLocalSheetsAdapter(options = {}) {
  const {
    transport,
    endpoint = 'http://local-fixture.test/gas-outbound',
    dummySecret = 'local-fixture-dummy-secret-not-prod',
    timeoutMs = 5000
  } = options;

  if (typeof transport !== 'function') {
    throw new Error('LocalSheetsAdapter requires an explicit mock/local transport function');
  }

  return {
    name: 'local-sheets-fixture',
    isLocalOnly: true,
    async deliver(outboxRow) {
      const built = buildOutboundEnvelope(outboxRow);
      if (!built.valid) {
        return {
          ok: false,
          status: 400,
          errorCode: built.errorCode,
          errorMessage: built.errorMessage
        };
      }

      const rawBody = JSON.stringify(built.envelope);
      const signature = await signOutboundBody(dummySecret, rawBody);

      const requestDescriptor = {
        url: endpoint,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${dummySecret}`,
          'X-FCMS-SIGNATURE': signature
        },
        body: rawBody,
        envelope: built.envelope
      };

      let timerId = null;
      try {
        const response = await Promise.race([
          Promise.resolve(transport(requestDescriptor)),
          new Promise((_, reject) => {
            timerId = setTimeout(() => {
              const err = new Error(`Local adapter timeout after ${timeoutMs}ms`);
              err.code = 'TIMEOUT';
              reject(err);
            }, timeoutMs);
          })
        ]);

        if (!response || typeof response !== 'object') {
          return { ok: false, status: 502, errorCode: 'INVALID_RESPONSE', errorMessage: 'Non-object transport response' };
        }

        const status = Number(response.status ?? (response.ok ? 200 : 500));
        let body = response.body;
        if (typeof body === 'string') {
          try {
            body = JSON.parse(body);
          } catch {
            return { ok: false, status: 502, errorCode: 'INVALID_RESPONSE', errorMessage: 'Invalid JSON response from Sheets adapter' };
          }
        }

        if (status >= 200 && status < 300) {
          if (!body || typeof body !== 'object' || body.ok !== true) {
            return { ok: false, status: 502, errorCode: 'INVALID_RESPONSE', errorMessage: 'Missing ok:true confirmation in 2xx response' };
          }
          return {
            ok: true,
            status,
            duplicate: body.duplicate === true,
            envelope: built.envelope,
            body
          };
        }

        if (status === 409 && body?.duplicate === true) {
          return { ok: true, status: 200, duplicate: true, envelope: built.envelope, body };
        }

        return {
          ok: false,
          status,
          conflict: status === 409,
          retryAfterSeconds: response.retryAfterSeconds ?? body?.retryAfterSeconds ?? null,
          errorCode: body?.errorCode || (status === 429 ? 'RATE_LIMIT' : undefined),
          errorMessage: sanitizeOutboxError(body?.error || body?.message || `HTTP ${status}`)
        };
      } catch (err) {
        const msg = sanitizeOutboxError(err);
        const code = String(err?.code || '').toUpperCase();
        return {
          ok: false,
          status: 0,
          errorCode: code === 'TIMEOUT' || /timeout/i.test(msg) ? 'TIMEOUT' : 'NETWORK_ERROR',
          errorMessage: msg
        };
      } finally {
        if (timerId !== null) clearTimeout(timerId);
      }
    }
  };
}

/**
 * Creates a DRY_RUN simulation adapter that validates payloads locally without
 * executing any transport or database mutation.
 */
export function createDryRunSimulationAdapter() {
  return {
    name: 'dry-run-simulation',
    isLocalOnly: true,
    async deliver(outboxRow) {
      const built = buildOutboundEnvelope(outboxRow);
      if (!built.valid) {
        return {
          ok: false,
          status: 400,
          errorCode: built.errorCode,
          errorMessage: built.errorMessage
        };
      }
      return {
        ok: true,
        status: 200,
        dryRun: true,
        envelope: built.envelope,
        body: { ok: true, simulated: true }
      };
    }
  };
}
