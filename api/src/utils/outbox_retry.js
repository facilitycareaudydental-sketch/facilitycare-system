/**
 * Outbox Retry Manager & Error Classifier
 *
 * Separates transient/recoverable errors from permanent rejections, auth errors,
 * duplicates, and OCC conflicts, while sanitizing all error messages before persistence.
 */

export const DEFAULT_MAX_ATTEMPTS = 5;
export const DEFAULT_BASE_DELAY_SECONDS = 30;
export const DEFAULT_MAX_DELAY_SECONDS = 3600;
export const DEFAULT_AUTH_COOLDOWN_SECONDS = 300;
export const MAX_ERROR_MESSAGE_LENGTH = 500;

export const OUTBOX_ERROR_CATEGORIES = {
  DUPLICATE_EVENT: 'DUPLICATE_EVENT',
  OCC_CONFLICT: 'OCC_CONFLICT',
  AUTH_ERROR: 'AUTH_ERROR',
  PERMANENT_CLIENT_ERROR: 'PERMANENT_CLIENT_ERROR',
  RATE_LIMIT: 'RATE_LIMIT',
  TIMEOUT: 'TIMEOUT',
  TRANSIENT_NETWORK: 'TRANSIENT_NETWORK',
  SERVER_ERROR: 'SERVER_ERROR',
  INVALID_RESPONSE: 'INVALID_RESPONSE'
};

/**
 * Strips credentials, secrets, signatures, tokens, and authorization headers
 * from error messages before storing in sync_outbox.last_error or logs.
 */
export function sanitizeOutboxError(rawInput) {
  let msg = '';
  if (rawInput instanceof Error) {
    msg = rawInput.message || rawInput.name || 'Unknown error';
  } else if (typeof rawInput === 'string') {
    msg = rawInput;
  } else if (rawInput && typeof rawInput === 'object') {
    msg = String(rawInput.errorMessage || rawInput.error || rawInput.message || JSON.stringify(rawInput));
  } else {
    msg = String(rawInput || 'Unknown error');
  }

  msg = msg
    .replace(/Bearer\s+[A-Za-z0-9\-._~+/=]+/gi, 'Bearer [REDACTED]')
    .replace(/(authorization|x-fcms-signature|api[_-]?key|secret|password|token|credential)\s*[:=]\s*[^\s,;]+/gi, '$1=[REDACTED]')
    .replace(/\beyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\b/g, '[JWT_REDACTED]')
    .replace(/\b[0-9a-f]{64}\b/gi, '[HEX64_REDACTED]')
    .replace(/[\r\n\t]+/g, ' ')
    .trim();

  if (msg.length > MAX_ERROR_MESSAGE_LENGTH) {
    return msg.slice(0, MAX_ERROR_MESSAGE_LENGTH - 3) + '...';
  }
  return msg || 'Unknown error';
}

/**
 * Classifies an adapter delivery result or thrown exception into a retry policy category.
 */
export function classifyOutboxFailure(resultOrError) {
  if (!resultOrError) {
    return {
      category: OUTBOX_ERROR_CATEGORIES.INVALID_RESPONSE,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message: 'Empty delivery response'
    };
  }

  if (resultOrError instanceof Error) {
    const code = String(resultOrError.code || resultOrError.name || '').toUpperCase();
    const msg = sanitizeOutboxError(resultOrError);
    if (code === 'ABORT_ERR' || code === 'ETIMEDOUT' || code === 'TIMEOUT' || /timeout|timed out|abort/i.test(msg)) {
      return {
        category: OUTBOX_ERROR_CATEGORIES.TIMEOUT,
        retryable: true,
        pauseDelivery: false,
        deadLetterImmediately: false,
        message: msg
      };
    }
    return {
      category: OUTBOX_ERROR_CATEGORIES.TRANSIENT_NETWORK,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message: msg
    };
  }

  const status = Number(resultOrError.status || 0);
  const errorCode = String(resultOrError.errorCode || '').toUpperCase();
  const message = sanitizeOutboxError(resultOrError.errorMessage || resultOrError.error || `HTTP ${status || 'error'}`);

  if (resultOrError.duplicate === true || errorCode === 'DUPLICATE_EVENT') {
    return {
      category: OUTBOX_ERROR_CATEGORIES.DUPLICATE_EVENT,
      retryable: false,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message: 'Duplicate event acknowledged idempotently'
    };
  }

  if (
    errorCode === 'INVALID_PAYLOAD' ||
    errorCode === 'SCHEMA_MISMATCH' ||
    errorCode === 'FORBIDDEN_FIELD' ||
    errorCode === 'SYNC_DISABLED' ||
    status === 400 ||
    status === 422
  ) {
    return {
      category: OUTBOX_ERROR_CATEGORIES.PERMANENT_CLIENT_ERROR,
      retryable: false,
      pauseDelivery: false,
      deadLetterImmediately: true,
      message
    };
  }

  if (status === 401 || status === 403 || errorCode === 'AUTH_ERROR') {
    // Pause delivery (circuit breaker / cooldown) instead of immediately dead-lettering
    // valid events during a non-atomic credential rotation window.
    return {
      category: OUTBOX_ERROR_CATEGORIES.AUTH_ERROR,
      retryable: true,
      pauseDelivery: true,
      deadLetterImmediately: false,
      message
    };
  }

  if (status === 409 || resultOrError.conflict === true || errorCode === 'OCC_CONFLICT') {
    return {
      category: OUTBOX_ERROR_CATEGORIES.OCC_CONFLICT,
      retryable: false,
      pauseDelivery: false,
      deadLetterImmediately: true,
      message
    };
  }

  if (status === 429 || errorCode === 'RATE_LIMIT') {
    return {
      category: OUTBOX_ERROR_CATEGORIES.RATE_LIMIT,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      retryAfterSeconds: Number(resultOrError.retryAfterSeconds) || null,
      message
    };
  }

  if (errorCode === 'TIMEOUT' || status === 408 || status === 504) {
    return {
      category: OUTBOX_ERROR_CATEGORIES.TIMEOUT,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message
    };
  }

  if (errorCode === 'NETWORK_ERROR') {
    return {
      category: OUTBOX_ERROR_CATEGORIES.TRANSIENT_NETWORK,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message
    };
  }

  if (errorCode === 'INVALID_RESPONSE') {
    return {
      category: OUTBOX_ERROR_CATEGORIES.INVALID_RESPONSE,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message
    };
  }

  if (status >= 500 && status <= 599) {
    return {
      category: OUTBOX_ERROR_CATEGORIES.SERVER_ERROR,
      retryable: true,
      pauseDelivery: false,
      deadLetterImmediately: false,
      message
    };
  }

  return {
    category: OUTBOX_ERROR_CATEGORIES.INVALID_RESPONSE,
    retryable: true,
    pauseDelivery: false,
    deadLetterImmediately: false,
    message
  };
}

/**
 * Computes bounded exponential backoff delay (in seconds) with optional jitter.
 */
export function computeBackoffDelaySeconds(attemptNumber, options = {}) {
  const base = Math.max(1, Number(options.baseDelaySeconds ?? DEFAULT_BASE_DELAY_SECONDS));
  const maxDelay = Math.max(base, Number(options.maxDelaySeconds ?? DEFAULT_MAX_DELAY_SECONDS));
  const exp = Math.max(0, Number(attemptNumber) - 1);
  const rawDelay = Math.min(maxDelay, base * Math.pow(2, exp));

  if (options.jitter === false) {
    return Math.round(rawDelay);
  }
  const jitterRatio = Math.min(0.5, Math.max(0, Number(options.jitterRatio ?? 0.2)));
  const factor = 1 + ((Math.random() * 2 - 1) * jitterRatio);
  return Math.max(1, Math.min(maxDelay, Math.round(rawDelay * factor)));
}

/**
 * Formats a Date as SQLite UTC datetime string ('YYYY-MM-DD HH:MM:SS') so
 * SQLite datetime('now') comparisons behave consistently.
 */
export function toSqliteDatetime(date = new Date()) {
  return date.toISOString().slice(0, 19).replace('T', ' ');
}

/**
 * Determines the next outbox state after a delivery failure or duplicate ack.
 */
export function evaluateRetryDecision(currentRetryCount, resultOrError, options = {}) {
  const maxAttempts = Math.max(1, Number(options.maxAttempts ?? DEFAULT_MAX_ATTEMPTS));
  const nowMs = Number(options.nowMs ?? Date.now());
  const classified = classifyOutboxFailure(resultOrError);

  if (classified.category === OUTBOX_ERROR_CATEGORIES.DUPLICATE_EVENT) {
    return {
      nextStatus: 'DELIVERED',
      nextRetryCount: Number(currentRetryCount || 0),
      nextRetryAt: null,
      pauseDelivery: false,
      sanitizedError: null,
      category: classified.category
    };
  }

  const nextRetryCount = Number(currentRetryCount || 0) + 1;

  if (classified.deadLetterImmediately || !classified.retryable || nextRetryCount >= maxAttempts) {
    return {
      nextStatus: 'DEAD_LETTER',
      nextRetryCount,
      nextRetryAt: null,
      pauseDelivery: classified.pauseDelivery,
      sanitizedError: `[${classified.category}] ${classified.message}`,
      category: classified.category
    };
  }

  let delaySeconds;
  if (classified.category === OUTBOX_ERROR_CATEGORIES.AUTH_ERROR) {
    delaySeconds = Math.max(60, Number(options.authCooldownSeconds ?? DEFAULT_AUTH_COOLDOWN_SECONDS));
  } else if (classified.category === OUTBOX_ERROR_CATEGORIES.RATE_LIMIT && classified.retryAfterSeconds) {
    const maxDelay = Number(options.maxDelaySeconds ?? DEFAULT_MAX_DELAY_SECONDS);
    delaySeconds = Math.min(maxDelay, Math.max(1, classified.retryAfterSeconds));
  } else {
    delaySeconds = computeBackoffDelaySeconds(nextRetryCount, options);
  }

  const nextRetryAt = toSqliteDatetime(new Date(nowMs + delaySeconds * 1000));
  return {
    nextStatus: 'RETRY',
    nextRetryCount,
    nextRetryAt,
    delaySeconds,
    pauseDelivery: classified.pauseDelivery,
    sanitizedError: `[${classified.category}] ${classified.message}`,
    category: classified.category
  };
}
