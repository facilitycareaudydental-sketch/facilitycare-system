/**
 * FCMS Google Sheets Bidirectional Sync - Utilities
 */

/**
 * Retrieve the webhook secret strictly from Google Apps Script Properties.
 * Fail-closed: never falls back to Config.js or any hardcoded literal.
 */
function getWebhookSecret() {
  if (typeof PropertiesService === 'undefined' || !PropertiesService || typeof PropertiesService.getScriptProperties !== 'function') {
    throw new Error('Script Properties service unavailable');
  }
  const props = PropertiesService.getScriptProperties();
  const secret = props && typeof props.getProperty === 'function' ? props.getProperty('WEBHOOK_SECRET') : null;
  if (!secret || typeof secret !== 'string' || secret.trim() === '') {
    throw new Error('WEBHOOK_SECRET is not configured in Script Properties');
  }
  return secret.trim();
}

const Utils = {
  getWebhookSecret: getWebhookSecret,

  /**
   * Compute HMAC SHA256 Signature for the payload
   */
  computeSignature: function(payloadString, secret) {
    const signatureBytes = Utilities.computeHmacSha256Signature(payloadString, secret);
    // Convert byte array to hex string
    return signatureBytes.map(function(byte) {
      const v = (byte < 0 ? byte + 256 : byte);
      return ("0" + v.toString(16)).slice(-2);
    }).join('');
  },

  /**
   * Generate a UUID V4
   */
  generateUUID: function() {
    return Utilities.getUuid();
  }
};
