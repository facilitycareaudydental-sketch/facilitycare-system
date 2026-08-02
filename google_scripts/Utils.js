/**
 * FCMS Google Sheets Bidirectional Sync - Utilities
 */

const Utils = {
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
