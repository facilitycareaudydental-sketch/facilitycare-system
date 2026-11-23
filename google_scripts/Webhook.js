/**
 * FCMS Google Sheets Bidirectional Sync - Webhook Logic
 */

function sendWebhook(payloadObj) {
  let webhookSecret;
  try {
    webhookSecret = getWebhookSecret();
  } catch (err) {
    console.error('Webhook aborted (fail-closed):', err && err.message ? err.message : 'Missing webhook secret');
    return { ok: false, error: 'Webhook secret not configured' };
  }

  const payloadString = JSON.stringify(payloadObj);
  const signature = Utils.computeSignature(payloadString, webhookSecret);

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${webhookSecret}`,
      'X-FCMS-SIGNATURE': signature
    },
    payload: payloadString,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    const statusCode = response.getResponseCode();
    if (statusCode !== 200) {
      console.error('Webhook failed with status:', statusCode);
      return { ok: false, status: statusCode };
    }
    return { ok: true, status: statusCode };
  } catch (err) {
    console.error('Fetch error:', err && err.message ? err.message : 'Network error');
    return { ok: false, error: 'Fetch error' };
  }
}
