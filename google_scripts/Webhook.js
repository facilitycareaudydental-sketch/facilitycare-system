/**
 * FCMS Google Sheets Bidirectional Sync - Webhook Logic
 */

function sendWebhook(payloadObj) {
  const payloadString = JSON.stringify(payloadObj);
  const signature = Utils.computeSignature(payloadString, CONFIG.WEBHOOK_SECRET);

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${CONFIG.WEBHOOK_SECRET}`,
      'X-FCMS-SIGNATURE': signature
    },
    payload: payloadString,
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(CONFIG.API_URL, options);
    if (response.getResponseCode() !== 200) {
      console.error('Webhook failed:', response.getContentText());
      // Logic for Retry Queue could be placed here
    }
  } catch (err) {
    console.error('Fetch error:', err.message);
  }
}
