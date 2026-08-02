/**
 * FCMS Google Sheets Bidirectional Sync Bridge
 * 
 * Instructions:
 * 1. Open Google Sheets -> Extensions -> Apps Script
 * 2. Paste this code
 * 3. Add your FCMS_API_URL and WEBHOOK_SECRET
 * 4. Run `setupTriggers()` once.
 */

const FCMS_API_URL = 'https://fm-operations-api.facilitycare-audydental.workers.dev/api/sync/webhook';
const WEBHOOK_SECRET = 'YOUR_WEBHOOK_SECRET_HERE'; // Set this in Cloudflare and here
const BOT_EMAIL = 'fcms-bot@your-gcp-project.iam.gserviceaccount.com'; // Change to actual service account email

/**
 * Triggered automatically when a user edits a cell
 */
function onEdit(e) {
  if (!e) return;
  
  // 1. Anti-Loop Check: Was this edit made by our Cloudflare Bot?
  const activeUser = e.user ? e.user.getEmail() : Session.getActiveUser().getEmail();
  if (activeUser === BOT_EMAIL) {
    // This is an echo from FCMS pushing to Sheets. Do nothing.
    return;
  }

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  // Example sheets to ignore
  const ignoredSheets = ['Dashboard', 'Dropdowns', 'Log'];
  if (ignoredSheets.includes(sheetName)) return;

  const range = e.range;
  const row = range.getRow();
  
  // Ignore header edits
  if (row === 1) return;

  // Get full row data
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowDataArray = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const payloadData = {};
  let fcmsId = null;
  let lastSyncSourceColIdx = -1;

  for (let i = 0; i < headers.length; i++) {
    const colName = headers[i];
    if (!colName) continue;
    
    payloadData[colName] = rowDataArray[i];
    
    if (colName === 'FCMS_ID') {
      fcmsId = rowDataArray[i];
    }
    if (colName === 'LAST_SYNC_SOURCE') {
      lastSyncSourceColIdx = i + 1; // 1-indexed for Apps Script
    }
  }

  // 2. Further Anti-Loop Check (Metadata Propagation)
  if (payloadData['LAST_SYNC_SOURCE'] === 'FCMS') {
    // If we just detected a change and the source was FCMS, it means we are in the middle of a bot update.
    return;
  }

  // Determine Action
  const action = fcmsId ? 'UPDATE' : 'INSERT';

  // Overwrite LAST_SYNC_SOURCE to 'SHEET'
  if (lastSyncSourceColIdx > 0) {
    sheet.getRange(row, lastSyncSourceColIdx).setValue('SHEET');
  }

  const payload = {
    event_id: Utilities.getUuid(),
    sheet_name: sheetName,
    action: action,
    source: 'SHEET',
    user_email: activeUser,
    timestamp: new Date().toISOString(),
    data: payloadData
  };

  sendWebhook(payload);
}

/**
 * Triggered on row deletion or addition
 */
function onChange(e) {
  // To handle DELETE, we need to track cached versions of sheets.
  // For Phase 1, we just set up the stub.
  if (e.changeType === 'REMOVE_ROW') {
    // Extract deleted FCMS_ID from cache
    // sendWebhook({ action: 'DELETE' ... })
  }
}

/**
 * Send HTTP POST to FCMS Cloudflare API
 */
function sendWebhook(payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${WEBHOOK_SECRET}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(FCMS_API_URL, options);
    if (response.getResponseCode() !== 200) {
      console.error('Webhook failed:', response.getContentText());
      // Here we could implement a Retry Queue (writing to a hidden sheet)
    }
  } catch (err) {
    console.error('Fetch error:', err.message);
  }
}

/**
 * Run this once manually to establish triggers
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActive();
  
  // Clear old triggers
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  ScriptApp.newTrigger('onChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
    
  SpreadsheetApp.getUi().alert('FCMS Sync Triggers Installed!');
}
