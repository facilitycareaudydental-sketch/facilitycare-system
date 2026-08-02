/**
 * FCMS Google Sheets Bidirectional Sync - Triggers
 * 
 * Run `setupTriggers()` once from the editor to initialize.
 */

function onEdit(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  
  if (CONFIG.IGNORED_SHEETS.includes(sheetName)) return;

  const range = e.range;
  const row = range.getRow();
  
  // Ignore header edits
  if (row === 1) return;

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
      lastSyncSourceColIdx = i + 1; // 1-indexed
    }
  }

  // Anti-Loop Check: Rely strictly on LAST_SYNC_SOURCE metadata instead of BOT_EMAIL
  if (payloadData['LAST_SYNC_SOURCE'] === 'FCMS') {
    // This edit was made by the FCMS D1 -> Sheets sync pushing data here.
    // Do not bounce it back.
    return;
  }

  const action = fcmsId ? 'UPDATE' : 'INSERT';

  // Overwrite LAST_SYNC_SOURCE to 'SHEET'
  if (lastSyncSourceColIdx > 0) {
    sheet.getRange(row, lastSyncSourceColIdx).setValue('SHEET');
  }

  const payload = {
    payload_version: CONFIG.PAYLOAD_VERSION,
    event_id: Utils.generateUUID(),
    sheet_name: sheetName,
    action: action,
    source: 'SHEET',
    user_email: e.user ? e.user.getEmail() : Session.getActiveUser().getEmail(),
    timestamp: new Date().toISOString(),
    data: payloadData
  };

  sendWebhook(payload);
}

function onChange(e) {
  if (e.changeType === 'REMOVE_ROW') {
    // Handle row deletion logic here
  }
}

function setupTriggers() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger('onEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChange').forSpreadsheet(ss).onChange().create();
    
  SpreadsheetApp.getUi().alert('FCMS Sync Triggers Installed!');
}
