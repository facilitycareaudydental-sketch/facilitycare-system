/**
 * FCMS Google Sheets Bidirectional Sync - Configuration
 */

const CONFIG = {
  API_URL: 'https://fm-operations-api.facilitycare-audydental.workers.dev/api/sync/webhook',
  // WEBHOOK_SECRET must NEVER be hardcoded here; configure via Script Properties ('WEBHOOK_SECRET')
  PAYLOAD_VERSION: '1.0',
  IGNORED_SHEETS: ['Dashboard', 'Dropdowns', 'Log']
};
