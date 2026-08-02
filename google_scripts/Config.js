/**
 * FCMS Google Sheets Bidirectional Sync - Configuration
 */

const CONFIG = {
  API_URL: 'https://fm-operations-api.facilitycare-audydental.workers.dev/api/sync/webhook',
  WEBHOOK_SECRET: 'YOUR_WEBHOOK_SECRET_HERE',
  PAYLOAD_VERSION: '1.0',
  IGNORED_SHEETS: ['Dashboard', 'Dropdowns', 'Log']
};
