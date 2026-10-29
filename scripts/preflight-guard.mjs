/**
 * preflight-guard.mjs — Deployment Target Guard for FCMS (Facility Care Management System)
 * ────────────────────────────────────────────────────────────────────────────────────────
 * Fail-Closed verification gate: ensures deployment commands never accidentally target
 * FCVM Cloudflare Account, FCVM D1 databases, or FCVM Workers.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');

const REQUIRED_FCMS_ACCOUNT_ID = '293440e04e5d0c8073e7e2893750b527';
const REQUIRED_FCMS_D1_NAME = 'fcms-production-db';
const REQUIRED_FCMS_D1_ID = '62b4b6e5-a718-40bc-9aab-6a700b82a9c9';

const FORBIDDEN_STRINGS = [
  { id: 'a41392e52e3f79e56af325ddf27a6ed4', desc: 'FCVM Cloudflare Account ID' },
  { id: '1b9822f6-e92f-4cff-9ef1-5d1c7fe27196', desc: 'FCVM Production D1 UUID' },
  { id: '91abff4a-dbe8-4605-9f6f-3e24dd022b11', desc: 'FCVM OLD D1 UUID' },
  { id: 'facilitycare-fcvm-api', desc: 'FCVM Worker name' },
  { id: 'fcvm-temp', desc: 'FCVM worker subdomain' },
];

console.log('🛡️  [FCMS Preflight Guard] Verifying deployment target configuration...');

let hasErrors = false;

function checkFile(relPath, requiredPairs = []) {
  const fullPath = path.join(rootDir, relPath);
  if (!fs.existsSync(fullPath)) {
    console.error(`❌ [GUARD] Required configuration file missing: ${relPath}`);
    hasErrors = true;
    return;
  }

  const content = fs.readFileSync(fullPath, 'utf8');

  // Check forbidden strings
  for (const forbidden of FORBIDDEN_STRINGS) {
    if (content.toLowerCase().includes(forbidden.id.toLowerCase())) {
      console.error(`❌ [GUARD FATAL] ${relPath} contains forbidden identifier: ${forbidden.desc} (${forbidden.id})`);
      hasErrors = true;
    }
  }

  // Check required key-value patterns
  for (const { key, value, pattern } of requiredPairs) {
    const regex = pattern || new RegExp(`^[\\s]*${key}[\\s]*=[\\s]*["']?${value}["']?`, 'm');
    if (!regex.test(content)) {
      console.error(`❌ [GUARD FATAL] ${relPath} is missing required setting: ${key} = "${value}"`);
      hasErrors = true;
    }
  }
}

// 1. Verify root wrangler.toml
checkFile('wrangler.toml', [
  { key: 'account_id', value: REQUIRED_FCMS_ACCOUNT_ID },
  { key: 'database_name', value: REQUIRED_FCMS_D1_NAME },
  { key: 'database_id', value: REQUIRED_FCMS_D1_ID },
]);

// 2. Verify api/wrangler.toml
checkFile('api/wrangler.toml', [
  { key: 'account_id', value: REQUIRED_FCMS_ACCOUNT_ID },
  { key: 'database_name', value: REQUIRED_FCMS_D1_NAME },
  { key: 'database_id', value: REQUIRED_FCMS_D1_ID },
]);

// 3. Verify frontend/wrangler.toml
checkFile('frontend/wrangler.toml', [
  { key: 'account_id', value: REQUIRED_FCMS_ACCOUNT_ID },
]);

// 4. Check Environment Variables (if set)
const envAccountId = process.env.CLOUDFLARE_ACCOUNT_ID ? process.env.CLOUDFLARE_ACCOUNT_ID.trim() : undefined;
if (envAccountId && envAccountId !== REQUIRED_FCMS_ACCOUNT_ID) {
  console.error(`❌ [GUARD FATAL] Environment CLOUDFLARE_ACCOUNT_ID (${envAccountId}) does not match FCMS (${REQUIRED_FCMS_ACCOUNT_ID})`);
  hasErrors = true;
}

// 5. Check cached wrangler files if present
const cachedAccountFiles = [
  path.join(rootDir, 'node_modules', '.cache', 'wrangler', 'wrangler-account.json'),
  path.join(rootDir, 'api', 'node_modules', '.cache', 'wrangler', 'wrangler-account.json')
];
for (const cacheFile of cachedAccountFiles) {
  if (fs.existsSync(cacheFile)) {
    try {
      const data = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
      if (data.account?.id && data.account.id !== REQUIRED_FCMS_ACCOUNT_ID) {
        console.error(`❌ [GUARD FATAL] Stale wrangler cache file detected (${cacheFile}) pointing to wrong account: ${data.account.id}.`);
        console.error(`   Please delete the file or let the guard clear it.`);
        hasErrors = true;
      }
    } catch {}
  }
}

if (hasErrors) {
  console.error('\n🚨 [FCMS Preflight Guard] DEPLOYMENT BLOCKED. Configuration failed isolation audit.\n');
  process.exit(1);
} else {
  console.log('✅ [FCMS Preflight Guard] All target assertions PASSED: Account ID, D1 UUID, and isolation verified.\n');
  process.exit(0);
}
