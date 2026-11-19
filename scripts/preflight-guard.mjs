/**
 * preflight-guard.mjs — Deployment Target Guard for FCMS (Facility Care Management System)
 * ────────────────────────────────────────────────────────────────────────────────────────
 * Fail-Closed verification gate: ensures deployment commands never accidentally target
 * FCVM Cloudflare Account, FCVM D1 databases, or FCVM Workers.
 *
 * Hardened in Stage N7B.4.1-B:
 * - Requires explicit environment (--env=production or --env=staging).
 * - Enforces strict fail-closed policy (zero default fallback).
 * - Strictly isolates Staging from Production UUIDs and FCVM identifiers.
 * - Preserves 100% of existing Production assertions.
 */

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');

export const REQUIRED_FCMS_ACCOUNT_ID = '293440e04e5d0c8073e7e2893750b527';
export const REQUIRED_FCMS_D1_NAME = 'fcms-production-db';
export const REQUIRED_FCMS_D1_ID = '62b4b6e5-a718-40bc-9aab-6a700b82a9c9';

export const FORBIDDEN_STRINGS = [
  { id: 'a41392e52e3f79e56af325ddf27a6ed4', desc: 'FCVM Cloudflare Account ID' },
  { id: '1b9822f6-e92f-4cff-9ef1-5d1c7fe27196', desc: 'FCVM Production D1 UUID' },
  { id: '91abff4a-dbe8-4605-9f6f-3e24dd022b11', desc: 'FCVM OLD D1 UUID' },
  { id: 'facilitycare-fcvm-api', desc: 'FCVM Worker name' },
  { id: 'facilitycare-fcvm-db', desc: 'FCVM D1 database name' },
  { id: 'fcvm-temp', desc: 'FCVM worker subdomain' },
  { id: 'fcvm', desc: 'FCVM generic token' }
];

/**
 * Functional verification engine (Testable without exiting process)
 */
export function runPreflightVerification(options = {}) {
  const errors = [];
  const warnings = [];

  const rawEnv = (options.env || options.environment || '').toLowerCase().trim();

  // 1. Explicit Environment Policy (No default fallback allowed)
  if (!rawEnv) {
    errors.push('Target environment must be explicitly specified (--env=production or --env=staging). No default fallback permitted.');
    return { success: false, errors, warnings };
  }

  if (rawEnv !== 'production' && rawEnv !== 'staging') {
    errors.push(`Unauthorized deployment environment: '${rawEnv}'. Permitted environments are 'production' and 'staging'.`);
    return { success: false, errors, warnings };
  }

  // 1.1 Mandatory Production Authorization Enforcement (Strict Positive Boolean Only)
  if (rawEnv === 'production') {
    if (options.authorizeProduction !== true) {
      errors.push('PRODUCTION_RELEASE_UNAUTHORIZED: Deployment to Production requires explicit authorization (--authorize-production or env FCMS_PRODUCTION_RELEASE_AUTHORIZED=true). options.authorizeProduction must be boolean true.');
    }
  }

  const baseDir = options.rootDir || rootDir;

  function checkFileContent(relPath, requiredPairs = [], forbiddenList = FORBIDDEN_STRINGS) {
    const fullPath = path.join(baseDir, relPath);
    if (!fs.existsSync(fullPath)) {
      errors.push(`Required configuration file missing: ${relPath}`);
      return;
    }

    const content = fs.readFileSync(fullPath, 'utf8');

    // Scan for forbidden strings
    for (const forbidden of forbiddenList) {
      const isMatch = forbidden.pattern
        ? forbidden.pattern.test(content)
        : content.toLowerCase().includes(forbidden.id.toLowerCase());
      if (isMatch) {
        errors.push(`${relPath} contains forbidden identifier: ${forbidden.desc} (${forbidden.id})`);
      }
    }

    // Scan for required pairs
    for (const { key, value, pattern } of requiredPairs) {
      const regex = pattern || new RegExp(`^[\\s]*${key}[\\s]*=[\\s]*["']?${value}["']?`, 'm');
      if (!regex.test(content)) {
        errors.push(`${relPath} is missing required setting: ${key} = "${value}"`);
      }
    }
  }

  // 2. PRODUCTION ENVIRONMENT AUDIT
  if (rawEnv === 'production') {
    // A. Verify root wrangler.toml
    checkFileContent('wrangler.toml', [
      { key: 'account_id', value: REQUIRED_FCMS_ACCOUNT_ID },
      { key: 'database_name', value: REQUIRED_FCMS_D1_NAME },
      { key: 'database_id', value: REQUIRED_FCMS_D1_ID },
      { key: 'ENVIRONMENT', value: 'production' },
    ]);

    // B. Verify api/wrangler.toml
    checkFileContent('api/wrangler.toml', [
      { key: 'account_id', value: REQUIRED_FCMS_ACCOUNT_ID },
      { key: 'database_name', value: REQUIRED_FCMS_D1_NAME },
      { key: 'database_id', value: REQUIRED_FCMS_D1_ID },
      { key: 'ENVIRONMENT', value: 'production' },
    ]);

    // B.1 Fail-Closed SYNC_MODE validation for Production configs
    for (const relCfg of ['wrangler.toml', 'api/wrangler.toml']) {
      const cfgPath = path.join(baseDir, relCfg);
      if (fs.existsSync(cfgPath)) {
        const cfgContent = fs.readFileSync(cfgPath, 'utf8');
        const syncMatch = cfgContent.match(/^[ \t]*SYNC_MODE[ \t]*=[ \t]*["']([^"']*)["']/m);
        const syncVal = syncMatch ? syncMatch[1].trim() : '';
        if (!syncVal) {
          errors.push(`${relCfg} is missing explicit SYNC_MODE setting (must default to "DRY_RUN"; empty or omitted SYNC_MODE is rejected fail-closed).`);
        } else if (syncVal !== 'DRY_RUN' && syncVal !== 'LIVE') {
          errors.push(`${relCfg} has invalid SYNC_MODE="${syncVal}". Permitted values: "DRY_RUN" or "LIVE".`);
        } else if (syncVal === 'LIVE' && options.authorizeLiveSync !== true) {
          errors.push(`LIVE_SYNC_UNAUTHORIZED: ${relCfg} specifies SYNC_MODE="LIVE" without explicit separate approval (--authorize-live-sync or FCMS_LIVE_SYNC_AUTHORIZED=true).`);
        }
      }
    }

    // C. Verify frontend/wrangler.toml
    checkFileContent('frontend/wrangler.toml', [
      { key: 'account_id', value: REQUIRED_FCMS_ACCOUNT_ID },
    ]);

    // D. Check Environment Variables
    const envAccountId = options.cloudflareAccountId || (process.env.CLOUDFLARE_ACCOUNT_ID ? process.env.CLOUDFLARE_ACCOUNT_ID.trim() : undefined);
    if (envAccountId && envAccountId !== REQUIRED_FCMS_ACCOUNT_ID) {
      errors.push(`Environment CLOUDFLARE_ACCOUNT_ID (${envAccountId}) does not match Production FCMS (${REQUIRED_FCMS_ACCOUNT_ID})`);
    }

    // E. Check cached wrangler files if present
    const cachedAccountFiles = [
      path.join(baseDir, 'node_modules', '.cache', 'wrangler', 'wrangler-account.json'),
      path.join(baseDir, 'api', 'node_modules', '.cache', 'wrangler', 'wrangler-account.json')
    ];
    for (const cacheFile of cachedAccountFiles) {
      if (fs.existsSync(cacheFile)) {
        try {
          const data = JSON.parse(fs.readFileSync(cacheFile, 'utf8'));
          if (data.account?.id && data.account.id !== REQUIRED_FCMS_ACCOUNT_ID) {
            errors.push(`Stale wrangler cache file detected (${cacheFile}) pointing to wrong account: ${data.account.id}.`);
          }
        } catch {}
      }
    }
  }

  // 3. STAGING ENVIRONMENT AUDIT
  if (rawEnv === 'staging') {
    const stagingAcc = (options.stagingAccountId || process.env.STAGING_ACCOUNT_ID || '').trim();
    const stagingDbId = (options.stagingDatabaseId || process.env.STAGING_D1_ID || '').trim();
    const stagingDbName = (options.stagingDatabaseName || process.env.STAGING_D1_NAME || 'fcms-staging-db').trim();

    // Staging must have explicit non-empty values
    if (!stagingAcc || stagingAcc.startsWith('UNASSIGNED')) {
      errors.push('Staging target requires an explicit, provisioned Cloudflare Account ID (STAGING_ACCOUNT_ID).');
    }
    if (!stagingDbId || stagingDbId.startsWith('UNASSIGNED')) {
      errors.push('Staging target requires an explicit, provisioned D1 Database UUID (STAGING_D1_ID).');
    }

    // Anti-Production Sieve in Staging
    if (stagingDbId && stagingDbId.toLowerCase() === REQUIRED_FCMS_D1_ID.toLowerCase()) {
      errors.push(`SECURITY_VIOLATION: Staging cannot target Production Database UUID (${REQUIRED_FCMS_D1_ID}).`);
    }
    if (stagingDbName && stagingDbName.toLowerCase() === REQUIRED_FCMS_D1_NAME.toLowerCase()) {
      errors.push(`SECURITY_VIOLATION: Staging cannot target Production Database Name (${REQUIRED_FCMS_D1_NAME}).`);
    }
    const stagingWorkerName = (options.stagingWorkerName || process.env.STAGING_WORKER_NAME || '').trim();
    if (stagingWorkerName && stagingWorkerName.toLowerCase() === 'fm-operations-api') {
      errors.push('SECURITY_VIOLATION: Staging cannot target Production Worker Name (fm-operations-api).');
    }

    // Anti-FCVM Sieve in Staging
    for (const f of FORBIDDEN_STRINGS) {
      if (stagingAcc.toLowerCase().includes(f.id.toLowerCase())) {
        errors.push(`SECURITY_VIOLATION: Staging account references forbidden FCVM identifier: ${f.desc}`);
      }
      if (stagingDbId.toLowerCase().includes(f.id.toLowerCase())) {
        errors.push(`SECURITY_VIOLATION: Staging database ID references forbidden FCVM identifier: ${f.desc}`);
      }
    }

    // Check staging config file if specified
    const stagingConfigFile = options.stagingConfigFile;
    if (stagingConfigFile) {
      const stagingConfigPath = path.join(baseDir, stagingConfigFile);
      if (!fs.existsSync(stagingConfigPath)) {
        errors.push(`Required staging configuration file missing: ${stagingConfigFile}`);
      } else {
        const forbiddenStaging = [
          ...FORBIDDEN_STRINGS,
          { id: REQUIRED_FCMS_D1_ID, desc: 'Production D1 Database UUID' },
          { id: REQUIRED_FCMS_D1_NAME, desc: 'Production D1 Database Name' },
          {
            id: 'fm-operations-api',
            desc: 'Production Worker API Name',
            pattern: /^[ \t]*name[ \t]*=[ \t]*["']fm-operations-api["']/m
          }
        ];
        checkFileContent(stagingConfigFile, [
          { key: 'database_name', value: stagingDbName },
          { key: 'database_id', value: stagingDbId },
          { key: 'ENVIRONMENT', value: 'staging' },
          { key: 'SYNC_MODE', value: 'DRY_RUN' }
        ], [
          ...forbiddenStaging,
          { id: 'SYNC_MODE_LIVE', desc: 'LIVE Sync Mode in Staging', pattern: /^[ \t]*SYNC_MODE[ \t]*=[ \t]*["']LIVE["']/mi }
        ]);
      }
    }
  }

  return {
    success: errors.length === 0,
    env: rawEnv,
    errors,
    warnings
  };
}

// CLI Execution Entrypoint
const isDirectCli = process.argv[1] && path.resolve(process.argv[1]) === path.resolve(__filename);

if (isDirectCli) {
  console.log('🛡️  [FCMS Preflight Guard] Verifying deployment target configuration...');

  const args = process.argv.slice(2);
  const envArg = args.find(a => a.startsWith('--env='));
  const targetEnv = envArg ? envArg.slice(6) : (process.env.TARGET_ENV || process.env.ENVIRONMENT || '');

  const configArg = args.find(a => a.startsWith('--staging-config-file=') || a.startsWith('--config='));
  const stagingConfigFile = configArg
    ? (configArg.startsWith('--config=') ? configArg.slice(9) : configArg.slice(22))
    : 'api/wrangler.staging.toml';

  const stagingConfigPath = path.join(rootDir, stagingConfigFile);
  let tomlAcc, tomlDbId, tomlDbName, tomlWorker;
  if (fs.existsSync(stagingConfigPath)) {
    const content = fs.readFileSync(stagingConfigPath, 'utf8');
    tomlAcc = (content.match(/^[ \t]*account_id[ \t]*=[ \t]*["']([^"']+)["']/m) || [])[1];
    tomlDbId = (content.match(/^[ \t]*database_id[ \t]*=[ \t]*["']([^"']+)["']/m) || [])[1];
    tomlDbName = (content.match(/^[ \t]*database_name[ \t]*=[ \t]*["']([^"']+)["']/m) || [])[1];
    tomlWorker = (content.match(/^[ \t]*name[ \t]*=[ \t]*["']([^"']+)["']/m) || [])[1];
  }

  const accArg = args.find(a => a.startsWith('--staging-account-id='));
  const stagingAcc = accArg ? accArg.slice(21) : (process.env.STAGING_ACCOUNT_ID || tomlAcc);

  const dbIdArg = args.find(a => a.startsWith('--staging-db-id='));
  const stagingDbId = dbIdArg ? dbIdArg.slice(16) : (process.env.STAGING_D1_ID || tomlDbId);

  const dbNameArg = args.find(a => a.startsWith('--staging-db-name='));
  const stagingDbName = dbNameArg ? dbNameArg.slice(18) : (process.env.STAGING_D1_NAME || tomlDbName || 'fcms-staging-db');

  const workerArg = args.find(a => a.startsWith('--staging-worker-name='));
  const stagingWorkerName = workerArg ? workerArg.slice(22) : (process.env.STAGING_WORKER_NAME || tomlWorker);

  const authArg = args.includes('--authorize-production') || args.includes('--confirm-production');
  const isAuthorized = authArg || process.env.FCMS_PRODUCTION_RELEASE_AUTHORIZED === 'true' || process.env.FCMS_PRODUCTION_RELEASE_AUTHORIZED === '1';
  const liveSyncArg = args.includes('--authorize-live-sync');
  const isLiveSyncAuthorized = liveSyncArg || process.env.FCMS_LIVE_SYNC_AUTHORIZED === 'true';

  const result = runPreflightVerification({
    env: targetEnv,
    stagingAccountId: stagingAcc,
    stagingDatabaseId: stagingDbId,
    stagingDatabaseName: stagingDbName,
    stagingWorkerName,
    stagingConfigFile,
    authorizeProduction: isAuthorized,
    authorizeLiveSync: isLiveSyncAuthorized
  });

  if (!result.success) {
    console.error(`\n🚨 [FCMS Preflight Guard] DEPLOYMENT BLOCKED for environment '${targetEnv || 'UNSPECIFIED'}':`);
    for (const err of result.errors) {
      console.error(`   ❌ ${err}`);
    }
    console.error('');
    process.exit(1);
  } else {
    console.log(`✅ [FCMS Preflight Guard] All target assertions PASSED for environment '${result.env}'.`);
    console.log('   Target identity and system isolation verified.\n');
    process.exit(0);
  }
}
