/**
 * deploy-staging.js — FCMS Frontend Staging Deployment Runner (Isolated Workspace Architecture)
 * ─────────────────────────────────────────────────────────────────────────────────────────────
 * Safely deploys FCMS Staging Frontend to Cloudflare Pages using an isolated workspace.
 * 
 * Key Security Guarantees:
 * 1. Zero Modification of Production Configuration:
 *    frontend/wrangler.toml is NEVER touched, moved, overwritten, or modified (even temporarily).
 * 2. Full Workspace Isolation:
 *    A dedicated, temporary workspace (.staging-deploy-workspace) is created.
 *    frontend/wrangler.staging.toml is copied into the isolated workspace as wrangler.toml.
 *    frontend/dist-staging is replicated into the isolated workspace.
 * 3. Schema Compliance:
 *    Asserts that frontend/wrangler.staging.toml contains zero "account_id" fields.
 * 4. Target Isolation:
 *    Strictly targets project 'facilitycare-system-staging' and branch 'staging'.
 *    Enforces Cloudflare Account ID: 293440e04e5d0c8073e7e2893750b527 (FCMS).
 * 5. Deterministic Process Execution:
 *    Uses process.execPath + local Wrangler JS binary (bypassing shell/npx path mutability).
 * 6. Fail-Safe Cleanup:
 *    Guaranteed cleanup of temporary workspace via try ... finally and signal traps.
 * ─────────────────────────────────────────────────────────────────────────────────────────────
 */

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const workspaceDir = __dirname;
const stagingTomlPath = path.join(workspaceDir, 'wrangler.staging.toml');
const prodTomlPath = path.join(workspaceDir, 'wrangler.toml');
const distStagingSource = path.join(workspaceDir, 'dist-staging');
const isolatedWorkspace = path.join(workspaceDir, '.staging-deploy-workspace');

const REQUIRED_FCMS_ACCOUNT_ID = '293440e04e5d0c8073e7e2893750b527';
const EXPECTED_PROJECT_NAME = 'facilitycare-system-staging';
const EXPECTED_BRANCH = 'staging';

const isValidateOnly = process.argv.includes('--validate-only') || process.argv.includes('--dry-run');

console.log('🛡️  [FCMS Frontend Staging Deploy Runner — Isolated Architecture]');
console.log(`   Mode: ${isValidateOnly ? 'VALIDATION ONLY (Dry-Run)' : 'FULL DEPLOY'}`);

// 1. Verify build output exists
if (!fs.existsSync(distStagingSource) || !fs.existsSync(path.join(distStagingSource, 'index.html'))) {
  console.error('❌ [DEPLOY ERROR] Staging build output missing: frontend/dist-staging/index.html');
  console.error('   Please run "npm run build:staging" first.');
  process.exit(1);
}

// 2. Verify wrangler.staging.toml exists and conforms to Pages schema
if (!fs.existsSync(stagingTomlPath)) {
  console.error('❌ [DEPLOY ERROR] Missing configuration file: frontend/wrangler.staging.toml');
  process.exit(1);
}

const stagingTomlContent = fs.readFileSync(stagingTomlPath, 'utf8');

// Strict check: Cloudflare Pages projects forbid account_id
if (/^[ \t]*account_id[ \t]*=/m.test(stagingTomlContent)) {
  console.error('❌ [DEPLOY ERROR] Configuration file for Pages projects does not support "account_id"!');
  console.error('   Please remove "account_id" from frontend/wrangler.staging.toml.');
  process.exit(1);
}

// Verify required Pages fields in staging config
if (!stagingTomlContent.includes(`name = "${EXPECTED_PROJECT_NAME}"`)) {
  console.error(`❌ [DEPLOY ERROR] Staging Pages configuration must specify: name = "${EXPECTED_PROJECT_NAME}"`);
  process.exit(1);
}

if (!stagingTomlContent.includes('pages_build_output_dir = "dist-staging"')) {
  console.error('❌ [DEPLOY ERROR] Staging Pages configuration must specify: pages_build_output_dir = "dist-staging"');
  process.exit(1);
}

console.log('✅ Staging configuration verified (zero account_id, target: facilitycare-system-staging).');

// 3. Verify production wrangler.toml integrity check (read-only verification)
let prodTomlSnapshot = null;
if (fs.existsSync(prodTomlPath)) {
  prodTomlSnapshot = fs.readFileSync(prodTomlPath, 'utf8');
}

// 4. Resolve Wrangler Binary
let wranglerBin;
try {
  wranglerBin = require.resolve('wrangler/bin/wrangler.js', { paths: [workspaceDir, path.resolve(workspaceDir, '..')] });
} catch (resolveErr) {
  console.error('❌ [DEPLOY ERROR] Could not locate local Wrangler CLI binary:', resolveErr.message);
  process.exit(1);
}

// Cleanup helper
function cleanupWorkspace() {
  if (fs.existsSync(isolatedWorkspace)) {
    try {
      fs.rmSync(isolatedWorkspace, { recursive: true, force: true });
    } catch {}
  }
}

// Signal traps for clean abort
process.on('SIGINT', () => {
  cleanupWorkspace();
  process.exit(1);
});
process.on('SIGTERM', () => {
  cleanupWorkspace();
  process.exit(1);
});

try {
  // 5. Construct Isolated Staging Workspace
  cleanupWorkspace();
  fs.mkdirSync(isolatedWorkspace, { recursive: true });

  // Place staging config as wrangler.toml inside isolated workspace ONLY
  fs.copyFileSync(stagingTomlPath, path.join(isolatedWorkspace, 'wrangler.toml'));

  // Replicate dist-staging into isolated workspace
  fs.cpSync(distStagingSource, path.join(isolatedWorkspace, 'dist-staging'), { recursive: true });
  console.log('✅ Isolated staging deployment workspace constructed successfully.');

  if (isValidateOnly) {
    console.log('🔎 [Dry-Run Validation]');
    console.log(`   Isolated Directory: ${isolatedWorkspace}`);
    console.log(`   Config File: ${path.join(isolatedWorkspace, 'wrangler.toml')} (Clean Pages Schema)`);
    console.log(`   Build Directory: ${path.join(isolatedWorkspace, 'dist-staging')}`);
    console.log(`   Target Project: ${EXPECTED_PROJECT_NAME}`);
    console.log(`   Target Branch: ${EXPECTED_BRANCH}`);
    console.log(`   Target Account: ${REQUIRED_FCMS_ACCOUNT_ID}`);
    console.log('✅ Validation completed successfully. Ready for deployment.');
  } else {
    console.log(`📡 Executing Cloudflare Pages deployment to ${EXPECTED_PROJECT_NAME} (branch: ${EXPECTED_BRANCH})...`);
    execFileSync(
      process.execPath,
      [
        wranglerBin,
        'pages',
        'deploy',
        'dist-staging',
        `--project-name=${EXPECTED_PROJECT_NAME}`,
        `--branch=${EXPECTED_BRANCH}`
      ],
      {
        cwd: isolatedWorkspace,
        stdio: 'inherit',
        env: {
          ...process.env,
          CLOUDFLARE_ACCOUNT_ID: REQUIRED_FCMS_ACCOUNT_ID
        }
      }
    );
    console.log('✨ Staging deployment completed successfully!');
  }
} catch (deployErr) {
  console.error('❌ [DEPLOY FAILED]:', deployErr.message);
  process.exit(1);
} finally {
  // 6. Cleanup Isolated Workspace
  cleanupWorkspace();

  // 7. Verify production configuration was never modified
  if (prodTomlSnapshot !== null) {
    const currentProdToml = fs.readFileSync(prodTomlPath, 'utf8');
    if (currentProdToml !== prodTomlSnapshot) {
      console.error('🚨 [SECURITY FATAL] Production frontend/wrangler.toml was altered during deployment!');
      process.exit(1);
    }
  }
  console.log('🔒 Production configuration frontend/wrangler.toml integrity verified (100% UNTOUCHED).');
}
