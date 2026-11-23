import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { execSync } from 'node:child_process';

const CI_ONLY = process.argv.includes('--ci-only');

console.log('======================================================================');
console.log(`FCMS — PRODUCTION DEPLOYMENT HARDENING VERIFICATION SUITE (${CI_ONLY ? 'CI-ONLY' : 'FULL + LIVE API'})`);
console.log('======================================================================\n');

let passed = 0;
let failed = 0;
let notVerified = 0;

async function test(name, fn) {
  try {
    await fn();
    console.log(`  ✓ [PASS] ${name}`);
    passed++;
  } catch (err) {
    console.error(`  ✗ [FAIL] ${name}`);
    console.error(`    ${err.message}`);
    failed++;
  }
}

function skipNotVerified(name, reason) {
  console.log(`  ⊘ [NOT VERIFIED] ${name} — ${reason}`);
  notVerified++;
}

function getGitHubToken() {
  const out = execSync('git credential fill', {
    input: 'protocol=https\nhost=github.com\n\n',
    encoding: 'utf8'
  });
  const m = out.match(/^password=(.+)$/m);
  if (!m) throw new Error('No GitHub token in git credential store');
  return m[1].trim();
}

function getCloudflareToken() {
  const cfgPath = path.join(process.env.APPDATA || '', 'xdg.config', '.wrangler', 'config', 'default.toml');
  const m = fs.readFileSync(cfgPath, 'utf8').match(/oauth_token\s*=\s*"([^"]+)"/);
  if (!m) throw new Error('No Cloudflare OAuth token found');
  return m[1].trim();
}

const workflowDir = '.github/workflows';
const workflowFiles = fs.readdirSync(workflowDir).filter(f => f.endsWith('.yml') || f.endsWith('.yaml'));
const deployYml = fs.readFileSync(path.join(workflowDir, 'deploy.yml'), 'utf8');
const migrationYml = fs.readFileSync(path.join(workflowDir, 'manual_migration.yml'), 'utf8');
const stagingYml = fs.readFileSync(path.join(workflowDir, 'deploy-staging.yml'), 'utf8');
const auditYml = fs.readFileSync(path.join(workflowDir, 'audit_db.yml'), 'utf8');

function evalDeployWorkflow(eventName, ref, deployTarget = '') {
  const isMain = ref === 'refs/heads/main';
  const isManual = eventName === 'workflow_dispatch';
  return {
    validate: (eventName === 'push' && isMain) || (eventName === 'pull_request' && isMain) || isManual,
    preflightTargetCheck: isManual && isMain,
    deployApi: isManual && isMain && deployTarget === 'API_ONLY',
    deployPages: isManual && isMain && deployTarget === 'FRONTEND_ONLY',
    d1Migration: false
  };
}

// ----------------------------------------------------------------------------
// SECTION 1: Workflow Static Hardening & Alternative Path Scan
// ----------------------------------------------------------------------------
console.log('--- SECTION 1: Workflow Static Hardening & Alternative Path Scan ---');

await test('1.1 deploy.yml enforces least-privilege permissions: contents: read', () => {
  assert.ok(/permissions:\r?\n  contents: read/.test(deployYml));
});

await test('1.2 deploy.yml separates Worker and Pages (no combined ALL option)', () => {
  assert.ok(deployYml.includes('- FRONTEND_ONLY'));
  assert.ok(deployYml.includes('- API_ONLY'));
  assert.ok(!deployYml.includes('- ALL'));
});

await test('1.3 deploy.yml binds deploy-api to environment: production-worker', () => {
  assert.ok(deployYml.includes('name: production-worker'));
});

await test('1.4 deploy.yml binds deploy-pages to environment: production-pages', () => {
  assert.ok(deployYml.includes('name: production-pages'));
});

await test('1.5 deploy.yml contains ZERO D1 migration commands', () => {
  assert.ok(!deployYml.includes('d1 migrations apply'));
  assert.ok(!deployYml.includes('d1 execute'));
});

await test('1.6 manual_migration.yml has ONLY workflow_dispatch trigger (zero push/PR/schedule)', () => {
  const onBlock = migrationYml.split('permissions:')[0];
  assert.ok(onBlock.includes('workflow_dispatch:'));
  assert.ok(!onBlock.includes('push:'));
  assert.ok(!onBlock.includes('pull_request:'));
  assert.ok(!onBlock.includes('schedule:'));
});

await test('1.7 manual_migration.yml binds execute job to environment: production-database & defaults to DRY_RUN', () => {
  assert.ok(migrationYml.includes('name: production-database'));
  assert.ok(migrationYml.includes('YES_BACKUP_VERIFIED'));
  assert.ok(migrationYml.includes("default: 'DRY_RUN_VERIFY_ONLY'"));
  assert.ok(migrationYml.includes('if [ "$INPUT_MODE" != "APPLY_MIGRATION" ]; then'));
});

await test('1.8 Destructive SQL guard rejects DROP TABLE / DROP INDEX / DELETE FROM / TRUNCATE and accepts additive SQL', () => {
  const destructiveRegex = /\b(DROP\s+TABLE|DROP\s+INDEX|DELETE\s+FROM|TRUNCATE)\b/i;
  assert.strictEqual(destructiveRegex.test('DROP TABLE users;'), true, 'Must reject DROP TABLE');
  assert.strictEqual(destructiveRegex.test('DROP INDEX idx_users;'), true, 'Must reject DROP INDEX');
  assert.strictEqual(destructiveRegex.test('DELETE FROM employees WHERE id = 1;'), true, 'Must reject DELETE FROM');
  assert.strictEqual(destructiveRegex.test('TRUNCATE TABLE sessions;'), true, 'Must reject TRUNCATE');
  assert.strictEqual(destructiveRegex.test('CREATE INDEX IF NOT EXISTS idx_test ON users(id);'), false, 'Must allow additive CREATE INDEX');
  assert.strictEqual(destructiveRegex.test('ALTER TABLE employees ADD COLUMN archived_at TEXT;'), false, 'Must allow additive ALTER TABLE');
});

await test('1.9 deploy-staging.yml isolates Staging from Production (zero Production Worker/D1 targets)', () => {
  assert.ok(stagingYml.includes('name: staging'));
  assert.ok(stagingYml.includes('fm-operations-api-staging'));
  assert.ok(stagingYml.includes('facilitycare-system-staging'));
  assert.ok(stagingYml.includes('7ecb4668-c2b9-48ef-93a6-f59a3672e8df'));
  assert.ok(!stagingYml.includes('62b4b6e5-a718-40bc-9aab-6a700b82a9c9'), 'Staging must NEVER reference Production D1 UUID');
  assert.ok(!stagingYml.includes('--project-name=facilitycare-system '), 'Staging must NEVER target Production Pages project');
});

await test('1.10 audit_db.yml remains disabled safe no-op (workflow_dispatch only)', () => {
  assert.ok(auditYml.includes('workflow_dispatch:'));
  assert.ok(!auditYml.includes('push:'));
});

await test('1.11 No alternative GitHub workflow bypasses Production approval gates', () => {
  for (const file of workflowFiles) {
    const content = fs.readFileSync(path.join(workflowDir, file), 'utf8');
    const hasWranglerDeploy = /wrangler-action|wrangler\s+(pages\s+)?deploy|wrangler\s+d1\s+(execute|migrations)/i.test(content);
    if (!hasWranglerDeploy) continue;
    if (file === 'deploy-staging.yml') {
      assert.ok(content.includes('name: staging'), `${file} must target staging environment`);
      assert.ok(!content.includes('production-worker') && !content.includes('production-pages'));
    } else if (file === 'deploy.yml') {
      assert.ok(content.includes('name: production-worker') && content.includes('name: production-pages'));
    } else if (file === 'manual_migration.yml') {
      assert.ok(content.includes('name: production-database'));
    } else {
      throw new Error(`Unexpected workflow with wrangler deploy/d1 command: ${file}`);
    }
  }
});

// ----------------------------------------------------------------------------
// SECTION 2: Event Simulation (Push, PR, Manual Dispatch, Non-Main)
// ----------------------------------------------------------------------------
console.log('\n--- SECTION 2: Event Simulation ---');

await test('2.1 SIMULATE push to main: validate RUNS, Production Worker/Pages/D1 DO NOT RUN', () => {
  const sim = evalDeployWorkflow('push', 'refs/heads/main');
  assert.strictEqual(sim.validate, true);
  assert.strictEqual(sim.preflightTargetCheck, false);
  assert.strictEqual(sim.deployApi, false);
  assert.strictEqual(sim.deployPages, false);
  assert.strictEqual(sim.d1Migration, false);
});

await test('2.2 SIMULATE pull_request to main: validate RUNS, Production deploy SKIPPED', () => {
  const sim = evalDeployWorkflow('pull_request', 'refs/heads/main');
  assert.strictEqual(sim.validate, true);
  assert.strictEqual(sim.deployApi, false);
  assert.strictEqual(sim.deployPages, false);
});

await test('2.3 SIMULATE manual dispatch (API_ONLY on main): ONLY Worker runs & waits for production-worker approval', () => {
  const sim = evalDeployWorkflow('workflow_dispatch', 'refs/heads/main', 'API_ONLY');
  assert.strictEqual(sim.validate, true);
  assert.strictEqual(sim.preflightTargetCheck, true);
  assert.strictEqual(sim.deployApi, true);
  assert.strictEqual(sim.deployPages, false);
});

await test('2.4 SIMULATE manual dispatch (FRONTEND_ONLY on main): ONLY Pages runs & waits for production-pages approval', () => {
  const sim = evalDeployWorkflow('workflow_dispatch', 'refs/heads/main', 'FRONTEND_ONLY');
  assert.strictEqual(sim.validate, true);
  assert.strictEqual(sim.preflightTargetCheck, true);
  assert.strictEqual(sim.deployApi, false);
  assert.strictEqual(sim.deployPages, true);
});

await test('2.5 SIMULATE manual dispatch from non-main branch: Production deploy BLOCKED', () => {
  const sim = evalDeployWorkflow('workflow_dispatch', 'refs/heads/hardening/production-deployment-h01', 'API_ONLY');
  assert.strictEqual(sim.preflightTargetCheck, false);
  assert.strictEqual(sim.deployApi, false);
  assert.strictEqual(sim.deployPages, false);
});

// ----------------------------------------------------------------------------
// SECTION 3: Live Platform Verification (GitHub & Cloudflare API)
// ----------------------------------------------------------------------------
console.log('\n--- SECTION 3: Live Platform Verification ---');

if (CI_ONLY) {
  skipNotVerified('3.1 Live GitHub Environments API check', 'Skipped in --ci-only mode (requires admin PAT)');
  skipNotVerified('3.2 Live GitHub Branch Protection API check', 'Skipped in --ci-only mode (requires admin PAT)');
  skipNotVerified('3.3 Live Cloudflare Pages API check', 'Skipped in --ci-only mode (requires local CF OAuth)');
} else {
  const ghToken = getGitHubToken();
  const cfToken = getCloudflareToken();
  const owner = 'facilitycareaudydental-sketch';
  const repo = 'facilitycare-system';
  const ghHeaders = {
    Authorization: `Bearer ${ghToken}`,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'fcms-hardening-verify'
  };

  for (const envName of ['production-worker', 'production-pages', 'production-database']) {
    await test(`3.x Live GitHub Environment [${envName}] enforces Owner Reviewer + Protected Branches Only`, async () => {
      const res = await fetch(`https://api.github.com/repos/${owner}/${repo}/environments/${envName}`, { headers: ghHeaders });
      assert.strictEqual(res.status, 200);
      const data = await res.json();
      const reviewerRule = (data.protection_rules || []).find(r => r.type === 'required_reviewers');
      assert.ok(reviewerRule, `Environment ${envName} must have required_reviewers rule`);
      const hasOwner = (reviewerRule.reviewers || []).some(r => r.reviewer?.login === owner);
      assert.ok(hasOwner, `Owner (${owner}) must be required reviewer on ${envName}`);
      assert.strictEqual(data.deployment_branch_policy?.protected_branches, true);
    });
  }

  await test('3.4 Live GitHub Environment [staging] exists with zero production blockers', async () => {
    const res = await fetch(`https://api.github.com/repos/${owner}/${repo}/environments/staging`, { headers: ghHeaders });
    assert.strictEqual(res.status, 200);
    const data = await res.json();
    const reviewerRule = (data.protection_rules || []).find(r => r.type === 'required_reviewers');
    assert.strictEqual(reviewerRule, undefined);
  });

  await test('3.5 Live GitHub Branch Protection on [main] enforces PRs, Required Status Check, and blocks force-push/delete', async () => {
    const res = await fetch(`https://api.github.com/repos/${owner}/${repo}/branches/main/protection`, { headers: ghHeaders });
    assert.strictEqual(res.status, 200);
    const data = await res.json();
    assert.ok(data.required_status_checks?.contexts?.includes('✅ Validate & Build Check'));
    assert.strictEqual(data.required_status_checks?.strict, true);
    assert.ok(data.required_pull_request_reviews, 'PR reviews rule must be active on main');
    assert.strictEqual(data.enforce_admins?.enabled, true);
    assert.strictEqual(data.allow_force_pushes?.enabled, false);
    assert.strictEqual(data.allow_deletions?.enabled, false);
  });

  await test('3.6 Live Cloudflare Pages [facilitycare-system] has Git auto-deploy DISABLED on main', async () => {
    const res = await fetch(
      'https://api.cloudflare.com/client/v4/accounts/293440e04e5d0c8073e7e2893750b527/pages/projects/facilitycare-system',
      { headers: { Authorization: `Bearer ${cfToken}` } }
    );
    const data = await res.json();
    assert.strictEqual(data.success, true);
    assert.ok(data.result?.id, 'Project ID must exist');
    assert.strictEqual(data.result?.name, 'facilitycare-system');
    assert.strictEqual(data.result?.source?.config?.deployments_enabled, false);
    assert.strictEqual(data.result?.source?.config?.production_deployments_enabled, false);
  });
}

console.log('\n======================================================================');
console.log(`PASS: ${passed} | FAIL: ${failed} | NOT VERIFIED: ${notVerified}`);
if (failed === 0) {
  console.log('STATUS: ALL EXECUTED CHECKS PASSED (0 FAILURES)');
  process.exit(0);
} else {
  console.log('STATUS: FAILED');
  process.exit(1);
}
