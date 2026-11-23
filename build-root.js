const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('=== Starting Root Build Process ===');

const ALLOWED_BUILD_ENVS = new Set(['production', 'staging', 'development']);
const cliArgs = process.argv.slice(2).join(' ');
const cliEnvArg = process.argv.slice(2).find(a => a.startsWith('--env='));
const cliEnv = cliEnvArg ? cliEnvArg.slice(6).trim().toLowerCase() : '';
const envVarEnv = (process.env.ENVIRONMENT || '').trim().toLowerCase();
const targetVarEnv = (process.env.TARGET_ENV || '').trim().toLowerCase();

if (cliEnv && envVarEnv && cliEnv !== envVarEnv) {
  console.error(`[BUILD SECURITY FATAL] Environment conflict: --env=${cliEnv} differs from ENVIRONMENT=${envVarEnv}`);
  process.exit(1);
}
if (cliEnv && targetVarEnv && cliEnv !== targetVarEnv) {
  console.error(`[BUILD SECURITY FATAL] Environment conflict: --env=${cliEnv} differs from TARGET_ENV=${targetVarEnv}`);
  process.exit(1);
}
if (envVarEnv && targetVarEnv && envVarEnv !== targetVarEnv) {
  console.error(`[BUILD SECURITY FATAL] Environment conflict: ENVIRONMENT=${envVarEnv} differs from TARGET_ENV=${targetVarEnv}`);
  process.exit(1);
}

const envName = cliEnv || envVarEnv || targetVarEnv;
if (!envName) {
  console.error('[BUILD FATAL] Explicit target environment required (--env=production, --env=staging, or --env=development). No default fallback allowed.');
  process.exit(1);
}
if (!ALLOWED_BUILD_ENVS.has(envName)) {
  console.error(`[BUILD FATAL] Unknown or unauthorized target environment '${envName}'. Permitted: production, staging, development.`);
  process.exit(1);
}

// 1. Run the frontend build
try {
  console.log('Executing frontend build script...');
  execSync(`node frontend/build.js ${cliArgs}`.trim(), { stdio: 'inherit' });
} catch (error) {
  console.error('Error during frontend build execution:', error);
  process.exit(1);
}

// 2. Copy frontend/dist to root/dist
const srcDist = path.join(__dirname, 'frontend', 'dist');
const destDist = path.join(__dirname, 'dist');

console.log(`Copying build output from ${srcDist} to ${destDist}...`);

try {
  if (fs.existsSync(destDist)) {
    fs.rmSync(destDist, { recursive: true, force: true });
  }
  
  function copyDir(src, dest) {
    fs.mkdirSync(dest, { recursive: true });
    fs.readdirSync(src).forEach(file => {
      const srcFile = path.join(src, file);
      const destFile = path.join(dest, file);
      if (fs.statSync(srcFile).isDirectory()) {
        copyDir(srcFile, destFile);
      } else {
        fs.copyFileSync(srcFile, destFile);
      }
    });
  }

  if (fs.existsSync(srcDist)) {
    copyDir(srcDist, destDist);

    // Root-level distribution validation
    const rootIndexHtml = fs.readFileSync(path.join(destDist, 'index.html'), 'utf8');
    if (rootIndexHtml.includes('__API_BASE_URL__') || rootIndexHtml.includes('__ENVIRONMENT__')) {
      console.error('[BUILD FATAL] Unreplaced placeholders detected in root dist/index.html');
      process.exit(1);
    }
    if (!rootIndexHtml.includes(`ENVIRONMENT: '${envName}'`)) {
      console.error(`[BUILD FATAL] Root dist/index.html does not contain expected ENVIRONMENT: '${envName}'`);
      process.exit(1);
    }
    if (envName === 'staging') {
      if (rootIndexHtml.includes('fm-operations-api.facilitycare-audydental.workers.dev') || rootIndexHtml.toLowerCase().includes('fcvm')) {
        console.error('[BUILD SECURITY FATAL] Production or FCVM Worker URL detected in root dist/index.html for staging build!');
        process.exit(1);
      }
      if (!rootIndexHtml.includes('https://fm-operations-api-staging.facilitycare-audydental.workers.dev')) {
        console.error('[BUILD SECURITY FATAL] Approved Staging Worker URL missing in root dist/index.html for staging build!');
        process.exit(1);
      }
      // Mirror frontend/dist-staging to root dist-staging
      const srcStagingDist = path.join(__dirname, 'frontend', 'dist-staging');
      const destStagingDist = path.join(__dirname, 'dist-staging');
      if (fs.existsSync(srcStagingDist)) {
        if (fs.existsSync(destStagingDist)) {
          fs.rmSync(destStagingDist, { recursive: true, force: true });
        }
        copyDir(srcStagingDist, destStagingDist);
        const stagingHtml = fs.readFileSync(path.join(destStagingDist, 'index.html'), 'utf8');
        if (!stagingHtml.includes("ENVIRONMENT: 'staging'") || stagingHtml.includes('fm-operations-api.facilitycare-audydental.workers.dev')) {
          console.error('[BUILD SECURITY FATAL] Staging distribution assertion failed for dist-staging/index.html!');
          process.exit(1);
        }
        console.log(`✅ Staging output synchronized to root ${destStagingDist}`);
      }
    } else if (envName === 'production') {
      if (rootIndexHtml.includes('fm-operations-api-staging') || rootIndexHtml.toLowerCase().includes('fcvm')) {
        console.error('[BUILD SECURITY FATAL] Staging or FCVM Worker URL detected in root dist/index.html for production build!');
        process.exit(1);
      }
      if (!rootIndexHtml.includes('https://fm-operations-api.facilitycare-audydental.workers.dev')) {
        console.error('[BUILD SECURITY FATAL] Approved Production Worker URL missing in root dist/index.html for production build!');
        process.exit(1);
      }
    } else if (envName === 'development') {
      if (rootIndexHtml.includes('fm-operations-api.facilitycare-audydental.workers.dev') || rootIndexHtml.toLowerCase().includes('fcvm')) {
        console.error('[BUILD SECURITY FATAL] Production or FCVM Worker URL detected in root dist/index.html for development build!');
        process.exit(1);
      }
    }

    const bundleMatch = rootIndexHtml.match(/\/assets\/(app(?:-[a-zA-Z0-9]+)?\.js)/);
    if (!bundleMatch || !fs.existsSync(path.join(destDist, 'assets', bundleMatch[1]))) {
      console.error('[BUILD FATAL] Active JS bundle referenced in root dist/index.html does not exist on disk!');
      process.exit(1);
    }

    console.log('=== Root Build Completed Successfully! ===');
  } else {
    console.error('Error: Build output folder frontend/dist not found.');
    process.exit(1);
  }
} catch (error) {
  console.error('Error copying build output to root directory:', error);
  process.exit(1);
}
