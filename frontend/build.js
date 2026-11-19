/**
 * Simple build script - bundles all JS modules into a single app.js
 * Uses esbuild if available, otherwise generates a bundle manually
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { execSync } = require('child_process');

const crypto = require('crypto');

const srcDir = path.join(__dirname, 'src');
const publicDir = path.join(__dirname, 'public');
const cliOutArg = process.argv.slice(2).find(a => a.startsWith('--outdir='));
const distDir = cliOutArg ? path.resolve(cliOutArg.slice(9).trim()) : path.join(__dirname, 'dist');

function sleepSync(ms) {
  try {
    Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, ms);
  } catch {
    const start = Date.now();
    while (Date.now() - start < ms) {}
  }
}

function readFileSyncRetry(filePath, encoding = 'utf8', maxRetries = 5) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return fs.readFileSync(filePath, encoding);
    } catch (err) {
      if (attempt === maxRetries) throw err;
      sleepSync(30 * attempt);
    }
  }
}

function writeFileSyncRetry(filePath, content, encoding = 'utf8', maxRetries = 5) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return fs.writeFileSync(filePath, content, encoding);
    } catch (err) {
      if (attempt === maxRetries) throw err;
      sleepSync(30 * attempt);
    }
  }
}

const APPROVED_STAGING_API_URL = 'https://fm-operations-api-staging.facilitycare-audydental.workers.dev';
const PRODUCTION_API_URL = 'https://fm-operations-api.facilitycare-audydental.workers.dev';
const PRODUCTION_WORKER_HOST = 'fm-operations-api.facilitycare-audydental.workers.dev';
const STAGING_WORKER_HOST = 'fm-operations-api-staging.facilitycare-audydental.workers.dev';
const ALLOWED_BUILD_ENVS = new Set(['production', 'staging', 'development']);

// 0. Fail-Closed Explicit Environment Resolution (Zero silent fallback allowed)
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
  console.error(`[BUILD FATAL] Invalid or unknown target environment '${envName}'. Permitted: production, staging, development.`);
  process.exit(1);
}

// Clean previous build output so stale cross-environment assets never linger
if (fs.existsSync(distDir)) {
  fs.rmSync(distDir, { recursive: true, force: true });
}
fs.mkdirSync(distDir, { recursive: true });
fs.mkdirSync(path.join(distDir, 'assets'), { recursive: true });

// Copy public files
function copyDir(src, dest) {
  if (!fs.existsSync(dest)) fs.mkdirSync(dest, { recursive: true });
  fs.readdirSync(src).forEach(f => {
    const srcF = path.join(src, f);
    const destF = path.join(dest, f);
    if (fs.statSync(srcF).isDirectory()) copyDir(srcF, destF);
    else fs.copyFileSync(srcF, destF);
  });
}

copyDir(publicDir, distDir);

// Copy Pages Functions to dist so Wrangler deploys them
const functionsDir = path.join(__dirname, '..', 'functions');
if (fs.existsSync(functionsDir)) {
  copyDir(functionsDir, path.join(distDir, 'functions'));
}

// Scan non-config source & public files to guarantee no hardcoded Worker URLs bypass config.js
function scanDirForCrossEnvUrls(dir, skipFileNames = new Set()) {
  const findings = [];
  for (const entry of fs.readdirSync(dir)) {
    const full = path.join(dir, entry);
    if (fs.statSync(full).isDirectory()) {
      findings.push(...scanDirForCrossEnvUrls(full, skipFileNames));
    } else if (!skipFileNames.has(path.basename(full))) {
      const content = fs.readFileSync(full, 'utf8');
      const lower = content.toLowerCase();
      if (lower.includes(PRODUCTION_WORKER_HOST) || lower.includes(STAGING_WORKER_HOST) || lower.includes('fcvm')) {
        findings.push(full);
      }
    }
  }
  return findings;
}

const leakedSourceFiles = [
  ...scanDirForCrossEnvUrls(srcDir, new Set(['config.js'])),
  ...scanDirForCrossEnvUrls(publicDir)
];
if (leakedSourceFiles.length > 0) {
  console.error(`[SECURITY FATAL] Hardcoded Worker/FCVM URL detected outside config.js: ${leakedSourceFiles.join(', ')}`);
  process.exit(1);
}

let apiUrl = process.env.API_BASE_URL !== undefined ? String(process.env.API_BASE_URL).trim() : '';

if (envName === 'staging') {
  if (apiUrl) {
    if (apiUrl.toLowerCase().includes(PRODUCTION_WORKER_HOST)) {
      console.error('[SECURITY FATAL] Production Worker URL detected in staging build output!');
      process.exit(1);
    }
    if (apiUrl.toLowerCase().includes('fcvm')) {
      console.error('[SECURITY FATAL] FCVM Worker URL detected in staging build configuration!');
      process.exit(1);
    }
    const normalizedApiUrl = apiUrl.replace(/\/+$/, '');
    if (normalizedApiUrl !== APPROVED_STAGING_API_URL) {
      console.error(`[SECURITY FATAL] Unapproved API_BASE_URL for staging build: '${apiUrl}'. Only '${APPROVED_STAGING_API_URL}' is permitted.`);
      process.exit(1);
    }
    apiUrl = APPROVED_STAGING_API_URL;
  } else {
    apiUrl = APPROVED_STAGING_API_URL;
  }
} else if (envName === 'production') {
  if (apiUrl) {
    if (apiUrl.toLowerCase().includes(STAGING_WORKER_HOST)) {
      console.error('[SECURITY FATAL] Staging Worker URL detected in production build configuration!');
      process.exit(1);
    }
    if (apiUrl.toLowerCase().includes('fcvm')) {
      console.error('[SECURITY FATAL] FCVM Worker URL detected in production build configuration!');
      process.exit(1);
    }
    const normalizedProdUrl = apiUrl.replace(/\/+$/, '');
    if (normalizedProdUrl !== PRODUCTION_API_URL) {
      console.error(`[SECURITY FATAL] Unapproved API_BASE_URL for production build: '${apiUrl}'. Only '${PRODUCTION_API_URL}' is permitted.`);
      process.exit(1);
    }
    apiUrl = PRODUCTION_API_URL;
  } else {
    apiUrl = PRODUCTION_API_URL;
  }
} else if (envName === 'development') {
  if (apiUrl) {
    const lowerDevUrl = apiUrl.toLowerCase();
    if (lowerDevUrl.includes(PRODUCTION_WORKER_HOST)) {
      console.error('[SECURITY FATAL] Production Worker URL is forbidden in development build!');
      process.exit(1);
    }
    if (lowerDevUrl.includes('fcvm')) {
      console.error('[SECURITY FATAL] FCVM Worker URL is forbidden in development build!');
      process.exit(1);
    }
  }
}

let html = readFileSyncRetry(path.join(distDir, 'index.html'), 'utf8');
html = html.replace('__API_BASE_URL__', apiUrl);
html = html.replace('__ENVIRONMENT__', envName);

// Validate build output to prevent unreplaced placeholders or cross-environment leaks
if (html.includes('__API_BASE_URL__') || html.includes('__ENVIRONMENT__')) {
  console.error('[BUILD ERROR] Unreplaced configuration placeholders remain in dist/index.html');
  process.exit(1);
}

if (!html.includes(`ENVIRONMENT: '${envName}'`) || !html.includes(`API_BASE_URL: '${apiUrl}'`)) {
  console.error('[BUILD ERROR] Failed to inject exact ENVIRONMENT and API_BASE_URL into dist/index.html');
  process.exit(1);
}

if (envName === 'staging') {
  if (html.includes(PRODUCTION_WORKER_HOST) || html.toLowerCase().includes('fcvm')) {
    console.error('[SECURITY FATAL] Production or FCVM Worker URL detected in staging build output!');
    process.exit(1);
  }
  if (!html.includes(APPROVED_STAGING_API_URL)) {
    console.error('[SECURITY FATAL] Approved Staging Worker URL missing in staging build output!');
    process.exit(1);
  }
}

if (envName === 'production') {
  if (html.includes(STAGING_WORKER_HOST) || html.toLowerCase().includes('fcvm')) {
    console.error('[SECURITY FATAL] Staging or FCVM Worker URL detected in production build output!');
    process.exit(1);
  }
  if (!html.includes(PRODUCTION_API_URL)) {
    console.error('[SECURITY FATAL] Approved Production Worker URL missing in production build output!');
    process.exit(1);
  }
}

if (envName === 'development') {
  if (html.includes(PRODUCTION_WORKER_HOST) || html.toLowerCase().includes('fcvm')) {
    console.error('[SECURITY FATAL] Production or FCVM Worker URL detected in development build output!');
    process.exit(1);
  }
}

// Verify script execution order: window.__FM_CONFIG must appear before app.js script tag
const configScriptIndex = html.indexOf('window.__FM_CONFIG');
const appScriptIndex = html.indexOf('/assets/app');
if (configScriptIndex === -1 || appScriptIndex === -1 || configScriptIndex >= appScriptIndex) {
  console.error('[BUILD ERROR] Runtime configuration window.__FM_CONFIG must be loaded before app.js executes!');
  process.exit(1);
}

writeFileSyncRetry(path.join(distDir, 'index.html'), html);


// Bundle JS with esbuild
let esbuildCmd = null;
try {
  execSync('esbuild --version', { stdio: 'ignore' });
  esbuildCmd = 'esbuild'; // globally installed
} catch {
  try {
    execSync('npx --no esbuild --version', { stdio: 'ignore' });
    esbuildCmd = 'npx --no esbuild'; // via npx
  } catch {
    // Install esbuild locally
    console.log('Installing esbuild locally...');
    try {
      execSync('npm install --save-dev esbuild', { stdio: 'inherit', cwd: path.join(__dirname, '..') });
      esbuildCmd = 'npx esbuild';
    } catch {
      esbuildCmd = null;
    }
  }
}

let activeBundleFilename = 'app.js';

if (esbuildCmd) {
  console.log(`Bundling with esbuild (${esbuildCmd})...`);
  const metafile = path.join(os.tmpdir(), `fcms_esbuild_meta_${process.pid}_${Date.now()}.json`);
  try {
    // Use content hash in filename so Cloudflare CDN always serves fresh file
    execSync(
      `${esbuildCmd} "${path.join(srcDir, 'app.js')}" --bundle --outdir="${path.join(distDir, 'assets')}" --entry-names=[name]-[hash] --format=esm --minify --platform=browser --metafile="${metafile}"`,
      { stdio: 'inherit' }
    );

    // Read metafile to find actual output filename
    const meta = JSON.parse(fs.readFileSync(metafile, 'utf8'));
    const outputs = Object.keys(meta.outputs).filter(f => f.endsWith('.js') && !f.endsWith('.css'));
    const outFilename = outputs.length ? path.basename(outputs[0]) : 'app.js';
    activeBundleFilename = outFilename;
    console.log(`✅ Bundle created: dist/assets/${outFilename}`);

    // Patch index.html to reference the hashed filename deterministically
    let htmlContent = readFileSyncRetry(path.join(distDir, 'index.html'), 'utf8');
    const cacheBuster = `?v=${Date.now()}`;
    htmlContent = htmlContent.replace(/\/assets\/app(?:-[a-zA-Z0-9]+)?\.js(?:\?[^"]*)?/, `/assets/${outFilename}${cacheBuster}`);
    writeFileSyncRetry(path.join(distDir, 'index.html'), htmlContent);
    console.log(`✅ index.html patched: /assets/${outFilename}`);
  } catch (err) {
    console.error('esbuild bundle failed, falling back to source copy:', err.message);
    esbuildCmd = null;
  } finally {
    try { if (fs.existsSync(metafile)) fs.unlinkSync(metafile); } catch {}
  }
}

// Keep dist/src synchronized if present so no stale config.js remains
if (fs.existsSync(path.join(distDir, 'src'))) {
  copyDir(srcDir, path.join(distDir, 'src'));
}

if (!esbuildCmd) {
  console.log('Falling back: copying source files (native ESM mode)...');
  copyDir(srcDir, path.join(distDir, 'src'));
  let html2 = fs.readFileSync(path.join(distDir, 'index.html'), 'utf8');
  html2 = html2.replace('/assets/app.js', '/src/app.js');
  fs.writeFileSync(path.join(distDir, 'index.html'), html2);
}

// Create _redirects for SPA routing on Cloudflare Pages
// form.html must be matched BEFORE the SPA catch-all
fs.writeFileSync(path.join(distDir, '_redirects'),
  '/form.html /form.html 200\n' +
  '/form      /form.html 200\n' +
  '/* /index.html 200\n'
);

// Compute strict CSP header using SHA-256 hash of inline window.__FM_CONFIG script
const configScriptMatch = html.match(/<script>([\s\S]*?window\.__FM_CONFIG[\s\S]*?)<\/script>/);
if (!configScriptMatch) {
  console.error('[SECURITY FATAL] Missing inline window.__FM_CONFIG script for CSP hashing!');
  process.exit(1);
}
const normalizedConfigScript = configScriptMatch[1].replace(/\r\n?/g, '\n');
const configScriptHash = crypto.createHash('sha256').update(normalizedConfigScript, 'utf8').digest('base64');

// Verify all external <script> tags in index.html are pinned with SRI + crossorigin (no -latest/@latest)
const externalScripts = [...html.matchAll(/<script\b[^>]*\bsrc=["'](https?:\/\/[^"']+)["'][^>]*>/gi)];
for (const m of externalScripts) {
  const fullTag = m[0];
  const srcUrl = m[1];
  if (/(?:^|[@/-])latest(?:[./]|$)/i.test(srcUrl)) {
    console.error(`[SECURITY FATAL] Unpinned latest CDN script forbidden: ${srcUrl}`);
    process.exit(1);
  }
  if (!/\bintegrity=["']sha(?:256|384|512)-[^"']+["']/i.test(fullTag) || !/\bcrossorigin=["']anonymous["']/i.test(fullTag)) {
    console.error(`[SECURITY FATAL] External script missing SRI integrity or crossorigin="anonymous": ${srcUrl}`);
    process.exit(1);
  }
}

let apiConnectOrigin = '';
try {
  if (apiUrl) apiConnectOrigin = new URL(apiUrl).origin;
} catch {}
const connectSrcValues = ["'self'", apiConnectOrigin].filter(Boolean).join(' ');
const strictCspHeader = [
  "default-src 'self'",
  `script-src 'self' 'sha256-${configScriptHash}' https://cdn.jsdelivr.net`,
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  "font-src 'self' https://fonts.gstatic.com data:",
  "img-src 'self' data: blob:",
  `connect-src ${connectSrcValues}`,
  "worker-src 'self'",
  "manifest-src 'self'",
  "object-src 'none'",
  "base-uri 'self'",
  "form-action 'self'",
  "frame-ancestors 'none'",
].join('; ');

// Create _headers for security
fs.writeFileSync(path.join(distDir, '_headers'), `/*
  X-Frame-Options: DENY
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  Permissions-Policy: camera=(), microphone=(), geolocation=()
  Strict-Transport-Security: max-age=31536000; includeSubDomains; preload
  Content-Security-Policy: ${strictCspHeader}

/*.html
  Cache-Control: no-cache, no-store, must-revalidate

/sw.js
  Cache-Control: no-cache, no-store, must-revalidate

/manifest.json
  Cache-Control: public, max-age=86400

/assets/*.js
  Cache-Control: public, max-age=31536000, immutable

/assets/*.css
  Cache-Control: public, max-age=31536000, immutable
`);

// When building for staging, also populate clean dedicated staging output dir (dist-staging)
if (envName === 'staging') {
  const stagingDistDir = path.join(__dirname, 'dist-staging');
  if (fs.existsSync(stagingDistDir)) {
    fs.rmSync(stagingDistDir, { recursive: true, force: true });
  }
  fs.mkdirSync(path.join(stagingDistDir, 'assets'), { recursive: true });
  copyDir(publicDir, stagingDistDir);
  if (fs.existsSync(functionsDir)) {
    copyDir(functionsDir, path.join(stagingDistDir, 'functions'));
  }
  fs.copyFileSync(path.join(distDir, 'index.html'), path.join(stagingDistDir, 'index.html'));
  fs.copyFileSync(path.join(distDir, '_redirects'), path.join(stagingDistDir, '_redirects'));
  fs.copyFileSync(path.join(distDir, '_headers'), path.join(stagingDistDir, '_headers'));
  if (fs.existsSync(path.join(distDir, 'assets', activeBundleFilename))) {
    fs.copyFileSync(
      path.join(distDir, 'assets', activeBundleFilename),
      path.join(stagingDistDir, 'assets', activeBundleFilename)
    );
  }
  if (fs.existsSync(path.join(distDir, 'assets', 'meta.json'))) {
    fs.copyFileSync(
      path.join(distDir, 'assets', 'meta.json'),
      path.join(stagingDistDir, 'assets', 'meta.json')
    );
  }
  console.log(`✅ Staging isolated build complete! Output in /dist-staging (Active bundle: ${activeBundleFilename})`);
}

console.log('✅ Build complete! Output in /dist');
