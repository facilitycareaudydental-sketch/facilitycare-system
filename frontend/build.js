/**
 * Simple build script - bundles all JS modules into a single app.js
 * Uses esbuild if available, otherwise generates a bundle manually
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const srcDir = path.join(__dirname, 'src');
const publicDir = path.join(__dirname, 'public');
const distDir = path.join(__dirname, 'dist');

// Ensure dist exists
if (!fs.existsSync(distDir)) fs.mkdirSync(distDir, { recursive: true });
if (!fs.existsSync(path.join(distDir, 'assets'))) fs.mkdirSync(path.join(distDir, 'assets'), { recursive: true });

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

// Replace API URL placeholder in index.html
const apiUrl = process.env.API_BASE_URL || 'https://fm-operations-api.YOUR_SUBDOMAIN.workers.dev';
let html = fs.readFileSync(path.join(distDir, 'index.html'), 'utf8');
html = html.replace('__API_BASE_URL__', apiUrl);
fs.writeFileSync(path.join(distDir, 'index.html'), html);

// Also update form.html
let formHtml = fs.readFileSync(path.join(distDir, 'form.html'), 'utf8');
formHtml = formHtml.replace('https://fm-operations-api.YOUR_SUBDOMAIN.workers.dev', apiUrl);
fs.writeFileSync(path.join(distDir, 'form.html'), formHtml);

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

if (esbuildCmd) {
  console.log(`Bundling with esbuild (${esbuildCmd})...`);
  try {
    // Use content hash in filename so Cloudflare CDN always serves fresh file
    const metafile = path.join(distDir, 'assets', 'meta.json');
    execSync(
      `${esbuildCmd} ${path.join(srcDir, 'app.js')} --bundle --outdir=${path.join(distDir, 'assets')} --entry-names=[name]-[hash] --format=esm --minify --platform=browser --metafile=${metafile}`,
      { stdio: 'inherit' }
    );

    // Read metafile to find actual output filename
    const meta = JSON.parse(fs.readFileSync(metafile, 'utf8'));
    const outputs = Object.keys(meta.outputs).filter(f => f.endsWith('.js') && !f.endsWith('.css'));
    const outFilename = outputs.length ? path.basename(outputs[0]) : 'app.js';
    console.log(`✅ Bundle created: dist/assets/${outFilename}`);

    // Patch index.html to reference the hashed filename
    let htmlContent = fs.readFileSync(path.join(distDir, 'index.html'), 'utf8');
    htmlContent = htmlContent.replace(/\/assets\/app(?:-[a-z0-9]+)?\.js/, `/assets/${outFilename}`);
    fs.writeFileSync(path.join(distDir, 'index.html'), htmlContent);
    console.log(`✅ index.html patched: /assets/${outFilename}`);
  } catch (err) {
    console.error('esbuild bundle failed, falling back to source copy:', err.message);
    esbuildCmd = null;
  }
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

// Create _headers for security
fs.writeFileSync(path.join(distDir, '_headers'), `/*
  X-Frame-Options: DENY
  X-Content-Type-Options: nosniff
  Referrer-Policy: strict-origin-when-cross-origin
  Permissions-Policy: camera=(), microphone=(), geolocation=()

/*.html
  Cache-Control: no-cache, no-store, must-revalidate

/assets/*.js
  Cache-Control: public, max-age=31536000, immutable

/assets/*.css
  Cache-Control: public, max-age=31536000, immutable
`);

console.log('✅ Build complete! Output in /dist');
