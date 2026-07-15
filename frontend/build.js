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
try {
  execSync('npx esbuild --version', { stdio: 'ignore' });
  console.log('Bundling with esbuild...');
  execSync(
    `npx esbuild ${path.join(srcDir, 'app.js')} --bundle --outfile=${path.join(distDir, 'assets', 'app.js')} --format=esm --minify`,
    { stdio: 'inherit' }
  );
} catch {
  console.log('esbuild not found, copying source files...');
  // Copy src files as-is for development (works with native ESM in modern browsers)
  copyDir(srcDir, path.join(distDir, 'src'));
  // Update script tag to point to src
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

/assets/*
  Cache-Control: public, max-age=31536000, immutable
`);

console.log('✅ Build complete! Output in /dist');
