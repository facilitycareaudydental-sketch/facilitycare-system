const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('=== Starting Root Build Process ===');

// 1. Run the frontend build
try {
  console.log('Executing frontend build script...');
  execSync('node frontend/build.js', { stdio: 'inherit' });
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
    console.log('=== Root Build Completed Successfully! ===');
  } else {
    console.error('Error: Build output folder frontend/dist not found.');
    process.exit(1);
  }
} catch (error) {
  console.error('Error copying build output to root directory:', error);
  process.exit(1);
}
