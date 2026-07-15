#!/bin/bash
# FM Operations — Local Setup Script
# Run: bash setup.sh

set -e

echo "🏥 FM Operations — Setup"
echo "========================"

# Check Node.js
if ! command -v node &> /dev/null; then
  echo "❌ Node.js tidak ditemukan. Install dari https://nodejs.org"
  exit 1
fi
echo "✅ Node.js $(node -v)"

# Check wrangler
if ! command -v wrangler &> /dev/null; then
  echo "📦 Menginstall Wrangler..."
  npm install -g wrangler@3
fi
echo "✅ Wrangler $(wrangler --version)"

# Install dependencies
echo ""
echo "📦 Install dependencies..."
npm install
cd api && npm install && cd ..
cd frontend && npm install && cd ..
echo "✅ Dependencies terinstall"

# Setup .dev.vars if not exists
if [ ! -f "api/.dev.vars" ]; then
  cp api/.dev.vars.example api/.dev.vars
  echo "✅ api/.dev.vars dibuat dari template"
  echo "⚠️  Edit api/.dev.vars dan isi JWT_SECRET Anda"
fi

echo ""
echo "📋 Langkah selanjutnya:"
echo "  1. Edit api/wrangler.toml — isi database_id dan kv namespace id"
echo "  2. Edit api/.dev.vars — isi JWT_SECRET"
echo "  3. Jalankan migrasi: npx wrangler d1 execute fm-operations-db --file=schema/migrations.sql"
echo "  4. Jalankan API: cd api && npx wrangler dev"
echo "  5. Build frontend: cd frontend && API_BASE_URL=http://localhost:8787 node build.js"
echo "  6. Serve frontend: cd frontend && npx serve dist -p 3000"
echo ""
echo "📖 Baca README.md untuk panduan deploy lengkap ke Cloudflare."
