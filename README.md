# 🏥 FM Operations — Facility Management System

Sistem manajemen operasional Facility Care berbasis web, dibangun di atas **Cloudflare Workers + D1 + Pages**. Berjalan sepenuhnya di edge — tanpa server, biaya hosting nyaris nol.

## Fitur Utama

| Modul | Keterangan |
|---|---|
| 📊 Dashboard | Statistik real-time, grafik tren, kontrak yang akan habis |
| 📅 Kalender | Visualisasi semua kegiatan per bulan |
| 👥 Karyawan | Master data karyawan + riwayat kontrak |
| 📋 Kontrak | Data PKWT, sisa hari otomatis, alert expiry |
| 🗓️ Jadwal Kegiatan | Timeline Inspeksi, GC, DC, Fogging per periode |
| ⚠️ Permasalahan | Log issues per cabang, tracking status & hari |
| 🤝 One on One | Catatan meeting one-on-one karyawan |
| 🎓 Training | Jadwal & rekap training karyawan |
| 🔄 Reliefer | Jadwal karyawan pengganti sementara |
| 🔍 Laporan Inspeksi | Score FC & SPV per cabang per periode |
| 🧹 Laporan GC/DC | Rekap General & Deep Cleaning |
| 💨 Rekap Fogging | Log kegiatan fogging per cabang |
| 📝 Laporan Basecamp | Rekap permasalahan dari Basecamp |
| 📖 SOP | Direktori SOP dengan link dokumen |
| ✅ Master Checklist | Checklist jobdesk karyawan |
| 📄 Master Form | Form management dengan akses publik opsional |
| 📦 Permintaan Barang | Form publik (tanpa login) untuk request chemical/alat |
| 🔐 Manajemen User | CRUD user + role-based access control |
| 🏢 Cabang | Data 70+ cabang |

---

## Tech Stack

```
Frontend  : Vanilla JS/HTML/CSS → Cloudflare Pages
Backend   : Cloudflare Workers (JS)
Database  : Cloudflare D1 (SQLite at edge)
Auth      : JWT (Web Crypto API, tanpa library)
CI/CD     : GitHub Actions → Cloudflare
```

---

## Panduan Deploy (Step by Step)

### Prasyarat
- Akun [Cloudflare](https://cloudflare.com) (free tier cukup)
- Akun [GitHub](https://github.com)
- Node.js 18+ terinstall di komputer
- `npm install -g wrangler` sudah dijalankan

---

### LANGKAH 1 — Clone & Push ke GitHub

```bash
# Buat repo baru di GitHub (misalnya: fm-operations)
# Lalu push project ini:

git init
git add .
git commit -m "feat: initial FM Operations system"
git remote add origin https://github.com/USERNAME/fm-operations.git
git push -u origin main
```

---

### LANGKAH 2 — Login ke Cloudflare via Wrangler

```bash
npx wrangler login
```

Browser akan terbuka. Login dengan akun Cloudflare Anda.

---

### LANGKAH 3 — Buat D1 Database

```bash
npx wrangler d1 create fm-operations-db
```

Output akan menampilkan `database_id`. Salin ID tersebut, lalu update file `api/wrangler.toml`:

```toml
[[d1_databases]]
binding = "DB"
database_name = "fm-operations-db"
database_id = "PASTE_ID_DISINI"   # ← ganti ini
```

---

### LANGKAH 4 — Jalankan Migrasi Database

```bash
# Local (untuk testing)
npx wrangler d1 execute fm-operations-db --file=schema/migrations.sql

# Remote (production — data tersimpan di Cloudflare)
npx wrangler d1 execute fm-operations-db --remote --file=schema/migrations.sql
```

Perintah ini akan membuat semua tabel dan mengisi data awal (cabang, roles, SOP seed).

---

### LANGKAH 5 — Buat KV Namespace (untuk session)

```bash
npx wrangler kv namespace create SESSIONS
```

Salin `id` dari output, update `api/wrangler.toml`:

```toml
[[kv_namespaces]]
binding = "SESSIONS"
id = "PASTE_KV_ID_DISINI"   # ← ganti ini
```

---

### LANGKAH 6 — Set JWT Secret

Buat file `api/.dev.vars` untuk development lokal:

```bash
cp api/.dev.vars.example api/.dev.vars
# Edit .dev.vars, isi JWT_SECRET dengan string acak yang panjang
```

Untuk production, set via Wrangler:

```bash
npx wrangler secret put JWT_SECRET --cwd api
# Ketik secret Anda (min. 32 karakter), lalu Enter
```

---

### LANGKAH 7 — Deploy Worker API

```bash
cd api
npx wrangler deploy
```

Catat URL Worker yang muncul, contoh:
```
https://fm-operations-api.USERNAME.workers.dev
```

---

### LANGKAH 8 — Build & Deploy Frontend

Update URL API di `frontend/wrangler.toml`:

```toml
[vars]
API_BASE_URL = "https://fm-operations-api.USERNAME.workers.dev"
```

Kemudian build dan deploy:

```bash
cd frontend
npm install
npm install -g esbuild

# Build (otomatis inject API URL)
API_BASE_URL="https://fm-operations-api.USERNAME.workers.dev" node build.js

# Deploy ke Cloudflare Pages
npx wrangler pages deploy dist --project-name=fm-operations
```

Saat pertama kali deploy Pages, Cloudflare akan membuat project baru. Konfirmasi dengan `Y`.

---

### LANGKAH 9 — Setup GitHub Secrets untuk CI/CD

Di GitHub repository → **Settings → Secrets and variables → Actions**, tambahkan:

| Secret Name | Nilai |
|---|---|
| `CLOUDFLARE_API_TOKEN` | API Token Cloudflare (buat di dashboard CF → My Profile → API Tokens → Create Token → "Edit Cloudflare Workers" template) |
| `CLOUDFLARE_ACCOUNT_ID` | Account ID Cloudflare (lihat di dashboard CF kanan bawah) |
| `API_BASE_URL` | `https://fm-operations-api.USERNAME.workers.dev` |

Setelah secrets ditambahkan, setiap push ke `main` akan otomatis deploy via GitHub Actions.

---

### LANGKAH 10 — Login Pertama

Akses URL Cloudflare Pages Anda (contoh: `https://fm-operations.pages.dev`)

**Kredensial default:**
```
Username : superadmin
Password : Admin@123
```

> ⚠️ **WAJIB ganti password** segera setelah login pertama di menu **Profil → Ganti Password**.

---

## Development Lokal

### Jalankan API (Worker) lokal

```bash
cd api
npm install
npx wrangler dev
# API berjalan di http://localhost:8787
```

### Jalankan Frontend lokal

```bash
cd frontend
npm install
# Build terlebih dahulu
API_BASE_URL="http://localhost:8787" node build.js
# Serve dist folder
npx serve dist -p 3000
# Buka http://localhost:3000
```

---

## Struktur Project

```
fm-operations/
├── api/                        # Cloudflare Worker (REST API)
│   ├── src/
│   │   ├── index.js            # Entry point, router utama
│   │   ├── routes/
│   │   │   ├── auth.js         # Login, logout, me, ganti password
│   │   │   ├── users.js        # CRUD users
│   │   │   ├── branches.js     # CRUD cabang
│   │   │   ├── employees.js    # CRUD karyawan
│   │   │   ├── contracts.js    # CRUD kontrak
│   │   │   ├── schedule.js     # CRUD jadwal kegiatan
│   │   │   ├── issues.js       # CRUD permasalahan
│   │   │   ├── one_on_one.js   # CRUD one on one
│   │   │   ├── training.js     # CRUD training
│   │   │   ├── relievers.js    # CRUD reliefer
│   │   │   ├── reports.js      # Laporan inspeksi/GC/DC/fogging/basecamp/supply
│   │   │   ├── misc.js         # SOP, checklist, forms, PIC list
│   │   │   └── dashboard.js    # Stats, kalender events, grafik
│   │   └── utils/
│   │       ├── auth.js         # JWT & password hashing (Web Crypto)
│   │       ├── response.js     # Helper response + CORS
│   │       └── pagination.js   # Pagination & query helpers
│   └── wrangler.toml
│
├── frontend/                   # Cloudflare Pages (SPA)
│   ├── src/
│   │   ├── app.js              # Entry point, layout, routing
│   │   ├── config.js           # API URL, token helpers
│   │   ├── router.js           # Hash-based router
│   │   ├── components/
│   │   │   ├── badges.js       # Status badges
│   │   │   ├── form.js         # Form builder
│   │   │   ├── modal.js        # Modal dialog
│   │   │   ├── table.js        # Data table + pagination
│   │   │   └── toast.js        # Notifikasi toast
│   │   └── pages/
│   │       ├── _crud.js        # Generic CRUD page builder
│   │       ├── login.js
│   │       ├── dashboard.js
│   │       ├── calendar.js
│   │       ├── employees.js
│   │       ├── contracts.js
│   │       ├── schedule.js
│   │       ├── issues.js
│   │       ├── one_on_one.js
│   │       ├── training.js
│   │       ├── relievers.js
│   │       ├── inspection_reports.js
│   │       ├── cleaning_reports.js
│   │       ├── fogging_reports.js
│   │       ├── basecamp_reports.js
│   │       ├── sop.js
│   │       ├── checklist.js
│   │       ├── forms.js
│   │       ├── users.js
│   │       ├── branches.js
│   │       └── profile.js
│   ├── public/
│   │   ├── index.html          # SPA entry point
│   │   ├── form.html           # Form publik (tanpa login)
│   │   ├── manifest.json       # PWA manifest
│   │   └── assets/
│   │       └── style.css       # Design system CSS
│   └── build.js                # Build script (esbuild)
│
├── schema/
│   └── migrations.sql          # D1 database schema + seed data
│
├── .github/
│   └── workflows/
│       └── deploy.yml          # CI/CD pipeline
│
└── package.json                # Monorepo root
```

---

## Role & Hak Akses

| Role | Keterangan |
|---|---|
| `superadmin` | Akses penuh semua modul + admin user |
| `admin` | CRUD semua modul, manajemen user |
| `manager` | Read semua + write operasional |
| `spv` | Read semua + write issues & one-on-one |
| `viewer` | Read only |

---

## API Endpoints

### Public (tanpa auth)
```
POST /api/auth/login
POST /api/reports/supply      ← Form permintaan barang
GET  /api/branches?all=1
GET  /api/sop
GET  /api/checklist
GET  /api/pic
```

### Protected (Bearer token)
```
GET|POST       /api/employees
GET|PUT|DELETE /api/employees/:id

GET|POST       /api/contracts
GET|PUT|DELETE /api/contracts/:id

GET|POST       /api/schedule
GET|PUT|DELETE /api/schedule/:id

GET|POST       /api/issues
GET|PUT|DELETE /api/issues/:id

GET|POST       /api/one-on-one
GET|POST       /api/training
GET|POST       /api/relievers

GET|POST       /api/reports/inspection
GET|POST       /api/reports/cleaning
GET|POST       /api/reports/fogging
GET|POST       /api/reports/basecamp

GET            /api/dashboard/stats
GET            /api/dashboard/calendar?month=YYYY-MM
GET            /api/dashboard/issues-summary

GET|POST       /api/users
GET|PUT|DELETE /api/users/:id
```

---

## Troubleshooting

**Error "YOUR_D1_DATABASE_ID" saat deploy API**
→ Update `database_id` di `api/wrangler.toml` dengan ID dari `wrangler d1 create`.

**CORS error di browser**
→ Pastikan `CORS_ORIGIN` di `api/wrangler.toml` atau `.dev.vars` sesuai URL frontend.

**Login gagal dengan credentials default**
→ Pastikan migrasi database sudah dijalankan (Langkah 4). Cek dengan:
```bash
npx wrangler d1 execute fm-operations-db --remote --command "SELECT username FROM users"
```

**GitHub Actions gagal**
→ Pastikan 3 secrets sudah ditambahkan: `CLOUDFLARE_API_TOKEN`, `CLOUDFLARE_ACCOUNT_ID`, `API_BASE_URL`.

**Kontrak "sisa hari" tidak muncul (-46217)**
→ Ada data dengan `start_date` atau `end_date` kosong di database. Edit kontrak tersebut dan isi tanggalnya.

---

## Lisensi

MIT — bebas digunakan dan dimodifikasi untuk keperluan internal perusahaan.
