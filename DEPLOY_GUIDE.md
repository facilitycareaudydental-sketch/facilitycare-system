# 🚀 Panduan Deploy FM Operations ke Cloudflare

**Repo GitHub:** https://github.com/facilitycareaudydental-sketch/facilitycare-system

---

## Prasyarat

| Kebutuhan | Link Download |
|---|---|
| Node.js 18+ (LTS) | https://nodejs.org |
| Git | https://git-scm.com |
| Akun Cloudflare (gratis) | https://cloudflare.com |

---

## LANGKAH 1 — Buat Akun Cloudflare

1. Buka https://cloudflare.com → **Sign Up**
2. Daftar dengan email `facilitycare.audydental@gmail.com`
3. Verifikasi email
4. Setelah login, catat **Account ID** di pojok kanan bawah dashboard

---

## LANGKAH 2 — Install Wrangler

Buka **PowerShell** atau **Command Prompt**, jalankan:

```bash
npm install -g wrangler@3
wrangler --version
```

---

## LANGKAH 3 — Login Wrangler ke Cloudflare

```bash
wrangler login
```

Browser terbuka → klik **Allow**. Selesai.

---

## LANGKAH 4 — Clone Repo dari GitHub

```bash
git clone https://github.com/facilitycareaudydental-sketch/facilitycare-system.git
cd facilitycare-system
```

---

## LANGKAH 5 — Buat D1 Database

```bash
wrangler d1 create fm-operations-db
```

Output contoh:
```
✅ Successfully created DB 'fm-operations-db'
database_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

Buka file `api/wrangler.toml`, ganti `database_id`:

```toml
[[d1_databases]]
binding = "DB"
database_name = "fm-operations-db"
database_id = "PASTE_ID_DISINI"
```

---

## LANGKAH 6 — Jalankan Migrasi Database

```bash
wrangler d1 execute fm-operations-db --remote --file=schema/migrations.sql
```

Ketik `Y` jika diminta konfirmasi.

> Perintah ini membuat semua tabel + mengisi data awal:
> 70+ cabang, roles, SOP, PIC list, dll.

---

## LANGKAH 7 — Set JWT Secret

```bash
cd api
wrangler secret put JWT_SECRET
```

Ketik secret yang kuat (min. 32 karakter), contoh:
```
fm-ops-audy-dental-secret-2026-secure-key
```
Tekan **Enter**.

---

## LANGKAH 8 — Deploy Worker API

```bash
# Masih di folder api
wrangler deploy
```

Catat URL yang muncul:
```
✅ Deployed: https://fm-operations-api.NAMAKAMU.workers.dev
```

---

## LANGKAH 9 — Update URL API di Frontend

Buka `frontend/wrangler.toml`, ganti URL:

```toml
[vars]
API_BASE_URL = "https://fm-operations-api.NAMAKAMU.workers.dev"
```

Buka `frontend/public/form.html`, ganti baris ini (sekitar baris 110):

```js
const API = window.__FM_CONFIG?.API_BASE_URL
  || 'https://fm-operations-api.NAMAKAMU.workers.dev';
```

---

## LANGKAH 10 — Build & Deploy Frontend

```bash
cd frontend
npm install
npm install -g esbuild
```

**Windows PowerShell:**
```powershell
$env:API_BASE_URL="https://fm-operations-api.NAMAKAMU.workers.dev"
node build.js
```

**Mac / Linux:**
```bash
API_BASE_URL="https://fm-operations-api.NAMAKAMU.workers.dev" node build.js
```

Kemudian deploy:
```bash
wrangler pages deploy dist --project-name=fm-operations
```

Saat pertama kali → ketik `Y` untuk buat project baru.

URL frontend setelah deploy:
```
https://fm-operations.pages.dev
```

---

## LANGKAH 11 — Setup GitHub Actions (CI/CD Otomatis)

Setiap push ke `main` akan otomatis deploy. Setup secret dulu:

1. Buka: https://github.com/facilitycareaudydental-sketch/facilitycare-system
2. Klik **Settings** → **Secrets and variables** → **Actions**
3. Klik **New repository secret**, tambahkan 3 secret berikut:

| Nama Secret | Cara Mendapatkan |
|---|---|
| `CLOUDFLARE_API_TOKEN` | Cloudflare Dashboard → My Profile → API Tokens → **Create Token** → pilih template **"Edit Cloudflare Workers"** |
| `CLOUDFLARE_ACCOUNT_ID` | Cloudflare Dashboard → pojok kanan bawah halaman |
| `API_BASE_URL` | URL dari Langkah 8, contoh: `https://fm-operations-api.NAMAKAMU.workers.dev` |

Setelah 3 secret disimpan, workflow `deploy.yml` akan otomatis berjalan setiap push ke `main`.

---

## LANGKAH 12 — Login Pertama ke Aplikasi

Buka: **https://fm-operations.pages.dev**

```
Username : superadmin
Password : Admin@123
```

> ⚠️ Segera ganti password di **Profil → Ganti Password** setelah login pertama!

---

## Akses Form Publik (Tanpa Login)

Bagikan link ini ke karyawan cabang untuk request barang & chemical:

```
https://fm-operations.pages.dev/form.html
```

Tidak perlu username/password.

---

## Struktur URL Aplikasi

| Halaman | URL |
|---|---|
| Login | `https://fm-operations.pages.dev` |
| Dashboard | `https://fm-operations.pages.dev/#/dashboard` |
| Karyawan | `https://fm-operations.pages.dev/#/employees` |
| Kontrak | `https://fm-operations.pages.dev/#/contracts` |
| Jadwal | `https://fm-operations.pages.dev/#/schedule` |
| Permasalahan | `https://fm-operations.pages.dev/#/issues` |
| One on One | `https://fm-operations.pages.dev/#/one-on-one` |
| Training | `https://fm-operations.pages.dev/#/training` |
| Reliefer | `https://fm-operations.pages.dev/#/relievers` |
| Laporan Inspeksi | `https://fm-operations.pages.dev/#/reports/inspection` |
| Laporan GC/DC | `https://fm-operations.pages.dev/#/reports/cleaning` |
| Rekap Fogging | `https://fm-operations.pages.dev/#/reports/fogging` |
| Laporan Basecamp | `https://fm-operations.pages.dev/#/reports/basecamp` |
| Permintaan Barang | `https://fm-operations.pages.dev/#/reports/supply` |
| SOP | `https://fm-operations.pages.dev/#/sop` |
| Master Checklist | `https://fm-operations.pages.dev/#/checklist` |
| Master Form | `https://fm-operations.pages.dev/#/forms` |
| Manajemen User | `https://fm-operations.pages.dev/#/users` |
| Cabang | `https://fm-operations.pages.dev/#/branches` |
| Kalender | `https://fm-operations.pages.dev/#/calendar` |
| Form Publik | `https://fm-operations.pages.dev/form.html` |

---

## Role & Hak Akses

| Role | Hak Akses |
|---|---|
| `superadmin` | Semua akses + admin user |
| `admin` | CRUD semua modul + manajemen user |
| `manager` | Read semua + write operasional |
| `spv` | Read semua + write issues & one-on-one |
| `viewer` | Read only |

---

## Perintah Sehari-hari

```bash
# Push update kode → CI/CD otomatis deploy
git add .
git commit -m "update: keterangan perubahan"
git push

# Deploy manual API
cd api && wrangler deploy

# Deploy manual frontend
cd frontend
node build.js
wrangler pages deploy dist --project-name=fm-operations

# Lihat log API real-time
cd api && wrangler tail

# Cek data database
wrangler d1 execute fm-operations-db --remote --command "SELECT COUNT(*) FROM employees"
wrangler d1 execute fm-operations-db --remote --command "SELECT COUNT(*) FROM contracts"

# Tambah user baru via database langsung
wrangler d1 execute fm-operations-db --remote --command \
  "INSERT INTO users (username,email,password_hash,full_name,role) VALUES ('berlin','berlin@fm.com','HASH','Berlin Manager','manager')"
```

---

## Troubleshooting

**❌ Error: YOUR_D1_DATABASE_ID**
→ Belum diisi `database_id` di `api/wrangler.toml`. Ikuti Langkah 5.

**❌ CORS error di browser**
→ Pastikan URL frontend sudah benar di `api/wrangler.toml`:
```toml
[vars]
CORS_ORIGIN = "https://fm-operations.pages.dev"
```
Lalu `wrangler deploy` ulang di folder `api`.

**❌ Login gagal (credentials default tidak bisa)**
→ Migrasi belum dijalankan. Ulang Langkah 6.

**❌ GitHub Actions gagal**
→ Cek 3 secret sudah ditambahkan (Langkah 11).

**❌ Kontrak sisa hari tampil -46217**
→ Ada data kontrak tanpa tanggal. Edit data tersebut, isi tanggal mulai & selesai.

**❌ Frontend tampil blank / loading terus**
→ Periksa `API_BASE_URL` di `frontend/public/index.html` sudah benar, bukan masih `__API_BASE_URL__`.

---

## ⚠️ Keamanan — Wajib Dilakukan

1. **Ganti password superadmin** setelah login pertama
2. **Revoke token GitHub lama** di https://github.com/settings/tokens (token yang dipakai saat setup)
3. **Buat token GitHub baru** jika perlu CI/CD
4. **Jangan commit** file `.dev.vars` atau file yang berisi token/secret ke GitHub

---

*Dibuat untuk: FM Operations - Facility Care Audy Dental*
*Stack: Cloudflare Workers + D1 + Pages | Repo: facilitycareaudydental-sketch/facilitycare-system*
