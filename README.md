# 🏥 Facility Care Management System (FCMS) — Audy Dental

Sistem manajemen operasional Facility Care berbasis web untuk jaringan klinik **Audy Dental**, dibangun di atas arsitektur serverless **Cloudflare Workers + D1 + Pages**. Berjalan sepenuhnya di edge computing dengan keandalan tinggi dan latensi rendah.

> [!IMPORTANT]
> **Status Lingkungan Saat Ini: STAGING (Pra-Cutover)**
> - **FCMS Lama** masih merupakan sistem operasional harian yang aktif digunakan.
> - **FCMS Baru (Staging)** berada dalam tahap verifikasi pra-cutover (testing terisolasi, audit rekonsiliasi, dan sinkronisasi mode `DRY_RUN`).
> - **Struktur Biaya:** Menggunakan Cloudflare Workers Free Tier dengan batas kuota operasional harian (100.000 requests/hari untuk Workers, 5M row reads/hari untuk D1), bukan sistem tanpa batas biaya.

---

## 24 Menu & Modul Resmi

FCMS Baru Audy Dental mengorganisasikan 24 modul resmi yang dikelompokkan ke dalam 6 kategori operasional:

| No | Kategori | Menu / Modul | Rute Frontend | Deskripsi |
|---|---|---|---|---|
| 1 | **Utama** | Dashboard | `#/dashboard` | Statistik real-time, ringkasan KPI, grafik tren, alert kontrak |
| 2 | **Utama** | Kalender | `#/calendar` | Visualisasi kalender seluruh agenda dan kegiatan operasional |
| 3 | **SDM** | Master Karyawan | `#/employees` | Master data karyawan internal Facility Care & Security **hanya area Management HQ (`000. Management`)** |
| 4 | **SDM** | Data Kontrak | `#/contracts` | Tracking kontrak kerja (PKWT), hitungan sisa masa aktif, alert expiry |
| 5 | **SDM** | Data SP | `#/sp` | Pencatatan dan monitoring Surat Peringatan karyawan |
| 6 | **SDM** | Data Mutasi | `#/mutasi` | Riwayat perpindahan dan mutasi penempatan karyawan |
| 7 | **SDM** | Data Lembur | `#/overtime` | Log pengajuan, verifikasi, dan rekapitulasi lembur |
| 8 | **Operasional** | Time Line / Jadwal | `#/timeline` | Jadwal kegiatan rutin (Inspeksi, GC, DC, Fogging) per periode |
| 9 | **Operasional** | Migrasi OS | `#/os-migrations` | Tracking program peralihan status karyawan OS ke internal |
| 10 | **Operasional** | Permasalahan / Issues | `#/issues` | Pelaporan komplain, kendala operasional cabang, dan tracking status |
| 11 | **Operasional** | Training | `#/training` | Jadwal dan rekapitulasi pelatihan kompetensi karyawan |
| 12 | **Operasional** | One on One | `#/one-on-one` | Log konseling dan evaluasi berkala karyawan |
| 13 | **Operasional** | Data Reliever | `#/relievers` | Jadwal backup dan penempatan tenaga pengganti sementara |
| 14 | **Laporan** | Report Inspeksi Hygiene | `#/reports/inspection` | Evaluasi skor kebersihan FC & SPV per cabang per periode |
| 15 | **Laporan** | Report GC-DC | `#/reports/cleaning` | Rekapitulasi pelaksanaan General Cleaning & Deep Cleaning |
| 16 | **Laporan** | Report Fogging | `#/reports/fogging` | Log jadwal dan realisasi fogging disinfeksi cabang |
| 17 | **Laporan** | Rekap Laporan Basecamp | `#/reports/basecamp` | Log penanganan komplain dan isu dari channel Basecamp |
| 18 | **Referensi** | SOP Facility Care | `#/sop-fc` | Direktori dokumen Standar Operasional Prosedur kebersihan |
| 19 | **Referensi** | SOP Security | `#/sop-security` | Direktori dokumen SOP pengamanan dan operasional security |
| 20 | **Referensi** | Master Dokumen / Form | `#/checklist` | Formulir kerja dan checklist harian operasional klinik |
| 21 | **Referensi** | Master Standar Hygiene | `#/hygiene-standards` | Parameter dan standar baku penilaian kebersihan |
| 22 | **Referensi** | Master Standar DAC | `#/dac` | Standar Dental Assistant Care dan fasilitas klinik |
| 23 | **Admin** | Manajemen Pengguna | `#/users` | Administrasi akun pengguna, penetapan role, dan reset akses |
| 24 | **Admin** | Data Cabang | `#/branches` | Master data 70+ klinik Audy Dental beserta PIC dan wilayah |

> [!NOTE]
> Modul administratif tambahan meliputi: **Audit Log** (`#/audit-logs`), **Import Data Awal** (`#/settings/import`), **Sync Dashboard** (`#/sync`), dan **Profil Pengguna** (`#/profile`).

---

## Role & Hak Akses (RBAC)

FCMS Baru menerapkan pembatasan akses berbasis peran resmi yang ketat:

| Role Resmi | Level Akses | Keterangan |
|---|---|---|
| `superadmin` | Akses Penuh (Administrator) | Memiliki hak penuh atas seluruh modul, konfigurasi sistem, manajemen user, audit log, dan sinkronisasi database. Dibatasi maksimal 2 akun aktif. |
| `fc_specialist_leader` | Operasional & Monitoring | Akses membaca (`read`) seluruh modul non-admin, serta hak tulis (`write`) **khusus** pada modul **Overtime (`overtime`)** dan **Permasalahan (`issues`)** (tanpa hak hapus `delete` atau akses modul Admin). Dibatasi maksimal 2 akun aktif. Role legacy (`admin`, `supervisor`, `editor_khusus`, `input_lembur`, `viewer`, `user`) ditolak fail-closed (HTTP 403). |

> [!WARNING]
> **Catatan Penting Mengenai "Reliever":**  
> `reliever` **BUKAN** merupakan role sistem (bukan akun pengguna/login). Reliever adalah **entitas data operasional** untuk mencatat jadwal petugas pengganti sementara di cabang klinik. Jangan pernah membuat user dengan role `reliever`.

---

## Tech Stack

- **Frontend:** Vanilla JavaScript (SPA), Semantic HTML5, CSS Design System, Responsive & Mobile PWA, di-host pada **Cloudflare Pages**.
- **Backend API:** **Cloudflare Workers** (RESTful API berbasis Web Standards & Web Crypto API).
- **Database:** **Cloudflare D1** (SQLite at Edge) dengan skema kanonikal v2 + migrasi v2 (37 tabel baseline / **38 tabel pasca-0017**, 114 index baseline / **122 index pasca-0018**).
- **Autentikasi & Keamanan:** JWT HS256 (Web Crypto API) dalam cookie `HttpOnly` + proteksi double-submit CSRF (`X-CSRF-Token`) + in-memory token fallback (tanpa `localStorage`), password hashing PBKDF2-SHA256 (100.000 iterasi), Content Security Policy ketat (`script-src` berbasis SHA-256 hash tanpa `'unsafe-inline'`), CDN pinned dengan SRI SHA-384, in-isolate rate limiting (per-isolate Map — penegakan global lintas-isolate bergantung pada Cloudflare WAF Edge Rate Limiting), serta immutable audit logging.
- **Data Synchronization:** Bidirectional Google Sheets Sync Engine dengan perlindungan fail-closed (`SYNC_MODE=DRY_RUN` di Staging).

---

## Panduan Setup & Deployment Lokal (Isolated Test Environment)

### Prasyarat
- Node.js versi 18 LTS atau 20+ LTS
- Akun Cloudflare (untuk manajemen Workers & D1 saat cutover)
- Cloudflare Wrangler CLI (`npm install -g wrangler`)

---

### Langkah 1 — Instalasi Dependensi

```bash
# Clone repository
git clone https://github.com/USERNAME/facilitycare-system.git
cd facilitycare-system

# Install dependensi Worker API
cd api
npm install

# Install dependensi Frontend
cd ../frontend
npm install
cd ..
```

---

### Langkah 2 — Konfigurasi Environment & Keamanan

Buat file environment lokal di `api/.dev.vars`:

```bash
cp api/.dev.vars.example api/.dev.vars
```

Konfigurasikan variabel rahasia dengan nilai acak yang aman:
- `JWT_SECRET`: Minimal 32 karakter string kriptografis aman.
- `SYNC_MODE`: Setel ke `DRY_RUN` untuk lingkungan pengujian dan Staging.
- `ENVIRONMENT`: Setel ke `development` atau `staging`.

> [!CAUTION]
> **Kebijakan Kredensial:**  
> Sistem ini **TIDAK MENGGUNAKAN** default password yang di-hardcode. Akun awal wajib diinisialisasi melalui environment variable atau script provisioning yang aman (`scripts/provision_final_users.mjs`). Jangan pernah menyimpan password atau token ke dalam repository.

---

### Langkah 3 — Inisialisasi Database (Baseline Kanonikal v2)

> [!CAUTION]
> **PERINGATAN KERAS DATABASE MIGRATION:**  
> File `schema/migrations.sql` sudah **USANG (DEPRECATED)** dan **DILARANG DIGUNAKAN**. Menjalankan file tersebut akan merusak struktur integritas database.
>
> Selalu gunakan skema kanonikal v2 resmi:
> - **Inisialisasi Database Bersih:** `schema/fresh_canonical_baseline_v2.sql`
> - **Migrasi Bertahap:** Direktori `schema/migrations_v2/`

Untuk inisialisasi lokal via Wrangler:

```bash
# Menjalankan baseline v2 pada database D1 lokal
npx wrangler d1 execute fcms-staging-db --local --file=schema/fresh_canonical_baseline_v2.sql
```

Untuk memverifikasi integritas 114 index kanonikal baseline v2 (atau 122 index pasca-migrasi `migrations_v2`):

```bash
node scripts/canonical-index-validator.mjs --baseline=v2
node scripts/canonical-index-validator.mjs --baseline=v2-migrated
```

---

### Langkah 4 — Menjalankan Worker API & Frontend Lokal

**Menjalankan API lokal:**
```bash
cd api
npx wrangler dev
# API akan berjalan pada http://localhost:8787
```

**Membangun dan menjalankan Frontend:**
```bash
cd frontend
API_BASE_URL="http://localhost:8787" node build.js
npx serve dist -p 3000
# Akses aplikasi pada http://localhost:3000
```

---

## Struktur Direktori Repository

```
facilitycare-system/
├── api/                             # Cloudflare Worker (Edge REST API)
│   ├── src/
│   │   ├── index.js                 # Router utama & security headers
│   │   ├── routes/                  # Endpoint handlers (auth, employees, issues, dll)
│   │   └── utils/                   # Helper (auth, sync_engine, sync_mapper, occ)
│   ├── wrangler.toml                # Konfigurasi Workers & D1
│   └── package.json
│
├── frontend/                        # Cloudflare Pages (SPA & PWA)
│   ├── src/
│   │   ├── app.js                   # Layout, router mounting, 24 menu rendering
│   │   ├── config.js                # Definisi modul dan konfigurasi API
│   │   ├── router.js                # Client-side hash routing
│   │   ├── components/              # UI components (modal, toast, form, table)
│   │   └── pages/                   # 24 Halaman aplikasi
│   ├── public/                      # Static assets & PWA manifest
│   └── build.js                     # Script kompilasi esbuild
│
├── schema/
│   ├── fresh_canonical_baseline_v2.sql  # Schema kanonikal resmi v2 (37 tabel base, 114 index)
│   ├── migrations/                  # Jalur upgrade V1 (0001..0018 + production_remediation_runner.mjs & variants/)
│   ├── migrations_v2/               # Migrasi berurutan pasca-v2 (0011..0018; 0014 pre-baked di v2)
│   └── canonical_baseline_v1.sql    # Referensi historis baseline v1 (read-only)
│
├── scripts/                         # Security & operational tooling
│   ├── canonical-index-validator.mjs# Validator index kanonikal (v1=79, v2=114, v2-migrated=122)
│   ├── canonical_index_manifest.json# Manifest integritas index v2 (114 index)
│   ├── canonical_index_manifest_v2_migrated.json # Manifest pasca-migrasi v2 (122 index)
│   ├── d1-migrator.mjs              # Local SQLite migration engine (mendukung chain v1 & v2)
│   ├── restore-layer-driver.mjs     # Topological 3-layer restore driver (38 tabel)
│   ├── safe-synthetic-cleanup.mjs   # Identity-verified synthetic data cleanup engine
│   └── backup-validator.mjs         # Validator backup database offline
│
├── reports/                         # Laporan audit independen & verifikasi
└── package.json
```

---

## Ringkasan API Endpoints Utama

### Publik (Tanpa Autentikasi)
```http
POST /api/auth/login                 # Autentikasi user & penerbitan JWT
POST /api/reports/supply             # Form publik pengajuan supply/chemical
GET  /api/branches?all=1             # Referensi daftar cabang aktif
```

### Terproteksi (Bearer Token Wajib)
```http
# Operasional SDM
GET|POST        /api/employees       # Data Karyawan (D1 operational fields)
GET|PUT|DELETE  /api/employees/:id
GET|POST        /api/contracts       # Data Kontrak PKWT
GET|POST        /api/overtime        # Data Lembur Karyawan

# Operasional Lapangan & Laporan
GET|POST        /api/schedule        # Timeline Jadwal Kegiatan
GET|POST        /api/issues          # Permasalahan Cabang
GET|POST        /api/reports/*       # Laporan Inspeksi, GC/DC, Fogging, Basecamp

# Sinkronisasi & Webhook
POST            /api/sync/webhook    # Inbound Google Sheets Webhook (DRY_RUN / LIVE)
GET             /api/sync/status     # Status antrean sinkronisasi & outbox

# Administrasi (Superadmin Only)
GET|POST        /api/users           # Manajemen Akun Pengguna
GET             /api/audit-logs      # Audit Trail Aktivitas Pengguna
```

---

## Verifikasi Kepatuhan & Prosedur Pra-Cutover

Sebelum cutover Production dilakukan:
1. Pastikan seluruh 10 uji skenario pada `test_ws01_remediation.mjs` berstatus **PASS**.
2. Pastikan validator index kanonikal `node scripts/canonical-index-validator.mjs --baseline=v2` berstatus **PASS (114/114 verified)**.
3. Pastikan `SYNC_MODE` tetap berada pada mode `DRY_RUN` hingga owner memberikan persetujuan resmi.
4. Jangan pernah melakukan perubahan langsung pada sistem FCMS Lama yang sedang berjalan.

---

## Lisensi & Hak Cipta

Hak Cipta © 2026 PT Audy Dental Indonesia. Seluruh hak cipta dilindungi undang-undang.  
Sistem ini merupakan perangkat lunak operasional internal perusahaan.
