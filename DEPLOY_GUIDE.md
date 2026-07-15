# 🚀 Panduan Deploy FM Operations via Cloudflare Web Dashboard

Panduan ini menjelaskan cara mendeploy seluruh sistem **FM Operations** secara langsung melalui antarmuka web (Cloudflare Dashboard & GitHub Web), tanpa memerlukan penggunaan terminal (CLI/Wrangler) sama sekali di komputer Anda.

---

## 📋 Prasyarat & Kebutuhan

Sebelum memulai, pastikan Anda memiliki:
1.  Akun **GitHub** yang berisi repository proyek ini.
2.  Akun **Cloudflare** (gratis).

---

## 🗄️ LANGKAH 1 — Membuat & Menginisialisasi D1 Database

D1 adalah database relasional SQLite serverless dari Cloudflare. Anda bisa membuatnya langsung dari dashboard:

1.  Login ke [Cloudflare Dashboard](https://dash.cloudflare.com).
2.  Pada menu navigasi sebelah kiri, klik **Workers & Pages** → **D1**.
3.  Klik tombol **Create Database** (atau **Create** lalu pilih **D1**).
4.  Masukkan nama database: `fm-operations-db`
5.  Klik **Create**.
6.  Setelah database dibuat, masuk ke tab **Console** di dalam halaman database tersebut.
7.  Buka file [migrations.sql](file:///C:/Users/Facility%20Care/.gemini/antigravity/scratch/facilitycare-system/schema/migrations.sql), **salin seluruh isi kode SQL** di dalamnya, lalu **tempel (paste)** ke dalam kolom input D1 Console di dashboard web.
8.  Klik **Execute** untuk membuat semua tabel dan memasukkan data awal (cabang, SOP, roles, PIC list).

---

## 🔑 LANGKAH 2 — Membuat KV Namespace (Session Store)

KV Namespace digunakan untuk melacak sesi pengguna yang aktif:

1.  Di menu navigasi kiri, masuk ke **Workers & Pages** → **KV**.
2.  Klik tombol **Create Namespace**.
3.  Masukkan nama namespace: `SESSIONS`
4.  Klik **Add**.

---

## 🔌 LANGKAH 3 — Mendeploy Worker (Backend API) via GitHub Integration

Cloudflare dapat terhubung langsung ke GitHub Anda untuk mendeploy Workers secara otomatis:

1.  Kembali ke menu **Workers & Pages** → click **Create** → pilih **Worker**.
2.  Beri nama Worker Anda (misal: `fm-operations-api`), lalu klik **Deploy**. (Ini akan membuat Worker kosong sementara).
3.  Setelah selesai dideploy, masuk ke halaman Worker tersebut, pilih tab **Settings** → **Build & deployments**.
4.  Di bagian **Integrations**, klik **Connect GitHub**.
5.  Otorisasi Cloudflare ke akun GitHub Anda, lalu pilih repository `facilitycare-system`.
6.  Atur konfigurasi deployment:
    *   **Production Branch**: `main`
    *   **Root Directory**: `api`
7.  Klik **Save and Deploy**. Cloudflare akan mulai menarik kode dari subfolder `api` repository Anda dan mendeploynya.

---

## 🔗 LANGKAH 4 — Mengonfigurasi Environment Variables & Bindings Worker

Worker API membutuhkan akses ke database D1, KV, dan JWT Secret. Mari kita sambungkan:

1.  Buka Worker `fm-operations-api` yang telah dibuat.
2.  Masuk ke tab **Settings** → **Variables**.
3.  Di bagian **Environment Variables**, klik **Add variable**:
    *   **Name**: `JWT_SECRET`
    *   **Type**: Pilih **Secret** (supaya aman/terenkripsi).
    *   **Value**: Masukkan teks rahasia acak yang panjang (misalnya: `fm-ops-audy-dental-secret-2026-secure-key`).
4.  Di bagian **Bindings**, klik **Add binding**:
    *   Pilih **D1 database binding**:
        *   **Variable name**: `DB` (harus huruf besar semua).
        *   **D1 database**: Pilih `fm-operations-db` dari dropdown.
    *   Klik **Add binding** lagi, lalu pilih **KV namespace binding**:
        *   **Variable name**: `SESSIONS` (harus huruf besar semua).
        *   **KV namespace**: Pilih `SESSIONS` dari dropdown.
5.  Klik **Save and Deploy** di bagian bawah halaman untuk menerapkan perubahan konfigurasi.

---

## 🌐 LANGKAH 5 — Mendeploy Frontend via Cloudflare Pages (Git Integration)

Untuk frontend (antarmuka web), kita akan menghubungkannya langsung ke GitHub dan mengatur proses kompilasi otomatis:

1.  Masuk ke **Workers & Pages** → klik **Create** → pilih tab **Pages** (bukan Workers).
2.  Klik **Connect to Git**.
3.  Pilih repository `facilitycare-system` dari akun GitHub Anda, lalu klik **Begin setup**.
4.  Atur konfigurasi build berikut:
    *   **Project name**: `fm-operations` (ini akan menjadi bagian dari URL Anda, misal `fm-operations.pages.dev`).
    *   **Production branch**: `main`
    *   **Framework preset**: Pilih **None** (atau biarkan default).
    *   **Root directory**: `frontend` (sangat penting agar Cloudflare masuk ke subfolder frontend).
    *   **Build command**: `node build.js`
    *   **Build output directory**: `dist`
5.  Di bagian **Environment variables (advanced)**, tambahkan variabel berikut:
    *   **Variable name**: `API_BASE_URL`
    *   **Value**: URL Worker API Anda (didapatkan dari Langkah 3, contoh: `https://fm-operations-api.username.workers.dev`).
6.  Klik **Save and Deploy**. Cloudflare Pages akan mem-build kode JavaScript dan mempublikasikan situs web Anda.

---

## 🔐 LANGKAH 6 — Pengaturan CORS Origin (Opsional tapi Direkomendasikan)

Agar API Worker Anda aman dan hanya bisa diakses oleh frontend resmi Anda:

1.  Buka Worker `fm-operations-api` di dashboard.
2.  Masuk ke **Settings** → **Variables**.
3.  Tambahkan satu **Environment Variable** baru:
    *   **Name**: `CORS_ORIGIN`
    *   **Value**: URL Cloudflare Pages Anda (didapatkan dari Langkah 5, contoh: `https://fm-operations.pages.dev`).
4.  Klik **Save and Deploy**.

---

## 🚪 LANGKAH 7 — Login Pertama Kali

Setelah deployment frontend selesai, buka URL Pages Anda (misalnya: `https://fm-operations.pages.dev`):

1.  Gunakan kredensial administrator awal berikut untuk masuk:
    *   **Username**: `superadmin`
    *   **Password**: `Admin@123`
2.  ⚠️ **PENTING:** Segera setelah berhasil masuk, klik menu **Profil** di pojok kanan atas, pilih **Ganti Password**, dan ubah password bawaan tersebut demi keamanan sistem Anda.

---

*Panduan ini disesuaikan khusus untuk sistem operasional Facility Care Audy Dental.*
*Teknologi: Cloudflare Workers (API) + Pages (Frontend SPA) + D1 (SQLite Database) + KV (Sessions).*
