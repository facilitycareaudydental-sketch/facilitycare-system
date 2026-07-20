/**
 * Script untuk mengimport data SP dan Mutasi dari Google Sheet ke FCMS
 * Jalankan: node import_sp_mutasi.js
 *
 * Ganti TOKEN dengan token dari login FCMS Bapak
 */

const BASE_URL = 'https://facilitycare-system.pages.dev';

// ─── Data SP dari Google Sheet (Bapak isi sesuai data yang ada) ──────────────
// Kolom: tanggal, employee_name, branch_id, sp_type, status, document_link
// Contoh data – jika data Google Sheet kosong, ini sebagai placeholder
const SP_DATA = [
  // { tanggal: '2026-01-15', employee_name: 'Nama Karyawan', branch_name: 'Nama Cabang', sp_type: 'SP 1', status: 'Aktif', document_link: '' },
];

// ─── Data Mutasi dari Google Sheet ───────────────────────────────────────────
const MUTASI_DATA = [
  // { tanggal: '2026-01-15', employee_name: 'Nama Karyawan', from_branch_name: 'Cabang Asal', to_branch_name: 'Cabang Tujuan', status: 'Selesai', document_link: '' },
];

async function main() {
  // 1. Login
  console.log('🔑 Logging in...');
  const loginRes = await fetch(`${BASE_URL}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username: 'superadmin', password: 'Admin@123' })
  });
  const loginData = await loginRes.json();
  if (!loginData.data?.token) {
    console.error('❌ Login gagal:', loginData.error);
    process.exit(1);
  }
  const TOKEN = loginData.data.token;
  console.log('✅ Login berhasil sebagai', loginData.data.user.username);

  // 2. Ambil daftar branch untuk resolve nama → id
  const branchRes = await fetch(`${BASE_URL}/api/branches?all=1`, {
    headers: { Authorization: `Bearer ${TOKEN}` }
  });
  const branchData = await branchRes.json();
  const branches = branchData.data || [];
  console.log(`📍 ${branches.length} cabang ditemukan`);

  const findBranch = (name) => {
    if (!name) return null;
    const n = name.toLowerCase().trim();
    const b = branches.find(b =>
      b.full_name?.toLowerCase().includes(n) ||
      b.name?.toLowerCase().includes(n) ||
      b.code?.toLowerCase() === n
    );
    return b ? b.id : null;
  };

  // 3. Test apakah tabel sp_data sudah ada
  const testRes = await fetch(`${BASE_URL}/api/sp?limit=1`, {
    headers: { Authorization: `Bearer ${TOKEN}` }
  });
  const testData = await testRes.json();
  if (!testRes.ok) {
    console.error('❌ Tabel sp_data belum ada di database:', testData.error);
    console.log('⏳ Tunggu deploy GitHub Actions selesai terlebih dahulu (sekitar 5 menit), lalu jalankan script ini kembali.');
    process.exit(1);
  }
  console.log('✅ Tabel sp_data sudah ada');

  // 4. Import SP data
  let spOk = 0, spErr = 0;
  for (const item of SP_DATA) {
    const payload = {
      tanggal: item.tanggal,
      employee_name: item.employee_name,
      branch_id: findBranch(item.branch_name),
      sp_type: item.sp_type,
      status: item.status || 'Aktif',
      document_link: item.document_link || null
    };
    const r = await fetch(`${BASE_URL}/api/sp`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${TOKEN}` },
      body: JSON.stringify(payload)
    });
    if (r.ok) { spOk++; } else { spErr++; const d = await r.json(); console.error('SP error:', d.error, item); }
  }
  console.log(`✅ SP import: ${spOk} berhasil, ${spErr} gagal`);

  // 5. Import Mutasi data
  let mutasiOk = 0, mutasiErr = 0;
  for (const item of MUTASI_DATA) {
    const payload = {
      tanggal: item.tanggal,
      employee_name: item.employee_name,
      from_branch_id: findBranch(item.from_branch_name),
      to_branch_id: findBranch(item.to_branch_name),
      status: item.status || 'Proses',
      document_link: item.document_link || null
    };
    const r = await fetch(`${BASE_URL}/api/mutasi`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${TOKEN}` },
      body: JSON.stringify(payload)
    });
    if (r.ok) { mutasiOk++; } else { mutasiErr++; const d = await r.json(); console.error('Mutasi error:', d.error, item); }
  }
  console.log(`✅ Mutasi import: ${mutasiOk} berhasil, ${mutasiErr} gagal`);

  if (SP_DATA.length === 0 && MUTASI_DATA.length === 0) {
    console.log('\n⚠️  Data SP dan Mutasi di Google Sheet masih kosong.');
    console.log('📋 Silakan Bapak isi data di Google Sheet terlebih dahulu,');
    console.log('   atau gunakan tombol "Tambah SP/Mutasi" di aplikasi FCMS secara langsung.');
  }

  console.log('\n✅ Selesai!');
}

main().catch(e => { console.error('Fatal error:', e.message); process.exit(1); });
