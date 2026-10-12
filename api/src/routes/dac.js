// Daily Activity Cleaning (DAC) API - Shift Pagi & Shift Siang
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

const SEED_PAGI = [
  { item_no: 1, start_time: '08:00', end_time: '08:05', category: 'Persiapan', location: 'Semua area', activity: 'Persiapkan peralatan kerja yang dibutuhkan, pastikan bersih dan proper, rapikan grooming' },
  { item_no: 2, start_time: '08:05', end_time: '08:15', category: 'Sampah Umum', location: 'Ruangan & koridor Lt 1 s.d Lt 4', activity: 'Kosongkan sampah umum pada semua dustbin di semua ruangan dan koridor' },
  { item_no: 3, start_time: '08:15', end_time: '08:25', category: 'Area Umum', location: 'Counter receptionist & meja lemari receptionist', activity: 'Dusting, sweeping, mopping' },
  { item_no: 4, start_time: '08:25', end_time: '08:30', category: 'Area Umum', location: 'Lobby / teras lobby', activity: 'Dusting, glass cleaning, sweeping, mopping' },
  { item_no: 5, start_time: '08:30', end_time: '08:40', category: 'Area Umum', location: 'Lobby (area dalam)', activity: 'Dusting, sweeping, mopping' },
  { item_no: 6, start_time: '08:40', end_time: '08:50', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 101', activity: 'Dusting, sweeping, glass cleaning, spotting, mopping' },
  { item_no: 7, start_time: '08:50', end_time: '08:55', category: 'Ruang Penunjang', location: 'Ruang ronsen', activity: 'Dusting, sweeping, mopping' },
  { item_no: 8, start_time: '08:55', end_time: '09:05', category: 'Toilet', location: 'Toilet Lt GF', activity: 'Dusting, sweeping, mopping; cek tools dan suplai tisu & handsoap' },
  { item_no: 9, start_time: '09:05', end_time: '09:10', category: 'Area Umum', location: 'Snacks corner', activity: 'Bersihkan meja dan lemari' },
  { item_no: 10, start_time: '09:10', end_time: '09:20', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 202', activity: 'Dusting, sweeping, glass cleaning, mopping' },
  { item_no: 11, start_time: '09:20', end_time: '09:30', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 201', activity: 'Dusting, sweeping, glass cleaning, spotting, mopping' },
  { item_no: 12, start_time: '09:30', end_time: '09:40', category: 'Toilet', location: 'Toilet Lt 1', activity: 'Dusting, sweeping, spotting, glass cleaning, mopping; cek tools dan suplai tisu & handsoap' },
  { item_no: 13, start_time: '09:40', end_time: '09:45', category: 'Area Umum', location: 'Playground', activity: 'Dusting, sweeping, mopping' },
  { item_no: 14, start_time: '09:45', end_time: '09:55', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 302', activity: 'Dusting, sweeping, spotting, glass cleaning, mopping' },
  { item_no: 15, start_time: '09:55', end_time: '10:05', category: 'Toilet', location: 'Toilet Lt 2', activity: 'Dusting, sweeping, spotting, glass cleaning, mopping; cek tools dan suplai tisu & handsoap' },
  { item_no: 16, start_time: '10:05', end_time: '10:15', category: 'Area Umum', location: 'Lobby 2', activity: 'Dusting, sweeping, mopping' },
  { item_no: 17, start_time: '10:15', end_time: '10:25', category: 'Ruang Penunjang', location: 'Pantry', activity: 'Dusting, spotting, sweeping, mopping' },
  { item_no: 18, start_time: '10:25', end_time: '10:30', category: 'Ruang Penunjang', location: 'Ruang kompresor', activity: 'Dusting, sweeping, mopping' },
  { item_no: 19, start_time: '10:30', end_time: '10:35', category: 'Ruang Penunjang', location: 'Ruang bahan perawat', activity: 'Dusting, sweeping, mopping' },
  { item_no: 20, start_time: '10:35', end_time: '10:40', category: 'Ruang Penunjang', location: 'Ruang steril 1', activity: 'Dusting, spotting, sweeping, mopping' },
  { item_no: 21, start_time: '10:40', end_time: '10:50', category: 'Area Luar', location: 'Rooftop', activity: 'Sweeping dan cleaning lantai; floordrain bebas dari sampah' },
  { item_no: 22, start_time: '10:50', end_time: '11:20', category: 'Kontrol & Check', location: 'Semua area dan ruangan', activity: 'Mobilitas area dan pengecekan semua area/ruangan yang sudah dibersihkan' },
  { item_no: 23, start_time: '11:20', end_time: '12:00', category: 'Kontrol & Check', location: 'Sesuai check sheet', activity: 'Pengerjaan weekly / check sheet secara bertahap (tahap 1)' },
  { item_no: 24, start_time: '12:00', end_time: '13:00', category: 'Istirahat', location: '-', activity: 'Istirahat kerja' },
  { item_no: 25, start_time: '13:00', end_time: '14:00', category: 'Kontrol & Check', location: 'Sesuai check sheet', activity: 'Lanjutan pengerjaan weekly / check sheet (tahap 2)' },
  { item_no: 26, start_time: '14:00', end_time: '15:00', category: 'Ruang Tindakan', location: 'Ruang tindakan Du-Du', activity: 'Pengecekan dan pembersihan setelah praktik dokter selesai' },
  { item_no: 27, start_time: '15:00', end_time: '16:00', category: 'Kontrol & Check', location: 'Lobby GF s.d Lantai 4', activity: 'Mobilitas semua area' },
  { item_no: 28, start_time: '16:00', end_time: '16:30', category: 'Layanan Pasien', location: 'Depan pintu lobby', activity: 'Standby, memberikan sambutan dan membukakan pintu untuk pasien yang datang' },
  { item_no: 29, start_time: '16:30', end_time: '16:40', category: 'Sampah Umum', location: 'Semua area dan ruangan', activity: 'Kosongkan sampah umum pada semua dustbin' },
  { item_no: 30, start_time: '16:40', end_time: '16:50', category: 'Sampah Medis', location: 'Ruang tindakan (Du) dan area penghasil sampah medis', activity: 'Ambil sampah medis (dipisah dari sampah umum) dan buang ke tempat penampungan sampah medis' },
  { item_no: 31, start_time: '16:50', end_time: '17:00', category: 'Penutup', location: 'Ruang janitor', activity: 'Bersihkan peralatan kerja dan rapikan di tempat seharusnya' }
];

const SEED_SIANG = [
  { item_no: 1, start_time: '13:00', end_time: '13:10', category: 'Sampah Umum', location: 'Semua area dan dalam ruangan', activity: 'Kosongkan sampah umum pada semua dustbin, buang ke container bin / Sulo' },
  { item_no: 2, start_time: '13:10', end_time: '13:15', category: 'Area Umum', location: 'Teras lobby klinik', activity: 'Bersihkan area teras lobby' },
  { item_no: 3, start_time: '13:15', end_time: '13:25', category: 'Area Umum', location: 'Lobby klinik (area dalam)', activity: 'Bersihkan area dalam lobby' },
  { item_no: 4, start_time: '13:25', end_time: '13:40', category: 'Area Umum', location: 'Counter receptionist', activity: 'Bersihkan lantai, meja, lemari, dan aksesorisnya' },
  { item_no: 5, start_time: '13:40', end_time: '13:50', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 101 (Lt GF)', activity: 'Bersihkan ruang tindakan' },
  { item_no: 6, start_time: '13:50', end_time: '14:00', category: 'Ruang Penunjang', location: 'Ruang ronsen', activity: 'Bersihkan area ruang ronsen' },
  { item_no: 7, start_time: '14:00', end_time: '14:10', category: 'Area Umum', location: 'Snacks corner', activity: 'Bersihkan dan rapikan meja dan lemari' },
  { item_no: 8, start_time: '14:10', end_time: '14:25', category: 'Toilet', location: 'Toilet Lt GF', activity: 'Bersihkan toilet' },
  { item_no: 9, start_time: '14:25', end_time: '14:35', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 202 (Lt 1)', activity: 'Bersihkan ruang tindakan' },
  { item_no: 10, start_time: '14:35', end_time: '14:45', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 201 (Lt 1)', activity: 'Bersihkan ruang tindakan' },
  { item_no: 11, start_time: '14:45', end_time: '15:00', category: 'Area Umum', location: 'Playground', activity: 'Bersihkan area playground beserta aksesorisnya' },
  { item_no: 12, start_time: '15:00', end_time: '15:10', category: 'Toilet', location: 'Toilet Lt 1', activity: 'Bersihkan toilet; cek tools dan suplai tisu & handsoap' },
  { item_no: 13, start_time: '15:10', end_time: '15:25', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 302 (Lt 2)', activity: 'Bersihkan ruang tindakan' },
  { item_no: 14, start_time: '15:25', end_time: '15:35', category: 'Area Umum', location: 'Area Lt 2', activity: 'Bersihkan area Lt 2 dan aksesorisnya' },
  { item_no: 15, start_time: '15:35', end_time: '15:45', category: 'Kontrol & Check', location: 'Semua area', activity: 'Mobilitas' },
  { item_no: 16, start_time: '15:45', end_time: '16:00', category: 'Toilet', location: 'Toilet Lt 2', activity: 'Bersihkan toilet; cek tools dan suplai tisu & handsoap' },
  { item_no: 17, start_time: '16:00', end_time: '16:10', category: 'Area Umum', location: 'Mushola dan tempat wudhu (Lt 2)', activity: 'Bersihkan mushola dan tempat wudhu' },
  { item_no: 18, start_time: '16:10', end_time: '16:20', category: 'Ruang Penunjang', location: 'Pantry / ruang team (Lt 3)', activity: 'Dusting, spotting, sweeping, mopping pantry (ruang team)' },
  { item_no: 19, start_time: '16:20', end_time: '16:30', category: 'Ruang Penunjang', location: 'Ruang kompresor dan mesin kompresor', activity: 'Bersihkan ruang kompresor dan mesin kompresor' },
  { item_no: 20, start_time: '16:30', end_time: '16:40', category: 'Ruang Penunjang', location: 'Ruang bahan perawat (Lt 3)', activity: 'Bersihkan ruang bahan perawat' },
  { item_no: 21, start_time: '16:40', end_time: '16:45', category: 'Ruang Penunjang', location: 'Ruang steril 1 (Lt 3)', activity: 'Bersihkan ruang steril 1 dan aksesorisnya' },
  { item_no: 22, start_time: '16:45', end_time: '16:50', category: 'Area Umum', location: 'Lobby 3', activity: 'Bersihkan area lobby 3 dan aksesorisnya' },
  { item_no: 23, start_time: '16:50', end_time: '17:00', category: 'Toilet', location: 'Toilet Lt 3', activity: 'Bersihkan toilet; cek tools dan suplai tisu & handsoap' },
  { item_no: 24, start_time: '17:00', end_time: '17:30', category: 'Kontrol & Check', location: 'Semua area klinik', activity: 'Mobilitas / cek ulang semua area yang sudah dibersihkan' },
  { item_no: 25, start_time: '17:30', end_time: '18:00', category: 'Kontrol & Check', location: 'Sesuai check sheet', activity: 'Kerjakan rencana kerja mingguan (check sheet) secara bertahap (tahap 1)' },
  { item_no: 26, start_time: '18:00', end_time: '19:00', category: 'Istirahat', location: '-', activity: 'Istirahat kerja / break' },
  { item_no: 27, start_time: '19:00', end_time: '20:00', category: 'Kontrol & Check', location: 'Sesuai check sheet', activity: 'Lanjutan rencana kerja mingguan (check sheet) secara bertahap (tahap 2)' },
  { item_no: 28, start_time: '20:00', end_time: '20:15', category: 'Area Luar', location: 'Rooftop', activity: 'Pengecekan floordrain dan pembersihan area rooftop' },
  { item_no: 29, start_time: '20:15', end_time: '21:00', category: 'Layanan Pasien', location: 'Area lobby', activity: 'Standby, membukakan pintu dan menyambut pasien' },
  { item_no: 30, start_time: '21:00', end_time: '21:10', category: 'Sampah Umum', location: 'Semua area dan dalam ruangan', activity: 'Kosongkan sampah umum pada semua dustbin' },
  { item_no: 31, start_time: '21:10', end_time: '21:45', category: 'Ruang Tindakan', location: 'Ruang tindakan Du 1–Du 4, koridor klinik, toilet', activity: 'Bersihkan semua ruang tindakan, area koridor, dan toilet' },
  { item_no: 32, start_time: '21:45', end_time: '21:55', category: 'Sampah Medis', location: 'Ruang tindakan (Du) dan area penghasil sampah medis', activity: 'Ambil sampah medis (dipisah dari sampah umum) dan buang ke tempat penampungan sampah medis' },
  { item_no: 33, start_time: '21:55', end_time: '22:00', category: 'Penutup', location: 'Ruang janitor', activity: 'Bersihkan dan rapikan peralatan kerja, simpan di tempat seharusnya' }
];

let _initialized = false;
async function ensureInitialized(env) {
  if (_initialized) return;
  try {
    await env.DB.prepare(`
      CREATE TABLE IF NOT EXISTS master_dac_standards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        shift TEXT NOT NULL,
        item_no INTEGER NOT NULL,
        start_time TEXT NOT NULL,
        end_time TEXT NOT NULL,
        category TEXT NOT NULL,
        location TEXT NOT NULL,
        activity TEXT NOT NULL,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
      )
    `).run();

    await env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_dac_shift ON master_dac_standards(shift)`).run();
    await env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_dac_category ON master_dac_standards(category)`).run();

    const countRow = await env.DB.prepare('SELECT COUNT(*) as cnt FROM master_dac_standards').first();
    if (!countRow || countRow.cnt === 0) {
      for (const item of SEED_PAGI) {
        await env.DB.prepare(`
          INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity)
          VALUES ('pagi', ?, ?, ?, ?, ?, ?)
        `).bind(item.item_no, item.start_time, item.end_time, item.category, item.location, item.activity).run();
      }
      for (const item of SEED_SIANG) {
        await env.DB.prepare(`
          INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity)
          VALUES ('siang', ?, ?, ?, ?, ?, ?)
        `).bind(item.item_no, item.start_time, item.end_time, item.category, item.location, item.activity).run();
      }
    }
    _initialized = true;
  } catch (err) {
    console.warn('master_dac_standards initialization warning:', err.message);
  }
}

export async function handleDAC(request, env, origin) {
  await ensureInitialized(env);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/dac', '');
  const idMatch = path.match(/^\/(\d+)$/);

  // GET summary
  if (request.method === 'GET' && path === '/summary') {
    return getDACSummary(request, env, origin);
  }

  // GET categories list
  if (request.method === 'GET' && path === '/categories') {
    return getDACCategories(request, env, origin);
  }

  // GET list (Public read for all authenticated users)
  if (request.method === 'GET' && path === '') {
    return listDAC(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      return getDACItem(id, env, origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'dac', 'write')) return forbidden(origin);
      return updateDACItem(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'dac', 'delete')) return forbidden(origin);
      return deleteDACItem(id, env, origin);
    }
  }

  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'dac', 'write')) return forbidden(origin);
    return createDACItem(request, env, origin);
  }

  return error('Not found', 404, origin);
}

async function listDAC(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getPagination(request.url);
  const shift = url.searchParams.get('shift') || '';
  const category = url.searchParams.get('category') || '';
  const search = url.searchParams.get('search') || '';
  const all = url.searchParams.get('all');

  let conditions = [];
  let values = [];

  if (shift) {
    conditions.push('shift = ?');
    values.push(shift);
  }
  if (category) {
    conditions.push('category = ?');
    values.push(category);
  }
  if (search) {
    conditions.push('(location LIKE ? OR activity LIKE ? OR category LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }

  const whereClause = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const baseSelect = `SELECT * FROM master_dac_standards ${whereClause} ORDER BY shift ASC, item_no ASC`;

  if (all === '1') {
    const rows = await env.DB.prepare(baseSelect).bind(...values).all();
    return ok(rows.results || [], 200, origin);
  }

  const countQuery = `SELECT COUNT(*) as total FROM master_dac_standards ${whereClause}`;

  const [rows, count] = await Promise.all([
    env.DB.prepare(`${baseSelect} LIMIT ? OFFSET ?`).bind(...values, limit, offset).all(),
    env.DB.prepare(countQuery).bind(...values).first()
  ]);

  return paginated(rows.results || [], count?.total || 0, page, limit, origin);
}

async function getDACSummary(request, env, origin) {
  const rows = (await env.DB.prepare('SELECT * FROM master_dac_standards ORDER BY shift ASC, item_no ASC').all()).results || [];
  
  const pagi = rows.filter(r => r.shift === 'pagi');
  const siang = rows.filter(r => r.shift === 'siang');

  const categories = Array.from(new Set(rows.map(r => r.category)));

  return ok({
    total_all: rows.length,
    total_pagi: pagi.length,
    total_siang: siang.length,
    categories,
    pagi,
    siang
  }, 200, origin);
}

async function getDACCategories(request, env, origin) {
  const rows = (await env.DB.prepare('SELECT DISTINCT category FROM master_dac_standards ORDER BY category ASC').all()).results || [];
  return ok(rows.map(r => r.category), 200, origin);
}

async function getDACItem(id, env, origin) {
  const row = await env.DB.prepare('SELECT * FROM master_dac_standards WHERE id = ?').bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createDACItem(request, env, origin) {
  try {
    const body = await request.json();
    const { shift, item_no, start_time, end_time, category, location, activity } = body;

    if (!shift || !start_time || !end_time || !category || !location || !activity) {
      return error('Shift, Jam Mulai, Jam Selesai, Kategori, Lokasi, dan Aktivitas wajib diisi', 400, origin);
    }

    // Auto calculate item_no if not provided
    let num = item_no;
    if (!num) {
      const last = await env.DB.prepare('SELECT MAX(item_no) as max_no FROM master_dac_standards WHERE shift = ?').bind(shift).first();
      num = (last?.max_no || 0) + 1;
    }

    const res = await env.DB.prepare(`
      INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).bind(shift, num, start_time, end_time, category, location, activity).run();

    const created = await env.DB.prepare('SELECT * FROM master_dac_standards WHERE id = ?').bind(res.meta.last_row_id).first();
    return ok(created, 201, origin);
  } catch (err) {
    return error(`Gagal menambahkan item DAC: ${err.message}`, 500, origin);
  }
}

async function updateDACItem(id, request, env, origin) {
  try {
    const body = await request.json();
    const { shift, item_no, start_time, end_time, category, location, activity } = body;

    const existing = await env.DB.prepare('SELECT id FROM master_dac_standards WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    if (!shift || !start_time || !end_time || !category || !location || !activity) {
      return error('Shift, Jam Mulai, Jam Selesai, Kategori, Lokasi, dan Aktivitas wajib diisi', 400, origin);
    }

    await env.DB.prepare(`
      UPDATE master_dac_standards
      SET shift = ?, item_no = ?, start_time = ?, end_time = ?, category = ?, location = ?, activity = ?, updated_at = datetime('now')
      WHERE id = ?
    `).bind(shift, item_no || 1, start_time, end_time, category, location, activity, id).run();

    const updated = await env.DB.prepare('SELECT * FROM master_dac_standards WHERE id = ?').bind(id).first();
    return ok(updated, 200, origin);
  } catch (err) {
    return error(`Gagal memperbarui item DAC: ${err.message}`, 500, origin);
  }
}

async function deleteDACItem(id, env, origin) {
  try {
    const existing = await env.DB.prepare('SELECT id FROM master_dac_standards WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    await env.DB.prepare('DELETE FROM master_dac_standards WHERE id = ?').bind(id).run();
    return ok({ message: 'Item DAC berhasil dihapus' }, 200, origin);
  } catch (err) {
    return error(`Gagal menghapus item DAC: ${err.message}`, 500, origin);
  }
}
