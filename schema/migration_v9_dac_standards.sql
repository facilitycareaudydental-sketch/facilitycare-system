-- ============================================================
-- MIGRATION V9: MASTER DAC (DAILY ACTIVITY CLEANING)
-- Standar Acuan Rutinitas Pembersihan Harian Facility Care
-- Shift Pagi: 31 Aktivitas (08:00 - 17:00)
-- Shift Siang: 33 Aktivitas (13:00 - 22:00)
-- Total: 64 Parameter Aktivitas Rutin
-- ============================================================

CREATE TABLE IF NOT EXISTS master_dac_standards (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  shift TEXT NOT NULL, -- 'pagi' atau 'siang'
  item_no INTEGER NOT NULL,
  start_time TEXT NOT NULL,
  end_time TEXT NOT NULL,
  category TEXT NOT NULL,
  location TEXT NOT NULL,
  activity TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_dac_shift ON master_dac_standards(shift);
CREATE INDEX IF NOT EXISTS idx_dac_category ON master_dac_standards(category);

-- ============================================================
-- SEED DATA: SHIFT PAGI (31 AKTIVITAS)
-- ============================================================
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 1, '08:00', '08:05', 'Persiapan', 'Semua area', 'Persiapkan peralatan kerja yang dibutuhkan, pastikan bersih dan proper, rapikan grooming');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 2, '08:05', '08:15', 'Sampah Umum', 'Ruangan & koridor Lt 1 s.d Lt 4', 'Kosongkan sampah umum pada semua dustbin di semua ruangan dan koridor');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 3, '08:15', '08:25', 'Area Umum', 'Counter receptionist & meja lemari receptionist', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 4, '08:25', '08:30', 'Area Umum', 'Lobby / teras lobby', 'Dusting, glass cleaning, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 5, '08:30', '08:40', 'Area Umum', 'Lobby (area dalam)', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 6, '08:40', '08:50', 'Ruang Tindakan', 'Ruang tindakan Du 101', 'Dusting, sweeping, glass cleaning, spotting, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 7, '08:50', '08:55', 'Ruang Penunjang', 'Ruang ronsen', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 8, '08:55', '09:05', 'Toilet', 'Toilet Lt GF', 'Dusting, sweeping, mopping; cek tools dan suplai tisu & handsoap');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 9, '09:05', '09:10', 'Area Umum', 'Snacks corner', 'Bersihkan meja dan lemari');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 10, '09:10', '09:20', 'Ruang Tindakan', 'Ruang tindakan Du 202', 'Dusting, sweeping, glass cleaning, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 11, '09:20', '09:30', 'Ruang Tindakan', 'Ruang tindakan Du 201', 'Dusting, sweeping, glass cleaning, spotting, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 12, '09:30', '09:40', 'Toilet', 'Toilet Lt 1', 'Dusting, sweeping, spotting, glass cleaning, mopping; cek tools dan suplai tisu & handsoap');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 13, '09:40', '09:45', 'Area Umum', 'Playground', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 14, '09:45', '09:55', 'Ruang Tindakan', 'Ruang tindakan Du 302', 'Dusting, sweeping, spotting, glass cleaning, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 15, '09:55', '10:05', 'Toilet', 'Toilet Lt 2', 'Dusting, sweeping, spotting, glass cleaning, mopping; cek tools dan suplai tisu & handsoap');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 16, '10:05', '10:15', 'Area Umum', 'Lobby 2', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 17, '10:15', '10:25', 'Ruang Penunjang', 'Pantry', 'Dusting, spotting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 18, '10:25', '10:30', 'Ruang Penunjang', 'Ruang kompresor', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 19, '10:30', '10:35', 'Ruang Penunjang', 'Ruang bahan perawat', 'Dusting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 20, '10:35', '10:40', 'Ruang Penunjang', 'Ruang steril 1', 'Dusting, spotting, sweeping, mopping');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 21, '10:40', '10:50', 'Area Luar', 'Rooftop', 'Sweeping dan cleaning lantai; floordrain bebas dari sampah');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 22, '10:50', '11:20', 'Kontrol & Check', 'Semua area dan ruangan', 'Mobilitas area dan pengecekan semua area/ruangan yang sudah dibersihkan');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 23, '11:20', '12:00', 'Kontrol & Check', 'Sesuai check sheet', 'Pengerjaan weekly / check sheet secara bertahap (tahap 1)');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 24, '12:00', '13:00', 'Istirahat', '-', 'Istirahat kerja');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 25, '13:00', '14:00', 'Kontrol & Check', 'Sesuai check sheet', 'Lanjutan pengerjaan weekly / check sheet (tahap 2)');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 26, '14:00', '15:00', 'Ruang Tindakan', 'Ruang tindakan Du-Du', 'Pengecekan dan pembersihan setelah praktik dokter selesai');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 27, '15:00', '16:00', 'Kontrol & Check', 'Lobby GF s.d Lantai 4', 'Mobilitas semua area');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 28, '16:00', '16:30', 'Layanan Pasien', 'Depan pintu lobby', 'Standby, memberikan sambutan dan membukakan pintu untuk pasien yang datang');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 29, '16:30', '16:40', 'Sampah Umum', 'Semua area dan ruangan', 'Kosongkan sampah umum pada semua dustbin');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 30, '16:40', '16:50', 'Sampah Medis', 'Ruang tindakan (Du) dan area penghasil sampah medis', 'Ambil sampah medis (dipisah dari sampah umum) dan buang ke tempat penampungan sampah medis');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('pagi', 31, '16:50', '17:00', 'Penutup', 'Ruang janitor', 'Bersihkan peralatan kerja dan rapikan di tempat seharusnya');

-- ============================================================
-- SEED DATA: SHIFT SIANG (33 AKTIVITAS)
-- ============================================================
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 1, '13:00', '13:10', 'Sampah Umum', 'Semua area dan dalam ruangan', 'Kosongkan sampah umum pada semua dustbin, buang ke container bin / Sulo');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 2, '13:10', '13:15', 'Area Umum', 'Teras lobby klinik', 'Bersihkan area teras lobby');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 3, '13:15', '13:25', 'Area Umum', 'Lobby klinik (area dalam)', 'Bersihkan area dalam lobby');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 4, '13:25', '13:40', 'Area Umum', 'Counter receptionist', 'Bersihkan lantai, meja, lemari, dan aksesorisnya');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 5, '13:40', '13:50', 'Ruang Tindakan', 'Ruang tindakan Du 101 (Lt GF)', 'Bersihkan ruang tindakan');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 6, '13:50', '14:00', 'Ruang Penunjang', 'Ruang ronsen', 'Bersihkan area ruang ronsen');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 7, '14:00', '14:10', 'Area Umum', 'Snacks corner', 'Bersihkan dan rapikan meja dan lemari');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 8, '14:10', '14:25', 'Toilet', 'Toilet Lt GF', 'Bersihkan toilet');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 9, '14:25', '14:35', 'Ruang Tindakan', 'Ruang tindakan Du 202 (Lt 1)', 'Bersihkan ruang tindakan');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 10, '14:35', '14:45', 'Ruang Tindakan', 'Ruang tindakan Du 201 (Lt 1)', 'Bersihkan ruang tindakan');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 11, '14:45', '15:00', 'Area Umum', 'Playground', 'Bersihkan area playground beserta aksesorisnya');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 12, '15:00', '15:10', 'Toilet', 'Toilet Lt 1', 'Bersihkan toilet; cek tools dan suplai tisu & handsoap');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 13, '15:10', '15:25', 'Ruang Tindakan', 'Ruang tindakan Du 302 (Lt 2)', 'Bersihkan ruang tindakan');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 14, '15:25', '15:35', 'Area Umum', 'Area Lt 2', 'Bersihkan area Lt 2 dan aksesorisnya');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 15, '15:35', '15:45', 'Kontrol & Check', 'Semua area', 'Mobilitas');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 16, '15:45', '16:00', 'Toilet', 'Toilet Lt 2', 'Bersihkan toilet; cek tools dan suplai tisu & handsoap');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 17, '16:00', '16:10', 'Area Umum', 'Mushola dan tempat wudhu (Lt 2)', 'Bersihkan mushola dan tempat wudhu');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 18, '16:10', '16:20', 'Ruang Penunjang', 'Pantry / ruang team (Lt 3)', 'Dusting, spotting, sweeping, mopping pantry (ruang team)');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 19, '16:20', '16:30', 'Ruang Penunjang', 'Ruang kompresor dan mesin kompresor', 'Bersihkan ruang kompresor dan mesin kompresor');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 20, '16:30', '16:40', 'Ruang Penunjang', 'Ruang bahan perawat (Lt 3)', 'Bersihkan ruang bahan perawat');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 21, '16:40', '16:45', 'Ruang Penunjang', 'Ruang steril 1 (Lt 3)', 'Bersihkan ruang steril 1 dan aksesorisnya');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 22, '16:45', '16:50', 'Area Umum', 'Lobby 3', 'Bersihkan area lobby 3 dan aksesorisnya');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 23, '16:50', '17:00', 'Toilet', 'Toilet Lt 3', 'Bersihkan toilet; cek tools dan suplai tisu & handsoap');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 24, '17:00', '17:30', 'Kontrol & Check', 'Semua area klinik', 'Mobilitas / cek ulang semua area yang sudah dibersihkan');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 25, '17:30', '18:00', 'Kontrol & Check', 'Sesuai check sheet', 'Kerjakan rencana kerja mingguan (check sheet) secara bertahap (tahap 1)');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 26, '18:00', '19:00', 'Istirahat', '-', 'Istirahat kerja / break');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 27, '19:00', '20:00', 'Kontrol & Check', 'Sesuai check sheet', 'Lanjutan rencana kerja mingguan (check sheet) secara bertahap (tahap 2)');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 28, '20:00', '20:15', 'Area Luar', 'Rooftop', 'Pengecekan floordrain dan pembersihan area rooftop');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 29, '20:15', '21:00', 'Layanan Pasien', 'Area lobby', 'Standby, membukakan pintu dan menyambut pasien');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 30, '21:00', '21:10', 'Sampah Umum', 'Semua area dan dalam ruangan', 'Kosongkan sampah umum pada semua dustbin');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 31, '21:10', '21:45', 'Ruang Tindakan', 'Ruang tindakan Du 1–Du 4, koridor klinik, toilet', 'Bersihkan semua ruang tindakan, area koridor, dan toilet');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 32, '21:45', '21:55', 'Sampah Medis', 'Ruang tindakan (Du) dan area penghasil sampah medis', 'Ambil sampah medis (dipisah dari sampah umum) dan buang ke tempat penampungan sampah medis');
INSERT INTO master_dac_standards (shift, item_no, start_time, end_time, category, location, activity) VALUES ('siang', 33, '21:55', '22:00', 'Penutup', 'Ruang janitor', 'Bersihkan dan rapikan peralatan kerja, simpan di tempat seharusnya');
