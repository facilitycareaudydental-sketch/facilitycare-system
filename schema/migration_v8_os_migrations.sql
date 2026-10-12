-- ============================================================
-- MIGRATION V8: TIMELINE MIGRASI PERPINDAHAN OUTSOURCING (OS)
-- Total: 62 Cabang (AEON: 20 Cabang, Triple S: 42 Cabang)
-- Periode: September 2026 - Desember 2026
-- ============================================================

CREATE TABLE IF NOT EXISTS os_migrations (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER NOT NULL,
  vendor TEXT NOT NULL,
  target_month TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'scheduled',
  actual_date TEXT,
  pic_vendor TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

CREATE INDEX IF NOT EXISTS idx_os_migrations_branch ON os_migrations(branch_id);
CREATE INDEX IF NOT EXISTS idx_os_migrations_vendor_month ON os_migrations(vendor, target_month);

-- ============================================================
-- SEED DATA: 62 CABANG TIMELINE MIGRASI OS
-- ============================================================
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (4, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (003. Depok)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (8, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (007. Bekasi)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (13, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (012. BSD)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (15, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (014. Sunter)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (20, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (019. Cinere)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (22, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (021. Kelapa Gading Boulevard)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (12, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (011. Kuningan)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (47, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (046. Kaliurang Jogja)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (60, 'AEON', '2026-09', 'completed', 'Migrasi OS ke AEON (059. Malang Suhat)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (3, 'AEON', '2026-11', 'scheduled', 'Migrasi OS ke AEON (002. Kemang)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (5, 'AEON', '2026-11', 'scheduled', 'Migrasi OS ke AEON (004. Cipete)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (17, 'AEON', '2026-11', 'scheduled', 'Migrasi OS ke AEON (016. Cibubur)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (6, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (005. Greenville)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (10, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (009. Bintaro)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (11, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (010. Bogor)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (14, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (013. Bandung Gatsu)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (16, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (015. Gading Serpong)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (25, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (024. Grand Wisata)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (28, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (027. Rawamangun)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (32, 'AEON', '2026-12', 'scheduled', 'Migrasi OS ke AEON (031. Sawangan)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (30, 'Triple S', '2026-09', 'completed', 'Migrasi OS ke Triple S (029. Karawang)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (9, 'Triple S', '2026-09', 'completed', 'Migrasi OS ke Triple S (008. PIK)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (51, 'Triple S', '2026-09', 'completed', 'Migrasi OS ke Triple S (050. Cibinong)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (53, 'Triple S', '2026-09', 'completed', 'Migrasi OS ke Triple S (052. BSD Rawa Buntu)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (55, 'Triple S', '2026-09', 'completed', 'Migrasi OS ke Triple S (054. Bogor Gunung Batu)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (27, 'Triple S', '2026-10', 'in_progress', 'Migrasi OS ke Triple S (026. Tangerang City)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (29, 'Triple S', '2026-10', 'in_progress', 'Migrasi OS ke Triple S (028. Taman Palem)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (43, 'Triple S', '2026-10', 'in_progress', 'Migrasi OS ke Triple S (042. Muara Karang)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (57, 'Triple S', '2026-10', 'in_progress', 'Migrasi OS ke Triple S (056. Citra Garden 2)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (63, 'Triple S', '2026-10', 'in_progress', 'Migrasi OS ke Triple S (062. Summarecon Bekasi)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (35, 'Triple S', '2026-11', 'scheduled', 'Migrasi OS ke Triple S (034. Puri Indah)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (39, 'Triple S', '2026-11', 'scheduled', 'Migrasi OS ke Triple S (038. Greenlake)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (40, 'Triple S', '2026-11', 'scheduled', 'Migrasi OS ke Triple S (039. Bandung Mekarwangi)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (41, 'Triple S', '2026-11', 'scheduled', 'Migrasi OS ke Triple S (040. Bekasi Timur)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (44, 'Triple S', '2026-11', 'scheduled', 'Migrasi OS ke Triple S (043. Tebet)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (7, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (006. Kelapa Gading MOI)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (31, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (030. Ciputat)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (23, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (022. Cikarang)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (18, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (017. Bandung Setrasari)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (19, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (018. Harapan Indah)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (21, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (020. Surabaya)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (24, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (023. Semarang)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (26, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (025. Bogor Baru)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (33, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (032. Cawang)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (34, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (033. Taman Mini)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (36, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (035. Ciledug)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (37, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (036. Jagakarsa)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (42, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (041. Karawaci)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (45, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (044. Depok GDC)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (46, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (045. FX Mall)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (48, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (047. Bogor Tajur)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (49, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (048. PIK 2)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (50, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (049. Surabaya Merr)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (52, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (051. Bintaro Sektor 9)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (54, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (053. Alam Sutera)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (56, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (055. Gajah Mada)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (58, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (057. Cikupa)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (61, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (060. Palmerah)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (64, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (063. Pasar Minggu)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (65, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (064. Pondok Indah)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (66, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (065. Kebon Sirih)');
INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes) VALUES (68, 'Triple S', '2026-12', 'scheduled', 'Migrasi OS ke Triple S (067. Agora Mall)');
