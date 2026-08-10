-- ============================================================
-- SQL Update: Perbaikan Tanggal Training (Mengambil Tanggal Awal Tiap Batch)
-- Format: YYYY-MM-DD (Tampil di aplikasi: 9 Juni 2026, 12 Juni 2026, dst)
-- ============================================================

-- 1. Batch 1 (9 - 11 Juni 2026) -> 9 Juni 2026
UPDATE training 
SET training_date = '2026-06-09', updated_at = datetime('now')
WHERE batch = 'Batch 1';

-- 2. Batch 2 (12, 14, 15 Juni 2026) -> 12 Juni 2026
UPDATE training 
SET training_date = '2026-06-12', updated_at = datetime('now')
WHERE batch = 'Batch 2' OR participants LIKE '%DEDY SUPANDI%';

-- 3. Batch 3 (17 - 19 Juni 2026) -> 17 Juni 2026
UPDATE training 
SET training_date = '2026-06-17', updated_at = datetime('now')
WHERE batch = 'Batch 3' AND participants NOT LIKE '%DEDY SUPANDI%';

-- 4. Batch 4 (24 - 26 Juni 2026) -> 24 Juni 2026
UPDATE training 
SET training_date = '2026-06-24', updated_at = datetime('now')
WHERE batch = 'Batch 4';

-- 5. Batch 5 (29, 30 Juni & 1 Juli 2026) -> 29 Juni 2026
UPDATE training 
SET training_date = '2026-06-29', updated_at = datetime('now')
WHERE batch = 'Batch 5';
