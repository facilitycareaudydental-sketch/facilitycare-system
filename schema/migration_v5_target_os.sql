-- Migration v5: Add target_pindah_os and target_selesai_os columns to employees
ALTER TABLE employees ADD COLUMN target_pindah_os TEXT;
ALTER TABLE employees ADD COLUMN target_selesai_os TEXT;
