-- Delete all relievers for July 2026 to clear duplicates/ghost data
DELETE FROM relievers 
WHERE 
  (CAST(backup_date AS INTEGER) BETWEEN 46204 AND 46234)
  OR strftime('%Y-%m', backup_date) = '2026-07'
  OR strftime('%Y-%m', REPLACE(backup_date, '/', '-')) = '2026-07'
  OR backup_date LIKE '%2026-07%'
  OR backup_date LIKE '%/07/2026%';
