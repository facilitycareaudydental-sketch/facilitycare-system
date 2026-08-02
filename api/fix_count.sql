DELETE FROM contracts WHERE employee_id = 1019 AND id = (SELECT max(id) FROM contracts WHERE employee_id = 1019);
DELETE FROM contracts WHERE employee_id = 1071 AND id = (SELECT max(id) FROM contracts WHERE employee_id = 1071);
INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, status) VALUES ('1082', (SELECT id FROM branches WHERE full_name LIKE '%Cikupa%' LIMIT 1), 'FACILITY CARE', '', '2099-12-31', 'Aktif');
