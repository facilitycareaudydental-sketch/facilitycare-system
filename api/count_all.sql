SELECT 'employees' as tbl, COUNT(*) as cnt FROM employees
UNION ALL SELECT 'contracts', COUNT(*) FROM contracts
UNION ALL SELECT 'issues', COUNT(*) FROM issues
UNION ALL SELECT 'one_on_one', COUNT(*) FROM one_on_one
UNION ALL SELECT 'training', COUNT(*) FROM training
UNION ALL SELECT 'relievers', COUNT(*) FROM relievers
UNION ALL SELECT 'activity_schedule', COUNT(*) FROM activity_schedule
UNION ALL SELECT 'inspection_reports', COUNT(*) FROM inspection_reports
UNION ALL SELECT 'cleaning_reports', COUNT(*) FROM cleaning_reports
UNION ALL SELECT 'fogging_reports', COUNT(*) FROM fogging_reports
UNION ALL SELECT 'basecamp_reports', COUNT(*) FROM basecamp_reports
UNION ALL SELECT 'sop', COUNT(*) FROM sop
UNION ALL SELECT 'master_checklist', COUNT(*) FROM master_checklist
UNION ALL SELECT 'master_forms', COUNT(*) FROM master_forms
UNION ALL SELECT 'supply_requests', COUNT(*) FROM supply_requests
