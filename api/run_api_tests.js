const baseUrl = 'https://fm-operations-api.facilitycare-audydental.workers.dev';

const endpoints = [
  '/api/employees',
  '/api/contracts',
  '/api/schedule',
  '/api/issues',
  '/api/one-on-one',
  '/api/training',
  '/api/relievers',
  '/api/reports/inspection',
  '/api/reports/cleaning',
  '/api/reports/fogging',
  '/api/reports/basecamp',
  '/api/sop',
  '/api/checklist',
  '/api/forms',
  '/api/pic',
  '/api/sp',
  '/api/mutasi',
  '/api/dashboard/kpi',
  '/api/dashboard/reliever-audit',
  '/api/options/branches',
  '/api/options/employees',
  '/api/options/sop',
  '/api/options/checklist'
];

async function runTests() {
  console.log('--- LOGGING IN ---');
  let token = '';
  try {
    const res = await fetch(baseUrl + '/api/auth/login', {
      method: 'POST',
      body: JSON.stringify({username: 'superadmin', password: 'admin123'}),
      headers: {'Content-Type': 'application/json'}
    });
    const data = await res.json();
    if (!data.data || !data.data.token) {
      console.log('Login failed', data);
      return;
    }
    token = data.data.token;
    console.log('Login successful! Token acquired.');
  } catch(e) {
    console.log('Login error', e.message);
    return;
  }

  console.log('--- STARTING API TESTS ---');
  let passCount = 0;
  for (const ep of endpoints) {
    try {
      const res = await fetch(baseUrl + ep, {
        headers: {
          'Authorization': 'Bearer ' + token
        }
      });
      const isJson = (res.headers.get('content-type') || '').includes('application/json');
      const bodyText = await res.text();
      let status = res.status;
      
      if (status === 200 && isJson) {
        let json;
        try {
          json = JSON.parse(bodyText);
          if (json.error) {
            console.log(`[FAIL] ${ep} - Returned 200 but JSON contains error: ${json.error}`);
          } else {
            console.log(`[PASS] ${ep} (Data length: ${bodyText.length})`);
            passCount++;
          }
        } catch(e) {
          console.log(`[FAIL] ${ep} - Invalid JSON: ${bodyText.substring(0, 50)}`);
        }
      } else {
        console.log(`[FAIL] ${ep} - Status: ${status}, Body: ${bodyText.substring(0, 100)}`);
      }
    } catch (err) {
      console.log(`[ERROR] ${ep} - Request failed: ${err.message}`);
    }
  }
  console.log(`--- RESULTS: ${passCount} / ${endpoints.length} PASSED ---`);
}

runTests();
