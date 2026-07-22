import { handleEmployees } from './src/routes/employees.js';

const mockEnv = {
  DB: {
    prepare: (query) => {
      console.log('PREPARE:', query);
      return {
        bind: (...args) => {
          console.log('BIND:', args);
          return {
            all: async () => ({ results: [] }),
            first: async () => null,
            run: async () => {
              console.log('RUN');
              return { meta: { last_row_id: 999 } };
            }
          }
        }
      }
    }
  }
};

const req = new Request('http://localhost/api/employees', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Authorization': 'Bearer test'
  },
  body: JSON.stringify({
    full_name: "Test Karyawan",
    branch_id: 15,
    division: "FACILITY CARE",
    phone: "",
    join_date: "",
    status: "Aktif",
    notes: ""
  })
});

async function run() {
  try {
    const res = await handleEmployees(req, mockEnv, 'http://localhost');
    console.log('STATUS:', res.status);
    console.log('BODY:', await res.text());
  } catch (e) {
    console.error('CRASH:', e);
  }
}
run();
