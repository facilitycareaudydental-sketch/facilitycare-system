import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { downloadExcel } from '../utils/excel.js';

// 35 Ruangan acuan standar FCMS (Source of Truth)
const DEFAULT_ROOMS = [
  'Lantai & Sudut Ruangan',
  'Dinding, Partisi & Plint',
  'Kaca, Cermin & Partisi Kaca',
  'Sanitair & Area Basah Toilet',
  'Plafon, Lampu & Kisi Ventilasi',
  'Mebel, Meja & Kursi Kerja/Pasien',
  'Peralatan Khusus & Medis Non-Steril',
  'Wadah Sampah & Utilitas',
  'Teras',
  'Lobby 1',
  'Receptionist',
  'Snack Corner',
  'DU 101',
  'Ronsen',
  'Toilet 1',
  'Janitor',
  'Ruang Limbah',
  'Ruang Steril',
  'Area Tangga',
  'Lobby 2',
  'Playground',
  'Musholla',
  'Tempat Wudhu',
  'Ruang Dokter',
  'Ruang Team',
  'Gudang Farmasi & Bahan',
  'RoofTop',
  'Ruang Kompressor',
  'Lift',
  'Area Genset',
  'Area Tangga Exit',
  'Taman',
  'Pos Security',
  'Toilet Security',
  'Halaman Parkir'
];

export async function renderHygieneStandards(container) {
  let roomOptions = DEFAULT_ROOMS;
  try {
    const res = await apiFetch('/api/hygiene-standards/rooms');
    if (res.ok && Array.isArray(res.data) && res.data.length > 0) {
      roomOptions = res.data;
    }
  } catch {
    // Fallback to DEFAULT_ROOMS
  }

  buildCrudPage({
    container,
    title: 'Master Hygiene',
    icon: '✨',
    apiPath: '/api/hygiene-standards',
    bulkDelete: false, // Operasi destructive massal dimatikan sesuai kebijakan sistem
    enableMobileFilterSheet: true,
    itemLabel: 'Parameter Kebersihan',
    columns: [
      {
        key: 'room_name',
        label: 'Ruangan',
        render: (v) => `<span class="badge badge-info" style="font-weight:600; white-space:nowrap;">${v || '-'}</span>`
      },
      {
        key: 'item_name',
        label: 'Item / Objek',
        render: (v) => `<strong style="color:var(--text-1);">${v || '-'}</strong>`
      },
      {
        key: 'cleanliness_standard',
        label: 'Standar Kebersihan',
        render: (v) => `<span style="line-height:1.5; color:var(--text-2);">${v || '-'}</span>`
      }
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari ruangan, item, atau standar kebersihan...' },
      { type: 'select', name: 'room_name', label: 'Ruangan', options: roomOptions }
    ],
    formFields: (data) => [
      {
        type: 'row',
        fields: [
          {
            name: 'room_name',
            label: 'Ruangan / Kategori',
            type: 'select',
            required: true,
            options: roomOptions,
            value: data?.room_name || ''
          },
          {
            name: 'item_name',
            label: 'Item / Objek',
            required: true,
            placeholder: 'Nama item/objek pembersihan',
            value: data?.item_name || ''
          }
        ]
      },
      {
        name: 'cleanliness_standard',
        label: 'Standar Kebersihan Fisik',
        type: 'textarea',
        required: true,
        rows: 3,
        placeholder: 'Kondisi fisik yang dipersyaratkan (misal: Bebas debu, tidak lengket, kering...)',
        value: data?.cleanliness_standard || ''
      }
    ],
    exportOptions: {
      moduleName: 'hygiene_standards',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/hygiene-standards?all=1&${qs}`);
        if (res.ok) {
          const list = res.data?.data || res.data || [];
          const data = list.map(d => ({
            'Ruangan': d.room_name || '',
            'Item/Objek': d.item_name || '',
            'Standar Kebersihan': d.cleanliness_standard || ''
          }));
          downloadExcel(data, `Master_Hygiene_FCMS_${new Date().toISOString().slice(0, 10)}`);
        } else {
          throw new Error('Gagal mengambil data master hygiene');
        }
      },
      onTemplate: async () => {
        const template = [
          {
            'Ruangan': 'Teras',
            'Item/Objek': 'Sudut/Pojok',
            'Standar Kebersihan': 'Bebas noda, kerak lumut, puntung rokok, dan tumpukan kotoran kering'
          },
          {
            'Ruangan': 'Lobby 1',
            'Item/Objek': 'Lantai Lobby',
            'Standar Kebersihan': 'Mengkilap bersih, bebas minyak, tidak licin, nat ubin cerah'
          }
        ];
        downloadExcel(template, 'Template_Master_Hygiene_FCMS');
      },
      onImport: async (json) => {
        const payload = json.map(row => ({
          room_name: String(row['Ruangan'] || row['room_name'] || '').trim(),
          item_name: String(row['Item/Objek'] || row['Item'] || row['item_name'] || '').trim(),
          cleanliness_standard: String(row['Standar Kebersihan'] || row['Standar'] || row['cleanliness_standard'] || '').trim()
        })).filter(r => r.room_name && r.item_name && r.cleanliness_standard);

        if (payload.length === 0) {
          throw new Error('Tidak ada baris valid untuk diimpor. Pastikan header: Ruangan, Item/Objek, Standar Kebersihan');
        }

        const res = await apiFetch('/api/hygiene-standards/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });

        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
        return res.data;
      }
    }
  });
}
