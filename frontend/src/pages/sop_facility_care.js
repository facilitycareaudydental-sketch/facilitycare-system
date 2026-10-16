import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// ── DEFAULT DATA SETS DARI DOKUMEN SOP - FC & HOUSEKEEPING ───────────────────────
export const DEFAULT_INSTRUKSI_KERJA = [
  {
    id: 1,
    title: 'Dusting (Pembersihan Debu)',
    category: 'Pembersihan Rutin',
    icon: '🧹',
    color: '#2563EB',
    bg: '#EFF6FF',
    target: 'Meja resepsionis, meja pasien, meja dokter, furniture, kaca, monitor & keyboard',
    chemical_ratio: '30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray',
    tools: [
      'Kain Lap Microfiber Biru (Area Umum)',
      'Kanebo (Kaca & Layar Monitor)',
      'Kemoceng',
      'Botol Spray Glass Cleaner',
      'Kuas 4 Inch (Keyboard & Sela Sempit)'
    ],
    steps: [
      'Siapkan larutan glass cleaner di botol spray khusus kaca dan perabot.',
      'Lipat kain microfiber menjadi 8 bidang kerja agar efisien dan bersih merata.',
      'Semprotkan cairan pembersih ke permukaan kain lap, bukan langsung ke monitor atau bodi elektronik.',
      'Usap permukaan perabot dari bagian paling atas ke bagian bawah secara searah atau pola S.',
      'Gunakan kuas cat 4 inch untuk membersihkan sela-sela keyboard dan lubang ventilasi.',
      'Keringkan kaca atau monitor menggunakan lap microfiber kering atau kanebo bersih.'
    ],
    notes: 'Khusus layar monitor medis & dental unit, dilarang menyemprotkan cairan langsung ke layar.'
  },
  {
    id: 2,
    title: 'Sweeping (Lobby Duster & Sapu Nylon)',
    category: 'Pembersihan Rutin',
    icon: '🌾',
    color: '#D97706',
    bg: '#FFFBEB',
    target: 'Lobby utama, lorong, ruang tindakan, balkon, pantry, teras',
    chemical_ratio: '-',
    tools: [
      'Lobby Duster Set 60cm (Area Lobby & Lorong)',
      'Lobby Duster Set 40cm (Area Ruang Tindakan)',
      'Sapu & Pengki Nagata 742s',
      'Kapi Scrapper (Kerak/Permen Karet)'
    ],
    steps: [
      'Pastikan kain lobby duster dalam keadaan kering, bersih, dan bebas pasir kasar.',
      'Dorong lobby duster menyusuri tepi dinding ruangan terlebih dahulu membentuk perimeter.',
      'Lanjutkan mendorong dari sudut terjauh menuju arah pintu keluar tanpa mengangkat kain pel lobi dari lantai.',
      'Kumpulkan debu dan kotoran di dekat pintu keluar, lalu angkat menggunakan sapu dan pengki Nagata.',
      'Jika ada noda kerak atau bekas stiker/permen karet menempel di lantai, kikis perlahan menggunakan kapi scrapper.',
      'Kibaskan kain duster di luar ruangan atau sedot sisa debu pada kain duster.'
    ],
    notes: 'Jangan pernah menggunakan lobby duster di area lantai yang masih basah atau tergenang air.'
  },
  {
    id: 3,
    title: 'Mopping (Pengepelan Lantai)',
    category: 'Pembersihan Rutin',
    icon: '🪣',
    color: '#059669',
    bg: '#ECFDF5',
    target: 'Seluruh lantai ruangan, koridor, dan area tunggu klinik',
    chemical_ratio: '25 ml Floor Cleaner per 1 Liter air bersih',
    tools: [
      'Single Bucket Mop Wringer 20L / Double Bucket 46L',
      'Mopping Stick / Tongkat & Kain Pel',
      'Chemical Floor Cleaner Prokleen Lavender',
      'Warning Sign Lantai Licin'
    ],
    steps: [
      'Pasang papan peringatan "Caution Wet Floor" di titik strategis yang mudah terlihat.',
      'Campurkan cairan Floor Cleaner ke dalam ember air bersih dengan takaran 25 ml per 1 Liter air.',
      'Celupkan kain pel, masukkan ke dalam alat pemeras (wringer), dan peras hingga kondisi kain pel lembap (tidak menetes).',
      'Mulai mengepel dari sudut terjauh ruangan mundur ke arah pintu keluar dengan gerakan membentuk pola angka delapan (8).',
      'Bilas kain pel secara berkala apabila air pada ember bilas sudah terlihat keruh.',
      'Biarkan lantai mengering sempurna secara alami sebelum papan peringatan wet floor dilepas.'
    ],
    notes: 'Wajib pasang papan peringatan "Caution Wet Floor" sebelum mulai mengepel untuk mencegah risiko terpeleset.'
  },
  {
    id: 4,
    title: 'Membersihkan Toilet & Wastafel',
    category: 'Sanitasi Toilet',
    icon: '🚽',
    color: '#DC2626',
    bg: '#FEF2F2',
    target: 'Toilet pasien, toilet tim, closet, wastafel, dinding keramik, floor drain',
    chemical_ratio: 'Karbol 25 ml:1 L Air / Bowl Cleaner 100 ml spray (kerak) / Bubuk KIFA',
    tools: [
      'Sarung Tangan Karet (Wajib)',
      'Sikat Kloset Scotch-Brite 3M',
      'Kain Lap Microfiber Merah (Khusus Toilet)',
      'Sponge Scotch Brite & Sikat Lantai ID-51',
      'Floor Squeegee 75cm',
      'Chemical Karbol Pine, Bowl Cleaner, Bubuk KIFA'
    ],
    steps: [
      'Kenakan sarung tangan karet dan siapkan perlengkapan sanitasi toilet.',
      'Flush kloset, tuangkan Toilet Bowl Cleaner ke bagian dalam dinding kloset dan leher angsa, diamkan 5–10 menit.',
      'Sikat bagian dalam kloset secara menyeluruh dengan sikat kloset, lalu bilas (flush) kembali hingga bersih.',
      'Seka dudukan, penutup, dan tabung luar kloset menggunakan kain lap microfiber merah yang dibasahi larutan disinfektan karbol.',
      'Bersihkan wastafel dan kran stainless menggunakan sponge tapas dengan bubuk pembersih KIFA (jangan gunakan Bowl Cleaner pada kran stainless).',
      'Sikat lantai kamar mandi dengan sikat lantai ID-51 dan cairan karbol disinfektan.',
      'Tarik genangan air ke arah floor drain menggunakan floor squeegee hingga lantai tidak licin dan cepat kering.',
      'Pasang kamper/urinal block, isi ulang sabun cuci tangan, dan pasang plastik sampah baru.'
    ],
    notes: 'Kain lap warna MERAH mutlak hanya untuk toilet dan mangkuk spittoon. Dilarang mencampur cairan kimia sembarangan.'
  },
  {
    id: 5,
    title: 'Pembersihan Area Parkir & Luar',
    category: 'Pembersihan Rutin',
    icon: '🅿️',
    color: '#4F46E5',
    bg: '#EEF2FF',
    target: 'Halaman parkir mobil & motor, saluran air luar, teras depan',
    chemical_ratio: '-',
    tools: [
      'Sapu Lidi Gagang Panjang',
      'Pengki Nagata',
      'Plastik Sampah Hitam Besar (90x120)',
      'Selang Air'
    ],
    steps: [
      'Sapu daun kering, puntung rokok, plastik, dan pasir dari batas trotoar/pagar menuju satu titik.',
      'Kumpulkan dan angkat sampah ke dalam pengki, masukkan langsung ke plastik sampah hitam besar.',
      'Periksa dan bersihkan saluran drainase air di area parkir dari sumbatan sampah padat.',
      'Bila ada tumpahan oli atau kotoran tanah membandel, siram dengan air dan sikat.',
      'Rapikan pembatas parkir dan pastikan area depan klinik terlihat bersih dan rapi menyambut pasien.'
    ],
    notes: 'Lakukan sebelum jam operasional klinik dibuka (pukul 07.30 - 08.00).'
  },
  {
    id: 6,
    title: 'Pembersihan Kaca & Jendela',
    category: 'Pembersihan Rutin',
    icon: '🪟',
    color: '#0891B2',
    bg: '#ECFEFF',
    target: 'Pintu kaca utama, jendela kaca klinik, partisi kaca, cermin toilet',
    chemical_ratio: '30 ml Glass Cleaner : 1 Liter Air / Siap pakai di botol spray',
    tools: [
      'Botol Spray Glass Cleaner Prokleen',
      'Window Wiper / Window Squeegee',
      'Kain Microfiber Kering / Kanebo',
      'Tongkat Teleskopik (Kaca Tinggi)'
    ],
    steps: [
      'Semprotkan larutan glass cleaner merata pada permukaan kaca yang hendak dibersihkan.',
      'Tarik wiper karet dari arah atas ke bawah secara lurus vertikal atau zigzag horizontal tanpa terputus.',
      'Lap bilah karet wiper setiap selesai satu tarikan menggunakan kain lap bersih agar tidak meninggalkan garis bekas air.',
      'Gunakan kain microfiber kering untuk membersihkan sisa tetesan air di sudut-sudut lis bingkai kaca.',
      'Untuk kaca bagian luar yang tinggi, pasang wiper pada tongkat sambungan teleskopik.'
    ],
    notes: 'Hindari membersihkan kaca saat terpapar terik sinar matahari langsung karena cairan akan cepat mengering dan membekas noda.'
  },
  {
    id: 7,
    title: 'Mencuci Bola Mandi Bola (Playground)',
    category: 'Pembersihan Khusus',
    icon: '🤹',
    color: '#7C3AED',
    bg: '#F5F3FF',
    target: 'Area bermain anak (Kids Playground) dan bola-bola mandi bola',
    chemical_ratio: '30 ml Dishwashing Liquid per 1 Liter Air + Karbol Disinfektan',
    tools: [
      'Baskom Plastik Besar 55cm',
      'Dishwashing Liquid Prokleen',
      'Karbol Disinfektan Pine',
      'Lap Microfiber Bersih',
      'Karung / Wadah Pengering Jaring'
    ],
    steps: [
      'Kumpulkan seluruh bola mandi bola ke dalam wadah atau karung bersih.',
      'Siapkan baskom plastik besar berisi air hangat kuku dicampur Dishwashing Liquid dan sedikit karbol disinfektan.',
      'Rendam bola dan cuci secara bertahap sambil digosok perlahan menggunakan sponge/lap bersih.',
      'Bilas bola dengan air bersih mengalir hingga tidak ada sisa busa sabun yang menempel.',
      'Tiriskan dan angin-anginkan bola di tempat bersih beralaskan terpal/kain kering hingga benar-benar kering.',
      'Seka dan sterilkan area kolam mandi bola sebelum bola dimasukkan kembali.'
    ],
    notes: 'Jadwal pencucian berkala dilakukan minimal 1 kali seminggu atau sesuai kebutuhan kebersihan arena anak.'
  },
  {
    id: 8,
    title: 'Membersihkan Filter Air Purifier',
    category: 'Elektronik & Utility',
    icon: '💨',
    color: '#0D9488',
    bg: '#F0FDFA',
    target: 'Unit Air Purifier di ruang tunggu, ruang dokter, dan koridor',
    chemical_ratio: '-',
    tools: [
      'Vacuum Cleaner (Nozzle Sikat Halus)',
      'Kain Microfiber Kering',
      'Kuas Cat Halus'
    ],
    steps: [
      'Matikan saklar power unit air purifier dan cabut kabel daya dari stop kontak dinding.',
      'Buka penutup bodi (casing) depan atau belakang air purifier secara hati-hati.',
      'Keluarkan pre-filter jaring nilon dari unit.',
      'Bersihkan debu tebal pada pre-filter menggunakan vacuum cleaner nozzle sikat halus atau cuci dengan air mengalir (jika tipe wash-able).',
      'Untuk filter HEPA dan Karbon Aktif: HANYA sedot debu permukaannya secara perlahan dengan vacuum, JANGAN DICUCI DENGAN AIR.',
      'Lap bodi luar dan kisi-kisi udara masuk menggunakan lap microfiber kering.',
      'Pastikan pre-filter benar-benar kering sebelum dipasang kembali ke posisinya semula.'
    ],
    notes: 'Jangan pernah mencuci filter HEPA dengan air karena akan merusak serat mikro penyaring bakteri/virus.'
  },
  {
    id: 9,
    title: 'Membersihkan Filter & Bodi AC',
    category: 'Elektronik & Utility',
    icon: '❄️',
    color: '#2563EB',
    bg: '#EFF6FF',
    target: 'Filter udara dan penutup indoor unit AC split di seluruh ruangan',
    chemical_ratio: '-',
    tools: [
      'Tangga Besi Lipat',
      'Kain Lap Microfiber Lembap',
      'Kuas Cat Halus',
      'Kran Air Mengalir'
    ],
    steps: [
      'Matikan AC melalui remote dan matikan MCB/saklar AC bila diperlukan.',
      'Gunakan tangga besi lipat yang stabil di lantai datar.',
      'Buka tutup cover depan unit indoor AC ke atas hingga terkunci.',
      'Lepaskan kedua kisi filter saringan udara plastik secara perlahan agar debu tidak berjatuhan.',
      'Cuci saringan filter di bawah kran air mengalir dari arah bagian dalam ke bagian luar.',
      'Keringkan filter di tempat teduh (jangan dijemur langsung di bawah terik matahari).',
      'Lap bodi casing dan louver/swing AC dengan kain microfiber lembap.',
      'Pasang kembali filter yang sudah kering, tutup cover AC, dan nyalakan AC untuk pengetesan.'
    ],
    notes: 'Pembersihan rutin mandiri dilakukan 2 minggu sekali. Deep cleaning / cuci evaporator dilakukan oleh teknisi AC vendor.'
  },
  {
    id: 10,
    title: 'Membersihkan Dental Unit',
    category: 'Klinis & Medis',
    icon: '🦷',
    color: '#E11D48',
    bg: '#FFF1F2',
    target: 'Kursi dental pasien, spittoon bowl, meja instrumen, lampu dental, cover bawah',
    chemical_ratio: 'Glass Cleaner (netral) / Karbol Disinfektan (mangkuk) / Bowl Cleaner (kerak)',
    tools: [
      'Sarung Tangan Karet (Wajib)',
      'Kain Lap Microfiber Biru (Khusus Bodi Luar)',
      'Kain Lap Microfiber Merah (Khusus Mangkuk Spittoon)',
      'Botol Spray Glass Cleaner (25ml + Air)',
      'Botol Spray Karbol (20ml + Air)',
      'Botol Spray Bowl Cleaner (100ml + Air, bila berkerak)'
    ],
    steps: [
      'Pastikan saklar utama dental unit dalam posisi OFF dan kursi dinaikkan ke posisi tertinggi untuk memudahkan akses.',
      'Lap seluruh jok kulit, sandaran kepala, sandaran tangan, dan tiang lampu menggunakan lap microfiber biru dan cairan pembersih netral/glass cleaner.',
      'Bersihkan mangkuk kumur (spittoon bowl) menggunakan kain lap merah dan cairan karbol disinfektan.',
      'Jika ditemukan kerak sisa darah/plak membandel di mangkuk spittoon, gunakan sedikit bowl cleaner, sikat perlahan, lalu bilas hingga bersih.',
      'Lap permukaan meja instrumen meja dokter. Bila ada instrumen steril, geser dengan hati-hati dan jangan mengotori instrumen.',
      'Buka penutup cover bawah unit dental, lap debu yang terjangkau dan pastikan kondisi area selang dalam keadaan kering.'
    ],
    notes: 'Kain lap warna MERAH hanya untuk mangkuk spittoon. Dilarang keras menggosok jok kulit dental unit dengan bahan kimia asam keras.'
  },
  {
    id: 11,
    title: 'Membersihkan Wastafel & Keran Air',
    category: 'Sanitasi Toilet',
    icon: '🚰',
    color: '#0284C7',
    bg: '#F0F9FF',
    target: 'Wastafel keramik ruang tindakan, wastafel lorong, kran stainless steel, afur sink',
    chemical_ratio: 'Bubuk Pembersih KIFA / Dishwashing Liquid Prokleen',
    tools: [
      'Sponge Scotch Brite Hijau/Kuning',
      'Kain Microfiber Hijau (Area Basah)',
      'Bubuk Pembersih KIFA',
      'Sikat Gigi Bekas / Sikat Kecil'
    ],
    steps: [
      'Basahi seluruh mangkuk wastafel keramik dan kran air stainless steel.',
      'Taburkan sedikit Bubuk KIFA pada tapas sponge yang basah.',
      'Gosok perlahan noda sabun dan kerak air putih pada kran stainless steel searah guratan logam hingga mengkilap kembali.',
      'Gunakan sikat kecil untuk membersihkan sela leher kran dan sudut sambungan afur pembuangan wastafel.',
      'Bilas mangkuk wastafel dan kran dengan air bersih mengalir hingga seluruh sisa residu bubuk KIFA hilang.',
      'Keringkan kran air stainless steel dan cermin wastafel menggunakan lap microfiber hijau kering.'
    ],
    notes: 'DILARANG KERAS menggunakan Toilet Bowl Cleaner pada kran air stainless steel karena menyebabkan karat dan warna menghitam permanen.'
  },
  {
    id: 12,
    title: 'Membersihkan Kulkas Klinik & Pantry',
    category: 'Elektronik & Utility',
    icon: '🧊',
    color: '#475569',
    bg: '#F8FAFC',
    target: 'Kulkas pantry staf dan kulkas penyimpan obat/material medis (eksterior & rak)',
    chemical_ratio: '30 ml Dishwashing Liquid : 1 Liter Air Hangat + Lap Kering',
    tools: [
      'Kain Microfiber Biru & Hijau',
      'Dishwashing Liquid',
      'Ember Air Hangat',
      'Sponge Scotch Brite Halus'
    ],
    steps: [
      'Keluarkan makanan/minuman yang telah kedaluwarsa atau berbau tidak sedap.',
      'Lepas rak-rak kaca/plastik interior kulkas yang dapat dilepas, cuci dengan sabun cuci piring dan bilas air bersih.',
      'Seka dinding interior kulkas menggunakan kain microfiber yang dicelupkan ke dalam air hangat bersabun lembut.',
      'Lap kering seluruh permukaan dalam kulkas dengan microfiber kering bersih.',
      'Pasang kembali rak interior dan susun kembali bahan/makanan secara tertib dan rapi.',
      'Bersihkan karet pintu kulkas dari noda hitam/jamur menggunakan sikat halus dan lap kering.',
      'Lap bodi luar kulkas dan pegangan pintu menggunakan cairan pembersih netral.'
    ],
    notes: 'Untuk kulkas penyimpanan obat/material dental, koordinasikan terlebih dahulu dengan Perawat/PIC Obat sebelum membersihkan rak interior.'
  },
  {
    id: 13,
    title: 'Membersihkan Sofa & Kursi Ruang Tunggu',
    category: 'Pembersihan Khusus',
    icon: '🛋️',
    color: '#9333EA',
    bg: '#FAF5FF',
    target: 'Sofa ruang tunggu pasien, kursi reception, kursi kerja tim',
    chemical_ratio: 'Baking Soda (penghilang bau) / Pembersih Upholstery secukupnya',
    tools: [
      'Vacuum Cleaner (Nozzle Kecil & Sikat Lembut)',
      'Kain Microfiber Bersih',
      'Sikat Bulu Halus',
      'Ember Air Hangat',
      'Cairan Pembersih Khusus Sofa (Upholstery Cleaner)',
      'Baking Soda (Penghilang Bau Tidak Sedap)',
      'Sarung Tangan Karet'
    ],
    steps: [
      'Kenali bahan sofa (kain fabric, beludru, kulit asli, atau kulit sintetis/oscar).',
      'Sedot debu di seluruh permukaan, lipatan jahitan, dan celah bantal sofa menggunakan vacuum cleaner nozzle kecil.',
      'Gunakan sikat bulu halus untuk mengangkat rambut atau kotoran yang menempel kuat di serat kain.',
      'Jika ada noda bercak (kopi, makanan, tinta), semprotkan pembersih upholstery secukupnya, gosok memutar perlahan dengan microfiber.',
      'Celupkan kain microfiber ke air hangat suam-suam kuku, peras hingga lembap, lalu seka seluruh permukaan sofa.',
      'Lap kembali dengan kain kering bersih untuk menyerap sisa kelembapan.',
      'Jika sofa berbau apek: taburkan baking soda tipis di permukaan, diamkan 20 menit, lalu sedot kembali dengan vacuum cleaner.',
      'Biarkan sofa mengering secara alami di dalam ruangan berventilasi atau ber-AC.'
    ],
    notes: 'Jangan menjemur sofa langsung di bawah terik matahari karena dapat menyebabkan bahan kulit pecah-pecah dan warna kain memudar.'
  }
];

export const DEFAULT_ALAT_KEBERSIHAN = [
  { no: 1,  name: 'Sapu Lidi Gagang', qty: '1 unit', area: 'Area parkir mobil & motor klinik' },
  { no: 2,  name: 'Sapu & Pengki Nagata 742s', qty: '1 set', area: 'Area pantry, balkon, dak lantai atas, atau teras' },
  { no: 3,  name: 'Lobby Duster Set 60 cm', qty: '1 set', area: 'Lobby utama dan koridor/lorong klinik' },
  { no: 4,  name: 'Lobby Duster Set 40 cm / Refill', qty: '1 unit', area: 'Area ruang tindakan dokter' },
  { no: 5,  name: 'Single Bucket Mop Wringer 20L', qty: '1 paket', area: 'Ember pel dengan pemeras + tongkat pel' },
  { no: 6,  name: 'Double Bucket Wringer Trolley 46L', qty: '1 paket', area: 'Khusus klinik besar / klinik dengan fasilitas lift' },
  { no: 7,  name: 'Kemoceng Bulu Lembut', qty: '1 unit', area: 'Membersihkan debu kasar meja dan perabot' },
  { no: 8,  name: 'Lap Microfiber Warna Merah', qty: '2 pcs', area: 'Khusus area toilet, sanitasi, dan mangkuk spittoon' },
  { no: 9,  name: 'Lap Microfiber Warna Biru', qty: '2 pcs', area: 'Area umum kering (meja, kursi, bodi peralatan)' },
  { no: 10, name: 'Lap Microfiber Warna Hijau', qty: '2 pcs', area: 'Area umum basah (pantry, wastafel luar)' },
  { no: 11, name: 'Kanebo / Synthetic Chamois', qty: '1 pcs', area: 'Area umum, cermin, dan pengering kaca' },
  { no: 12, name: 'Caddy Bag Cleaning', qty: '1 unit', area: 'Tas perlengkapan pinggang petugas cleaning' },
  { no: 13, name: 'Sponge Scotch Brite 3M', qty: '2 pcs', area: 'Pembersihan noda lantai dan kerak ringan (spotting)' },
  { no: 14, name: 'Scotch-Brite 3M Sikat Toilet', qty: '1 unit', area: 'Sikat khusus kloset dan leher angsa' },
  { no: 15, name: 'Kapi Gagang Karet / Scrapper', qty: '1 unit', area: 'Pembersihan kerak lantai, noda permen karet, stiker' },
  { no: 16, name: 'Kuas Cat 4 Inch', qty: '2 pcs', area: 'Pembersih sela keyboard, ventilasi, dan sudut sempit' },
  { no: 17, name: 'Warning Sign (Lantai Licin)', qty: '1 unit', area: 'Digunakan saat mopping dan general cleaning' },
  { no: 18, name: 'Teleskopik + Pembersih Kaca/Langit', qty: '1 set', area: 'Membersihkan jendela tinggi, plafon, dan sarang laba-laba' },
  { no: 19, name: 'Sarung Tangan Karet 3M', qty: '1 pasang', area: 'Perlindungan tangan area toilet dan bahan kimia' },
  { no: 20, name: 'Botol Spray 500ml', qty: '5 unit', area: 'Wadah larutan chemical (diberi label masing-masing)' },
  { no: 21, name: 'Pengharum Ruangan Stella Aerosol', qty: '1 kaleng', area: 'Digunakan darurat apabila timbul bau menyengat' },
  { no: 22, name: 'Kleaner Gantungan Sapu & Pel', qty: '1 set', area: 'Gantungan penyimpanan alat cleaning di janitor' },
  { no: 23, name: 'Floor Squeegee 75 cm', qty: '1 unit', area: 'Pendorong genangan air lantai toilet/teras' },
  { no: 24, name: 'Sikat Botol Ikea', qty: '1 unit', area: 'Di setiap ruang tindakan dokter gigi' },
  { no: 25, name: 'Scotch-Brite 3M Sikat Lantai ID-51', qty: '1 unit', area: 'Sikat lantai kamar mandi dan teras' },
  { no: 26, name: 'Kamper Toilet / Urinal Block', qty: '1 pack', area: 'Pengharum di toilet dan urinoir' },
  { no: 27, name: 'Plastik Sampah Hitam Sedang (60x100)', qty: '1 pack', area: 'Tempat sampah indoor ruangan' },
  { no: 28, name: 'Plastik Sampah Hitam Besar (90x120)', qty: '1 pack', area: 'Tempat sampah luar & penampung sampah akhir' },
  { no: 29, name: 'Plastik Sampah Medis Kuning (90x120)', qty: '1 pack', area: 'Limbah medis padat B3 klinik' },
  { no: 30, name: 'Plastik Sampah Medis Kuning (40x60)', qty: '1 pack', area: 'Tempat sampah medis kecil di ruang tindakan' },
  { no: 31, name: 'Tangga Besi Lipat', qty: '1 unit', area: 'Alat bantu membersihkan plafon, lampu, dan AC' },
  { no: 32, name: 'Baskom Plastik Besar 55cm', qty: '1 unit', area: 'Mencuci bola kids playground' }
];

export const DEFAULT_BAHAN_CHEMICAL = [
  { no: 1, name: 'Floor Cleaner Prokleen Lavender', size: '5 Liter', ratio: '25 ml : 1 Liter Air', function: 'Pembersihan dan pengepelan rutin seluruh lantai ruangan' },
  { no: 2, name: 'Karbol Prokleen Pine Disinfektan', size: '5 Liter', ratio: '25 ml : 1 Liter Air (Mop) / 20 ml per botol spray', function: 'Mengepel lantai toilet, disinfeksi area basah dan berbau tak sedap' },
  { no: 3, name: 'Glass Cleaner Prokleen', size: '5 Liter', ratio: '30 ml : 1 Liter Air / Siap pakai di botol spray', function: 'Membersihkan pintu kaca, jendela, cermin, dan layar monitor' },
  { no: 4, name: 'Manual Dishwashing Liquid Prokleen', size: '5 Liter', ratio: '30 ml : 1 Liter Air', function: 'Mencuci piring, gelas, sendok pantry, dan mencuci bola playground' },
  { no: 5, name: 'Hand Soap Prokleen Green Tea', size: '5 Liter', ratio: 'Siap pakai (tanpa campur)', function: 'Refill sabun pencuci tangan di wastafel pasien dan tim' },
  { no: 6, name: 'Liquid Detergent Prokleen', size: '5 Liter', ratio: 'Sesuai kapasitas mesin', function: 'Mencuci handuk, kain microfiber, dan seragam' },
  { no: 7, name: 'Toilet Bowl Cleaner Prokleen', size: '5 Liter', ratio: '100 ml per botol spray / Khusus kerak', function: 'Pembersih kerak membandel pada kloset, wastafel keramik, dan nat keramik' },
  { no: 8, name: 'Bubuk Pembersih KIFA', size: '500 gram', ratio: 'Taburkan tipis pada tapas basah', function: 'Membersihkan kerak air pada kran air stainless steel, jet washer, dan sink' }
];

// ── LOCAL STORAGE REPOSITORY ────────────────────────────────────────────────────
const KEY_IK = 'fcms_sop_fc_ik_v1';
const KEY_TOOLS = 'fcms_sop_fc_tools_v1';
const KEY_CHEM = 'fcms_sop_fc_chem_v1';

function getStoredIK() {
  try {
    const raw = localStorage.getItem(KEY_IK);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.error('Error load IK', e); }
  return JSON.parse(JSON.stringify(DEFAULT_INSTRUKSI_KERJA));
}

function saveStoredIK(list) {
  localStorage.setItem(KEY_IK, JSON.stringify(list));
}

function getStoredTools() {
  try {
    const raw = localStorage.getItem(KEY_TOOLS);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.error('Error load tools', e); }
  return JSON.parse(JSON.stringify(DEFAULT_ALAT_KEBERSIHAN));
}

function saveStoredTools(list) {
  localStorage.setItem(KEY_TOOLS, JSON.stringify(list));
}

function getStoredChem() {
  try {
    const raw = localStorage.getItem(KEY_CHEM);
    if (raw) return JSON.parse(raw);
  } catch (e) { console.error('Error load chem', e); }
  return JSON.parse(JSON.stringify(DEFAULT_BAHAN_CHEMICAL));
}

function saveStoredChem(list) {
  localStorage.setItem(KEY_CHEM, JSON.stringify(list));
}

// ── MAIN RENDER FUNCTION ────────────────────────────────────────────────────────
export async function renderSOPFacilityCare(container) {
  let activeTab = 'ik'; // 'ik' | 'tools' | 'policy' | 'links'
  let currentSearch = '';
  let currentCategory = 'all';
  let toolView = 'tools'; // 'tools' | 'chemical'
  let toolSearch = '';

  let listIK = getStoredIK();
  let listTools = getStoredTools();
  let listChem = getStoredChem();

  function renderPage() {
    container.innerHTML = `
      <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
        <div>
          <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
            <span>📚</span> SOP Facility Care & Housekeeping
          </h1>
          <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
            Pedoman Standar Operasional Prosedur Kebersihan & Pemeliharaan Fasilitas PT Audy Mandiri Indonesia
          </p>
        </div>
        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <button id="btn-reset-default" class="btn btn-outline" title="Kembalikan data ke dokumen standar asli" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            🔄 Reset ke Standar
          </button>
          <button id="btn-print-sop" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            🖨️ Cetak Ringkasan SOP
          </button>
          <button id="btn-export-tools" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            📥 Download Data Excel
          </button>
        </div>
      </div>

      <!-- Navigation Tabs -->
      <div style="display:flex; gap:8px; margin-bottom:22px; border-bottom:2px solid #E2E8F0; padding-bottom:8px; overflow-x:auto;">
        <button class="sop-tab-btn ${activeTab === 'ik' ? 'active' : ''}" data-tab="ik" style="${getTabStyle(activeTab === 'ik')}">
          📋 Instruksi Kerja (IK) <span class="badge" style="${getBadgeStyle(activeTab === 'ik')}">${listIK.length}</span>
        </button>
        <button class="sop-tab-btn ${activeTab === 'tools' ? 'active' : ''}" data-tab="tools" style="${getTabStyle(activeTab === 'tools')}">
          🛠️ Alat & Chemical <span class="badge" style="${getBadgeStyle(activeTab === 'tools')}">${listTools.length + listChem.length}</span>
        </button>
        <button class="sop-tab-btn ${activeTab === 'policy' ? 'active' : ''}" data-tab="policy" style="${getTabStyle(activeTab === 'policy')}">
          📖 Pedoman & Ketentuan
        </button>
        <button class="sop-tab-btn ${activeTab === 'links' ? 'active' : ''}" data-tab="links" style="${getTabStyle(activeTab === 'links')}">
          🔗 Form & Tautan
        </button>
      </div>

      <!-- Content Container -->
      <div id="sop-content-area"></div>
    `;

    bindTabEvents();
    renderActiveTabContent();
  }

  function getTabStyle(isActive) {
    if (isActive) {
      return 'border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;';
    }
    return 'border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem; white-space:nowrap; transition:all 0.2s;';
  }

  function getBadgeStyle(isActive) {
    if (isActive) {
      return 'background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;';
    }
    return 'background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 7px; border-radius:10px; font-weight:700;';
  }

  function bindTabEvents() {
    container.querySelectorAll('.sop-tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        activeTab = btn.dataset.tab;
        renderPage();
      });
    });

    const btnReset = container.querySelector('#btn-reset-default');
    if (btnReset) {
      btnReset.addEventListener('click', () => {
        confirmDialog('Apakah Anda yakin ingin mengembalikan seluruh data Instruksi Kerja, Alat, dan Chemical ke data standar dokumen asli?', () => {
          localStorage.removeItem(KEY_IK);
          localStorage.removeItem(KEY_TOOLS);
          localStorage.removeItem(KEY_CHEM);
          listIK = getStoredIK();
          listTools = getStoredTools();
          listChem = getStoredChem();
          toastSuccess('Data SOP Facility Care berhasil di-reset ke standar asli');
          renderPage();
        }, 'Konfirmasi Reset Standar');
      });
    }

    const btnPrint = container.querySelector('#btn-print-sop');
    if (btnPrint) {
      btnPrint.addEventListener('click', () => {
        window.print();
      });
    }

    const btnExport = container.querySelector('#btn-export-tools');
    if (btnExport) {
      btnExport.addEventListener('click', () => {
        exportAllData();
      });
    }
  }

  function renderActiveTabContent() {
    const content = container.querySelector('#sop-content-area');
    if (!content) return;

    if (activeTab === 'ik') {
      renderIKTab(content);
    } else if (activeTab === 'tools') {
      renderToolsTab(content);
    } else if (activeTab === 'policy') {
      renderPolicyTab(content);
    } else if (activeTab === 'links') {
      renderLinksTab(content);
    }
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 1: INSTRUKSI KERJA DENGAN TAMBAH, EDIT & HAPUS
  // ─────────────────────────────────────────────────────────────────────────────
  function renderIKTab(parent) {
    parent.innerHTML = `
      <!-- Toolbar Filter & Search -->
      <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
          <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
            <div style="position:relative; flex:1; min-width:220px; max-width:340px;">
              <input type="text" id="ik-search" class="form-control" placeholder="Cari pekerjaan, alat, atau chemical..." value="${currentSearch}" style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
              <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">🔍</span>
            </div>
            
            <div style="display:flex; gap:6px; flex-wrap:wrap;" id="cat-pills-wrap">
              <button class="cat-pill ${currentCategory === 'all' ? 'active' : ''}" data-cat="all">Semua (${listIK.length})</button>
              <button class="cat-pill ${currentCategory === 'Pembersihan Rutin' ? 'active' : ''}" data-cat="Pembersihan Rutin">Rutin (${listIK.filter(x => x.category === 'Pembersihan Rutin').length})</button>
              <button class="cat-pill ${currentCategory === 'Sanitasi Toilet' ? 'active' : ''}" data-cat="Sanitasi Toilet">Toilet & Wastafel (${listIK.filter(x => x.category === 'Sanitasi Toilet').length})</button>
              <button class="cat-pill ${currentCategory === 'Klinis & Medis' ? 'active' : ''}" data-cat="Klinis & Medis">Klinis Dental (${listIK.filter(x => x.category === 'Klinis & Medis').length})</button>
              <button class="cat-pill ${currentCategory === 'Elektronik & Utility' ? 'active' : ''}" data-cat="Elektronik & Utility">Elektronik & AC (${listIK.filter(x => x.category === 'Elektronik & Utility').length})</button>
              <button class="cat-pill ${currentCategory === 'Pembersihan Khusus' ? 'active' : ''}" data-cat="Pembersihan Khusus">Khusus (${listIK.filter(x => x.category === 'Pembersihan Khusus').length})</button>
            </div>
          </div>
          <div style="display:flex; gap:10px; align-items:center;">
            <div id="ik-count-label" style="font-size:0.85rem; color:var(--text-3); font-weight:600;">
              Menampilkan ${listIK.length} dari ${listIK.length} Instruksi Kerja
            </div>
            <button id="btn-add-ik" class="btn btn-primary" style="font-weight:600; padding:7px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
              ➕ Tambah IK Baru
            </button>
          </div>
        </div>
      </div>

      <!-- IK Cards Grid Container -->
      <div id="ik-cards-grid-container"></div>
    `;

    // Inject Cat pill styles
    const styleTag = document.createElement('style');
    styleTag.innerHTML = `
      .cat-pill {
        border: 1px solid var(--border);
        background: #F8FAFC;
        color: var(--text-2);
        font-size: 0.8rem;
        font-weight: 600;
        padding: 5px 12px;
        border-radius: 20px;
        cursor: pointer;
        transition: all 0.15s;
      }
      .cat-pill:hover {
        background: #E2E8F0;
      }
      .cat-pill.active {
        background: #2563EB;
        color: #fff;
        border-color: #2563EB;
      }
      .ik-card:hover {
        box-shadow: 0 6px 16px rgba(0,0,0,0.08) !important;
        transform: translateY(-2px);
      }
    `;
    parent.appendChild(styleTag);

    const gridContainer = parent.querySelector('#ik-cards-grid-container');
    const countLabel = parent.querySelector('#ik-count-label');

    function updateIKGrid() {
      const filteredIK = listIK.filter(item => {
        const matchCat = currentCategory === 'all' || item.category === currentCategory;
        const q = currentSearch.toLowerCase().trim();
        const matchQuery = !q ||
          item.title.toLowerCase().includes(q) ||
          item.target.toLowerCase().includes(q) ||
          item.tools.some(t => t.toLowerCase().includes(q)) ||
          item.steps.some(s => s.toLowerCase().includes(q));
        return matchCat && matchQuery;
      });

      if (countLabel) {
        countLabel.textContent = `Menampilkan ${filteredIK.length} dari ${listIK.length} Instruksi Kerja`;
      }

      if (filteredIK.length === 0) {
        gridContainer.innerHTML = `
          <div style="text-align:center; padding:48px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
            <div style="font-size:2.5rem; margin-bottom:8px;">🔍</div>
            <h3 style="margin:0 0 4px; font-size:1.1rem; color:var(--text-1);">Tidak ada Instruksi Kerja yang cocok</h3>
            <p style="margin:0; font-size:0.85rem; color:var(--text-3);">Coba ubah kata kunci pencarian atau pilih kategori lain.</p>
          </div>
        `;
        return;
      }

      gridContainer.innerHTML = `
        <div style="display:grid; grid-template-columns:repeat(auto-fill, minmax(360px, 1fr)); gap:18px;">
          ${filteredIK.map(ik => `
            <div class="card ik-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:18px; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 1px 4px rgba(0,0,0,0.04); transition:all 0.2s;">
              <div>
                <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                  <div style="display:flex; align-items:center; gap:10px;">
                    <span style="font-size:1.6rem; width:44px; height:44px; display:inline-flex; align-items:center; justify-content:center; background:${ik.bg || '#EFF6FF'}; border-radius:10px;">${ik.icon || '📋'}</span>
                    <div>
                      <h3 style="margin:0 0 2px 0; font-size:1.05rem; font-weight:700; color:var(--text-1);">
                        ${ik.id}. ${ik.title}
                      </h3>
                      <span class="badge" style="background:${ik.bg || '#EFF6FF'}; color:${ik.color || '#2563EB'}; font-size:0.75rem; font-weight:700; border-radius:6px; padding:2px 8px;">
                        ${ik.category}
                      </span>
                    </div>
                  </div>

                  <!-- Quick Action Buttons: Edit & Hapus -->
                  <div style="display:flex; gap:4px;">
                    <button class="btn btn-sm btn-ghost btn-edit-ik" data-id="${ik.id}" title="Edit Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; cursor:pointer;">
                      ✏️
                    </button>
                    <button class="btn btn-sm btn-ghost btn-delete-ik" data-id="${ik.id}" title="Hapus Instruksi Kerja" style="padding:4px 8px; border-radius:6px; font-size:0.9rem; color:#DC2626; cursor:pointer;">
                      🗑️
                    </button>
                  </div>
                </div>

                <div style="font-size:0.85rem; color:var(--text-2); margin-bottom:14px; background:#F8FAFC; padding:8px 12px; border-radius:8px; border-left:3px solid ${ik.color || '#2563EB'};">
                  <strong>Area:</strong> ${ik.target}
                </div>

                <div style="margin-bottom:14px;">
                  <div style="font-size:0.8rem; font-weight:700; color:var(--text-3); text-transform:uppercase; margin-bottom:6px;">🧰 Alat & Bahan Utama:</div>
                  <ul style="margin:0 0 0 16px; padding:0; font-size:0.83rem; color:var(--text-2);">
                    ${ik.tools.slice(0, 3).map(t => `<li style="margin-bottom:2px;">${t}</li>`).join('')}
                    ${ik.tools.length > 3 ? `<li style="color:var(--text-3); font-style:italic;">+ ${ik.tools.length - 3} alat lainnya...</li>` : ''}
                  </ul>
                </div>

                ${ik.chemical_ratio && ik.chemical_ratio !== '-' ? `
                  <div style="font-size:0.8rem; background:#FEF3C7; color:#92400E; padding:6px 10px; border-radius:6px; margin-bottom:14px; font-weight:600;">
                    🧪 Takaran: ${ik.chemical_ratio}
                  </div>
                ` : ''}
              </div>

              <div style="border-top:1px solid #F1F5F9; padding-top:12px; margin-top:8px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
                  ${ik.steps.length} Langkah Kerja
                </span>
                <div style="display:flex; gap:6px;">
                  <button class="btn btn-sm btn-primary btn-view-ik" data-id="${ik.id}" style="padding:6px 14px; font-weight:600; border-radius:6px;">
                    Lihat Panduan ➔
                  </button>
                </div>
              </div>
            </div>
          `).join('')}
        </div>
      `;

      // Event detail modal
      gridContainer.querySelectorAll('.btn-view-ik').forEach(btn => {
        btn.addEventListener('click', () => {
          const ikId = parseInt(btn.dataset.id, 10);
          const ik = listIK.find(x => x.id === ikId);
          if (ik) openIKDetailModal(ik);
        });
      });

      // Event edit modal
      gridContainer.querySelectorAll('.btn-edit-ik').forEach(btn => {
        btn.addEventListener('click', () => {
          const ikId = parseInt(btn.dataset.id, 10);
          const ik = listIK.find(x => x.id === ikId);
          if (ik) openIKFormModal(ik);
        });
      });

      // Event delete
      gridContainer.querySelectorAll('.btn-delete-ik').forEach(btn => {
        btn.addEventListener('click', () => {
          const ikId = parseInt(btn.dataset.id, 10);
          const ik = listIK.find(x => x.id === ikId);
          if (!ik) return;
          confirmDialog(`Apakah Anda yakin ingin menghapus Instruksi Kerja: "${ik.title}"?`, () => {
            listIK = listIK.filter(x => x.id !== ikId);
            saveStoredIK(listIK);
            toastSuccess('Instruksi Kerja berhasil dihapus');
            renderPage();
          }, 'Hapus Instruksi Kerja');
        });
      });
    }

    updateIKGrid();

    // Event search
    const searchInput = parent.querySelector('#ik-search');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        currentSearch = e.target.value;
        updateIKGrid();
      });
    }

    // Event cat pills
    parent.querySelectorAll('.cat-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        parent.querySelectorAll('.cat-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        currentCategory = pill.dataset.cat;
        updateIKGrid();
      });
    });

    // Event Add New IK
    const btnAddIK = parent.querySelector('#btn-add-ik');
    if (btnAddIK) {
      btnAddIK.addEventListener('click', () => {
        openIKFormModal();
      });
    }
  }

  // MODAL VIEW DETAIL IK
  function openIKDetailModal(ik) {
    createModal({
      title: `${ik.icon || '📋'} IK-${ik.id}: ${ik.title}`,
      size: 'lg',
      content: `
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="background:${ik.bg || '#EFF6FF'}; color:${ik.color || '#2563EB'}; padding:10px 14px; border-radius:8px; font-weight:700; font-size:0.9rem; display:flex; justify-content:space-between; align-items:center;">
            <span>Kategori: ${ik.category}</span>
            <span>ID: IK-${String(ik.id).padStart(2, '0')}</span>
          </div>

          <div style="background:#F8FAFC; padding:12px 14px; border-radius:8px; border-left:4px solid ${ik.color || '#2563EB'};">
            <strong style="color:var(--text-1); font-size:0.85rem; display:block; margin-bottom:4px;">📍 Area & Sasaran Target:</strong>
            <span style="font-size:0.88rem; color:var(--text-2);">${ik.target}</span>
          </div>

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">🧰 Alat & Bahan yang Dibutuhkan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 14px;">
              <ul style="margin:0 0 0 16px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.7;">
                ${ik.tools.map(t => `<li>${t}</li>`).join('')}
              </ul>
            </div>
          </div>

          ${ik.chemical_ratio && ik.chemical_ratio !== '-' ? `
            <div style="background:#FEF3C7; border:1px solid #FCD34D; border-radius:8px; padding:10px 14px;">
              <strong style="color:#92400E; font-size:0.85rem; display:block; margin-bottom:2px;">🧪 Rasio & Takaran Chemical:</strong>
              <span style="color:#78350F; font-size:0.9rem; font-weight:600;">${ik.chemical_ratio}</span>
            </div>
          ` : ''}

          <div>
            <h4 style="margin:0 0 8px 0; font-size:0.95rem; font-weight:700; color:var(--text-1);">📝 Langkah-Langkah Pelaksanaan:</h4>
            <div style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:12px 16px;">
              <ol style="margin:0 0 0 18px; padding:0; font-size:0.875rem; color:var(--text-2); line-height:1.8;">
                ${ik.steps.map(s => `<li style="margin-bottom:6px;">${s}</li>`).join('')}
              </ol>
            </div>
          </div>

          ${ik.notes ? `
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; padding:10px 14px;">
              <strong style="color:#1E40AF; font-size:0.85rem; display:block; margin-bottom:2px;">⚠️ Catatan Penting & Safety:</strong>
              <span style="color:#1E3A8A; font-size:0.85rem; line-height:1.6;">${ik.notes}</span>
            </div>
          ` : ''}
        </div>
      `,
      confirmText: 'Tutup Panduan',
      onConfirm: (overlay, close) => {
        close();
      }
    });
  }

  // MODAL TAMBAH & EDIT IK
  function openIKFormModal(existingIK = null) {
    const isEdit = !!existingIK;
    const initialTools = existingIK ? existingIK.tools.join('\n') : '';
    const initialSteps = existingIK ? existingIK.steps.join('\n') : '';

    const contentHtml = `
      <form id="ik-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div style="display:grid; grid-template-columns:1fr 140px; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Judul Instruksi Kerja *</label>
            <input type="text" id="form-ik-title" class="form-control" required value="${existingIK ? escapeHtml(existingIK.title) : ''}" placeholder="Contoh: Membersihkan Dental Unit" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Ikon Emoji</label>
            <input type="text" id="form-ik-icon" class="form-control" value="${existingIK ? existingIK.icon : '📋'}" placeholder="🧹" style="width:100%; border-radius:6px; height:36px; text-align:center; font-size:1.2rem;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kategori *</label>
            <select id="form-ik-cat" class="form-control" style="width:100%; border-radius:6px; height:36px;">
              <option value="Pembersihan Rutin" ${existingIK?.category === 'Pembersihan Rutin' ? 'selected' : ''}>Pembersihan Rutin</option>
              <option value="Sanitasi Toilet" ${existingIK?.category === 'Sanitasi Toilet' ? 'selected' : ''}>Sanitasi Toilet</option>
              <option value="Klinis & Medis" ${existingIK?.category === 'Klinis & Medis' ? 'selected' : ''}>Klinis & Medis</option>
              <option value="Elektronik & Utility" ${existingIK?.category === 'Elektronik & Utility' ? 'selected' : ''}>Elektronik & Utility</option>
              <option value="Pembersihan Khusus" ${existingIK?.category === 'Pembersihan Khusus' ? 'selected' : ''}>Pembersihan Khusus</option>
            </select>
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Takaran / Rasio Chemical</label>
            <input type="text" id="form-ik-ratio" class="form-control" value="${existingIK ? escapeHtml(existingIK.chemical_ratio || '') : ''}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Area & Sasaran Target *</label>
          <input type="text" id="form-ik-target" class="form-control" required value="${existingIK ? escapeHtml(existingIK.target) : ''}" placeholder="Contoh: Kursi pasien, meja instrumen, mangkuk spittoon" style="width:100%; border-radius:6px; height:36px;">
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Alat & Bahan yang Digunakan (1 item per baris) *</label>
          <textarea id="form-ik-tools" class="form-control" rows="4" placeholder="Kain Microfiber Biru&#10;Sarung Tangan Karet&#10;Botol Spray Glass Cleaner" style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.5;">${escapeHtml(initialTools)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Langkah-Langkah Pelaksanaan (1 langkah per baris) *</label>
          <textarea id="form-ik-steps" class="form-control" rows="6" placeholder="1. Pastikan posisi saklar OFF.&#10;2. Lap jok dental unit dengan kain microfiber biru.&#10;3. Sikat mangkuk spittoon..." style="width:100%; border-radius:6px; font-size:0.85rem; line-height:1.6;">${escapeHtml(initialSteps)}</textarea>
        </div>

        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Catatan Penting / Hal yang Perlu Diperhatikan (Safety)</label>
          <textarea id="form-ik-notes" class="form-control" rows="2" placeholder="Catatan larangan atau peringatan keselamatan..." style="width:100%; border-radius:6px; font-size:0.85rem;">${existingIK ? escapeHtml(existingIK.notes || '') : ''}</textarea>
        </div>
      </form>
    `;

    createModal({
      title: isEdit ? `✏️ Edit IK-${existingIK.id}: ${existingIK.title}` : `➕ Tambah Instruksi Kerja Baru`,
      content: contentHtml,
      size: 'lg',
      confirmText: 'Simpan Perubahan',
      onConfirm: (overlay, close) => {
        const title = overlay.querySelector('#form-ik-title').value.trim();
        const icon = overlay.querySelector('#form-ik-icon').value.trim() || '📋';
        const category = overlay.querySelector('#form-ik-cat').value;
        const chemical_ratio = overlay.querySelector('#form-ik-ratio').value.trim();
        const target = overlay.querySelector('#form-ik-target').value.trim();
        const toolsRaw = overlay.querySelector('#form-ik-tools').value.trim();
        const stepsRaw = overlay.querySelector('#form-ik-steps').value.trim();
        const notes = overlay.querySelector('#form-ik-notes').value.trim();

        if (!title || !target || !toolsRaw || !stepsRaw) {
          toastError('Harap lengkapi seluruh field wajib (*)');
          return;
        }

        const tools = toolsRaw.split('\n').map(s => s.trim().replace(/^[-•*]\s*/, '')).filter(Boolean);
        const steps = stepsRaw.split('\n').map(s => s.trim().replace(/^\d+[\.\)]\s*/, '')).filter(Boolean);

        const categoryColors = {
          'Pembersihan Rutin': { color: '#2563EB', bg: '#EFF6FF' },
          'Sanitasi Toilet': { color: '#DC2626', bg: '#FEF2F2' },
          'Klinis & Medis': { color: '#E11D48', bg: '#FFF1F2' },
          'Elektronik & Utility': { color: '#0D9488', bg: '#F0FDFA' },
          'Pembersihan Khusus': { color: '#7C3AED', bg: '#F5F3FF' },
        };
        const styling = categoryColors[category] || { color: '#2563EB', bg: '#EFF6FF' };

        if (isEdit) {
          existingIK.title = title;
          existingIK.icon = icon;
          existingIK.category = category;
          existingIK.chemical_ratio = chemical_ratio;
          existingIK.target = target;
          existingIK.tools = tools;
          existingIK.steps = steps;
          existingIK.notes = notes;
          existingIK.color = styling.color;
          existingIK.bg = styling.bg;
          saveStoredIK(listIK);
          toastSuccess(`Instruksi Kerja "${title}" berhasil diperbarui`);
        } else {
          const nextId = listIK.length > 0 ? Math.max(...listIK.map(x => x.id)) + 1 : 1;
          const newIK = {
            id: nextId,
            title,
            category,
            icon,
            color: styling.color,
            bg: styling.bg,
            target,
            chemical_ratio: chemical_ratio || '-',
            tools,
            steps,
            notes
          };
          listIK.push(newIK);
          saveStoredIK(listIK);
          toastSuccess(`Instruksi Kerja baru IK-${nextId} berhasil ditambahkan`);
        }

        close();
        renderPage();
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 2: ALAT & CHEMICAL DENGAN TAMBAH, EDIT & HAPUS
  // ─────────────────────────────────────────────────────────────────────────────
  function renderToolsTab(parent) {
    parent.innerHTML = `
      <!-- Sub-tab Switcher & Actions -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px; margin-bottom:18px;">
        <div class="btn-group" style="background:#E2E8F0; padding:3px; border-radius:8px; display:inline-flex;">
          <button id="view-tools-btn" class="btn btn-sm ${toolView === 'tools' ? 'btn-primary' : 'btn-ghost'}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            🧹 Alat Kebersihan (${listTools.length} Item)
          </button>
          <button id="view-chem-btn" class="btn btn-sm ${toolView === 'chemical' ? 'btn-primary' : 'btn-ghost'}" style="font-weight:700; border-radius:6px; padding:6px 16px; cursor:pointer;">
            🧪 Bahan Chemical (${listChem.length} Item)
          </button>
        </div>

        <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
          <input type="text" id="tool-search" class="form-control" placeholder="Cari nama / peruntukan / rasio..." value="${toolSearch}" style="height:36px; border-radius:8px; font-size:0.875rem; width:240px;">
          <button id="btn-add-tool-chem" class="btn btn-primary" style="font-weight:600; padding:6px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; font-size:0.85rem;">
            ➕ ${toolView === 'tools' ? 'Tambah Alat Baru' : 'Tambah Chemical Baru'}
          </button>
        </div>
      </div>

      <!-- Table Container -->
      <div id="tool-table-container"></div>
    `;

    const subContent = parent.querySelector('#tool-table-container');
    const searchInput = parent.querySelector('#tool-search');

    function renderSubTable() {
      const q = toolSearch.toLowerCase().trim();
      if (toolView === 'tools') {
        const filtered = listTools.filter(x => !q || x.name.toLowerCase().includes(q) || x.area.toLowerCase().includes(q));
        subContent.innerHTML = `
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Standar Peralatan Kebersihan (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-primary">${filtered.length} Item</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
                <thead>
                  <tr style="background:#F1F5F9; border-bottom:2px solid #E2E8F0; text-align:left;">
                    <th style="padding:10px 14px; width:60px;">No</th>
                    <th style="padding:10px 14px;">Nama Alat / Item</th>
                    <th style="padding:10px 14px; width:130px;">Kuantitas Standar</th>
                    <th style="padding:10px 14px;">Peruntukan Area & Kegunaan</th>
                    <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${filtered.map((item, idx) => `
                    <tr style="border-bottom:1px solid #F1F5F9; background:${idx % 2 === 0 ? '#fff' : '#FAFAFA'};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${item.no || idx + 1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${escapeHtml(item.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#E0F2FE; color:#0369A1; font-weight:600;">${escapeHtml(item.qty)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${escapeHtml(item.area)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-tool" data-no="${item.no}" title="Edit Alat" style="padding:4px 6px; cursor:pointer;">✏️</button>
                        <button class="btn btn-sm btn-ghost btn-delete-tool" data-no="${item.no}" title="Hapus Alat" style="padding:4px 6px; color:#DC2626; cursor:pointer;">🗑️</button>
                      </td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>
        `;

        // Bind Tool actions
        subContent.querySelectorAll('.btn-edit-tool').forEach(btn => {
          btn.addEventListener('click', () => {
            const no = parseInt(btn.dataset.no, 10);
            const tool = listTools.find(x => x.no === no);
            if (tool) openToolFormModal(tool);
          });
        });

        subContent.querySelectorAll('.btn-delete-tool').forEach(btn => {
          btn.addEventListener('click', () => {
            const no = parseInt(btn.dataset.no, 10);
            const tool = listTools.find(x => x.no === no);
            if (!tool) return;
            confirmDialog(`Apakah Anda yakin ingin menghapus alat: "${tool.name}"?`, () => {
              listTools = listTools.filter(x => x.no !== no);
              // re-number
              listTools.forEach((t, i) => { t.no = i + 1; });
              saveStoredTools(listTools);
              toastSuccess('Alat kebersihan berhasil dihapus');
              renderPage();
            }, 'Hapus Alat Kebersihan');
          });
        });

      } else {
        const filtered = listChem.filter(x => !q || x.name.toLowerCase().includes(q) || x.function.toLowerCase().includes(q));
        subContent.innerHTML = `
          <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div style="padding:14px 18px; background:#F8FAFC; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
              <strong style="color:var(--text-1); font-size:0.95rem;">Daftar Bahan Kimia & Standar Takaran (Lampiran 8.1 SOP)</strong>
              <span class="badge badge-success">${filtered.length} Chemical</span>
            </div>
            <div style="overflow-x:auto;">
              <table class="table" style="width:100%; margin:0; border-collapse:collapse; font-size:0.875rem;">
                <thead>
                  <tr style="background:#F1F5F9; border-bottom:2px solid #E2E8F0; text-align:left;">
                    <th style="padding:10px 14px; width:60px;">No</th>
                    <th style="padding:10px 14px;">Nama Produk Chemical</th>
                    <th style="padding:10px 14px; width:110px;">Kemasan</th>
                    <th style="padding:10px 14px; width:220px;">Rasio & Standar Takaran</th>
                    <th style="padding:10px 14px;">Fungsi & Cara Penggunaan</th>
                    <th style="padding:10px 14px; width:110px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${filtered.map((item, idx) => `
                    <tr style="border-bottom:1px solid #F1F5F9; background:${idx % 2 === 0 ? '#fff' : '#FAFAFA'};">
                      <td style="padding:10px 14px; font-weight:600; color:var(--text-3);">${item.no || idx + 1}</td>
                      <td style="padding:10px 14px; font-weight:700; color:var(--text-1);">${escapeHtml(item.name)}</td>
                      <td style="padding:10px 14px;"><span class="badge" style="background:#FEF3C7; color:#92400E; font-weight:600;">${escapeHtml(item.size)}</span></td>
                      <td style="padding:10px 14px;"><span style="background:#EFF6FF; color:#1E40AF; padding:3px 8px; border-radius:6px; font-weight:600; font-size:0.82rem;">${escapeHtml(item.ratio)}</span></td>
                      <td style="padding:10px 14px; color:var(--text-2);">${escapeHtml(item.function)}</td>
                      <td style="padding:10px 14px; text-align:center;">
                        <button class="btn btn-sm btn-ghost btn-edit-chem" data-no="${item.no}" title="Edit Chemical" style="padding:4px 6px; cursor:pointer;">✏️</button>
                        <button class="btn btn-sm btn-ghost btn-delete-chem" data-no="${item.no}" title="Hapus Chemical" style="padding:4px 6px; color:#DC2626; cursor:pointer;">🗑️</button>
                      </td>
                    </tr>
                  `).join('')}
                </tbody>
              </table>
            </div>
          </div>
        `;

        // Bind Chem actions
        subContent.querySelectorAll('.btn-edit-chem').forEach(btn => {
          btn.addEventListener('click', () => {
            const no = parseInt(btn.dataset.no, 10);
            const chem = listChem.find(x => x.no === no);
            if (chem) openChemFormModal(chem);
          });
        });

        subContent.querySelectorAll('.btn-delete-chem').forEach(btn => {
          btn.addEventListener('click', () => {
            const no = parseInt(btn.dataset.no, 10);
            const chem = listChem.find(x => x.no === no);
            if (!chem) return;
            confirmDialog(`Apakah Anda yakin ingin menghapus chemical: "${chem.name}"?`, () => {
              listChem = listChem.filter(x => x.no !== no);
              // re-number
              listChem.forEach((c, i) => { c.no = i + 1; });
              saveStoredChem(listChem);
              toastSuccess('Bahan chemical berhasil dihapus');
              renderPage();
            }, 'Hapus Bahan Chemical');
          });
        });
      }
    }

    renderSubTable();

    parent.querySelector('#view-tools-btn').addEventListener('click', () => {
      toolView = 'tools';
      renderToolsTab(parent);
    });

    parent.querySelector('#view-chem-btn').addEventListener('click', () => {
      toolView = 'chemical';
      renderToolsTab(parent);
    });

    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        toolSearch = e.target.value;
        renderSubTable();
      });
    }

    parent.querySelector('#btn-add-tool-chem').addEventListener('click', () => {
      if (toolView === 'tools') openToolFormModal();
      else openChemFormModal();
    });
  }

  // MODAL FORM TAMBAH & EDIT ALAT
  function openToolFormModal(existingTool = null) {
    const isEdit = !!existingTool;
    const contentHtml = `
      <form id="tool-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Alat / Item *</label>
          <input type="text" id="form-tool-name" class="form-control" required value="${existingTool ? escapeHtml(existingTool.name) : ''}" placeholder="Contoh: Lobby Duster Set 60 cm" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Kuantitas Standar *</label>
          <input type="text" id="form-tool-qty" class="form-control" required value="${existingTool ? escapeHtml(existingTool.qty) : '1 unit'}" placeholder="Contoh: 1 set, 2 pcs, 1 paket" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Peruntukan Area & Kegunaan *</label>
          <textarea id="form-tool-area" class="form-control" rows="3" required placeholder="Contoh: Lobby utama dan koridor/lorong klinik" style="width:100%; border-radius:6px; font-size:0.85rem;">${existingTool ? escapeHtml(existingTool.area) : ''}</textarea>
        </div>
      </form>
    `;

    createModal({
      title: isEdit ? `✏️ Edit Alat: ${existingTool.name}` : `➕ Tambah Alat Kebersihan Baru`,
      content: contentHtml,
      confirmText: 'Simpan',
      onConfirm: (overlay, close) => {
        const name = overlay.querySelector('#form-tool-name').value.trim();
        const qty = overlay.querySelector('#form-tool-qty').value.trim();
        const area = overlay.querySelector('#form-tool-area').value.trim();

        if (!name || !qty || !area) {
          toastError('Harap lengkapi semua kolom');
          return;
        }

        if (isEdit) {
          existingTool.name = name;
          existingTool.qty = qty;
          existingTool.area = area;
          saveStoredTools(listTools);
          toastSuccess(`Alat "${name}" berhasil diperbarui`);
        } else {
          const nextNo = listTools.length > 0 ? Math.max(...listTools.map(x => x.no)) + 1 : 1;
          listTools.push({ no: nextNo, name, qty, area });
          saveStoredTools(listTools);
          toastSuccess(`Alat baru "${name}" berhasil ditambahkan`);
        }

        close();
        renderPage();
      }
    });
  }

  // MODAL FORM TAMBAH & EDIT CHEMICAL
  function openChemFormModal(existingChem = null) {
    const isEdit = !!existingChem;
    const contentHtml = `
      <form id="chem-edit-form" style="display:flex; flex-direction:column; gap:14px; font-size:0.88rem;">
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Nama Produk Chemical *</label>
          <input type="text" id="form-chem-name" class="form-control" required value="${existingChem ? escapeHtml(existingChem.name) : ''}" placeholder="Contoh: Floor Cleaner Prokleen Lavender" style="width:100%; border-radius:6px; height:36px;">
        </div>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Kemasan Standar *</label>
            <input type="text" id="form-chem-size" class="form-control" required value="${existingChem ? escapeHtml(existingChem.size) : '5 Liter'}" placeholder="Contoh: 5 Liter, 500 gram" style="width:100%; border-radius:6px; height:36px;">
          </div>
          <div>
            <label style="display:block; font-weight:600; margin-bottom:4px;">Rasio & Takaran Standar *</label>
            <input type="text" id="form-chem-ratio" class="form-control" required value="${existingChem ? escapeHtml(existingChem.ratio) : ''}" placeholder="Contoh: 25 ml : 1 Liter Air" style="width:100%; border-radius:6px; height:36px;">
          </div>
        </div>
        <div>
          <label style="display:block; font-weight:600; margin-bottom:4px;">Fungsi & Cara Penggunaan *</label>
          <textarea id="form-chem-function" class="form-control" rows="3" required placeholder="Contoh: Pembersihan dan pengepelan rutin seluruh lantai ruangan" style="width:100%; border-radius:6px; font-size:0.85rem;">${existingChem ? escapeHtml(existingChem.function) : ''}</textarea>
        </div>
      </form>
    `;

    createModal({
      title: isEdit ? `✏️ Edit Chemical: ${existingChem.name}` : `➕ Tambah Chemical Baru`,
      content: contentHtml,
      confirmText: 'Simpan',
      onConfirm: (overlay, close) => {
        const name = overlay.querySelector('#form-chem-name').value.trim();
        const size = overlay.querySelector('#form-chem-size').value.trim();
        const ratio = overlay.querySelector('#form-chem-ratio').value.trim();
        const func = overlay.querySelector('#form-chem-function').value.trim();

        if (!name || !size || !ratio || !func) {
          toastError('Harap lengkapi semua kolom');
          return;
        }

        if (isEdit) {
          existingChem.name = name;
          existingChem.size = size;
          existingChem.ratio = ratio;
          existingChem.function = func;
          saveStoredChem(listChem);
          toastSuccess(`Chemical "${name}" berhasil diperbarui`);
        } else {
          const nextNo = listChem.length > 0 ? Math.max(...listChem.map(x => x.no)) + 1 : 1;
          listChem.push({ no: nextNo, name, size, ratio, function: func });
          saveStoredChem(listChem);
          toastSuccess(`Chemical baru "${name}" berhasil ditambahkan`);
        }

        close();
        renderPage();
      }
    });
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 3: PEDOMAN & KETENTUAN
  // ─────────────────────────────────────────────────────────────────────────────
  function renderPolicyTab(parent) {
    parent.innerHTML = `
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(360px, 1fr)); gap:20px;">
        <!-- Card 1: Grooming -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#EFF6FF; padding:8px; border-radius:8px;">👔</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Standar Grooming Petugas (6.1)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">8 Persyaratan Wajib Penampilan Facility Care</span>
            </div>
          </div>
          <ol style="margin:0 0 0 18px; padding:0; font-size:0.88rem; color:var(--text-2); line-height:1.8;">
            <li>Menggunakan seragam resmi setiap hari (Senin s/d Minggu).</li>
            <li>Menggunakan celana bahan berwarna hitam (bukan jeans/chino kasual).</li>
            <li>Wajib mengenakan name badge atau ID Card di dada kiri.</li>
            <li>Menggunakan kaos kaki warna hitam.</li>
            <li>Menggunakan sepatu pantofel hitam yang bersih dan disemir.</li>
            <li>Menggunakan parfum dan deodorant untuk menjaga kesegaran badan.</li>
            <li>Merapikan rambut (pria pendek rapi, wanita diikat/hairnet rapi).</li>
            <li>Dilarang menggunakan aksesoris perhiasan selain jam tangan standar.</li>
          </ol>
        </div>

        <!-- Card 2: Jam Kerja -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#FEF3C7; padding:8px; border-radius:8px;">⏰</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Pengaturan Jam Kerja & Shift (6.2)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Sistem Kerja 6 Hari Kerja : 1 Hari Libur</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:10px; border-left:3px solid #2563EB;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">🏥 Klinik Primer (1 Orang Facility Care):</strong>
              <span>Shift 1 (Normal): <strong>08.00 s/d 17.00</strong> (Istirahat 13.00 s/d 14.00)</span>
            </div>
            <div style="background:#F8FAFC; border-radius:8px; padding:10px 14px; margin-bottom:12px; border-left:3px solid #059669;">
              <strong style="color:var(--text-1); display:block; margin-bottom:2px;">✨ Klinik Luxury (> 1 Orang Facility Care):</strong>
              <div style="margin-bottom:2px;">• Shift 1 (Pagi): <strong>07.00 s/d 16.00</strong> (Istirahat 12.00 s/d 13.00)</div>
              <div>• Shift 2 (Siang): <strong>13.00 s/d 22.00</strong> (Istirahat 17.00 s/d 18.00)</div>
            </div>
            <p style="margin:0; font-size:0.82rem; color:var(--text-3);">
              <strong>Ketentuan Backup:</strong> Saat Facility Care cabang libur, tim cabang wajib melakukan backup pekerjaan yang ditentukan oleh Supervisor (SPV) Cabang.
            </p>
          </div>
        </div>

        <!-- Card 3: Tanggung Jawab & Pembagian Area -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#ECFDF5; padding:8px; border-radius:8px;">🏢</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Tanggung Jawab & Area (6.3)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Pembagian Pelaksana & Penanggung Jawab Area</span>
            </div>
          </div>
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem; margin-bottom:14px;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; text-align:left;">
                <th style="padding:6px 10px;">Area Kerja</th>
                <th style="padding:6px 10px;">Pelaksana</th>
                <th style="padding:6px 10px;">PIC Pengawas</th>
              </tr>
            </thead>
            <tbody>
              <tr style="border-bottom:1px solid #F1F5F9;">
                <td style="padding:8px 10px; font-weight:600;">HQ & Gudang Pusat</td>
                <td style="padding:8px 10px;">Facility Care Pusat</td>
                <td style="padding:8px 10px; color:#2563EB; font-weight:600;">FC Supervisor</td>
              </tr>
              <tr>
                <td style="padding:8px 10px; font-weight:600;">Klinik / Cabang</td>
                <td style="padding:8px 10px;">Facility Care Cabang</td>
                <td style="padding:8px 10px; color:#059669; font-weight:600;">SPV Cabang</td>
              </tr>
            </tbody>
          </table>
          <div style="font-size:0.83rem; color:var(--text-2); line-height:1.6;">
            <div style="margin-bottom:6px;"><strong>Deep Cleaning:</strong> Pembersihan total eventual pasca renovasi, pembukaan cabang baru, atau relokasi cabang lama.</div>
            <div><strong>General Cleaning (GC):</strong> Pembersihan menyeluruh berkala untuk cabang eksisting sesuai jadwal yang telah ditetapkan.</div>
          </div>
        </div>

        <!-- Card 4: Monitoring & Evaluasi -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
          <div style="display:flex; align-items:center; gap:10px; margin-bottom:14px; border-bottom:1px solid #F1F5F9; padding-bottom:10px;">
            <span style="font-size:1.5rem; background:#F5F3FF; padding:8px; border-radius:8px;">🔍</span>
            <div>
              <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Sistem Monitoring & Inspeksi (6.4)</h3>
              <span style="font-size:0.8rem; color:var(--text-3);">Mekanisme Evaluasi Kinerja Harian & Triwulanan</span>
            </div>
          </div>
          <div style="font-size:0.88rem; color:var(--text-2); line-height:1.7;">
            <div style="margin-bottom:10px;">
              <strong style="color:var(--text-1);">1. Monitor Harian & Mingguan:</strong>
              <div style="font-size:0.84rem; color:var(--text-2); margin-top:2px;">
                Ceklist harian diisi 2 kali sehari (Pagi & Siang) dan pekerjaan mingguan diisi seminggu sekali. Wajib ditandatangani oleh SPV Cabang / FC Supervisor.
              </div>
            </div>
            <div>
              <strong style="color:var(--text-1);">2. Inspeksi Cabang Triwulanan (3 Bulan):</strong>
              <div style="font-size:0.84rem; color:var(--text-2); margin-top:2px;">
                Dilakukan oleh FC Specialist langsung ke lokasi klinik. Laporan inspeksi diunggah ke Basecamp dan wajib diberikan tindak lanjut oleh PIC cabang maksimal <strong>3 hari</strong> setelah laporan terbit. Hasil penilaian mempengaruhi performance cabang.
              </div>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  // ─────────────────────────────────────────────────────────────────────────────
  // TAB 4: FORM & TAUTAN TERKAIT
  // ─────────────────────────────────────────────────────────────────────────────
  function renderLinksTab(parent) {
    parent.innerHTML = `
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:18px;">
        <!-- Card: Ceklist Harian Google Sheets -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">📊</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Ceklist Harian & Mingguan</h3>
                <span class="badge" style="background:#E0F2FE; color:#0369A1; font-size:0.75rem; font-weight:700;">Google Sheets</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Format lembar ceklist kebersihan harian (Shift Pagi & Siang) dan ceklist pembersihan berkala mingguan yang wajib diisi oleh Facility Care dan disetujui Supervisor Cabang.
            </p>
          </div>
          <a href="https://docs.google.com/spreadsheets/d/1alq6ppT7syB8Q0_RjjeqwYhjeHgrWvVe-KJ3dHXWrss/edit?usp=sharing" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px;">
            📄 Buka Google Sheets Ceklist ➔
          </a>
        </div>

        <!-- Card: Form Inspeksi Hygiene Google Forms -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">📑</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Form Inspeksi Hygiene Cabang</h3>
                <span class="badge" style="background:#F3E8FF; color:#7E22CE; font-size:0.75rem; font-weight:700;">Google Forms</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              Formulir digital inspeksi evaluasi kebersihan dan kepatuhan hygiene cabang yang diisi oleh Facility Care Specialist setiap periode 3 bulanan.
            </p>
          </div>
          <a href="https://docs.google.com/forms/d/e/1FAIpQLScn8iPabIt51E_GORkbD86IURSS-EsUYINEQ_Zwdq_uC3XOCQ/viewform" target="_blank" rel="noopener" class="btn btn-primary" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#7E22CE; border-color:#7E22CE;">
            📋 Buka Google Form Inspeksi ➔
          </a>
        </div>

        <!-- Card: Basecamp Monitoring Project -->
        <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; flex-direction:column; justify-content:space-between;">
          <div>
            <div style="display:flex; align-items:center; gap:10px; margin-bottom:12px;">
              <span style="font-size:2rem;">🏕️</span>
              <div>
                <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:var(--text-1);">Laporan Inspeksi Basecamp</h3>
                <span class="badge" style="background:#FEF3C7; color:#92400E; font-size:0.75rem; font-weight:700;">Basecamp 3 Project</span>
              </div>
            </div>
            <p style="font-size:0.85rem; color:var(--text-2); line-height:1.6; margin-bottom:16px;">
              To-do list dan kanal pelaporan temuan inspeksi cabang. PIC Cabang wajib memberikan respon tindak lanjut maksimal 3 hari kerja setelah laporan terbit.
            </p>
          </div>
          <a href="https://3.basecamp.com/5259614/buckets/33617501/todosets/6367925585" target="_blank" rel="noopener" class="btn btn-outline" style="display:inline-flex; align-items:center; justify-content:center; gap:6px; font-weight:600; padding:10px; border-radius:8px; background:#fff;">
            🚀 Buka Laporan di Basecamp ➔
          </a>
        </div>
      </div>

      <!-- Validasi & Pengesahan SOP -->
      <div class="card" style="margin-top:20px; padding:20px; background:#fff; border:1px solid var(--border); border-radius:12px;">
        <h3 style="margin:0 0 14px 0; font-size:1rem; font-weight:700; color:var(--text-1); display:flex; align-items:center; gap:8px;">
          <span>📜</span> Validasi & Pengesahan Dokumen SOP
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:16px;">
          <div style="border-left:3px solid #2563EB; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disiapkan Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Berlin Ariansyah</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Facility Care Supervisor</div>
          </div>
          <div style="border-left:3px solid #059669; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Diketahui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Catur Meiyogo</div>
            <div style="font-size:0.8rem; color:var(--text-2);">General Support Manager</div>
          </div>
          <div style="border-left:3px solid #7C3AED; padding-left:12px;">
            <div style="font-size:0.75rem; color:var(--text-3); font-weight:700; text-transform:uppercase;">Disetujui Oleh:</div>
            <div style="font-size:0.95rem; font-weight:700; color:var(--text-1); margin-top:2px;">Drg. Yulita</div>
            <div style="font-size:0.8rem; color:var(--text-2);">Chief Executive Officer</div>
          </div>
        </div>
      </div>
    `;
  }

  // EXPORT EXCEL DATA
  async function exportAllData() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const ikSheet = listIK.map(ik => ({
        'No IK': `IK-${ik.id}`,
        'Judul Pekerjaan': ik.title,
        'Kategori': ik.category,
        'Sasaran Target Area': ik.target,
        'Rasio Chemical': ik.chemical_ratio || '-',
        'Alat & Bahan': ik.tools.join(', '),
        'Langkah Kerja': ik.steps.map((s, i) => `${i + 1}. ${s}`).join(' | '),
        'Catatan Safety': ik.notes || '-'
      }));

      const dateStr = new Date().toISOString().slice(0, 10);
      downloadExcel(ikSheet, `SOP_Facility_Care_Instruksi_Kerja_${dateStr}`);
      toastSuccess('Data Excel Instruksi Kerja berhasil diunduh');
    } catch (e) {
      console.error(e);
      toastError('Gagal mendownload data Excel: ' + e.message);
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  renderPage();
}
