/* ============================================================
   SIPEKA — Backend Google Apps Script
   Dinas Pendidikan Kab. Deli Serdang
   ------------------------------------------------------------
   CARA PAKAI (lihat README.md):
   1. Buat Google Spreadsheet baru → Extensions → Apps Script
   2. Tempel seluruh file ini ke Code.gs
   3. Jalankan fungsi setup() sekali (beri izin)
   4. Deploy → New deployment → Web app:
      - Execute as: Me
      - Who has access: Anyone
   5. Salin URL /exec ke js/config.js (API_URL)
   Password admin awal: admin123  → GANTI lewat halaman Admin.
   ============================================================ */

const NAMA_SHEET = {
  SEKOLAH: 'MASTER_SEKOLAH',
  GURU_M: 'MASTER_GURU',
  SDSMP: 'KS_SD_SMP',
  PAUD: 'KS_PAUD',
  GURU: 'GURU',
  SETTING: 'PENGATURAN',
  A_PENGAWAS: 'AKUN_PENGAWAS', // username = NIP pengawas
  A_KS: 'AKUN_KS',             // username = NPSN sekolah
  A_KEC: 'AKUN_KECAMATAN'      // username = nama kecamatan
};

/* ---------- Definisi indikator (identik dgn js/indikator.js) ---------- */
const KS_SDSMP = [
  ['Performance Sekolah', [
    ['1.1','Kebersihan Lingkungan Satuan Pendidikan'],
    ['1.2','Kebersihan Lingkungan Belajar'],
    ['1.3','Kelengkapan Sanitasi dan Kebersihan Kamar Mandi'],
    ['1.4','Kerindangan dan Apotik Hidup Sekolah'],
    ['1.5','Mewujudkan Sekolah Adiwiyata'],
    ['1.6','Mewujudkan sekolah dan kantin Sehat']]],
  ['Tata Kelola Sampah', [
    ['2.1','Ketersediaan Tong Sampah Sesuai Peruntukan'],
    ['2.2','Pengelolaan bank Sampah'],
    ['2.3','Tempat Pengomposan dan Pengolahan Sampah Organik'],
    ['2.4','Produk Daur Ulang']]],
  ['Inovasi Sekolah', [
    ['3.1','Jumlah Inovasi di Sekolah'],
    ['3.2','Pemanfaatan Teknologi dalam Pembelajaran'],
    ['3.3','Upaya meningkatkan literasi dan numerasi']]],
  ['Indeks Kepuasan & Kepemimpinan', [
    ['4.1','Indeks kepuasan Orang Tua Murid'],
    ['4.2','Indeks kepuasan Siswa'],
    ['4.3','Indeks Kepuasan guru dan staff']]],
  ['Prestasi Murid', [
    ['5.1','Jumlah murid dalam mengikuti perlombaan di bidang akademik dan non akademik'],
    ['5.2','Jumlah Murid yang Berprestasi di bidang akademik'],
    ['5.3','Jumlah Murid yang Berprestasi di bidang non akademik'],
    ['5.4','Jumlah Murid yang tidak Lancar Calistung'],
    ['5.5','Partisipasi Murid dalam Ekstrakurikuler']]],
  ['Kelengkapan Dokumen Sekolah', [
    ['6.1','Dokumen Rencana Kerja Sekolah'],
    ['6.2','Dokumen Kurikulum Satuan Pendidikan'],
    ['6.3','Dokumen Inventaris Barang'],
    ['6.4','Dokumen Administrasi Murid'],
    ['6.5','Laporan Penggunaan Dana BOS']]]
];
const KS_PAUD = [
  ['Performance', [
    ['1.1','Kebersihan Lingkungan Satuan Pendidikan'],
    ['1.2','Kebersihan Lingkungan Belajar'],
    ['1.3','Kelengkapan Sanitasi dan Kebersihan Kamar Mandi'],
    ['1.4','Kerindangan dan Apotik Hidup Sekolah']]],
  ['Kebijakan Sekolah Ramah Anak (SRA)', [
    ['2.1','Area bermain aman bagi anak'],
    ['2.2','Terpisah toilet laki-laki dan perempuan'],
    ['2.3','Parenting']]],
  ['Pemanfaatan Teknologi & 7 KAIH', [
    ['3.1','Pemanfaatan Teknologi dan tersedianya APE dalam pembelajaran'],
    ['3.2','Pelaksanaan kegiatan 7 KAIH']]],
  ['Data Tumbuh Kembang Anak', [
    ['4.1','Pertumbuhan dan perkembangan anak']]],
  ['Kemitraan', [
    ['5.1','Kerjasama dengan puskesmas, posyandu, atau lembaga lain'],
    ['5.2','Keterlibatan orang tua dalam kegiatan sekolah']]],
  ['Dokumen Sekolah', [
    ['6.1','Dokumen Rencana Kerja Sekolah'],
    ['6.2','Dokumen Kurikulum Satuan Pendidikan'],
    ['6.3','Dokumen Inventaris Barang'],
    ['6.4','Dokumen Administrasi Murid'],
    ['6.5','Laporan Penggunaan Dana BOS']]]
];
const GURU_GRP = [
  ['Manajemen Ruang & Lingkungan Sekolah', 7],
  ['Tata Kelola Sampah', 4],
  ['Inovasi & Pengembangan Profesi', 8],
  ['Kepuasan, Komunikasi & Prestasi Siswa', 10],
  ['Praktek Pembelajaran-Perangkat Pembelajaran', 4]
];

const PENGATURAN_DEFAULT = {
  kategoriGuru: [{min:90,label:'Sangat Baik'},{min:76,label:'Baik'},{min:61,label:'Cukup'},{min:0,label:'Kurang'}],
  kategoriKS:   [{min:4.5,label:'Sangat Baik'},{min:3.5,label:'Baik'},{min:2.5,label:'Cukup'},{min:0,label:'Kurang'}],
  kelompokBaseline: [
    {min:1,max:1.9,label:'Kelompok A (1–1,9)'},
    {min:2,max:2.9,label:'Kelompok B (2–2,9)'},
    {min:3,max:3.9,label:'Kelompok C (3–3,9)'},
    {min:4,max:5,  label:'Kelompok D (4–5)'}
  ]
};

/* ================= SETUP DATABASE ================= */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buatSheet_(ss, NAMA_SHEET.SEKOLAH, ['NPSN','NAMA SEKOLAH','JENJANG','KECAMATAN','NAMA KEPALA SEKOLAH','HP KEPALA SEKOLAH','NAMA PENGAWAS','HP PENGAWAS','NIP KEPALA SEKOLAH']); // NIP pengawas TIDAK di master — diambil dari AKUN_PENGAWAS
  buatSheet_(ss, NAMA_SHEET.GURU_M, ['NPSN','NAMA GURU','NIP/NUPTK','STATUS','PANGKAT/GOL.RUANG','MATA PELAJARAN','HP GURU']);
  buatSheet_(ss, NAMA_SHEET.A_PENGAWAS, ['NIP (USERNAME)','NAMA PENGAWAS','HP','HASH PASSWORD','KUNCI','WAJIB GANTI','TERAKHIR LOGIN']);
  buatSheet_(ss, NAMA_SHEET.A_KS, ['NPSN (USERNAME)','NAMA SEKOLAH','NAMA KEPALA SEKOLAH','HASH PASSWORD','KUNCI','WAJIB GANTI','TERAKHIR LOGIN']);
  buatSheet_(ss, NAMA_SHEET.A_KEC, ['KECAMATAN (USERNAME)','NAMA PETUGAS','HP','HASH PASSWORD','KUNCI','WAJIB GANTI','TERAKHIR LOGIN']); // struktur kolom = AKUN_PENGAWAS

  // KS SD & SMP
  let h = ['Timestamp','Email Address','NPSN','Kecamatan','JENJANG SEKOLAH','NAMA SEKOLAH','NAMA PENGAWAS','Nomor HP Pengawas','NAMA KEPALA SEKOLAH','Nomor HP Kepala Sekolah','Periode Penilaian','Bulan','Minggu'];
  KS_SDSMP.forEach(g => g[1].forEach(i => {
    h.push(i[0]+' '+i[1]+' [Baseline]', i[0]+' '+i[1]+' [Target]', i[0]+' '+i[1]+' [Capaian]', 'Link Dokumentasi '+i[0]+' '+i[1]);
  }));
  h.push('Catatan dan Rekomendasi','Nomor Urut','Rata Baseline','Rata Target','Rata Capaian','Kinerja Keseluruhan');
  KS_SDSMP.forEach((g, i) => h.push('BASELINE '+(i+1),'TARGET '+(i+1),'CAPAIAN '+(i+1),'Kinerja '+g[0]));
  h.push('ID KIRIM'); // anti-duplikat kiriman ulang (kolom 148)
  buatSheet_(ss, NAMA_SHEET.SDSMP, h);

  // KS PAUD
  h = ['Timestamp','Email Address','NPSN','Kecamatan','JENJANG SEKOLAH','NAMA SEKOLAH','NAMA PENGAWAS','Nomor HP Pengawas','NAMA KEPALA SEKOLAH','Nomor HP Kepala Sekolah','Bulan','Minggu'];
  KS_PAUD.forEach(g => g[1].forEach(i => {
    h.push(i[0]+' '+i[1]+' [Baseline]', i[0]+' '+i[1]+' [Target]', i[0]+' '+i[1]+' [Capaian]', 'Link Dokumentasi '+i[0]+' '+i[1]);
  }));
  h.push('Catatan dan Rekomendasi','Rata Baseline','Rata Target','Rata Capaian','Kinerja Keseluruhan');
  KS_PAUD.forEach(g => h.push('Rata Baseline '+g[0],'Rata Target '+g[0],'Rata Capaian '+g[0],'Kinerja '+g[0]));
  h.push('ID KIRIM'); // anti-duplikat (kolom 110)
  buatSheet_(ss, NAMA_SHEET.PAUD, h);

  // GURU
  h = ['Timestamp','Email Address','KECAMATAN','JENJANG','NPSN','SATUAN PENDIDIKAN','NAMA KEPALA SEKOLAH','NOMOR HP KEPALA SEKOLAH','NAMA GURU','STATUS','PANGKAT/GOL.RUANG','NIP/NUPTK','MATA PELAJARAN','NOMOR HP GURU','BULAN'];
  GURU_GRP.forEach((g, gi) => {
    for (let b = 1; b <= g[1]; b++) h.push('Penilaian '+g[0]+' [Butir '+b+']');
    h.push('Link Bukti Dukung Indikator '+(gi+1)+'. '+g[0]);
  });
  GURU_GRP.forEach(g => h.push('Rata Capaian '+g[0]));
  h.push('Rata Keseluruhan');
  h.push('ID KIRIM'); // anti-duplikat (kolom 60)
  buatSheet_(ss, NAMA_SHEET.GURU, h);

  // PENGATURAN
  const st = buatSheet_(ss, NAMA_SHEET.SETTING, ['KUNCI','NILAI']);
  if (st.getLastRow() < 2) {
    st.getRange(2, 1, 3, 2).setValues([
      ['adminHash', hash_('admin123')],
      ['pengaturan', JSON.stringify(PENGATURAN_DEFAULT)],
      ['apiToken', Utilities.getUuid()] // token anti-spam form publik (opsional)
    ]);
  }
  SpreadsheetApp.flush();
  Logger.log('Setup selesai. Password admin awal: admin123');
}
function buatSheet_(ss, nama, header) {
  let sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ================= HTTP HANDLER ================= */
function doGet(e) {
  try {
    const a = ((e && e.parameter && e.parameter.action) || '').toLowerCase();
    // URL dibuka tanpa parameter (mis. dicek di address bar) → tampilkan status, bukan error
    if (!a) {
      return json_({
        ok: true,
        aplikasi: 'SIPEKA — Penilaian Kinerja Kepala Sekolah & Guru',
        instansi: 'Dinas Pendidikan Kabupaten Deli Serdang',
        status: 'Backend AKTIF dan siap menerima permintaan ✅',
        waktuServer: new Date().toISOString(),
        catatan: 'URL ini adalah API untuk website SIPEKA, bukan halaman yang dibuka langsung. ' +
                 'Salin URL ini ke API_URL pada js/config.js. Untuk uji cepat, buka: ' +
                 '...?action=master (daftar sekolah) atau ...?action=pengaturan'
      });
    }
    if (a === 'master')      return json_(getMaster_());
    if (a === 'gurusekolah') return json_(getGuruSekolah_(e.parameter.npsn));
    if (a === 'pengaturan')  return json_({ pengaturan: bacaPengaturan_() });
    if (a === 'rapor')       return json_(getRapor_(e.parameter));
    if (a === 'raporbanyak') return json_(getRaporBanyak_(e.parameter));
    if (a === 'hasilkecamatan') {
      if (e.parameter.token) {
        cekToken_(e.parameter.token);
        if (e.parameter.kecamatan) return json_(getHasilKecamatan_(e.parameter.kecamatan));
        return json_({ daftarKecamatan: daftarKecamatan_() });
      }
      cekAuthUser_({ peran: 'kecamatan', username: e.parameter.username, kunci: e.parameter.kunci });
      return json_(getHasilKecamatan_(e.parameter.username));
    }
    if (a === 'laporan')     { cekToken_(e.parameter.token); return json_(getLaporan_(false)); }
    if (a === 'tarik')       { cekToken_(e.parameter.token); return json_(getLaporan_(true)); }
    return json_({ error: 'action "' + a + '" tidak dikenal. Action GET yang tersedia: master, gurusekolah, pengaturan, rapor, raporbanyak, hasilkecamatan, laporan, tarik' });
  } catch (err) { return json_({ error: String(err.message || err) }); }
}
function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    const a = (b.action || '').toLowerCase();
    if (a === 'login')          return json_(login_(b.password));
    if (a === 'gantipassword')  { cekToken_(b.token); return json_(gantiPassword_(b.passwordBaru)); }
    if (a === 'savepengaturan') { cekToken_(b.token); return json_(savePengaturan_(b.pengaturan)); }
    if (a === 'loginuser')        return json_(loginUser_(b.peran, b.username, b.password));
    if (a === 'gantipassworduser')return json_(gantiPwUser_(b.auth, b.passwordBaru));
    if (a === 'admingenerateakun'){ cekToken_(b.token); return json_(generateAkun_(b.peran, b.passwordAwal)); }
    if (a === 'adminresetpassword'){ cekToken_(b.token); return json_(resetPwUser_(b.peran, b.username)); }
    if (a === 'admindaftarakun')  { cekToken_(b.token); return json_(daftarAkun_(b.peran)); }
    if (a === 'editks')         return json_(editKS_(b, b.jenis, b.baris, b.nilai, b.catatan));
    if (a === 'editguru')       return json_(editGuru_(b, b.baris, b.nilai, b.links));
    if (a === 'deleteks')       return json_(deleteKS_(b, b.jenis, b.baris));
    if (a === 'deleteguru')     return json_(deleteGuru_(b, b.baris));
    if (a === 'submitks')       { cekAuthSubmit_(b.auth, 'pengawas', b.data); return json_(submitKS_(b.data, false)); }
    if (a === 'submitpaud')     { cekAuthSubmit_(b.auth, 'pengawas', b.data); return json_(submitKS_(b.data, true)); }
    if (a === 'submitguru')     { cekAuthSubmit_(b.auth, 'ks', b.data); return json_(submitGuru_(b.data)); }
    return json_({ error: 'action tidak dikenal' });
  } catch (err) { return json_({ error: String(err.message || err) }); }
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= AUTH ADMIN ================= */
function hash_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}
function bacaSetting_(kunci) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMA_SHEET.SETTING);
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
  for (const r of v) if (r[0] === kunci) return r[1];
  return null;
}
function tulisSetting_(kunci, nilai) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMA_SHEET.SETTING);
  const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === kunci) { sh.getRange(i + 2, 2).setValue(nilai); return; }
  sh.appendRow([kunci, nilai]);
}
function login_(password) {
  if (hash_(String(password || '')) !== bacaSetting_('adminHash')) throw new Error('Password salah');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + token, '1', 21600); // 6 jam
  return { token: token };
}
function cekToken_(token) {
  if (!token || !CacheService.getScriptCache().get('tok_' + token)) throw new Error('Sesi admin tidak valid / kedaluwarsa. Silakan login ulang.');
}
function gantiPassword_(baru) {
  if (!baru || String(baru).length < 6) throw new Error('Password minimal 6 karakter');
  tulisSetting_('adminHash', hash_(String(baru)));
  return { ok: true };
}
function bacaPengaturan_() {
  try { return JSON.parse(bacaSetting_('pengaturan')) || PENGATURAN_DEFAULT; }
  catch (e) { return PENGATURAN_DEFAULT; }
}
function savePengaturan_(p) { tulisSetting_('pengaturan', JSON.stringify(p)); return { ok: true }; }

/* ================= AKUN PENGAWAS & KEPALA SEKOLAH =================
   AKUN_PENGAWAS: NIP(username) | NAMA | HP | HASH | KUNCI | WAJIB GANTI | TERAKHIR LOGIN
   AKUN_KS      : NPSN(username)| SEKOLAH | NAMA KS | HASH | KUNCI | WAJIB GANTI | TERAKHIR LOGIN */
function sheetAkun_(peran) {
  const peta = { pengawas: NAMA_SHEET.A_PENGAWAS, ks: NAMA_SHEET.A_KS, kecamatan: NAMA_SHEET.A_KEC };
  if (!peta[peran]) throw new Error('Peran tidak dikenal');
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(peta[peran]);
}
function norm_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ''); }
function cariAkun_(peran, username) {
  const sh = sheetAkun_(peran);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  for (let i = 0; i < v.length; i++) {
    const u = String(v[i][0]).replace(/^'/, '').trim();
    const cocok = peran === 'kecamatan' ? norm_(u) === norm_(username) : u === String(username || '').trim();
    if (cocok && v[i][0] !== '') return { baris: i + 2, data: v[i], sh: sh };
  }
  return null;
}
function samaNama_(a, b) { return String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase(); }

function loginUser_(peran, username, password) {
  const a = cariAkun_(peran, username);
  if (!a) throw new Error('Akun tidak ditemukan. Periksa username (NIP/NPSN) atau hubungi admin.');
  if (hash_(String(password || '')) !== a.data[3]) throw new Error('Password salah');
  let kunci = a.data[4];
  if (!kunci) { kunci = Utilities.getUuid(); a.sh.getRange(a.baris, 5).setValue(kunci); }
  a.sh.getRange(a.baris, 7).setValue(new Date());
  const hasil = { kunci: kunci, peran: peran, username: String(a.data[0]).replace(/^'/, ''), wajibGanti: String(a.data[5]).toUpperCase() === 'Y' };
  if (peran === 'pengawas') {
    hasil.nama = a.data[1]; hasil.hp = String(a.data[2] || '');
    hasil.binaan = getMaster_().sekolah.filter(s => samaNama_(s.pengawas, a.data[1]));
  } else if (peran === 'kecamatan') {
    hasil.kecamatan = String(a.data[0]).replace(/^'/, '');
    hasil.nama = a.data[1] || ('Admin Kec. ' + hasil.kecamatan);
  } else {
    const m = getMaster_().sekolah.find(s => String(s.npsn) === hasil.username);
    hasil.nama = a.data[2]; hasil.profil = m || { npsn: hasil.username, sekolah: a.data[1], ks: a.data[2] };
  }
  return hasil;
}
function cekAuthUser_(auth) {
  if (!auth || !auth.peran || !auth.username || !auth.kunci) throw new Error('Harus login untuk mengirim penilaian');
  const a = cariAkun_(auth.peran, auth.username);
  if (!a || !a.data[4] || String(a.data[4]) !== String(auth.kunci)) throw new Error('Sesi tidak valid — silakan login ulang');
  return a;
}
function gantiPwUser_(auth, baru) {
  if (!baru || String(baru).length < 6) throw new Error('Password minimal 6 karakter');
  const a = cekAuthUser_(auth);
  a.sh.getRange(a.baris, 4).setValue(hash_(String(baru)));
  a.sh.getRange(a.baris, 6).setValue('T');
  return { ok: true };
}
/* pembatasan kirim: pengawas hanya sekolah binaan, KS hanya sekolahnya */
function cekAuthSubmit_(auth, peranWajib, data) {
  const a = cekAuthUser_(auth);
  if (auth.peran !== peranWajib) throw new Error('Akun Anda tidak berhak mengisi form ini');
  const npsn = String((data && data.ident && data.ident.npsn) || '').trim();
  if (peranWajib === 'pengawas') {
    const m = getMaster_().sekolah.find(s => String(s.npsn) === npsn);
    if (m && !samaNama_(m.pengawas, a.data[1]))
      throw new Error('Sekolah ' + npsn + ' bukan binaan Anda');
    // sekolah di luar master (ketik manual) diizinkan, tercatat atas akun pengawas ybs
  } else {
    if (npsn !== String(a.data[0]).replace(/^'/, '').trim())
      throw new Error('Anda hanya dapat menilai guru di satuan pendidikan Anda sendiri (NPSN ' + String(a.data[0]).replace(/^'/, '') + ')');
  }
}
/* admin: generate massal akun yang belum ada, password awal seragam */
function generateAkun_(peran, pwAwal) {
  if (!pwAwal || String(pwAwal).length < 6) throw new Error('Password awal minimal 6 karakter');
  const sh = sheetAkun_(peran);
  const ada = new Set();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .forEach(r => ada.add(
      peran === 'pengawas' ? String(r[1]).trim().toLowerCase() :
      peran === 'kecamatan' ? norm_(r[0]) :
      String(r[0]).replace(/^'/, '').trim()));
  const master = getMaster_().sekolah;
  const hashAwal = hash_(String(pwAwal));
  let baru = 0;
  if (peran === 'pengawas') {
    const unik = {};
    master.forEach(s => { if (s.pengawas && !unik[s.pengawas.trim().toLowerCase()]) unik[s.pengawas.trim().toLowerCase()] = s; });
    Object.values(unik).forEach(s => {
      if (!ada.has(s.pengawas.trim().toLowerCase())) {
        sh.appendRow(['', s.pengawas, "'" + (s.hpPengawas || ''), hashAwal, Utilities.getUuid(), 'Y', '']);
        baru++;
      }
    });
  } else if (peran === 'kecamatan') {
    const unik = {};
    master.forEach(s => { if (s.kecamatan) unik[norm_(s.kecamatan)] = s.kecamatan; });
    Object.keys(unik).forEach(k => {
      if (!ada.has(k)) {
        sh.appendRow([unik[k], '', '', hashAwal, Utilities.getUuid(), 'Y', '']);
        baru++;
      }
    });
  } else {
    master.forEach(s => {
      if (!ada.has(String(s.npsn).trim())) {
        sh.appendRow(["'" + s.npsn, s.sekolah, s.ks, hashAwal, Utilities.getUuid(), 'Y', '']);
        baru++;
      }
    });
  }
  tulisSetting_('pwAwalHash_' + peran, hashAwal);
  return { ok: true, baru: baru, catatan:
    peran === 'pengawas' ? 'Isi kolom NIP (USERNAME) di sheet AKUN_PENGAWAS — pengawas baru bisa login setelah NIP terisi.' :
    peran === 'kecamatan' ? 'Username = nama kecamatan (boleh diketik tanpa spasi saat login). Isi NAMA PETUGAS di sheet AKUN_KECAMATAN bila perlu.' : '' };
}
/* admin: reset password akun ke password awal seragam */
function resetPwUser_(peran, username) {
  const hashAwal = bacaSetting_('pwAwalHash_' + peran);
  if (!hashAwal) throw new Error('Password awal belum ditetapkan — jalankan Generate Akun dahulu');
  const a = cariAkun_(peran, username);
  if (!a) throw new Error('Akun "' + username + '" tidak ditemukan');
  a.sh.getRange(a.baris, 4).setValue(hashAwal);
  a.sh.getRange(a.baris, 6).setValue('Y');
  return { ok: true, nama: peran === 'pengawas' ? a.data[1] : a.data[2] };
}
function daftarAkun_(peran) {
  const sh = sheetAkun_(peran);
  if (!sh || sh.getLastRow() < 2) return { akun: [] };
  return { akun: sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().map(r => ({
    username: String(r[0]).replace(/^'/, ''), nama: peran === 'pengawas' ? r[1] : r[2], ket: peran === 'pengawas' ? '' : r[1]
  })) };
}

/* ================= MASTER DATA ================= */
function bacaSemua_(nama) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function getMaster_() {
  const rows = bacaSemua_(NAMA_SHEET.SEKOLAH);
  // NIP pengawas diambil dari AKUN_PENGAWAS (sumber tunggal, anti-konflik), dipetakan via nama
  const nipPengawas = {};
  bacaSemua_(NAMA_SHEET.A_PENGAWAS).forEach(r => {
    if (r[1]) nipPengawas[String(r[1]).trim().toLowerCase()] = String(r[0]).replace(/^'/, '');
  });
  return { sekolah: rows.filter(r => r[0]).map(r => ({
    npsn: String(r[0]), sekolah: r[1], jenjang: r[2], kecamatan: r[3],
    ks: r[4], hpKS: String(r[5] || ''), pengawas: r[6], hpPengawas: String(r[7] || ''),
    nipKS: String(r[8] || ''),
    nipPengawas: nipPengawas[String(r[6] || '').trim().toLowerCase()] || ''
  })) };
}
function getGuruSekolah_(npsn) {
  const rows = bacaSemua_(NAMA_SHEET.GURU_M);
  return { guru: rows.filter(r => String(r[0]) === String(npsn)).map(r => ({
    nama: r[1], nip: String(r[2] || ''), status: r[3], pangkat: r[4], mapel: r[5], hp: String(r[6] || '')
  })) };
}

/* ================= SINKRON MASTER_SEKOLAH =================
   - NPSN ditemukan  → perbarui baris (hanya kolom terisi yang ditimpa)
   - NPSN tidak ada  → hasil "ketik manual": TAMBAHKAN sekolah baru ke master
     (NPSN, nama sekolah, jenjang, kecamatan, nama+HP KS, nama+HP pengawas, NIP KS) */
function perbaruiMaster_(id) {
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMA_SHEET.SEKOLAH);
    if (!sh) return;
    const npsn = String(id.npsn || '').trim();
    if (!npsn) return;
    const n = sh.getLastRow();
    const v = n > 1 ? sh.getRange(2, 1, n - 1, 1).getValues() : [];
    for (let i = 0; i < v.length; i++) {
      if (String(v[i][0]).replace(/^'/, '').trim() === npsn) {
        const baris = i + 2;
        // kolom: 2 NAMA SEKOLAH, 3 JENJANG, 4 KECAMATAN, 5 NAMA KS, 6 HP KS,
        //        7 NAMA PENGAWAS, 8 HP PENGAWAS, 9 NIP KS (NIP pengawas TIDAK di master)
        const isi = [
          [2, id.sekolah], [3, id.jenjang], [4, id.kecamatan], [5, id.ks],
          [6, id.hpKS, true], [7, id.pengawas], [8, id.hpPengawas, true],
          [9, id.nipKS, true]
        ];
        isi.forEach(([kol, val, teks]) => {
          if (val != null && String(val).trim() !== '')
            sh.getRange(baris, kol).setValue(teks ? "'" + String(val).trim() : String(val).trim());
        });
        return;
      }
    }
    // tidak ditemukan → sekolah baru dari mode ketik manual: daftarkan ke master
    sh.appendRow([
      "'" + npsn,
      String(id.sekolah || '').trim(), String(id.jenjang || '').trim(), String(id.kecamatan || '').trim(),
      String(id.ks || '').trim(), "'" + String(id.hpKS || '').trim(),
      String(id.pengawas || '').trim(), "'" + String(id.hpPengawas || '').trim(),
      "'" + String(id.nipKS || '').trim()
    ]);
  } catch (e) { /* kegagalan sinkron master tidak boleh menggagalkan penilaian */ }
}

/* ================= ANTI-DUPLIKAT & KUNCI TULIS =================
   idKirim dibuat sekali di perangkat dan ikut dikirim ulang dari antrean —
   jika sudah pernah tersimpan (respons hilang di jaringan), kiriman ulang
   dikenali sebagai duplikat dan TIDAK ditulis dua kali. */
function adaIdKirim_(sh, kolom, id) {
  if (!id) return false;
  const n = sh.getLastRow();
  if (n < 2) return false;
  const mulai = Math.max(2, n - 499); // periksa 500 baris terakhir
  const v = sh.getRange(mulai, kolom, n - mulai + 1, 1).getValues();
  return v.some(r => String(r[0]) === String(id));
}

/* ================= SUBMIT PENILAIAN KS ================= */
function submitKS_(d, paud) {
  if (!d || !d.ident || !d.nilai) throw new Error('Data tidak lengkap');
  const grpDef = paud ? KS_PAUD : KS_SDSMP;
  const jumlahInd = grpDef.reduce((n, g) => n + g[1].length, 0);
  if (d.nilai.length !== jumlahInd) throw new Error('Jumlah indikator tidak sesuai');

  const id = d.ident;
  const row = [new Date(), id.email || '', "'" + (id.npsn || ''), id.kecamatan || '', id.jenjang || '', id.sekolah || '',
               id.pengawas || '', "'" + (id.hpPengawas || ''), id.ks || '', "'" + (id.hpKS || '')];
  if (!paud) row.push(id.periode || '');
  row.push(id.bulan || '', id.minggu || '');

  let sB = 0, sT = 0, sC = 0;
  const grpStat = grpDef.map(() => ({ b: 0, t: 0, c: 0, n: 0 }));
  let idx = 0;
  grpDef.forEach((g, gi) => g[1].forEach(() => {
    const v = d.nilai[idx++];
    const b = Number(v.b), t = Number(v.t), c = Number(v.c);
    if ([b, t, c].some(x => isNaN(x) || x < 1 || x > 5)) throw new Error('Nilai harus 1–5');
    row.push(b, t, c, v.link || '-');
    sB += b; sT += t; sC += c;
    grpStat[gi].b += b; grpStat[gi].t += t; grpStat[gi].c += c; grpStat[gi].n++;
  }));

  row.push(d.catatan || '');
  if (!paud) row.push(''); // Nomor Urut (kompatibilitas data lama)
  row.push(sB / jumlahInd, sT / jumlahInd, sC / jumlahInd, (sC - sB) / jumlahInd);
  grpStat.forEach(s => row.push(s.b / s.n, s.t / s.n, s.c / s.n, (s.c - s.b) / s.n));

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(paud ? NAMA_SHEET.PAUD : NAMA_SHEET.SDSMP);
  const kolomId = row.length + 1; // ID KIRIM di kolom terakhir (PAUD 110 / SDSMP 148)
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // kunci tulis: cegah baris saling menyela saat pengisian serentak
  try {
    if (adaIdKirim_(sh, kolomId, d.idKirim)) {
      return { ok: true, duplikat: true, rataCapaian: sC / jumlahInd, kinerja: (sC - sB) / jumlahInd };
    }
    row.push(d.idKirim || '');
    sh.appendRow(row);
  } finally { lock.releaseLock(); }
  perbaruiMaster_(id); // sinkronkan perbaikan identitas ke MASTER_SEKOLAH
  return { ok: true, rataCapaian: sC / jumlahInd, kinerja: (sC - sB) / jumlahInd };
}

/* ================= SUBMIT PENILAIAN GURU ================= */
function submitGuru_(d) {
  if (!d || !d.ident || !d.nilai) throw new Error('Data tidak lengkap');
  const id = d.ident;
  const row = [new Date(), id.email || '', id.kecamatan || '', id.jenjang || '', "'" + (id.npsn || ''), id.sekolah || '',
               id.ks || '', "'" + (id.hpKS || ''), id.nama || '', id.status || '', id.pangkat || '',
               "'" + (id.nip || ''), id.mapel || '', "'" + (id.hp || ''), id.bulan || ''];
  const rata = [];
  GURU_GRP.forEach((g, gi) => {
    const vals = d.nilai[gi];
    if (!vals || vals.length !== g[1]) throw new Error('Jumlah butir kelompok ' + (gi + 1) + ' tidak sesuai');
    let s = 0;
    vals.forEach(v => {
      // checklist 0/1 ATAU skor survey desimal 0–1 (terima koma/titik, 2 desimal)
      let n = Number(String(v).replace(',', '.'));
      if (isNaN(n) || n < 0) n = 0;
      if (n > 1) n = 1;
      n = Math.round(n * 100) / 100;
      row.push(n); s += n;
    });
    row.push((d.links && d.links[gi]) || '-');
    rata.push(Math.round(s / g[1] * 10000) / 100);
  });
  rata.forEach(r => row.push(r));
  const rataAll = Math.round(rata.reduce((a, b) => a + b, 0) / rata.length * 100) / 100;
  row.push(rataAll);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMA_SHEET.GURU);
  const kolomId = row.length + 1; // ID KIRIM = kolom 60
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (adaIdKirim_(sh, kolomId, d.idKirim)) {
      return { ok: true, duplikat: true, rata: rata, rataKeseluruhan: rataAll };
    }
    row.push(d.idKirim || '');
    sh.appendRow(row);
  } finally { lock.releaseLock(); }
  return { ok: true, rata: rata, rataKeseluruhan: rataAll };
}

/* ================= RAPOR (publik) =================
   Mengembalikan SEMUA pengisian yang cocok pada bulan tsb —
   jika lebih dari satu, pengguna memilih mana yang dicetak. */
function getRapor_(p) {
  const jenis = (p.jenis || '').toLowerCase();
  let hit;
  if (jenis === 'guru') {
    const rows = bacaSemua_(NAMA_SHEET.GURU);
    hit = rows.filter(r => String(r[4]).replace(/^'/,'') === String(p.npsn) && String(r[14]) === String(p.bulan) &&
      (!p.nip || String(r[11]).replace(/\D/g,'') === String(p.nip).replace(/\D/g,'')) &&
      (!p.nama || String(r[8]).trim().toLowerCase() === String(p.nama).trim().toLowerCase()));
  } else {
    const paud = jenis === 'paud';
    const rows = bacaSemua_(paud ? NAMA_SHEET.PAUD : NAMA_SHEET.SDSMP);
    const bulanCol = paud ? 10 : 11;
    hit = rows.filter(r => String(r[2]).replace(/^'/,'') === String(p.npsn) && String(r[bulanCol]) === String(p.bulan));
  }
  if (!hit.length) return { ditemukan: false };
  // row: kompatibilitas lama (pengisian terakhir); rows: semua pengisian
  return { ditemukan: true, jumlah: hit.length, rows: hit, row: hit[hit.length - 1] };
}

/* ================= CETAK BANYAK (publik) =================
   KS  : semua pengisian sekolah binaan satu PENGAWAS pada bulan tsb
   Guru: semua pengisian guru yang dinilai satu KEPALA SEKOLAH pada bulan tsb */
function getRaporBanyak_(p) {
  const jenis = (p.jenis || '').toLowerCase();
  const sama = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
  let hit;
  if (jenis === 'guru') {
    const rows = bacaSemua_(NAMA_SHEET.GURU);
    hit = rows.filter(r => sama(r[6], p.ks) && String(r[14]) === String(p.bulan)); // kolom 7 = NAMA KEPALA SEKOLAH
  } else {
    const paud = jenis === 'paud';
    const rows = bacaSemua_(paud ? NAMA_SHEET.PAUD : NAMA_SHEET.SDSMP);
    const bulanCol = paud ? 10 : 11;
    hit = rows.filter(r => sama(r[6], p.pengawas) && String(r[bulanCol]) === String(p.bulan)); // kolom 7 = NAMA PENGAWAS
  }
  // urutkan: nama sekolah lalu waktu pengisian
  hit.sort((a, b) => String(a[5]).localeCompare(String(b[5])) || String(a[0]).localeCompare(String(b[0])));
  return { ditemukan: hit.length > 0, jumlah: hit.length, rows: hit };
}

/* ================= HASIL PENILAIAN PER KECAMATAN (admin kecamatan & admin dinas) ================= */
/* daftar nama kecamatan unik dari MASTER_SEKOLAH, untuk selector admin dinas */
function daftarKecamatan_() {
  const set = new Set();
  getMaster_().sekolah.forEach(s => { if (s.kecamatan) set.add(String(s.kecamatan).trim()); });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'id'));
}
/* akses Hasil Penilaian & edit/hapus: admin dinas (token, semua kecamatan) ATAU admin kecamatan (auth, kecamatannya sendiri) */
function resolveAksesHasil_(b) {
  if (b && b.token) {
    cekToken_(b.token);
    return { admin: true, kecamatan: null };
  }
  if (!b || !b.auth || b.auth.peran !== 'kecamatan') throw new Error('Hanya admin kecamatan atau admin Dinas yang dapat mengakses data ini');
  const akun = cekAuthUser_(b.auth);
  return { admin: false, kecamatan: norm_(String(akun.data[0]).replace(/^'/, '')) };
}
function getHasilKecamatan_(usernameKec) {
  const ambil = (nama, kolKec) => {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
    if (!sh || sh.getLastRow() < 2) return [];
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const out = [];
    v.forEach((r, i) => {
      if (norm_(r[kolKec]) === norm_(usernameKec)) out.push({ baris: i + 2, row: r });
    });
    return out;
  };
  return {
    kecamatan: usernameKec,
    sdsmp: ambil(NAMA_SHEET.SDSMP, 3), // kolom 4 = Kecamatan
    paud: ambil(NAMA_SHEET.PAUD, 3),
    guru: ambil(NAMA_SHEET.GURU, 2)    // kolom 3 = KECAMATAN
  };
}

/* ===== EDIT HASIL PENILAIAN KS (admin kecamatan / admin dinas; identitas tidak diubah) ===== */
function editKS_(b, jenis, baris, nilai, catatan) {
  const akses = resolveAksesHasil_(b);
  const paud = String(jenis).toLowerCase() === 'paud';
  const grpDef = paud ? KS_PAUD : KS_SDSMP;
  const jumlahInd = grpDef.reduce((n, g) => n + g[1].length, 0);
  if (!nilai || nilai.length !== jumlahInd) throw new Error('Jumlah indikator tidak sesuai');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(paud ? NAMA_SHEET.PAUD : NAMA_SHEET.SDSMP);
  baris = Number(baris);
  if (!baris || baris < 2 || baris > sh.getLastRow()) throw new Error('Baris tidak valid');

  const lebar = sh.getLastColumn();
  const row = sh.getRange(baris, 1, 1, lebar).getValues()[0];
  // keamanan: baris harus milik kecamatan akun (kolom 4 = Kecamatan) — admin Dinas bebas akses semua kecamatan
  if (!akses.admin && norm_(row[3]) !== akses.kecamatan)
    throw new Error('Data ini bukan milik kecamatan Anda');

  const identN = paud ? 12 : 13;
  let sB = 0, sT = 0, sC = 0;
  const grpStat = grpDef.map(() => ({ b: 0, t: 0, c: 0, n: 0 }));
  let idx = 0;
  grpDef.forEach((g, gi) => g[1].forEach(() => {
    const v = nilai[idx];
    const b = Number(v.b), t = Number(v.t), c = Number(v.c);
    if ([b, t, c].some(x => isNaN(x) || x < 1 || x > 5)) throw new Error('Nilai harus 1–5');
    const dasar = identN + idx * 4;
    row[dasar] = b; row[dasar + 1] = t; row[dasar + 2] = c; row[dasar + 3] = v.link || '-';
    sB += b; sT += t; sC += c;
    grpStat[gi].b += b; grpStat[gi].t += t; grpStat[gi].c += c; grpStat[gi].n++;
    idx++;
  }));
  let p = identN + jumlahInd * 4;
  row[p++] = catatan != null ? catatan : row[identN + jumlahInd * 4];
  if (!paud) p++; // lewati Nomor Urut
  row[p++] = sB / jumlahInd; row[p++] = sT / jumlahInd; row[p++] = sC / jumlahInd; row[p++] = (sC - sB) / jumlahInd;
  grpStat.forEach(s => { row[p++] = s.b / s.n; row[p++] = s.t / s.n; row[p++] = s.c / s.n; row[p++] = (s.c - s.b) / s.n; });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sh.getRange(baris, 1, 1, lebar).setValues([row]); }
  finally { lock.releaseLock(); }
  return { ok: true, rataCapaian: sC / jumlahInd, kinerja: (sC - sB) / jumlahInd };
}

/* ===== EDIT HASIL PENILAIAN GURU (admin kecamatan / admin dinas) ===== */
function editGuru_(b, baris, nilai, links) {
  const akses = resolveAksesHasil_(b);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMA_SHEET.GURU);
  baris = Number(baris);
  if (!baris || baris < 2 || baris > sh.getLastRow()) throw new Error('Baris tidak valid');
  const lebar = sh.getLastColumn();
  const row = sh.getRange(baris, 1, 1, lebar).getValues()[0];
  if (!akses.admin && norm_(row[2]) !== akses.kecamatan)
    throw new Error('Data ini bukan milik kecamatan Anda');

  let p = 15; // kolom 16 = butir pertama
  const rata = [];
  GURU_GRP.forEach((g, gi) => {
    const vals = nilai && nilai[gi];
    if (!vals || vals.length !== g[1]) throw new Error('Jumlah butir kelompok ' + (gi + 1) + ' tidak sesuai');
    let s = 0;
    vals.forEach(v => {
      let n = Number(String(v).replace(',', '.'));
      if (isNaN(n) || n < 0) n = 0;
      if (n > 1) n = 1;
      n = Math.round(n * 100) / 100;
      row[p++] = n; s += n;
    });
    if (links && links[gi] != null && String(links[gi]).trim() !== '') row[p] = links[gi];
    p++; // kolom link
    rata.push(Math.round(s / g[1] * 10000) / 100);
  });
  rata.forEach(r2 => { row[p++] = r2; });
  row[p] = Math.round(rata.reduce((a, b) => a + b, 0) / rata.length * 100) / 100;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sh.getRange(baris, 1, 1, lebar).setValues([row]); }
  finally { lock.releaseLock(); }
  return { ok: true, rata: rata, rataKeseluruhan: row[p] };
}

/* ===== HAPUS HASIL PENILAIAN KS (admin kecamatan / admin dinas) ===== */
function deleteKS_(b, jenis, baris) {
  const akses = resolveAksesHasil_(b);
  const paud = String(jenis).toLowerCase() === 'paud';
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(paud ? NAMA_SHEET.PAUD : NAMA_SHEET.SDSMP);
  baris = Number(baris);
  if (!baris || baris < 2 || baris > sh.getLastRow()) throw new Error('Baris tidak valid');
  if (!akses.admin) {
    const row = sh.getRange(baris, 1, 1, 4).getValues()[0];
    if (norm_(row[3]) !== akses.kecamatan) throw new Error('Data ini bukan milik kecamatan Anda');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sh.deleteRow(baris); }
  finally { lock.releaseLock(); }
  return { ok: true };
}

/* ===== HAPUS HASIL PENILAIAN GURU (admin kecamatan / admin dinas) ===== */
function deleteGuru_(b, baris) {
  const akses = resolveAksesHasil_(b);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NAMA_SHEET.GURU);
  baris = Number(baris);
  if (!baris || baris < 2 || baris > sh.getLastRow()) throw new Error('Baris tidak valid');
  if (!akses.admin) {
    const row = sh.getRange(baris, 1, 1, 3).getValues()[0];
    if (norm_(row[2]) !== akses.kecamatan) throw new Error('Data ini bukan milik kecamatan Anda');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { sh.deleteRow(baris); }
  finally { lock.releaseLock(); }
  return { ok: true };
}

/* ================= LAPORAN & TARIK DATA (admin) ================= */
function getLaporan_(denganDetail) {
  const sdsmp = bacaSemua_(NAMA_SHEET.SDSMP);
  const paud = bacaSemua_(NAMA_SHEET.PAUD);
  const guru = bacaSemua_(NAMA_SHEET.GURU);

  // ringkasan KS — kolom agregat dihitung ulang bila kosong (data lama dari Google Form)
  const ringkasKS = (rows, isPaud) => rows.map(r => {
    const nInd = isPaud ? 17 : 26;
    const identN = isPaud ? 12 : 13;
    let rataB = Number(r[identN + nInd * 4 + (isPaud ? 1 : 2)]);
    let rataC = Number(r[identN + nInd * 4 + (isPaud ? 3 : 4)]);
    if (!rataB || isNaN(rataB)) { // hitung ulang dari kolom indikator
      let sb = 0, sc = 0;
      for (let i = 0; i < nInd; i++) { sb += Number(r[identN + i * 4]) || 0; sc += Number(r[identN + i * 4 + 2]) || 0; }
      rataB = sb / nInd; rataC = sc / nInd;
    }
    return {
      jenis: isPaud ? 'paud' : 'sdsmp', ts: r[0], npsn: String(r[2]).replace(/^'/,''), kecamatan: r[3],
      jenjang: r[4], sekolah: r[5], pengawas: r[6], ks: r[8],
      periode: isPaud ? '' : r[10], bulan: isPaud ? r[10] : r[11], minggu: isPaud ? r[11] : r[12],
      rataB: rataB, rataC: rataC, kinerja: rataC - rataB
    };
  });
  const ringkasGuru = guru.map(r => ({
    ts: r[0], kecamatan: r[2], jenjang: r[3], npsn: String(r[4]).replace(/^'/,''), sekolah: r[5],
    nama: r[8], nip: String(r[11]), mapel: r[12], bulan: r[14],
    nilai: [Number(r[53]), Number(r[54]), Number(r[55]), Number(r[56]), Number(r[57])],
    rata: Number(r[58])
  }));

  const out = {
    pengaturan: bacaPengaturan_(),
    ks: ringkasKS(sdsmp, false).concat(ringkasKS(paud, true)),
    guru: ringkasGuru
  };
  if (denganDetail) {
    out.master = getMaster_();
    out.masterGuru = bacaSemua_(NAMA_SHEET.GURU_M);
    out.detail = { sdsmp: sdsmp, paud: paud, guru: guru };
  }
  return out;
}
