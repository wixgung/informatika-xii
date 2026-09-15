// ============================================================
// SISWA — auto-isi Nama, Kelas & No. Absen begitu siswa mengetik
// NIS, datanya diambil dari tab "Siswa" di Google Sheet.
// Pasang di file kuis, SETELAH config.js:
//   <script src="../config.js"></script>
//   <script src="../siswa.js"></script>
// Lalu panggil sekali di bagian script kuis:
//   pasangAutoisiSiswa('gNis', 'gNama', 'gKelas', 'gAbsen', 'gCekNisBtn');
// (ganti id sesuai id input di file kuis tsb; idAbsen dan idTombol
// boleh dikosongkan/diabaikan kalau file itu tidak punya field/tombol itu)
// Kalau tab "Siswa" belum dibuat / kosong, fitur ini diam saja
// dan form tetap bisa diisi manual seperti biasa.
//
// Perilaku pengecekan:
// - Otomatis dicek begitu NIS sudah diketik minimal 4 digit (tiap
//   ketikan berikutnya dicek ulang) — cepat, karena daftar siswa
//   di-cache sekali di memori setelah pengambilan pertama.
// - Juga dicek saat kolom NIS kehilangan fokus (blur), untuk NIS
//   yang panjangnya kurang dari 4 digit.
// - Tombol manual (opsional) untuk siswa yang ingin memicu
//   pengecekan sendiri kapan saja.
// ============================================================
let _daftarSiswaCache = null;

async function ambilDaftarSiswa() {
  if (_daftarSiswaCache) return _daftarSiswaCache;
  if (typeof CONTROL_URL === "undefined") return [];
  try {
    const res = await fetch(CONTROL_URL + "?action=siswa");
    _daftarSiswaCache = await res.json();
  } catch (err) {
    console.warn("Gagal mengambil daftar siswa:", err);
    _daftarSiswaCache = [];
  }
  return _daftarSiswaCache;
}

function pasangAutoisiSiswa(idNis, idNama, idKelas, idAbsen, idTombol) {
  const inputNis = document.getElementById(idNis);
  const inputNama = document.getElementById(idNama);
  const inputKelas = document.getElementById(idKelas);
  const inputAbsen = idAbsen ? document.getElementById(idAbsen) : null;
  const tombolCek = idTombol ? document.getElementById(idTombol) : null;
  if (!inputNis || !inputNama || !inputKelas) return;

  async function cariSiswa() {
    const nis = inputNis.value.trim();
    if (!nis) return;
    if (tombolCek) { tombolCek.disabled = true; tombolCek.textContent = '...'; }
    const daftar = await ambilDaftarSiswa();
    const cocok = daftar.find(s => String(s.nis).trim() === nis);
    if (cocok) {
      inputNama.value = cocok.nama || '';
      inputKelas.value = cocok.kelas || '';
      if (inputAbsen) inputAbsen.value = cocok.absen || '';
    } else {
      inputNama.value = '';
      inputKelas.value = '';
      if (inputAbsen) inputAbsen.value = '';
    }
    if (tombolCek) { tombolCek.disabled = false; tombolCek.textContent = '🔍 Cek'; }
  }

  // Cek instan begitu NIS sudah 4 digit atau lebih
  inputNis.addEventListener('input', () => {
    if (inputNis.value.trim().length >= 4) cariSiswa();
  });
  // Cek juga saat pindah kolom (menjaring NIS < 4 digit)
  inputNis.addEventListener('blur', cariSiswa);
  // Tombol manual, opsional
  if (tombolCek) tombolCek.addEventListener('click', cariSiswa);
}
