// ============================================================
// SISWA — auto-isi Nama, Kelas & No. Absen begitu siswa mengetik
// NIS, datanya diambil dari tab "Siswa" di Google Sheet.
// Pasang di file kuis, SETELAH config.js:
//   <script src="../config.js"></script>
//   <script src="../siswa.js"></script>
// Lalu panggil sekali di bagian script kuis:
//   pasangAutoisiSiswa('gNis', 'gNama', 'gKelas', 'gAbsen');
// (ganti id sesuai id input di file kuis tsb; idAbsen boleh
// dikosongkan/diabaikan kalau file itu tidak punya field absen)
// Kalau tab "Siswa" belum dibuat / kosong, fitur ini diam saja
// dan form tetap bisa diisi manual seperti biasa.
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

function pasangAutoisiSiswa(idNis, idNama, idKelas, idAbsen) {
  const inputNis = document.getElementById(idNis);
  const inputNama = document.getElementById(idNama);
  const inputKelas = document.getElementById(idKelas);
  const inputAbsen = idAbsen ? document.getElementById(idAbsen) : null;
  if (!inputNis || !inputNama || !inputKelas) return;

  inputNis.addEventListener("blur", async () => {
    const nis = inputNis.value.trim();
    if (!nis) return;
    const daftar = await ambilDaftarSiswa();
    const cocok = daftar.find(s => String(s.nis).trim() === nis);
    if (cocok) {
      inputNama.value = cocok.nama;
      inputKelas.value = cocok.kelas;
      if (inputAbsen) inputAbsen.value = cocok.absen || '';
    } else {
      inputNama.value = '';
      inputKelas.value = '';
      if (inputAbsen) inputAbsen.value = '';
    }
  });
}
