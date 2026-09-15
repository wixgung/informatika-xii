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
// Perilaku pengecekan & kecepatan:
// - Daftar siswa mulai diambil begitu pasangAutoisiSiswa() dipanggil
//   (saat halaman kuis dibuka), BUKAN menunggu siswa mulai mengetik.
//   Jadi saat siswa selesai mengetik NIS, datanya biasanya sudah siap
//   di memori dan pencariannya jadi instan.
// - Selama menunggu, kalau ada beberapa pemicu sekaligus (mengetik +
//   blur + klik tombol), semuanya berbagi satu proses pengambilan
//   yang sama — tidak memicu beberapa request sekaligus.
// - Hasilnya juga disimpan sementara di sessionStorage (kadaluarsa
//   30 menit) supaya kalau siswa membuka kuis lain di jam yang sama,
//   tidak perlu mengambil ulang dari awal — halaman kuis berikutnya
//   langsung baca dari cache ini tanpa menunggu.
// - Otomatis dicek begitu NIS sudah diketik minimal 4 digit, juga
//   saat kolom NIS kehilangan fokus (blur), dan lewat tombol manual.
// ============================================================
const _KUNCI_CACHE_SISWA = 'daftarSiswaCache';
const _MASA_BERLAKU_CACHE_MS = 30 * 60 * 1000; // 30 menit

let _daftarSiswaPromise = null;

function ambilDaftarSiswa() {
  if (_daftarSiswaPromise) return _daftarSiswaPromise; // sudah/sedang diambil, jangan ulang

  // 1) Coba dari sessionStorage dulu (bertahan antar-halaman kuis)
  try {
    const mentah = sessionStorage.getItem(_KUNCI_CACHE_SISWA);
    if (mentah) {
      const cache = JSON.parse(mentah);
      if (Date.now() - cache.waktu < _MASA_BERLAKU_CACHE_MS) {
        _daftarSiswaPromise = Promise.resolve(cache.data);
        return _daftarSiswaPromise;
      }
    }
  } catch (e) { /* sessionStorage bermasalah -> abaikan, lanjut fetch biasa */ }

  // 2) Kalau tidak ada cache valid, ambil dari server sekali saja
  if (typeof CONTROL_URL === "undefined") return Promise.resolve([]);
  _daftarSiswaPromise = fetch(CONTROL_URL + "?action=siswa")
    .then(res => res.json())
    .then(data => {
      try {
        sessionStorage.setItem(_KUNCI_CACHE_SISWA, JSON.stringify({ waktu: Date.now(), data }));
      } catch (e) { /* sessionStorage penuh/nonaktif -> tidak masalah, tetap jalan */ }
      return data;
    })
    .catch(err => {
      console.warn("Gagal mengambil daftar siswa:", err);
      _daftarSiswaPromise = null; // biar bisa dicoba ulang lain kali
      return [];
    });
  return _daftarSiswaPromise;
}

function pasangAutoisiSiswa(idNis, idNama, idKelas, idAbsen, idTombol) {
  const inputNis = document.getElementById(idNis);
  const inputNama = document.getElementById(idNama);
  const inputKelas = document.getElementById(idKelas);
  const inputAbsen = idAbsen ? document.getElementById(idAbsen) : null;
  const tombolCek = idTombol ? document.getElementById(idTombol) : null;
  if (!inputNis || !inputNama || !inputKelas) return;

  // Mulai ambil data dari SEKARANG (saat halaman kuis dibuka), jangan
  // tunggu sampai siswa mengetik — ini yang memangkas rasa "lama".
  ambilDaftarSiswa();

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
