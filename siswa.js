// ============================================================
// SISWA — auto-isi Nama, Kelas & No. Absen begitu siswa mengetik
// NIS, datanya diambil dari tab "Siswa" di Google Sheet.
// Pasang di file kuis, SETELAH config.js:
//   <script src="../config.js"></script>
//   <script src="../siswa.js"></script>
// Lalu panggil sekali di bagian script kuis:
//   pasangAutoisiSiswa('gNis', 'gNama', 'gKelas', 'gAbsen', 'gCekNisBtn', 'gStatusNis');
// (idAbsen, idTombol, idStatus boleh dikosongkan kalau file itu
// tidak punya field/tombol/teks status itu)
// Kalau tab "Siswa" belum dibuat / kosong, fitur ini diam saja
// dan form tetap bisa diisi manual seperti biasa.
//
// Perilaku pengecekan & kecepatan:
// - Daftar siswa mulai diambil begitu pasangAutoisiSiswa() dipanggil
//   (saat halaman kuis dibuka), BUKAN menunggu siswa mulai mengetik.
// - Hasilnya disimpan sementara di sessionStorage (30 menit) supaya
//   kuis lain di jam yang sama tidak perlu mengambil ulang. Cache
//   divalidasi bentuknya (harus array berisi field "nis") sebelum
//   dipakai — kalau bentuknya salah/rusak (mis. tersimpan saat
//   Apps Script belum ter-deploy dengan benar), cache itu dibuang
//   otomatis dan diambil ulang dari server, bukan dipakai terus.
// - Ada teks status (opsional) yang membedakan 3 keadaan dengan
//   jelas: "Memeriksa...", "NIS tidak ditemukan...", dan
//   "Gagal memuat data siswa..." — supaya kalau ada masalah,
//   penyebabnya kelihatan, bukan diam saja seperti field kosong.
// ============================================================
const _KUNCI_CACHE_SISWA = 'daftarSiswaCache';
const _MASA_BERLAKU_CACHE_MS = 30 * 60 * 1000; // 30 menit

let _daftarSiswaPromise = null;

// Cache/hasil dianggap valid hanya kalau berupa array DAN (kosong,
// atau) baris pertamanya benar-benar punya field "nis". Ini mencegah
// data yang salah bentuk (mis. daftar Materi ke-fetch karena Apps
// Script belum di-deploy ulang) ikut dipakai sebagai daftar siswa.
function _bentukValid(data) {
  return Array.isArray(data) && (data.length === 0 || Object.prototype.hasOwnProperty.call(data[0], 'nis'));
}

function ambilDaftarSiswa() {
  if (_daftarSiswaPromise) return _daftarSiswaPromise; // sudah/sedang diambil, jangan ulang

  // 1) Coba dari sessionStorage dulu (bertahan antar-halaman kuis)
  try {
    const mentah = sessionStorage.getItem(_KUNCI_CACHE_SISWA);
    if (mentah) {
      const cache = JSON.parse(mentah);
      if (Date.now() - cache.waktu < _MASA_BERLAKU_CACHE_MS && _bentukValid(cache.data)) {
        _daftarSiswaPromise = Promise.resolve(cache.data);
        return _daftarSiswaPromise;
      }
      if (!_bentukValid(cache.data)) {
        console.warn("Cache daftar siswa bentuknya tidak valid, diabaikan & diambil ulang.");
        sessionStorage.removeItem(_KUNCI_CACHE_SISWA);
      }
    }
  } catch (e) { /* sessionStorage bermasalah -> abaikan, lanjut fetch biasa */ }

  // 2) Kalau tidak ada cache valid, ambil dari server sekali saja
  if (typeof CONTROL_URL === "undefined") return Promise.resolve([]);
  _daftarSiswaPromise = fetch(CONTROL_URL + "?action=siswa")
    .then(res => res.json())
    .then(data => {
      if (!_bentukValid(data)) {
        // Bentuk data tidak seperti yang diharapkan (bukan daftar
        // siswa) -> jangan dicache, jangan dianggap berhasil.
        console.warn("Respons ?action=siswa bentuknya tidak sesuai (bukan daftar siswa). Cek apakah Apps Script sudah di-deploy sebagai versi terbaru.", data);
        _daftarSiswaPromise = null;
        return [];
      }
      try {
        sessionStorage.setItem(_KUNCI_CACHE_SISWA, JSON.stringify({ waktu: Date.now(), data }));
      } catch (e) { /* sessionStorage penuh/nonaktif -> tidak masalah, tetap jalan */ }
      return data;
    })
    .catch(err => {
      console.warn("Gagal mengambil daftar siswa (cek CONTROL_URL / koneksi / deployment Apps Script):", err);
      _daftarSiswaPromise = null; // biar bisa dicoba ulang lain kali
      return [];
    });
  return _daftarSiswaPromise;
}

function pasangAutoisiSiswa(idNis, idNama, idKelas, idAbsen, idTombol, idStatus) {
  const inputNis = document.getElementById(idNis);
  const inputNama = document.getElementById(idNama);
  const inputKelas = document.getElementById(idKelas);
  const inputAbsen = idAbsen ? document.getElementById(idAbsen) : null;
  const tombolCek = idTombol ? document.getElementById(idTombol) : null;
  const statusEl = idStatus ? document.getElementById(idStatus) : null;
  if (!inputNis || !inputNama || !inputKelas) return;

  function tulisStatus(teks) {
    if (statusEl) statusEl.textContent = teks;
  }

  // Mulai ambil data dari SEKARANG (saat halaman kuis dibuka), jangan
  // tunggu sampai siswa mengetik — ini yang memangkas rasa "lama".
  ambilDaftarSiswa();

  async function cariSiswa() {
    const nis = inputNis.value.trim();
    if (!nis) { tulisStatus(''); return; }
    if (tombolCek) { tombolCek.disabled = true; tombolCek.textContent = '...'; }
    tulisStatus('Memeriksa...');

    const daftar = await ambilDaftarSiswa();

    if (!Array.isArray(daftar) || daftar.length === 0) {
      inputNama.value = ''; inputKelas.value = ''; if (inputAbsen) inputAbsen.value = '';
      tulisStatus('Gagal memuat data siswa dari server. Periksa koneksi internet, lalu coba klik 🔍 Cek lagi.');
      if (tombolCek) { tombolCek.disabled = false; tombolCek.textContent = '🔍 Cek'; }
      return;
    }

    const cocok = daftar.find(s => String(s.nis).trim() === nis);
    if (cocok) {
      inputNama.value = cocok.nama || '';
      inputKelas.value = cocok.kelas || '';
      if (inputAbsen) inputAbsen.value = cocok.absen || '';
      tulisStatus('');
    } else {
      inputNama.value = '';
      inputKelas.value = '';
      if (inputAbsen) inputAbsen.value = '';
      tulisStatus('NIS tidak ditemukan di data siswa. Periksa kembali NIS-mu.');
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
