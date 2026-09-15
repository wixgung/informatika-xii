# Panduan Lengkap: Portal Belajar Informatika XII

Dokumen ini adalah **satu-satunya panduan yang perlu dilampirkan** di sesi berikutnya — sudah mencakup setup awal maupun semua perbaikan (anti-gagal-kirim, cek duplikat + kirim ulang, auto-isi identitas dari NIS, mode Kelompok dinonaktifkan).

Konteks: portal ini **baru dan terpisah** dari portal Informatika Kelas XI — Google Sheet dan deployment Apps Script sendiri, supaya data kedua kelas tidak tercampur.

---

## Bagian A — Struktur Google Sheet (3 tab)

Buat/pastikan ada Spreadsheet, bebas namanya (misal **"Data Portal Analisis Data XII"**), dengan 3 tab berikut.

### Tab "Materi" — pengganti penuh modules.json

Header baris pertama, urutan kolom harus persis:

| ID | Judul | Deskripsi | Kategori | Semester | Icon | Gambar | File | Tanggal | Tampilkan |
|---|---|---|---|---|---|---|---|---|---|

- **ID** — kode unik, misal `XII-105`. Harus sama dengan `MODUL_ID` di dalam file HTML materi.
- **Semester** — isi `1` atau `2`.
- **Icon** — satu emoji (boleh kosong, default 📄).
- **Gambar** — boleh kosong (kartu dapat warna pastel otomatis). Kalau diisi: link gambar langsung (imgur/Unsplash/Drive direct-link), atau path lokal `images/nama.jpg` yang diupload ke folder `images/` di repo.
- **File** — path relatif ke file HTML di folder `modules/`, persis nama filenya.
- **Tanggal** — bebas, hanya catatan.
- **Tampilkan** — kolom ini di-set jadi **Insert → Checkbox**. Centang = materi tampil ke siswa.

### Tab "HasilKuis" — terisi otomatis saat siswa submit

| Waktu | Mode | Nama | NIS | No Absen | Kelas | Anggota Kelompok | Materi | Skor | Total |
|---|---|---|---|---|---|---|---|---|---|

- Satu siswa/materi = **satu baris**. Kalau siswa kirim ulang, baris ini **ditimpa otomatis** dengan hasil terbaru (tidak dobel).
- Kolom "Mode" akan selalu `Individu` (fitur Kelompok sudah dinonaktifkan di sisi kuis — lihat Bagian D).

### Tab "Siswa" — untuk auto-isi identitas dari NIS

Minimal berisi kolom **NIS, Nama, Kelas, Absen** — boleh urutannya berbeda, boleh huruf besar/kecil bebas (mis. `NAMA` atau `Nama` sama-sama kebaca), dan boleh ada kolom tambahan seperti `JK` (akan diabaikan). Contoh:

| NIS | NAMA | JK | ABSEN | KELAS |
|---|---|---|---|---|

- **Wajib diisi lengkap** untuk semua siswa yang akan mengerjakan kuis. Sejak field Nama/Kelas/Absen di form kuis dikunci (readonly), siswa yang NIS-nya tidak ada di tab ini **tidak bisa lanjut mengerjakan** — dia hanya akan melihat pesan untuk memeriksa NIS atau menghubungi guru.

---

## Bagian B — Pasang Apps Script

1. Di Spreadsheet: **Extensions → Apps Script**.
2. Hapus isi default, tempel kode berikut secara utuh:

```javascript
// ============================================================
// APPS SCRIPT — Portal Informatika XII
// ============================================================

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "siswa") {
    return jsonOutput(getSiswaList(ss));
  }
  if (action === "cekHasil") {
    return jsonOutput(cekHasilExisting(ss, e.parameter));
  }

  // Default (tanpa action): daftar Materi
  const sheet = ss.getSheetByName("Materi");
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const idx = {};
  header.forEach((h, i) => idx[h.trim()] = i);

  const data = rows
    .filter(r => r[idx["ID"]] !== "")
    .map(r => ({
      id: String(r[idx["ID"]]).trim(),
      title: r[idx["Judul"]],
      description: r[idx["Deskripsi"]],
      category: r[idx["Kategori"]],
      semester: String(r[idx["Semester"]]).trim(),
      icon: r[idx["Icon"]] || "📄",
      image: r[idx["Gambar"]] || "",
      file: r[idx["File"]],
      tanggal: r[idx["Tanggal"]] ? String(r[idx["Tanggal"]]) : "",
      tampilkan: r[idx["Tampilkan"]] === true
    }));

  return jsonOutput(data);
}

function doPost(e) {
  // Lock: kalau 2 siswa submit bersamaan, yang kedua menunggu
  // giliran (maks 30 detik) supaya tidak saling menimpa baris.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("HasilKuis");
    const data = JSON.parse(e.postData.contents);

    const rows = sheet.getDataRange().getValues();
    const header = rows[0];
    const idx = {};
    header.forEach((h, i) => idx[h.trim()] = i);
    const kunciBaru = buatKunci(data);

    let baris = -1;
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const dataBaris = {
        mode: r[idx["Mode"]], nis: r[idx["NIS"]], nama: r[idx["Nama"]],
        kelas: r[idx["Kelas"]], materi: r[idx["Materi"]]
      };
      if (buatKunci(dataBaris) === kunciBaru) { baris = i; break; }
    }

    const nilaiBaris = [
      data.waktu, data.mode, data.nama, data.nis || "", data.absen || "",
      data.kelas || "", data.anggota || "", data.materi, data.skor, data.total
    ];

    let diperbarui = false;
    if (baris === -1) {
      sheet.appendRow(nilaiBaris);
    } else {
      sheet.getRange(baris + 1, 1, 1, nilaiBaris.length).setValues([nilaiBaris]);
      diperbarui = true;
    }

    // Pastikan penulisan benar-benar tersimpan sebelum menjawab ke client.
    SpreadsheetApp.flush();

    return jsonOutput({ status: "ok", diperbarui: diperbarui });
  } catch (err) {
    return jsonOutput({ status: "error", pesan: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- Util bersama ----------

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buatKunci(data) {
  if (data.mode === "Individu") {
    return "IND|" + String(data.nis || "").trim() + "|" + String(data.materi || "").trim();
  }
  return "KEL|" + String(data.nama || "").trim().toLowerCase() + "|" +
    String(data.kelas || "").trim().toLowerCase() + "|" + String(data.materi || "").trim();
}

function cekHasilExisting(ss, params) {
  const sheet = ss.getSheetByName("HasilKuis");
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  const idx = {};
  header.forEach((h, i) => idx[h.trim()] = i);
  const kunciDicari = buatKunci(params);

  for (const r of rows) {
    const dataBaris = {
      mode: r[idx["Mode"]], nis: r[idx["NIS"]], nama: r[idx["Nama"]],
      kelas: r[idx["Kelas"]], materi: r[idx["Materi"]]
    };
    if (buatKunci(dataBaris) === kunciDicari) {
      return { ada: true, skor: r[idx["Skor"]], total: r[idx["Total"]], waktu: r[idx["Waktu"]] };
    }
  }
  return { ada: false };
}

function getSiswaList(ss) {
  const sheet = ss.getSheetByName("Siswa");
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  const header = rows.shift();
  // Dicocokkan tanpa peduli besar-kecil huruf & urutan kolom, supaya
  // header "NIS/NAMA/KELAS/ABSEN" atau "Nis/Nama/Kelas/Absen" sama-sama
  // kebaca, dan kolom tambahan (mis. "JK") aman diabaikan.
  const idx = {};
  header.forEach((h, i) => idx[String(h).trim().toUpperCase()] = i);
  return rows
    .filter(r => String(r[idx["NIS"]]).trim() !== "")
    .map(r => ({
      nis: String(r[idx["NIS"]]).trim(),
      nama: idx["NAMA"] !== undefined ? r[idx["NAMA"]] : "",
      kelas: idx["KELAS"] !== undefined ? r[idx["KELAS"]] : "",
      absen: idx["ABSEN"] !== undefined ? String(r[idx["ABSEN"]]).trim() : ""
    }));
}
```

3. Simpan project (nama bebas, misal "API Portal Informatika XII").
4. **Deploy → New deployment** → ikon gear ⚙️ → **Web app** → **Execute as: Me**, **Who has access: Anyone** → **Deploy** → setujui izin akses.
5. Salin **Web app URL** (`https://script.google.com/macros/s/xxxxx/exec`).

> ⚠️ **Kalau nanti mengedit ulang kode ini** (bukan setup pertama kali): setelah tempel kode baru, wajib **Deploy → Manage deployments → klik ikon pensil → Version: New version → Deploy**. Kalau cuma disimpan tanpa bikin versi baru, URL exec lama tetap menjalankan kode **lama**.

---

## Bagian C — File di repo GitHub (root, sejajar folder `modules/`)

### `config.js` — satu-satunya file yang diedit rutin

```javascript
// ============================================================
// KONFIGURASI PORTAL — edit HANYA baris di bawah ini
// ============================================================
const CONTROL_URL = "https://script.google.com/macros/s/GANTI_DENGAN_ID_DEPLOYMENT/exec";
```

Ganti dengan URL asli dari Bagian B langkah 5.

### `guard.js` — jangan diedit, cukup dipasang

Mengecek status "Tampilkan" dari Sheet sebelum materi bisa diakses. Kalau `Tampilkan` = FALSE, siswa yang buka link langsung akan melihat halaman "Materi belum dibuka" alih-alih isi materi.

```javascript
(async function () {
  const kontenUtama = document.getElementById("konten-utama");

  function sembunyikanLoader() {
    const loader = document.getElementById("guard-loading");
    if (loader) loader.remove();
  }

  function tampilkanTerkunci() {
    document.body.innerHTML = `
      <div style="font-family:Inter,ui-sans-serif,system-ui,sans-serif;text-align:center;
                  padding:100px 24px;color:#787774;max-width:420px;margin:0 auto">
        <div style="font-size:48px">🔒</div>
        <h2 style="color:#37352F;margin-top:14px;font-size:20px">Materi belum dibuka</h2>
        <p style="margin-top:8px;font-size:14px;line-height:1.5">
          Materi ini belum diaktifkan oleh guru. Silakan cek kembali nanti.
        </p>
        <p style="margin-top:20px">
          <a href="../index.html" style="color:#2383E2;font-size:14px;text-decoration:none">&larr; Kembali ke Portal</a>
        </p>
      </div>`;
  }

  if (typeof MODUL_ID === "undefined" || typeof CONTROL_URL === "undefined") {
    console.warn("Guard: MODUL_ID atau CONTROL_URL tidak terpasang. Menampilkan konten apa adanya.");
    if (kontenUtama) kontenUtama.style.visibility = "visible";
    sembunyikanLoader();
    return;
  }

  try {
    const res = await fetch(CONTROL_URL + "?t=" + Date.now());
    const data = await res.json();
    const item = data.find(d => String(d.id).trim() === String(MODUL_ID).trim());

    if (item && item.tampilkan === false) {
      tampilkanTerkunci();
      return;
    }
    if (kontenUtama) kontenUtama.style.visibility = "visible";
    sembunyikanLoader();
  } catch (err) {
    console.warn("Guard: tidak bisa memeriksa status visibilitas.", err);
    if (kontenUtama) kontenUtama.style.visibility = "visible";
    sembunyikanLoader();
  }
})();
```

### `siswa.js` — file baru, auto-isi Nama/Kelas/Absen dari NIS

Data siswa mulai diambil sejak halaman kuis dibuka (bukan menunggu siswa mengetik), dan disimpan sementara di `sessionStorage` selama 30 menit supaya kuis-kuis berikutnya di jam yang sama tidak perlu mengambil ulang.

```javascript
// ============================================================
// SISWA — auto-isi Nama, Kelas & No. Absen begitu siswa mengetik
// NIS, datanya diambil dari tab "Siswa" di Google Sheet.
// Pasang di file kuis, SETELAH config.js:
//   <script src="../config.js"></script>
//   <script src="../siswa.js"></script>
// Lalu panggil sekali di bagian script kuis:
//   pasangAutoisiSiswa('gNis', 'gNama', 'gKelas', 'gAbsen', 'gCekNisBtn');
// (idAbsen dan idTombol boleh dikosongkan kalau file itu tidak
// punya field/tombol itu)
// ============================================================
const _KUNCI_CACHE_SISWA = 'daftarSiswaCache';
const _MASA_BERLAKU_CACHE_MS = 30 * 60 * 1000; // 30 menit

let _daftarSiswaPromise = null;

function ambilDaftarSiswa() {
  if (_daftarSiswaPromise) return _daftarSiswaPromise;

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
      _daftarSiswaPromise = null;
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
  // tunggu sampai siswa mengetik.
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

  inputNis.addEventListener('input', () => {
    if (inputNis.value.trim().length >= 4) cariSiswa();
  });
  inputNis.addEventListener('blur', cariSiswa);
  if (tombolCek) tombolCek.addEventListener('click', cariSiswa);
}
```

Upload ketiga file ini (`config.js`, `guard.js`, `siswa.js`) di root repo, sejajar dengan folder `modules/`.

---

## Bagian D — Pola wajib di setiap file kuis (HTML)

Setiap file materi/kuis di folder `modules/` harus mengikuti pola ini. `fitur-dasar-excel.html` sudah jadi contoh lengkapnya — untuk kuis lain, terapkan 5 hal yang sama:

### 1. Script include, tepat sebelum `</body>`

```html
<script>const MODUL_ID = "XII-XXX";</script>
<script src="../config.js"></script>
<script src="../guard.js"></script>
<script src="../siswa.js"></script>
```

Ganti `XII-XXX` dengan ID unik materi tersebut (harus sama dengan kolom **ID** di tab Materi).

### 2. Form identitas: NIS paling atas, field lain dikunci

```html
<div class="idform" id="gFormIndividu">
  <div class="field">
    <label for="gNis">NIS</label>
    <div class="field-nis-row">
      <input id="gNis" type="text" placeholder="Ketik NIS (min. 4 digit)" inputmode="numeric">
      <button type="button" class="btn-cek-nis" id="gCekNisBtn">🔍 Cek</button>
    </div>
  </div>
  <div class="field"><label for="gNama">Nama Lengkap</label><input id="gNama" type="text" placeholder="Otomatis terisi dari NIS" readonly></div>
  <div class="field"><label for="gKelas">Kelas</label><input id="gKelas" type="text" placeholder="Otomatis terisi dari NIS" readonly></div>
  <div class="field"><label for="gAbsen">No. Absen</label><input id="gAbsen" type="text" placeholder="Otomatis terisi dari NIS" readonly></div>
</div>
```

CSS tambahan (taruh dekat aturan `.field input:focus`):

```css
.field input[readonly]{background:var(--paper-2);color:#6b6355;cursor:not-allowed;}
.field-nis-row{display:flex;gap:8px;align-items:stretch;}
.field-nis-row input{flex:1;}
.btn-cek-nis{
  flex-shrink:0;font-family:'IBM Plex Mono',monospace;font-weight:600;font-size:.82rem;
  background:var(--teal);color:#fff;border:none;border-radius:8px;padding:0 16px;cursor:pointer;white-space:nowrap;
}
.btn-cek-nis:hover{background:var(--teal-dark);}
.btn-cek-nis:disabled{opacity:.6;cursor:default;}
```

**Mode Kelompok dinonaktifkan** — tombol pemilihan mode disembunyikan (form tetap ada di HTML untuk kompatibilitas kode, tapi tidak pernah terlihat siswa):

```html
<div class="mode-toggle hidden">
  <button class="pill-btn mode-btn active" id="gModeIndividu" type="button">👤 Individu</button>
  <button class="pill-btn mode-btn" id="gModeKelompok" type="button">👥 Kelompok</button>
</div>
```

### 3. Panggil auto-isi siswa (di blok script utama, setelah toggle mode didefinisikan)

```javascript
if(typeof pasangAutoisiSiswa === 'function'){
  pasangAutoisiSiswa('gNis', 'gNama', 'gKelas', 'gAbsen', 'gCekNisBtn');
}
```

### 4. Validasi saat tombol "Mulai Kuis" diklik

```javascript
if(modeAktif === 'Individu'){
  const nama = document.getElementById('gNama').value.trim();
  const kelas = document.getElementById('gKelas').value.trim();
  const nis = document.getElementById('gNis').value.trim();
  const absen = document.getElementById('gAbsen').value.trim();
  if(!nis){
    alert('Mohon isi NIS terlebih dahulu.');
    return;
  }
  if(!nama || !kelas || !absen){
    alert('NIS tidak ditemukan di data siswa, atau data belum lengkap. Periksa kembali NIS-mu, atau hubungi guru.');
    return;
  }
  ident.nama = nama; ident.kelas = kelas; ident.nis = nis; ident.absen = absen;
}
```

### 5. Fungsi kirim hasil (cek duplikat + verifikasi + retry)

```javascript
const MATERI_LABEL = "Nama Materi Ini"; // sesuaikan per file

async function cekHasilSebelumnya(materi, identitas){
  if(typeof CONTROL_URL === 'undefined') return { ada:false };
  try{
    const params = new URLSearchParams({
      action: 'cekHasil', mode: identitas.mode, materi,
      nis: identitas.nis || '', nama: identitas.nama || '', kelas: identitas.kelas || ''
    });
    const res = await fetch(CONTROL_URL + '?' + params.toString());
    return await res.json();
  }catch(err){
    console.warn('Gagal memeriksa hasil sebelumnya:', err);
    return { ada:false };
  }
}

// Return true = terkirim, false = gagal, null = dibatalkan siswa.
async function kirimKeSheet(materi, skor, total, identitas){
  if(typeof CONTROL_URL === 'undefined'){
    console.warn('CONTROL_URL belum tersedia — lewati pengiriman ke Sheet.');
    return false;
  }

  const cekLama = await cekHasilSebelumnya(materi, identitas);
  if(cekLama.ada){
    const lanjut = confirm(
      'Kamu sudah pernah mengirim hasil kuis ini sebelumnya (skor ' + cekLama.skor + '/' + cekLama.total + ', ' + cekLama.waktu + ').\n\n' +
      'Kirim ulang untuk mengganti dengan hasil yang baru (' + skor + '/' + total + ')?'
    );
    if(!lanjut) return null;
  }

  const payload = {
    waktu: new Date().toLocaleString('id-ID'),
    mode: identitas.mode,
    nama: identitas.nama,
    nis: identitas.nis || '',
    absen: identitas.absen || '',
    kelas: identitas.kelas,
    anggota: identitas.anggota || '',
    materi: materi,
    skor: skor,
    total: total
  };

  for(let percobaan = 1; percobaan <= 2; percobaan++){
    try{
      const res = await fetch(CONTROL_URL, { method: 'POST', body: JSON.stringify(payload) });
      const hasil = await res.json();
      if(hasil.status === 'ok') return true;
      console.warn('Server menolak pengiriman:', hasil);
    }catch(err){
      console.error('Percobaan ' + percobaan + ' gagal mengirim hasil:', materi, err);
    }
    if(percobaan === 1) await new Promise(r => setTimeout(r, 1000));
  }
  return false;
}
```

Dan handler tombol kirimnya:

```javascript
document.getElementById('gSendBtn').addEventListener('click', async ()=>{
  if(!identitas) return;
  const btn = document.getElementById('gSendBtn');
  const note = document.getElementById('gSendNote');
  btn.disabled = true;
  note.textContent = 'Memeriksa...';
  const ok = await kirimKeSheet(MATERI_LABEL + ' — nama kuisnya', skor, TOTAL_SOAL, identitas);
  if(ok === true){
    note.textContent = '✓ Hasil terkirim ke guru.';
  } else if(ok === null){
    note.textContent = 'Dibatalkan — hasil lama tetap tersimpan.';
    btn.disabled = false;
  } else {
    note.textContent = 'Gagal mengirim — coba lagi (periksa koneksi internet).';
    btn.disabled = false;
  }
});
```

---

## Bagian E — Langkah setup dari nol, urut

1. **Buat Google Sheet** 3 tab sesuai Bagian A (Materi, HasilKuis, Siswa — Siswa diisi lengkap NIS/Nama/Kelas/Absen semua murid).
2. **Pasang Apps Script** sesuai Bagian B, deploy sebagai Web App, salin URL exec. Kalau di kemudian hari mengedit `Code.gs` lagi, WAJIB pakai **Manage deployments → New version** (bukan bikin deployment baru), supaya URL exec tidak berubah dan `config.js` tidak perlu diutak-atik lagi.
3. **Siapkan 3 file di root repo**: `config.js` (isi URL exec), `guard.js`, `siswa.js` — sesuai Bagian C.
4. **Buat file materi/kuis** di folder `modules/`, ikuti pola lengkap di Bagian D.
5. **Tambah 1 baris** di tab Materi untuk materi tersebut (ID, Judul, Deskripsi, Kategori, Semester, Icon, File, centang Tampilkan saat siap dibuka).
6. **Uji coba sebagai siswa**: isi NIS yang terdaftar → cek Nama/Kelas/Absen otomatis muncul → kerjakan kuis → kirim hasil → cek muncul di tab HasilKuis. Coba kirim kedua kalinya dengan NIS sama → pastikan muncul dialog konfirmasi dan baris lama ikut terganti (bukan baris baru).

---

## Bagian F — Menambah materi baru (workflow rutin setelah setup)

1. Buat file HTML materinya mengikuti pola Bagian D.
2. Upload ke folder `modules/` di GitHub.
3. Tambah 1 baris di tab **Materi** dengan ID, judul, dsb.
4. Centang **Tampilkan** saat materi siap dibuka untuk siswa.

Tidak ada file kode lain yang perlu disentuh untuk menambah materi — kecuali file HTML materinya sendiri.

---

## Catatan jujur soal batasan

Ini situs statis (GitHub Pages), bukan server sungguhan. Pengecekan visibilitas dan identitas berjalan lewat JavaScript di browser siswa — cukup efektif untuk kebutuhan kelas sehari-hari (materi tersembunyi dari daftar, halaman terkunci walau diakses lewat link langsung, field identitas tidak bisa diedit manual), tapi bukan keamanan tingkat tinggi. Siswa yang sangat mahir teknis (mis. membaca kode sumber atau memanggil Apps Script langsung lewat console) secara teori masih bisa melihat isi mentah atau mengirim data palsu. Untuk kontrol kelas sehari-hari, ini sudah lebih dari cukup.
