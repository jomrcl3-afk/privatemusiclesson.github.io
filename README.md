# 🎵 Les Musik Privat — Presensi Siswa (Google Apps Script)

Aplikasi web presensi les musik privat berbasis Google Sheets, dengan gaya claymorphism biru muda.

## Fitur
- **Presensi** per periode (4x atau 8x pertemuan), materi dan catatan tiap pertemuan
- **Data siswa**: nama, instrumen, level, orang tua, WhatsApp, jadwal, durasi, biaya, foto, link dokumentasi
- **Penilaian** skala 1–5 dengan laporan yang bisa dicetak / PDF / kirim WhatsApp
- **Invoice** orang tua per periode (terbitkan, tandai lunas, PDF / gambar, kirim WhatsApp)
- **Link orang tua** (hanya baca, khusus satu anak, tanpa sandi)
- Satu sandi untuk pengajar, tanpa admin

## File
| File | Isi |
|---|---|
| `Kode.gs` | Kode server Apps Script |
| `Index.html` | Tampilan aplikasi |

## Cara memasang
1. Buat Google Sheet baru → **Ekstensi → Apps Script**.
2. Ganti isi `Code.gs` dengan isi `Kode.gs`.
3. Tambah file HTML bernama `Index`, isi dengan `Index.html`.
4. Di baris atas `Kode.gs`, ganti `SANDI` dan `NAMA_LES`.
5. Jalankan fungsi `setup` sekali, lalu izinkan akses.
6. **Terapkan → Deployment baru → Aplikasi web** (Jalankan sebagai: Saya · Akses: Siapa saja).

Untuk memperbarui: ganti kodenya, jalankan `setup`, lalu **Kelola deployment → Edit → Versi baru** agar link tetap sama.

> ⚠️ Jangan unggah sandi asli ke GitHub. Biarkan `SANDI = 'GANTI-SANDI-INI'` di repositori, dan isi sandi asli hanya di Apps Script.

> Aplikasi ini berjalan di Google Apps Script, tidak bisa dijalankan lewat GitHub Pages.
