# instrumen-praktik-mtssw

Mesin dokumen **Instrumen Ujian Praktik Asesmen Madrasah** untuk **Madrasah Tsanawiyah Syubbanul Wathon**, Tegalrejo (kelas IX, Fase D). Mesin ini dipakai oleh skill Gemini `instrumen-ujian-praktik-umum-mtssw` dan `instrumen-ujian-praktik-pai-arab-mtssw`.

Gemini cukup menulis **data isi** (`window.INSTRUMEN` berisi konfig, isi, dan siswa). Mesin merender 9 komponen A4, toolbar 4 menu (Unggah Cover, Unggah Kop, Unduh Word .docx, Unduh PDF tanpa dialog cetak), kop MTs otomatis, dan pemeriksaan otomatis. Pemeriksaan itu mencakup data siswa kelas IX yang diverifikasi dengan hash per rombel Fattan 1–4, level C3–C6, integrasi KBC SK Dirjen Pendis 6077/2025, dan halaman yang melebihi A4.

Mesin ini dibangun dari kode sumber yang sama dengan mesin MA (`muflihan-ak/instrumen-praktik-masw`), dengan profil satuan MTs.

## Isi repo

| Path | Fungsi |
|---|---|
| `mesin/mesin-instrumen.js` | Mesin v2.2.0 (profil MTs) |
| `assets/kop-mtssw.png` | Kop resmi MTs (1000 px) yang dimuat mesin |

## Alamat mesin yang dipakai skill

- Utama (versi dikunci): `https://cdn.jsdelivr.net/gh/muflihan-ak/instrumen-praktik-mtssw@v2.2.0/mesin/mesin-instrumen.js`
- Cadangan (GitHub Pages): `https://muflihan-ak.github.io/instrumen-praktik-mtssw/mesin/mesin-instrumen.js`

## Ketentuan

- Repo **Public**. **DILARANG** menaruh data pribadi (nama, NIS, NISN siswa, atau data guru) di repo ini. Mesin hanya menyimpan jumlah siswa dan hash per rombel.
- Setiap perubahan mesin dirilis dengan tag versi baru; tag yang sudah dirilis jangan diubah.
