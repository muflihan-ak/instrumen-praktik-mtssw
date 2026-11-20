/*!
 * Mesin Instrumen Ujian Praktik AM — Madrasah Tsanawiyah Syubbanul Wathon, Tegalrejo
 * Versi 2.2.0 · https://github.com/muflihan-ak/instrumen-praktik-mtssw
 * Merender 9 komponen A4 dari window.INSTRUMEN (ditulis Gemini): toolbar 4 menu, kop otomatis,
 * unggah cover/kop, ekspor Word (.docx) dan PDF tanpa dialog cetak, pemeriksaan otomatis.
 * Repo ini TIDAK memuat data pribadi; data siswa dikirim oleh dokumen dan diverifikasi dengan hash.
 */
(function () {
"use strict";
if (window.__MESIN_MASW__) return; window.__MESIN_MASW__ = true;
const VERSI = "2.2.0";
const SRC = (document.currentScript && document.currentScript.src) || "";
const BASE = SRC ? SRC.replace(/mesin\/mesin-instrumen(\.min)?\.js.*$/, "") : "https://cdn.jsdelivr.net/gh/muflihan-ak/instrumen-praktik-mtssw@v2.2.0/";
const SATUAN = {"kode": "MTS", "singkat": "MTs", "kelas": "IX", "fase": "D", "mode": "rombel", "kop": "assets/kop-mtssw.png", "logo": [5, 24, 124]};
const IDENTITAS = Object.freeze({"madrasah": "Madrasah Tsanawiyah Syubbanul Wathon", "madrasahKapital": "MADRASAH TSANAWIYAH SYUBBANUL WATHON", "akreditasi": "A", "nsm": "121233080068", "npsn": "69994000", "alamat": "Jalan Kyai Abdan No.03 Tepo Dlimas Tegalrejo Magelang 56192", "kota": "Tegalrejo", "kepala": "Achmad Jamari, S.Pd.I., Gr.", "nipyKepala": "NIPY. 0103101"});
const REF_SISWA = {"IX FATTAN 1": {"peminatan": "-", "n": 35, "h": 46911652}, "IX FATTAN 2": {"peminatan": "-", "n": 34, "h": 2512165495}, "IX FATTAN 3": {"peminatan": "-", "n": 34, "h": 1586519312}, "IX FATTAN 4": {"peminatan": "-", "n": 36, "h": 2204836315}};
const PROFIL = {"UMUM": {"regulasi": "Kepka BSKAP No. 046/H/KR/2025, KMA No. 1503 Tahun 2025, dan Kepdirjen Pendis No. 6077 Tahun 2025", "sumberCP": "Lampiran Keputusan Kepala BSKAP Kemendikdasmen Nomor 046/H/KR/2025 tentang Capaian Pembelajaran", "k3Label": "Keselamatan dan Kesehatan Kerja (K3)", "integrasiLabel": "KBC (SK Dirjen Pendis 6077/2025): Tema Panca Cinta, Materi Insersi, dan asesmen berbasis proses", "aspekRubrik": ["Persiapan", "Proses", "Hasil/Produk", "Sikap & Nilai Cinta"], "aspekSingkat": ["Persp.", "Proses", "Hasil", "Sikap"]}, "PAI_ARAB": {"regulasi": "Kepdirjen Pendis No. 3302 Tahun 2024, KMA No. 1503 Tahun 2025, dan Kepdirjen Pendis No. 6077 Tahun 2025", "sumberCP": "Keputusan Direktur Jenderal Pendidikan Islam Nomor 3302 Tahun 2024 tentang Capaian Pembelajaran PAI dan Bahasa Arab", "k3Label": "Adab dan Keselamatan Kerja", "integrasiLabel": "KBC (SK Dirjen Pendis 6077/2025): Tema Panca Cinta, Materi Insersi, dan asesmen berbasis proses", "aspekRubrik": ["Persiapan", "Proses", "Hasil", "Adab & Nilai Cinta"], "aspekSingkat": ["Persp.", "Proses", "Hasil", "Adab"]}};
const PANCA = ["Cinta Allah dan Rasul-Nya", "Cinta Ilmu", "Cinta Lingkungan", "Cinta Diri dan Sesama", "Cinta Tanah Air"];
const CSS = "\n  @page { size: A4 portrait; margin: 20mm 20mm 20mm 20mm;\n          @bottom-right { content: \"Hal. \" counter(page); font-family: \"Times New Roman\", \"Liberation Serif\", serif; font-size: 9pt; color: #555; } }\n  @page cover_page { size: A4 portrait; margin: 0mm; @bottom-right { content: none; } }\n\n  :root { --hijau: #0b6b3a; --hijau-muda: #e7f2ea; --emas: #c9a227; --teks: #111; }\n  * { box-sizing: border-box; }\n  html, body { margin: 0; padding: 0; }\n  body { background: #dfe3e6; font-family: \"Times New Roman\", Times, \"Liberation Serif\", serif; color: var(--teks);\n         -webkit-print-color-adjust: exact; print-color-adjust: exact; padding-top: 64px; }\n\n  /* ---------- TOOLBAR MENGAMBANG (tidak ikut tercetak) ---------- */\n  .toolbar { position: fixed; top: 0; left: 0; right: 0; z-index: 999; background: var(--hijau); color: #fff;\n             display: flex; flex-wrap: wrap; align-items: center; gap: 8px; padding: 10px 14px;\n             font-family: \"Segoe UI\", Arial, sans-serif; font-size: 13px; box-shadow: 0 2px 8px rgba(0,0,0,.25); }\n  .toolbar .judul-app { font-weight: 700; margin-right: auto; letter-spacing: .2px; }\n  .toolbar .btn { background: #fff; color: var(--hijau); border: 0; border-radius: 6px; padding: 8px 12px; font-weight: 600;\n                  cursor: pointer; font-size: 13px; display: inline-flex; align-items: center; gap: 6px; }\n  .toolbar .btn:hover { background: #f3e7b8; }\n  .toolbar input[type=file] { display: none; }\n  .toolbar .status { flex-basis: 100%; font-size: 12px; padding: 4px 8px; border-radius: 4px; background: rgba(255,255,255,.15); }\n  .toolbar .status.warn { background: #ffe08a; color: #5a3d00; }\n  .toast { position: fixed; bottom: 18px; left: 50%; transform: translateX(-50%); background: #222; color: #fff; padding: 12px 16px;\n           border-radius: 8px; font-family: Arial, sans-serif; font-size: 13px; max-width: 560px; z-index: 1000; display: none; line-height: 1.5; }\n  .toast a { color: #ffd76a; cursor: pointer; text-decoration: underline; }\n\n  /* ---------- LEMBAR A4 (layar) ---------- */\n  .page-a4 { width: 210mm; min-height: 297mm; padding: 20mm; margin: 0 auto 10mm auto; background: #fff;\n             box-shadow: 0 2px 10px rgba(0,0,0,.18); font-size: 12pt; line-height: 1.35; position: relative; outline: none; }\n  .page-a4[contenteditable=\"true\"]:focus-within { box-shadow: 0 0 0 2px var(--emas), 0 2px 10px rgba(0,0,0,.18); }\n  .kop-slot { margin: 0 0 4mm 0; }\n  .kop-image { width: 100%; height: auto; display: block; margin-bottom: 8px; }\n\n  h1.judul, h2.judul { text-align: center; font-size: 14pt; font-weight: 700; margin: 0 0 1mm 0; text-transform: uppercase; }\n  .subjudul { text-align: center; font-size: 12pt; font-weight: 700; margin: 0 0 5mm 0; }\n  h3 { font-size: 12pt; margin: 4mm 0 2mm 0; break-after: avoid; page-break-after: avoid; }\n  p { margin: 0 0 2mm 0; text-align: justify; }\n  ol, ul { margin: 0 0 2mm 0; padding-left: 7mm; }\n\n  table { width: 100%; table-layout: fixed; border-collapse: collapse; word-wrap: break-word; overflow-wrap: break-word; }\n  table.grid th, table.grid td { border: 0.75pt solid #000; padding: 1.2mm 1.6mm; vertical-align: top; font-size: 11pt; }\n  table.grid th { background: var(--hijau-muda); font-weight: 700; text-align: center; vertical-align: middle; }\n  table.info td { padding: 0.6mm 1mm; vertical-align: top; font-size: 12pt; }\n  table.info td.lbl { width: 32%; } table.info td.sep { width: 3%; }\n  td.c, th.c { text-align: center; }\n  .kotak-kosong { border: 0.75pt solid #000; min-height: 14mm; padding: 2mm; }\n  .hal-verif table.grid td, .hal-verif table.grid th { font-size: 10pt; padding: 0.7mm 1.4mm; line-height: 1.25; }\n  .hal-verif table.info td { font-size: 10.5pt; padding: 0.2mm 1mm; }\n  .hal-verif h3 { margin: 2mm 0 1mm 0; } .hal-verif p { font-size: 11pt; }\n  .hal-verif table.ttd { margin-top: 3mm; } .hal-verif table.ttd td.ruang { height: 13mm; }\n  #tabel-kisi td, #tabel-kisi th { font-size: 10pt; padding: 1mm 1.2mm; }\n  /* Halaman isi 3–7: tipografi padat agar muat 1 A4 */\n  #hal-3, #hal-4, #hal-5, #hal-6, #hal-7 { font-size: 11pt; line-height: 1.3; }\n  #hal-3 table.grid td, #hal-3 table.grid th, #hal-5 table.grid td, #hal-5 table.grid th,\n  #hal-6 table.grid td, #hal-6 table.grid th, #hal-7 table.grid td, #hal-7 table.grid th { font-size: 10.5pt; padding: 0.8mm 1.4mm; line-height: 1.25; }\n  #hal-3 table.info td, #hal-4 table.info td, #hal-5 table.info td { font-size: 11pt; padding: 0.2mm 1mm; }\n  #hal-3 table.info2 td { font-size: 10.5pt; padding: 0.1mm 0.8mm; }\n  #hal-3 table.grid td, #hal-3 table.grid th { font-size: 10pt !important; line-height: 1.2 !important; padding: 0.6mm 1.2mm !important; }\n  #hal-3 ol li { font-size: 10.5pt; }\n  #hal-3 h3, #hal-4 h3, #hal-5 h3, #hal-6 h3, #hal-7 h3 { font-size: 11.5pt; margin: 2.5mm 0 1.2mm 0; }\n  #hal-3 p, #hal-5 p, #hal-6 p, #hal-7 p { margin-bottom: 1.2mm; }\n  #hal-3 .subjudul, #hal-4 .subjudul, #hal-5 .subjudul, #hal-6 .subjudul, #hal-7 .subjudul { margin-bottom: 3mm; }\n  #hal-4 table.ttd td.ruang, #hal-6 table.ttd td.ruang { height: 16mm; }\n  #hal-6 table.grid ol { margin: 0; }\n  .garis-isi { border-bottom: 0.75pt dotted #000; height: 6mm; }\n  .ket-kbc { font-size: 9.5pt !important; margin-top: 1mm; }\n  ol.refleksi li { margin-bottom: 1mm; }\n\n  /* ---------- BLOK TANDA TANGAN (nama 1 baris) ---------- */\n  table.ttd { margin-top: 6mm; page-break-inside: avoid; break-inside: avoid; }\n  table.ttd td { width: 50%; vertical-align: top; text-align: left; padding: 0 0 0 5mm; font-size: 12pt; }\n  table.ttd td.ruang { height: 20mm; }\n  .nama-ttd { font-weight: 700; text-decoration: underline; white-space: nowrap; }\n  .nipy-ttd { white-space: nowrap; }\n\n  /* ---------- COVER (Halaman 1, tanpa margin) ---------- */\n  #page-cover { page: cover_page; padding: 0; width: 210mm; height: 297mm; min-height: 297mm; overflow: hidden; }\n  #page-cover.full-bleed-cover { padding: 0; width: 210mm; height: 297mm; }\n  #page-cover .cover-image { width: 210mm; height: 297mm; object-fit: cover; display: block; }\n  .cv { position: absolute; inset: 0; display: block; }\n  .cv-band { position: absolute; left: 0; right: 0; top: 0; height: 62mm; background: var(--hijau); }\n  .cv-band::after { content: \"\"; position: absolute; left: 0; right: 0; bottom: -3mm; height: 3mm; background: var(--emas); }\n  .cv-top { position: absolute; top: 13mm; left: 0; right: 0; text-align: center; color: #fff; }\n  .cv-top .l1 { font-size: 13pt; letter-spacing: 2px; }\n  .cv-top .l2 { font-size: 19pt; font-weight: 700; letter-spacing: 1px; margin-top: 2mm; }\n  .cv-logo { position: absolute; top: 44mm; left: 50%; transform: translateX(-50%); width: 40mm; height: 40mm; border-radius: 50%;\n             background: #fff; box-shadow: 0 0 0 2mm #fff, 0 0 0 2.8mm var(--emas); display: flex; align-items: center; justify-content: center; }\n  .cv-logo img { width: 36mm; height: 36mm; }\n  .cv-title { position: absolute; top: 98mm; left: 18mm; right: 18mm; text-align: center; }\n  .cv-title .t1 { font-size: 15pt; letter-spacing: 3px; color: #333; }\n  .cv-title .t2 { font-size: 30pt; font-weight: 700; color: var(--hijau); margin: 2mm 0 1mm; line-height: 1.1; }\n  .cv-title .t3 { font-size: 15pt; font-weight: 700; color: #333; }\n  .cv-title .t4 { display: inline-block; margin-top: 5mm; padding: 2mm 8mm; border: 1pt solid var(--emas); font-size: 13pt; font-weight: 700; letter-spacing: 1px; }\n  .cv-box { position: absolute; top: 160mm; left: 26mm; right: 26mm; border: 1pt solid var(--hijau); border-left: 3mm solid var(--hijau); padding: 5mm 7mm; background: #fbfcfb; }\n  .cv-box table td { font-size: 12.5pt; padding: 1mm 0; vertical-align: top; }\n  .cv-box td.lbl { width: 38%; } .cv-box td.sep { width: 4%; }\n  .cv-foot { position: absolute; left: 0; right: 0; bottom: 0; height: 40mm; background: var(--hijau); color: #fff; text-align: center; padding-top: 8mm; }\n  .cv-foot::before { content: \"\"; position: absolute; left: 0; right: 0; top: -3mm; height: 3mm; background: var(--emas); }\n  .cv-foot .f1 { font-size: 15pt; font-weight: 700; letter-spacing: 1px; }\n  .cv-foot .f2 { font-size: 11pt; margin-top: 1.5mm; }\n\n  /* ---------- DAFTAR NILAI (font 8–8.5 pt, 1 halaman per rombel) ---------- */\n  .dn-judul { text-align: center; font-weight: 700; font-size: 12pt; margin: 0; line-height: 1.25; }\n  table.dn-info td { font-size: 9.5pt; padding: 0.3mm 1mm; }\n  table.tbl-nilai th, table.tbl-nilai td { border: 0.75pt solid #000; font-size: 8.5pt; padding: 0.2mm 1mm; line-height: 1.05; vertical-align: middle; }\n  table.tbl-nilai th { font-size: 8pt; line-height: 1.1; padding: 0.6mm 0.5mm; }\n  .page-dn .kop-slot { margin-bottom: 1.5mm; } .page-dn .kop-image { margin-bottom: 2px !important; }\n  table.tbl-nilai th { background: var(--hijau-muda); text-align: center; font-weight: 700; }\n  table.tbl-nilai td.nama { white-space: nowrap; overflow: hidden; text-overflow: clip; }\n  table.tbl-nilai td.skor { text-align: center; }\n  .dn-ket { font-size: 8pt; margin: 1mm 0 0 0; }\n  .page-dn table.ttd { margin-top: 2mm; } .page-dn table.ttd td { font-size: 10pt; line-height: 1.2; } .page-dn table.ttd td.ruang { height: 13mm; }\n\n\n  /* ---------- TEKS ARAB (RTL) ---------- */\n  .arabic-text { direction: rtl; unicode-bidi: isolate; text-align: right;\n                 font-family: \"Traditional Arabic\", \"Amiri\", \"Scheherazade New\", \"Times New Roman\", serif; font-size: 15pt; line-height: 1.9; }\n\n  /* ---------- CETAK / PDF (margin 2-2-2-2 cm dari @page) ---------- */\n  @media print {\n    body { background: #fff; padding-top: 0; }\n    .no-print, .toolbar, .toast { display: none !important; }\n    .page-a4 { width: 170mm !important; max-width: 170mm !important; min-height: 0 !important; padding: 0 !important; margin: 0 auto !important;\n               box-shadow: none !important; break-after: page; page-break-after: always; }\n    .page-a4:last-child { break-after: auto; page-break-after: auto; }\n    #page-cover { width: 210mm !important; max-width: 210mm !important; height: 296.5mm !important; padding: 0 !important; margin: 0 !important; }\n    table { width: 100%; table-layout: fixed; word-wrap: break-word; }\n    tr, img { break-inside: avoid; page-break-inside: avoid; }\n  }\n";
const TOOLBAR = "<!-- ======================= TOOLBAR 4 MENU ======================= -->\n<div class=\"toolbar no-print\" contenteditable=\"false\">\n  <span class=\"judul-app\">Instrumen Ujian Praktik AM · MTs Syubbanul Wathon</span>\n  <label class=\"btn\" title=\"Ganti Halaman 1 dengan gambar cover penuh A4\">🖼️ Unggah / Ganti Cover<input type=\"file\" id=\"inputCover\" accept=\"image/*\"></label>\n  <label class=\"btn\" title=\"Ganti gambar kop di semua halaman\">🏫 Unggah / Ganti Kop<input type=\"file\" id=\"inputKop\" accept=\"image/*\"></label>\n  <button class=\"btn\" id=\"btnWord\" type=\"button\">📄 Unduh Word (.docx)</button>\n  <button class=\"btn\" id=\"btnPdf\" type=\"button\">🖨️ Unduh PDF</button>\n  <div class=\"status\" id=\"statusCek\">Memeriksa dokumen…</div>\n</div>\n<div class=\"toast no-print\" id=\"toast\"></div>\n\n";
const HALAMAN = "\n\n<!-- ======================= HALAMAN 1 · COVER ======================= -->\n<section class=\"page-a4\" id=\"page-cover\" contenteditable=\"true\">\n  <div class=\"cv\">\n    <div class=\"cv-band\"></div>\n    <div class=\"cv-top\"><div class=\"l1\">KEMENTERIAN AGAMA REPUBLIK INDONESIA</div><div class=\"l2\" data-bind=\"madrasahKapital\"></div></div>\n    <div class=\"cv-logo\" contenteditable=\"false\"><img class=\"logo-image\" alt=\"Logo MA Syubbanul Wathon\"></div>\n    <div class=\"cv-title\">\n      <div class=\"t1\">PERANGKAT INSTRUMEN</div>\n      <div class=\"t2\">UJIAN PRAKTIK</div>\n      <div class=\"t3\">ASESMEN MADRASAH (AM)</div>\n      <div class=\"t4\">BENTUK UJIAN: <span data-bind=\"bentuk\"></span></div>\n    </div>\n    <div class=\"cv-box\">\n      <table>\n        <tr><td class=\"lbl\">Mata Pelajaran</td><td class=\"sep\">:</td><td><b data-bind=\"mapel\"></b></td></tr>\n        <tr><td class=\"lbl\">Kelas / Fase</td><td class=\"sep\">:</td><td>IX / D</td></tr>\n        <tr><td class=\"lbl\">Rombel</td><td class=\"sep\">:</td><td data-bind=\"peminatanTeks\"></td></tr>\n        <tr><td class=\"lbl\">Tahun Ajaran</td><td class=\"sep\">:</td><td data-bind=\"ta\"></td></tr>\n        <tr><td class=\"lbl\">Disusun oleh</td><td class=\"sep\">:</td><td><b class=\"nama-ttd\" style=\"text-decoration:none\" data-bind=\"guru\"></b><br><span class=\"nipy-ttd\" data-bind=\"nipyGuru\"></span></td></tr>\n      </table>\n    </div>\n    <div class=\"cv-foot\"><div class=\"f1\" data-bind=\"madrasahKapital\"></div><div class=\"f2\">Terakreditasi <span data-bind=\"akreditasi\"></span> | NSM: <span data-bind=\"nsm\"></span> | NPSN: <span data-bind=\"npsn\"></span></div><div class=\"f2\" data-bind=\"alamat\"></div></div>\n  </div>\n</section>\n\n<!-- ======================= HALAMAN 2 · LEMBAR VERIFIKASI ======================= -->\n<section class=\"page-a4 hal-verif\" id=\"hal-2\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">LEMBAR VERIFIKASI INSTRUMEN UJIAN PRAKTIK</h1>\n  <div class=\"subjudul\">ASESMEN MADRASAH TAHUN AJARAN <span data-bind=\"ta\"></span></div>\n  <table class=\"info\">\n    <tr><td class=\"lbl\">Mata Pelajaran</td><td class=\"sep\">:</td><td data-bind=\"mapel\"></td></tr>\n    <tr><td class=\"lbl\">Bentuk Ujian Praktik</td><td class=\"sep\">:</td><td data-bind=\"bentuk\"></td></tr>\n    <tr><td class=\"lbl\">Kelas / Rombel</td><td class=\"sep\">:</td><td data-bind=\"peminatanTeks\"></td></tr>\n    <tr><td class=\"lbl\">Guru Penyusun</td><td class=\"sep\">:</td><td data-bind=\"guru\"></td></tr>\n  </table>\n  <p style=\"margin-top:2mm\"><i>Petunjuk: validator memberi tanda centang (✓) pada kolom Sesuai / Valid atau Perlu Revisi dan menuliskan catatan bila diperlukan.</i></p>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:6%\"><col style=\"width:52%\"><col style=\"width:10%\"><col style=\"width:10%\"><col style=\"width:22%\"></colgroup>\n    <tr><th>No</th><th>Aspek yang Diverifikasi</th><th>Sesuai / Valid</th><th>Perlu Revisi</th><th>Catatan Validator</th></tr>\n    <tr><td class=\"c\">1</td><td>Kesesuaian instrumen dengan regulasi: <span data-bind=\"regulasi\"></span></td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">2</td><td>Kesesuaian CP, TP, dan ATP Fase D dengan dokumen resmi</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">3</td><td>Tingkat kesulitan berada pada level C3–C6 (bebas C1–C2)</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">4</td><td>Kesesuaian kisi-kisi dengan bentuk ujian (<span data-bind=\"bentuk\"></span>)</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">5</td><td>LKPD memuat petunjuk <span data-bind=\"k3Label\"></span> yang jelas</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">6</td><td>Kartu soal memuat stimulus kontekstual, instruksi tugas, dan kriteria keberhasilan</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">7</td><td>Objektivitas rubrik analitik skala 1–4, termasuk aspek sikap dan nilai cinta</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">8</td><td>Ketepatan pedoman konversi nilai (KKTP 76)</td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">9</td><td>Integrasi <span data-bind=\"integrasiLabel\"></span></td><td></td><td></td><td></td></tr>\n    <tr><td class=\"c\">10</td><td>Kelengkapan 9 komponen perangkat dalam format A4</td><td></td><td></td><td></td></tr>\n  </table>\n  <h3>Keputusan Verifikasi:</h3>\n  <div class=\"kotak-kosong\"></div>\n  <table class=\"ttd\">\n    <tr><td>Guru Penyusun,</td><td><span data-bind=\"tanggal\"></span><br>Kepala Madrasah selaku Validator,</td></tr>\n    <tr><td class=\"ruang\"></td><td class=\"ruang\"></td></tr>\n    <tr><td><span class=\"nama-ttd\" data-bind=\"guru\"></span><br><span class=\"nipy-ttd\" data-bind=\"nipyGuru\"></span></td>\n        <td><span class=\"nama-ttd\" data-bind=\"kepala\"></span><br><span class=\"nipy-ttd\" data-bind=\"nipyKepala\"></span></td></tr>\n  </table>\n</section>\n\n<!-- ======================= HALAMAN 3 · IDENTITAS, CP, TP, ATP ======================= -->\n<section class=\"page-a4\" id=\"hal-3\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">IDENTITAS, CAPAIAN PEMBELAJARAN, TUJUAN PEMBELAJARAN, DAN ALUR TUJUAN PEMBELAJARAN</h1>\n  <div class=\"subjudul\">FASE D · TAHUN AJARAN <span data-bind=\"ta\"></span></div>\n  <h3>A. Identitas Mata Pelajaran</h3>\n  <table class=\"info info2\">\n    <colgroup><col style=\"width:21%\"><col style=\"width:2%\"><col style=\"width:30%\"><col style=\"width:17%\"><col style=\"width:2%\"><col style=\"width:28%\"></colgroup>\n    <tr><td>Mata Pelajaran</td><td>:</td><td data-bind=\"mapel\"></td><td>Fase / Kelas</td><td>:</td><td>D / IX</td></tr>\n    <tr><td>Bentuk Ujian</td><td>:</td><td data-bind=\"bentuk\"></td><td>Alokasi Waktu</td><td>:</td><td>%%alokasi%%</td></tr>\n    <tr><td>Guru Penyusun</td><td>:</td><td colspan=\"4\" data-bind=\"guru\"></td></tr>\n    <tr><td>Rombel</td><td>:</td><td colspan=\"4\" data-bind=\"peminatanTeks\"></td></tr>\n    <tr><td>Materi Pokok</td><td>:</td><td colspan=\"4\" data-bind=\"materi\"></td></tr>\n    <tr><td>Tema KBC</td><td>:</td><td colspan=\"4\" data-bind=\"temaKBC\"></td></tr>\n    <tr><td>Materi Insersi</td><td>:</td><td colspan=\"4\" data-bind=\"materiInsersi\"></td></tr>\n  </table>\n  <h3>B. Capaian Pembelajaran (CP) Fase D</h3>\n  <p style=\"font-size:10.5pt\"><i>Sumber: <span data-bind=\"sumberCP\"></span></i></p>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:21%\"><col style=\"width:67%\"><col style=\"width:12%\"></colgroup>\n    <tr><th>Elemen</th><th>Capaian Pembelajaran (kutipan resmi)</th><th>Tema KBC</th></tr>\n    %%cpRows%%\n  </table>\n  <p class=\"ket-kbc\"><i>Tema KBC: 1 Cinta Allah dan Rasul-Nya · 2 Cinta Ilmu · 3 Cinta Lingkungan · 4 Cinta Diri dan Sesama · 5 Cinta Tanah Air (pola Lampiran 1 SK Dirjen Pendis 6077/2025).</i></p>\n  <h3>C. Tujuan Pembelajaran (TP)</h3>\n  <ol>%%tpLis%%</ol>\n  <h3>D. Alur Tujuan Pembelajaran (ATP)</h3>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:8%\"><col style=\"width:44%\"><col style=\"width:28%\"><col style=\"width:20%\"></colgroup>\n    <tr><th>No</th><th>Tujuan Pembelajaran</th><th>Materi / Lingkup</th><th>Alokasi Waktu</th></tr>\n    %%atpRows%%\n  </table>\n</section>\n\n<!-- ======================= HALAMAN 4 · KISI-KISI ======================= -->\n<section class=\"page-a4\" id=\"hal-4\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">KISI-KISI UJIAN PRAKTIK</h1>\n  <div class=\"subjudul\">ASESMEN MADRASAH TAHUN AJARAN <span data-bind=\"ta\"></span></div>\n  <table class=\"info\">\n    <tr><td class=\"lbl\">Mata Pelajaran</td><td class=\"sep\">:</td><td data-bind=\"mapel\"></td></tr>\n    <tr><td class=\"lbl\">Kelas / Rombel</td><td class=\"sep\">:</td><td data-bind=\"peminatanTeks\"></td></tr>\n    <tr><td class=\"lbl\">Bentuk Ujian Praktik</td><td class=\"sep\">:</td><td data-bind=\"bentuk\"></td></tr>\n    <tr><td class=\"lbl\">Materi Pokok</td><td class=\"sep\">:</td><td data-bind=\"materi\"></td></tr>\n    <tr><td class=\"lbl\">Tema KBC (Panca Cinta)</td><td class=\"sep\">:</td><td data-bind=\"temaKBC\"></td></tr>\n  </table>\n  <br>\n  <table class=\"grid\" id=\"tabel-kisi\">\n    <colgroup><col style=\"width:5%\"><col style=\"width:14%\"><col style=\"width:22%\"><col style=\"width:14%\"><col style=\"width:23%\"><col style=\"width:9%\"><col style=\"width:13%\"></colgroup>\n    <tr><th>No</th><th>Elemen</th><th>Tujuan Pembelajaran</th><th>Materi</th><th>Indikator Ujian Praktik</th><th>Level Kognitif</th><th>Topik Panca Cinta</th></tr>\n    %%kisiRows%%\n  </table>\n  <p style=\"font-size:10pt;margin-top:2mm\"><i>Keterangan level kognitif: C3 Mengaplikasikan, C4 Menganalisis, C5 Mengevaluasi, C6 Mencipta. Topik Panca Cinta mengacu SK Dirjen Pendis Nomor 6077 Tahun 2025.</i></p>\n  <table class=\"ttd\">\n    <tr><td>Mengetahui,<br>Kepala Madrasah</td><td><span data-bind=\"tanggal\"></span><br>Guru Penyusun</td></tr>\n    <tr><td class=\"ruang\"></td><td class=\"ruang\"></td></tr>\n    <tr><td><span class=\"nama-ttd\" data-bind=\"kepala\"></span><br><span class=\"nipy-ttd\" data-bind=\"nipyKepala\"></span></td>\n        <td><span class=\"nama-ttd\" data-bind=\"guru\"></span><br><span class=\"nipy-ttd\" data-bind=\"nipyGuru\"></span></td></tr>\n  </table>\n</section>\n\n<!-- ======================= HALAMAN 5 · LKPD ======================= -->\n<section class=\"page-a4\" id=\"hal-5\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">LEMBAR KERJA PESERTA DIDIK (LKPD)</h1>\n  <div class=\"subjudul\">UJIAN PRAKTIK <span data-bind=\"mapelKapital\"></span> · BENTUK: <span data-bind=\"bentuk\"></span></div>\n  <table class=\"info\">\n    <tr><td class=\"lbl\">Nama Peserta Didik</td><td class=\"sep\">:</td><td>.............................................................</td></tr>\n    <tr><td class=\"lbl\">NIS / NISN</td><td class=\"sep\">:</td><td>.............................................................</td></tr>\n    <tr><td class=\"lbl\">Kelas</td><td class=\"sep\">:</td><td>.............................................................</td></tr>\n    <tr><td class=\"lbl\">Alokasi Waktu</td><td class=\"sep\">:</td><td>%%alokasi%%</td></tr>\n    <tr><td class=\"lbl\">Tema KBC</td><td class=\"sep\">:</td><td data-bind=\"temaKBC\"></td></tr>\n  </table>\n  <h3>A. Tujuan</h3>\n  <p>%%lkpd.tujuan%%</p>\n  <h3>B. Alat dan Bahan</h3>\n  <ol>%%lkpd.alat%%</ol>\n  <h3>C. Petunjuk <span data-bind=\"k3Label\"></span></h3>\n  <ol>%%lkpd.petunjuk%%</ol>\n  <h3>D. Langkah Kerja (alur %%lkpd.alur%%)</h3>\n  <ol>%%lkpd.langkah%%</ol>\n  <h3>E. Lembar Hasil Kerja</h3>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:8%\"><col style=\"width:42%\"><col style=\"width:50%\"></colgroup>\n    <tr><th>No</th><th>%%lkpd.kolom%%</th><th>Hasil / Catatan Peserta Didik</th></tr>\n    %%lkpd.item%%\n  </table>\n  <h3>F. Refleksi Diri (Nilai Cinta)</h3>\n  <ol class=\"refleksi\">%%lkpd.refleksi%%</ol>\n</section>\n\n<!-- ======================= HALAMAN 6 · KARTU SOAL ======================= -->\n<section class=\"page-a4\" id=\"hal-6\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">KARTU SOAL UJIAN PRAKTIK</h1>\n  <div class=\"subjudul\">ASESMEN MADRASAH TAHUN AJARAN <span data-bind=\"ta\"></span></div>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:30%\"><col style=\"width:70%\"></colgroup>\n    <tr><th colspan=\"2\">Metadata Soal</th></tr>\n    <tr><td>Mata Pelajaran</td><td data-bind=\"mapel\"></td></tr>\n    <tr><td>Kelas / Fase</td><td>IX / D</td></tr>\n    <tr><td>Bentuk Ujian Praktik</td><td data-bind=\"bentuk\"></td></tr>\n    <tr><td>Elemen CP</td><td>%%ks.elemen%%</td></tr>\n    <tr><td>Tujuan Pembelajaran</td><td>%%ks.tp%%</td></tr>\n    <tr><td>Indikator Soal</td><td>%%ks.indikator%%</td></tr>\n    <tr><td>Level Kognitif</td><td>%%ks.level%%</td></tr>\n    <tr><td>Tema KBC / Materi Insersi</td><td><span data-bind=\"temaKBC\"></span> / <span data-bind=\"materiInsersi\"></span></td></tr>\n    <tr><td>Acuan Regulasi</td><td data-bind=\"regulasi\"></td></tr>\n    <tr><th colspan=\"2\">Stimulus</th></tr>\n    <tr><td colspan=\"2\">%%ks.stimulus%%</td></tr>\n    <tr><th colspan=\"2\">Instruksi Tugas Praktik</th></tr>\n    <tr><td colspan=\"2\"><ol>%%ks.instruksi%%</ol></td></tr>\n    <tr><th colspan=\"2\">Garis Besar Penyelesaian</th></tr>\n    <tr><td colspan=\"2\"><ol>%%ks.garis%%</ol></td></tr>\n    <tr><th colspan=\"2\">Kriteria Keberhasilan</th></tr>\n    <tr><td colspan=\"2\"><ol>%%ks.kriteria%%</ol></td></tr>\n  </table>\n  <table class=\"ttd\">\n    <tr><td>Mengetahui,<br>Kepala Madrasah</td><td><span data-bind=\"tanggal\"></span><br>Guru Penyusun</td></tr>\n    <tr><td class=\"ruang\"></td><td class=\"ruang\"></td></tr>\n    <tr><td><span class=\"nama-ttd\" data-bind=\"kepala\"></span><br><span class=\"nipy-ttd\" data-bind=\"nipyKepala\"></span></td>\n        <td><span class=\"nama-ttd\" data-bind=\"guru\"></span><br><span class=\"nipy-ttd\" data-bind=\"nipyGuru\"></span></td></tr>\n  </table>\n</section>\n\n<!-- ======================= HALAMAN 7 · RUBRIK ======================= -->\n<section class=\"page-a4\" id=\"hal-7\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">RUBRIK PENILAIAN ANALITIK UJIAN PRAKTIK</h1>\n  <div class=\"subjudul\"><span data-bind=\"mapelKapital\"></span> · BENTUK: <span data-bind=\"bentuk\"></span></div>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:16%\"><col style=\"width:21%\"><col style=\"width:21%\"><col style=\"width:21%\"><col style=\"width:21%\"></colgroup>\n    <tr><th>Aspek</th><th>4<br>(Sangat Baik)</th><th>3<br>(Baik)</th><th>2<br>(Cukup)</th><th>1<br>(Perlu Bimbingan)</th></tr>\n    %%rubrikRows%%\n  </table>\n  <p style=\"margin-top:3mm\"><b>Skor maksimal</b> = 4 aspek × 4 = <b>16</b>. Skor minimal = 4.</p>\n  <p style=\"font-size:10.5pt\"><i>Aspek <b data-bind=\"aspek4\"></b> dinilai melalui observasi selama ujian berdasarkan Tema KBC: <span data-bind=\"temaKBC\"></span> (evaluasi berbasis proses, SK Dirjen Pendis Nomor 6077 Tahun 2025).</i></p>\n</section>\n\n<!-- ======================= HALAMAN 8 · PEDOMAN KONVERSI NILAI (tetap) ======================= -->\n<section class=\"page-a4\" id=\"hal-8\" contenteditable=\"true\">\n  <div class=\"kop-slot\" contenteditable=\"false\"></div>\n  <h1 class=\"judul\">PEDOMAN KONVERSI NILAI AKHIR</h1>\n  <div class=\"subjudul\">UJIAN PRAKTIK <span data-bind=\"mapelKapital\"></span></div>\n  <h3>A. Rumus Nilai Akhir</h3>\n  <p style=\"text-align:center;font-size:13pt\"><b>Nilai Akhir = (Total Skor Perolehan ÷ 16) × 100</b></p>\n  <p>Nilai akhir dibulatkan ke bilangan bulat terdekat (0,5 dibulatkan ke atas) sebelum menentukan predikat. Karena skor setiap aspek minimal 1, rentang nilai akhir adalah 25–100. Kriteria Ketercapaian Tujuan Pembelajaran (<b>KKTP</b>) = <b>76</b>.</p>\n  <h3>B. Predikat dan Tindak Lanjut</h3>\n  <table class=\"grid\">\n    <colgroup><col style=\"width:15%\"><col style=\"width:22%\"><col style=\"width:23%\"><col style=\"width:40%\"></colgroup>\n    <tr><th>Predikat</th><th>Rentang Nilai</th><th>Keterangan</th><th>Tindak Lanjut</th></tr>\n    <tr><td class=\"c\"><b>A</b></td><td class=\"c\">91 – 100</td><td>Sangat Baik</td><td>Pengayaan (tuntas, nilai ≥ 76)</td></tr>\n    <tr><td class=\"c\"><b>B</b></td><td class=\"c\">83 – 90</td><td>Baik</td><td>Pengayaan (tuntas, nilai ≥ 76)</td></tr>\n    <tr><td class=\"c\"><b>C</b></td><td class=\"c\">76 – 82</td><td>Cukup</td><td>Pengayaan (tuntas, nilai ≥ 76)</td></tr>\n    <tr><td class=\"c\"><b>D</b></td><td class=\"c\">&lt; 76</td><td>Perlu Bimbingan</td><td>Remedial (belum tuntas, nilai &lt; 76)</td></tr>\n  </table>\n  <h3>C. Tabel Konversi Total Skor</h3>\n  <table class=\"grid\">\n    <tr><th>Total Skor</th><th>4</th><th>5</th><th>6</th><th>7</th><th>8</th><th>9</th><th>10</th><th>11</th><th>12</th><th>13</th><th>14</th><th>15</th><th>16</th></tr>\n    <tr><td class=\"c\"><b>Nilai Akhir</b></td><td class=\"c\">25</td><td class=\"c\">31</td><td class=\"c\">38</td><td class=\"c\">44</td><td class=\"c\">50</td><td class=\"c\">56</td><td class=\"c\">63</td><td class=\"c\">69</td><td class=\"c\">75</td><td class=\"c\">81</td><td class=\"c\">88</td><td class=\"c\">94</td><td class=\"c\">100</td></tr>\n    <tr><td class=\"c\"><b>Predikat</b></td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">D</td><td class=\"c\">C</td><td class=\"c\">B</td><td class=\"c\">A</td><td class=\"c\">A</td></tr>\n  </table>\n  <h3>D. Ketentuan Tindak Lanjut</h3>\n  <ol>\n    <li><b>Remedial</b> (nilai &lt; 76): pembimbingan ulang pada aspek yang skornya rendah, lalu ujian praktik ulang pada aspek tersebut. Nilai hasil remedial paling tinggi sama dengan KKTP (76).</li>\n    <li><b>Pengayaan</b> (nilai ≥ 76): tugas pengembangan/pendalaman yang menantang (level C5–C6) tanpa mengubah nilai yang telah dicapai.</li>\n  </ol>\n  <p><i>Contoh:</i> skor aspek 4 + 3 + 3 + 4 = 14 → (14 ÷ 16) × 100 = 87,5 → dibulatkan <b>88</b> → predikat <b>B</b> → pengayaan.</p>\n</section>\n\n\n";
let KOP_SRC = "", KOP_DEFAULT = "", LOGO_SRC = "", COVER_SRC = "", KONFIG = {}, ISI = {}, DATA_SISWA = { rombel: [] };
const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
/* ====================== BAGIAN TERUJI DARI TEMPLATE v2.1 ====================== */
const esc = s => String(s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
function nipyGuru() { const n = String(KONFIG.guruNipy || "").trim(); return !n || n === "-" ? "NIPY. -" : (/^nipy/i.test(n) ? n : "NIPY. " + n); }
function binds() {
  return { ...IDENTITAS, guru: KONFIG.guruNama, nipyGuru: nipyGuru(), mapel: KONFIG.mapel, mapelKapital: String(KONFIG.mapel).toUpperCase(),
    bentuk: String(KONFIG.bentukUjian).toUpperCase(), peminatanTeks: teksPeminatan(), materi: KONFIG.materiPokok, ta: KONFIG.tahunAjaran,
    tanggal: IDENTITAS.kota + ", " + KONFIG.tanggalDokumen, regulasi: KONFIG.regulasi, sumberCP: KONFIG.sumberCP,
    k3Label: KONFIG.k3Label, integrasiLabel: KONFIG.integrasiLabel,
    temaKBC: [].concat(KONFIG.temaKBC || []).join("; "), materiInsersi: KONFIG.materiInsersi,
    aspek1: KONFIG.aspekRubrik[0], aspek2: KONFIG.aspekRubrik[1], aspek3: KONFIG.aspekRubrik[2], aspek4: KONFIG.aspekRubrik[3] };
}
function isiBind(root = document) { const b = binds(); $$("[data-bind]", root).forEach(el => { if (b[el.dataset.bind] !== undefined) el.textContent = b[el.dataset.bind]; }); }
function pasNamaTtd() {
  $$(".nama-ttd, .nipy-ttd").forEach(el => {
    const td = el.closest("td"); if (!td || !td.closest("table.ttd")) return;
    el.style.fontSize = ""; const cs = getComputedStyle(td);
    const ruang = td.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight) - 2;
    let f = parseFloat(getComputedStyle(el).fontSize) * 0.75;
    while (el.scrollWidth > ruang && f > 9) { f -= 0.25; el.style.fontSize = f + "pt"; }
  });
}
function pasangKop(root = document) {
  $$(".kop-slot", root).forEach(slot => {
    let img = $(".kop-image", slot);
    if (!img) { img = document.createElement("img"); img.className = "kop-image"; img.alt = "Kop Madrasah"; img.setAttribute("width", "642");
      img.style.cssText = "width:100%; height:auto; display:block; margin-bottom:8px;"; slot.appendChild(img); }
    img.src = KOP_SRC;
  });
}
function ttdHTML(kiriJudul, kananJudul) {
  return `<table class="ttd"><tr><td>${kiriJudul}</td><td><span data-bind="tanggal"></span><br>${kananJudul}</td></tr>
  <tr><td class="ruang"></td><td class="ruang"></td></tr>
  <tr><td><span class="nama-ttd" data-bind="kepala"></span><br><span class="nipy-ttd" data-bind="nipyKepala"></span></td>
  <td><span class="nama-ttd" data-bind="guru"></span><br><span class="nipy-ttd" data-bind="nipyGuru"></span></td></tr></table>`;
}
function buatDaftarNilai() {
  $$(".page-dn").forEach(e => e.remove());
  const doc = $("#doc"), A = KONFIG.aspekSingkat.map(s => String(s).trim().slice(0, 8));
  rombelTerpilih().forEach((r, idx) => {
    const sec = document.createElement("section");
    sec.className = "page-a4 page-dn"; sec.id = "hal-dn-" + (idx + 1); sec.contentEditable = "true";
    const baris = r.siswa.map((s, i) => `<tr><td class="skor">${i + 1}</td><td class="skor">${esc(s.nis)}</td><td class="skor">${esc(s.nisn)}</td><td class="nama">${esc(s.nama)}</td>` +
      `<td class="skor in"></td><td class="skor in"></td><td class="skor in"></td><td class="skor in"></td><td class="skor tot"></td><td class="skor na"></td><td class="skor pr"></td></tr>`).join("");
    sec.innerHTML = `<div class="kop-slot" contenteditable="false"></div>
      <p class="dn-judul">DAFTAR NILAI UJIAN PRAKTIK KELAS IX</p>
      <p class="dn-judul">ASESMEN MADRASAH TAHUN AJARAN <span data-bind="ta"></span></p>
      <table class="dn-info" style="margin:1.5mm 0 1mm 0"><colgroup><col style="width:17%"><col style="width:2%"><col style="width:36%"><col style="width:15%"><col style="width:2%"><col style="width:28%"></colgroup>
        <tr><td>Mata Pelajaran</td><td>:</td><td data-bind="mapel"></td><td>Kelas</td><td>:</td><td>${esc(r.rombel)}</td></tr>
        <tr><td>Bentuk Ujian</td><td>:</td><td data-bind="bentuk"></td><td>KKTP</td><td>:</td><td>76</td></tr>
        <tr><td>Guru Penguji</td><td>:</td><td data-bind="guru"></td><td>Jumlah Siswa</td><td>:</td><td>${r.siswa.length} peserta didik</td></tr></table>
      <table class="tbl-nilai"><colgroup><col style="width:4.5%"><col style="width:7.5%"><col style="width:11%"><col style="width:34%"><col style="width:6.5%"><col style="width:6.5%"><col style="width:6.5%"><col style="width:6.5%"><col style="width:6%"><col style="width:6%"><col style="width:5%"></colgroup>
        <tr><th>No</th><th>NIS</th><th>NISN</th><th>Nama Peserta Didik</th><th>${esc(A[0])}</th><th>${esc(A[1])}</th><th>${esc(A[2])}</th><th>${esc(A[3])}</th><th>Total</th><th>Nilai</th><th>Pred.</th></tr>${baris}</table>
      <p class="dn-ket">Skor tiap aspek 1–4 · Nilai Akhir = (Total Skor ÷ 16) × 100, dibulatkan · Predikat: A 91–100, B 83–90, C 76–82, D &lt; 76 (remedial).</p>
      ${ttdHTML("Mengetahui,<br>Kepala Madrasah", "Guru Penguji")}`;
    doc.appendChild(sec);
  });
  pasangKop(); isiBind();
}
function hitungBaris(tr) {
  const ins = $$("td.in", tr), v = ins.map(td => parseInt(td.textContent.trim(), 10));
  const ok = v.every(x => x >= 1 && x <= 4);
  const tot = $("td.tot", tr), na = $("td.na", tr), pr = $("td.pr", tr);
  if (!ok) { if (v.every(isNaN)) { tot.textContent = na.textContent = pr.textContent = ""; } return; }
  const t = v.reduce((a, b) => a + b, 0), n = Math.round(t / 16 * 100);
  tot.textContent = t; na.textContent = n; pr.textContent = n >= 91 ? "A" : n >= 83 ? "B" : n >= 76 ? "C" : "D";
}
const DX = { tw: px => Math.round(px * 15), emu: px => Math.round(px * 9525),
  x: s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;") };
function dxWarna(c) { const m = String(c).match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/); if (!m || (m[4] !== undefined && +m[4] === 0)) return null;
  return [m[1], m[2], m[3]].map(v => (+v).toString(16).padStart(2, "0")).join("").toUpperCase(); }
function dxBuat() {
  const media = [], petaMedia = new Map(); let rid = 10, did = 1;
  const gambar = (src, wpx, hpx) => {
    let r = petaMedia.get(src);
    if (!r) { const m = src.match(/^data:image\/(png|jpe?g|gif);base64,(.*)$/); if (!m) return "";
      r = "rId" + (++rid); const ext = m[1] === "jpeg" ? "jpg" : m[1]; media.push({ rid: r, nama: "media/img" + rid + "." + ext, b64: m[2], ext }); petaMedia.set(src, r); }
    const cx = DX.emu(wpx), cy = DX.emu(hpx), id = ++did;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Gambar ${id}"/>` +
      `<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="img${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${r}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  };
  const rPr = el => { const s = getComputedStyle(el), sz = Math.round(parseFloat(s.fontSize) * 1.5), w = dxWarna(s.color);
    const rtl = s.direction === "rtl";
    return `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="${rtl ? "Traditional Arabic" : "Times New Roman"}"/>` +
      (parseInt(s.fontWeight) >= 600 ? "<w:b/><w:bCs/>" : "") + (s.fontStyle === "italic" ? "<w:i/><w:iCs/>" : "") +
      (w && w !== "000000" && w !== "111111" ? `<w:color w:val="${w}"/>` : "") + `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>` +
      (s.textDecorationLine.includes("underline") ? '<w:u w:val="single"/>' : "") + (rtl ? "<w:rtl/>" : "") + "</w:rPr>"; };
  const BLOK = /^(P|DIV|H1|H2|H3|H4|OL|UL|LI|TABLE|SECTION)$/;
  function runs(node) { let o = "";
    node.childNodes.forEach(n => {
      if (n.nodeType === 3) { const t = n.textContent.replace(/\s+/g, " "); if (t.trim() || t === " ") o += `<w:r>${rPr(n.parentElement)}<w:t xml:space="preserve">${DX.x(t)}</w:t></w:r>`; }
      else if (n.nodeType === 1 && getComputedStyle(n).display !== "none") {
        if (n.tagName === "BR") o += "<w:r><w:br/></w:r>";
        else if (n.tagName === "IMG") { const r = n.getBoundingClientRect(); if (n.src) o += gambar(n.src, r.width, r.height); }
        else if (!BLOK.test(n.tagName)) o += runs(n);
      } });
    return o; }
  function pPr(el, o = {}) { const s = getComputedStyle(el), fs = parseFloat(s.fontSize), lh = parseFloat(s.lineHeight) || fs * 1.2;
    let al = { center: "center", right: "right", end: "right", justify: "both" }[s.textAlign] || "left";
    if (s.direction === "rtl" && (al === "right" || al === "left")) al = al === "right" ? "left" : "right"; // bidi: kiri/kanan bertukar
    const rapat = !!el.closest("table.tbl-nilai"), line = rapat ? Math.round(fs * 0.75 * 20 * 1.12) : Math.max(200, Math.round(240 * (lh / fs) / 1.15));
    return `<w:pPr>${o.tabs || ""}${s.direction === "rtl" ? "<w:bidi/>" : ""}<w:spacing w:before="${DX.tw(parseFloat(s.marginTop) || 0)}" w:after="${DX.tw(parseFloat(s.marginBottom) || 0)}" w:line="${line}" w:lineRule="${rapat ? "exact" : "auto"}"/>` +
      `${o.ind || ""}<w:jc w:val="${al}"/></w:pPr>`; }
  function para(el, isi) { return `<w:p>${pPr(el)}${isi}</w:p>`; }
  function blok(el, ctx = {}) { // mengembalikan XML untuk isi elemen (campuran inline & blok)
    let o = "", buf = [];
    const flush = () => { if (!buf.length) return; let isi = ""; buf.forEach(n => isi += runs({ childNodes: [n] }));
      if (/<w:t xml:space="preserve">[^<]*[^\s<]|<w:drawing>|<w:br\/>/.test(isi)) o += para(el, isi); buf = []; };
    el.childNodes.forEach(n => {
      if (n.nodeType === 1 && getComputedStyle(n).display === "none") return;
      if (n.nodeType === 1 && BLOK.test(n.tagName)) { flush(); o += konversi(n, ctx); }
      else if (n.nodeType === 3 || n.nodeType === 1) buf.push(n);
    });
    flush(); return o; }
  function konversi(el, ctx = {}) {
    const tag = el.tagName;
    if (tag === "TABLE") return tabel(el);
    if (el.classList.contains("kotak-kosong")) { const h = el.getBoundingClientRect().height;
      return `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${["top", "left", "bottom", "right"].map(b => `<w:${b} w:val="single" w:sz="6" w:color="000000"/>`).join("")}</w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="${DX.tw(h)}" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>`; }
    if (tag === "OL" || tag === "UL") { let i = 0, o = ""; const lvl = (ctx.lvl || 0) + 1;
      [...el.children].forEach(li => { if (li.tagName !== "LI") return; i++;
        const pre = `<w:r>${rPr(li)}<w:t xml:space="preserve">${tag === "OL" ? i + "." : "•"}\t</w:t></w:r>`;
        const opsi = { tabs: `<w:tabs><w:tab w:val="left" w:pos="${360 * lvl}"/></w:tabs>`, ind: `<w:ind w:left="${360 * lvl}" w:hanging="360"/>` };
        const sub = [...li.childNodes].filter(n => n.nodeType === 1 && BLOK.test(n.tagName));
        let isi = ""; li.childNodes.forEach(n => { if (!(n.nodeType === 1 && BLOK.test(n.tagName))) isi += runs({ childNodes: [n] }); });
        o += `<w:p>${pPr(li, opsi)}${pre}${isi}</w:p>`;
        sub.forEach(s => o += konversi(s, { lvl })); });
      return o; }
    if (el.classList.contains("kop-slot")) { const im = el.querySelector("img"); if (!im) return "";
      const r = im.getBoundingClientRect(); return `<w:p><w:pPr><w:spacing w:before="0" w:after="${DX.tw(6)}"/><w:jc w:val="center"/></w:pPr>${gambar(im.src, r.width, r.height)}</w:p>`; }
    const adaBlok = [...el.children].some(c => BLOK.test(c.tagName));
    if (!adaBlok) { const isi = runs(el); return isi || tag === "P" ? para(el, isi) : ""; }
    return blok(el, ctx);
  }
  function tabel(t) {
    const s = getComputedStyle(t), bergaris = t.classList.contains("grid") || t.classList.contains("tbl-nilai");
    const rows = [...t.rows]; let acuan = rows.reduce((a, r) => (r.cells.length > a.cells.length && ![...r.cells].some(c => c.colSpan > 1)) ? r : a, rows[0]);
    const lebarTotal = t.getBoundingClientRect().width;
    const kol = [...acuan.cells].map(c => DX.tw(c.getBoundingClientRect().width));

    const garis = bergaris ? `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(b => `<w:${b} w:val="single" w:sz="6" w:color="000000"/>`).join("")}</w:tblBorders>` : "";
    const mar = t.classList.contains("tbl-nilai") ? 40 : bergaris ? 70 : 30;
    let o = `<w:tbl><w:tblPr><w:tblW w:w="${kol.reduce((a, b) => a + b, 0)}" w:type="dxa"/>${t.style.marginLeft === "auto" ? '<w:jc w:val="center"/>' : ""}${garis}<w:tblLayout w:type="fixed"/>` +
      `<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="${mar}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="${mar}" w:type="dxa"/></w:tblCellMar></w:tblPr>` +
      `<w:tblGrid>${kol.map(w => `<w:gridCol w:w="${w}"/>`).join("")}</w:tblGrid>`;
    rows.forEach(r => { const h = r.getBoundingClientRect().height;
      o += `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="${Math.round(DX.tw(h) * (t.classList.contains("tbl-nilai") ? 0.8 : 0.92))}" w:hRule="atLeast"/></w:trPr>`;
      let gi = 0;
      [...r.cells].forEach(c => { const span = c.colSpan || 1, w = kol.slice(gi, gi + span).reduce((a, b) => a + b, 0) || DX.tw(c.getBoundingClientRect().width); gi += span;
        const cs = getComputedStyle(c), bg = dxWarna(cs.backgroundColor), va = { middle: "center", bottom: "bottom" }[cs.verticalAlign] || "top";
        let isi = konversi(c); if (!isi || /<\/w:tbl>$/.test(isi)) isi += `<w:p>${pPr(c)}</w:p>`;
        o += `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ""}${bg && bg !== "FFFFFF" ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : ""}<w:vAlign w:val="${va}"/></w:tcPr>${isi}</w:tc>`; });
      o += "</w:tr>"; });
    return o + "</w:tbl>";
  }
  const sect = (m, akhir) => `<w:sectPr>${akhir ? "" : '<w:type w:val="nextPage"/>'}<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="${m}" w:right="${m}" w:bottom="${m}" w:left="${m}" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr>`;
  // ---- Cover ----
  let body = "", cv = $("#page-cover");
  if (cv.classList.contains("full-bleed-cover") && COVER_SRC) {
    body += `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:jc w:val="center"/>${sect(0, false)}</w:pPr>${gambar(COVER_SRC, 793.7, 1118)}</w:p>`;
  } else {
    const b = binds(), box = document.createElement("div"); box.className = "page-a4"; box.style.cssText = "position:absolute;left:-9999px;top:0;padding:0;width:170mm;min-height:0";
    box.innerHTML = `<table style="width:100%;border:3pt solid #0B6B3A"><tr><td style="padding:10mm 8mm;text-align:center">
      <p style="text-align:center;font-size:13pt">KEMENTERIAN AGAMA REPUBLIK INDONESIA</p><p style="text-align:center;font-size:18pt;font-weight:700;color:#0B6B3A">${esc(b.madrasahKapital)}</p>
      <p style="text-align:center;margin:8mm 0">${LOGO_SRC ? `<img src="${LOGO_SRC}" style="width:40mm;height:40mm">` : ""}</p>
      <p style="text-align:center;font-size:15pt">PERANGKAT INSTRUMEN</p><p style="text-align:center;font-size:30pt;font-weight:700;color:#0B6B3A;margin:2mm 0">UJIAN PRAKTIK</p>
      <p style="text-align:center;font-size:15pt;font-weight:700">ASESMEN MADRASAH (AM)</p><p style="text-align:center;font-size:13pt;font-weight:700;margin:4mm 0 10mm">BENTUK UJIAN: ${esc(b.bentuk)}</p>
      <table style="width:82%;margin-left:auto;margin-right:auto"><colgroup><col style="width:36%"><col style="width:4%"><col style="width:60%"></colgroup>
        <tr><td>Mata Pelajaran</td><td>:</td><td><b>${esc(b.mapel)}</b></td></tr><tr><td>Kelas / Fase</td><td>:</td><td>IX / D</td></tr>
        <tr><td>Rombel</td><td>:</td><td>${esc(b.peminatanTeks)}</td></tr><tr><td>Tahun Ajaran</td><td>:</td><td>${esc(b.ta)}</td></tr>
        <tr><td>Disusun oleh</td><td>:</td><td><b>${esc(b.guru)}</b><br>${esc(b.nipyGuru)}</td></tr></table>
      <p style="text-align:center;font-size:15pt;font-weight:700;color:#0B6B3A;margin-top:16mm">${esc(b.madrasahKapital)}</p>
      <p style="text-align:center;font-size:11pt">Terakreditasi ${esc(b.akreditasi)} | NSM: ${esc(b.nsm)} | NPSN: ${esc(b.npsn)}</p><p style="text-align:center;font-size:11pt">${esc(b.alamat)}</p></td></tr></table>`;
    document.body.appendChild(box); const tb = box.querySelector("table");
    body += tabel(tb).replace('<w:tblLayout', `<w:tblBorders>${["top", "left", "bottom", "right"].map(x => `<w:${x} w:val="double" w:sz="12" w:color="0B6B3A"/>`).join("")}</w:tblBorders><w:tblLayout`);
    body += `<w:p><w:pPr>${sect(1134, false)}</w:pPr></w:p>`; box.remove();
  }
  // ---- Halaman 2 dst. (margin 2 cm) ----
  const hal = $$(".page-a4").filter(s => s.id !== "page-cover");
  hal.forEach((sec, i) => { let x = blok(sec);
    if (i > 0) x = x.replace(/^<w:p><w:pPr>/, '<w:p><w:pPr><w:pageBreakBefore/>').replace(/^<w:tbl>/, '<w:p><w:pPr><w:pageBreakBefore/><w:spacing w:after="0"/></w:pPr></w:p><w:tbl>');
    body += x; });
  body += sect(1134, true);
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"';
  const files = {
    "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${media.map(m => `<Relationship Id="${m.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${m.nama}"/>`).join("")}</Relationships>`,
    "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="id-ID"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`,
    "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}</w:body></w:document>`
  };
  const enc = new TextEncoder(), daftar = Object.entries(files).map(([n, s]) => ({ n, d: enc.encode(s) }));
  media.forEach(m => { const bin = atob(m.b64), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); daftar.push({ n: "word/" + m.nama, d: u }); });
  return zipStore(daftar);
}
function zipStore(files) {
  const T = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; T[n] = c >>> 0; }
  const crc = d => { let c = 0xFFFFFFFF; for (let i = 0; i < d.length; i++) c = T[(c ^ d[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const enc = new TextEncoder(), lokal = [], pusat = []; let off = 0;
  files.forEach(f => { const nm = enc.encode(f.n), c = crc(f.d), L = new DataView(new ArrayBuffer(30));
    L.setUint32(0, 0x04034b50, true); L.setUint16(4, 20, true); L.setUint32(14, c, true); L.setUint32(18, f.d.length, true); L.setUint32(22, f.d.length, true); L.setUint16(26, nm.length, true);
    lokal.push(new Uint8Array(L.buffer), nm, f.d);
    const C = new DataView(new ArrayBuffer(46)); C.setUint32(0, 0x02014b50, true); C.setUint16(4, 20, true); C.setUint16(6, 20, true); C.setUint32(16, c, true);
    C.setUint32(20, f.d.length, true); C.setUint32(24, f.d.length, true); C.setUint16(28, nm.length, true); C.setUint32(42, off, true);
    pusat.push(new Uint8Array(C.buffer), nm); off += 30 + nm.length + f.d.length; });
  const ukPusat = pusat.reduce((a, b) => a + b.length, 0), E = new DataView(new ArrayBuffer(22));
  E.setUint32(0, 0x06054b50, true); E.setUint16(8, files.length, true); E.setUint16(10, files.length, true); E.setUint32(12, ukPusat, true); E.setUint32(16, off, true);
  return new Blob([...lokal, ...pusat, new Uint8Array(E.buffer)], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}
function unduhWord() {
  document.activeElement && document.activeElement.blur();
  try { const blob = dxBuat(), a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = namaFile("docx");
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 3000);
    toast("File Word (.docx) diunduh: A4, margin 2-2-2-2 cm, kop & cover ikut tersimpan. Bisa dibuka di MS Word, WPS, Google Docs, dan LibreOffice.", 8000);
  } catch (err) { toast("Gagal membuat file Word: " + err.message); console.error(err); }
}
/* ====================== BAGIAN BARU v2.2 ====================== */
function siapkanData() {
  const D = window.INSTRUMEN || {};
  const k = D.konfig || {}, prof = PROFIL[k.jenis] || PROFIL.UMUM;
  KONFIG = Object.assign({ jenis: "UMUM", tahunAjaran: "2026/2027", tanggalDokumen: "2 Januari 2027", alokasiWaktu: "2 × 45 menit" }, k);
  KONFIG.regulasi = prof.regulasi; KONFIG.sumberCP = prof.sumberCP; KONFIG.k3Label = prof.k3Label; KONFIG.integrasiLabel = prof.integrasiLabel;
  KONFIG.aspekRubrik = [].concat(k.aspekRubrik || prof.aspekRubrik).slice(0, 4); while (KONFIG.aspekRubrik.length < 4) KONFIG.aspekRubrik.push("{{aspek belum diisi}}");
  KONFIG.aspekSingkat = [].concat(k.aspekSingkat || prof.aspekSingkat).slice(0, 4); while (KONFIG.aspekSingkat.length < 4) KONFIG.aspekSingkat.push("-");
  KONFIG.temaKBC = [].concat(k.temaKBC || []);
  ISI = D.isi || {};
  const rs = (D.siswa && D.siswa.rombel) || [];
  DATA_SISWA = { rombel: rs.map(r => ({ rombel: normRombel(r.rombel), peminatan: r.peminatan || (REF_SISWA[normRombel(r.rombel)] || {}).peminatan || "",
    siswa: (r.siswa || []).map(s => Array.isArray(s) ? { nis: String(s[0]), nisn: String(s[1]), nama: String(s[2]) } : { nis: String(s.nis), nisn: String(s.nisn), nama: String(s.nama) }) })) };
}
const T = (v, nama) => (v === undefined || v === null || v === "") ? `{{belum diisi: ${nama}}}` : String(v);
const daftar = (a, nama) => { a = [].concat(a || []); return a.length ? a.map(x => `<li>${T(x, nama)}</li>`).join("") : `<li>{{belum diisi: ${nama}}}</li>`; };
function isiSlot(html) {
  const L = ISI.lkpd || {}, K = ISI.kartuSoal || {};
  const cp = [].concat(ISI.cp || [{}]), atp = [].concat(ISI.atp || [{}]), kisi = [].concat(ISI.kisi || [{}]), item = [].concat(L.itemHasil || [""]);
  const rub = [].concat(ISI.rubrik || []);
  const S = {
    alokasi: T(KONFIG.alokasiWaktu, "alokasi waktu"),
    cpRows: cp.map(r => `<tr><td>${T(r.elemen, "elemen CP")}</td><td>${T(r.cp, "kutipan CP")}</td><td class="c">${T(r.tema, "nomor tema KBC")}</td></tr>`).join(""),
    tpLis: daftar(ISI.tp, "TP"),
    atpRows: atp.map((r, i) => `<tr><td class="c">${i + 1}</td><td>${T(r.tp, "TP ATP")}</td><td>${T(r.materi, "materi ATP")}</td><td class="c">${T(r.waktu, "alokasi ATP")}</td></tr>`).join(""),
    kisiRows: kisi.map((r, i) => `<tr><td class="c">${i + 1}</td><td>${T(r.elemen, "elemen")}</td><td>${T(r.tp, "TP")}</td><td>${T(r.materi, "materi")}</td><td>${T(r.indikator, "indikator")}</td><td class="c">${T(r.level, "level")}</td><td class="c">${T(r.topik, "topik Panca Cinta")}</td></tr>`).join(""),
    "lkpd.tujuan": T(L.tujuan, "tujuan LKPD"), "lkpd.alat": daftar(L.alatBahan, "alat dan bahan"), "lkpd.petunjuk": daftar(L.petunjuk, "petunjuk K3/adab"),
    "lkpd.alur": T(L.alur, "alur FIDS/ARKA"), "lkpd.langkah": daftar(L.langkah, "langkah kerja"), "lkpd.kolom": T(L.kolomPengamatan, "kolom pengamatan"),
    "lkpd.item": item.map((x, i) => `<tr><td class="c">${i + 1}</td><td>${T(x, "item hasil kerja")}</td><td></td></tr>`).join(""),
    "lkpd.refleksi": [].concat(L.refleksi || [""]).map(x => `<li>${T(x, "pertanyaan refleksi")}<div class="garis-isi"></div></li>`).join(""),
    "ks.elemen": T(K.elemen, "elemen kartu soal"), "ks.tp": T(K.tp, "TP kartu soal"), "ks.indikator": T(K.indikator, "indikator soal"), "ks.level": T(K.level, "level kartu soal"),
    "ks.stimulus": T(K.stimulus, "stimulus"), "ks.instruksi": daftar(K.instruksi, "instruksi"), "ks.garis": daftar(K.garisBesar, "garis besar penyelesaian"), "ks.kriteria": daftar(K.kriteria, "kriteria keberhasilan"),
    rubrikRows: [0, 1, 2, 3].map(i => { const r = [].concat(rub[i] || []); return `<tr><td><b data-bind="aspek${i + 1}"></b></td>` + [0, 1, 2, 3].map(j => `<td>${T(r[j], "deskriptor rubrik")}</td>`).join("") + "</tr>"; }).join("")
  };
  return html.replace(/%%([\w.]+)%%/g, (m, k) => S[k] !== undefined ? S[k] : m);
}
const normRombel = x => String(x || "").toUpperCase().replace(/\s+/g, " ").trim();
function pilihanRombel() { // mode MTs: konfig.rombel = "Semua" | "IX FATTAN 1" | ["IX FATTAN 1", "IX FATTAN 3"]
  const v = [].concat(KONFIG.rombel === undefined ? [] : KONFIG.rombel).map(normRombel);
  if (v.some(x => /^SEMUA/.test(x))) return Object.keys(REF_SISWA);
  return Object.keys(REF_SISWA).filter(k => v.some(x => x === k || ("IX " + x) === k || k.endsWith(" " + x)));
}
function rombelDiharapkan() {
  if (SATUAN.mode === "rombel") return pilihanRombel();
  const p = (KONFIG.peminatan || "").toLowerCase();
  return Object.keys(REF_SISWA).filter(k => p.includes("kedua") || p.includes("semua") || p.includes(REF_SISWA[k].peminatan.toLowerCase()));
}
function rombelTerpilih() { const h = rombelDiharapkan(); return DATA_SISWA.rombel.filter(r => h.includes(r.rombel)); }
function teksPeminatan() {
  const rs = rombelTerpilih();
  if (SATUAN.mode === "rombel") return rs.map(r => r.rombel).join(", ");
  const pem = [...new Set(rs.map(r => r.peminatan))];
  return SATUAN.kelas + " Peminatan " + pem.join(" & ") + " (" + rs.map(r => r.rombel).join(", ") + ")";
}
function buatLogo() {
  return new Promise(res => { if (!KOP_DEFAULT) return res(false); const im = new Image(); im.onload = () => {
      const s = im.naturalWidth / 1000, L = SATUAN.logo, c = document.createElement("canvas"); c.width = c.height = 244;
      const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, 244, 244); x.imageSmoothingQuality = "high";
      x.drawImage(im, L[0] * s, L[1] * s, L[2] * s, L[2] * s, 0, 0, 244, 244);
      LOGO_SRC = c.toDataURL("image/png"); $$(".logo-image").forEach(i => i.src = LOGO_SRC); res(true); };
    im.onerror = () => res(false); im.src = KOP_DEFAULT; });
}
function djb2(s) { let h = 5381; for (let i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; return h; }
function pemeriksaanOtomatis() {
  const w = [];
  if (!KOP_SRC) w.push("Kop belum termuat — periksa internet atau klik 'Unggah / Ganti Kop'");
  if (!window.INSTRUMEN) w.push("Data window.INSTRUMEN tidak ditemukan");
  const harap = rombelDiharapkan(), ada = DATA_SISWA.rombel.map(r => r.rombel);
  if (!harap.length) w.push(SATUAN.mode === "rombel" ? "konfig.rombel belum valid (\"Semua\" atau nama rombel, mis. \"IX FATTAN 1\")" : "konfig.peminatan belum valid (Agama / MIPA / Keduanya)");
  harap.filter(r => !ada.includes(r)).forEach(r => w.push("Data siswa " + r + " belum disertakan"));
  DATA_SISWA.rombel.forEach(r => { const ref = REF_SISWA[r.rombel];
    if (!ref) { w.push("Rombel tidak dikenal: " + r.rombel); return; }
    const h = djb2(r.siswa.map(s => [r.rombel, s.nis, s.nisn, s.nama].join("|")).join("\n"));
    if (r.siswa.length !== ref.n || h !== ref.h) w.push("Data " + r.rombel + " tidak sama dengan master resmi (" + r.siswa.length + "/" + ref.n + " siswa) — salin ulang dari master_siswa_kelas_12.json"); });
  const tema = KONFIG.temaKBC;
  if (!tema.length || tema.some(x => !PANCA.some(p => String(x).toLowerCase().startsWith(p.toLowerCase())))) w.push("temaKBC harus berisi topik Panca Cinta (SK 6077/2025)");
  if (!String(KONFIG.materiInsersi || "").trim()) w.push("materiInsersi belum diisi");
  if (!/cinta|adab|akhlak/i.test(KONFIG.aspekRubrik[3] || "")) w.push("Aspek ke-4 rubrik sebaiknya 'Sikap & Nilai Cinta' / 'Adab & Nilai Cinta'");
  const teks = $("#doc").innerText;
  const ph = (teks.match(/\{\{[^}]*\}\}/g) || []); if (ph.length) w.push(ph.length + " bagian belum diisi, mis. " + ph[0]);
  const kisi = $$("#tabel-kisi tr").slice(1).map(tr => (tr.cells[5] || {}).innerText || "");
  if (kisi.some(t => /\bC[12]\b/i.test(t))) w.push("Kisi-kisi memuat level C1/C2 (wajib C3–C6)");
  const n = $$(".page-a4").length; if (n < 9) w.push("Jumlah halaman " + n + " (< 9 komponen)");
  const pxA4 = 297 * 96 / 25.4 + 2, lebih = $$(".page-a4").filter(s => s.scrollHeight > pxA4).map(s => s.id.replace("hal-", "Hal. "));
  if (lebih.length) w.push("Isi melebihi 1 halaman A4: " + lebih.join(", "));
  const st = $("#statusCek"); st.className = "status" + (w.length ? " warn" : "");
  st.textContent = w.length ? "⚠ Pemeriksaan otomatis: " + w.join(" · ") : "✔ Pemeriksaan otomatis lulus (mesin v" + VERSI + " · " + SATUAN.singkat + "): kop, identitas madrasah, data siswa resmi, C3–C6, integrasi KBC, " + n + " halaman A4.";
}
function toast(html, ms = 9000) { const t = $("#toast"); t.innerHTML = html; t.style.display = "block"; clearTimeout(t._t); t._t = setTimeout(() => t.style.display = "none", ms); }
function bacaFile(f) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); }); }
function namaFile(ext) { return ("Instrumen_UP_" + KONFIG.mapel + "_" + KONFIG.bentukUjian).replace(/[^\w\-]+/g, "_").replace(/_+/g, "_") + "." + ext; }

/* ---------- PDF tanpa dialog cetak (berfungsi di Canvas Gemini) ---------- */
const CSS_VARS = "--hijau:#0b6b3a;--hijau-muda:#e7f2ea;--emas:#c9a227;--teks:#111;";
async function gambarLembar(sec, skala) {
  const w = sec.offsetWidth, h = Math.max(sec.scrollHeight, sec.offsetHeight);
  const k = sec.cloneNode(true); k.removeAttribute("contenteditable"); k.style.margin = "0"; k.style.boxShadow = "none";
  const xhtml = new XMLSerializer().serializeToString(k);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><foreignObject x="0" y="0" width="${w}" height="${h}">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="${CSS_VARS}width:${w}px;height:${h}px;background:#fff;font-family:'Times New Roman',Times,'Liberation Serif',serif;color:#111">` +
    `<style><![CDATA[${CSS}]]></style>${xhtml}</div></foreignObject></svg>`;
  const im = new Image(); im.src = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  await im.decode(); await new Promise(r => setTimeout(r, 30));
  const PW = Math.round(794 * skala), PH = Math.round(1123 * skala), c = document.createElement("canvas"); c.width = PW; c.height = PH;
  const x = c.getContext("2d"); x.fillStyle = "#fff"; x.fillRect(0, 0, PW, PH);
  const f = Math.min(1, 1123 / h), dw = w * skala * f, dh = h * skala * f; x.drawImage(im, (PW - dw) / 2, 0, dw, dh);
  const url = c.toDataURL("image/jpeg", 0.8), bin = atob(url.split(",")[1]), u = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
  return { data: u, w: PW, h: PH };
}
function rakitPdf(hal) {
  const enc = new TextEncoder(), parts = [], off = []; let len = 0;
  const push = x => { const u = typeof x === "string" ? enc.encode(x) : x; parts.push(u); len += u.length; };
  const obj = (n, f) => { off[n] = len; push(n + " 0 obj\n"); f(); push("\nendobj\n"); };
  push("%PDF-1.4\n"); const N = hal.length;
  obj(1, () => push("<< /Type /Catalog /Pages 2 0 R >>"));
  obj(2, () => push(`<< /Type /Pages /Kids [${hal.map((_, i) => (3 + 3 * i) + " 0 R").join(" ")}] /Count ${N} >>`));
  hal.forEach((g, i) => { const p = 3 + 3 * i, cs = `q 595.28 0 0 841.89 0 0 cm /Im${i} Do Q`;
    obj(p, () => push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.28 841.89] /Resources << /XObject << /Im${i} ${p + 2} 0 R >> >> /Contents ${p + 1} 0 R >>`));
    obj(p + 1, () => push(`<< /Length ${cs.length} >>\nstream\n${cs}\nendstream`));
    obj(p + 2, () => { push(`<< /Type /XObject /Subtype /Image /Width ${g.w} /Height ${g.h} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${g.data.length} >>\nstream\n`); push(g.data); push("\nendstream"); }); });
  const xref = len, total = 3 + 3 * N;
  push(`xref\n0 ${total}\n0000000000 65535 f \n` + off.slice(1).map(o => String(o).padStart(10, "0") + " 00000 n \n").join(""));
  push(`trailer\n<< /Size ${total} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`);
  return new Blob(parts, { type: "application/pdf" });
}
async function unduhPdf() {
  document.activeElement && document.activeElement.blur(); const sel = window.getSelection(); sel && sel.removeAllRanges();
  const hal = $$(".page-a4"), hasil = [];
  try {
    for (let i = 0; i < hal.length; i++) { toast(`Membuat PDF… halaman ${i + 1} dari ${hal.length}`, 60000); hasil.push(await gambarLembar(hal[i], 2)); }
    const a = document.createElement("a"); a.href = URL.createObjectURL(rakitPdf(hasil)); a.download = namaFile("pdf");
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    toast("PDF diunduh: " + hal.length + " halaman A4, margin 2-2-2-2 cm (cover tanpa margin)." + (window.self === window.top ? " Untuk PDF berteks (bisa diblok/salin), <a id='pdfCetak'>gunakan dialog cetak</a>." : ""), 10000);
    const pc = $("#pdfCetak"); if (pc) pc.onclick = () => window.print();
  } catch (e) { console.error(e); toast("Gagal membuat PDF: " + e.message + ". Coba unduh Word lalu simpan sebagai PDF."); }
}
function pasangHandler() {
  $("#inputCover").addEventListener("change", async e => {
    const f = e.target.files[0]; if (!f) return; COVER_SRC = await bacaFile(f);
    const c = $("#page-cover"); c.classList.add("full-bleed-cover"); c.style.padding = "0";
    c.innerHTML = `<img class="cover-image" src="${COVER_SRC}" alt="Cover">`; toast("Cover diganti: gambar tampil penuh 1 halaman A4 tanpa margin.", 4000); e.target.value = "";
  });
  $("#inputKop").addEventListener("change", async e => {
    const f = e.target.files[0]; if (!f) return; KOP_SRC = await bacaFile(f);
    $$("img.kop-image").forEach(i => i.src = KOP_SRC); toast("Kop diperbarui di semua halaman (Halaman 2 s.d. terakhir).", 4000); e.target.value = ""; pemeriksaanOtomatis();
  });
  $("#btnWord").addEventListener("click", unduhWord);
  $("#btnPdf").addEventListener("click", unduhPdf);
  document.addEventListener("input", () => { const sel = window.getSelection();
    const node = sel && sel.anchorNode ? (sel.anchorNode.nodeType === 1 ? sel.anchorNode : sel.anchorNode.parentElement) : null;
    const tr = node && node.closest("table.tbl-nilai tr"); if (tr) hitungBaris(tr); });
}
async function ambilKop() {
  try { const r = await fetch(BASE + SATUAN.kop, { mode: "cors" }); if (!r.ok) throw new Error("HTTP " + r.status);
    const b = await r.blob(); return await new Promise(res => { const f = new FileReader(); f.onload = () => res(f.result); f.readAsDataURL(b); });
  } catch (e) { console.warn("Kop gagal dimuat:", e); return ""; }
}
async function mulai() {
  siapkanData();
  const st = document.createElement("style"); st.id = "css-mesin-masw"; st.textContent = CSS; document.head.appendChild(st);
  if (!document.querySelector("meta[name=viewport]")) { const m = document.createElement("meta"); m.name = "viewport"; m.content = "width=device-width, initial-scale=1.0"; document.head.appendChild(m); }
  const wadah = document.createElement("div"); wadah.innerHTML = TOOLBAR + `<main id="doc">${isiSlot(HALAMAN)}</main>`;
  const app = document.getElementById("app"); if (app) app.replaceWith(...wadah.childNodes); else document.body.append(...wadah.childNodes);
  document.title = "Instrumen UP " + (KONFIG.mapel || "") + " — " + SATUAN.singkat + " Syubbanul Wathon";
  pasangHandler(); isiBind(); buatDaftarNilai();
  KOP_DEFAULT = await ambilKop(); KOP_SRC = KOP_DEFAULT; pasangKop();
  await buatLogo(); if (document.fonts && document.fonts.ready) await document.fonts.ready;
  pasNamaTtd(); pemeriksaanOtomatis();
}
if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", mulai); else mulai();

})();
