# CHANGELOG — AI Article Reviewer v2.9

## Fokus revisi
Versi ini memprioritaskan mutu akademik hasil review, keterlacakan bukti, dan keterhubungan Review Bertahap dengan Simpulan Catatan Reviewer.

### 1. Evidence-grounded review
- Setiap temuan pada Review Bertahap sekarang memiliki Dasar/Bukti pemeriksaan.
- Bukti positif menggunakan kutipan verbatim dari naskah dan lokasi tekstual bila dapat ditentukan.
- Bukti negatif menggunakan status **BELUM TERDETEKSI SECARA EKSPLISIT**, bukan klaim bahwa unsur tersebut pasti tidak ada dalam penelitian.
- Narasi bukti dibedakan berdasarkan jenis unsur: tujuan, masalah, state of the art, gap, desain, sumber data, pengumpulan data, analisis, validitas/reliabilitas, etika, hasil, perbandingan literatur, dan kontribusi.
- AI evidence yang tidak dapat dicocokkan dengan naskah diberi status **BELUM TERKONFIRMASI**.

### 2. Rekomendasi reviewer
- Rekomendasi tidak lagi memakai satu narasi generik untuk semua artikel.
- Rekomendasi disusun berdasarkan jenis temuan dan karakter metodologi.
- Setiap rekomendasi diarahkan pada tindakan revisi yang dapat dikerjakan penulis.
- Rekomendasi tetap melarang penambahan data/prosedur yang sebenarnya tidak dilakukan.

### 3. Alternatif solusi
- Setiap temuan memperoleh tiga jalur solusi yang berbeda.
- Alternatif membedakan antara memperjelas unsur yang sudah ada, menyusun ulang uraian, dan melakukan audit konsistensi.
- Alternatif tidak boleh menjadi instruksi untuk mengarang data atau prosedur penelitian.

### 4. Simpulan Catatan Reviewer
- Menu Simpulan Catatan Reviewer mengambil temuan, bukti, rekomendasi, dan alternatif langsung dari Review Bertahap.
- Tidak dibuatkan temuan baru pada tahap simpulan.
- Rekomendasi editorial final tidak ditetapkan sebelum seluruh bagian Judul–Referensi selesai direview.

### 5. Scope
- Kartu alternatif scope dibuat langsung dari judul/bidang aktif.
- Lima alternatif: Fokus Utama, Objek/Unit Analisis, Konseptual/Teoretis, Empiris, dan Kontribusi.
- Scope merupakan alternatif telaah, bukan fakta baru tentang penelitian.

### 6. Referensi
- Temuan audit referensi juga menggunakan struktur bukti yang sama.
- Tabel referensi menggunakan fixed layout dan wrapping agar tidak memerlukan geser horizontal.

### 7. Perbaikan teknis
- Memperbaiki referensi variabel pada fungsi evidence helper.
- Menyimpan teks bagian yang direview agar bukti dan alternatif tetap dapat digunakan saat agregasi.

## v2.9.1
- Removed unexplained/empty evidence cards; evidence is rendered only when substantive text exists.
- Forced high-contrast, readable text inside evidence and recommendation cards.
- Fixed Audit Referensi table header contrast so all column titles are readable.
