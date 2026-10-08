# AI Article Reviewer v2.9

Aplikasi Streamlit untuk membantu review artikel ilmiah secara sistematis, dengan fokus pada keterlacakan temuan ke naskah, rekomendasi revisi yang spesifik, dan alternatif solusi yang dapat ditindaklanjuti.

## Prinsip utama
1. Reviewer tidak boleh menganggap sesuatu tidak ada hanya karena tidak ditemukan dalam satu bagian.
2. Setiap temuan harus memiliki dasar pemeriksaan.
3. Kutipan bukti harus berasal dari naskah yang diunggah.
4. Bukti negatif harus diberi status **BELUM TERDETEKSI SECARA EKSPLISIT**.
5. Bukti AI yang tidak cocok dengan naskah diberi status **BELUM TERKONFIRMASI**.
6. Rekomendasi harus menjawab temuan tertentu, bukan memakai kalimat generik untuk semua artikel.
7. Alternatif solusi tidak boleh mendorong penulis mengarang data, prosedur, hasil, atau referensi.
8. Simpulan Catatan Reviewer mengambil hasil langsung dari Review Bertahap.

## Fitur
- Upload PDF/DOCX.
- Ekstraksi teks lengkap.
- Tampilan naskah lengkap yang dibaca sistem.
- Identifikasi judul dan tampilan judul utuh.
- Kartu alternatif scope berdasarkan judul/bidang.
- Full Review dan Review Bertahap.
- Review: Judul, Abstrak, Pendahuluan, Metodologi, Hasil, Discussion, Conclusion, References.
- Temuan dengan severity CRITICAL/MAJOR/MINOR/EDITORIAL.
- Evidence-grounded review dengan kutipan verbatim dan status verifikasi.
- Rekomendasi reviewer spesifik menurut jenis masalah.
- Tiga alternatif solusi per temuan.
- Audit referensi dengan tabel fixed-layout tanpa geser horizontal.
- Simpulan Catatan Reviewer yang terkoneksi dengan Review Bertahap.
- Mode Demo/Offline.
- Mode AI API OpenAI-compatible.
- Riwayat review.

## Instalasi
Gunakan Python 3.10 atau lebih baru.

```bash
pip install -r requirements.txt
```

Jalankan:

```bash
streamlit run app.py
```

## Mode AI API
Masukkan API key dan endpoint OpenAI-compatible melalui sidebar. API key tidak disimpan ke file oleh aplikasi.

## Catatan penggunaan akademik
Aplikasi ini adalah alat bantu reviewer. Keputusan ilmiah dan editorial tetap memerlukan pemeriksaan manusia. Evidence-grounded berarti sistem membatasi klaim pada apa yang dapat ditelusuri dari teks yang dibaca; sistem tidak menjamin kebenaran substantif sumber eksternal atau kualitas penelitian secara otomatis.

## Alur yang disarankan
Upload artikel → periksa ekstraksi → periksa judul → Review Bertahap dari Judul sampai References → telaah bukti setiap temuan → telaah rekomendasi dan alternatif solusi → buka Simpulan Catatan Reviewer → lakukan pemeriksaan konsistensi akhir.

### v2.9.1 tampilan
Kotak bukti kosong tidak lagi ditampilkan. Header Audit Referensi menggunakan kontras tinggi agar seluruh judul kolom terbaca.
