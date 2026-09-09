# Token nama Export Related

Perubahan ini berada di Class Module `ExportTemplateParser`, yang dipakai oleh
`ExportRelatedSettings` dan export queue. Perbarui seluruh isi class di VBE dengan
`ExportTemplateParser.cls`, lalu jalankan Debug > Compile.

## Nilai huruf, angka, dan alfanumerik pada command existing

Bagian nilai di luar `{}` kini menerima huruf (A, B, AA), angka (1, 2), atau
gabungannya (A1, B2, 1A). Teks tetap di dalam `{}` dipertahankan. Nilai yang
berikutnya mengganti bagian nilai dari pola pertama, termasuk prefix/suffix-nya.

Contoh nama dokumen tanpa ekstensi: `ORDER NAME 7 NAME`.

| txbName | Output pertama | Output kedua |
| --- | --- | --- |
| `(/*NAME 7 NAME): {CUT_}A{_END}, B` | `ORDER CUT_A_END` | `ORDER CUT_B_END` |
| `(/*NAME 7 NAME): {CUT_}A1{_END}, B2` | `ORDER CUT_A1_END` | `ORDER CUT_B2_END` |
| `(/*NAME 7 NAME): {CUT_}1{_END}, 2` | `ORDER CUT_1_END` | `ORDER CUT_2_END` |

Huruf besar/kecil pada nilai dipertahankan. Gunakan `{}` untuk membatasi teks tetap
agar huruf/angka yang merupakan bagian prefix atau suffix tidak dianggap nilai.
Parser tetap mendahulukan bagian angka pada pola lama; huruf yang melekat langsung
pada angka, seperti A1 atau 1A, sekarang menjadi satu nilai utuh.

Aturan koma pada command existing tetap dipertahankan, termasuk koma di antara dua
digit tanpa spasi. Gunakan `, ` untuk memisahkan nilai angka pada command seperti
`{}1, 2`. Token sync `;` tetap menggunakan grup sync page existing.

## Template nama langsung tanpa /*

Jika template memiliki bagian `{}` dan nilai di luarnya, prefix/suffix pola pertama
dipakai ulang untuk setiap nilai berikutnya. `_` adalah teks biasa; separator lain
seperti spasi atau `-` juga dipertahankan sebagai teks.

| txbName | Output pertama | Output kedua |
| --- | --- | --- |
| `{NAMA}_A, B` | `NAMA_A` | `NAMA_B` |
| `{NAMA}_1, 2` | `NAMA_1` | `NAMA_2` |
| `{NAMA}_A1,B2` | `NAMA_A1` | `NAMA_B2` |
| `{NAMA}-A, B` | `NAMA-A` | `NAMA-B` |
| `{NAMA} A{ END}, B` | `NAMA A END` | `NAMA B END` |

Untuk contoh dua output di atas, gunakan dua grup page, misalnya `txbPage = (1),(2)`.
Koma di luar `{}` memisahkan nilai, baik dengan maupun tanpa spasi. Koma di dalam
`{}` tetap merupakan teks nama. Jumlah nilai harus sesuai jumlah file output,
kecuali satu nilai yang tetap digunakan untuk semua output (suffix duplikat
existing seperti `(1)`, `(2)` tetap berlaku).

Perilaku nama langsung ini tidak aktif jika bagian template mengandung `/*`, baik sebelum
maupun sesudah bagian `{}`. Contoh `OLD/*{NAMA}_A, B` tetap merupakan replacement
existing, bukan daftar nama langsung. Tidak ada token `/` baru; penyebutan A atau 1
pada requirement merupakan alternatif contoh nilai.

## Prefix/suffix __ dan penggabungan |

- `String__` menambahkan String sebelum nama yang sedang diproses.
- `__String` menambahkan String sesudah nama tersebut, sebelum ekstensi file.
- `|` memisahkan bagian pengolahan. Nama utama dibentuk terlebih dahulu dari
  bagian non-affix, mengikuti urutannya dari kiri ke kanan; setelah itu semua
  prefix/suffix diterapkan. Gunakan pemisah ini untuk menggabungkan replacement
  `/*` dengan prefix/suffix. Posisi affix sebelum atau sesudah bagian nama utama
  tidak membuatnya terhapus. Urutan penerapan antar-affix tetap kiri ke kanan.
- Token `__` yang berada dalam satu bagian replacement `/*` tetap diperlakukan
  sebagai teks replacement existing. Bagian `|` kosong ditolak.
- Nilai berupa string, termasuk huruf, angka, dan alfanumerik (`A1`, `A2`, `A3`),
  didukung di prefix maupun suffix. Satu nilai dipakai untuk seluruh output.
- Daftar `,` mengikuti urutan file output; koma tidak wajib diikuti spasi pada
  prefix/suffix. Daftar `;` mengikuti indeks grup sync `txbPage`. Jumlah nilai
  diperiksa terhadap jumlah output atau grup sync yang sesuai.

Contoh dengan nama dasar `BASE` dan tiga output:

| txbName | Output 1 | Output 2 | Output 3 |
| --- | --- | --- | --- |
| `A1__, A2, A3` | `A1BASE` | `A2BASE` | `A3BASE` |
| `A1, A2, A3__` | `A1BASE` | `A2BASE` | `A3BASE` |
| `A1___, A2, A3` | `A1_BASE` | `A2_BASE` | `A3_BASE` |
| `__A1, A2, A3` | `BASEA1` | `BASEA2` | `BASEA3` |
| `___A1, A2, A3` | `BASE_A1` | `BASE_A2` | `BASE_A3` |

Pada `SUSUN___`, dua underscore terakhir adalah marker prefix; satu underscore
lainnya tetap menjadi bagian string `SUSUN_`. Pada `___A, B, C`, dua underscore
pertama adalah marker suffix; underscore sisanya dipakai pada setiap nilai, menjadi
`_A`, `_B`, `_C`. Jika ingin underscore tambahan pada semua nilai prefix, gunakan
`A___, B, C` atau `A, B, C___`. Tidak ada spasi/underscore yang ditambahkan otomatis
selain teks yang dinyatakan pada pola tersebut. Kurung `{}` bisa melindungi koma
atau semicolon literal, misalnya `___{A,B}, C`.

Contoh lengkap:

```text
Nama dasar:
VERTICAL - HCR (UV PREMIUM) - POT CUTTER - UK 100 X 100 CM - TOTAL 3 LBR - (DAVIN)

txbPage: {1-3}
txbName: (/*TOTAL 3): {}1 | SUSUN___ | ___A, B, C

Nama hasil sebelum ekstensi:
SUSUN_VERTICAL - HCR (UV PREMIUM) - POT CUTTER - UK 100 X 100 CM - 1 LBR - (DAVIN)_A
SUSUN_VERTICAL - HCR (UV PREMIUM) - POT CUTTER - UK 100 X 100 CM - 1 LBR - (DAVIN)_B
SUSUN_VERTICAL - HCR (UV PREMIUM) - POT CUTTER - UK 100 X 100 CM - 1 LBR - (DAVIN)_C
```

Suffix dapat diganti menjadi `___1, 2, 3` atau `___A1, A2, A3`. Nama final yang
berbeda tidak mendapat suffix duplikat `(1)`, `(2)`, dan seterusnya. Jika nama final
masih sama, helper duplikat existing tetap berlaku.

Contoh sync: `txbPage = {1-2};{5-6}` dan
`txbName = (/*TOTAL 3): {}1; 2 | A1___; A2 | ___X1, X2, X3, X4`.
Untuk nama dasar `TOTAL 3`, hasilnya `A1_1_X1`, `A1_1_X2`, `A2_2_X3`, `A2_2_X4`.
Prefix dan replacement mengikuti grup sync, sedangkan suffix mengikuti empat
output. Di dalam satu bagian yang menggunakan `;`, pemisahan mengikuti grup sync
existing; gunakan bagian `|` terpisah untuk daftar per-output berbasis koma.

Perbaikan posisi affix (2026-09-09):

```text
txbName: Contoh __ | {pola}_A, B, C, D | __ Fix
txbPage: {1-4}

Contoh pola_A Fix
Contoh pola_B Fix
Contoh pola_C Fix
Contoh pola_D Fix
```

Pada implementasi sebelumnya, `{pola}_A, B, C, D` mengganti seluruh nama sementara
sehingga prefix di depannya hilang. Kini affix dipasang setelah nama utama selesai,
baik nama utama berupa template langsung, string tetap, maupun replacement.
Replacement nama utama tidak mengubah teks prefix/suffix yang ditambahkan kemudian.
Jika ada beberapa prefix, urutan penerapannya tetap seperti sebelumnya: misalnya
`P1__ | FIXED | P2__` menghasilkan `P2P1FIXED`.

## Validasi

Harness menggunakan source parser saat ini dan source sebelum perubahan yang
diadaptasi ke VBScript. 108 pemeriksaan lulus: contoh huruf/angka/alfanumerik,
prefix/suffix, koma, jumlah output, sync `;`, nilai tunggal konstan, serta regresi
replacement `/*`, command numeric, dan suffix duplikat.

Setelah penambahan `__` dan `|`, total 203 pemeriksaan lulus. Tambahan mencakup
prefix/suffix huruf/angka/alfanumerik, literal underscore, marker prefix pada awal
atau akhir daftar, contoh nama lengkap, kombinasi daftar output/grup sync, validasi
jumlah nilai, bagian kosong, serta pemeriksaan suffix duplikat pada nama final.

Perbaikan posisi affix pada 2026-09-09 lulus total 253 pemeriksaan simulasi,
termasuk reproduksi prefix hilang pada source sebelumnya, enam urutan bagian
prefix/nama/suffix, alfanumerik, sync, serta regresi kasus yang sudah ada.
Perbaikan posisi ini belum memiliki konfirmasi runtime CorelDRAW pada catatan ini.

Hasil harness tersebut bukan compile VBA atau export CorelDRAW oleh agent. Jumlah
output mengikuti parser page existing dan tidak otomatis bertambah karena input nama.

Validasi runtime user pada 2026-09-08 berhasil dengan:

```text
txbName: {pola}_A, B, C, D | Contoh __
txbPage: {1-4}

Contoh pola_A.pdf
Contoh pola_B.pdf
Contoh pola_C.pdf
Contoh pola_D.pdf
```

Screenshot dan konfirmasi user membuktikan kombinasi nama langsung, koma, pipe,
prefix dengan spasi, serta empat PDF dengan nama unik. Kasus suffix alfanumerik,
sync `;`, dan replacement `/*` dengan affix masih memiliki bukti simulasi pada
catatan ini; tidak diklaim sebagai uji runtime terpisah dari screenshot tersebut.
