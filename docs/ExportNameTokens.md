# Token nama dan halaman ExportRelated

`txbPage` menentukan halaman dan cara membaginya menjadi file. `txbName` menentukan nama dasar tiap file, lalu format ekspor menentukan ekstensi. Aturan ini berlaku untuk editor, antrean, dan MacroRunner Behavior yang memakai editor yang sama. [Panduan penggunaan](../README.md) · [Referensi algoritma](algorithm-references.md).

## Pemilihan halaman

| Input `txbPage` | Hasil |
| --- | --- |
| Kosong | Halaman aktif dokumen sumber. Pada antrean, halaman ini ditetapkan saat item disimpan. |
| `1,3` | Satu PDF berisi halaman 1 dan 3. |
| `1-3` | Satu PDF berisi halaman 1 sampai 3. |
| `(1),(2-3)` | Dua PDF. File pertama berisi halaman 1 dan file kedua berisi halaman 2–3. |
| `{1-3}` | Tiga file, masing-masing berisi satu halaman. |
| `{1,(2-5),6,7}` | Empat file. Halaman 2–5 digabung menjadi satu PDF. |
| `{1,2};{3,4}` | Empat file dalam dua grup sinkronisasi. |

Kurung `()` menggabungkan halaman dan `{}` memecah isinya sambil mempertahankan grup `()` di dalamnya. PDF mendukung output multi-page. PNG, JPEG, dan DXF memerlukan satu halaman per file.

### Halaman dinamis dengan `/*n`

| Input `txbPage` | Hasil pada dokumen yang cukup panjang |
| --- | --- |
| `{1/*n}` | Satu file per halaman mulai dari halaman 1 sampai akhir dokumen. |
| `{3/*n-3}` | Tiga file untuk halaman 3, 4, dan 5. |
| `(1/*n-5)` | Satu PDF berisi halaman 1–5. |
| `{(1/*n-3),4,5}` | Tiga file. PDF halaman 1–3, lalu file halaman 4 dan file halaman 5. |

`/*n` adalah token dinamis untuk halaman. Akhiran `-n` membatasi jumlah halaman fisik dari seed. Pada dokumen tiga halaman, `{1/*n-5}` menghasilkan tiga file dan `(1/*n-5)` menghasilkan satu PDF, keduanya disertai peringatan. Grup dinamis berikutnya boleh kosong saat halaman sudah habis. Jika seluruh input tidak menghasilkan halaman, validasi berhenti. Rentang eksplisit seperti `{1-5}` atau `(1-5)` pada dokumen yang sama ditolak sebelum ekspor, termasuk jika batas nama lebih kecil.

## Pola nama

Jika `txbName` kosong, nama dasar berasal dari nama dokumen tanpa ekstensi. Pola tetap dapat dipakai untuk beberapa file. Daftar dengan lebih dari satu nilai harus cocok dengan jumlah file atau jumlah grup sinkronisasi sesuai operatornya.

### Nama langsung, daftar, dan spasi literal

| Input `txbName` | Nama dasar untuk dua output |
| --- | --- |
| `{A}1,2` | `A1`, `A2` |
| `{A}1, 2` | `A1`, `A 2` |
| `{NAMA}_A,B` | `NAMA_A`, `NAMA_B` |
| `{NAMA}_A1,B2` | `NAMA_A1`, `NAMA_B2` |
| `{NAMA} A{ END},B` | `NAMA A END`, `NAMA B END` |
| `{ORDER, A}_1,2` | `ORDER, A_1`, `ORDER, A_2` |

Di `txbName`, `{}` melindungi teks tetap. Nilai di luarnya dapat berupa huruf, angka, atau alfanumerik. Koma di luar kurung kurawal memisahkan nilai, termasuk di antara dua digit. Spasi setelah koma menjadi bagian output, termasuk pada command replacement dan affix. Spasi pemisah di sekitar `;` dan `|` diabaikan. Karakter `_` tunggal dan `/` sendiri tidak berfungsi sebagai operator.

### Replacement dengan `/*`

Replacement biasa `Lama/*Baru` mengubah `Desain Lama` menjadi `Desain Baru`. Gunakan `(/*target): nilai` untuk menentukan pengganti terstruktur. Pencocokan target mengabaikan kapitalisasi.

| Input `txbName` untuk dokumen `ORDER TOTAL 3` | Nama dasar untuk dua output |
| --- | --- |
| `(/*TOTAL 3): {CUT_}A{_END},B` | `ORDER CUT_A_END`, `ORDER CUT_B_END` |
| `(/*TOTAL 3): {CUT_}A1,B2` | `ORDER CUT_A1`, `ORDER CUT_B2` |
| `(/*TOTAL 3): {CUT_}1, 2` | `ORDER CUT_1`, `ORDER CUT_ 2` |
| `(/*TOTAL 3): {}1` | `ORDER 1` untuk kedua output, sebelum suffix duplikat. |

Pada target terstruktur, kurung kurawal melindungi angka yang merupakan bagian dari teks tetap. Contoh `(/*{Target 1} 1 Akurat): {Target 2} 3 {MOA}` mengganti `Target 1 1 Akurat` menjadi `Target 2 3 MOA`. Teks tetap dapat ditulis di salah satu sisi angka atau di kedua sisi.

Command tetap seperti `(/*TOTAL 3): {}1` memakai nilai yang sama di setiap grup halaman. Nama dasar berulang dalam satu item diberi suffix `(1)`, `(2)`, dan seterusnya, termasuk kemunculan pertama. [Benturan nama file](../README.md#file-name-conflicts).

### Prefix, suffix, dan pipeline

- `String__` menambah prefix. `__String` menambah suffix sebelum ekstensi.
- Underscore di luar marker menjadi teks literal. `SUSUN___` memberi prefix `SUSUN_` dan `___A,B` memberi suffix `_A` serta `_B`.
- Daftar affix dengan koma mengikuti indeks output. Daftar affix dengan `;` dapat mengikuti indeks grup sinkronisasi halaman. Spasi pada `___A, B` menghasilkan suffix `_A` dan `_ B`.
- `|` memisahkan operasi. Parser membentuk nama utama, lalu menerapkan affix sesuai urutan antar-affix. Bagian pipeline kosong dan nilai affix kosong ditolak.

Dengan `txbPage = {1-4}` dan `txbName = Contoh __ | {A}1,2; {B}1,2 | __ Fix`, hasilnya `Contoh A1 Fix`, `Contoh A2 Fix`, `Contoh B1 Fix`, dan `Contoh B2 Fix`.

## Grup sinkronisasi dan blok dengan `;`

Pada command replacement dan affix yang memakai nilai per grup, `;` memasangkan nilai nama dengan segmen `txbPage`. Contoh untuk dokumen `ORDER TOTAL 3`:

```text
txbPage: {3-7,9};{1,2,8}
txbName: (/*TOTAL 3): {}1; 2
```

Enam file pertama memakai nama dasar `ORDER 1`, tiga file berikutnya memakai `ORDER 2`. Suffix membedakan file dengan nama dasar yang sama.

Pada nama langsung, setiap blok mencantumkan nilai untuk file berikutnya. `{A}1,2,3;{B}X,Y` menghasilkan `A1`, `A2`, `A3`, `BX`, dan `BY`, baik untuk `{1-5}` maupun `{1-3};{4-5}`. Panjang blok boleh berbeda. Total nilai harus sesuai jumlah file. `{A}1;{B}1` memberi tepat dua nama. Blok kosong, nilai kosong, dan blok tanpa pola ditolak.

## Token dinamis untuk nama dan affix

| Pola | Deret nilai |
| --- | --- |
| `1/*n` | `1,2,3,...` |
| `008/*n` | `008,009,010,...` |
| `Y/*a` | `Y,Z,AA,AB,...` |
| `y/*a` | `y,z,aa,ab,...` |
| `A1B/*a` | `A1B,A1C,...,A1Z,A1AA,...` |
| `A008/*an` | `A008,A009,A010,...` |
| `A1B008/*an` | `A1B008,A1B009,...` |
| `A4/*an-5` | Maksimal lima nilai, dari `A4` sampai `A8`. |

`/*n` memerlukan seed angka. `/*a` menaikkan rangkaian huruf paling belakang dan mempertahankan prefix. Akhiran huruf bercampur kapitalisasi seperti `aB/*a` ditolak. `/*an` menaikkan angka paling belakang, sehingga `A1B/*an` ditolak. Padding nol dipertahankan dan boleh bertambah panjang saat angka melebar. Token dalam `{teks literal}` tetap menjadi teks.

### Batas blok dan counter

Di `txbName`, akhiran `-n` membatasi jumlah file dari blok tersebut. Contoh untuk dokumen `AYAM TOTAL 10 EKOR` dengan sepuluh halaman:

```text
txbPage: {1/*n}
txbName: (/*TOTAL 10): {}A1/*an-5;{}B1/*an-5 | ___1/*n
```

Halaman 1–5 menghasilkan `AYAM A1 EKOR_1` sampai `AYAM A5 EKOR_5`. Halaman 6–10 menghasilkan `AYAM B1 EKOR_1` sampai `AYAM B5 EKOR_5`. Counter nama dan affix kembali ke awal saat blok berganti. Urutan halaman terus berjalan.

- Jika `txbPage` memiliki grup eksplisit seperti `{1-5};{6-10}`, blok nama boleh tanpa batas. Jumlah blok nama dan grup sinkronisasi harus cocok.
- Tanpa grup eksplisit, setiap blok sebelum blok terakhir memerlukan batas. Daftar nilai langsung dapat menentukan panjang blok. Contoh `{A}1,2;{B}1,2 | ___1/*n` menghasilkan `A1_1,A2_2,B1_1,B2_2`.
- Batas eksplisit nama dan affix dalam satu blok harus sama. Untuk `{}A1/*an-5`, pakai `___1/*n` jika suffix perlu mengikuti panjang blok. Prefix dan suffix boleh ditempatkan pada kedua sisi pola utama di antara `|`. Marker affix diwariskan, sehingga `___X;Y` tetap merupakan suffix pada blok kedua.
- Pisahkan deret dinamis dan daftar nilai koma ke blok yang berbeda. Bentuk `{}1/*n-3,9` ditolak. Koma di dalam `{}` tetap literal.

### Wildcard target, alternatif `&*`, dan preservasi `{~}`

Pada **target** replacement, `/*a`, `/*n`, dan `/*an` mencocokkan token huruf, angka, atau alfanumerik utuh. Misalnya, `/*n` melewatkan `10A`. Setiap target yang cocok diganti seluruhnya.

```text
Nama sumber: AYAM TOTAL 10 EKOR
txbPage: {1/*n}
txbName: TOTAL /*n /* 1 | ___1/*n
Hasil: AYAM 1 EKOR_1, AYAM 1 EKOR_2, ...
```

`&*` memisahkan alternatif literal. Target `TOTAL&*TTL&*JUMLAH /*n EKOR` cocok dengan `TOTAL 10 EKOR`, `TTL 20 EKOR`, atau `JUMLAH 3 EKOR`. Alternatif boleh terdiri dari beberapa kata, misalnya `JUMLAH TOTAL&*TTL /*n`. Teks sebelum atau sesudah wildcard memberi konteks pencocokan. `&` biasa bersifat literal. Jika target yang memakai wildcard atau alternatif tidak ditemukan, validasi gagal sebelum ekspor.

`{~}` di sisi pengganti menyalin teks aktual dari sisi terkait pada setiap kecocokan. Spasi tepi teks yang disalin dibuang. Tulis spasi hasil di luar kurung. Untuk target literal tanpa wildcard, angka pertama menjadi titik pemisah. `{}` berarti teks kosong. Parser menolak `{~}` tanpa target atau dengan lebih dari satu wildcard.

| Target aktual | Replacement | Hasil |
| --- | --- | --- |
| `TTL 10 EKOR` | `{~} 1 {~}` | `TTL 1 EKOR` |
| `TTL 10 EKOR` | `{~}1 {~}` | `TTL1 EKOR` |
| `TTL 10 EKOR` | `{}1 {~}` | `1 EKOR` |
| `TTL 10 EKOR` | `{} 1 {~}` | ` 1 EKOR` (dengan spasi awal) |

## Verifikasi dan rujukan pengembang

Untuk memasang source, siapkan Class Module `ExportDynamicTokens` dan `ExportTokenPlan` beserta parser dan pemanggilnya sesuai struktur proyek. File class yang headerless dapat disalin ke Class Module yang sesuai. Jalankan **Debug > Compile**, lalu uji hasil file pada dokumen representatif. [Pemasangan](../README.md#setup-and-notes) · [Struktur source](../README.md#project-structure).

| Kasus regresi | Pemeriksaan |
| --- | --- |
| Sepuluh halaman dengan dua blok dinamis | Jalankan lewat antrean dan MacroRunner. Periksa nama pertama blok kedua, halaman 6, reset suffix, dan `File Qty`. |
| Tiga halaman dengan batas dinamis dan rentang eksplisit | Bandingkan `{1/*n-5}` dan `(1/*n-5)` yang memperingatkan dengan `{1-5}` dan `(1-5)` yang harus ditolak sebelum file dibuat. |
| Wildcard dan alternatif dengan target berulang | Periksa semua kecocokan, konteks `{~}` per target, serta penolakan target yang tidak ditemukan. |
| Template lama | Coba PDF, PNG, JPEG, dan DXF via single export dan antrean. Periksa layer, preset, Parent Directory, dan opsi format yang relevan. |
| Batas dan literal | Periksa penolakan batas `-5` dan `-3` dalam satu blok, `A1B/*an`, serta blok awal tanpa batas. Uji `Z` → `AA`, `009` → `010`, dan spasi setelah koma. |

Pemeriksaan adaptasi terdahulu mencatat 108 kasus parser nama dan sync, 203 setelah penambahan affix dan pipeline, 253 untuk perbaikan posisi affix, serta 160 untuk blok nama dan spasi literal. Gunakan contoh pada dokumen ini untuk aturan spasi aktif. Riwayat dan konfirmasi runtime pengguna berada di `../../Development.log`.

Harness lokal `../VerifyExportTokens.ps1` mengadaptasi VBA ke VBScript dengan dokumen tiruan. Harness itu mencari file class di layout lama dan memerlukan penyesuaian path class sebelum dijalankan pada layout `src/` ini. Script juga memakai helper dari proyek MacroRunner di workspace yang sama. Bukti harness tidak mencakup compile VBE, MSForms, atau file hasil ekspor CorelDRAW.
