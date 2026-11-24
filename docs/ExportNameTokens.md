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
| `TTL 10 EKOR` | `{} 1 {~}` | `1 EKOR` (dengan spasi awal) |

## Urutan evaluasi operator

| Tahap | Aturan |
| --- | --- |
| 1. Kelompokkan halaman | Bagi `txbPage` pada `;` tingkat terluar. Di setiap segmen, perluas daftar, rentang, dan deret sesuai aturan `()` dan `{}`. Simpan indeks segmen untuk setiap grup output. |
| 2. Bagi pipeline nama | Pisahkan `txbName` pada `\|` di luar `{}`. Kenali `/*an`, `/*a`, dan `/*n` sebelum marker replacement `/*`. Kenali command `(/*target):` sebelum replacement biasa. Marker `__` di dalam ekspresi replacement tetap menjadi teks. |
| 3. Pilih nilai dan perluas deret | Koma mengindeks file hasil. Dalam command replacement atau affix statis, `;` memilih nilai per segmen halaman. Pada nama langsung, `;` membagi blok berdasarkan jumlah nilainya. Untuk deret dinamis, planner menyelaraskan blok dengan segmen halaman eksplisit bila keduanya dipakai, atau mengalokasikan file menurut batas blok. Counter nama dan affix mulai lagi saat blok berganti. |
| 4. Bentuk nama utama | Jalankan bagian non-affix dari kiri ke kanan. Pada replacement, ekspansi deret berlaku pada nilai pengganti, sedangkan wildcard dan `&*` digunakan untuk mencocokkan target. Dalam satu operasi, semua kecocokan diambil dari nama yang masuk ke operasi itu. `{~}` memakai teks aktual dari tiap kecocokan. |
| 5. Terapkan affix | Setelah nama utama terbentuk, terapkan prefix dan suffix sesuai urutan kemunculan antar-affix. Prefix yang ditulis sebelum nama langsung tetap diterapkan pada nama yang sudah terbentuk. |
| 6. Bentuk nama file | Bandingkan nama akhir dalam satu item tanpa membedakan kapitalisasi. Tambahkan nomor pada setiap kemunculan nama yang berulang, lalu gunakan ekstensi format yang dipilih. Preflight antrean memeriksa path akhir lintas item. |

Untuk ekspresi `TOTAL /*n /* 1 | ___1/*n`, `/*n` di kiri mencocokkan angka pada target, `/*` kedua memisahkan target dari pengganti `1`, dan `1/*n` pada suffix menghitung urutan output. [`BuildExportFileName` dan `BuildUniqueExportName`](../src/classes/ExportTemplateParser.cls) · [`Prepare`](../src/classes/ExportTokenPlan.cls) · [`Expand` dan `ReplaceTarget`](../src/classes/ExportDynamicTokens.cls) · [`BuildPaths`](../src/classes/ExportQueueRunner.cls).

## Grammar formal (EBNF)

Grammar ini mencakup bentuk sintaks yang didokumentasikan. Produksi `? … ?` memakai aturan leksikal yang dijelaskan sesudah grammar. `[ ... ]` berarti opsional, `{ ... }` berarti pengulangan, dan `A - B` mengecualikan bentuk `B` dari `A`. Tanda dalam kutip merupakan karakter yang diketik pengguna.

### Grammar `txbPage`

```ebnf
page-input    = [ page-segment, { ";", page-segment } ] ;
page-segment  = page-item, { ",", page-item } ;
page-item     = page-span | page-group | page-split ;
page-group    = "(", page-span, { ",", page-span }, ")" ;
page-split    = "{", page-item, { ",", page-item }, "}" ;
page-span     = positive-int, [ "-", positive-int | "/*n", [ "-", positive-int ] ] ;
positive-int  = digit, { digit } ;
digit         = "0" | "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8" | "9" ;
```

- `positive-int` bernilai 1 sampai 2.147.483.647. Rentang eksplisit harus menaik atau setara dan berada dalam jumlah halaman dokumen. Batas deret `/*n-n` menghitung halaman fisik dari seed. Spasi di sekitar pemisah halaman dapat diabaikan. Seed, `/*n`, dan batas deret dinamis ditulis rapat.
- Bila segmen hanya berisi `page-span`, seluruh daftarnya menjadi satu grup. Bila segmen memuat `()` atau `{}`, setiap item tingkat terluar menghasilkan grup tersendiri. Di dalam `{}`, setiap halaman dari `page-span` menjadi satu file dan `()` tetap menyatukan semua halamannya. `{}` boleh bersarang.
- Input kosong memilih halaman aktif. Grup berisi lebih dari satu halaman memerlukan PDF. Deret dinamis yang kekurangan halaman boleh menghasilkan bagian yang tersedia disertai peringatan.

### Grammar `txbName`

```ebnf
name-input        = [ name-part, { "|", name-part } ] ;
name-part         = command | replacement | prefix | suffix | direct-name ;
command           = "(/*", target, "):", command-block, { ";", command-block } ;
command-block     = replacement-value, { ",", replacement-value } ;
replacement       = target, "/*", replacement-body ;
replacement-body  = { replacement-piece | "," | ";" } ;
replacement-value = { replacement-piece } ;
replacement-piece = preserve | replacement-braced-text | name-series | name-text ;
preserve          = "{~}" ;
replacement-braced-text
                  = ( "{", { braced-character | replacement-braced-text | preserve }, "}" )
                    - preserve ;

prefix            = name-value, "__", [ affix-tail ] | name-value, affix-tail, "__" ;
suffix            = "__", name-value, [ affix-tail ] ;
affix-tail        = ( ",", name-value, { ",", name-value } )
                  | ( ";", name-value, { ";", name-value } ) ;
direct-name       = plain-name | direct-block, { ";", direct-block } ;
direct-block      = name-value, { ",", name-value } ;
name-value        = name-piece, { name-piece } ;
name-piece        = name-text | braced-text | name-series ;
name-series       = seed, dynamic-marker, [ "-", positive-int ] ;
dynamic-marker    = "/*an" | "/*a" | "/*n" ;

target            = target-literal
                  | [ target-literal ], wildcard, { [ target-literal ], wildcard },
                    [ target-literal ] ;
wildcard          = dynamic-marker ;
target-literal    = target-text, { "&*", target-text } ;
target-text       = target-piece, { target-piece } ;
target-piece      = target-character | braced-text ;
braced-text       = "{", { braced-character | braced-text }, "}" ;
seed              = ? run alfanumerik tepat sebelum marker dinamis ? ;
plain-name        = ? nama literal yang tidak dikenali sebagai pola operator ? ;
name-text         = ? teks nilai yang berhenti pada delimiter atau marker aktif ? ;
target-character  = ? karakter target di luar pembatas command, wildcard, dan &* ? ;
braced-character  = ? karakter selain kurung kurawal pembuka dan penutup ? ;
```

## Aturan leksikal dan semantik

- Parser nama melindungi `,`, `;`, dan `|` dengan `{}`. Spasi di sekitar `|` dan `;` diabaikan. Pada daftar yang memakai koma, spasi setelah koma dipertahankan dalam nama. Pada target, `&*` membagi alternatif literal di antara wildcard. Wildcard target memakai marker tanpa seed atau batas.
- Pada replacement biasa, marker pemisah `/*` adalah marker pertama di luar `{}` yang tidak cocok dengan marker dinamis. Gunakan spasi seperti `Lama /* anyar` jika awal teks pengganti dapat terbaca sebagai bagian marker dinamis.
- Pada nilai pengganti, kenali bentuk persis `{~}` sebagai `preserve` sebelum memproses teks berkurung. Aturan ini berlaku untuk replacement biasa dan command `(/*target):`. `{}` tetap menjadi teks kosong. Parser mempertahankan token `preserve` saat melepas kurung teks literal, lalu memakai konteks tiap kecocokan untuk mengisinya. [Parser nama](../src/classes/ExportTemplateParser.cls) · [`PreserveValue`](../src/classes/ExportDynamicTokens.cls).
- Suatu `direct-block` statis dengan daftar atau batas blok perlu pola berkurung kurawal `{}` dan nilai di luarnya. Nama literal biasa dapat berdiri sendiri. Deret dinamis memakai aturan seed untuk marker masing-masing. Batas `-n` di nama menghitung file.
- Kecocokan jumlah nilai dengan file atau grup sync, konsistensi panjang blok, dan syarat `{~}` diperiksa saat merencanakan nama. Pada `{~}`, gunakan satu wildcard, atau target literal dengan angka sebagai titik pemisah. Pemilihan ekstensi, validitas path, dan benturan antarjob diperiksa pada tahap pembentukan file. [Parser halaman](../src/classes/ExportPageParser.cls) · [Parser nama](../src/classes/ExportTemplateParser.cls) · [Token dinamis](../src/classes/ExportDynamicTokens.cls) · [Perencana token](../src/classes/ExportTokenPlan.cls).

## Matriks regresi

Gunakan dokumen dengan nama dan jumlah halaman pada kolom fixture, tujuan ekspor kosong yang valid, serta layer yang dapat diekspor. Format default matriks adalah PDF. `—` berarti input kosong. `file.pdf ← [1,3]` berarti file berisi halaman sumber 1 dan 3 dalam urutan tersebut. Untuk pengujian native, periksa file yang terbentuk dan isi tiap PDF. Kasus yang sama bisa dijalankan lewat single export dan antrean jika berlaku pada keduanya.

### Hasil yang diharapkan

| ID | Fixture | `txbPage` | `txbName` | Output yang diharapkan |
| --- | --- | --- | --- | --- |
| P01 | `Order.cdr`, 3 halaman, halaman aktif 2 | — | — | `Order.pdf` ← [2]. |
| P02 | `Order.cdr`, 3 halaman | `1,3` | `Dipilih` | `Dipilih.pdf` ← [1,3]. |
| P03 | `Order.cdr`, 3 halaman | `1-3` | `Dipilih` | `Dipilih.pdf` ← [1,2,3]. |
| P04 | `Order.cdr`, 7 halaman | `{1,(2-5),6,7}` | `P1/*an` | `P1.pdf` ← [1], `P2.pdf` ← [2,3,4,5], `P3.pdf` ← [6], `P4.pdf` ← [7]. |
| N01 | `TOTAL 3.cdr`, 6 halaman | `(1,2),(3,4),(5,6)` | `(/*TOTAL 3): {}1` | `1 (1).pdf` ← [1,2], `1 (2).pdf` ← [3,4], `1 (3).pdf` ← [5,6]. |
| N02 | `Order.cdr`, 2 halaman | `{1-2}` | `{A}1, 2` | `A1.pdf` ← [1], `A 2.pdf` ← [2]. |
| N03 | `Order.cdr`, 2 halaman | `{1-2}` | `{ORDER, A}_1,2` | `ORDER, A_1.pdf` ← [1], `ORDER, A_2.pdf` ← [2]. |
| N04 | `Order.cdr`, 2 halaman | `{1-2}` | `Contoh __ \| {A}1,2 \| __ Fix` | `Contoh A1 Fix.pdf` ← [1], `Contoh A2 Fix.pdf` ← [2]. |
| N05 | `ORDER TOTAL 3.cdr`, 9 halaman | `{3-7,9};{1,2,8}` | `(/*TOTAL 3): {}1; 2` | `ORDER 1 (1).pdf` sampai `ORDER 1 (6).pdf` masing-masing ← [3], [4], [5], [6], [7], [9]. `ORDER 2 (1).pdf` sampai `ORDER 2 (3).pdf` masing-masing ← [1], [2], [8]. |
| N06 | `Order.cdr`, 4 halaman | `{1-4}` | `{A}1,2;{B}1,2` | `A1.pdf` ← [1], `A2.pdf` ← [2], `B1.pdf` ← [3], `B2.pdf` ← [4]. |
| N07 | `AYAM TTL 10 EKOR - TOTAL 20 EKOR.cdr`, 1 halaman | `1` | `(/*TOTAL&*TTL /*n EKOR): {~} 1 {~}` | `AYAM TTL 1 EKOR - TOTAL 1 EKOR.pdf` ← [1]. |
| N08 | `AYAM TOTAL 10 EKOR.cdr`, 10 halaman | `{1/*n}` | `(/*TOTAL 10): {}A1/*an-5;{}B1/*an-5 \| ___1/*n` | Lima file `AYAM A{k} EKOR_{k}.pdf` ← [k] dan lima file `AYAM B{k} EKOR_{k}.pdf` ← [5+k], dengan k = 1–5. `File Qty` = 10. Ulangi lewat antrean dan MacroRunner. |
| N09 | `Order.cdr`, 3 halaman | `{1-3}` | `Y/*a` | `Y.pdf` ← [1], `Z.pdf` ← [2], `AA.pdf` ← [3]. |
| N10 | `Order.cdr`, 3 halaman | `{1-3}` | `008/*n` | `008.pdf` ← [1], `009.pdf` ← [2], `010.pdf` ← [3]. |
| W01 | `Order.cdr`, 3 halaman | `{1/*n-5}` | `{A}1/*n` | `A1.pdf` ← [1], `A2.pdf` ← [2], `A3.pdf` ← [3]. Peringatan menyebut **3 file, 3 halaman**. |
| W02 | `Order.cdr`, 3 halaman | `(1/*n-5)` | `{A}1/*n` | `A1.pdf` ← [1,2,3]. Peringatan menyebut **1 file, 3 halaman**. |

Pada W01 dan W02, cocokkan kedua angka dalam peringatan dengan jumlah file dan total halaman dari seluruh output. [`ExportTokenPlan.Prepare`](../src/classes/ExportTokenPlan.cls).

### Penolakan yang diharapkan

Semua kasus di bawah harus berhenti sebelum ekspor dan menghasilkan **nol file**. Kolom terakhir menyebut tahap dan inti pesan untuk diperiksa. Pembungkus pesan dapat berbeda menurut jalur single export atau antrean.

| ID | Fixture | `txbPage` | `txbName` | Tahap dan pesan yang diharapkan |
| --- | --- | --- | --- | --- |
| E01 | `Order.cdr`, 3 halaman | `{1-5}` atau `(1-5)`, masing-masing diuji | — | Parser halaman. `Page 5 tidak tersedia`. |
| E02 | `Order.cdr`, 3 halaman | `{1/*a}` | — | Parser halaman. `txbPage hanya menerima /*n`. |
| E03 | `Order.cdr`, 3 halaman | `{1-3}` | `{A}1,2` | Perencana token. `Jumlah value nama tidak sesuai`. |
| E04 | `ORDER TOTAL 3.cdr`, 3 halaman | `{1-3}` | `(/*TOTAL 3): {}1; 2` | Perencana token. `Jumlah value nama tidak sesuai` karena hanya ada satu grup sync. |
| E05 | `Order.cdr`, 1 halaman | `1` | `MISSING /*n /* 1` | Replacement. `Target nama tidak ditemukan`. |
| E06 | `AYAM TTL A1 10 EKOR.cdr`, 1 halaman | `1` | `(/*TTL /*an /*n EKOR): {~}1` | Replacement. `{~} memerlukan tepat satu wildcard`. |
| E07 | `Order.cdr`, 3 halaman | `{1-3}` | `A1B/*an` | Ekspansi nama. `memerlukan akhiran angka`. |
| E08 | `Order.cdr`, 5 halaman | `{1-5}` | `{}1/*n-3,9` | Perencana token. `Pisahkan deret dinamis dan daftar koma menjadi blok ;`. |
| E09 | `Order.cdr`, 5 halaman | `{1-5}` | `{}A1/*an-5 \| ___1/*n-3` | Perencana token. `Batas eksplisit nama/affix` tidak konsisten. |
| E10 | `Order.cdr`, 5 halaman | `{1-5}` | `{}A1/*an;{}B1/*an` | Perencana token. Blok awal `memerlukan batas` atau segmen halaman eksplisit. |
| E11 | `Order.cdr`, 3 halaman, format PNG, JPEG, atau DXF | `(1-3)` | `Gabung` | Validasi format. `Grup multi-page hanya didukung untuk PDF` atau pesan setara. |
| E12 | `Order.cdr`, 1 halaman, jalur antrean | `1` | `Cut?` | Validasi nama file pada preflight antrean. `Nama file mengandung karakter tidak valid`. |
| E13 | `Order.cdr`, 3 halaman | `{5/*n}` | — | Perencana token. `Format page tidak valid atau kosong` karena deret tidak menghasilkan halaman. |

Setelah kasus token, jalankan smoke test per format dengan `txbPage = {1-2}` dan `txbName = {PROOF}_A,B`. Harapkan `PROOF_A` dari halaman 1 dan `PROOF_B` dari halaman 2, memakai ekstensi format terpilih. Gunakan konfigurasi valid untuk memeriksa layer, preset, dan Parent Directory secara terpisah.

## Verifikasi dan rujukan pengembang

Untuk memasang source, siapkan Class Module `ExportDynamicTokens` dan `ExportTokenPlan` beserta parser dan pemanggilnya sesuai struktur proyek. File class yang headerless dapat disalin ke Class Module yang sesuai. Jalankan **Debug > Compile**, lalu uji hasil file pada dokumen representatif. [Pemasangan](../README.md#setup-and-notes) · [Struktur source](../README.md#project-structure).

Pemeriksaan adaptasi terdahulu mencatat 108 kasus parser nama dan sync, 203 setelah penambahan affix dan pipeline, 253 untuk perbaikan posisi affix, serta 160 untuk blok nama dan spasi literal. Gunakan contoh pada dokumen ini untuk aturan spasi aktif. Riwayat dan konfirmasi runtime pengguna berada di `../../Development.log`.

Harness lokal `../VerifyExportTokens.ps1` mengadaptasi VBA ke VBScript dengan dokumen tiruan. Harness itu mencari file class di layout lama dan memerlukan penyesuaian path class sebelum dijalankan pada layout `src/` ini. Script juga memakai helper dari proyek MacroRunner di workspace yang sama. Bukti harness tidak mencakup compile VBE, MSForms, atau file hasil ekspor CorelDRAW.
