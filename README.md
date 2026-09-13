# ExportRelated

ExportRelated adalah macro VBA untuk CorelDRAW yang dirancang untuk mengelola pekerjaan export **PDF**, **DXF**, **PNG**, dan **JPEG** melalui antrean dengan pengaturan terpisah untuk setiap job.

Macro ini dibuat untuk mengurangi pekerjaan manual ketika beberapa page atau dokumen perlu diekspor ke format, directory, nama file, dan kombinasi layer yang berbeda.

ExportRelated menyediakan **Export Queue**, pengelompokan page, template nama berbasis token, pengaturan per format, serta **Parent Directory** untuk menggunakan folder dokumen sumber sebagai lokasi hasil.

## Main Features

### Export Queue

Setiap pekerjaan dimasukkan ke antrean sebagai satu item `ExportSettingItem` dan ditampilkan melalui `lbxSettingLists` pada `ExportRelatedMenu`.

Satu item menyimpan:

- dokumen sumber;
- directory tujuan atau pilihan Parent Directory;
- format export;
- pilihan page dan pembagian file hasil;
- template nama file;
- pengaturan format;
- pilihan `Layer 1`, `Layer 2`, dan `Layer 3`.

Satu job dapat menghasilkan satu atau beberapa file. Karena itu, jumlah item dalam antrean tidak selalu sama dengan jumlah file hasil.

Summary pada ListBox menampilkan format, preset, `File Qty`, `Page`, dan `Layer`. Teks summary merupakan representasi visual; data pekerjaan tetap disimpan pada item queue.

### Multi-Document Export

Antrean dapat berisi pekerjaan dari beberapa dokumen CorelDRAW yang masih terbuka.

Saat item baru ditambahkan, dokumen aktif menjadi sumber item tersebut. Untuk menambahkan pekerjaan dari dokumen lain, aktifkan dokumen itu lalu gunakan **Add** kembali.

Contoh ilustratif:

| Job | Dokumen Sumber | Format | Directory Tujuan |
|---|---|---|---|
| 1 | `Label.cdr` | `.png` | `D:\Output\Preview` |
| 2 | `Cutting.cdr` | `.pdf` | `D:\Output\Cutting` |
| 3 | `Label.cdr` | `.jpg` | `D:\Output\Approval` |

Ketiga pekerjaan dapat dijalankan melalui satu antrean. Dokumen sumber harus tetap terbuka sampai proses export selesai.

---

## Queue Controls

| Kontrol | Perilaku |
|---|---|
| `cmdAddSetting` | Membuka editor untuk menambahkan job dari dokumen aktif. |
| `cmdModify` | Mengubah job terpilih dengan tetap menggunakan dokumen sumber item tersebut. |
| `cmdRemoveSetting` | Menghapus job terpilih dari antrean. |
| `cmdClearLists` | Mengosongkan antrean dan mengembalikan tampilan checkbox ke default Layer 1. |
| `cmdExport` | Memvalidasi dan menjalankan seluruh antrean. |
| `cmdClose` | Menutup Main UserForm. |

**Remove** dan **Clear Lists** hanya menghapus entry antrean, bukan objek dokumen atau file hasil di disk.

Antrean disimpan dalam memori selama instance Main UserForm masih tersedia. Membuka kembali form setelah ditutup membuat antrean baru. Setelah export berhasil, daftar job tetap tersedia.

---

## Export Settings

Editor `ExportRelatedSettings` digunakan untuk menentukan kebutuhan setiap job.

| Kontrol | Fungsi |
|---|---|
| `txbDirectory` | Directory tujuan export. |
| `cmdBrowse` | Memilih folder tujuan. |
| `chkParentDirectory` | Menggunakan folder dokumen sumber. |
| `cmbExFormat` | Memilih `.pdf`, `.dxf`, `.png`, atau `.jpg`. |
| `txbPage` | Menentukan page dan pembagian file hasil. |
| `txbName` | Menentukan nama atau template nama file. |
| `cmdFormatSettings` | Membuka pengaturan format yang dipilih. |
| `cmdSave` | Menyimpan draft job ke antrean pada mode queue. |
| `cmdCancel` | Membatalkan perubahan editor. |

Pengaturan PNG, JPEG, dan DXF disalin sebagai **snapshot** per item. Perubahan default untuk pekerjaan berikutnya tidak otomatis mengganti snapshot job yang sudah ada.

Untuk PDF, item menyimpan nama preset dan pilihan compatibility. Isi preset tetap dimuat dari CorelDRAW saat export, sehingga preset tersebut harus tersedia dan perubahan isinya dapat memengaruhi hasil.

### Parent Directory

Ketika `chkParentDirectory` aktif, tujuan export mengikuti folder dokumen sumber item. `txbDirectory` dan `cmdBrowse` dinonaktifkan selama pilihan ini digunakan.

Pada queue lintas dokumen, setiap job mengikuti folder dokumennya masing-masing. Lokasi diperiksa kembali sebelum proses berjalan.

Dokumen harus sudah tersimpan untuk menggunakan Parent Directory. Jika belum, simpan dokumen terlebih dahulu atau gunakan directory manual yang sudah tersedia.

---

## Page Selection

`txbPage` menentukan page yang diproses sekaligus cara membaginya menjadi file hasil.

### Single and Grouped Pages

| Input `txbPage` | Hasil |
|---|---|
| Kosong | Menggunakan page aktif saat item queue disimpan. |
| `1` | Satu file dari page 1. |
| `1,3` | Satu PDF berisi page 1 dan 3. |
| `1-3` | Satu PDF berisi page 1 sampai 3. |
| `(1),(2,3)` | Dua PDF: page 1, lalu page 2 dan 3. |
| `{1-3}` | Tiga file terpisah, masing-masing dari satu page. |
| `{(1,2),3}` | Dua PDF: page 1 dan 2, lalu page 3. |

Kurung biasa `()` mengelompokkan page menjadi satu file. Kurung kurawal `{}` memisahkan daftar atau rentang menjadi output per page, dengan tetap memungkinkan grup `()` di dalamnya.

**Grup multi-page hanya didukung untuk PDF.** Untuk PNG, JPEG, dan DXF, setiap grup harus berisi satu page. Gunakan `{1-3}` atau `(1),(2),(3)` untuk menghasilkan tiga file terpisah.

Nomor page harus tersedia pada dokumen sumber. Rentang ditulis dari kecil ke besar, seperti `2-5`.

### Sync Groups

Tanda `;` membagi input page menjadi segmen sinkronisasi yang dapat dipasangkan dengan nilai nama tertentu.

Contoh:

`{1,2};{3,4}`

Input tersebut menghasilkan empat file: page 1 dan 2 berada pada segmen pertama, sedangkan page 3 dan 4 berada pada segmen kedua.

Segmen sinkronisasi dan jumlah file adalah dua hal yang berbeda. Satu segmen dapat menghasilkan beberapa file.

---

## File Naming

`txbName` mendukung nama langsung, replacement, nilai berurutan, prefix, suffix, dan kombinasi beberapa operasi.

Jika input kosong, macro menggunakan nama dokumen tanpa ekstensi. Ekstensi hasil ditambahkan sesuai `cmbExFormat`.

### Direct Name and Replacement

Dengan nama dokumen `Desain Lama.cdr`:

| Input `txbName` | Nama Dasar Hasil |
|---|---|
| Kosong | `Desain Lama` |
| `Hasil Final` | `Hasil Final` |
| `Lama/*Baru` | `Desain Baru` |

Token `/*` memisahkan teks yang dicari dan penggantinya. Replacement tidak membedakan huruf besar dan kecil.

### Sequence Values

Untuk nama dokumen `Order A.cdr` dan tiga file hasil:

`(/*A):1,2,3`

akan membentuk `Order 1`, `Order 2`, dan `Order 3`.

Template langsung juga dapat memakai bagian tetap dalam `{}` dan nilai yang berubah di luarnya:

| Template untuk Tiga Output | Nama Dasar Hasil |
|---|---|
| `{LABEL}_A,B,C` | `LABEL_A`, `LABEL_B`, `LABEL_C` |
| `{LABEL}_1,2,3` | `LABEL_1`, `LABEL_2`, `LABEL_3` |
| `{LABEL}_A1,A2,A3` | `LABEL_A1`, `LABEL_A2`, `LABEL_A3` |

Pada bentuk ini, `{}` menandai teks tetap dan tidak ikut tampil pada nama hasil. Nilai sequence dapat berupa huruf, angka, atau gabungannya.

Contoh menggunakan daftar nilai eksplisit. Satu pola seperti `{LABEL}_A` tetap konstan untuk seluruh output; nama berikutnya tidak otomatis berubah menjadi B atau C. Daftar beberapa nilai harus cocok dengan jumlah file hasil, bukan jumlah page fisik di dalam satu PDF.

### Literal Text and Spaces

Koma di luar `{}` menjadi delimiter nilai, termasuk ketika berada di antara angka. Koma di dalam `{}` tetap menjadi teks literal.

Spasi setelah koma dipertahankan sebagai bagian nilai pada template langsung, command replacement, dan daftar affix. Spasi pemisah di sekitar `;` dan `|` diabaikan.

| Template untuk Dua Output | Nama Dasar Hasil |
|---|---|
| `{A}1,2` | `A1`, `A2` |
| `{A}1, 2` | `A1`, `A 2` |
| `{ORDER, A}_1,2` | `ORDER, A_1`, `ORDER, A_2` |
| `{NAMA} A{ END},B` | `NAMA A END`, `NAMA B END` |

Gunakan `1,2,3` ketika tidak menginginkan spasi tambahan. Karakter `_` adalah teks biasa, sedangkan `/` sendiri bukan token nama; replacement menggunakan `/*`.

Command dengan pola tetap seperti `(/*TOTAL 3): {}1` mengganti `TOTAL 3` menjadi `1` pada setiap output. Bentuk terstruktur ini tidak otomatis mengambil nomor page.

### Prefix and Suffix

Marker `__` digunakan untuk menambahkan teks pada nama dasar.

Dengan nama dasar `Desain`:

| Input | Hasil |
|---|---|
| `FINAL__` | `FINALDesain` |
| `__FINAL` | `DesainFINAL` |
| `FINAL___` | `FINAL_Desain` |
| `___FINAL` | `Desain_FINAL` |

Dua underscore merupakan marker. Underscore tambahan menjadi karakter literal, sehingga pemisah pada hasil perlu ditulis secara eksplisit.

### Naming Pipeline

Tanda `|` menggabungkan beberapa operasi penamaan.

Contoh untuk dua file hasil:

`{LABEL}_A,B | ___FINAL`

menghasilkan `LABEL_A_FINAL` dan `LABEL_B_FINAL`.

Parser menyelesaikan operasi pembentukan nama utama terlebih dahulu, kemudian menerapkan prefix dan suffix. Karena itu, pipeline tidak sepenuhnya menjalankan semua jenis operasi dari kiri ke kanan tanpa pengelompokan.

Contoh gabungan prefix, blok nama, dan suffix untuk empat output:

`Contoh __ | {A}1,2; {B}1,2 | __ Fix`

Hasilnya adalah `Contoh A1 Fix`, `Contoh A2 Fix`, `Contoh B1 Fix`, dan `Contoh B2 Fix`.

### Name and Page Synchronization

Pada replacement dan affix, `;` dapat memilih nilai berdasarkan segmen sinkronisasi page.

Dengan nama dokumen `Order X.cdr`:

| Pengaturan | Input |
|---|---|
| `txbPage` | `{1,2};{3,4}` |
| `txbName` | `(/*X):A;B` |

Dua output pertama memakai nama dasar `Order A`, sedangkan dua berikutnya memakai `Order B`. Karena ada nama dasar yang berulang dalam satu item, suffix nomor ditambahkan pada hasilnya.

Template langsung berbasis brace juga mendukung blok seperti `{A}1,2;{B}1,2`. Bentuk ini menghasilkan `A1`, `A2`, `B1`, dan `B2` secara berurutan; pembagiannya mengikuti jumlah nilai setiap blok, bukan batas segmen `txbPage`.

Panjang blok boleh berbeda. `{A}1,2,3;{B}X,Y` menghasilkan lima nama: `A1`, `A2`, `A3`, `BX`, dan `BY`. Jumlah total nilai seluruh blok harus sama dengan jumlah file output; `{A}1;{B}1` berarti tepat dua nama, bukan dua pola yang diulang tanpa batas.

Daftar nilai nama divalidasi terhadap jumlah output atau jumlah segmen yang relevan. Pola nama dan pengelompokan page perlu disusun sebagai satu kesatuan.

---

## Layer Selection

Checkbox `chkLayer1`, `chkLayer2`, dan `chkLayer3` mengatur flag print/export layer lokal untuk job terpilih.

Default item baru adalah `Layer 1` aktif, sedangkan `Layer 2` dan `Layer 3` tidak dipilih.

Saat job diproses:

- layer lokal yang dipilih diaktifkan untuk print/export;
- layer lokal biasa lainnya dinonaktifkan sementara;
- seluruh master layer mengikuti flag print/export pengguna;
- layer otomatis lokal tertentu juga mengikuti flag pengguna;
- flag layer yang diubah dipulihkan setelah pemrosesan item.

Layer otomatis lokal yang dikecualikan dari perubahan adalah:

- `scpro2_printmargin`
- `scpro2_printonly`
- `scpro2_regmarks`
- `Regmark`

Layer yang dipilih harus ditemukan tepat satu kali sebagai layer lokal pada setiap page terkait. Nama `Layer 1`, `Layer 2`, dan `Layer 3` mengikuti penulisan persis di source.

Semua checkbox boleh tidak dipilih jika masih ada master layer atau layer otomatis dengan print/export aktif. Pengaturan ini tidak mengubah visibility, editability, atau urutan layer.

---

## PDF Settings

Export PDF dijalankan melalui `PublishToPDF` dan diatur menggunakan `PDFSettings`.

Pengaturan utama:

- `cmbPDFPresets`: memilih preset bawaan atau custom preset CorelDRAW.
- `cmbCompatibility`: memilih compatibility PDF yang tersedia pada form.

PDF dapat berisi satu page atau gabungan beberapa page sesuai `txbPage`.

`cmdAddPreset` dan `cmdRemovePreset` masih dinonaktifkan. Pengelolaan preset lanjutan belum tersedia melalui form ini.

Item queue menyimpan pilihan preset dan compatibility, tetapi tidak menyimpan salinan penuh isi preset PDF. Pastikan preset yang dipilih masih dapat dimuat saat export dijalankan.

---

## PNG Settings

`PNGSettings` menyediakan pengaturan berikut:

- preset XML melalui `cmbPreset`;
- Black and White, Grayscale, Paletted, atau RGB melalui `cmbColorMode`;
- transparency melalui `chkTransparency`;
- warna matte/background melalui `cmbBgColor` ketika transparency tidak aktif;
- antialias melalui `chkAntialiased`;
- penyertaan profil warna melalui `chkEmbedColorProfile`;
- interlace melalui `chkInterlaced`;
- batas area export mengikuti page melalui `chkCropToPageOnExport`;
- resolusi melalui `cmbResolution`.

Resolusi menerima bilangan bulat 1 sampai 1200 dpi, dengan default 300 dpi.

Untuk mode Paletted, runner mendukung palette `optimized`, jumlah warna 2 sampai 256, serta dithering `none` atau `floyd-steinberg` melalui atribut preset yang dikenali. Palette dan dithering lain belum didukung.

---

## JPEG Settings

`JPEGSettings` menyediakan pilihan Grayscale, RGB, dan CMYK, serta pengaturan resolusi, background, antialias, penyertaan profil warna, dan crop ke page.

Pengaturan tambahan mencakup:

| Kontrol | Fungsi |
|---|---|
| `cmbSubFormat` | Memilih `Standard (4:2:2)` atau `Optional (4:4:4)`. |
| `cmbQuality` | Menentukan quality dari 0 sampai 100%. |
| `cmbBlur` | Menentukan blur dari 0 sampai 100%. |
| `optDocColorSetting` | Menggunakan alur warna dokumen. |
| `optColorProofSetting` | Menggunakan color proof settings dari active view saat export. |
| `chkOptimize` | Mengaktifkan opsi optimized pada filter JPEG. |
| `chkProgressive` | Mengaktifkan progressive JPEG. |

JPEG menggunakan background/matte dan tidak menyediakan transparency.

Nilai Quality dipetakan berlawanan dengan Compression pada filter CorelDRAW: Quality 100% menggunakan Compression 0. Pemetaan ini telah diperbaiki pada perubahan tanggal 2026-09-10.

Pilihan penggunaan color proof disimpan per item; konfigurasi proof aktual dibaca dari active view ketika proses dijalankan.

---

## DXF Settings

`DXFSettings` menyediakan pengaturan versi AutoCAD, unit, format bitmap, text, dan unmapped fills.

| Kontrol | Perilaku |
|---|---|
| `cmbAutoCADVersion` | Memilih versi format AutoCAD/DXF. |
| `cmbAutoCADUnits` | Menentukan unit hasil. |
| `cmbAutoCADBitmapAs` | Memilih tipe bitmap untuk filter DXF. |
| `optTextAsCurves` / `optTextAsText` | Menentukan penyimpanan text sebagai curves atau text. |
| `optAutoCADColor` / `optAutoCADUnfilled` | Mengatur penanganan unmapped fills. |
| `cmbAutoCADColor` | Memilih warna RGB untuk unmapped fills. |
| `cmbAutoCADCurvesAs` | Menyimpan pilihan Polycurves atau Splines; belum diterapkan ke export. |
| `cmbAutoCADTolerance` | Menyimpan tolerance; belum diterapkan ke export. |

**R2007 menjadi checkpoint yang tercatat berhasil digunakan.** Pilihan versi lain tersedia pada form, tetapi masih memerlukan validasi lanjutan.

`CurvesAs` dan `Tolerance` sudah masuk ke pengaturan dan snapshot, tetapi belum diterapkan karena pemetaan properti terkait belum tersedia pada filter yang digunakan source ini.

Default DXF pada source adalah R2007 dengan unit **Centimeters**. Periksa unit sebelum export agar sesuai dengan workflow tujuan.

---

## Presets and Saved Settings

Preset PNG dan JPEG dibaca dari file XML di folder:

`Documents\Corel\Corel Content\Export Presets`

Pencarian mengikuti lokasi Documents pengguna serta beberapa lokasi OneDrive yang dikenali. Parser hanya menerapkan atribut dan nilai yang didukung oleh masing-masing format.

Pengaturan macro disimpan menggunakan registry VBA pada aplikasi `RinCorelMacros`, section `ExportRelatedMacro`, antara lain directory manual terakhir, pilihan Parent Directory, format terakhir pada alur single export, serta pengaturan format.

Registry digunakan sebagai default ketika diperlukan. Antrean job sendiri tidak disimpan sebagai antrean permanen.

---

## File Name Conflicts

Penanganan nama perlu dibedakan berdasarkan lokasi benturannya:

| Kondisi | Perilaku |
|---|---|
| Nama dasar berulang dalam satu item | Macro memberi suffix seperti `Desain (1)` dan `Desain (2)`. |
| Dua hasil queue mempunyai path tujuan yang sama | Validasi menghentikan antrean sebelum export dimulai. |
| File dengan path tujuan yang sama sudah ada di disk | File tersebut dapat ditimpa saat hasil disalin ke tujuan. |

Suffix duplikat dibentuk berdasarkan nama hasil dalam item, bukan dengan mencari nomor kosong di folder tujuan. Tidak ada pilihan interaktif duplicate, overwrite, atau cancel untuk benturan dengan file lama pada jalur queue.

---

## Basic Workflow

1. Buka dokumen CorelDRAW yang akan diekspor.
2. Jalankan `ExportRelatedWizard` untuk membuka Main UserForm.
3. Aktifkan dokumen sumber, lalu gunakan **Add** melalui `cmdAddSetting`.
4. Tentukan directory manual atau aktifkan `chkParentDirectory`.
5. Pilih format melalui `cmbExFormat`.
6. Isi `txbPage` dan `txbName` sesuai pembagian file yang dibutuhkan.
7. Buka `cmdFormatSettings`, periksa opsi format, lalu simpan.
8. Simpan job melalui `cmdSave`.
9. Pilih kombinasi layer job tersebut pada Main UserForm.
10. Ulangi untuk pekerjaan atau dokumen lainnya.
11. Periksa summary pada `lbxSettingLists`, lalu tekan **Export**.

Main UserForm dibuka secara `vbModeless`, sehingga user dapat berpindah dokumen saat menyiapkan antrean. Editor job dan pengaturan format dibuka secara `vbModal`.

---

## Processing and Recovery

Sebelum membuat file atau mengubah flag layer, runner memvalidasi seluruh antrean: dokumen sumber, directory, format, snapshot, page, pola nama, pilihan layer, dan benturan path antarhasil.

Export kemudian dijalankan berurutan. Hasil ditulis ke file sementara dan diperiksa keberadaan serta ukurannya sebelum disalin ke tujuan. Runner PNG dan JPEG juga memeriksa penanda awal dan akhir format file; pemeriksaan ini bukan validasi visual seluruh isi gambar.

Setelah setiap item, macro berusaha memulihkan flag print/export layer, active page, active layer, dan selection. Setelah antrean selesai atau dihentikan, dokumen aktif awal juga dipulihkan.

Jika terjadi kegagalan saat export, antrean berhenti dan menampilkan item yang gagal serta jumlah file yang sudah selesai. File yang telah berhasil diekspor tetap berada di tujuan; seluruh proses bukan transaksi yang otomatis membatalkan hasil sebelumnya.

File sementara yang belum dapat dihapus dicoba kembali dan dilaporkan sebagai peringatan cleanup bila masih tersisa. Tombol pengubahan antrean dan penutupan form dinonaktifkan selama export berjalan.

---

## Project Structure

| File atau Kelompok Source | Tanggung Jawab |
|---|---|
| `ExportRelatedMasterWizard.bas` | Entry point `ExportRelatedWizard`. |
| `ExportRelatedMenu.vba` | Main UserForm, antrean, checkbox layer, dan pemanggilan export. |
| `ExportRelatedSettings.vba` | Editor job serta alur single export. |
| `ExportSettings.cls` | Directory, format, page, template nama, dan Parent Directory. |
| `ExportSettingItem.cls` | Dokumen sumber, snapshot format, pilihan layer, validasi item, dan summary. |
| `ExportPageParser.cls` | Page range, grouping, pemisahan output, dan sync group. |
| `ExportTemplateParser.cls` | Replacement, sequence, template langsung, prefix/suffix, pipeline, dan suffix duplikat. |
| `ExportQueueRunner.cls` | Validasi antrean, eksekusi lintas dokumen, serta pemulihan state. |
| `ExportLayerState.cls` | Validasi, penerapan sementara, dan pemulihan print/export layer. |
| `ExportRunner.cls` | Pemilihan runner format dan penerapan filter DXF. |
| `ExportTempFiles.cls` | Cleanup file sementara dan percobaan ulang penghapusan. |
| `PDFSettings.vba`, `PDFExport.cls` | Pemilihan preset/compatibility dan PublishToPDF. |
| `PNGSettings.vba`, `PNGPreset.cls`, `PNGPresetReader.cls`, `PNGSettingsStore.cls`, `PNGExportRunner.cls` | Form, preset XML, snapshot, dan eksekusi PNG. |
| `JPEGSettings.vba`, `JPEGPreset.cls`, `JPEGPresetReader.cls`, `JPEGSettingsStore.cls`, `JPEGExportRunner.cls` | Form, preset XML, snapshot, dan eksekusi JPEG. |
| `DXFSettings.vba`, `DXFExportSettings.cls`, `DXFSettingsStore.cls` | Form, data pengaturan, dan snapshot DXF. |
| `BitmapExportSupport.bas` | Helper bersama untuk preset XML, warna, resolusi, registry, dan opsi bitmap. |
| `Changelog.log` | Riwayat fitur, perubahan behavior, perbaikan, dan status pengembangan. |

---

## Setup and Notes

Source ini ditujukan untuk project VBA di CorelDRAW pada Windows, dengan rujukan API CorelDRAW 2024 pada implementasi filter tertentu. Kompatibilitas lintas versi CorelDRAW belum dinyatakan teruji secara menyeluruh.

Catatan referensi token membedakan pemeriksaan simulasi parser dan pengujian runtime CorelDRAW. Perubahan blok nama dan spasi literal telah diperiksa melalui harness adaptasi parser, tetapi catatan tersebut belum mengonfirmasi pengujian runtime untuk seluruh kombinasi. Contoh historis dengan spasi setelah koma perlu disesuaikan jika hasil yang diinginkan tidak menggunakan spasi tambahan.

Paket source berisi standard module `.bas`, class module `.cls`, dan code-behind UserForm `.vba`. Paket ini belum menyertakan project `.gms` siap pakai maupun file designer UserForm `.frm`/`.frx`.

Untuk merakit project, modul dan class perlu dimasukkan ke project VBA, lalu UserForm beserta kontrolnya disiapkan dengan nama yang sesuai source. Entry point utama adalah `ExportRelatedWizard`.

`cmdHintName` dan `cmdHintPage` merujuk ke `NameHintMenu` dan `PageHintMenu`. Source kedua form bantuan tersebut belum disertakan dalam paket ini, sehingga perlu dilengkapi atau pemanggilannya disesuaikan saat merakit project.

Alur single export tetap tersedia ketika `ExportRelatedSettings` dibuka tanpa `BeginQueueEdit`; pada alur tersebut, `cmdSave` langsung menjalankan export. Pengaturan layer per job dan validasi seluruh antrean merupakan perilaku jalur queue.

Directory tujuan harus sudah tersedia. ExportRelated berfokus pada export dan pengaturan pekerjaan, bukan pembuatan struktur folder produksi atau penyimpanan dokumen CDR.

---

## Current Scope

ExportRelated saat ini mencakup:

- Export Queue lintas dokumen;
- PDF, DXF, PNG, dan JPEG;
- directory, format, page, nama, dan layer per item;
- snapshot PNG/JPEG/DXF serta pilihan preset PDF per item;
- page grouping, split output, dan sync group;
- template nama alfanumerik, replacement, prefix/suffix, dan pipeline;
- Parent Directory;
- pengaturan bitmap, profil warna, serta opsi format;
- validasi queue, pemulihan state dokumen, dan cleanup file sementara.

Pengembangan DXF di luar checkpoint R2007, penerapan CurvesAs/Tolerance, dan pengelolaan preset PDF lanjutan masih memerlukan pekerjaan berikutnya.

---

## Feedback and Development

Source boleh dipelajari dan dikembangkan, dan issue/feedback tentang bug, edge case, CorelDRAW API, architecture, atau improvement sangat dihargai.
