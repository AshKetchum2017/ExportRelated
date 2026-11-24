# Referensi Algoritma ExportRelated

ExportRelated menyusun rencana halaman, nama, dan path sebelum menjalankan filter CorelDRAW. Rujukan ini menjelaskan alur implementasi, alasan pemisahan tahap, dan batas yang perlu diperhatikan saat memakai ulang polanya. [Panduan penggunaan](../README.md) · [Sintaks nama dan halaman](ExportNameTokens.md).

## Perencanaan output

### Pengelompokan halaman dan ekspansi token

`ResolvePageGroups` mengubah input menjadi daftar grup output berurutan. Satu grup memuat halaman untuk satu file, sedangkan indeks sinkronisasi menghubungkan beberapa grup dengan satu segmen nama. Pemisahan ini memungkinkan pilihan seperti `{1-3};(4-7)` menghasilkan tiga file per halaman dan satu PDF empat halaman tanpa menyamakan jumlah file dengan jumlah halaman. Untuk rentang eksplisit, angka di luar dokumen menyebabkan validasi gagal. Deret dinamis dapat menghasilkan bagian yang tersedia dengan peringatan. [`ResolvePageGroups` dan `SyncGroupIndex`](../src/classes/ExportPageParser.cls).

`ExportTokenPlan.Prepare` memasangkan grup dengan nama, memvalidasi panjang daftar dan batas blok, serta mengekspansi deret sesuai indeks output. Parser nama menangani replacement dan affix, sementara `ExportDynamicTokens` menangani deret dan wildcard. Editor memakai rencana yang sama untuk preview dan runner menghitungnya kembali dari dokumen sumber saat proses berjalan. Dengan demikian, perubahan dokumen atau input masih dapat memengaruhi hasil preflight. [`Prepare`](../src/classes/ExportTokenPlan.cls) · [Parser nama](../src/classes/ExportTemplateParser.cls) · [Token dinamis](../src/classes/ExportDynamicTokens.cls) · [Editor](../src/forms/ExportRelatedSettings.vba).

### Nama unik dan benturan antrean

Setelah nama dibentuk, parser menambahkan suffix numerik pada nama berulang dalam satu item, termasuk kemunculan pertama. Preflight antrean memeriksa path akhir lintas item dengan perbandingan tanpa membedakan kapitalisasi. Benturan lintas item menghentikan antrean sebelum penulisan file. Preview menampilkan nama pertama untuk item yang sedang diedit, sehingga benturan antarjob masih diperiksa saat antrean dijalankan. [Penanganan nama](../src/classes/ExportTemplateParser.cls) · [`Run` dan `BuildPaths`](../src/classes/ExportQueueRunner.cls).

## Antrean dan pemulihan state

### Snapshot pengaturan dan preflight

Satu `ExportSettingItem` mengikat dokumen sumber, pengaturan, dan pilihan layer. PNG, JPEG, dan DXF memakai snapshot pengaturan per item. Untuk PDF, item menyimpan identitas preset dan compatibility, lalu runner memuat preset saat ekspor. Perubahan isi preset PDF dapat memengaruhi job yang sudah dibuat. Runner mengklon item dan memvalidasi dokumen, pengaturan, grup halaman, layer, serta path seluruh antrean sebelum menjalankan item pertama. [Model item](../src/classes/ExportSettingItem.cls) · [Runner antrean](../src/classes/ExportQueueRunner.cls) · [Runner PDF](../src/classes/PDFExport.cls).

### Layer sementara dan penulisan hasil

Runner mencatat flag print/export layer yang akan diubah serta halaman, layer aktif, dan selection dokumen sumber. Setelah ekspor berhasil atau gagal, runner mencoba memulihkan state item dan dokumen aktif awal. Kegagalan pemulihan ikut dilaporkan. [`Apply` dan `Restore`](../src/classes/ExportLayerState.cls) · [`ExecuteItem` dan `RestoreItemState`](../src/classes/ExportQueueRunner.cls).

Pada jalur antrean, runner menulis ke file sementara di folder tujuan, memeriksa keberadaan dan ukuran file, lalu menyalinnya ke nama akhir. PNG dan JPEG juga memeriksa penanda format pada file bitmap sebelum menyalinnya. Pola ini menunda penimpaan file tujuan sampai hasil sementara lolos pemeriksaan. Penyalinan akhir tetap dapat gagal, dan penghentian antrean tidak membatalkan file yang sudah selesai. File sementara yang belum terhapus masuk daftar percobaan ulang. [Runner antrean](../src/classes/ExportQueueRunner.cls) · [Runner PNG](../src/classes/PNGExportRunner.cls) · [Runner JPEG](../src/classes/JPEGExportRunner.cls) · [Cleanup](../src/classes/ExportTempFiles.cls).

## Penerapan API per format

### PDF, bitmap, dan DXF

PDF memakai `Document.PDFSettings` untuk menerapkan preset, compatibility, dan rentang halaman sebelum `PublishToPDF`. Tampilan preset dapat memakai nama yang mudah dibaca, sementara pemuatan memakai key internal CorelDRAW bila tersedia. [Pengaturan PDF](../src/forms/PDFSettings.vba) · [`PublishDocument`](../src/classes/PDFExport.cls).

PNG, JPEG, dan DXF memakai `ExportEx`, mengisi pengaturan filter, lalu memanggil `Finish`. Untuk JPEG, nilai UI tetap menyatakan `Quality` dan runner mengubahnya menjadi `Compression = 100 - Quality` saat mengisi filter. Untuk DXF, periksa `$ACADVER` pada file jika versi hasil penting. Dukungan enum dan setter bergantung pada host dan filter. [Runner PNG](../src/classes/PNGExportRunner.cls) · [Runner JPEG](../src/classes/JPEGExportRunner.cls) · [`ExportDXFDocument`](../src/classes/ExportRunner.cls).

## Referensi publik

- **Dipakai** · CorelDRAW 2024 SDK. Referensi untuk mengatur PDF per dokumen dan menjalankan publikasi. [Document.PDFSettings](https://community.coreldraw.com/sdk/api/draw/25/p/ivgdocument.pdfsettings?lang=js) · [Document.PublishToPDF](https://community.coreldraw.com/sdk/api/draw/25/m/document.publishtopdf?lang=js).
- **Dipakai** · CorelDRAW 2024 SDK. Referensi untuk memperoleh `ExportFilter` dan menerapkan parameter filter sebelum menyelesaikan ekspor. [Document.ExportEx](https://community.coreldraw.com/sdk/api/draw/25/m/document.exportex?lang=cs).

## Dokumen dan bukti pendukung

### Tersedia di Git

- [Panduan penggunaan ExportRelated](../README.md) menjelaskan alur operator, kontrol, dan batas per format.
- [Referensi token nama dan halaman](ExportNameTokens.md) memuat grammar, contoh, dan matriks regresi.
- [Changelog](../Changelog.log) mencatat perubahan yang terlihat pengguna.

### Lokal — khusus komputer pengembang

Path dihitung relatif dari `ExportRelated/docs`.

- Keputusan lintas proyek dan batas reuse. `../../.agents/references/coreldraw-integration.md`.
- Riwayat investigasi, hasil harness, dan konfirmasi pengguna. `../../Development.log` dan `../../KDDK.log`.
- Checklist antrean, catatan filter DXF, dan harness token. `../ExportQueueVerification.md`, `../DXFSettings.md`, dan `../VerifyExportTokens.ps1`. Hasil harness adaptasi, compile VBE, dan ekspor native memberikan bukti pada tingkat yang berbeda.
