# Export Queue: pemasangan dan verifikasi

Status akhir 2026-09-08: scope Export Queue selesai dan diterima user setelah
pengujian di CorelDRAW, dengan konfirmasi "Untuk scope ini, uji coba berjalan sesuai
ekspektasi." Bukti sesi mencakup tiga output JPG/PDF/PNG dalam satu queue serta
konfirmasi aturan layer. Pemeriksaan statis/harness dilakukan oleh agent; agent
tidak menjalankan compile VBE atau export CorelDRAW langsung. Rincian setiap
kombinasi skenario tidak dicatat terpisah.

## Pasang kode dan rename kontrol

Perubahan file workspace tidak otomatis memperbarui GMS. Hentikan macro sebelum
memperbarui code-behind di VBE, lalu gunakan pemetaan berikut.

| UserForm | Sumber code-behind | Rename properti `(Name)` kontrol |
| --- | --- | --- |
| `ExportRelatedSettings` | `ExportRelatedSettings.vba` | `txtDirectory` -> `txbDirectory`, `txtPage` -> `txbPage`, `txtName` -> `txbName` |
| `ExportRelatedMenu` | `NewMasterMenu.vba` | `txtSettingLists` -> `lbxSettingLists` (tetap MSForms.ListBox) |

Menu memakai referensi langsung `Me.lbxSettingLists`, mempertahankan perbaikan
startup sebelumnya. Jangan mengembalikannya ke pencarian `Me.Controls(...)`.

Tambahkan Class Module baru dengan `(Name)` persis `ExportQueueRunner` dan
`ExportLayerState`, lalu salin isi file `.cls` masing-masing. File ini headerless.
Perbarui isi Class Module existing berikut tanpa membuat modul duplikat:

- `ExportSettingItem`
- `ExportSettings`
- `ExportPageParser`
- `ExportTemplateParser`
- `ExportRunner`
- `PDFExport`
- `PNGExportRunner`
- `JPEGExportRunner`

Entry point `ExportRelatedWizard` tetap memakai `ERMasterWizard.bas` existing.
Jangan import file backup. Jalankan **Debug > Compile** setelah seluruh kode dan
nama kontrol sinkron.

## Perilaku queue

- Directory valid terakhir disimpan ke key existing `LastDirectory` setelah Browse,
  setelah edit manual `txbDirectory` selesai (AfterUpdate), dan saat Save item queue.
  Add berikutnya membaca directory tersebut. Modify tetap membuka directory milik
  item yang dipilih; setiap item mempertahankan directory sendiri. Input kosong,
  path file, atau folder yang tidak tersedia tidak mengganti nilai registry.
  Cancel membuang draft item, tetapi tidak membatalkan preferensi directory yang
  sudah disimpan setelah Browse/edit manual.
- Semua item diperiksa sebelum perubahan layer atau pembuatan file: dokumen masih
  terbuka, directory/page/nama valid, setting format tersedia, dan layer lokal
  terpilih tersedia dengan nama persis pada setiap page target. Pilihan kosong
  diperbolehkan jika page target memiliki layer otomatis lokal atau ada master
  layer dengan print/export aktif. Tanpa keduanya, validasi berhenti. Layer terpilih yang
  hilang/duplikat atau tujuan file identik di dalam queue juga menghentikan proses.
- Item menggunakan dokumen sumber, directory, format, page, template nama, snapshot
  PNG/JPG, pilihan preset/compatibility PDF, dan checkbox layer masing-masing.
  PDF memuat preset bernama tersebut saat export; perubahan isi preset di luar
  macro tetap akan terbaca. DXF menggunakan opsi default existing.
- Layer lokal yang dicentang diaktifkan print/export sementara. Layer lokal lain
  dikecualikan selain empat nama otomatis di bawah. Flag awal disimpan sebelum
  perubahan dan dipulihkan setelah setiap item pada jalur sukses maupun error.
  Visibility/editability tidak diubah.
- Pengecualian berdasarkan nama persis: `scpro2_printmargin`, `scpro2_printonly`,
  `scpro2_regmarks`, dan `Regmark`. Keempat layer otomatis ini, baik lokal maupun
  master, tetap mengikuti print/export milik pengguna tanpa dipengaruhi checkbox.
  Macro tidak mengubah flag mereka: aktif ikut export, disabled tidak ikut.
- Semua master layer, tanpa pembatasan nama, mengikuti print/export pengguna:
  aktif tetap ikut export, disabled tetap tidak ikut. Macro tidak mengubah flag
  master saat Apply/Restore. Aturan cakupan page master tetap mengikuti dokumen.
- Page aktif, layer aktif, selection, dan dokumen aktif dipulihkan. Jika pemulihan
  gagal, error dilaporkan dan queue berhenti; item berikutnya tidak dijalankan.
- Hasil ditulis sementara di directory tujuan sebelum disalin ke nama akhir.
  Error menghentikan queue dan melaporkan item serta jumlah file yang sudah selesai.
  File yang sudah selesai tetap ada; queue tetap tersedia untuk Modify/Remove.
- Preset PDF yang hilang akan menghasilkan error saat item itu diekspor. Queue
  tidak membuat preset global pengganti. Registry tidak digunakan sebagai pengganti
  snapshot item yang kosong. Alur single tetap memakai default registry existing.

## Skenario regresi untuk perubahan berikutnya

Daftar ini disimpan sebagai acuan uji ulang; penerimaan scope saat ini sudah dicatat
di atas. Pekerjaan berikutnya (DXF atau TIF) belum dipilih.

Gunakan dokumen uji dengan dua page dan objek yang mudah dibedakan pada `Layer 1`,
`Layer 2`, `Layer 3`, satu layer lokal tambahan, serta satu master layer.

1. Buka menu, Add dua item, Modify item pertama, lalu Remove item kedua. Periksa
   summary, index pilihan, dan state checkbox; item baru default Layer 1.
2. Buat dua item berbeda format/page/nama/directory/setting. Ekspor dan periksa
   masing-masing hasil, termasuk dua item PNG/JPG dengan resolusi berbeda.
3. Atur Layer 1 non-printable, Layer 2 printable. Ekspor item Layer 1 saja, lalu
   item Layer 2 saja. Objek layer terkait, layer otomatis lokal aktif, dan master aktif boleh muncul; flag awal semua
   layer harus kembali. Ulangi dengan kombinasi Layer 1+2 dan layer hidden.
4. Pastikan layer lokal tambahan tidak masuk ke hasil PDF, DXF, PNG, dan JPG,
   kecuali empat nama layer otomatis dengan print/export aktif. Semua master aktif
   harus tetap ikut, sedangkan master disabled tidak ikut. Uji juga PDF grup
   multi-page, keempat nama otomatis secara terpisah, flag aktif/nonaktif, serta
   semua checkbox kosong. Layer otomatis disabled harus tetap disabled.
5. Item terakhir memilih layer yang tidak ada pada salah satu page target: seluruh
   queue harus berhenti sebelum ada file baru. Ulangi dengan semua checkbox kosong
   tanpa layer otomatis lokal aktif maupun master aktif,
   nama berbeda kapitalisasi, dokumen sumber ditutup, dan page yang dihapus.
6. Picu kegagalan export/copy, misalnya file tujuan dikunci aplikasi lain. Periksa
   error item, jumlah hasil sebelumnya, penghentian queue, serta pemulihan seluruh
   flag layer/page/selection/dokumen aktif.
7. Buka `ExportRelatedSettings` langsung di luar queue. Uji single PDF, DXF, PNG,
   dan JPG untuk memeriksa workflow registry dan hasil yang sebelumnya stabil.

## Pemeriksaan yang sudah dilakukan

- Perbandingan dengan snapshot source sebelum perubahan: form single hanya berubah
  pada rename dan pemanggilan helper nama yang dipindahkan tanpa perubahan logic;
  parser page dan ExportSettings hanya mengalami rename kontrol.
- Pipeline PNG/JPG setelah pemilihan sumber setting serta dispatch single
  PDF/DXF/bitmap tidak berubah. Pemeriksaan batas procedure, deklarasi duplikat,
  panjang baris, dan referensi kontrol aktif lulus.
- Harness VBScript yang mengadaptasi source `ExportLayerState` memakai objek
  simulasi: isolasi layer per item, exact name, validasi kosong/
  missing layer, partial apply failure, restore failure, dan retry restore lulus.
- Simulasi tambahan untuk keempat nama otomatis pada local/master layer lulus:
  aktif/nonaktif, checkbox kosong, kombinasi dengan layer terpilih, dan pemulihan
  setelah partial apply gagal. Setter flag layer otomatis tidak dipanggil.
- Setelah cakupan diperluas ke semua master layer, simulasi nama master bebas,
  aktif/nonaktif, master-only multi-page, checkbox kosong, pengecualian lokal, dan
  partial apply failure lulus. Tidak ada setter flag master yang dipanggil.
- Harness PDF dengan objek simulasi: fallback preset single, preset/compatibility
  per item, page range, tidak membaca registry saat queue, serta propagasi error
  preset/compatibility/publish lulus.

Harness tersebut tidak menjalankan `ExportQueueRunner` end-to-end, MSForms, filter
CorelDRAW, atau compile VBA. Penerimaan runtime scope berasal dari pengujian dan
konfirmasi user, bukan dari harness tersebut.
