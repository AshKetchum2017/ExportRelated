# ExportEx E_FAIL Debug Checklist

Tujuan: membantu mengidentifikasi penyebab `E_FAIL` pada `ExportEx` dengan kode `-2147467259` (`Method 'ExportEx' of object 'IVGDocument' failed`).

## 1. Validasi dokumen dan halaman aktif

- Apakah `ActiveDocument` tidak `Nothing`?
- Apakah dokumen benar-benar terbuka dan siap diekspor?
- Apakah halaman target valid dan ada di `doc.Pages`?
- Apakah `txbPage` atau `pageValues` memuat halaman yang benar-benar ada?
- Apakah `pageNumber` pernah < 1 atau > `doc.Pages.Count`?

## 2. Validasi path dan output file

- Apakah folder tujuan benar-benar ada?
- Apakah path tidak berakhir dengan karakter yang invalid?
- Apakah output file name valid untuk sistem Windows?
- Apakah file output sedang dibuka oleh program lain atau terkunci?
- Apakah file lama masih tersisa dan mengganggu proses export?

## 3. Validasi format ekspor

- Apakah `cmbExFormat` menghasilkan format yang valid untuk CorelDRAW yang aktif?
- Apakah `cdrFilter` yang dipilih benar-benar didukung dalam versi CorelDRAW ini?
- Apakah `ResolveExportFormat` menghasilkan nilai yang konsisten sesuai kebutuhan?
- Apakah kombinasi `outputPath` + `GetExportExtension(exportFormat)` sesuai dengan format yang dipilih?

### Default PDF saat ini

- Output color: CMYK
- Color management: Use document color settings
- Embed color profiles: True
- Encoding: Binary
- Compression: JPEG, quality 100%
- Downsampling: Color 600 DPI, Grayscale 300 DPI, Monochrome 100 DPI
- Embed fonts in document: True
- Embed base 14 fonts: True
- Convert TrueType to Type 1: True
- Subset fonts: Under 80%
- EPS files: PostScript
- Compatibility: Acrobat DC

## 4. Validasi argumen `ExportEx`

- Apakah pemanggilan `doc.ExportEx(outputPath, exportFormat, cdrCurrentPage)` sesuai dengan kebutuhan?
- Apakah untuk ekspor multi-page, `cdrCurrentPage` bukan kandidat yang membingungkan?
- Apakah fungsi seharusnya menggunakan page tertentu, bukan page aktif?
- Apakah argumen `cdrCurrentPage` menjadi tidak tepat saat halaman target bukan aktif?

## 5. Validasi logika template dan nama file

- Apakah nama file hasil generate unique dan aman?
- Apakah template `txbName` menghasilkan string kosong, terlalu panjang, atau karakter ilegal?
- Apakah nama file duplikat dapat memicu problem tertentu di path yang sama?
- Apakah `BuildExportFileName` menghasilkan output yang valid untuk tiap page?

## 6. Validasi state CorelDRAW

- Apakah dokumen masih dalam state valid setelah operasi sebelumnya?
- Apakah ada object atau layer yang menyebabkan export gagal karena state dokumen tidak lengkap?
- Apakah macro mengeksekusi export dalam kondisi yang terlalu cepat setelah operasi lain?
- Apakah proses `On Error Resume Next` menutupi error yang sebenarnya perlu diperiksa secara lebih spesifik?

## 7. Pola yang harus dicatat saat debugging

Catat hal-hal berikut jika error terjadi:

- nama dokumen
- page aktif saat ekspor
- isi `txbPage`
- isi `txtName`
- format ekspor yang dipilih
- path lengkap output
- jumlah page yang diproses
- apakah file output berhasil dibuat atau tidak
- apakah error terjadi pada single page atau multiple page

## 8. Skenario uji minimal

Lakukan test secara terpisah:

1. Single page, page aktif, nama default
2. Single page, page manual, nama custom
3. Multi page, `txtPage = (1,2,3)`
4. Multi page, template sequence seperti `(/*A,B,C):`
5. Export ke folder baru yang kosong
6. Export ke folder yang sudah berisi file dengan nama sama
7. Export file `.pdf`, `.dxf`, `.png` satu per satu

## 9. Pertanyaan yang harus dijawab setelah tiap percobaan

- Apakah error selalu muncul pada kondisi tertentu?
- Apakah file selalu dibuat tetapi error tetap muncul?
- Apakah error hanya terjadi pada page tertentu?
- Apakah error hanya terjadi pada format tertentu?
- Apakah error hanya terjadi dengan template tertentu?

## 10. Kesimpulan yang dicari

Tujuan akhirnya adalah menemukan pola yang konsisten: apakah `ExportEx` gagal karena:

- halaman target salah
- nama output invalid
- format export tidak cocok
- path tidak valid
- `cdrCurrentPage` tidak sesuai
- dokumen state tidak stabil
- atau kombinasi dari beberapa faktor di atas
