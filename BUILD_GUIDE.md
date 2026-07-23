# Build Guide — Mapping Proses ke UiPath Activities

Panduan ini membantu kamu membangun ulang workflow di UiPath Studio
secara konsisten dengan desain di README.md. Disusun mengikuti pola
**REFramework** (Get Transaction → Process → Set Status), template
bawaan UiPath Studio (`File > New > Robotic Enterprise Framework`).

## Package yang Perlu Di-install

Buka **Manage Packages** di Studio, install:

| Package | Kegunaan |
|---|---|
| `UiPath.Mail.Activities` | Baca/pindah email, unduh attachment (IMAP/Outlook/Exchange) |
| `UiPath.Excel.Activities` | Baca/tulis file Excel (master tracking, batch invoice vendor) |
| `UiPath.PDF.Activities` | Ekstrak teks dari PDF (untuk PDF berbasis teks, bukan scan) |
| `UiPath.DocumentUnderstanding.Activities` | OCR + ekstraksi data terstruktur dari PDF scan/gambar invoice |
| `UiPath.System.Activities` | Activities umum (Assign, If, For Each, Log Message, dll — biasanya sudah default) |
| `UiPath.GSuite.Activities` | Google Workspace activities package |

## 1. Pantau Inbox Vendor

**Activities:** `Get Outlook Mail Messages` (atau `Get IMAP Mail Messages`)
- Folder: ambil dari `Config("EmailFolderInbox")`
- Filter: `UnreadOnly = True`, atau filter subject mengandung "Invoice"/"Faktur"
- Top: sesuaikan jumlah email yang diambil per run (mis. 20)

Di pola REFramework, ini ditaruh di `GetTransactionData.xaml` — satu
transaksi = satu email (atau satu attachment, tergantung granularitas
yang kamu mau dilog).

## 2. Unduh Lampiran

**Activities:** `Save Attachments`
- Output Path: `Config("AttachmentDownloadPath") + "\" + tanggal_hari_ini`
- Setelah unduh, gunakan `Assign` untuk membangun nama file terstruktur
  (`vendor_tanggal_nomorinvoice.ext`) supaya mudah ditelusuri.

## 3. Klasifikasi Format File

**Activities:** `Assign` + `Switch` (berdasarkan ekstensi file)
```
fileExtension = Path.GetExtension(attachmentPath).ToLower
```
- `.pdf` → cek dulu apakah PDF berbasis teks atau scan (coba `Read PDF Text`,
  kalau hasil kosong/sangat pendek → treat sebagai scan, lempar ke OCR)
- `.jpg` / `.jpeg` / `.png` → langsung ke alur OCR/Document Understanding
- `.xlsx` / `.xls` → langsung ke alur Excel

## 4a. Ekstraksi OCR / Document Understanding (PDF scan & gambar)

**Activities:** `Digitize Document` → `Classify Document` → `Extract Document Data`
(dari package Document Understanding, terhubung ke project DU di Orchestrator)

Field yang diekstrak (sesuaikan taxonomy di DU project):
`VendorName`, `InvoiceNumber`, `InvoiceDate`, `DueDate`, `LineItems`
(deskripsi, qty, harga satuan, total), `Subtotal`, `Tax`, `TotalAmount`

Simpan juga **confidence score** tiap field — dipakai di langkah validasi
untuk menentukan apakah invoice butuh review manual.

> Alternatif tanpa lisensi Document Understanding: gunakan `OCR` activity
> dasar (Google OCR/Microsoft OCR) + `Regex` untuk menarik pola seperti
> nomor invoice dan total, lalu cross-check manual ambang akurasinya lebih
> rendah dibanding Document Understanding.

## 4b. Baca Data Excel (langsung, tanpa OCR)

**Activities:** `Excel Application Scope` → `Read Range` → `For Each Row`

Karena ini data tabular langsung dari vendor, ekstraksi jauh lebih cepat
dan akurat dibanding OCR. Tiap baris di file Excel vendor = satu invoice
yang diproses dalam satu iterasi `For Each Row`.

## 5. Validasi & Bersihkan Data

**Activities:** `Assign` (logika cleaning) dibungkus dalam `Invoke Workflow File`
terpisah (mis. `CleanInvoiceData.xaml`) agar reusable untuk kedua alur (4a & 4b).

Logika yang perlu diimplementasikan (port dari logika Python project sebelumnya):
- `Trim` + `Regex.Replace(text, "\s+", " ")` untuk spasi berlebih
- Konversi nominal: hapus prefix "Rp"/spasi, tangani pemisah ribuan koma/titik,
  pakai `Decimal.TryParse` (kalau gagal parse → flag sebagai issue)
- Standarkan format tanggal dengan `DateTime.TryParseExact` mencoba beberapa
  format (`dd/MM/yyyy`, `yyyy-MM-dd`, `dd-MM-yyyy`, dst.)
- Standarkan nomor invoice ke huruf besar semua (`.ToUpper`)
- Field wajib kosong (vendor, no. invoice, total) → set `DataQuality = "Need Review"`

Gunakan `If`/`Switch` untuk routing: kalau ada issue → masuk daftar
"Need Review", kalau bersih → lanjut ke langkah 6 sebagai "OK".

## 6. Update Master Tracking & Arsipkan

**Activities:**
- `Append Range` (Excel Activities) → tulis baris baru ke
  `invoice_tracking_master.xlsx` (kolom: Vendor, NoInvoice, Tanggal,
  JatuhTempo, Total, Status, DataQuality, SourceFile)
- `Write CSV` / `Append Line` → tulis ke `data_quality_log.csv` kalau ada issue
- `Move Mail Message` → pindahkan email ke folder `Processed` atau
  `Need Review` sesuai hasil validasi
- (Opsional) `Send Outlook Mail Message` → kirim ringkasan harian ke
  `Config("NotificationEmailTo")`, dijadwalkan jalan sekali di akhir proses
  (bisa pakai workflow terpisah `SendDailySummary.xaml` yang dipanggil dari
  `Main.xaml` setelah semua transaksi selesai, pola umum di REFramework)

## Testing

1. Drop `TestData/sample_invoice_PT_Sumber_Sejuk_Abadi.pdf` dan
   `TestData/sample_invoice_batch_vendor.xlsx` ke folder simulasi inbox
   (atau attach ke email test di akun Outlook testing)
2. Jalankan bot dalam mode debug, cek tiap field hasil ekstraksi di
   **Locals panel**
3. Pastikan baris dengan data tidak lengkap (di `sample_invoice_batch_vendor.xlsx`
   ada beberapa baris sengaja kosong/format aneh) benar ditandai "Need Review",
   bukan ikut tersimpan sebagai data bersih
