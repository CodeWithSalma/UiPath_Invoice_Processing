# RPA Invoice Processing — Sinkronisasi Invoice Vendor dari Email ke Sistem Tracking

Project UiPath untuk mengotomasi penerimaan, ekstraksi, validasi, dan pencatatan
invoice vendor yang masuk via email — relevan untuk tim Finance/AP yang menerima banyak invoice spare part, jasa
maintenance, dan instalasi dari berbagai vendor setiap hari.

## Konteks Masalah

Tim Accounts Payable (AP) menerima invoice dari vendor melalui email dengan
format yang **tidak seragam**:

- **PDF** — invoice resmi (hasil cetak digital atau hasil scan/foto)
- **Excel** — beberapa vendor mengirim rekap batch invoice dalam satu file
- **Gambar (JPG/PNG)** — vendor kecil/lapangan kadang memfoto invoice langsung
- **Email body** — vendor tertentu menulis ringkasan invoice langsung di badan email

Proses manual saat ini: staff AP membuka tiap email, mengunduh lampiran,
membaca/mengetik ulang data ke spreadsheet tracking, lalu memindahkan email
ke folder arsip. Ini memakan waktu, rawan salah ketik nominal, dan sulit
dilacak invoice mana yang sudah/belum diproses.

## Solusi

Bot UiPath dengan arsitektur berikut (lihat diagram alur di atas):

1. **Pantau inbox vendor** — bot mengecek folder email `Inbox/Vendor Invoice`
   secara terjadwal (mis. tiap 30 menit via Orchestrator trigger)
2. **Unduh lampiran** — simpan semua attachment ke folder kerja lokal,
   diberi nama terstruktur (`{vendor}_{tanggal}_{nomor_invoice}`)
3. **Klasifikasi format file** — berdasarkan ekstensi: PDF/gambar vs Excel
4. **Ekstraksi data**:
   - PDF/Gambar → **UiPath Document Understanding** (atau OCR dasar) untuk
     menarik field: nama vendor, no. invoice, tanggal, jatuh tempo, item,
     subtotal, pajak, total tagihan
   - Excel → baca langsung pakai **Excel Application Scope / Read Range**,
     karena datanya sudah tabular (lebih cepat & akurat daripada OCR)
5. **Validasi & bersihkan data** — field kosong, spasi berlebih, format
   tanggal/angka tidak konsisten (mis. `"Rp 3.200.000"` vs `"3200000"`),
   nomor invoice tidak konsisten kapitalisasi (`INV-AC-0501` vs `inv-ac-0502`).
   field yang gagal validasi diberi status **"Need Review"**, bukan diteruskan
   begitu saja.
6. **Update master tracking & arsipkan**:
   - Tulis/append ke `invoice_tracking_master.xlsx`
   - Pindahkan email ke folder `Processed` (atau `Need Review` jika confidence
     ekstraksi rendah/data tidak lengkap)
   - Kirim ringkasan harian ke tim AP (jumlah invoice diproses, jumlah perlu
     review, total nominal)

## Hasil yang Diharapkan

- Waktu input invoice per dokumen turun dari ±5-10 menit manual menjadi
  diproses otomatis dalam hitungan detik per file
- Semua invoice (apapun formatnya) tercatat di satu master tracking yang
  konsisten, siap untuk rekonsiliasi/approval pembayaran
- Invoice dengan data tidak lengkap/ambigu otomatis ditandai untuk
  ditinjau manual — bukan diproses "asal jalan" yang berisiko salah bayar
- Jejak audit jelas: email mana sudah diproses, kapan, dan data apa yang
  perlu dicek ulang

## Data Uji (Test Data)

| File | Deskripsi |
|---|---|
| `TestData/sample_invoice_PT_Sumber_Sejuk_Abadi.pdf` | Contoh invoice PDF resmi (1 vendor, 4 item barang) untuk uji ekstraksi OCR/Document Understanding |
| `TestData/sample_invoice_batch_vendor.xlsx` | Contoh rekap batch invoice dari satu vendor jasa, sengaja dibuat dengan data berantakan (spasi, format tanggal campur, harga dengan prefix "Rp", nomor invoice kosong/tidak konsisten) untuk uji logika cleaning |

## Konfigurasi

`Config/Config.xlsx` berisi semua pengaturan bot (folder email, path lokal,
ekstensi file yang diizinkan, ambang batas confidence OCR, penerima notifikasi)
— mengikuti best practice **REFramework**, supaya tidak ada nilai hardcode
di dalam workflow.

## Struktur Proyek (Rencana Build di UiPath Studio)

```
InvoiceProcessing/
├── Main.xaml                      # Entry point (REFramework)
├── Framework/                     # Komponen standar REFramework
│   ├── InitAllSettings.xaml
│   ├── GetTransactionData.xaml    # Ambil 1 email/attachment per transaksi
│   ├── Process.xaml               # Logika inti: klasifikasi -> ekstrak -> validasi -> simpan
│   └── SetTransactionStatus.xaml
├── Config/
│   └── Config.xlsx
├── TestData/
│   ├── sample_invoice_PT_Sumber_Sejuk_Abadi.pdf
│   └── sample_invoice_batch_vendor.xlsx
└── Output/
    ├── invoice_tracking_master.xlsx
    └── data_quality_log.csv
```

## Panduan Build (mapping ke UiPath Activities)

Lihat `BUILD_GUIDE.md` untuk daftar activity per langkah, lengkap dengan
package UiPath yang perlu di-install.

## Pengembangan Lanjutan (Ide)

- Tambahkan **Action Center (Human-in-the-loop)** untuk invoice status
  "Need Review" — staff AP approve/reject langsung dari UiPath Action Center
  alih-alih buka email manual
- Hubungkan output ke API ERP (SAP/Accurate) untuk auto-posting invoice,
  bukan hanya Excel
- Tambahkan deteksi duplikat invoice (nomor invoice yang sama masuk 2x)
- Dashboard ringkas (Power BI/Excel) untuk memantau status invoice harian

tessttt