# Batch Print Settings untuk Excel (VBA)

Repositori ini berisi contoh file Excel dan kode VBA untuk **mengatur pengaturan print (Page Setup) secara otomatis** pada banyak file Excel sekaligus.

Use case utama: ketika Anda memiliki puluhan/ ratusan file Excel dengan struktur sheet yang sama, dan ingin menerapkan pengaturan print yang seragam untuk setiap sheet di semua file tersebut.

---

## Fitur

- Memproses **semua file Excel** (`.xlsx`, `.xlsm`, `.xlsb`) dalam satu folder.
- Mengatur pengaturan print per-sheet dengan aturan:
  - **Sheet 1, 2, 3**: orientasi **Portrait**
  - **Sheet 4**: orientasi **Landscape**
  - Ukuran kertas: **A4**
  - Margin: **Normal**
    - Top: 1,91 cm
    - Bottom: 1,91 cm
    - Left: 1,76 cm
    - Right: 1,76 cm
  - Skala: **Fit to 1 page wide** (lebar 1 halaman, tinggi boleh lebih dari 1 halaman)
  - Center on page: **Horizontal**
  - **Print Area** otomatis:
    - Dari sel `A1`
    - Sampai sel terakhir yang memiliki **border (garis tepi)**  
      → sehingga area yang dicetak mengikuti blok tabel yang diberi garis tepi.
- Bekerja untuk banyak file dalam satu kali eksekusi.

---

## Prasyarat

- Sistem operasi: **Windows**
- Aplikasi: **Microsoft Excel** (mendukung macro/VBA)
- Macro di-enable (centang `Enable all macros` atau izinkan macro saat membuka file).

---

## Struktur Repositori

> Sesuaikan nama file di bawah ini dengan isi repo Anda yang sebenarnya.

Contoh:

```text
.
├── README.md
├── FORMAT EXCEL.xlsx          # Contoh file Excel dengan 4 sheet
└── BatchPrintSettings.xlsm    # Workbook berisi macro VBA BatchSetPrintSettings
