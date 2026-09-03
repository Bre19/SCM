# Vendor Cleansing Otomatis

Sistem ini mengubah sembilan file vendor dan PO yang formatnya tetap menjadi satu workbook Excel siap ditinjau. Universe output adalah seluruh `NO SAP` unik yang terdapat pada PO HK atau PO JO. Status calon tidak menghalangi klasifikasi apabila calon tersebut dapat dihubungkan secara aman ke vendor PO.

Setiap pembaruan periode berikutnya cukup mengganti isi sembilan file sumber dengan data terbaru, mempertahankan nama dan struktur kolomnya, lalu menjalankan satu perintah.

## Input wajib

Letakkan file berikut di `data/raw/` dengan nama persis:

1. `PO HK.xlsx`
2. `PO JO.xlsx`
3. `DBCR.xls`
4. `DCR.xls`
5. `DRT.xls`
6. `DRT Lama.xls`
7. `DCM.xls`
8. `DM.xls`
9. `DM Lama.xls`

File `.xls` boleh berupa HTML export maupun workbook Excel. Pipeline mendeteksi format berdasarkan isi file.

## Instalasi

Windows PowerShell:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Menjalankan

```powershell
python scripts/run_vendor_cleansing_revamp.py
```

Menentukan folder input/output secara eksplisit:

```powershell
python scripts/run_vendor_cleansing_revamp.py `
  --raw-dir data/raw `
  --output-dir output/revamp
```

Hasil akhir hanya satu file:

- `output/revamp/Data Cleansing - Otomatis.xlsx`

Workbook tersebut selalu berisi dua sheet:

1. `Data Cleansing`: satu baris per `NO SAP` vendor PO, lengkap dengan status sumber, kategori, perlakuan, klasifikasi, dan item pekerjaan PO.
2. `Audit`: ringkasan proses, asumsi aktif, jumlah temuan berdasarkan severity, dan rincian anomali yang dapat difilter.

## Identity resolution

Urutan pencocokan otomatis:

1. `NO SAP` / `Ext Number` ke `PO.Vendor`.
2. `ID Vendor` / `ID` / `Kode Identitas` melalui relasi ID-SAP pada DRT/DM.
3. Nama exact yang hanya mengarah ke satu SAP PO.
4. Nama kanonik yang hanya mengarah ke satu SAP PO.

Nama ambigu tidak digabung otomatis dan masuk ke sheet `Audit`.

## Aturan kategori

Aturan dapat diubah di `config/revamp_settings.json`:

| Kategori | Kondisi | Perlakuan |
|---|---|---|
| A | Ada pada master aktif DRT atau DM | DRP |
| B | Hanya ada pada DRT Lama atau DM Lama | DRP + Daftar ulang |
| C | Tidak ada di master, tetapi ada pada DCR atau DCM | DRP + Prioritas Approve |
| D | Tidak ada di master/DCR/DCM, tetapi ada pada DBCR | DRP + Update data |
| E | Hanya ditemukan pada PO | DRP + Daftar ulang |

`DRT` pada output mewakili master aktif DRT/DM. `DRT Lama` mewakili master lama jika tidak ditemukan pada master aktif. `Inject` menandai vendor PO yang tidak berada pada master aktif maupun master lama.

## Klasifikasi

`Klasifikasi Final` dihasilkan dari deskripsi item PO menggunakan:

- `config/vocabulary_v2.csv`
- `config/po_rules_v2.csv`
- `config/classification_exclusions_v2.csv`

Tiga kolom Kelompok Klasifikasi Level 1–3 sengaja kosong pada versi ini. `Saldo Hutang` juga kosong sampai sumber datanya tersedia.

## Quality guard

Run dihentikan apabila:

- ada input wajib yang hilang;
- ada `NO SAP` output yang kosong atau duplikat;
- jumlah output berbeda dari jumlah vendor PO unik;
- Level 1–3 atau Saldo Hutang terisi pada versi ini;
- struktur kolom sumber berubah dan kolom wajib tidak ditemukan.

## Audit dan highlight

Temuan audit dibagi menjadi `HIGH`, `MEDIUM`, dan `LOW`. Contohnya mencakup vendor PO yang tidak ditemukan pada registri, nama calon yang ambigu, konflik ID/SAP/NPWP, ID atau NPWP kosong, format NPWP tidak wajar, atribut master belum lengkap, variasi nama pada PO, serta klasifikasi yang belum terdeteksi.

Sheet `Data Cleansing` memberi highlight otomatis:

- merah: anomali severity tinggi atau duplikasi;
- kuning: data wajib belum lengkap atau temuan menengah;
- ungu: `Klasifikasi Final` belum terdeteksi;
- oranye: vendor masih `Inject` atau belum berada di master aktif/lama.

Seluruh temuan tetap disimpan pada sheet `Audit`; highlight hanya membantu pemindaian cepat dan tidak menghapus data.

Jalankan tes:

```powershell
python -m unittest discover -s tests -v
```
