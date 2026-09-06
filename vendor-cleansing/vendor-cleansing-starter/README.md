# Vendor Cleansing Otomatis

Sistem ini mengubah sembilan file vendor dan PO menjadi satu workbook Excel. Daftar utama adalah seluruh record HK Circle dari tujuh sumber vendor/calon. Record duplikat, tanpa SAP, tanpa nama, dan tanpa PO dipertahankan. SAP PO yang belum terwakili juga ditambahkan agar pekerjaan dan nilainya tidak hilang. Status calon tidak menghalangi klasifikasi berbukti PO.

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

1. `Data Cleansing`: satu baris per record HK Circle, ditambah SAP PO yang belum terwakili. Atribut asli tidak diganti oleh record lain. Jumlah baris bukan jumlah vendor unik.
2. `Audit`: ringkasan proses, asumsi aktif, jumlah temuan berdasarkan severity, dan rincian anomali yang dapat difilter.

## Identity resolution

Urutan pencocokan otomatis:

1. `NO SAP` / `Ext Number` ke `PO.Vendor`.
2. `ID Vendor` / `ID` / `Kode Identitas` melalui relasi ID-SAP pada DRT/DM.
3. Nama exact yang hanya mengarah ke satu SAP PO.
4. Nama kanonik yang hanya mengarah ke satu SAP PO.

Nama ambigu tidak digabung otomatis dan masuk ke sheet `Audit`. Kecocokan nama unik hanya menyediakan bukti klasifikasi sementara yang perlu konfirmasi. SAP dan saldo tidak diisi dari nama saja. Nama sama dengan SAP berbeda ditandai HIGH, tanpa menggabungkan identitas atau saldo.

Penanda `not_posted` diperlakukan sebagai SAP belum tersedia, bukan nomor SAP bersama. Nilai aslinya tetap ada di kolom SAP Sumber pada tabel penelusuran Audit. File sumber tidak diubah.

## Aturan kategori

Aturan dapat diubah di `config/revamp_settings.json`:

| Kategori | Kondisi | Perlakuan |
|---|---|---|
| A | Ada pada master aktif DRT atau DM | DRP |
| B | Hanya ada pada DRT Lama atau DM Lama | DRP + Daftar ulang |
| C | Tidak ada di master, tetapi ada pada DCR atau DCM | DRP + Prioritas Approve |
| D | Tidak ada di master/DCR/DCM, tetapi ada pada DBCR | DRP + Update data |
| E | Hanya ditemukan pada PO | DRP + Daftar ulang |

`DRT` pada output mewakili master aktif DRT/DM. `DRT Lama` mewakili master lama jika tidak ditemukan pada master aktif. `Inject` adalah status tindak lanjut untuk vendor yang sudah mempunyai PO tetapi belum ditemukan pada master aktif maupun lama, sehingga perlu registrasi atau daftar ulang. Oranye bukan severity anomali.

## Klasifikasi

`Klasifikasi Final` dihasilkan dari deskripsi item PO menggunakan:

- `config/vocabulary_v2.csv`
- `config/po_rules_v2.csv`
- `config/classification_exclusions_v2.csv`

HK Circle dipetakan melalui `config/context_rules_v2.csv` sebagai validasi. Circle tidak
menambahkan klasifikasi yang tidak mempunyai bukti PO. Jika Circle kosong, klasifikasi
tetap dapat dihasilkan dari PO. Jika Circle dan PO tidak beririsan, hasil PO dipertahankan
dan perbedaannya dicatat pada Audit agar tidak terjadi penggabungan klasifikasi yang tidak
berkaitan.

Rule sengaja bersifat konservatif. Deskripsi umum seperti `material bantu`, `upah`, atau
`jasa temporary` tidak ditebak otomatis tanpa frasa pekerjaan yang lebih spesifik.

Tiga kolom Kelompok Klasifikasi Level 1–3 dibentuk langsung dari setiap deskripsi item PO menggunakan `config/classification_hierarchy.json`, dari `Klasifikasi_Rekanan_HK_Group copy.xlsx`. Level 2 menampilkan nama tanpa kode; Level 3 hanya cakupan yang terbukti, bukan seluruh cakupan kelompok. Satu baris teks per kelompok Level 2, sejajar dengan Level 1 dan Level 3. Beberapa cakupan dipisahkan titik koma. Kode master tetap ada di bukti Audit. Item ambigu tidak ditebak.

## Saldo Hutang

Definisi operasional sesuai konfirmasi pengguna: jumlah kolom `Nilai PO` setiap baris pekerjaan PO HK dan PO JO, dikelompokkan berdasarkan `Vendor` (nomor SAP). Ini bukan perhitungan saldo belum dibayar. Semua angka diperlakukan sebagai IDR tanpa konversi; perbedaan label currency sumber dicatat pada Audit.

Setiap baris PO tetap dihitung, termasuk nomor PO/item berulang. Pengulangan kunci dicatat untuk pemeriksaan, tidak dihapus otomatis. Baris total/footer bukan pekerjaan dan tidak dihitung lagi.

Saldo dicatat sekali per SAP pada record pertama menurut prioritas sumber. Record lain dengan SAP sama tetap ada, dengan saldo kosong dan rujukan ke baris pemilik saldo dalam Audit. Ini adalah aturan pencatatan agar total tidak berlipat, bukan penggabungan identitas. SAP berbeda selalu memiliki akumulasi terpisah. Tanpa PO, saldo kosong, bukan dianggap nol.

## Quality guard

Ekspor memakai penulisan Excel bertahap dan penyimpanan teks PO bersama untuk mengurangi pemakaian memori. Tidak memerlukan Node.js, npm, atau paket privat. Struktur Excel tetap dua sheet dengan tabel dan filter native.

Run dihentikan apabila:

- ada input wajib yang hilang;
- ada record HK Circle atau SAP PO yang tidak terwakili;
- ada baris PO valid yang tidak terhitung;
- Level 1–3 terisi tidak lengkap atau jumlah kelompoknya tidak sejajar;
- Nilai PO kosong/tidak valid atau jumlah saldo output tidak sama dengan sumber;
- struktur kolom sumber berubah dan kolom wajib tidak ditemukan.

## Audit dan highlight

Temuan audit dibagi menjadi `HIGH`, `MEDIUM`, dan `LOW`. Satu record dapat memiliki beberapa temuan; jumlah temuan bukan jumlah vendor unik. Audit menyediakan navigasi klik, ringkasan jenis temuan, tabel duplikasi/konflik, kelengkapan, pencocokan, klasifikasi, akumulasi Nilai PO per SAP, dan penelusuran seluruh record. Rujukan baris Data Cleansing disediakan, termasuk record tanpa SAP. Bagian Rekonsiliasi Baris Total PO tidak ditampilkan.

Hyperlink hanya digunakan pada navigasi antartabel Audit. Nomor baris rinci ditampilkan sebagai teks agar jumlah hyperlink tidak melewati batas Excel dan tidak memicu pesan `Removed Feature: Hyperlinks`.

Baris PO yang tidak mempunyai Vendor/SAP tidak dapat menjadi baris Data Cleansing, tetapi
tetap dicatat sebagai temuan `HIGH` pada Audit lengkap dengan perusahaan, nomor PO, item,
baris sumber, nama vendor, dan deskripsinya.

Baris total/footer PO dikenali dari kosongnya identitas item serta adanya mata uang dan nilai total. Footer direkonsiliasi terhadap penjumlahan seluruh item valid per mata uang dan ditampilkan pada tabel tersendiri di Audit. Footer bukan anomali Vendor/SAP dan nilainya tidak dibagikan ke `Saldo Hutang` vendor.

Temuan tersebut dipisahkan menjadi tabel duplikasi/konflik, kelengkapan data, pencocokan sumber terhadap PO, dan rekonsiliasi klasifikasi agar tindak lanjutnya tidak tercampur. Sheet `Audit` juga memuat tabel bukti klasifikasi per vendor dan rule, serta kelompok
deskripsi item PO yang belum terklasifikasi. Dengan demikian setiap Klasifikasi Final dapat
ditelusuri ke jumlah PO, jumlah item, contoh deskripsi, rule, dan dukungan Circle.

Audit juga memuat bukti hierarchy Level 1–3 per vendor, kode Level 2, jumlah PO/item, nomor baris PO sumber, istilah yang cocok, contoh deskripsi, dan Boundary Rule yang diterapkan. Item yang belum mempunyai bukti atau masih ambigu tersedia pada tabel terpisah untuk penyempurnaan rule berikutnya.

Sheet `Data Cleansing` memberi highlight otomatis:

- merah: anomali severity tinggi atau duplikasi;
- kuning: data wajib belum lengkap atau temuan menengah;
- ungu: `Klasifikasi Final` belum terdeteksi;
- oranye: status `Inject`, yaitu vendor sudah mempunyai PO tetapi belum berada di master aktif/lama dan perlu registrasi atau daftar ulang; warna ini bukan severity.

Seluruh temuan tetap disimpan pada sheet `Audit`; highlight hanya membantu pemindaian cepat dan tidak menghapus data.

Jalankan tes:

```powershell
python -m unittest discover -s tests -v
```
