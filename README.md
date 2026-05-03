# Kasir Online Google Apps Script + GitHub

Aplikasi kasir ini **tidak bisa dijalankan 100% langsung dari GitHub Pages** karena file `Index.html` memakai `google.script.run`, yaitu fitur khusus Google Apps Script.

Jadi:
- **Google Apps Script** = tempat aplikasi utama berjalan.
- **GitHub** = tempat menyimpan source code dan bisa membuat halaman pembuka/redirect.
- **GitHub Pages** = halaman pembuka yang mengarah ke URL Web App Apps Script.

## Isi project

- `Code.gs` → backend Google Apps Script.
- `Index.html` → tampilan aplikasi kasir.
- `appsscript.json` → konfigurasi Apps Script.
- `docs/index.html` → halaman GitHub Pages / landing page.
- `.github/workflows/deploy-pages.yml` → workflow deploy GitHub Pages.

## Cara deploy aplikasi utama ke Google Apps Script

1. Buka Google Spreadsheet.
2. Klik **Ekstensi → Apps Script**.
3. Isi file `Code.gs` dengan isi `Code.gs` dari project ini.
4. Buat / isi file `Index.html` dengan isi `Index.html`.
5. Klik **Save**.
6. Jalankan fungsi:

```text
setupSheets
```

7. Deploy:

```text
Deploy → Manage deployments → Edit → Version: New version → Deploy
```

8. Copy URL Web App yang berakhiran `/exec`.

## Cara deploy ke GitHub

### Cara cepat lewat browser GitHub

1. Buat repository baru di GitHub.
2. Upload semua file dalam folder ini ke repository.
3. Buka `docs/index.html`.
4. Ganti:

```javascript
const WEB_APP_URL = "GANTI_DENGAN_URL_WEB_APP_APPS_SCRIPT";
```

menjadi URL Web App Apps Script kamu, contoh:

```javascript
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxxxxxxxxxxxxxxxx/exec";
```

5. Commit perubahan.

### Aktifkan GitHub Pages

1. Buka repository GitHub.
2. Masuk **Settings → Pages**.
3. Pilih **GitHub Actions** sebagai source.
4. Push/commit ke branch `main`.
5. Tunggu workflow selesai.
6. Buka URL GitHub Pages.

## Login default

```text
Admin: admin / admin123
User : user / user123
RFID Login: 000768567
```

## Fitur

- Login username/password.
- Login RFID otomatis.
- RFID transaksi otomatis simpan + tampilkan struk.
- Responsive untuk HP, laptop, komputer.
- Barcode EAN-13.
- Supplier per barang.
- Pembelian stok dengan supplier.
- Diskon item.
- Diskon keseluruhan.
- PPN keseluruhan.
- Laporan HJ/HPP.
- Admin dan user melihat semua transaksi.
