# ODS → XLSX Dönüştürücü

OpenDocument tablolarını (`.ods`, `.fods`) Excel'in `.xlsx` formatına dönüştüren tek sayfalık bir web aracı.

**Canlı:** https://keremcankabadayi.github.io/ods-to-xlsx/

- Dönüşüm tamamen tarayıcıda yapılır, dosyalar hiçbir sunucuya yüklenmez.
- Sürükle-bırak, birden fazla dosyayı aynı anda dönüştürme.
- **Parola korumalı ODS desteği** (`odscrypt.js`):
  - LibreOffice 24.2+ varsayılanı: Argon2id + AES-256-GCM
  - ODF 1.2/1.3: PBKDF2 + AES-256-CBC

  Sadece makro kütüphanesi parolalı olan (verisi açık) dosyalar parola sorulmadan dönüştürülür; şifreli makrolar atlanır.
  Doğru girilen parola, aynı oturumdaki diğer şifreli dosyalarda otomatik denenir. Eski Blowfish şifrelemesi (OpenOffice 2.x) desteklenmez.
- Build adımı yok; düz HTML, CSS ve JS. `vendor/` altında gömülü: [SheetJS](https://sheetjs.com) 0.20.3, [hash-wasm](https://github.com/Daninet/hash-wasm) 4.13.0 (Argon2).

## Neler korunur

| Korunur | Korunmaz |
| --- | --- |
| Veriler, tüm sayfalar, formüller, birleştirilmiş hücreler, sütun genişlikleri, tarihler | Hücre renkleri, yazı tipleri, kenarlıklar, grafikler, resimler, makrolar |

## Yerelde çalıştırma

```sh
python3 -m http.server 8000
# http://localhost:8000
```

## Lisans

MIT. SheetJS Apache 2.0 (`vendor/xlsx-LICENSE`), hash-wasm MIT (`vendor/hash-wasm-LICENSE`) lisanslıdır.
