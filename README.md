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
- ODS'nin XML'i doğrudan okunur (`ods2xlsx.js`) ve xlsx [ExcelJS](https://github.com/exceljs/exceljs) ile yazılır.
  Formüller LibreOffice sözdiziminden (`[.A1];[Sayfa.B2]`) Excel sözdizimine (`A1,Sayfa!B2`) çevrilir.
- Build adımı yok; düz HTML, CSS ve JS. `vendor/` altında gömülü: ExcelJS 4.4.0, [hash-wasm](https://github.com/Daninet/hash-wasm) 4.13.0 (Argon2).

## Neler korunur

| Korunur | Korunmaz |
| --- | --- |
| Veriler, tüm sayfalar, formüller (Excel sözdizimine çevrilir) | Makrolar |
| Yazı tipleri, renkler, dolgular, kenarlıklar, hizalama, metin kaydırma, döndürme | Grafikler, şekiller, düğmeler |
| Sayı / tarih / saat biçimleri | Koşullu biçimlendirme |
| Birleştirilmiş hücreler, sütun genişlikleri, satır yükseklikleri, gizli satır/sütun, gruplar | EMF/WMF/SVM çizimler (PNG yedeği varsa o kullanılır) |
| Açılır listeler (veri doğrulama), köprüler, notlar | Sayfa koruma parolası (sayfalar parolasız korumalı açılır) |
| Resimler (PNG/JPEG/GIF), sekme renkleri, gizli sayfalar, dondurulmuş bölmeler | Üstbilgi / altbilgi |
| Tanımlı adlar, yazdırma alanları, sayfa yapısı (kağıt, yön, kenar boşluğu, sığdırma), otomatik filtre | |

Gerçek bir 13 sayfalık LibreOffice dosyasında, LibreOffice'in kendi "xlsx olarak kaydet" çıktısıyla hücre hücre
karşılaştırıldığında değerler, formüller (38 bin), dolgular, kenarlıklar, birleştirmeler, sütun/satır ölçüleri,
veri doğrulamaları ve yazdırma alanları birebir eşleşmektedir.

## Yerelde çalıştırma

```sh
python3 -m http.server 8000
# http://localhost:8000
```

## Lisans

MIT. ExcelJS MIT (`vendor/exceljs-LICENSE`), hash-wasm MIT (`vendor/hash-wasm-LICENSE`) lisanslıdır.
