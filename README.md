# ODS → XLSX Dönüştürücü

OpenDocument tablolarını (`.ods`, `.fods`) Excel'in `.xlsx` formatına dönüştüren tek sayfalık bir web aracı.

**Canlı:** https://keremcankabadayi.github.io/ods-to-xlsx/

- Dönüşüm tamamen tarayıcıda yapılır, dosyalar hiçbir sunucuya yüklenmez.
- Sürükle-bırak, birden fazla dosyayı aynı anda dönüştürme.
- Build adımı yok; düz HTML, CSS ve JS. [SheetJS](https://sheetjs.com) `vendor/` klasörüne gömülü (0.20.3).

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

MIT. SheetJS, Apache 2.0 lisansıyla dağıtılır (`vendor/xlsx-LICENSE`).
