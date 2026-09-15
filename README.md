# İYBF Sınav ve Gözetmenlik Programı

İstanbul Esenyurt Üniversitesi İşletme ve Yönetim Bilimleri Fakültesi sınav programı ve gözetmen
dağılımını hazırlayan araçlar.

- Talep dosyasını okur, dersleri benzersiz sınavlara indirir, çakışma gruplarını kurar
- Sınavları derslik müsaitliği, hoca talepleri ve kurallara göre yerleştirir; gözetmen atar
- Bağımsız kontrol yapar; yayın formu (Excel/PDF), üniversite derslik dosyası, gözetmenlik çizelgesi ve
  web sitesi üretir
- Yayınlanmış programa sonradan gelen değişiklikleri en az değişiklikle uygular

Kullanım ve kurallar: [Yöntem/REHBER.md](Yöntem/REHBER.md)

**Gizlilik:** `Belgeler/`, `Çıktılar/`, çalışma şablonları, üretilen veri dosyaları ve
`Yöntem/kod/donem_ayarlari.py` (şifreler, kişi bilgileri) bu depoya yüklenmez — bkz. `.gitignore`.

Gereksinim: Python 3, `openpyxl` (PDF için Windows'ta Microsoft Excel).
