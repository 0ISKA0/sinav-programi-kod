# Sınav ve Gözetmenlik Programı — Rehber (v2)

İstanbul Esenyurt Üniversitesi İşletme ve Yönetim Bilimleri Fakültesi (İYBF) için geliştirildi.
Bu rehber, 2025-2026 Bahar vize denemesinde (v2) oturan yöntemi anlatır. Claude ile birlikte ya da
tek başına kullanılabilir.

---

## 1. Klasör yapısı

```
<Dönem klasörü>/
├── <Dönem> - Şablon.xlsx          ← çalışma dosyası (Yöntem/Yeni Dönem Şablonu.xlsx kopyası)
├── Belgeler/                      ← GİRDİLER (GitHub'a yüklenmez)
│   ├── Sınav Programı Talep Dosyası.xlsx   (ders programı + hoca talepleri, tek dosya)
│   ├── Ders öğrenci sayıları.xlsx          (OBS'den, aynı dönem)
│   └── Boş Derslik Dosyası.xlsx            (üniversitenin kat kat derslik formu)
├── Çıktılar/                      ← ÇIKTILAR (GitHub'a yüklenmez)
└── Yöntem/
    ├── REHBER.md                  ← bu dosya
    ├── Yeni Dönem Şablonu.xlsx    ← boş şablon (Genel Kısıtlar dolu)
    ├── site_sablonu.html          ← verisiz, şifresiz web sitesi şablonu
    ├── Kod Anahtari.txt           ← bölüm kodları, derslikler
    └── kod/
        ├── donem_ayarlari_ornek.py   ← örnek dönem ayarları (GitHub'da)
        ├── donem_ayarlari.py         ← GERÇEK dönem ayarları + şifreler (GitHub'a YÜKLENMEZ)
        ├── 1_dersleri_yukle.py
        ├── 2_programi_kur.py
        ├── 3_formlari_uret.py
        └── 4_senaryo.py
```

**Kural:** kod dosyaları her dönem aynıdır. Döneme ait her şey (tarihler, dosya adları, şifreler,
kişiye özel istisnalar) yalnızca `donem_ayarlari.py` içinde ve şablonun **Özel Kısıtlar** sekmesinde durur.

---

## 2. Yeni dönem — adım adım

| # | Ne yapılır | Kim |
|---|---|---|
| 1 | `Yöntem/Yeni Dönem Şablonu.xlsx` → dönem klasörüne kopyala, adını ver | Kullanıcı |
| 2 | `kod/donem_ayarlari_ornek.py` → `kod/donem_ayarlari.py` kopyala; tarihleri, dosya adlarını, şifreleri yaz | Kullanıcı / Claude |
| 3 | Belgeleri `Belgeler/` klasörüne koy | Kullanıcı |
| 4 | Şablonda **Gözetmenler** (ad, cinsiyet, kısıt, izin günü) ve **Derslik Kısıtlamaları** (Müsait/Dolu) doldur | Kullanıcı |
| 5 | `python 1_dersleri_yukle.py --yaz` → **Dersler** sekmesi; sarı satırlardaki soruları kullanıcıyla netleştir | Claude + kullanıcı |
| 6 | Kararları `donem_ayarlari.py` ve **Özel Kısıtlar** sekmesine yaz, 5. adımı tekrarla | Claude |
| 7 | `python 2_programi_kur.py --deneme 200 --yaz` → program + gözetmenler + kontrol raporu | Claude |
| 8 | `python 3_formlari_uret.py --pdf` → yayın formu (Excel+PDF), derslik dosyası, gözetmenlik, web sitesi | Claude |
| 9 | Sonradan gelen değişiklikler: `python 4_senaryo.py "açıklama" [--simdi "Gün SS:00"] [--sabit "ad;ad"] --yaz` | Claude |

> **Programı kurmadan önce** derslik müsaitliği mutlaka sorulur (başka fakülteler de aynı derslikleri kullanır).

---

## 3. Her dönem kullanıcıya sorulacaklar

- Sınav türü, tarihleri, tatil günleri
- Derslik müsaitliği (hangi derslik hangi gün/saat boş)
- Gözetmen listesi, cinsiyet, izin günleri, kişisel saat kısıtları, muafiyetler
- Hoca talebinde belirsiz olanlar (Dersler sekmesinde sarı satırlar):
  - hoca gözetmen istedi ama öğrenci 30'un altında
  - "dersi alan bölümler"de programda olmayan bölüm yazılmış
  - talep ile sistem öğrenci sayısı çok farklı
  - hoca adı farklı yazılmış / tanımsız ders
- Birlikte yapılacak sınavlar, özel derslik istekleri (lab, belirli kat)

---

## 4. Kurallar

Kuralların tamamı şablonun **Genel Kısıtlar** sekmesindedir (Sert / Tercih / Yöntem olarak işaretli).
Özet:

- **Çakışma:** aynı program-yarıyıl grubunun iki sınavı aynı anda olmaz; aynı hocanın iki sınavı aynı anda olmaz.
  Çakışma grubu = programdaki bölüm-yarıyıl + hocanın "dersi alan bölümler" bilgisi + sistemde dersi alan
  görünen her bölüm (şüphede ekle).
- **Arka arkaya:** bir grubun kendi bölüm dersleri aynı gün en fazla 2 sınav arka arkaya.
- **Süre:** 50 dk; hocanın 60 dk'ya kadar yazdığı süre 50 sayılır; 90 dk derslik 2 oturum kapatır.
- **Derslik:** önce 6. kat; 60'tan fazla öğrenci 3. kata (hoca talebinden önce gelir); bölünen sınavda
  derslikler aynı katta; her bölüm mümkünse tek derslikte; kapasiteler ±6 kişi esnek; az derslik tercih.
- **Hoca:** akademik izin gününe sınav konmaz; izin günü bilinmeyen hocanın sınavı dersin normal saatine;
  hoca talebi mümkün olduğunca karşılanır.
- **Gözetmen ihtiyacı:** 30 ve üzeri → gözetmen (derslik başına 1, 3. katta 2); altı → dersin hocası girer.
- **Gözetmen atama:** izin günü ve saat kısıtları; aynı saatte tek görev; üst üste 3 görev yok;
  toplam ve İngilizce yük eşit; 3. kat çift gözetmende bir kadın bir erkek.
- **Sonradan değişiklik:** program baştan kurulmaz; en az değişiklik; başlamış / değiştirilemez sınav oynatılmaz.

---

## 5. Çıktılar

| Dosya | Açıklama |
|---|---|
| `… Sınav Programı - Yayın.xlsx / .pdf` | Okul sitesi için; talep dosyasının biçimi, her bölüm satırında kendi dersliği, gerçek bitiş saati |
| `… - Derslik Dosyası (6. Kat).xlsx` | Üniversite derslik formu; hücrede hoca / ders (bölünen sınavda bölüm etiketi) |
| `… - Gözetmenlik.xlsx` | Her gözetmen için haftalık blok + özet |
| `….html` | Web sitesi: bölüm / öğretim üyesi / gözetmen sekmeleri, şifreli Birim sekmesi |
| Şablon sekmeleri | Sınav Programı, Dersler, Derslik Kısıtlamaları, Gözetmenler (yük), Genel/Özel Kısıtlar |

> Sitedeki şifreler yalnızca göz önünden saklar; sayfa kaynağında veri görülebilir. Gerçekten gizli bilgi
> siteye konmaz.

---

## 6. Bilinen tuzaklar

- Talep dosyasının sütun düzeni dönemden döneme değişebilir (güz ve bahar farklıydı) — okuma kodu kontrol edilmeli.
- Fakülte Seçmeli Havuzu blokları bir sütun sola kayıktır.
- Hoca taleplerinde tarih/saat serbest metindir ("Perşembe 12-15", "8 Nisan 13:00", "sabah"); belirsizler sorulur.
- Aynı hoca farklı yazılır (unvan, büyük/küçük harf, bitişik yazım) — isimler tekilleştirilir.
- Aynı öğrenci sayısı birden çok satırda tekrar edebilir (toplam mı, bölüm mü?) — sorulur.
- Excel dosyası açıkken yazılamaz.
- HTML dosyası telefonda önizleme uygulamasında açılırsa JavaScript çalışmaz; tarayıcıda açılmalı.
