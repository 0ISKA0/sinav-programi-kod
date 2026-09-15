# -*- coding: utf-8 -*-
"""
DÖNEM AYARLARI — ÖRNEK. Yeni dönemde bu dosyayı 'donem_ayarlari.py' adıyla kopyalayıp doldurun.
'donem_ayarlari.py' GitHub'a yüklenmez (şifreler ve kişi bilgileri içerir).
Aşağıdaki değerler örnektir; boş bırakılabilecek istisnalar boş küme / boş liste olarak duruyor.
"""

# ---------------------------------------------------------------- dönem ve dosyalar
DONEM_ETIKETI = "2026-2027 Güz İYBF Vize"                 # çıktı dosya adlarının başı
SABLON_DOSYASI = "2026-2027 Güz İYBF Vize - Şablon.xlsx"  # proje kökündeki çalışma şablonu ('Yöntem/Yeni Dönem Şablonu.xlsx' kopyası)
TALEP_DOSYASI = "Sınav Programı Talep Dosyası.xlsx"        # Belgeler/ içinde
SISTEM_DOSYASI = "Ders öğrenci sayıları.xlsx"              # Belgeler/ içinde (OBS sayıları)
BOS_DERSLIK_DOSYASI = "Boş Derslik Dosyası.xlsx"           # Belgeler/ içinde (üniversite formu)
SITE_DOSYASI = "2026-27 Güz Vize Sınav Programı.html"      # Çıktılar/ içinde

# ---------------------------------------------------------------- takvim
SINAV_TARIHLERI = {                                        # gün -> tarih (gg.aa.yyyy)
    "Pazartesi": "09.11.2026", "Salı": "10.11.2026", "Çarşamba": "11.11.2026",
    "Perşembe": "12.11.2026", "Cuma": "13.11.2026",
}
AY_ADI = "Kasım"                                           # talepteki "11 Kasım" gibi tarihleri okumak için
DONEM_SURESI_METNI = "9–13 Kasım 2026"
HAVUZ_YARIYIL = {"HAVUZ2": 5, "HAVUZ4": 7}                 # güz: 5/7, bahar: 6/8

# ---------------------------------------------------------------- yayın ve site başlıkları, şifreler
YAYIN_BASLIGI = ("2026-2027 Eğitim Öğretim Yılı Güz Dönemi Vize Sınav Programı / "
                 "2026-2027 Fall Semester Midterm Exam Schedule")
SITE_BASLIGI = "2026-2027 Güz Dönemi Vize Sınav Programı"
SITE_SIFRESI = "BURAYA-SITE-SIFRESI"
BIRIM_SIFRESI = "BURAYA-BIRIM-SIFRESI"

# ---------------------------------------------------------------- ders istisnaları (dönem içinde kullanıcı kararlarıyla doldurulur)
LAB_TASMA_ONAYLI = set()          # LAB kapasitesini aşsa da tek LAB'da yapılacak sınavlar (ad anahtarı: küçük harf, boşluksuz)
TUM_BOLUMLER_ALIR = {}            # {"dersadianahtari": yarıyıl}: bütün bölümler alıyor kabul edilir
SISTEM_AD_ESI = {}                # {"talepteki ad anahtarı": "sistemdeki ad anahtarı"}: aynı ders, farklı ad
SURE_ISTISNA = {}                 # {"Ders Adı": 90}: sonradan bildirilen sınav süreleri (dk)
GOZETMEN_ONAYLI_HOCA = set()      # {"SOYAD"}: 30 altı olsa da hocanın istediği gözetmen verilir
BIRLIKTE_SINAVLAR = []            # [("dersanahtari1", "dersanahtari2")]: aynı saatte aynı derslikte birlikte
TEK_DERSLIK_TERCIH = set()        # {"Ders Adı"}: mümkünse tek derslikte
HOCA_EN_GEC_BASLANGIC = {}        # {"SOYAD": 15}: hocanın sınavı en geç bu saatte başlar
SERT_TALEPLI_DERSLIK = ("Kütüphane Lab", "7. kat")  # bu derslikleri isteyen hocanın gün/saati kesin

# ---------------------------------------------------------------- gözetmen istisnaları
GOZETMEN_ARDISIK_MUAF = set()     # {"Ad Soyad"}: üst üste 3 sınav kuralından muaf
GOZETMEN_AZ_GUN = set()           # {"Ad Soyad"}: görevleri az güne toplansın
GOZETMEN_COK_GUN = set()          # {"Ad Soyad"}: görevleri çok güne yayılsın

# ---------------------------------------------------------------- raporda ayrıca gösterilecek sınavlar
RAPOR_OZEL_ISTEKLER = []
