# -*- coding: utf-8 -*-
"""
Adım 0 — Hocalara gönderilecek bilgi toplama formunu üretir (tek dosya; her hocanın dersleri önceden yazılı).

Girdi : Belgeler/Ders Programı.pdf       (ESAS KAYNAK: hangi dersler, hoca, gün/saat/derslik)
        Belgeler/altan ve üsten .xlsx (öğrenci no, ders kodu; yalnızca ders başına SAYI için kullanılır)
        Belgeler/Öğrenci sayıları.xlsx   (yalnızca PDF ile karşılaştırma ve kodsuz dersleri eşleştirme için)
Çıktı : Çıktılar/Hoca Formları/<Dönem> - Hoca Bilgi Formu (Toplu).xlsx  (SharePoint'e yüklenecek tek dosya)
        Çıktılar/Hoca Formları/_Kontrol Raporu.xlsx  (PDF ile OBS arasındaki farklar; komisyon teyit eder)
OBS eski olabileceği için forma yalnızca PDF'te yazan dersler girer.
Formlarda öğrenci numarası YOKTUR; yalnızca ders başına toplam öğrenci sayısı yazılır.

Kullanım: python 0_hoca_formu.py [--hoca "SOYAD"]
"""
import math, os, re, sys, collections
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Protection, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kaynaklar import KOK, BELIRSIZ, tr_upper, ders_listesi  # noqa: E402

CIKTI = os.path.join(KOK, "Çıktılar", "Hoca Formları")
DONEM = "2026-2027 Güz Vize"
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]

# form seçenekleri (okuyucu da aynı metinleri kullanır)
SINAV_VAR = ["Evet - yazılı sınav", "Hayır - ödev/proje ile ölçülüyor", "Hayır - sınav yapılmayacak"]
DERS_SAATINDE = ["Evet - ders günü ve saatinde yapılsın", "Hayır - fark etmez"]
DERS_SAATINDE_BASLIK = "Sınav dersin gün ve saatinde yapılsın mı?"
SURELER = ["50 dk", "90 dk"]
GOZETMEN = ["Gözetmen istiyorum", "Kendim gireceğim"]
IZIN = ["İzin günüm yok"] + GUNLER
IZIN_BASLIK = "Akademik izin günü"
LAB = ["Hayır", "Evet - laboratuvar gerekli"]
# Excel açılır listesi tek seçimlik. Dersi alan bölümleri sistemden (kayıt dosyası) hazır yazıyoruz; hoca yalnızca
# listede OLMAYAN bir bölüm varsa tek sütundan seçer.
EK_BASLIK = "Listede olmayan başka bölüm var mı?"
BOLUMLER = ["TÜM BÖLÜMLER (bütün fakülte alıyor)", "Siyaset Bilimi ve Uluslararası İlişkiler",
            "Siyaset Bilimi ve Uluslararası İlişkiler (İng.)", "İşletme", "İşletme (İng.)", "Ekonomi ve Finans (İng.)",
            "Uluslararası Ticaret ve Finansman", "Lojistik Yönetimi", "Havacılık Yönetimi",
            "Yönetim Bilişim Sistemleri", "Elektronik Ticaret ve Yönetimi", "Başka fakülte/bölüm",
            "Birden fazla bölüm eksik (komisyonla görüşülecek)"]
EK_IPUCU = ("Soldaki 'dersi alan bölümler' sistemden gelir. Listede OLMAYAN bir bölüm de dersi alıyorsa onu seçin; "
            "bütün fakülte alıyorsa TÜM BÖLÜMLER'i, birden fazla bölüm eksikse son seçeneği seçin. Eksik yoksa boş bırakın.")
GOZETMEN_ESIK = 30   # bu sayının altındaki derslere gözetmen atanmaz, sınava dersin hocası girer
GOZETMEN_BASLIK = f"Gözetmen ({GOZETMEN_ESIK}'dan az öğrencili derse atanmaz)"
LAB_BASLIK = "Laboratuvar gerekli mi?"
AYNI_ANDA_BASLIK = "Aynı anda yapılacak sınav (aynı sınıf, gün, saat)"
IPUCU = {  # hücreye tıklayınca görünen kısa uyarılar (Excel sınırı: 255 karakter)
    GOZETMEN_BASLIK: f"Öğrenci sayısı {GOZETMEN_ESIK}'un altındaki derslerde gözetmen atanmaz; sınava dersin hocası girer. "
                     "Seçiminizi listeden yapın, yazı yazmayın.",
    LAB_BASLIK: "Sınav bilgisayar laboratuvarında yapılacaksa 'Evet' seçin; gerekmiyorsa boş bırakın.",
    DERS_SAATINDE_BASLIK: "Sınavın dersin normal gün ve saatinde (gri sütunda yazan) yapılmasını istiyorsanız 'Evet' seçin.",
    IZIN_BASLIK: "Sınav haftasında akademik izinli olduğunuz günü seçin; o güne sınavınız konmaz. "
                 "Birden fazla dersiniz varsa her satıra aynı günü seçin.",
}


def liste_dv(secenek=None, aralik=None, ipucu=None):
    """Listeden seçmeye zorlar: uymayan metin reddedilir (stop), hücreye tıklayınca uyarı görünür."""
    f = aralik if aralik else '"' + ",".join(secenek) + '"'
    dv = DataValidation(type="list", formula1=f, allow_blank=True, showErrorMessage=True,
                        showInputMessage=True, errorStyle="stop")
    dv.errorTitle, dv.error = "Yazı yazmayın", "Lütfen kutudaki oka tıklayıp listeden seçiniz."
    dv.promptTitle, dv.prompt = "Seçiniz", ipucu or "Yazı yazmayın; oka tıklayıp listeden seçiniz."
    return dv


TR_ALFABE = "ABCÇDEFGĞHIİJKLMNOÖPQRSŞTUÜVWXYZ"


def tr_sira(s):
    """Türkçe alfabe sırası için sıralama anahtarı (büyük harfli ad beklenir)."""
    return [TR_ALFABE.find(c) if c in TR_ALFABE else -1 for c in s]


# ---------------------------------------------------------------- form
KALIN = Font(bold=True)
BASLIK_DOLGU = PatternFill("solid", fgColor="1F3864")
GIRIS_DOLGU = PatternFill("solid", fgColor="FFF2CC")
HAZIR_DOLGU = PatternFill("solid", fgColor="EDEDED")
ince = Side(style="thin", color="999999")
CERCEVE = Border(left=ince, right=ince, top=ince, bottom=ince)

# (başlık, genişlik, giriş mi?, seçenekler)
SUTUNLAR = [
    ("Ders Kodu", 16, False, None),
    ("Ders", 38, False, None),
    ("Program / Sınıf", 46, False, None),
    ("Kayıtlı Öğrenci", 12, False, None),
    ("Normal Ders Saati", 36, False, None),
    ("Sistemde dersi alan bölümler", 46, False, None),
    ("Sınav yapılacak mı?", 36, True, SINAV_VAR),
    (DERS_SAATINDE_BASLIK, 40, True, DERS_SAATINDE),
    ("Sınav süresi", 12, True, SURELER),
    (GOZETMEN_BASLIK, 26, True, GOZETMEN),
    (IZIN_BASLIK, 18, True, IZIN),
    (LAB_BASLIK, 28, True, LAB),
    (EK_BASLIK, 50, True, "BOLUM"),
    ("Ek öğrenci sayısı (tahmini)", 16, True, None),
    (AYNI_ANDA_BASLIK, 46, True, None),
]
SATIR_YUKSEKLIGI = 15  # pt; Excel Online satırı kendiliğinden büyütmediği için yükseklik açıkça yazılır
TOPLU_ILK = 4  # 'Dersler' sayfasında tablo başlığının satırı


def toplu_form_yaz(yol, tum):
    """Tek dosya: Yönerge + Dersler (hoca x ders, filtrelenebilir) + gizli Listeler."""
    wb = openpyxl.Workbook()
    yon = wb.active
    yon.title = "Yönerge"
    satirlar = [
        f"{DONEM} Sınav Programı — Hoca Bilgi Formu",
        "",
        "1) 'Dersler' sayfasında 'Öğretim Elemanı' sütunundan kendi adınızı bulun (satırlar hoca bazında gruplu ve renkli).",
        "   Filtre kullanacaksanız mutlaka önce Görünüm sekmesinden 'Sayfa görünümleri' ile kendi görünümünüzü açın;",
        "   aksi halde filtreniz dosyayı aynı anda açık olan diğer hocalarda da uygulanır.",
        "2) Sarı sütunları açılır listelerden doldurun. Gri sütunlar bizden hazır gelir, değiştirilemez.",
        "3) Yalnızca kendi satırlarınıza yazın. Değişiklikler otomatik kaydedilir; ayrıca göndermeniz gerekmez.",
        "",
        "SORULAR (soldan sağa): sınav yapılacak mı; dersin gün ve saatinde yapılsın mı; süre; gözetmen; akademik izin günü; ",
        "laboratuvar; listede olmayan bölüm; aynı anda (aynı sınıf, gün ve saatte) yapılacak sınav.",
        "AKADEMİK İZİN GÜNÜ: Sınav haftasında izinli olduğunuz günü seçin, o güne sınavınız konmaz. Birden fazla dersiniz varsa her satıra aynı günü seçin.",
        f"GÖZETMEN: Kayıtlı öğrenci sayısı {GOZETMEN_ESIK}'dan az olan derslerde gözetmen atanmaz; sınava dersin hocası girer.",
        "LABORATUVAR: Sınav bilgisayar laboratuvarında yapılacaksa 'Evet' seçin; gerekmiyorsa boş bırakın.",
        "DERSİ ALAN BÖLÜMLER: Hangi bölümlerin öğrencilerinin dersi aldığı sistemden hazır yazılıdır (gri sütun). Yalnızca o listede",
        "          OLMAYAN bir bölüm de dersi alıyorsa 'Listede olmayan başka bölüm var mı?' sütunundan seçin; bütün fakülte alıyorsa",
        "          'TÜM BÖLÜMLER', birden fazla bölüm eksikse son seçeneği seçin. Eksik yoksa o sütunu boş bırakın.",
        "",
        "Not: Sınav tarihleri ve saatleri fakülte tarafından belirlenir; tercihleriniz mümkün olduğunca karşılanır.",
        "Boş bıraktığınız alanlar 'yazılı sınav / ders saatinde olması şart değil / 50 dk / izin günüm yok' kabul edilir.",
    ]
    for i, t in enumerate(satirlar, 1):
        yon.cell(i, 1, t)
    yon["A1"].font = Font(bold=True, size=14)
    yon.column_dimensions["A"].width = 120

    ws = wb.create_sheet("Dersler")
    ws["A1"] = "Dersleriniz için sarı sütunları doldurun (adınızı 'Öğretim Elemanı' filtresinden seçin)"
    ws["A1"].font = KALIN
    basl = ["Öğretim Elemanı"] + [s[0] for s in SUTUNLAR] + ["Anahtar"]
    gen = [40] + [s[1] for s in SUTUNLAR] + [18]
    for j, (b, w) in enumerate(zip(basl, gen), 1):
        c = ws.cell(TOPLU_ILK, j, b)
        c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), BASLIK_DOLGU, CERCEVE
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[TOPLU_ILK].height = 52
    ws.column_dimensions[get_column_letter(len(basl))].hidden = True
    k = TOPLU_ILK
    for sira, (key, unvan, ad, dersler) in enumerate(tum):
        gri = HAZIR_DOLGU if sira % 2 == 0 else PatternFill("solid", fgColor="D0D0D0")
        for d in dersler:
            k += 1
            degerler = [f"{unvan} {ad}".strip(), d["kod"], d["ad"], d["program"], d["sayi"], d["saat"], d["bolumler"]]
            for j in range(1, len(basl) + 1):
                if j <= len(degerler):
                    v = degerler[j - 1]
                elif j == len(basl):
                    v = f"{key}|{d['kod']}|{d['saat']}"  # aynı ders birden fazla şubeyse saat de anahtarın parçası
                else:
                    v = None
                c = ws.cell(k, j, v)
                c.border = CERCEVE
                c.font = Font(name="Calibri", size=11)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                giris = j > 1 and j <= len(SUTUNLAR) + 1 and SUTUNLAR[j - 2][2]
                c.fill = GIRIS_DOLGU if giris else gri
                if giris:
                    c.protection = Protection(locked=False)
            satir = 1  # satır yüksekliği: en çok satıra ihtiyaç duyan hücreye göre (kapasite ≈ sütun genişliği)
            for j, v in enumerate(degerler, 1):
                if v not in (None, ""):
                    kap = max(1, int(gen[j - 1] * 0.9))
                    satir = max(satir, sum(max(1, math.ceil(len(s) / kap)) for s in str(v).split("\n")))
            ws.row_dimensions[k].height = SATIR_YUKSEKLIGI * satir + 3
    for j, (b, w, giris, secenek) in enumerate(SUTUNLAR, 2):
        if isinstance(secenek, list):
            dv = liste_dv(secenek, ipucu=IPUCU.get(b))
            ws.add_data_validation(dv)
            L = get_column_letter(j)
            dv.add(f"{L}{TOPLU_ILK + 1}:{L}{k}")
    # 'Birlikte yapılacak ders' ve ek bölüm sütunları: serbest metin yerine gizli 'Listeler' sayfasından seçim
    lst = wb.create_sheet("Listeler")
    adlar = sorted({d["ad"] for _, _, _, ds in tum for d in ds})
    for i, a in enumerate(adlar, 1):
        lst.cell(i, 1, a)
    for i, a in enumerate(BOLUMLER, 1):
        lst.cell(i, 2, a)
    lst.sheet_state = "hidden"
    for j, (b, *_r) in enumerate(SUTUNLAR, 2):
        L = get_column_letter(j)
        if b == EK_BASLIK:
            dv = liste_dv(aralik=f"=Listeler!$B$1:$B${len(BOLUMLER)}")
            dv.promptTitle, dv.prompt = "Seçiniz", EK_IPUCU
            ws.add_data_validation(dv); dv.add(f"{L}{TOPLU_ILK + 1}:{L}{k}")
        elif b.startswith("Ek öğrenci"):
            dv = DataValidation(type="whole", operator="between", formula1="0", formula2="500", allow_blank=True,
                                showErrorMessage=True, showInputMessage=True, errorStyle="stop")
            dv.errorTitle, dv.error = "Geçersiz sayı", "0 ile 500 arasında bir sayı girin."
            dv.promptTitle, dv.prompt = "Sayı", "Ek bölümlerden dersi alan öğrenci sayısı (tahmini); bilmiyorsanız boş bırakın."
            ws.add_data_validation(dv); dv.add(f"{L}{TOPLU_ILK + 1}:{L}{k}")
        elif b == AYNI_ANDA_BASLIK:
            dv = liste_dv(aralik=f"=Listeler!$A$1:$A${len(adlar)}")
            dv.promptTitle = "Seçiniz"
            dv.prompt = ("Bu dersin sınavı, başka bir dersin sınavıyla aynı sınıfta, aynı gün ve aynı saatte "
                         "yapılacaksa o dersi listeden seçin. Yoksa boş bırakın.")
            ws.add_data_validation(dv); dv.add(f"{L}{TOPLU_ILK + 1}:{L}{k}")
    ws.freeze_panes = ws.cell(TOPLU_ILK + 1, 4)
    ws.auto_filter.ref = f"A{TOPLU_ILK}:{get_column_letter(len(basl))}{k}"
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False
    ws.protection.formatColumns = False
    ws.page_setup.orientation = "landscape"
    wb.save(yol)


def main():
    filtre = None
    if "--hoca" in sys.argv:
        filtre = tr_upper(sys.argv[sys.argv.index("--hoca") + 1])
    hocalar, rapor, _birimler = ders_listesi()
    os.makedirs(CIKTI, exist_ok=True)
    tum = []
    for key, h in sorted(hocalar.items(), key=lambda kv: (kv[1]["rank"], tr_sira(tr_upper(kv[1]["ad"])))):
        if filtre and filtre not in tr_upper(h["ad"]):
            continue
        dersler = sorted(h["dersler"], key=lambda d: tr_sira(tr_upper(d["ad"])))
        tum.append((key, h["unvan"], h["ad"], dersler))
    yol = os.path.join(CIKTI, f"{DONEM} - Hoca Bilgi Formu (Toplu).xlsx")
    try:
        toplu_form_yaz(yol, tum)
    except PermissionError:  # dosya Excel'de açık
        yol = yol.replace(".xlsx", " (yeni).xlsx")
        toplu_form_yaz(yol, tum)
        print("UYARI: asıl dosya açık olduğu için yeni adla kaydedildi.")
    print(f"{len(tum)} hoca, {sum(len(t[3]) for t in tum)} ders satırı -> {yol}")
    rapor_yol = os.path.join(CIKTI, "_Kontrol Raporu.xlsx")
    try:
        rapor_yaz(rapor_yol, rapor)
    except PermissionError:
        rapor_yol = rapor_yol.replace(".xlsx", " (yeni).xlsx")
        rapor_yaz(rapor_yol, rapor)
    print("Kontrol raporu:", rapor_yol)
    for ad, liste in RAPOR_SAYFALARI.items():
        print(f"  {len(rapor[liste[0]]):3d}  {ad}")


RAPOR_SAYFALARI = {
    "PDF ve OBS'te farklı hoca": ("hoca_farki", ["Ders Kodu", "Ders", "PDF'teki hoca (forma yazılan)", "OBS'teki hoca"]),
    "OBS'te var, PDF'te yok": ("obs_de_var_pdf_de_yok", ["Ders Kodu", "Ders", "OBS'teki hoca", "OBS program", "Kayıtlı öğrenci",
                                                "PDF'te aynı adlı ders var mı?"]),
    "PDF'te var, öğrencisi yok": ("ogrencisiz", ["Ders Kodu", "Ders", "Hoca", "Program", "PDF saati"]),
    "PDF'te kodsuz dersler": ("kodsuz", ["Ders", "Hoca", "Program", "OBS'te eşleşen kod"]),
    "PDF ve OBS'te farklı kod": ("kod_farki", ["PDF kodu", "Ders", "Hoca", "OBS'te aynı adlı kod", "Formda/çakışmada kullanılan kod"]),
}


def rapor_yaz(yol, rapor):
    """Komisyonun teyit edeceği farklar: PDF esas alındı, OBS eski olabilir."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for ad, (anahtar, baslik) in RAPOR_SAYFALARI.items():
        ws = wb.create_sheet(ad[:31])
        ws.append(baslik)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), BASLIK_DOLGU
        for satir in sorted(rapor[anahtar], key=lambda r: tr_sira(tr_upper(str(r[0])))):
            ws.append(list(satir))
        for j in range(1, len(baslik) + 1):
            ws.column_dimensions[get_column_letter(j)].width = 34
        ws.freeze_panes = "A2"
    wb.save(yol)


if __name__ == "__main__":
    main()
