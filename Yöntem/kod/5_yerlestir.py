# -*- coding: utf-8 -*-
"""
Aşamalı sınav yerleştirme (2026-2027 Güz Vize).  Her aşama sonunda kullanıcı kontrol eder.

  Aşama 1: ortak dersler (3 ve daha fazla bölümün aldığı dersler)

Girdi : Kullanılacak Belgeler/ (Ders Programı PDF, Öğrenci sayıları, altan ve üsten, Hoca Bilgi Formu)
Çıktı : Çıktılar/Sınav Programı/Aşama N - ....xlsx   (öğrenci numarası içermez)

Kullanım:  py -3 Yöntem\\kod\\5_yerlestir.py --asama 1
Kurallar (kısa): bkz. 'Özel Kısıtlar' sekmesi (2026-2027 Güz Vize - Şablon.xlsx).
"""
import os, sys, re, itertools, collections, argparse, warnings, functools
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kaynaklar as K  # noqa: E402

FORM = os.path.join(K.KOK, "Kullanılacak Belgeler", "2026-2027 Güz Vize - Hoca Bilgi Formu (Toplu) (1).xlsx")
CIKTI = os.path.join(K.KOK, "Çıktılar", "Sınav Programı")

# ---------------------------------------------------------------- dönem ayarları
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
TARIHLER = ["02.11.2026", "03.11.2026", "04.11.2026", "05.11.2026", "06.11.2026"]
SAATLER = list(range(9, 18))                       # oturum başlangıçları 09:00 ... 17:00
YASAK = {(2, 9), (2, 10), (2, 11)}                 # Çarşamba 09-12'ye sınav konmaz
DERSLIKLER = {"6A1": 35, "6A2": 33, "6D1": 17, "6D3": 23, "6D4": 25, "6D5": 23}   # sınav kapasitesi (6. kat)
TAM = {"6A1": 50, "6A2": 47, "6D1": 29, "6D3": 32, "6D4": 40, "6D5": 37}   # tam (yan yana) kapasite: kesin üst sınır
LABLAR = {"LAB": 40}                                # bilgisayar laboratuvarı (sınav kapasitesi 40; Güz bilgisi gelince güncellenecek)
TAM["LAB"] = 40
ESNEME_TOPLA = 8                                    # bir dersin gruplarını daha az dersliğe toplamak için kabul edilen esneme (kullanıcı örüntüsü)
ESNEME = 5                                         # normal esneme payı
ESNEME_BUYUK = 10                                  # >=100 öğrencili derslerde gerekirse
BUYUK = 100
MIN_PARCA = 8                                      # bir grup bölünecekse parça en az bu kadar kişi olur
# hoca anahtarı -> sert izin günü (formdaki yanlış girişleri geçersiz kılar)
IZIN_DUZELT = {"ELDEM|A": {0}, "HAJIAMIRI|M": {0}, "ZENGINKUZUCU|D": {3}}   # Dikran Zenginkuzucu: Perşembe sınav yok (20 Ekim bildirimi)
SABIT = {}                                          # (İşletmeye Giriş Cuma 12:00 sabiti 20 Ekim'de kaldırıldı: Zekeriya Şahin izin günü Cuma)
# Öğrencisi kayıtta görünmeyen dersler (kullanıcı bildirdi): kod -> (öğrenci, çakışma grubu bölümü)
ELLE_SAYI = {"20EG4009": 15, "2023G2006": 12, "20G1009": 8}
# Kullanıcının "bu bölüm bu dersi fiilen almıyor" dediği gruplar (OBS'te birkaç öğrenci görünüyor): derslik dosyasında yazılmaları beklenmez
# Kullanıcının "sayılar önemli değil, bölüm çok kişi olduğu için böldük" dediği dersler: hücredeki sayılara bakılmaz, yalnız hangi bölümün hangi derslikte
# olduğu okunur; öğrenci sayısı derste yazılı bölümlere ve derslikleri eşit paylaştırılır (22 Ekim)
SAYI_ONEMSIZ = {"Bölgesel Çalışmalar : Avrasya"}
DERSI_ALMAYAN = {"Bölgesel Çalışmalar : Avrasya": {"Yönetim Bilişim Sistemleri", "Elektronik Ticaret ve Yönetimi"}}   # 22 Ekim'de kullanıcı bildirdi
# Okul tarafından yapılan (programa girmeyen) dersler
OKUL_YAPAR = {"2075EG1002", "2075B1002", "2075EG1001", "2075G1001", "2075G1002", "2075G1005"}
# Ortak ders sayılmayacak (kullanıcı kararı) - şimdilik yok
# Aynı derslikte karışmaması gereken öğrenci grupları (aynı kodlu dersi veren farklı hocalar)
AYRI = {"20G1010": [{"İşletme", "Lojistik Yönetimi", "Yönetim Bilişim Sistemleri"},
                    {"Uluslararası Ticaret ve Finansman", "Uluslararası Ticaret ve Finans", "Elektronik Ticaret ve Yönetimi"}]}
# Sona bırakılan dersler: büyük derslik (3. kat) durumu belli olunca yerleştirilecek
ERTELE = {"20G1003", "20G1007"}                # İktisada Giriş I, Sosyal Bilimler için Matematik I
# Kullanıcının elle düzeltmeleri (derslik dosyasında değiştirdi): ders kodu -> {grup (bölüm adı, "*" = hepsi): derslik}
ELLE_GRUP = {"20G3004": {"*": "6A1"},                      # Tedarik Zinciri Yönetimi: hepsi tek derslik (6A1, 40 kişi)
             "20G1004": {"İşletme": "6A1"},                # BİT: İşletme 41 kişi 6A1'e
             "2014G4009": {"*": "6A2"}}                    # Managerial Leadership: 6A2'ye
# Aynı derslikte, aynı gün ve saatte yapılacak Türkçe/İngilizce karşılıklar (hoca talebi): tek sınav, birleşik ad
BIRLESTIR = [(("2017G3024", "2016G3030"), "History of European Integration / Avrupa Bütünleşme Tarihi"),
             (("2017G3003", "2016G3003"), "International Law-I / Uluslararası Hukuk I")]
# Eski (TALEP_ONCELIK=False iken) kullanıcı onaylı yerleşimler (gün, saat): yalnızca talep önceliği kapalıyken denenir
PRIORITY_HOCA = {"ZENGINKUZUCU|D"}               # öncelikli hocalar: talepleri sabit (KILIT) sınavlardan hemen sonra, diğerlerinden önce yerleşir
KILIT = {"20EG1004": (0, 11), "20G1002": (0, 12), "2020G4004": (0, 15), "20G1004": (0, 16), "20G2006": (1, 13),
         "2016G4023": (1, 15), "2016G3003": (2, 12), "20EG2006": (3, 12), "2020G3004": (3, 12),
         "2016G3030": (3, 13), "20EG2002": (4, 11), "20G3004": (4, 11), "20G2002": (4, 13),
         "2014G4009": (4, 14), "20G1003": (0, 13), "20G1007": (2, 13)}   # aşama 1 (onaylı) + sona bırakılanların geçici yeri
# Ortak olmasa da hoca talebiyle şimdiden yerleştirilen sınavlar (kod) ve gün kısıtı (0=Pazartesi)
EKLE = {"2016G3003"}                                  # International Law-I / Uluslararası Hukuk I (Zenginkuzucu)
GUN_KISITI = {"2016G3003": {2}}                       # Çarşamba (ders saati 12-15)
# --- aşama 2 kuralları ---
SABIT["2022G4001"] = (1, 9)                           # Derivatives and Risk Management (Sudi Apak): Salı 09:00, 6D5, kullanıcı elle girecek
ODA_KISIT = {"2022G4001": {"6D5"}}
IZIN_DUZELT["GÜLBAŞ|E"] = {3, 4}                      # Erdinç Gülbaş: Perşembe ve Cuma sınav olmasın
TERCIH_SAAT = {"GÜLBAŞ|E": 17}                        # Erdinç Gülbaş: sınavlar 17:00-18:00 arasına
SINAV_YOK = {"2014G4001", "2014B4002"}                # Consumer Behaviour, Brand Management: ödev
HOCA_DOLU = {"ŞAHIN|A": {(3, 15), (3, 16), (3, 17)}}  # Alperen Şahin: Perşembe 15-18 Sanat Fakültesi dersi
HOCA_DEGIS = {"EROL|U": ("HAJIAMIRI|M", "Misagh HAJIAMIRI")}   # Ufuk Erol'un derslerine Misagh Hajiamiri girer
TALEP_YOK = {"2018G2003", "2018G4002"}                # Erol dersleri: Pazartesi hariç herhangi gün (ders saati talebi yok)
for _k in ("2023G2005", "2023G3009", "2023G3007", "2023G4009", "20G1009"):    # Meysure Evren Çelik Sütiçer: Salı
    GUN_KISITI[_k] = {1}
ELLE_BOLUM = {"20EG4009": {"Ekonomi ve Finans"}, "2023G2006": {"Elektronik Ticaret ve Yönetimi"},
              "20G1009": {"Elektronik Ticaret ve Yönetimi"}}
CAP_YUMUSAK = [{"İşletme", "Yönetim Bilişim Sistemleri"}, {"Siyaset Bilimi ve Uluslararası İlişkiler", "Havacılık Yönetimi"}]
CAP_BOLUMLER = {"Havacılık Yönetimi", "Lojistik Yönetimi", "Elektronik Ticaret ve Yönetimi", "İşletme",
                "Yönetim Bilişim Sistemleri", "Siyaset Bilimi ve Uluslararası İlişkiler"}
CAP_SERT = [{"Havacılık Yönetimi", "Lojistik Yönetimi"}, {"Elektronik Ticaret ve Yönetimi", "Havacılık Yönetimi"}]


KAPASITE = {**DERSLIKLER, **LABLAR}


def bolum(p):
    return re.sub(r"\s*\((İngilizce|İng\.)\)", "", p).strip()


def slot_adi(s, sure=50):
    g, h = divmod(s, len(SAATLER))
    h += SAATLER[0]
    son = f"{h:02d}:50" if sure <= 50 else f"{h + 1:02d}:30"
    return f"{GUNLER[g]} {TARIHLER[g][:5]} {h:02d}:00-{son}"


def yasak_mi(s):
    g, h = divmod(s, len(SAATLER))
    return (g, h + SAATLER[0]) in YASAK


# ---------------------------------------------------------------- veri
def formu_oku(kt):
    rows = list(openpyxl.load_workbook(FORM)["Dersler"].iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r and r[0] == "Öğretim Elemanı")
    out = []
    for r in rows[hi + 1:]:
        if not r[0] or str(r[0]).startswith("—"):
            continue
        out.append({"hoca": kt.anahtar(r[0]), "hoca_ad": r[0], "kodlar": K.KOD_RE.findall(str(r[1] or "")),
                    "ad": r[2], "sinav": r[7], "gun_saat": r[8], "sure": r[9], "gozetmen": r[10], "izin": r[11],
                    "lab": r[12], "ek": r[14] if isinstance(r[14], (int, float)) else 0})
    return out


def sinif_oku():
    """kod -> program -> en sık 'Ders Sınıf' (dersin müfredattaki sınıfı)"""
    ws = openpyxl.load_workbook(K.KAYIT_DOSYASI, read_only=True).active
    it = ws.iter_rows(values_only=True)
    h = [re.sub(r"_\d+$", "", str(x).strip()) for x in next(it)]
    ik, ip, isn = h.index("D.Kodu"), h.index("Öğrenci Programı"), h.index("Ders Sınıf")
    out = collections.defaultdict(lambda: collections.defaultdict(collections.Counter))
    for r in it:
        if r[ik] and r[isn]:
            out[str(r[ik]).strip()][str(r[ip]).strip()][str(r[isn]).strip()] += 1
    return out


def program_hoca_haritasi(kt):
    """(OBS ders kodu, program) -> hoca anahtarı.  Şube listesinde hoca boşsa aynı dersin bir önceki satırındaki hoca."""
    out, onceki = {}, {}
    for r in K.subeleri_oku():
        h = kt.anahtar(r["hoca"]) if r["hoca"] else onceki.get(r["kod"], "")
        onceki[r["kod"]] = h
        out[(r["kod"], r["program"])] = h
    return out


# OBS'te aynı dersin ikinci bir kodu (PDF'teki koddan farklı) var: o öğrenciler de sınava girer.  ders adı (ad_anahtar) -> ek OBS kodları
EK_OBS = {"sehirvekultur": ["2075B3001"], "finansalyonetim": ["2010G2003"],
          "orgutseldavranisarastirmalari": ["2010G2002"], "havaciliktastratejikyonetim": ["2018G3008"],
          "depolamavenvanteryonetimi": ["LOJ211"], "istatistik": ["20G2002[ISL203]"]}


def birimleri_duzelt(birimler, ogr):
    """PDF kodu başka bir OBS dersine denk geldiği için yanlış birleşen iki durumu ayırır."""
    out = []
    for b in birimler:
        if b["ad"] == "Derivatives and Risk Management" and b["kodlar"] == ["2022G4001"]:
            dv = dict(b)                              # OBS'te Derivatives = 2022G4003 (2022G4001 Political Economics'in kodu)
            dv.update(tum={"2022G4003"}, ogr=set(ogr.get("2022G4003", set())), _kodsuz=True, _isim=True)
            out.append(dv)
        elif b["ad"] == "Political Economics":
            pe = dict(b)
            pe.update(_isim=True)
            out.append(pe)
        elif b["ad"].startswith("Endüstri 4.0") and "2019G4008" in b["kodlar"]:
            en = dict(b)
            en.update(tum={"2019G4010"}, ogr=set(ogr.get("2019G4010", set())), _kodsuz=True, _isim=True)
            li = dict(b)
            li.update(ad="Liman ve Terminal Yönetimi", tum={"2019G4008"}, ogr=set(ogr.get("2019G4008", set())), _isim=True)
            out += [en, li]
        else:
            out.append(b)
    return out


SPLIT_HOCA = {"2017G4021"}
# Introduction to Economics-I: önce Bayraktar (İşletme İng., SBUİ İng.) ve Şimşek (ECF) ayrı sınavdı; 23 Ekim'de Bayraktar'ın dersi vermediği
# bildirildi (hoca maili): ders İlknur Şimşek'te, üç grup tek sınav
DERS_HOCASI = {"20EG1003": "ŞIMŞEK|I", "20B1003": "ŞIMŞEK|I"}
INTFIN_PROG = ("İşletme (İngilizce)", "Siyaset Bilimi ve Uluslararası İlişkiler (İngilizce)", "Ekonomi ve Finans (İngilizce)")   # International Finance: bütün İngilizce bölümler


def hocaya_gore_ayir(sinavlar, kayit_prog, hoca_tbl):
    out = []
    for e in sinavlar:
        if not any(k in SPLIT_HOCA for k in e["kodlar"]):
            out.append(e)
            continue
        by = collections.defaultdict(list)
        for pr, h in e["prog_hoca"].items():
            by[h or "?"].append(pr)
        ps = {pr: set().union(*(kayit_prog.get(k, {}).get(pr, set()) for k in e["tum"])) for pr in e["prog"]}
        for h, progs in by.items():
            ogr2 = set().union(*(ps[pr] for pr in progs))
            gr, saat = set(), []
            for d in hoca_tbl.get(h, {}).get("dersler", []):
                if d["kod"] in e["kodlar"]:
                    gr |= {x.strip() for x in d["program"].split("\n") if x.strip()}
                    if d["saat"]:
                        saat.append(d["saat"])
            sd = []
            for t in saat:
                for m in re.finditer(r"(Pazartesi|Salı|Çarşamba|Perşembe|Cuma)\s+(\d{2}):\d{2}-(\d{2}):\d{2}", t):
                    sd.append((GUNLER.index(m.group(1)), int(m.group(2)), int(m.group(3))))
            adlar_h = {hh: next((d["ad"] for d in hoca_tbl.get(hh, {}).get("dersler", []) if d["kod"] in e["kodlar"]), e["ad"])
                       for hh in by}
            farkli = len({K.ad_anahtar(a) for a in adlar_h.values()}) > 1     # hocaların dersleri farklı adlı: kendi adıyla
            e2 = dict(e)
            e2.update(ad=adlar_h[h] if farkli else f"{e['ad']} ({h.split('|')[0].title()})", hocalar=[x for x in e["hocalar"] if x[0] == h] or e["hocalar"],
                      prog={pr: len(ps[pr]) for pr in progs}, ogr=ogr2, n=len(ogr2),
                      bolum={bolum(pr) for pr in progs if len(ps[pr]) >= 3},
                      form=[f for f in e["form"] if f["hoca"] == h] or e["form"],
                      sinif={pr: v for pr, v in e["sinif"].items() if pr in progs}, prog_hoca={pr: h for pr in progs},
                      gruplar=gr or e["gruplar"], ders_saat=sorted(set(sd)) or e["ders_saat"])
            out.append(e2)
    return out


def sinavlari_kur():
    """-> (sınav listesi, rapor).  Aynı dersin farklı kodlu/hocalı kayıtları tek sınavda birleşir."""
    hoca_tbl, rapor, birimler = K.ders_listesi()
    ogr, kayit_ad, tip, kayit_prog = K.kayitlari_oku()
    kt = K.KisiTablosu()
    form = formu_oku(kt)
    birimler = birimleri_duzelt(birimler, ogr)
    # 1) birimleri ortak kodla birleştir
    ebeveyn = list(range(len(birimler)))

    def bul(i):
        while ebeveyn[i] != i:
            ebeveyn[i] = ebeveyn[ebeveyn[i]]; i = ebeveyn[i]
        return i
    kod_sahibi = {}
    for i, b in enumerate(birimler):
        for k in (set(b["tum"]) if b.get("_kodsuz") else set(b["kodlar"]) | set(b["tum"])):
            if k in kod_sahibi:
                if K.ad_anahtar(birimler[kod_sahibi[k]]["ad"]) == K.ad_anahtar(b["ad"]):   # yalnızca aynı adlı dersler birleşir
                    ebeveyn[bul(i)] = bul(kod_sahibi[k])
            else:
                kod_sahibi[k] = i
    for kodlar_, _ad in BIRLESTIR:
        sahipler = [kod_sahibi[k] for k in kodlar_ if k in kod_sahibi]
        for o in sahipler[1:]:
            ebeveyn[bul(o)] = bul(sahipler[0])
    kume = collections.defaultdict(list)
    for i in range(len(birimler)):
        kume[bul(i)].append(birimler[i])
    # PDF'ten program satırları: kod -> program/yarıyıl grupları
    grup_kod = collections.defaultdict(set)
    for v in hoca_tbl.values():
        for d in v["dersler"]:
            for s in d["program"].split("\n"):
                s = s.strip()
                if s:
                    grup_kod[d["kod"]].add(s)
    saat_kod = collections.defaultdict(list)
    for v in hoca_tbl.values():
        for d in v["dersler"]:
            if d["saat"]:
                saat_kod[d["kod"]].append(d["saat"])
    saat_ad = collections.defaultdict(list)        # PDF'te kodsuz ya da farklı kodlu satırlar için ders adına göre
    for r in K.pdf_oku():
        if r["saat"]:
            saat_ad[K.ad_anahtar(r["ad"])].append(r["saat"])
            if r["kod"]:
                saat_kod[r["kod"]].append(r["saat"])
    sinif_tbl = sinif_oku()
    ph = program_hoca_haritasi(kt)
    sinavlar = []
    for uyeler in kume.values():
        kodlar = sorted({k for b in uyeler for k in b["kodlar"]})
        tum = {k for b in uyeler for k in b["tum"]} | {k for b in uyeler if not b.get("_kodsuz") for k in b["kodlar"]}
        ogrenci = set().union(*(b["ogr"] for b in uyeler))
        for k in EK_OBS.get(K.ad_anahtar(uyeler[0]["ad"]), []):
            tum.add(k)
            ogrenci |= set(ogr.get(k, set()))
        prog = collections.defaultdict(set)
        for k in tum:
            for p, s in kayit_prog.get(k, {}).items():
                prog[p] |= s
        hocalar = []
        for b in uyeler:
            for h in b["hocalar"]:
                a = kt.anahtar(h)
                if a and a not in [x[0] for x in hocalar]:
                    hocalar.append((a, h))
        pdf_h = [(hk, v["ad"]) for hk, v in hoca_tbl.items()
                 if any(d["kod"] in (tum | set(kodlar)) and (K.ad_anahtar(d["ad"]) in {K.ad_anahtar(b_["ad"]) for b_ in uyeler}
                                                              or any(k in SPLIT_HOCA for k in kodlar))
                        for d in v["dersler"])]
        if pdf_h:
            hocalar = pdf_h
        hocalar = [HOCA_DEGIS.get(a, (a, h)) for a, h in hocalar]
        for k_, h_ in DERS_HOCASI.items():
            if k_ in kodlar and any(a == h_ for a, _ in hocalar):
                hocalar = [x for x in hocalar if x[0] == h_]
        gruplar = set()
        for k in tum | set(kodlar):
            gruplar |= grup_kod.get(k, set())
        # form satırları
        fr = [f for f in form if set(f["kodlar"]) & (tum | set(kodlar))
              and (not any(b.get("_isim") for b in uyeler) or K.ad_anahtar(f["ad"]) == K.ad_anahtar(uyeler[0]["ad"]))]
        sure = max([90 if (f["sure"] or "").startswith("90") else 50 for f in fr] or [50])
        lab = any((f["lab"] or "").startswith("Evet") for f in fr) and (len(ogrenci) or ELLE_SAYI.get(next((k for k in kodlar if k in ELLE_SAYI), ""), 0)) <= LABLAR["LAB"]
        hayir = bool(fr) and all((f["sinav"] or "").startswith("Hayır") for f in fr if f["sinav"]) and any(f["sinav"] for f in fr)
        ex = ELLE_SAYI.get(next((k for k in kodlar if k in ELLE_SAYI), ""), 0)
        sc = collections.defaultdict(collections.Counter)
        for k in tum:
            for pr, c in sinif_tbl.get(k, {}).items():
                sc[pr] += c
        sinif = {pr: int(c.most_common(1)[0][0]) for pr, c in sc.items() if c.most_common(1)[0][0].isdigit()}
        prog_hoca = {}
        for pr in prog:
            hk = next((v for (kd, pn), v in ph.items() if kd in tum and v and
                       (K.ad_anahtar(pn).startswith(K.ad_anahtar(pr)[:22]) or K.ad_anahtar(pr).startswith(K.ad_anahtar(pn)[:22]))), "")
            prog_hoca[pr] = hk
        ders_saat = []
        adlar_ = {K.ad_anahtar(b_["ad"]) for b_ in uyeler}
        metinler = [t for k in tum | set(kodlar) for t in saat_kod.get(k, [])] + [t for a_ in adlar_ for t in saat_ad.get(a_, [])]
        for k in [0]:
            for t in metinler:
                for m in re.finditer(r"(Pazartesi|Salı|Çarşamba|Perşembe|Cuma)\s+(\d{2}):\d{2}-(\d{2}):\d{2}", t):
                    ders_saat.append((GUNLER.index(m.group(1)), int(m.group(2)), int(m.group(3))))
        ad_son = uyeler[0]["ad"]
        for kd, bad in BIRLESTIR:
            if set(kd) & (tum | set(kodlar)):
                ad_son = bad
        sinavlar.append({"ders_saat": sorted(set(ders_saat)), "prog_hoca": prog_hoca, "sinif": sinif, "kodlar": kodlar, "tum": tum, "ad": ad_son, "hocalar": hocalar, "ogr": ogrenci,
                         "prog": {p: len(s) for p, s in prog.items()}, "gruplar": gruplar,
                         "n": len(ogrenci) or ex, "sure": sure, "lab": lab, "var": not hayir and not any(k in SINAV_YOK for k in kodlar), "form": fr,
                         "bolum": {bolum(p) for p, s in prog.items() if len(s) >= 3} or
                                  {b for k in kodlar for b in ELLE_BOLUM.get(k, set())}})
    # hoca izinleri (hocanın bütün satırlarındaki gün birleşimi)
    gun_no = {g: i for i, g in enumerate(GUNLER)}
    izin = collections.defaultdict(set)
    for f in form:
        if f["izin"] in gun_no:
            izin[f["hoca"]].add(gun_no[f["izin"]])
    izin.update(IZIN_DUZELT)
    for e in sinavlar:
        e["izin"] = set().union(*(izin.get(h, set()) for h, _ in e["hocalar"]), set())
    sinavlar = hocaya_gore_ayir(sinavlar, kayit_prog, hoca_tbl)
    for e in sinavlar:
        if "2022EG3018" in e["kodlar"]:        # 23 Ekim: Bayraktar'ın "Tourism Economics" dersi aslında International Finance (hoca maili); öğrenci sayısı bilinmiyor
            e.update(ad="International Finance", var=True, bolum={"İşletme", "Siyaset Bilimi ve Uluslararası İlişkiler", "Ekonomi ve Finans"},
                     prog={p: 0 for p in INTFIN_PROG}, sinif={"Ekonomi ve Finans (İngilizce)": 4})
    for e in sinavlar:
        e["izin"] = set().union(*(izin.get(h, set()) for h, _ in e["hocalar"]), set())
    return sinavlar, izin, kt


# ---------------------------------------------------------------- kurallar
# Meysure Evren Çelik Sütiçer'in dersleri birbiriyle "aynı bölüm-yarıyıl / aynı hoca" kuralından muaf (hoca aynı gün, aynı sınıfta,
# art arda yapmayı kabul etti; ortak öğrencisi olanlar yine ayrı oturumda)
MUAF_KODLAR = {"2023G2005", "2023G3009", "2023G3007", "2023G4009", "20G1009"}
CAP_MOD = "capraz"          # "capraz": yalnızca iki bölümün dersleri birbiriyle; "birlesim": iki bölümün bütün dersleri birbiriyle


def cap_capraz(a, b, cift):
    if CAP_MOD == "birlesim":
        return bool(a["bolum"] & cift and b["bolum"] & cift)
    x, y = sorted(cift)
    return bool((x in a["bolum"] and y in b["bolum"]) or (y in a["bolum"] and x in b["bolum"]))


def cakisir(a, b):
    """İki sınav aynı anda olamaz mı?"""
    if a["ogr"] & b["ogr"]:
        return "ortak öğrenci"
    muaf = any(k in MUAF_KODLAR for k in a["kodlar"]) and any(k in MUAF_KODLAR for k in b["kodlar"])
    if a["gruplar"] & b["gruplar"] and not muaf:
        return "aynı bölüm-yarıyıl"
    if {h for h, _ in a["hocalar"]} & {h for h, _ in b["hocalar"]} and not muaf:
        return "aynı hoca"
    for cift in CAP_SERT:
        if cap_capraz(a, b, cift):
            return "ÇAP çifti"
    return ""


def paketle(gruplar, bos, esneme, ayri, sirala=True, max_parca=9):
    """gruplar: [(ad, n)]  bos: {derslik: kapasite}  ayri: {grup bölümü: küme no} (farklı kümeler aynı odaya girmez)
    -> {derslik: [(grup, n)]} ya da None.  Önce az derslik, bölüm mümkünse tek derslikte."""
    kalan = {r: min(c + esneme, TAM[r]) for r, c in bos.items()}
    yer = {}

    def uyumlu(r, ad):
        k = ayri.get(bolum(ad))
        return all(ayri.get(bolum(g)) == k for g, _ in yer.get(r, []))

    for ad, n in (sorted(gruplar, key=lambda x: -x[1]) if sirala else gruplar):
        uygun = [r for r in kalan if kalan[r] >= n and uyumlu(r, ad)]
        if uygun:
            r = min(uygun, key=lambda r: (not yer.get(r), kalan[r] - n, r))
            yer.setdefault(r, []).append((ad, n))
            kalan[r] -= n
            continue
        # sığmadı: grup en az sayıda parçaya EŞİT bölünür (2-3 kişilik artık parça olmaz)
        yerlesti = False
        for k in range(2, min(len(kalan), max_parca) + 1):
            parca = [n // k + (1 if i < n % k else 0) for i in range(k)]
            if min(parca) < MIN_PARCA:
                break
            kopya = dict(kalan)
            atama = []
            for pn in sorted(parca, reverse=True):
                uyg = [r for r in kopya if kopya[r] >= pn and uyumlu(r, ad) and r not in [x[0] for x in atama]]
                if not uyg:
                    atama = None
                    break
                r = min(uyg, key=lambda r: (not (yer.get(r) or any(x[0] == r for x in atama)), kopya[r] - pn, r))
                atama.append((r, pn))
                kopya[r] -= pn
            if not atama:                               # eşit parça sığmadı: derslik boşluğuna orantılı (eşitsiz) parçalar
                aday = sorted((r for r in kalan if kalan[r] > 0 and uyumlu(r, ad)), key=lambda r: -kalan[r])[:k]
                top = sum(kalan[r] for r in aday)
                if len(aday) == k and top >= n:
                    pc = {r: max(0, min(kalan[r], n * kalan[r] // top)) for r in aday}
                    kalanlar = n - sum(pc.values())
                    for r in aday:
                        ek = min(kalanlar, kalan[r] - pc[r])
                        pc[r] += ek
                        kalanlar -= ek
                    if kalanlar == 0 and min(pc.values()) >= MIN_PARCA:
                        atama = list(pc.items())
            if atama:
                for r, pn in atama:
                    yer.setdefault(r, []).append((ad, pn))
                    kalan[r] -= pn
                yerlesti = True
                break
        if not yerlesti:
            return None
    return {r: l for r, l in yer.items() if l}


def derslik_ata(e, bos):
    gruplar = list(e["prog"].items()) if e["prog"] else [("Öğrenci", e["n"])]
    toplam = sum(n for _, n in gruplar)
    if toplam < e["n"]:
        gruplar.append(("Diğer", e["n"] - toplam))
    ayri = {}
    for kod in e["kodlar"]:
        for k, kume in enumerate(AYRI.get(kod, [])):
            for b in kume:
                ayri[b] = k
    tek = any(k in EKLE for k in e["kodlar"])            # hoca "aynı sınıf" istedi: tek dersliğe sığıyorsa en az esnemeli olan
    return _en_iyi(tuple(gruplar), tuple(sorted(bos.items())), e["n"] >= BUYUK, tuple(sorted(ayri.items())), tek)


@functools.lru_cache(maxsize=4096)
def _en_iyi(gruplar, bos_t, buyuk, ayri_t, tek=False):
    bos, ayri = dict(bos_t), dict(ayri_t)
    if tek:
        toplam = sum(n for _, n in gruplar)
        aday = [(max(0, toplam - c), c, r) for r, c in bos.items() if toplam <= min(TAM[r], c + ESNEME_BUYUK)]
        if aday:
            asma, _, r = min(aday)
            return {r: list(gruplar)}, next(es for es in (0, ESNEME, ESNEME_BUYUK) if es >= asma)
    # ÖNCE esneme olmadan (gerekirse bölerek), olmazsa en az esnemeyle: esneme son çare
    # Önce bölünme en az, sonra en az derslik, sonra en az esneme.  Esneme toplama sınırı: ESNEME_TOPLA (+8); büyük derslerde +20.
    ladder = [0, 2, 4, 6, ESNEME_TOPLA] if not buyuk else [0, 2, 4, 6, ESNEME_TOPLA, 12, 16, 20]
    for mp in (2, 9):                                   # her grup en çok 2 derslikte; olmazsa daha fazla
        en = None
        for es in ladder:
            denemeler = [list(gruplar)]
            y0 = paketle(list(gruplar), bos, es, ayri, max_parca=mp)
            say0 = collections.Counter(g for l in (y0 or {}).values() for g, _ in l)
            if y0 is None or sum(c - 1 for c in say0.values()) > 0:
                denemeler += [list(x) for x in itertools.islice(itertools.permutations(gruplar), 720)]
            for i, sira in enumerate(denemeler):
                y = paketle(sira, bos, es, ayri, sirala=(i == 0), max_parca=mp)
                if not y:
                    continue
                say = collections.Counter(g for l in y.values() for g, _ in l)
                diger = sum(c - 1 for g, c in say.items() if bolum(g) != "İşletme")   # İşletme bilerek bölünebilir
                puan = (diger, sum(c - 1 for c in say.values()), len(y), es)
                if en is None or puan < en[0]:
                    en = (puan, y, es)
        if en:
            return en[1], en[2]
    return None, None


# ---------------------------------------------------------------- yerleştirme
class Program:
    def __init__(self):
        self.yer = {}                                   # id(e) -> (slot, derslikler, esneme)
        self.kayit = collections.defaultdict(list)      # slot -> [sınav]
        self.dolu = collections.defaultdict(dict)       # slot -> {derslik: kapasite kullanılan}
        self.ogr_sl = collections.defaultdict(set)      # (öğrenci no, gün) -> o gün sınavı olan saat indeksleri
        self.ardisik_sert = True                        # bir öğrenci günde 3 ARDIŞIK sınava girmez

    def slotlar(self, e, s):
        return [s + i for i in range(2 if e["sure"] > 50 else 1)]

    def ardisik_ogrenci(self, e, s):
        """e s'ye konursa herhangi bir öğrencinin aynı gün 3 ardışık oturumu dolar mı?"""
        g = s // len(SAATLER)
        yeni = [x % len(SAATLER) for x in self.slotlar(e, s)]
        for no in e["ogr"]:
            occ = self.ogr_sl.get((no, g))
            if not occ:
                continue
            t = occ | set(yeni)
            for hh in yeni:
                for a in (hh - 2, hh - 1, hh):
                    if all((a + i) in t for i in range(3)):
                        return True
        return False

    def uygun(self, e, s, sabit=False):
        g, h = divmod(s, len(SAATLER))
        sl = self.slotlar(e, s)
        if e["sure"] > 50 and h + 1 >= len(SAATLER):
            return None, "90 dk saat dışı"
        gk = next((GUN_KISITI[k] for k in e["kodlar"] if k in GUN_KISITI), None)
        if gk and g not in gk:
            return None, "gün kısıtı"
        if g in e["izin"] and not sabit:
            return None, "izin günü"
        if self.ardisik_sert and not sabit and self.ardisik_ogrenci(e, s):
            return None, "3 ardışık sınav"
        for x in sl:
            if yasak_mi(x):
                return None, "Çarşamba 09-12"
            gx, hx = divmod(x, len(SAATLER))
            for hk, _ in e["hocalar"]:
                if (gx, hx + SAATLER[0]) in HOCA_DOLU.get(hk, ()):
                    return None, "hoca dolu"
        bos = dict(LABLAR) if e.get("lab") else dict(DERSLIKLER)     # laboratuvar isteyen sınav yalnızca LAB'a girer
        kisit = next((ODA_KISIT[k] for k in e["kodlar"] if k in ODA_KISIT), None)
        if kisit:
            bos = {r: c for r, c in bos.items() if r in kisit}
        for x in sl:
            for o in self.kayit.get(x, []):
                if cakisir(e, o):
                    return None, "çakışma: " + o["ad"][:25]
                if e["n"] >= BUYUK and o["n"] >= BUYUK:
                    return None, "iki büyük ders"
            for r in self.dolu.get(x, {}):
                bos.pop(r, None)
        y, es = derslik_ata(e, bos)
        if not y:
            return None, "derslik yok"
        return (y, es), ""

    def cikar(self, e):
        s, y, es = self.yer.pop(id(e))
        for x in self.slotlar(e, s):
            for no in e["ogr"]:
                self.ogr_sl[(no, x // len(SAATLER))].discard(x % len(SAATLER))
            self.kayit[x].remove(e)
            for r in y:
                self.dolu[x].pop(r, None)
        return s

    def koy(self, e, s, y, es):
        self.yer[id(e)] = (s, y, es)
        for x in self.slotlar(e, s):
            for no in e["ogr"]:
                self.ogr_sl[(no, x // len(SAATLER))].add(x % len(SAATLER))
            self.kayit[x].append(e)
            for r in y:
                self.dolu[x][r] = True


def elle_duzelt(P, sec):
    """ELLE_GRUP düzeltmelerini, derslik o oturumda boşsa ve tam kapasiteyi aşmıyorsa uygular. -> uygulanamayanlar"""
    uyarilar = []
    for e in sec:
        kod = next((k for k in e["kodlar"] if k in ELLE_GRUP), None)
        if kod is None or id(e) not in P.yer:
            continue
        s_, y, es = P.yer[id(e)]
        yeni = {r: list(l) for r, l in y.items()}
        for grup, hedef in ELLE_GRUP[kod].items():
            tasi = []
            for r in list(yeni):
                kalan = []
                for a_, n in yeni[r]:
                    (tasi if grup == "*" or bolum(a_) == grup else kalan).append((a_, n))
                yeni[r] = kalan
            yeni = {r: l for r, l in yeni.items() if l}
            yeni.setdefault(hedef, []).extend(tasi)
        sl = P.slotlar(e, s_)
        baska = {r for x in sl for o in P.kayit.get(x, []) if o is not e for r in P.yer[id(o)][1]}
        if any(sum(n for _, n in l) > TAM[r] for r, l in yeni.items()) or set(yeni) & baska:
            uyarilar.append(f"{e['ad']}: elle düzeltme bu oturumda uygulanamadı")
            continue
        for x in sl:
            for r in y:
                P.dolu[x].pop(r, None)
            for r in yeni:
                P.dolu[x][r] = True
        P.yer[id(e)] = (s_, yeni, es)
    return uyarilar


TALEP_ONCELIK = True         # True: hocanın "ders günü ve saatinde" talebi yerleşimde yumuşak kural olarak gözetilir


def yumusak_ihlal(P, e, s):
    """ÇAP yumuşak çiftleri: aynı oturumda kaç sınav çifti bu kümelerden birine giriyor"""
    n = 0
    for x in P.slotlar(e, s):
        for o in P.kayit.get(x, []):
            if any(cap_capraz(e, o, k) for k in CAP_YUMUSAK):
                n += 1
    return n


def ardisik_ihlal(P, e, s):
    """Aynı bölüm-yarıyıl grubunun aynı gün 3 ardışık sınavı olur mu (e eklenirse)"""
    g, h = divmod(s, len(SAATLER))
    dolu = {h}
    for k in range(len(SAATLER)):
        for o in P.kayit.get(g * len(SAATLER) + k, []):
            if o["gruplar"] & e["gruplar"]:
                dolu.add(k)
    return sum(1 for a in range(h - 2, h + 1) if all(b in dolu for b in (a, a + 1, a + 2)))


def tercih_ihlali(e, s):
    h = s % len(SAATLER) + SAATLER[0]
    return 1 if any(k in TERCIH_SAAT and TERCIH_SAAT[k] != h for k, _ in e["hocalar"]) else 0


def talep_ihlali(e, s):
    """0: talep yok ya da karşılandı, 1: 'ders günü ve saatinde' istendi ama bu oturum dersin saatine uymuyor"""
    if any(k in TALEP_YOK for k in e["kodlar"]):
        return 0
    if not any((f["gun_saat"] or "").startswith("Evet") for f in e["form"] if f["gun_saat"]):
        return 0
    g, h = divmod(s, len(SAATLER))
    h += SAATLER[0]
    return 0 if any(dg == g and a <= h < b for dg, a, b in e["ders_saat"]) else 1


def ayni_oda_birlestir(P, sec):
    """Aynı oturuma düşen MUAF_KODLAR sınavlarını (aynı hoca) sığıyorsa tek dersliğe toplar ('aynı sınıfta')."""
    kume = collections.defaultdict(list)
    for e in sec:
        if id(e) in P.yer and any(k in MUAF_KODLAR for k in e["kodlar"]):
            kume[P.yer[id(e)][0]].append(e)
    bilgi = []
    for s_, lst in kume.items():
        if len(lst) < 2:
            continue
        toplam = sum(e["n"] for e in lst)
        baska = {r for x in P.slotlar(lst[0], s_) for o in P.kayit.get(x, []) if o not in lst for r in P.yer[id(o)][1]}
        aday = [r for r in (LABLAR if lst[0].get("lab") else DERSLIKLER) if r not in baska and toplam <= min(TAM[r], KAPASITE[r] + ESNEME)]
        if not aday:
            bilgi.append(f"{slot_adi(s_)}: {toplam} kişi tek dersliğe sığmadı")
            continue
        hedef = min(aday, key=lambda r: (KAPASITE[r], r))
        for e in lst:
            s0, y, es = P.yer[id(e)]
            for x in P.slotlar(e, s0):
                for r in y:
                    P.dolu[x].pop(r, None)
            P.yer[id(e)] = (s0, {hedef: [g for l in y.values() for g in l]}, es)
            for x in P.slotlar(e, s0):
                P.dolu[x][hedef] = True
        bilgi.append(f"{slot_adi(s_)}: {', '.join(e['ad'][:22] for e in lst)} -> {hedef} ({toplam} kişi)")
    return bilgi


def yerlestir(sinavlar, P=None):
    P = P or Program()
    sinavlar = [e for e in sinavlar if id(e) not in P.yer]
    sabit = [e for e in sinavlar if any(k in SABIT for k in e["kodlar"])]
    for e in sabit:
        kod = next(k for k in e["kodlar"] if k in SABIT)
        g, h = SABIT[kod]
        s = g * len(SAATLER) + (h - SAATLER[0])
        ok, neden = P.uygun(e, s, sabit=True)
        if not ok:
            raise SystemExit(f"Sabit sınav yerleşemedi: {e['ad']}: {neden}")
        P.koy(e, s, *ok)
    kalan = [e for e in sinavlar if e not in sabit]
    derece = {id(e): sum(1 for o in sinavlar if o is not e and cakisir(e, o)) for e in kalan}
    def _kilitli(e):
        return any(k in KILIT for k in list(e["kodlar"]) + sorted(e["tum"])) and not any(h in PRIORITY_HOCA for h, _ in e["hocalar"])
    kalan.sort(key=lambda e: (not _kilitli(e), not any(h in PRIORITY_HOCA for h, _ in e["hocalar"]),
                              -(e["n"] >= BUYUK), -derece[id(e)], -e["n"]))
    yerlesemeyen = []
    for e in kalan:
        gun_yuk = collections.Counter()
        buyuk_gun = collections.Counter()
        for x, lst in P.kayit.items():
            gun_yuk[x // len(SAATLER)] += len(lst)
            buyuk_gun[x // len(SAATLER)] += sum(1 for o in lst if o["n"] >= BUYUK)
        en = None
        nedenler = collections.Counter()
        kilit = next((KILIT[k] for k in list(e["kodlar"]) + sorted(e["tum"]) if k in KILIT), None)
        if kilit:
            ks = kilit[0] * len(SAATLER) + (kilit[1] - SAATLER[0])
            ok, neden = P.uygun(e, ks)
            if ok:
                en = ((0,), ks, ok)
                P.koy(e, en[1], *en[2][:2])
                continue
        for s in range(len(GUNLER) * len(SAATLER)):
            ok, neden = P.uygun(e, s)
            if not ok:
                nedenler[neden.split(":")[0]] += 1
                continue
            g, h = divmod(s, len(SAATLER))
            puan = ((talep_ihlali(e, s) if TALEP_ONCELIK else 0), tercih_ihlali(e, s), yumusak_ihlal(P, e, s),
                    ardisik_ihlal(P, e, s), ok[1], buyuk_gun[g] if e["n"] >= BUYUK else 0, gun_yuk[g],
                    abs(h + SAATLER[0] - 13), s)
            if en is None or puan < en[0]:
                en = (puan, s, ok)
        if en is None and P.ardisik_sert:                       # 3 ardışık kuralı yüzünden yerleşemedi: kuralı gevşetip dene
            P.ardisik_sert = False
            for s in range(len(GUNLER) * len(SAATLER)):
                ok, neden = P.uygun(e, s)
                if ok:
                    g, h = divmod(s, len(SAATLER))
                    puan = (ardisik_ihlal(P, e, s), s)
                    if en is None or puan < en[0]:
                        en = (puan, s, ok)
            P.ardisik_sert = True
            if en is not None:
                e["gevsek"] = True
        if en is None:
            yerlesemeyen.append((e, nedenler))
        else:
            P.koy(e, en[1], *en[2][:2])
    return P, yerlesemeyen


# ---------------------------------------------------------------- aşama 1 seçimi
_TR = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ"


def tr_anahtar(x):
    x = str(x or "").strip().replace("i", "İ").replace("ı", "I").upper()
    return [(_TR.index(c) if c in _TR else 99 + ord(c)) for c in x]


def ogrenci_oku():
    """kod -> [(no, ad, soyad, program)]  (kişisel veri: yalnızca 'Öğrenci Dağılımı' dosyasına yazılır)"""
    ws = openpyxl.load_workbook(K.KAYIT_DOSYASI, read_only=True).active
    it = ws.iter_rows(values_only=True)
    h = [re.sub(r"_\d+$", "", str(x).strip()) for x in next(it)]
    ik, ino, iad, isy, ip, itp = (h.index(x) for x in ("D.Kodu", "Öğrenci No", "Adı", "Soyadı", "Öğrenci Programı",
                                                         "Ders Alma Tip"))
    out = collections.defaultdict(list)
    for r in it:
        if r[ik] and r[ino]:
            out[str(r[ik]).strip()].append((str(r[ino]).strip(), str(r[iad] or "").strip(), str(r[isy] or "").strip(),
                                            str(r[ip] or "").strip(), str(r[itp] or "").strip()))
    return out


def dagilim_hesapla(P, sec, ogr_tbl):
    """Bölünen bölümlerde hangi öğrencinin hangi dersliğe gireceği: soyad (sonra ad) sırasına göre ardışık bloklar.
    -> {id(sınav): {(derslik, program): [(no, ad, soyad)]}}"""
    out = {}
    for e in sec:
        if id(e) not in P.yer:
            continue
        y = P.yer[id(e)][1]
        ogrler = collections.defaultdict(list)
        for k in e["tum"]:
            for no, ad, sy, pr, tp in ogr_tbl.get(k, []):
                ogrler[pr].append((no, ad, sy, tp))
        d = {}
        for pr, kisiler in ogrler.items():
            parcalar = sorted((r, n) for r, l in y.items() for a, n in l if a == pr)
            kisiler = sorted(set(kisiler), key=lambda x: x[0])      # öğrenci numarası sırası = giriş yılı sırası
            i = 0
            for r, n in parcalar:
                d[(r, pr)] = kisiler[i:i + n]
                i += n
            if parcalar and i < len(kisiler):                     # artan (kayıt/sayı farkı): son parçaya
                d[(parcalar[-1][0], pr)] += kisiler[i:]
        out[id(e)] = d
    return out


def alttan_say(kisiler):
    return sum(1 for k in kisiler if k[3] not in ("Zorunlu", "Seçmeli"))


def aralik(kisiler):
    """Bölünen grubun bu parçası: giriş yılı aralığı ve kaç kişinin alttan/üstten aldığı."""
    if not kisiler:
        return ""
    y1, y2 = "20" + kisiler[0][0][:2], "20" + kisiler[-1][0][:2]
    return f"giriş {y1 if y1 == y2 else y1 + '–' + y2}, {alttan_say(kisiler)} alttan"


def ortaklar(sinavlar):
    sec, haric = [], []
    for e in sinavlar:
        if len(e["bolum"]) < 3 or not e["ogr"]:
            continue
        if any(k in ERTELE for k in e["kodlar"]):
            haric.append((e, "SONA BIRAKILDI: büyük derslik (3. kat) durumuna göre en son yerleştirilecek"))
        elif any(k in OKUL_YAPAR for k in e["kodlar"]) or any(k in OKUL_YAPAR for k in e["tum"]):
            haric.append((e, "okul tarafından yapılıyor olabilir (Turkish Language I), sorulacak"))
        elif not e["var"]:
            haric.append((e, "hoca sınav yapmayacak"))
        else:
            sec.append(e)
    for e in sinavlar:
        if any(k in EKLE for k in e["kodlar"]) and e["ogr"] and e["var"] and e not in sec:
            sec.append(e)
    return sec, haric


# ---------------------------------------------------------------- Excel çıktıları
BASL = PatternFill("solid", fgColor="1F3864")
SINIF_RENK = {1: ("F8B4B4", "C00000"), 2: ("C6E0B4", "375623"), 3: ("BDD7EE", "1F4E79"), 4: ("F4B6D8", "A0155E")}
RENK = ["F4B183", "A9D18E", "9DC3E6", "FFD966", "C9A0DC", "F8CBAD", "B4C7E7", "C5E0B4", "FFE699", "D9D9D9"]


def ust(ws, satir=1):
    for c in ws[satir]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), BASL
        c.alignment = Alignment(wrap_text=True, vertical="center")


def genislik(ws, w):
    for j, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(j)].width = x


def hoca_yaz(e):
    return ", ".join(h for _, h in e["hocalar"]) or "—"


def oda_hoca(e, lst):
    """Dersliğe giren öğrenci gruplarının hocası (aynı kodlu dersi birden çok hoca veriyorsa ayrı ayrı)."""
    anahtarlar = {e["prog_hoca"].get(a, "") for a, _ in lst} - {""}
    ad = [h for k, h in e["hocalar"] if k in anahtarlar]
    return ", ".join(ad) if ad else hoca_yaz(e)


def prog_yaz(e):
    return "; ".join(f"{p} {n}" for p, n in sorted(e["prog"].items(), key=lambda x: -x[1]))


def bolum_hoca_sayfalari(wb, P, sec):
    """Yerleşim.xlsx'e 'Bölüm bazında' ve 'Hoca bazında' çakışma durumu sayfalarını ekler (sınavların kendi öğrenci listesiyle)."""
    ws0 = openpyxl.load_workbook(K.KAYIT_DOSYASI, read_only=True).active
    it = ws0.iter_rows(values_only=True)
    h = [re.sub(r"_\d+$", "", str(x).strip()) for x in next(it)]
    ix = {n: i for i, n in enumerate(h)}
    coh = {}
    for r in it:
        if r[1]:
            coh[r[ix["Öğrenci No"]]] = (str(r[ix["Öğrenci Programı"]]).strip(), r[ix["Öğr.Sınıf"]])
    sl = {id(e): set(P.slotlar(e, P.yer[id(e)][0])) for e in sec}
    ogr_ex = collections.defaultdict(list)
    for e in sec:
        for no in e["ogr"]:
            ogr_ex[no].append(e)
    prog_ogr = collections.defaultdict(set)
    for no, (pr, _) in coh.items():
        prog_ogr[pr].add(no)
    bol = []
    for pr in sorted(prog_ogr):
        ogrs = prog_ogr[pr]
        ex = [e for e in sec if len(e["ogr"] & ogrs) >= 3]
        celisen = ardisik = gun3 = 0
        for no in ogrs:
            occ = collections.defaultdict(set)
            for e in ogr_ex.get(no, []):
                for x in sl[id(e)]:
                    occ[x].add(e["ad"])
            if any(len(v) > 1 for v in occ.values()):
                celisen += 1
            gun = collections.defaultdict(set)
            for x in occ:
                gun[x // len(SAATLER)].add(x % len(SAATLER))
            if max([len(v) for v in gun.values()] or [0]) >= 3:
                gun3 += 1
            for g, st in gun.items():
                if any(all((a + i) in st for i in range(3)) for a in range(len(SAATLER) - 2)):
                    ardisik += 1
        kohort = collections.defaultdict(lambda: collections.defaultdict(set))
        for e in sec:
            for no in e["ogr"]:
                if no in coh and coh[no][0] == pr:
                    kohort[coh[no][1]][e["ad"]].add(no)
        cift, ornek = 0, []
        byad = {e["ad"]: e for e in sec}
        for sn_, exs in kohort.items():
            for a, b in itertools.combinations(exs, 2):
                if sl[id(byad[a])] & sl[id(byad[b])] and len(exs[a]) >= 3 and len(exs[b]) >= 3:
                    cift += 1
                    ornek.append(f"{sn_}. sınıf: {a[:28]} <-> {b[:28]}")
        bol.append([pr, len(ogrs), len(ex), celisen, cift, "; ".join(ornek) or "—", gun3, ardisik])
    ws = wb.create_sheet("Bölüm bazında")
    ws.append(["Bölüm / program", "Öğrenci", "Sınavı olan ders", "Aynı anda iki sınavı olan öğrenci",
               "Aynı bölüm+sınıftan aynı oturumda iki sınav (çift; ortak öğrenci yok)", "Örnekler", "Günde 3+ sınavı olan öğrenci (bilgi)",
               "3 ardışık sınavı olan öğrenci-gün"])
    ust(ws)
    for r in bol:
        ws.append(r)
    genislik(ws, (44, 9, 12, 16, 22, 70, 16, 16))
    ws.freeze_panes = "A2"
    hoca = collections.defaultdict(list)
    for e in sec:
        for k, ad in e["hocalar"]:
            hoca[(k, ad)].append(e)
    hs = wb.create_sheet("Hoca bazında")
    hs.append(["Hoca", "Sınav sayısı", "Sınav günleri", "Aynı anda iki sınavı (sayı)", "Ayrıntı", "İzin gününe düşen sınav", "Ayrıntı",
               "Ders günü/saati talebi karşılanmayan", "Dersler"])
    ust(hs)
    for (k, ad), exs in sorted(hoca.items(), key=lambda x: x[0][1]):
        cak = []
        for a, b in itertools.combinations(exs, 2):
            if sl[id(a)] & sl[id(b)]:
                bil = any(c in MUAF_KODLAR for c in a["kodlar"]) and any(c in MUAF_KODLAR for c in b["kodlar"])
                cak.append(f"{a['ad'][:24]} / {b['ad'][:24]}" + (" (Meysure, bilerek aynı sınıf)" if bil else ""))
        izin_ih = [e["ad"] for e in exs if (P.yer[id(e)][0] // len(SAATLER)) in e["izin"]]
        gunler = sorted({P.yer[id(e)][0] // len(SAATLER) for e in exs})
        talep = [e["ad"][:30] for e in exs
                 if any((f["gun_saat"] or "").startswith("Evet") for f in e["form"] if f["hoca"] == k)
                 and talep_ihlali(e, P.yer[id(e)][0]) == 1 and not any(c in TALEP_YOK for c in e["kodlar"])]
        hs.append([ad, len(exs), ", ".join(GUNLER[g][:3] for g in gunler), len(cak), "; ".join(cak) or "—", len(izin_ih),
                   ", ".join(izin_ih) or "—", len(talep), ", ".join(talep) or "—"])
    genislik(hs, (30, 9, 22, 12, 60, 10, 30, 14, 60))
    hs.freeze_panes = "A2"


def excel_yaz(P, sec, haric, yerlesemeyen, tum_ex, yol, dag=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Açıklama"
    for s in ("Aşama 1: Ortak dersler (3 ve daha fazla bölümün aldığı dersler).",
              "Sayfalar: Liste (ortak dersler), Yerleşim (gün/saat/derslik), Derslik Dosyası (derslik x oturum), "
              "Bölüm Görünümü (bölüm x oturum, renkli), Kontrol.",
              "Varsayımlar: yalnızca 6. kat derslikleri kullanıldı (3. kat dosyası ve derslik müsaitliği henüz yok, "
              "6. katın bütün gün boş olduğu varsayıldı).",
              "Derslik kapasiteleri sınav kapasitesidir; normalde +5, 100 ve üzeri öğrencili derslerde gerekirse +10 esneme.",
              "Bu dosya öğrenci numarası içermez."):
        ws.append([s])
    ws.column_dimensions["A"].width = 150
    # Liste
    ws = wb.create_sheet("Liste")
    ws.append(["Ders kodu", "Ders", "Hoca(lar)", "Öğrenci", "Bölüm sayısı", "Program dağılımı", "Süre (dk)",
               "Hoca izin günü", "Durum", "Not"])
    ust(ws)
    yerli = {id(e) for e in sec if id(e) in P.yer}
    for e in sorted(sec, key=lambda e: -e["n"]):
        ws.append([" / ".join(e["kodlar"]), e["ad"], hoca_yaz(e), e["n"], len(e["bolum"]), prog_yaz(e), e["sure"],
                   ", ".join(GUNLER[g] for g in sorted(e["izin"])) or "—",
                   "Yerleşti" if id(e) in yerli else "YERLEŞEMEDİ",
                   e.get("not") or ("" if len(e["bolum"]) >= 3 else "ortak ders değil: hoca talebiyle (aşama 3'ten) erken yerleştirildi")])
    for e, neden in haric:
        ws.append([" / ".join(e["kodlar"]), e["ad"], hoca_yaz(e), e["n"], len(e["bolum"]), prog_yaz(e), e["sure"],
                   "", "Yerleştirilmedi", neden])
    genislik(ws, (22, 40, 34, 9, 9, 70, 8, 20, 14, 50))
    ws.freeze_panes = "A2"
    # Yerleşim
    ws = wb.create_sheet("Yerleşim")
    ws.append(["Ders kodu", "Ders", "Hoca(lar)", "Gün", "Tarih", "Saat", "Süre (dk)", "Öğrenci", "Derslik ve dağılım",
               "Esneme"])
    ust(ws)
    for e in sorted((e for e in sec if id(e) in P.yer), key=lambda e: P.yer[id(e)][0]):
        s, y, es = P.yer[id(e)]
        g, h = divmod(s, len(SAATLER))
        bol_say = collections.Counter(a for l in y.values() for a, _ in l)

        def etk(r, a, n):
            return f"{a} {n}" + (f" [{aralik(dag[id(e)].get((r, a), []))}]" if dag and bol_say[a] > 1 else "")
        dagilim = "; ".join(f"{r}: " + ", ".join(etk(r, a, n) for a, n in l) + f" ({sum(n for _, n in l)}/{KAPASITE[r]})"
                            for r, l in sorted(y.items()))
        asan = [f"{r} +{sum(n for _, n in l) - KAPASITE[r]}" for r, l in sorted(y.items())
                if sum(n for _, n in l) > KAPASITE[r]]
        ws.append([" / ".join(e["kodlar"]), e["ad"], hoca_yaz(e), GUNLER[g], TARIHLER[g],
                   slot_adi(s, e["sure"]).split(" ", 2)[2], e["sure"], e["n"], dagilim, ", ".join(asan) or "—"])
    genislik(ws, (22, 40, 34, 11, 12, 14, 8, 8, 90, 24))
    ws.freeze_panes = "A2"
    # Derslik dosyası
    ws = wb.create_sheet("Derslik Dosyası")
    ws.append(["Oturum"] + [f"{r} (kap. {c})" for r, c in KAPASITE.items()])
    ust(ws)
    for s in range(len(GUNLER) * len(SAATLER)):
        g, h = divmod(s, len(SAATLER))
        satir = [f"{GUNLER[g]} {TARIHLER[g][:5]}  {h + SAATLER[0]:02d}:00"]
        for r in KAPASITE:
            hucre = ""
            for e in P.kayit.get(s, []):
                if r in P.yer[id(e)][1]:
                    l = P.yer[id(e)][1][r]
                    bs = collections.Counter(a for ll in P.yer[id(e)][1].values() for a, _ in ll)
                    yeni_ = f"{e['ad']} [{oda_hoca(e, l)}]\n" + ", ".join(
                        f"{a} {n}" + (f" [{aralik(dag[id(e)].get((r, a), []))}]" if dag and bs[a] > 1 else "") for a, n in l)
                    hucre = f"{hucre}\n--\n{yeni_}" if hucre else yeni_
            satir.append(hucre)
        ws.append(satir)
        ws.row_dimensions[ws.max_row].height = 45 if any(satir[1:]) else 18
        if yasak_mi(s):
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", fgColor="D9D9D9")
            ws.cell(ws.max_row, 2).value = "Çarşamba 09-12: sınav yok"
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    genislik(ws, (22,) + (30,) * len(KAPASITE))
    ws.freeze_panes = "B2"
    # Bölüm görünümü (sınıfa göre renk; ortak dersler koyu)
    progs = sorted({p for e in tum_ex for p, n in e["prog"].items() if n >= 1})
    ws = wb.create_sheet("Bölüm Görünümü")
    ws.append(["Oturum"] + progs)
    ust(ws)
    for s_ in range(len(GUNLER) * len(SAATLER)):
        g, h = divmod(s_, len(SAATLER))
        ws.append([f"{GUNLER[g]} {TARIHLER[g][:5]}  {h + SAATLER[0]:02d}:00"] + [""] * len(progs))
        satir = ws.max_row
        if yasak_mi(s_):
            for c in ws[satir]:
                c.fill = PatternFill("solid", fgColor="D9D9D9")
        for j, p in enumerate(progs, 2):
            lst = [e for e in P.kayit.get(s_, []) if e["prog"].get(p, 0) >= 3 and P.yer[id(e)][1]]
            c = ws.cell(satir, j)
            if lst:
                c.value = "\n".join(f"{e['ad']} ({e['sinif'].get(p, '?')}. sınıf)" for e in lst)
                if len(lst) > 1:
                    c.fill = PatternFill("solid", fgColor="000000")
                    c.font = Font(bold=True, color="FFFFFF")
                else:
                    e = lst[0]
                    acik, koyu = SINIF_RENK.get(e["sinif"].get(p), ("EEEEEE", "7F7F7F"))
                    ortak = len(e["bolum"]) >= 3
                    c.fill = PatternFill("solid", fgColor=koyu if ortak else acik)
                    c.font = Font(bold=ortak, color="FFFFFF" if ortak else "000000")
                c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.row_dimensions[satir].height = 30 * len(lst) + 6
    genislik(ws, (22,) + (28,) * len(progs))
    ws.freeze_panes = "B2"
    # Renk rehberi
    ws = wb.create_sheet("Renk Rehberi")
    ws.append(["Sınıf", "Açık renk (bölüme özgü ders)", "Koyu renk (ortak ders: 3+ bölüm)"])
    ust(ws)
    for sn, (acik, koyu) in SINIF_RENK.items():
        ws.append([f"{sn}. sınıf", "Örnek ders", "Ortak ders (koyu)"])
        ws.cell(ws.max_row, 2).fill = PatternFill("solid", fgColor=acik)
        ws.cell(ws.max_row, 3).fill = PatternFill("solid", fgColor=koyu)
        ws.cell(ws.max_row, 3).font = Font(bold=True, color="FFFFFF")
    ws.append(["Çarşamba 09-12", "Sınav konmaz", ""])
    ws.cell(ws.max_row, 2).fill = PatternFill("solid", fgColor="D9D9D9")
    ws.append(["Siyah hücre", "Aynı bölümün iki dersi aynı oturumda (çakışma)", ""])
    ws.cell(ws.max_row, 2).fill = PatternFill("solid", fgColor="000000")
    ws.cell(ws.max_row, 2).font = Font(bold=True, color="FFFFFF")
    genislik(ws, (18, 50, 36))
    # Hoca talepleri (aşama 3'te uygulanacak; burada yalnızca durum)
    ws = wb.create_sheet("Hoca Talepleri")
    ws.append(["Ders", "Hoca(lar)", "Hoca talebi", "Dersin normal saati", "Yerleşen", "Talebe uyuyor mu?"])
    ust(ws)
    for e in sorted((e for e in sec if id(e) in P.yer), key=lambda e: e["ad"]):
        s_ = P.yer[id(e)][0]
        g, h = divmod(s_, len(SAATLER))
        h += SAATLER[0]
        istek = {(f["gun_saat"] or "").startswith("Evet") for f in e["form"] if f["gun_saat"]}
        talep = "ders günü ve saatinde" if True in istek else ("fark etmez" if istek else "form boş")
        uy = ""
        if talep.startswith("ders"):
            uy = "Evet" if any(dg == g and a <= h < b for dg, a, b in e["ders_saat"]) else "HAYIR"
        ws.append([e["ad"], hoca_yaz(e), talep,
                   "; ".join(f"{GUNLER[dg]} {a:02d}-{b:02d}" for dg, a, b in e["ders_saat"]) or "—",
                   slot_adi(s_, e["sure"]), uy])
    genislik(ws, (44, 36, 24, 34, 28, 18))
    ws.freeze_panes = "A2"
    # Kontrol
    ws = wb.create_sheet("Kontrol")
    ws.append(["Kontrol", "Sonuç"])
    ust(ws)
    yer = [e for e in sec if id(e) in P.yer]
    ihlal_o = ihlal_h = 0
    for a, b in itertools.combinations(yer, 2):
        sa, sb = set(P.slotlar(a, P.yer[id(a)][0])), set(P.slotlar(b, P.yer[id(b)][0]))
        if sa & sb:
            if a["ogr"] & b["ogr"]:
                ihlal_o += 1
            if {h for h, _ in a["hocalar"]} & {h for h, _ in b["hocalar"]}:
                ihlal_h += 1
    sebep = collections.Counter()
    for a, b in itertools.combinations(yer, 2):
        if set(P.slotlar(a, P.yer[id(a)][0])) & set(P.slotlar(b, P.yer[id(b)][0])):
            sebep[cakisir(a, b) or "-"] += 1
    ard = 0
    gunler_ = collections.defaultdict(set)
    for e in yer:
        for x in P.slotlar(e, P.yer[id(e)][0]):
            for no in e["ogr"]:
                gunler_[(no, x // len(SAATLER))].add(x % len(SAATLER))
    for occ in gunler_.values():
        if any(all((a + i) in occ for i in range(3)) for a in range(len(SAATLER) - 2)):
            ard += 1
    ws.append(["3 ardışık sınavı olan öğrenci-gün sayısı", ard])
    ws.append(["Yerleşen / seçilen ders", f"{len(yer)} / {len(sec)}"])
    for nd in ("ortak öğrenci", "aynı bölüm-yarıyıl", "aynı hoca", "ÇAP çifti"):
        ws.append([f"Aynı oturumda çakışan çift: {nd}", sebep.get(nd, 0)])
    ws.append(["ÇAP yumuşak çiftlerden aynı oturumda olan",
               sum(1 for a, b in itertools.combinations(yer, 2)
                   if set(P.slotlar(a, P.yer[id(a)][0])) & set(P.slotlar(b, P.yer[id(b)][0]))
                   and any(cap_capraz(a, b, k) for k in CAP_YUMUSAK))])
    ws.append(["Aynı oturumda ortak öğrencisi olan sınav çifti", ihlal_o])
    ws.append(["Aynı oturumda aynı hocanın iki sınavı", ihlal_h])
    ws.append(["Çarşamba 09-12'de sınav", sum(1 for e in yer for x in P.slotlar(e, P.yer[id(e)][0]) if yasak_mi(x))])
    ws.append(["Hoca izin gününe düşen sınav", sum(1 for e in yer if P.yer[id(e)][0] // len(SAATLER) in e["izin"])])
    ws.append(["Kapasiteyi aşan derslik sayısı (esneme)",
               sum(1 for e in yer for r, l in P.yer[id(e)][1].items() if sum(n for _, n in l) > KAPASITE[r])])
    for e, nedenler in yerlesemeyen:
        ws.append([f"YERLEŞEMEDİ: {e['ad']}", "; ".join(f"{k}: {v}" for k, v in nedenler.most_common(4))])
    genislik(ws, (60, 70))
    bolum_hoca_sayfalari(wb, P, yer)
    os.makedirs(os.path.dirname(yol), exist_ok=True)
    try:
        wb.save(yol)
    except PermissionError:
        yol = yol.replace(".xlsx", " (yeni).xlsx")
        wb.save(yol)
        print("UYARI: dosya açık olduğu için yeni adla kaydedildi.")
    return yol


KISA = {"Siyaset Bilimi ve Uluslararası İlişkiler": "SBUİ", "Siyaset Bilimi ve Uluslararası İlişkiler (İngilizce)": "PSIR",
        "İşletme": "İŞL", "İşletme (İngilizce)": "BUS", "Uluslararası Ticaret ve Finansman": "UTF",
        "Uluslararası Ticaret ve Finans": "UTF", "Lojistik Yönetimi": "LOJ", "Havacılık Yönetimi": "HAV",
        "Yönetim Bilişim Sistemleri": "YBS", "Ekonomi ve Finans (İngilizce)": "ECF",
        "Elektronik Ticaret ve Yönetimi": "ETİC"}
BOS_DERSLIK = os.path.join(K.KOK, "Belgeler", "Derslik Dosyası 19.10.2025.xlsx")


def derslik_formu(P, sec, yol, dag=None):
    """Üniversitenin Derslik Dosyası biçimi: yalnızca 6.Kat sekmesi; hücrelere hoca / ders (bölünen sınavda bölüm etiketi)."""
    wb = openpyxl.load_workbook(BOS_DERSLIK)
    for ad in list(wb.sheetnames):
        if ad != "6.Kat":
            del wb[ad]
    ws = wb["6.Kat"]
    gun_ad = {"PAZARTESİ": 0, "SALI": 1, "ÇARŞAMBA": 2, "PERŞEMBE": 3, "CUMA": 4}
    harita = {}
    for r in range(1, ws.max_row + 1):
        m = re.match(r"\s*(6[AD]\d)\s*-", str(ws.cell(r, 2).value or ""))
        if not m:
            continue
        gs = {gun_ad[str(ws.cell(r + 1, k).value).strip().upper()]: k for k in range(3, 13)
              if str(ws.cell(r + 1, k).value or "").strip().upper() in gun_ad}
        for rr in range(r + 3, r + 16):
            t = re.match(r"(\d{2}):00-", str(ws.cell(rr, 2).value or ""))
            if t:
                for g, k in gs.items():
                    harita[(m.group(1), g, int(t.group(1)))] = (rr, k)
    eksik = 0
    for e in sec:
        if id(e) not in P.yer:
            continue
        s_, y, es = P.yer[id(e)]
        g, h = divmod(s_, len(SAATLER))
        for oda, lst in y.items():
            etiket = e["ad"]
            bs = collections.Counter(a for ll in y.values() for a, _ in ll)
            etiket += " (" + ", ".join(f"{KISA.get(a, a)} {n}" + (f" [{aralik(dag[id(e)].get((oda, a), []))}]" if dag and bs[a] > 1 else "")
                                       for a, n in lst) + ")"
            for hh in range(h + SAATLER[0], h + SAATLER[0] + (2 if e["sure"] > 50 else 1)):
                if (oda, g, hh) not in harita:
                    eksik += 1
                    continue
                rr, k = harita[(oda, g, hh)]
                for kk, v in ((k, oda_hoca(e, lst)), (k + 1, etiket)):
                    c = ws.cell(rr, kk)
                    c.value = f"{c.value} / {v}" if c.value else v
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["N3"].value = "Kısaltmalar: SBUİ, PSIR (İng.), İŞL, BUS (İng.), UTF, LOJ, HAV, YBS, ECF (İng.), ETİC; sayı = o dersliğe giren öğrenci."
    ws["N4"].value = "Dersliklerin yanındaki sayılar tam kapasitedir; sınav kapasitesi: 6A1:35, 6A2:33, 6D1:17, 6D3:23, 6D4:25, 6D5:23 (+5 esneme)."
    os.makedirs(os.path.dirname(yol), exist_ok=True)
    try:
        wb.save(yol)
    except PermissionError:
        yol = yol.replace(".xlsx", " (yeni).xlsx")
        wb.save(yol)
    return yol, eksik


def dagilim_dosyasi(P, sec, dag, yol):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Açıklama"
    for t in ("Bu dosya ÖĞRENCİ NUMARASI ve ADI içerir: paylaşmayın (Çıktılar/ klasörü GitHub'a gitmez).",
              "Bir bölümün öğrencileri birden fazla dersliğe bölündüğünde ÖĞRENCİ NUMARASI sırasına göre ardışık bloklara "
              "ayrılır (numara = giriş yılı sırası: eski girişliler bir derslikte, yeniler diğerinde). 'Bölünen Gruplar' "
              "sayfası her dersliğe giren giriş yılı aralığını ve alttan/üstten alanların sayısını gösterir.",
              "Bölünmeyen bölümlerin öğrencileri 'Öğrenci Listesi'nde tek dersliktedir."):
        ws.append([t])
    ws.column_dimensions["A"].width = 150
    ws = wb.create_sheet("Bölünen Gruplar")
    ws.append(["Ders", "Gün", "Saat", "Bölüm", "Derslik", "Kişi", "Giriş yılı", "Alttan/üstten", "Öğrenci no aralığı"])
    ust(ws)
    ls = wb.create_sheet("Öğrenci Listesi")
    ls.append(["Ders", "Gün", "Saat", "Derslik", "Bölüm", "Öğrenci No", "Ad", "Soyad", "Ders alma tipi"])
    ust(ls)
    for e in sorted((e for e in sec if id(e) in P.yer), key=lambda e: P.yer[id(e)][0]):
        s_, y, es = P.yer[id(e)]
        g, h = divmod(s_, len(SAATLER))
        saat = slot_adi(s_, e["sure"]).split(" ", 2)[2]
        bs = collections.Counter(a for l in y.values() for a, _ in l)
        for (oda, pr), kisiler in sorted(dag[id(e)].items()):
            if bs[pr] > 1:
                nolar = sorted(k[0] for k in kisiler)
                y1, y2 = "20" + nolar[0][:2], "20" + nolar[-1][:2]
                ws.append([e["ad"], f"{GUNLER[g]} {TARIHLER[g]}", saat, pr, oda, len(kisiler),
                           y1 if y1 == y2 else f"{y1}–{y2}", alttan_say(kisiler),
                           f"{nolar[0]} – {nolar[-1]}" if nolar else ""])
            for no, ad, sy, tp in kisiler:
                ls.append([e["ad"], f"{GUNLER[g]} {TARIHLER[g]}", saat, oda, pr, no, ad, sy, tp])
    genislik(ws, (40, 22, 14, 44, 8, 6, 12, 14, 28))
    genislik(ls, (40, 22, 14, 8, 44, 14, 20, 20, 16))
    ws.freeze_panes = ls.freeze_panes = "A2"
    ls.auto_filter.ref = ls.dimensions
    os.makedirs(os.path.dirname(yol), exist_ok=True)
    try:
        wb.save(yol)
    except PermissionError:
        yol = yol.replace(".xlsx", " (yeni).xlsx")
        wb.save(yol)
    return yol


# ---------------------------------------------------------------- aşama 2: alttan/üstten ve ÇAP listesi
ELLE_BOLUM = {"20EG4009": {"Ekonomi ve Finans"}, "2023G2006": {"Elektronik Ticaret ve Yönetimi"},
              "20G1009": {"Elektronik Ticaret ve Yönetimi"}}
OZEL_TIP = ("Zorunlu", "Seçmeli")                       # bunların dışındakiler: alttan, devamlı alttan, üstten, yükseltme
CAP_YUMUSAK = [{"İşletme", "Yönetim Bilişim Sistemleri"}, {"Siyaset Bilimi ve Uluslararası İlişkiler", "Havacılık Yönetimi"}]


def asama2_liste(sinavlar, P, sec, haric, yol):
    ogr_tbl = ogrenci_oku()
    alttan = {}                                         # id(e) -> {öğrenci no: program}
    for e in sinavlar:
        d = {}
        for k in e["tum"]:
            for no, ad, sy, pr, tp in ogr_tbl.get(k, []):
                if tp not in OZEL_TIP:
                    d[no] = pr
        alttan[id(e)] = d
    ertele = {id(e) for e, n in haric if n.startswith("SONA")}
    okul = {id(e) for e, n in haric if n.startswith("okul")}
    yerli = {id(e): slot_adi(P.yer[id(e)][0], e["sure"]) for e in sec if id(e) in P.yer}

    def durum(e):
        if id(e) in yerli:
            return "Aşama 1'de yerleşti: " + yerli[id(e)]
        if id(e) in ertele:
            return "Sona bırakıldı (3. kat)"
        if id(e) in okul:
            return "Okul yapıyor olabilir"
        if not e["var"]:
            return "Sınav yapılmayacak"
        return "Bekliyor"

    tum = [e for e in sinavlar if e["ogr"]]
    for e in tum:
        if not e["bolum"]:
            e["bolum"] = {b for k in e["kodlar"] for b in ELLE_BOLUM.get(k, set())}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Açıklama"
    for t in ("Aşama 2: Alttan/üstten alan öğrencilerin olduğu dersler ve ÇAP çiftlerinin dersleri (YALNIZCA LİSTE; yerleştirme onayınızdan sonra).",
              "Alttan/üstten = ders alma tipi Alttan, Devamlı Alttan, Üstten veya Yükseltme. Bu öğrenciler başka sınıfların derslerini de aldığı için "
              "normal programda çakışmayan dersler çakışabilir; çakışma kuralı (ortak öğrencisi olan dersler aynı oturumda olmaz) bunları zaten kapsar.",
              "ÇAP öğrencilerinin kimliği bilinmiyor: Havacılık+Lojistik ve ETİC+Havacılık çiftlerinin tüm dersleri çakışmaz (sert); "
              "İşletme+YBS ve Siyaset+Havacılık mümkünse çakışmaz (yumuşak).",
              "Bu dosya öğrenci numarası içermez."):
        ws.append([t])
    ws.column_dimensions["A"].width = 160

    # --- alttan/üstten dersler
    ws = wb.create_sheet("Alttan-Üstten Dersler")
    ws.append(["Ders kodu", "Ders", "Hoca(lar)", "Öğrenci", "Alttan/üstten", "Oran %", "Alttan/üstten öğrencilerin bölümü", "Durum"])
    ust(ws)
    satir = []
    for e in tum:
        a = alttan[id(e)]
        if not a:
            continue
        pc = collections.Counter(a.values())
        satir.append((len(a), [" / ".join(e["kodlar"]), e["ad"], hoca_yaz(e), e["n"], len(a), round(100 * len(a) / e["n"]),
                               "; ".join(f"{KISA.get(p, p)} {n}" for p, n in pc.most_common()), durum(e)]))
    for _, r in sorted(satir, key=lambda x: -x[0]):
        ws.append(r)
    genislik(ws, (22, 44, 36, 9, 12, 8, 50, 44))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # --- yalnızca alttan/üstten nedeniyle çakışan ders çiftleri
    ws = wb.create_sheet("Alttan Çakışmaları")
    ws.append(["Ders A", "Ders B", "Ortak öğrenci", "Bunlardan alttan/üstten", "Durum A", "Durum B", "Aşama 1'de aynı oturumda mı?"])
    ust(ws)
    toplam = yalniz = 0
    for a_, b_ in itertools.combinations(tum, 2):
        if a_["tum"] & b_["tum"]:
            continue
        ortak = a_["ogr"] & b_["ogr"]
        if not ortak:
            continue
        toplam += 1
        al = {n for n in ortak if n in alttan[id(a_)] or n in alttan[id(b_)]}
        if len(al) == len(ortak):                      # alttan olmasa çakışmayacaktı
            yalniz += 1
            ayni = ""
            if id(a_) in P.yer and id(b_) in P.yer:
                sa, sb = set(P.slotlar(a_, P.yer[id(a_)][0])), set(P.slotlar(b_, P.yer[id(b_)][0]))
                ayni = "EVET (ihlal)" if sa & sb else "hayır"
            ws.append([a_["ad"], b_["ad"], len(ortak), len(al), durum(a_), durum(b_), ayni])
    genislik(ws, (44, 44, 10, 12, 44, 44, 18))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # --- ÇAP çiftleri
    ws = wb.create_sheet("ÇAP Dersleri")
    ws.append(["ÇAP çifti", "Kural", "Ders kodu", "Ders", "Bölümler (≥3 öğrenci)", "Öğrenci", "Durum"])
    ust(ws)
    ozet = []
    for kume, kural in [(c, "SERT") for c in CAP_SERT] + [(c, "mümkünse") for c in CAP_YUMUSAK]:
        ad = " + ".join(sorted(KISA.get(b, b) for b in kume))
        dersler = [e for e in tum if e["bolum"] & kume and e["var"] and id(e) not in okul]
        for e in sorted(dersler, key=lambda e: e["ad"]):
            ws.append([ad, kural, " / ".join(e["kodlar"]), e["ad"], ", ".join(sorted(KISA.get(b, b) for b in e["bolum"])),
                       e["n"], durum(e)])
        ihlal = 0
        yl = [e for e in dersler if id(e) in P.yer]
        for a_, b_ in itertools.combinations(yl, 2):
            if set(P.slotlar(a_, P.yer[id(a_)][0])) & set(P.slotlar(b_, P.yer[id(b_)][0])):
                ihlal += 1
        ozet.append((ad, kural, len(dersler), len(yl), ihlal))
    genislik(ws, (28, 10, 22, 46, 40, 9, 50))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    ws = wb.create_sheet("Özet")
    ws.append(["Konu", "Değer"])
    ust(ws)
    ws.append(["Alttan/üstten alan öğrencisi olan ders sayısı", len(satir)])
    ws.append(["Toplam alttan/üstten kaydı (ders-öğrenci)", sum(len(alttan[id(e)]) for e in tum)])
    ws.append(["Ortak öğrencisi olan ders çifti (toplam)", toplam])
    ws.append(["  bunlardan yalnızca alttan/üstten nedeniyle çakışanlar", yalniz])
    for ad, kural, n, y, ih in ozet:
        ws.append([f"ÇAP {ad} ({kural}): ders sayısı / aşama 1'de yerleşen / aynı oturum ihlali", f"{n} / {y} / {ih}"])
    ws.append(["Kullanılabilir oturum (Çarşamba 09-12 hariç)", len(GUNLER) * len(SAATLER) - len(YASAK)])
    genislik(ws, (92, 22))
    os.makedirs(os.path.dirname(yol), exist_ok=True)
    try:
        wb.save(yol)
    except PermissionError:
        yol = yol.replace(".xlsx", " (yeni).xlsx")
        wb.save(yol)
    return yol, len(satir), toplam, yalniz, ozet


ABBR = {"SIY": "SBUİ", "SBUİ": "SBUİ", "PSIR": "PSIR", "İŞL": "İŞL", "ISL": "İŞL", "BUS": "BUS", "UTF": "UTF", "LOJ": "LOJ", "HAV": "HAV",
        "YBS": "YBS", "ECF": "ECF", "ETİC": "ETİC", "ETIC": "ETİC"}


# okul dosyasında (3./4./11. kat) yapılan ve 6.Kat sayfasında olmayan sınavlar: ad -> (gün, saat, süre_dk, {derslik: [(program, sayı)]})
EKSTRA = {
    "İktisada Giriş I": (4, 15, 50, {"3A1": [("Siyaset Bilimi ve Uluslararası İlişkiler", 28)],
                                      "3A3": [("Uluslararası Ticaret ve Finansman", 28), ("Lojistik Yönetimi", 44)],
                                      "3A5": [("İşletme", 46), ("Havacılık Yönetimi", 44)]}),
    "Sosyal Bilimler için Matematik I": (4, 16, 120, {"3A6": [("İşletme", 52)],
                                                      "4A3": [("Uluslararası Ticaret ve Finansman", 29), ("Havacılık Yönetimi", 48)],
                                                      "4A4": [("Lojistik Yönetimi", 43), ("Elektronik Ticaret ve Yönetimi", 16)]}),
    # 29 Ekim: Sosyoloji bölümü istedi; Cuma 10:00-11:00, kendi öğrencileriyle ortak derslikler
    "Sosyolojide Temel Kavramlar": (4, 10, 60, {"3A5-3A6-3D6-3D7": [("Siyaset Bilimi ve Uluslararası İlişkiler", 19), ("Havacılık Yönetimi", 1)]}),
    "Algoritma ve Programlama 1": (0, 12, 90, {"4L1": [("Yönetim Bilişim Sistemleri", 38)]}),
    "Veri Yapıları": (1, 9, 90, {"11L8": [("Yönetim Bilişim Sistemleri", 35)]}),
    "Web Programlama": (4, 14, 90, {"4D8": [("Yönetim Bilişim Sistemleri", 22)]}),
    "Python Programlama": (3, 16, 90, {"4D8": [("Yönetim Bilişim Sistemleri", 22), ("Elektronik Ticaret ve Yönetimi", 4),
                                               ("Siyaset Bilimi ve Uluslararası İlişkiler", 1)]}),
    "Excel VBA Programlama": (0, 15, 90, {"11L5": [("Yönetim Bilişim Sistemleri", 30)]}),
}


# 6.Kat dosyasında bir saatlik hücre olarak yazılan ama uzun süren sınavlar: ad -> dakika (30 Ekim: İstatistik 2 saat, Cuma 12:00-14:00)
SURE_DK = {"İstatistik": 120}


def ekstra_yukle(sinavlar, yuk):
    """EKSTRA sınavlarını (6.Kat dosyasında olmayanlar) dosyadan_yukle çıktısı gibi {id(e): (slot, {derslik: [(program, n)]}, 0)} üretir.
    90/120 dk sınavlar için e['sure'] güncellenir (iki oturum). Çakışma kontrolü ve bölüm programı bunları da görsün diye."""
    sonuc = {}
    for e in sinavlar:
        if id(e) in yuk and e["ad"] in SURE_DK:
            e["sure"] = SURE_DK[e["ad"]]
    for e in sinavlar:
        if id(e) in yuk:
            continue
        ad = e["ad"].replace("Dept Elective: ", "").strip()
        for k, (g, h, dk, o) in EKSTRA.items():
            if ad.startswith(k):
                if dk > 60:
                    e["sure"] = dk
                sonuc[id(e)] = (g * len(SAATLER) + (h - SAATLER[0]), {r: list(l) for r, l in o.items()}, 0)
    return sonuc


def dosyadan_yukle(yol, sinavlar):
    """Kullanıcının (elle düzenlediği) Derslik Dosyası 6.Kat sekmesini okuyup sınav -> (slot, {derslik: [(grup, n)]}) üretir.
    Her sınav kendi grup sayılarına göre doğrulanır; tutmayanlar uyarı olarak döner (ve yüklenmez)."""
    ws = openpyxl.load_workbook(yol)["6.Kat"]
    gun = {"C": 0, "E": 1, "G": 2, "I": 3, "K": 4}
    adlar = sorted({e["ad"] for e in sinavlar}, key=len, reverse=True)
    ad_e = {e["ad"]: e for e in sinavlar}
    hucre = collections.defaultdict(list)                   # (derslik, gün, saat) -> [(sınav adı, metin)]
    oda = None
    for r in range(1, ws.max_row + 1):
        m = re.match(r"\s*(6[AD]\d)\s*-", str(ws.cell(r, 2).value or ""))
        if m:
            oda = m.group(1)
        sa = str(ws.cell(r, 2).value or "")
        t = re.match(r"(\d\d):00-", sa)
        if not t:
            continue
        for col, g in gun.items():
            metin = ws[f"{chr(ord(col) + 1)}{r}"].value
            if not metin:
                continue
            metin = str(metin)
            konum = []
            for ad in adlar:
                for mm in re.finditer(re.escape(ad) + r"[ (\d]", metin):      # kullanıcı bazen '(' yazmayı unutuyor, bazen birleşik toplamı adın yanına yazıyor
                    konum.append((mm.start(), ad))
            if not konum:
                # kullanıcı yeni eklediği dersi grup sayısı olmadan, kısa adla yazmış ("Küresel Lojistik", "Uluslararası Finansman:")
                kucuk = lambda x: x.replace("İ", "i").replace("I", "i").replace("ı", "i").lower()      # İ, I, ı, i aynı sayılır
                t_ = kucuk(re.sub(r"[:\s]+$", "", metin))
                aday = [e for e in sinavlar if len(t_) >= 8 and t_ in kucuk(e["ad"])]
                if len(aday) == 1:
                    konum.append((0, aday[0]["ad"]))
                    metin = aday[0]["ad"] + " "
            konum.sort(key=lambda x: (x[0], -len(x[1])))
            ele = []
            for i, (st, ad) in enumerate(konum):
                if any(st < s2 + len(a2) and s2 <= st for s2, a2 in ele):    # başka bir adın içinde kalanları ele
                    continue
                ele.append((st, ad))
            for i, (st, ad) in enumerate(ele):
                son = ele[i + 1][0] if i + 1 < len(ele) else len(metin)
                hucre[(oda, g, int(t.group(1)))].append((ad, metin[st + len(ad):son]))
    by = collections.defaultdict(lambda: {"saat": set(), "oda": collections.defaultdict(dict)})
    for (o, g, h), lst in hucre.items():
        for ad, metin in lst:
            by[ad]["saat"].add((g, h))
            by[ad]["oda"][o][(g, h)] = metin
    yuk, uyari = {}, []
    for ad, d in by.items():
        e = ad_e[ad]
        g0, h0 = min(d["saat"])
        s0 = g0 * len(SAATLER) + (h0 - SAATLER[0])
        harita = {KISA.get(k, k): k for k in e["prog"]}
        odalar, toplam = {}, collections.Counter()
        bos_odalar = []
        sayisiz = {}
        for o, gh in d["oda"].items():
            metin = gh.get((g0, h0)) or next(iter(gh.values()))
            metin = re.sub(r"\[[^\]]*\]", "", metin)
            grup = []
            if ad in SAYI_ONEMSIZ:
                isim = set()
                for tok in re.findall(r"[A-Za-zİŞÜÇĞÖıçğöşü]+", metin):
                    k = ABBR.get(tok.upper().replace("İ", "I").replace("Ü", "U"), ABBR.get(tok.upper()))
                    if k and harita.get(k):
                        isim.add(harita[k])
                sayisiz[o] = isim
                continue
            if not re.search(r"\d", metin):
                # grup sayısı yazılmamış: dersin hepsi bu derslikte sayılır (tek derslikli sınav)
                if len(d["oda"]) > 1:
                    bos_odalar.append(o)                               # sayısı yazılmayan derslikler: kalan öğrenciler bunlara eşit bölünür
                    continue
                odalar[o] = list(e["prog"].items()) or [("Öğrenci", e["n"])]
                toplam.update(e["prog"])
                continue
            if re.match(r"\s*\d+\s*\(", metin) and len(set(re.findall(r"([A-Za-zİŞÜÇĞÖıçğöşü]+)\s*\d+", metin))) == 1:
                # kullanıcı iki derslikteki aynı bölüm gruplarını birleştirip toplamı adın yanına yazmış ("Giriş38 (YBS 19 ...)");
                # parantez eski tek-grup bilgisidir. Toplam listedekinden farklıysa listedeki sayı esas alınır.
                ab = re.findall(r"([A-Za-zİŞÜÇĞÖıçğöşü]+)\s*\d+", metin)[0]
                k = ABBR.get(ab.upper(), ab.upper())
                prog = harita.get(k)
                yazili = int(re.match(r"\s*(\d+)", metin).group(1))
                if prog is not None and e["prog"].get(prog) not in (None, yazili):
                    uyari.append(f"NOT {ad}: {o} hücresinde birleşik toplam {yazili} yazılmış, listede {e['prog'][prog]}; listedeki esas alındı")
                    yazili = e["prog"][prog]
                if prog is not None:
                    odalar[o] = [(prog, yazili)]
                    toplam[prog] += yazili
                    continue
            for ab, n in re.findall(r"([A-Za-zİŞÜÇĞÖıçğöşü]+)\s*(\d+)", metin):
                k = ABBR.get(ab.upper().replace("İ", "İ"), ab.upper())
                if k == "ÖĞRENCI" or ab.upper() == "ÖĞRENCI":
                    k = "Öğrenci"
                prog = harita.get(k, "Öğrenci" if k == "Öğrenci" else None)
                if prog is None:
                    uyari.append(f"{ad}: {o} hücresinde tanınmayan grup '{ab} {n}'")
                    continue
                grup.append((prog, int(n)))
                toplam[prog] += int(n)
            odalar[o] = grup
        if bos_odalar:
            kalan = [(pr, n - toplam[pr]) for pr, n in e["prog"].items() if n - toplam[pr] > 0]
            k_ = len(bos_odalar)
            for i, o in enumerate(bos_odalar):                         # eşit parça: artan ilk derslik(ler)e +1
                odalar[o] = [(pr, n // k_ + (1 if i < n % k_ else 0)) for pr, n in kalan if n // k_ + (1 if i < n % k_ else 0) > 0]
            for pr, n in kalan:
                toplam[pr] += n
            uyari.append(f"NOT {ad}: {', '.join(bos_odalar)} hücrelerinde sayı yok; kalan öğrenciler eşit bölündü: {[(o, odalar[o]) for o in bos_odalar]}")
        if sayisiz:
            for pr, n in e["prog"].items():
                if pr in DERSI_ALMAYAN.get(ad, ()):
                    continue
                rl = [o for o, s_ in sayisiz.items() if pr in s_]
                for i, o in enumerate(rl):
                    odalar.setdefault(o, []).append((pr, n // len(rl) + (1 if i < n % len(rl) else 0)))
                toplam[pr] = n
            uyari.append(f"NOT {ad}: hücredeki sayılar yok sayıldı, bölümler derslikler arasında eşit paylaştırıldı: {odalar}")
        beklenen = collections.Counter(e["prog"]) if e["prog"] else None
        ok = True
        if e["prog"]:
            for pr, n in e["prog"].items():
                if pr in DERSI_ALMAYAN.get(ad, ()):
                    continue
                if toplam[pr] != n:
                    ok = False
                    uyari.append(f"{ad}: {KISA.get(pr, pr)} dosyada {toplam[pr]}, listede {n}")
        if ok:
            es = max([0] + [sum(n for _, n in l) - KAPASITE[o] for o, l in odalar.items() if o in KAPASITE])
            yuk[id(e)] = (s0, odalar, es)
    return yuk, uyari


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--asama", type=int, default=1)
    ap.add_argument("--onayli", action="store_true",
                    help="Programı çözmek yerine kullanıcının Derslik Dosyası'ndan yükle (elle değişiklikler korunur)")
    ap.add_argument("--dosya", default=None, help="--onayli için kaynak Derslik Dosyası (varsayılan: Aşama 2 - Derslik Dosyası (6. Kat).xlsx)")
    a = ap.parse_args()
    sinavlar, izin, kt = sinavlari_kur()
    sec, haric = ortaklar(sinavlar)
    if a.onayli:
        kaynak = a.dosya or os.path.join(CIKTI, "Aşama 2 - Derslik Dosyası (6. Kat).xlsx")
        yuk, uyari = dosyadan_yukle(kaynak, sinavlar)
        P = Program()
        for e in sinavlar:
            if id(e) in yuk:
                P.koy(e, *yuk[id(e)])
        yerli = [e for e in sinavlar if id(e) in P.yer]
        print(f"{kaynak}: {len(yerli)} sınav yüklendi, {len(uyari)} uyarı")
        for u in uyari:
            print("  UYARI:", u)
        yerli_id = {id(e) for e in yerli}
        haric2 = [(e, n) for e, n in haric if id(e) not in yerli_id]
        for e in sinavlar:
            if e.get("lab") and e["var"] and id(e) not in yerli_id and all(id(e) != id(x) for x, _ in haric2):
                haric2.append((e, "laboratuvar sınavı: kullanıcı yerleştirecek (lab bilgisi sonra)"))
        ogr_tbl = ogrenci_oku()
        dag = dagilim_hesapla(P, yerli, ogr_tbl)
        yol = excel_yaz(P, yerli, haric2, [], sinavlar, os.path.join(CIKTI, "Aşama 2 - Yerleşim.xlsx"), dag)
        print("->", yol)
        y2, eksik = derslik_formu(P, yerli, os.path.join(CIKTI, "Aşama 2 - Derslik Dosyası (6. Kat) (kod çıktısı).xlsx"), dag)
        print("->", y2, "| haritada olmayan:", eksik, "(kullanıcının dosyası değiştirilmedi)")
        print("->", dagilim_dosyasi(P, yerli, dag, os.path.join(CIKTI, "Aşama 2 - Öğrenci Dağılımı (PAYLAŞMAYIN).xlsx")))
        return
    print(f"{len(sinavlar)} sınav birimi, {len(sec)} ortak ders (+{len(haric)} yerleştirilmeyen)")
    # sona bırakılan dersler de önce yerleşir (diğerlerinin yerleri değişmesin), sonra çıkarılır
    ertele = [e for e, n in haric if n.startswith("SONA BIRAKILDI")]
    P, yok = yerlestir(sec + ertele)
    gecici_s = {id(e): P.cikar(e) for e in ertele if id(e) in P.yer}
    gecici = {id(e): slot_adi(gecici_s[id(e)], e["sure"]) for e in ertele if id(e) in gecici_s}
    for i, (e, n) in enumerate(haric):
        if id(e) in gecici:
            haric[i] = (e, n + f" (geçici yer: {gecici[id(e)]}; bu yer boşaltıldı)")
    for u in elle_duzelt(P, sec):
        print("UYARI:", u)
    if a.asama == 2:
        liste = os.path.join(CIKTI, "Aşama 2 - Alttan ve ÇAP Listesi.xlsx")
        if "--liste" in sys.argv:
            asama2_liste(sinavlar, P, sec, haric, liste)
            return
        # sona bırakılanlar: yalnızca çakışma için yer tutar (derslik tutmaz)
        for e in ertele:
            if id(e) in gecici_s:
                P.koy(e, gecici_s[id(e)], {}, 0)
        ogr_tbl = ogrenci_oku()
        okul = {id(e) for e, n in haric if n.startswith("okul")}
        yerli_ids = {id(e) for e in sec}
        aday = []
        for e in sinavlar:
            if id(e) in yerli_ids or id(e) in P.yer or id(e) in okul or not e["var"]:
                continue
            if not (e["ogr"] or any(k in ELLE_SAYI for k in e["kodlar"])):
                continue
            alt = any(tp not in OZEL_TIP for k in e["tum"] for _, _, _, _, tp in ogr_tbl.get(k, []))
            if e.get("lab"):                                  # laboratuvar sınavları: kullanıcı kendisi yerleştirecek
                haric.append((e, "laboratuvar sınavı: kullanıcı yerleştirecek (lab bilgisi sonra)"))
                continue
            if alt or e["bolum"] & CAP_BOLUMLER:
                e["not"] = "Aşama 2: alttan/üstten alan öğrencisi var" if alt else "Aşama 2: ÇAP çifti bölümü"
                aday.append(e)
        print(f"aşama 2 aday: {len(aday)} sınav")
        P, yok2 = yerlestir(aday, P)
        for e in ertele:
            if id(e) in P.yer and not P.yer[id(e)][1]:
                P.cikar(e)
        sec2 = sec + [e for e in aday if id(e) in P.yer]
        for u in elle_duzelt(P, sec2):
            print("UYARI:", u)
        for b_ in ayni_oda_birlestir(P, sec2):
            print("Aynı sınıf:", b_)
        dag = dagilim_hesapla(P, sec2, ogr_tbl)
        yol = excel_yaz(P, sec2, haric, yok2, sinavlar, os.path.join(CIKTI, "Aşama 2 - Yerleşim.xlsx"), dag)
        print("->", yol)
        y2, eksik = derslik_formu(P, sec2, os.path.join(CIKTI, "Aşama 2 - Derslik Dosyası (6. Kat).xlsx"), dag)
        print("->", y2, "| haritada olmayan:", eksik)
        print("->", dagilim_dosyasi(P, sec2, dag, os.path.join(CIKTI, "Aşama 2 - Öğrenci Dağılımı (PAYLAŞMAYIN).xlsx")))
        print(f"yerleşen aşama 2: {sum(1 for e in aday if id(e) in P.yer)} / {len(aday)}")
        for e, n in yok2:
            print("  YERLEŞEMEDİ:", e["ad"][:40], e["n"], dict(n.most_common(4)))
        return
    dag = dagilim_hesapla(P, sec, ogrenci_oku())
    yol = excel_yaz(P, sec, haric, yok, sinavlar, os.path.join(CIKTI, "Aşama 1 - Ortak Dersler.xlsx"), dag)
    print("->", yol)
    y2, eksik = derslik_formu(P, sec, os.path.join(CIKTI, "Aşama 1 - Derslik Dosyası (6. Kat).xlsx"), dag)
    print("->", y2, "| haritada olmayan:", eksik)
    print("->", dagilim_dosyasi(P, sec, dag, os.path.join(CIKTI, "Aşama 1 - Öğrenci Dağılımı (PAYLAŞMAYIN).xlsx")))
    for e in sorted((e for e in sec if id(e) in P.yer), key=lambda e: P.yer[id(e)][0]):
        s, y, es = P.yer[id(e)]
        print(f"  {slot_adi(s, e['sure']):32s} {e['n']:4d}  {e['ad'][:38]:38s} {sorted(y)}  esneme +{es}")
    for e, n in yok:
        print("  YERLEŞEMEDİ:", e["ad"], dict(n))


if __name__ == "__main__":
    main()
