# -*- coding: utf-8 -*-
"""
Girdi okuyucular ve birleştirme (0_hoca_formu.py kullanır).

İLKE: Hangi derslerin sınavı yapılacağını ve dersin hocasını **Ders Programı PDF'i** belirler.
OBS verisi bazen eski olabilir; bu yüzden yalnızca PDF'teki derslere öğrenci sayısı katmak ve PDF ile
farkları raporlamak için kullanılır. OBS'te olup PDF'te olmayan dersler forma EKLENMEZ, kontrol raporuna yazılır.

Öğrenci numaraları yalnızca bellekte sayım için okunur, hiçbir çıktıya yazılmaz.
"""
import os, re, collections
import openpyxl

KOK = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SUBE_DOSYASI = os.path.join(KOK, "Kullanılacak Belgeler", "Öğrenci sayıları.xlsx")
KAYIT_DOSYASI = os.path.join(KOK, "Kullanılacak Belgeler", "altan ve üsten .xlsx")
PROGRAM_PDF = os.path.join(KOK, "Kullanılacak Belgeler", "İYBF 2026-2027 Akademik Yılı Güz Dönemi Ders Programı (Güncel).pdf")

UNVAN_KELIMELER = {"prof", "doç", "doğ", "dr", "öğr", "öğretim", "üyesi", "üye", "gör", "arş", "lab", "doktor"}
UNVAN_ETIKET = {0: "Prof. Dr.", 1: "Doç. Dr.", 2: "Dr. Öğr. Üyesi", 3: "Doktor", 4: "Öğr. Gör. Dr.",
                5: "Öğr. Gör.", 6: "Arş. Gör.", 99: ""}
OZEL_TIPLER = {"alttan", "devamlı alttan", "üstten", "yükseltme"}  # zorunlu/seçmeli dışındaki ders alma tipleri
BELIRSIZ = "~BELİRSİZ"
BELIRSIZ_AD = "— Hocası belirlenmemiş dersler (komisyon dolduracak) —"


# ---------------------------------------------------------------- metin yardımcıları
def tr_upper(s):
    return s.replace("i", "İ").replace("ı", "I").upper()


def tr_lower(s):
    return s.replace("I", "ı").replace("İ", "i").lower()


_ASCII = str.maketrans("çğıöşüâîû", "cgiosuaiu")


def ad_anahtar(ad):
    """Karşılaştırma için: küçük harf, Türkçe karakterler ve noktalama yok ('Economics-I' = 'Economics I')."""
    return re.sub(r"[^a-z0-9]", "", tr_lower(ad).translate(_ASCII)).replace("behaviour", "behavior")


def hoca_ayir(ham):
    """'Dr.Öğr.Üyesi M. Sercan SEVİMLİ' -> (ünvan etiketi, ad, anahtar 'SEVİMLİ|M', ünvan sırası)."""
    parcalar = [t for t in re.split(r"[\s.]+", (ham or "").strip()) if t]
    i = 0
    while i < len(parcalar) and tr_lower(parcalar[i]) in UNVAN_KELIMELER:
        i += 1
    kel = {tr_lower(t) for t in parcalar[:i]}
    adlar = parcalar[i:]
    if not adlar:
        return "", "", "", 99
    if "prof" in kel: sira = 0
    elif kel & {"doç", "doğ"}: sira = 1
    elif "arş" in kel: sira = 6
    elif "gör" in kel: sira = 4 if "dr" in kel else 5
    elif kel & {"üyesi", "üye", "öğretim", "öğr"}: sira = 2
    elif kel & {"doktor", "dr"}: sira = 3
    else: sira = 99
    ad = " ".join(t + "." if len(t) == 1 else t for t in adlar)
    anahtar = (tr_upper(adlar[-1]) + "|" + tr_upper(adlar[0][0])).replace("İ", "I")
    return UNVAN_ETIKET[sira], ad, anahtar, sira


def ad_parcalari(ad):
    return [t for t in re.split(r"[\s.]+", tr_upper(ad).replace("İ", "I")) if t]


class KisiTablosu:
    """Aynı kişinin farklı yazımlarını tek anahtara indirir: 'Merve DOĞRUEL' = 'MERVE DOĞRUEL ANUĞLU'.
    Kural: baş harf aynı ve en az 2 uzun ad parçası ortak (tek ortak soyad yetmez: Elif ŞİMŞEK ≠ İlknur ŞİMŞEK)."""

    def __init__(self):
        self.kisiler = []

    def anahtar(self, hoca_ham):
        if not hoca_ham:
            return ""
        _, ad, key, _ = hoca_ayir(hoca_ham)
        if not ad:
            return ""
        p = ad_parcalari(ad)
        for k, q in self.kisiler:
            if p[0][0] == q[0][0] and len({x for x in p if len(x) > 1} & {x for x in q if len(x) > 1}) >= 2:
                return k
        self.kisiler.append((key, p))
        return key

    def ayni(self, a, b):
        ka, kb = self.anahtar(a), self.anahtar(b)
        return not ka or not kb or ka == kb  # biri bilinmiyorsa çelişki yok sayılır


# ---------------------------------------------------------------- OBS
def subeleri_oku():
    ws = openpyxl.load_workbook(SUBE_DOSYASI, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    baslik = [re.sub(r"_\d+$", "", str(x).strip()) for x in rows[0]]
    ix ={b: i for i, b in enumerate(baslik)}
    out = []
    for r in rows[1:]:
        if not r[ix["Ders Kodu"]]:
            continue
        out.append({"kod": str(r[ix["Ders Kodu"]]).strip(), "ad": str(r[ix["Ders Adı"]]).strip(),
                    "program": str(r[ix["Program"]]).strip(), "sinif": str(r[ix["Sınıf" if "Sınıf" in ix else "Snf"]]).strip(),
                    "hoca": str(r[ix["Öğretim Elemanı"]] or "").strip()})
    return out


def kayitlari_oku():
    """-> (kod -> öğrenci kümesi, kod -> ad, kod -> {öğrenci no: ders alma tipi},
    kod -> {öğrenci programı: öğrenci kümesi})."""
    if not os.path.exists(KAYIT_DOSYASI):
        return {}, {}, {}, {}
    ws = openpyxl.load_workbook(KAYIT_DOSYASI, read_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    baslik = [re.sub(r"_\d+$", "", str(x).strip()) for x in rows[0]]
    i_no, i_kod, i_ad = baslik.index("Öğrenci No"), baslik.index("D.Kodu"), baslik.index("Ders Adı")
    i_tip, i_prog = baslik.index("Ders Alma Tip"), baslik.index("Öğrenci Programı")
    ogr, ad, tip = collections.defaultdict(set), {}, collections.defaultdict(dict)
    prog = collections.defaultdict(lambda: collections.defaultdict(set))
    for r in rows[1:]:
        if r[i_no] and r[i_kod]:
            k, no = str(r[i_kod]).strip(), str(r[i_no]).strip()
            ogr[k].add(no)
            ad.setdefault(k, str(r[i_ad]).strip())
            tip[k][no] = str(r[i_tip] or "").strip()
            prog[k][str(r[i_prog] or "").strip()].add(no)
    return ogr, ad, tip, prog


# ---------------------------------------------------------------- PDF
GUN_ESLE = {"pazartesi": "Pazartesi", "monday": "Pazartesi", "salı": "Salı", "tuesday": "Salı",
            "çarşamba": "Çarşamba", "wednesday": "Çarşamba", "perşembe": "Perşembe", "thursday": "Perşembe",
            "cuma": "Cuma", "friday": "Cuma",
            "thusrday": "Perşembe"}  # PDF'teki yazım hatası ('Introduction to Law Thusrday')
KOD_RE = re.compile(r"\b(\d{2,4}[A-Z]{1,3}\d{3,4})\b")
ZAMAN_RE = re.compile(r"\b(" + "|".join(GUN_ESLE) + r")\s+(\d{1,2})\.(\d{2})\s+(\d{1,2})\.(\d{2})\b", re.I)
UNVAN_BAS = re.compile(r"\b(?:Prof|Doç|Doğ|Dr|Öğr|Arş|Doktor|Lab)\b")
YARIYIL_RE = re.compile(r"(\d)\.\s*YARIYIL", re.I)
ON_EK = re.compile(r"^(?:Elective\s+(?:Dept|Faculty|University)\s*\d*|Department(?:al)?\s+Elect\w+|Faculty\s+Elective|"
                   r"Bölüm\s+Seçmeli|Fakülte\s+Seçmeli|Seçmeli\s+(?:Bölüm|Fakülte|Üniversite))\s*[:/\-–]*\s*", re.I)
SON_EK = re.compile(r"\s*\((?:Department|Faculty|Departmental)[^)]*\)\s*$", re.I)


def ders_adi_temizle(ad):
    """'Elective Dept 5 / Global Marketing' -> 'Global Marketing' (havuz/seçmeli ön ekleri atılır)"""
    ad = re.sub(r"\s+", " ", ad).strip()
    return SON_EK.sub("", ON_EK.sub("", ad)).strip()


def bolum_etiketi(prefix, yy):
    """PDF satır başındaki bölüm yazısı -> okunur program adı + yarıyıl"""
    f = ad_anahtar(prefix)
    ing = "(ing" in tr_lower(prefix).translate(_ASCII) or any(x in f for x in ("political", "business", "economics"))
    for anahtar, ad in (("siyaset", "Siyaset Bilimi ve Uluslararası İlişkiler"),
                        ("political", "Siyaset Bilimi ve Uluslararası İlişkiler"),
                        ("isletme", "İşletme"), ("business", "İşletme"),
                        ("ekonomi", "Ekonomi ve Finans"), ("economics", "Ekonomi ve Finans"),
                        ("ticaretvefinansman", "Uluslararası Ticaret ve Finansman"),
                        ("lojistik", "Lojistik Yönetimi"), ("havac", "Havacılık Yönetimi"),
                        ("yonetimbilisim", "Yönetim Bilişim Sistemleri"),
                        ("elektronikticaret", "Elektronik Ticaret ve Yönetimi")):
        if anahtar in f:
            return f"{ad}{' (İng.)' if ing else ''}" + (f" {yy}. yarıyıl" if yy else "")
    return (prefix.title() if prefix else "") + (f" {yy}. yarıyıl" if yy and prefix else "")


def pdf_oku():
    """Ders programı PDF'i -> [{kod ('' olabilir), ad, hoca, saat (None olabilir), prefix, yy}]"""
    if not os.path.exists(PROGRAM_PDF):
        return []
    import pypdf
    satirlar = []
    for p in pypdf.PdfReader(PROGRAM_PDF).pages:
        satirlar += (p.extract_text() or "").split("\n")
    satirlar = [re.sub(r"\s+", " ", s.replace("\xa0", " ")).strip() for s in satirlar]
    # ön geçiş: bilinen bölüm yazıları (kodsuz satırların başındaki bölümü ayırmak için)
    bolumler = set()
    for s in satirlar:
        m, h = KOD_RE.search(s), YARIYIL_RE.search(s)
        if m and m.start() > 3:
            bolumler.add(s[:m.start()].strip())
        elif h and not m:
            bolumler.add(s[:h.start()].strip())
    bolumler = sorted((b for b in bolumler if len(b) > 3), key=len, reverse=True)
    out, yy = [], None
    for s in satirlar:
        h, k, z = YARIYIL_RE.search(s), KOD_RE.search(s), ZAMAN_RE.search(s)
        if h and not k and not z:
            yy = int(h.group(1)); continue
        if not k and not z:
            continue
        if k:
            prefix, bas = s[:k.start()].strip(), k.end()
        else:
            prefix = next((b for b in bolumler if s.startswith(b)), "")
            bas = len(prefix)
        if z:
            ad = s[bas:z.start()]
            kalan = s[z.end():].strip()
            gun = GUN_ESLE[z.group(1).lower()]
            if UNVAN_BAS.match(kalan):
                derslik, hoca = "", kalan
            else:
                parca = kalan.split(None, 1)
                derslik, hoca = (parca[0] if parca else ""), (parca[1] if len(parca) > 1 else "")
            saat = f"{gun} {int(z.group(2)):02d}:{z.group(3)}-{int(z.group(4)):02d}:{z.group(5)}" + \
                   (f" ({derslik})" if derslik else "")
        else:  # ortak dersler: gün/saat yok
            t = UNVAN_BAS.search(s, bas)
            ad, hoca, saat = (s[bas:t.start()], s[t.start():], None) if t else (s[bas:], "", None)
        ad = ders_adi_temizle(ad)
        if len(ad) < 3:
            continue
        out.append({"kod": k.group(1) if k else "", "ad": ad, "hoca": hoca.strip(), "saat": saat,
                    "prefix": prefix, "yy": yy})
    return out


# ---------------------------------------------------------------- birleştirme
def subelere_ayir(d):
    """Aynı ders+hoca farklı bölümlere farklı gün/saatte veriliyorsa (ör. İktisada Giriş I: Salı 2 bölüm, Çarşamba 3 bölüm)
    her gün/saat grubu ayrı şube (ayrı satır) olur. Aynı bölümün haftada iki ayrı saati ise tek şube kalır."""
    grup = {}  # gün+saat -> {"taban": bölümler, "kayit": kayıtlar, "saatler": derslikli yazımlar}
    for taban, etiket, saat in d["kayit"]:
        if saat:
            z = re.sub(r"\s*\([^)]*\)\s*$", "", saat)  # yalnızca derslik farkı şube ayırmaz
            g = grup.setdefault(z, {"taban": set(), "kayit": [], "saatler": []})
            if taban:
                g["taban"].add(taban)
            g["kayit"].append((taban, etiket, saat))
            if saat not in g["saatler"]:
                g["saatler"].append(saat)
    bilesen = []  # ortak bölümü olan saat grupları birleşir
    for g in grup.values():
        eslesen = [b for b in bilesen if b["taban"] & g["taban"]]
        yeni = {"taban": set(g["taban"]), "kayit": list(g["kayit"]), "saatler": list(g["saatler"])}
        for b in eslesen:
            yeni["taban"] |= b["taban"]; yeni["kayit"] += b["kayit"]; yeni["saatler"] = b["saatler"] + yeni["saatler"]
            bilesen.remove(b)
        bilesen.append(yeni)
    # bölüm adı yazılmayan satırlar (PDF'te alta devam eden oturumlar) ayrı şube sayılmaz: bölümlü ilk şubeye eklenir
    ana = [b for b in bilesen if b["taban"]]
    if ana:
        for b in bilesen:
            if not b["taban"]:
                ana[0]["kayit"] += b["kayit"]; ana[0]["saatler"] += b["saatler"]
        bilesen = ana
    if len(bilesen) <= 1:
        return [d]
    sonuc = []
    for b in bilesen:
        s = dict(d)
        s["taban"], s["saatler"] = b["taban"], list(dict.fromkeys(b["saatler"]))
        s["prog"] = list(dict.fromkeys(e for _, e, _ in b["kayit"] if e))
        s["kodlar"], s["obs_kodlar"] = list(d["kodlar"]), list(d["obs_kodlar"])
        sonuc.append(s)
    return sonuc


def ders_listesi():
    """-> ({hoca anahtarı: {unvan, ad, rank, dersler:[...]}}, rapor)  rapor: kontrol sayfaları için listeler."""
    kisi = KisiTablosu()
    ogr, kayit_ad, tip, kayit_prog = kayitlari_oku()
    sube = subeleri_oku()
    sube_kod = collections.defaultdict(list)
    sube_ad = collections.defaultdict(list)
    for s in sube:
        sube_kod[s["kod"]].append(s)
        sube_ad[ad_anahtar(s["ad"])].append(s)
    kayit_ad_idx = collections.defaultdict(set)
    for k, a in kayit_ad.items():
        kayit_ad_idx[ad_anahtar(a)].add(k)
    ic_kod = collections.defaultdict(list)  # '2010G2002[2021B2021]' -> ic kod 2021B2021
    for k in ogr:
        if "[" in k:
            ic_kod[k.split("[")[1].rstrip("]")].append(k)

    # 1) PDF kayıtlarını derslere birleştir (kod ya da ad eşleşmesi; hoca çelişmiyorsa)
    dersler = []
    for r in pdf_oku():
        hedef = None
        for d in dersler:
            if not kisi.ayni(d["hoca"], r["hoca"]):
                continue
            if (r["kod"] and r["kod"] in d["kodlar"]) or ad_anahtar(d["ad"]) == ad_anahtar(r["ad"]):
                hedef = d
                break
        if hedef is None:
            hedef = {"kodlar": [], "ad": r["ad"], "hoca": r["hoca"], "prog": [], "saatler": [], "obs_kodlar": [],
                     "eslesme": "", "kayit": [], "taban": None}
            dersler.append(hedef)
        if r["kod"] and r["kod"] not in hedef["kodlar"]:
            hedef["kodlar"].append(r["kod"])
        if not hedef["hoca"]:
            hedef["hoca"] = r["hoca"]
        p = bolum_etiketi(r["prefix"], r["yy"]) if r["prefix"] else ""
        if p and p not in hedef["prog"]:
            hedef["prog"].append(p)
        if r["saat"] and r["saat"] not in hedef["saatler"]:
            hedef["saatler"].append(r["saat"])
        hedef["kayit"].append((bolum_etiketi(r["prefix"], None) if r["prefix"] else "", p, r["saat"]))
    dersler = [s for d in dersler for s in subelere_ayir(d)]

    # 2) OBS verisini PDF derslerine bağla: öğrenci sayısı + kod/hoca karşılaştırması
    bagli_kodlar, rapor = set(), {"hoca_farki": [], "kodsuz": [], "obs_de_var_pdf_de_yok": [], "ogrencisiz": [],
                                  "kod_farki": []}
    for d in dersler:
        kodlar = list(d["kodlar"])
        ak = ad_anahtar(d["ad"])
        if not kodlar:  # PDF'te kod yazmıyor: OBS'te ad (+hoca, +bölüm) ile ara
            adaylar = [s for s in sube_ad.get(ak, []) if kisi.ayni(s["hoca"], d["hoca"])]
            if len({s["kod"] for s in adaylar}) > 1 and d["prog"]:  # aynı ad birkaç bölümde: bölümü uyanı seç
                taban = {ad_anahtar(re.sub(r"\s*\(İng\.\)|\s*\d+\. yarıyıl", "", p)) for p in d["prog"]}
                daha = [s for s in adaylar if any(ad_anahtar(s["program"]).startswith(t) for t in taban)]
                adaylar = daha or adaylar
            d["obs_kodlar"] = list(dict.fromkeys(s["kod"] for s in adaylar))
            if not d["obs_kodlar"]:
                d["obs_kodlar"] = sorted(kayit_ad_idx.get(ak, []))
            rapor["kodsuz"].append((d["ad"], d["hoca"], ", ".join(d["prog"]), ", ".join(d["obs_kodlar"]) or "OBS'te bulunamadı"))
            kodlar = d["obs_kodlar"]
        tum = set(kodlar)
        for k in list(kodlar):
            tum.update(ic_kod.get(k, []))
        ogrenci = set().union(*[ogr.get(k, set()) for k in tum]) if tum else set()
        d["ogrenci"] = len(ogrenci)
        if not ogrenci and d["kodlar"]:  # PDF kodu OBS'te öğrencisiz: aynı adlı OBS kodu var mı? (kod farkı)
            alt = sorted(kayit_ad_idx.get(ak, set()) - set(d["kodlar"]))
            if alt:
                ogrenci = set().union(*[ogr.get(k, set()) for k in alt])
                d["ogrenci"] = len(ogrenci)
                rapor["kod_farki"].append((" / ".join(d["kodlar"]), d["ad"], d["hoca"], ", ".join(alt),
                                           " / ".join(d["kodlar"])))  # şimdilik PDF kodu esas
                tum.update(alt)
        bagli_kodlar |= tum
        d["tum"], d["ogr"] = tum, ogrenci
        sayac = collections.Counter()  # dersi alan öğrencilerin bölümleri (gürültüyü atmak için %5 / en az 2 öğrenci)
        for k in tum:
            for p, s in kayit_prog.get(k, {}).items():
                sayac[p.replace("(İngilizce)", "(İng.)")] += len(s)
        toplam = sum(sayac.values()) or 1
        d["bolumler"] = [p for p, n in sayac.most_common() if p and n >= max(2, 0.05 * toplam)]
        if d["taban"]:  # dersin şubelerinden biri: öğrenci sayısı ve bölümler yalnızca bu şubenin bölümleri için
            sube_ogr = set()
            for k in tum:
                for p, s in kayit_prog.get(k, {}).items():
                    if p.replace("(İngilizce)", "(İng.)") in d["taban"]:
                        sube_ogr |= s
            d["ogrenci"] = len(sube_ogr)
            d["bolumler"] = [t for t in sorted(d["taban"]) if sayac.get(t)]  # d["ogr"] tüm şubeler: çakışma analizi için
        d["ozel"] = {no for k in tum for no, t in tip.get(k, {}).items()
                     if no in ogrenci and tr_lower(t) in OZEL_TIPLER}
        # hoca karşılaştırması (yalnızca PDF kodu varsa): PDF hocası OBS'tekilerden hiçbiriyle uyuşmuyorsa
        obs_h = [s["hoca"] for k in d["kodlar"] for s in sube_kod.get(k, []) if s["hoca"]]
        if d["hoca"] and obs_h and not any(kisi.ayni(h, d["hoca"]) for h in obs_h):
            rapor["hoca_farki"].append((" / ".join(d["kodlar"]), d["ad"], d["hoca"],
                                        "; ".join(dict.fromkeys(obs_h))))
        if not d["ogrenci"]:
            rapor["ogrencisiz"].append((" / ".join(kodlar) or "—", d["ad"], d["hoca"], ", ".join(d["prog"]),
                                        " / ".join(d["saatler"]) or "saat yok"))

    # 3) OBS'te olup PDF'te olmayanlar (forma eklenmez, rapora yazılır)
    for kod in sorted(set(sube_kod) | set(ogr)):
        if kod in bagli_kodlar or ("[" in kod and kod.split("[")[1].rstrip("]") in bagli_kodlar):
            continue
        ad = (sube_kod[kod][0]["ad"] if sube_kod.get(kod) else kayit_ad.get(kod, ""))
        hocalar = "; ".join(dict.fromkeys(s["hoca"] for s in sube_kod.get(kod, []) if s["hoca"]))
        prog = "; ".join(dict.fromkeys(s["program"] for s in sube_kod.get(kod, [])))
        benzer = [" / ".join(d["kodlar"]) or "kodsuz" for d in dersler if ad_anahtar(d["ad"]) == ad_anahtar(ad)]
        rapor["obs_de_var_pdf_de_yok"].append((kod, ad, hocalar, prog, len(ogr.get(kod, ())),
                                               ("evet: " + ", ".join(benzer)) if benzer else "hayır"))

    # 4) sınav birimleri (çakışma analizi için): aynı dersin farklı hoca satırları tek birim sayılır
    birimler = {}
    for d in dersler:
        anahtar = frozenset(d["tum"]) or ad_anahtar(d["ad"])
        b = birimler.setdefault(anahtar, {"kodlar": [], "ad": d["ad"], "hocalar": [], "ogr": set(), "ozel": set(),
                                          "tum": set()})
        for k in d["kodlar"] or d["obs_kodlar"]:
            if k not in b["kodlar"]:
                b["kodlar"].append(k)
        if d["hoca"] and hoca_ayir(d["hoca"])[1] not in b["hocalar"]:
            b["hocalar"].append(hoca_ayir(d["hoca"])[1])
        b["ogr"] |= d["ogr"]; b["ozel"] |= d["ozel"]; b["tum"] |= d["tum"]

    # 5) hoca bazında grupla
    hocalar = {}
    for d in dersler:
        if d["hoca"]:
            unvan, ad, key, sira = hoca_ayir(d["hoca"])
            key = kisi.anahtar(d["hoca"])
        else:
            unvan, ad, key, sira = "", BELIRSIZ_AD, BELIRSIZ, 100
        h = hocalar.setdefault(key, {"unvan": unvan, "ad": ad, "rank": sira, "dersler": []})
        if key != BELIRSIZ and len(ad.split()) > len(h["ad"].split()):
            h["ad"] = ad
        kod = " / ".join(d["kodlar"]) or (" / ".join(d["obs_kodlar"]) + " (OBS)" if d["obs_kodlar"] else "—")
        h["dersler"].append({"kod": kod, "ad": d["ad"], "program": "\n".join(d["prog"]) or "—",
                             "bolumler": "\n".join(d["bolumler"]) or "—", "sayi": d["ogrenci"], "saat": " / ".join(d["saatler"]) or "Saat yok (ortak/belirsiz)"})
    return hocalar, rapor, list(birimler.values())
