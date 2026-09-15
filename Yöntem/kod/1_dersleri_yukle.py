# -*- coding: utf-8 -*-
"""
Adım 1 — Talep dosyasındaki dersleri benzersiz sınavlara indirip
şablonun "Dersler" sekmesine yazar. Yöntem: Sinav Programi Hazirlama Rehberi.md

Girdi : Belgeler/Sınav Programı Talep Dosyası.xlsx
        Belgeler/Ders öğrenci sayıları.xlsx   (sistem sayıları, yedek kaynak)
Çıktı : şablon (donem_ayarlari.SABLON_DOSYASI) -> yalnızca "Dersler" sekmesi
        Yöntem/kod/dersler.json                     -> sonraki adımlar için yapısal veri
"""
import json, os, re, sys, collections
import openpyxl

try:
    from donem_ayarlari import *   # dönem bilgileri ve şifreler (GitHub'a yüklenmez)
except ModuleNotFoundError:
    raise SystemExit("donem_ayarlari.py bulunamadı: 'donem_ayarlari_ornek.py' dosyasını 'donem_ayarlari.py' adıyla kopyalayıp doldurun.")
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

KOK = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
TALEP = os.path.join(KOK, "Belgeler", TALEP_DOSYASI)
SISTEM = os.path.join(KOK, "Belgeler", SISTEM_DOSYASI)
SABLON = os.path.join(KOK, SABLON_DOSYASI)
JSON_CIKTI = os.path.join(KOK, "Yöntem", "kod", "dersler.json")

GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
NISAN = {int(t.split(".")[0]): g for g, t in SINAV_TARIHLERI.items()}   # talepteki "8 Nisan" -> gün
BOLUM_SIRA = ["SBUİ", "PSIR", "İŞL", "BUS", "UTF", "LOJ", "HAV", "YBS", "ECF", "ETİC", "HAVUZ2", "HAVUZ4"]
INGILIZCE = {"PSIR", "BUS", "ECF"}
DERSLIK_SAAT_SINIRI = 50  # dk; üstü 2 oturum kapatır

# ---------------------------------------------------------------- genel kurallar (dönem istisnaları donem_ayarlari.py içinde)
SURE_YUVARLA_UST = 60          # 60 dk'ya kadar yazılan süreler 50 dk sayılır; üstü (90 dk) aynen kalır
GOZETMEN_ESIGI = 30            # öğrenci >= 30 -> gözetmen; altı -> dersin hocası girer
LAB_KAPASITE = 40
ONAYLI_BOLUM_EKLERI = "hepsi"  # hocanın "dersi alan bölümler"de yazdığı bölümler çakışma grubuna eklenir
SISTEM_BOLUM_EKLE = True       # sistemde dersi alan görünen her bölüm de eklenir (çakışma riski olmasın)
IZINSIZ_HOCA_DERS_SAATI = True # izin günü bilinmiyorsa (çoğu ders saati ücretli) sınav, dersin normal gün/saatine konmaya çalışılır
TUM_BOLUM_KODLARI = ["SBUİ", "PSIR", "İŞL", "BUS", "UTF", "LOJ", "HAV", "YBS", "ECF", "ETİC"]


# ---------------------------------------------------------------- metin yardımcıları
def tr_lower(s):
    return str(s).replace("I", "ı").replace("İ", "i").lower()

def tr_upper(s):
    return str(s).replace("i", "İ").replace("ı", "I").upper()

def sade(s):
    t = tr_lower(s)
    for a, b in zip("çğıöşüâîû", "cgiosuaiu"):
        t = t.replace(a, b)
    return t

def bosmu(v):
    return v is None or str(v).strip() in ("", "-")

def metin(v):
    return "" if v is None else re.sub(r"\s+", " ", str(v)).strip()

ONEK = re.compile(
    r"^\s*(?:(?:bölüm|bölm|fakülte|fakülteseçmeli|üniversite|seçmeli|departmental|department|faculty|"
    r"university|elective|dept|course)\s*[-]?\s*)+(?:(?:[ivx]+|\d+)\b)?\s*[/:]?\s*",
    re.IGNORECASE)
KOD_RE = re.compile(r"\b(\d{2,4}[A-Z]{1,2}\d{3,5})(?:_\d+)?\b")

def ders_adi_temizle(ad):
    ad = metin(ad)
    kod = None
    m = KOD_RE.search(ad)
    if m:
        kod = m.group(1)
        ad = (ad[:m.start()] + ad[m.end():]).strip()
    ad2 = ONEK.sub("", ad).strip()
    if ad2.startswith("(") and ad2.endswith(")"):
        ad2 = ad2[1:-1].strip()
    ad2 = re.sub(r"\((\d)\.\s*Sınıf\)", "", ad2).strip()
    return (ad2 or ad).rstrip(" .,;:"), kod

def ad_anahtar(ad):
    return re.sub(r"[^0-9a-z]", "", sade(ad))

def kod_temizle(k):
    k = metin(k)
    m = KOD_RE.search(k)
    return m.group(1) if m else None

def levenshtein(a, b):
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


# ---------------------------------------------------------------- öğretim elemanı
UNVAN = {"PROF", "DOC", "DR", "OGR", "GOR", "UYESI", "UYE", "ARS", "UZM", "OGRETIM", "MSC", "PHD", "OR", "DOKTOR"}

def hoca_parcala(ad):
    s = re.sub(r"(?<=[a-zçğıöşü])(?=[A-ZÇĞİÖŞÜ])", " ", metin(ad))  # "SercanSEVİMLİ"
    s = re.sub(r"[^\wçğıöşüÇĞİÖŞÜ]", " ", s)
    parca = [p for p in sade(s).upper().split() if p not in UNVAN]
    return parca

class Hocalar:
    def __init__(self):
        self.kisiler = []  # [{'soyad','bas','yazimlar':set}]

    def anahtar(self, ad):
        if bosmu(ad) or sade(ad).strip() in ("acilacak",):
            return None
        p = hoca_parcala(ad)
        if not p:
            return None
        soyad, bas = p[-1], p[0][0]
        for i, k in enumerate(self.kisiler):
            if self._ayni(k, soyad, bas, p):
                k["yazimlar"].add(metin(ad)); k["tum"].update(p)
                return i
        self.kisiler.append({"soyad": soyad, "bas": bas, "yazimlar": {metin(ad)}, "tum": set(p)})
        return len(self.kisiler) - 1

    @staticmethod
    def _ayni(k, soyad, bas, p):
        if k["bas"] != bas:
            return False
        esik = 1 if len(soyad) <= 6 else 2
        return (k["soyad"] == soyad or levenshtein(k["soyad"], soyad) <= esik
                or (len(soyad) > 3 and soyad in k["tum"]) or (len(k["soyad"]) > 3 and k["soyad"] in p))

    def ayni_kisi(self, ad, i):
        """Kayıt eklemeden: 'ad' i numaralı kişi mi?"""
        if i is None or bosmu(ad):
            return None
        p = hoca_parcala(ad)
        return bool(p) and self._ayni(self.kisiler[i], p[-1], p[0][0], p)

    def kanonik(self, i):
        if i is None:
            return ""
        return max(sorted(self.kisiler[i]["yazimlar"]), key=lambda y: (len(hoca_parcala(y)), len(y), sum(ch.isupper() for ch in y), y))  # kararlı seçim


# ---------------------------------------------------------------- blok başlıkları
def blok_bolum(baslik):
    u = tr_upper(baslik)
    if "HAVUZU" in u:
        n = re.search(r"HAVUZU\s*(\d)", u)
        return ("HAVUZ2", HAVUZ_YARIYIL["HAVUZ2"]) if n and n.group(1) == "2" else ("HAVUZ4", HAVUZ_YARIYIL["HAVUZ4"])
    m = re.search(r"(\d+)\s*\.\s*(YARIYIL|SEMESTER)", u)
    if not m:
        return None
    yy = int(m.group(1))
    if "POLITICAL SCIENCE" in u: return ("PSIR", yy)
    if "SİYASET" in u: return ("SBUİ", yy)
    if "BUSINESS ADMINISTRATION" in u or "İŞLETME (ING)" in u: return ("BUS", yy)
    if "İŞLETME" in u: return ("İŞL", yy)
    if "ELEKTRONİK TİCARET" in u: return ("ETİC", yy)
    if "ULUSLARARASI TİCARET" in u: return ("UTF", yy)
    if "LOJİSTİK" in u: return ("LOJ", yy)
    if "HAVACILIK" in u: return ("HAV", yy)
    if "BİLİŞİM" in u: return ("YBS", yy)
    if "EKONOMİ" in u or "ECONOMICS" in u: return ("ECF", yy)
    return None

def bolum_listesi(text, ingilizce):
    """'İşletme+YBS+UTF' gibi serbest metinden bölüm kodları."""
    t = sade(text)
    if "ayni saat" in t or "seçmeli" in tr_lower(text) and "uni" in t:
        return set()
    out = set()
    if re.search(r"siyaset|sbu|siy\.?\s*bil", t): out.add("PSIR" if ingilizce else "SBUİ")
    if re.search(r"isletme|isletne", t): out.add("BUS" if ingilizce else "İŞL")
    if re.search(r"eko|economics", t): out.add("ECF")
    if re.search(r"utf|uluslararasi ticaret", t): out.add("UTF")
    if "lojistik" in t: out.add("LOJ")
    if "havacilik" in t: out.add("HAV")
    if re.search(r"\bybs\b|bilisim", t): out.add("YBS")
    if re.search(r"\bety\b|elektronik ticaret|e-ticaret", t): out.add("ETİC")
    return out


# ---------------------------------------------------------------- talep ayrıştırma
GUN_KELIME = [("pazartesi", "Pazartesi"), ("monday", "Pazartesi"), ("sali", "Salı"), ("tuesday", "Salı"),
              ("carsamba", "Çarşamba"), ("cars", "Çarşamba"), ("wednesday", "Çarşamba"),
              ("persembe", "Perşembe"), ("thursday", "Perşembe"), ("cuma", "Cuma"), ("friday", "Cuma")]

def gunleri_bul(text):
    t = sade(text)
    bulunan = []
    for k, g in GUN_KELIME:
        if re.search(r"\b" + k, t) and g not in bulunan:
            bulunan.append(g)
    return sorted(bulunan, key=GUNLER.index)

def saat_int(v):
    if bosmu(v):
        return None
    if hasattr(v, "hour"):
        return v.hour
    m = re.match(r"\s*(\d{1,2})", str(v))
    return int(m.group(1)) if m else None

def talep_ayristir(text, ders_gunu, ders_bas, ders_bit):
    """-> dict(gunler, saatler, haric, ders_saatinde, tarih)"""
    sonuc = {"gunler": [], "saatler": [], "haric": [], "ders_saatinde": False, "tarih": []}
    if bosmu(text):
        return sonuc
    t = tr_lower(text)
    for m in re.finditer(r"(\d{1,2})\s*" + tr_lower(AY_ADI), t):
        g = NISAN.get(int(m.group(1)))
        if g:
            sonuc["tarih"].append(f"{SINAV_TARIHLERI[g]} {g}")
            sonuc["gunler"].append(g)
    t = re.sub(r"\d{1,2}\s*" + tr_lower(AY_ADI), " ", t)
    # hariç parçaları
    for m in re.finditer(r"\(([^)]*hariç[^)]*)\)", t):
        sonuc["haric"] += [h for h in _saatler(m.group(1)) if 9 <= h <= 17]
    t = re.sub(r"\([^)]*hariç[^)]*\)", " ", t)
    for g in gunleri_bul(t):
        if g not in sonuc["gunler"]:
            sonuc["gunler"].append(g)
    saatler = _saatler(t)
    if "sabah" in t:
        saatler = sorted(set(saatler) | {9, 10, 11})
    ders_flag = bool(re.search(r"ders\w*\s+gün", t) or re.search(r"ders\w*\s+saat", t))
    if ders_flag:
        sonuc["ders_saatinde"] = True
        if not sonuc["gunler"]:
            sonuc["gunler"] = gunleri_bul(ders_gunu or "")
        if not saatler and "saat" in t and ders_bas is not None and ders_bit is not None:
            saatler = list(range(ders_bas, ders_bit))
    sonuc["saatler"] = sorted({h for h in saatler if 9 <= h <= 17})
    sonuc["gunler"] = sorted(set(sonuc["gunler"]), key=GUNLER.index)
    return sonuc

def _saatler(t):
    out = []
    kalan = t
    for m in re.finditer(r"(\d{1,2})[:.](\d{2})(?:\s*-\s*\d{1,2}[:.]\d{2})?", t):
        out.append(int(m.group(1)))
    kalan = re.sub(r"(\d{1,2})[:.](\d{2})(?:\s*-\s*\d{1,2}[:.]\d{2})?", " ", kalan)
    for m in re.finditer(r"\b\d{1,2}(?:\s*[-/]\s*\d{1,2})+\b", kalan):
        parca = [int(x) for x in re.split(r"\s*[-/]\s*", m.group(0))]
        if "-" in m.group(0) and "/" not in m.group(0) and len(parca) == 2:
            out += list(range(parca[0], parca[1]))
        else:
            out += parca
    return out

def sure_dk(v):
    if bosmu(v):
        return None
    m = re.search(r"(\d+)\s*(dk|dak|dakika|saat)", tr_lower(v))
    if not m:
        return None
    return int(m.group(1)) * (60 if m.group(2) == "saat" else 1)

def gozetmen(v):
    if bosmu(v):
        return None
    t = sade(v)
    if t.startswith("var"): return "Var"
    if t.startswith("yok") or t.startswith("hayir"): return "Yok"
    return None


# ---------------------------------------------------------------- 1) satırları oku
def satirlari_oku():
    ws = openpyxl.load_workbook(TALEP, data_only=True).active
    satirlar, blok, onceki = [], None, None
    for no, r in enumerate(ws.iter_rows(values_only=True), 1):
        r = list(r) + [None] * (15 - len(r))
        dolu = [i for i, v in enumerate(r) if not bosmu(v)]
        if not dolu:
            onceki = None; continue
        ilk = metin(r[0])
        if dolu == [0]:
            b = blok_bolum(ilk)
            if b:
                blok = {"bolum": b[0], "yy": b[1], "havuz": b[0].startswith("HAVUZ")}
            onceki = None; continue
        tum = " ".join(metin(r[i]) for i in dolu)
        if re.search(r"\b(BÖLÜM|DEPARTMENT|SınavKODU|SınavADI|DERSKODU)\b", tum) or \
           all(re.search(r"SAATİ|BAŞLANGIÇ|BİTİŞ", metin(r[i])) for i in dolu):
            onceki = None; continue
        if blok is None:
            continue
        h = blok["havuz"]
        o = 1 if h else 0  # havuz blokları A-G'de bir sütun sola kayık
        kod_i, ad_i, gun_i, bas_i, bit_i, dl_i, hoca_i = (0, 1, 2, 3, 4, 5, 6) if h else (1, 2, 3, 4, 5, 6, 7)
        # devam satırı: önceki satırın bir hücresinin alt satıra taşmış kısmı
        if onceki is not None and bosmu(r[kod_i]) and (h or bosmu(r[0])) and len(dolu) == 1:
            i = dolu[0]
            alan = {ad_i: "ad", hoca_i: "hoca", gun_i: "gun"}.get(i)
            if alan:
                onceki[alan] = (metin(onceki[alan]) + " " + metin(r[i])).strip()
                continue
        ad_ham = metin(r[ad_i])
        s = {
            "satir": no, "bolum": blok["bolum"], "yy": blok["yy"], "havuz": h,
            "kod": kod_temizle(r[kod_i]), "ad": ad_ham, "gun": metin(r[gun_i]),
            "bas": saat_int(r[bas_i]), "bit": saat_int(r[bit_i]),
            "derslik": metin(r[dl_i]), "hoca": metin(r[hoca_i]),
            "bolumler": metin(r[8]), "sayi": r[9], "izin": metin(r[10]),
            "talep": metin(r[11]), "gozetmen": metin(r[12]), "sure": metin(r[13]),
            "ek": metin(r[14]),
        }
        if "ORTAK DERSLER" in tr_upper(" ".join(metin(x) for x in r[3:7])):
            s["ortak"] = True
        s["yer_tutucu"] = bool(re.search(r"alt tarafta|bakiniz|^university elective\s*[ivx]*$",
                                         sade(ad_ham))) or not ad_ham
        satirlar.append(s)
        onceki = s
    return satirlar


# ---------------------------------------------------------------- 2) sınavlara indir
def main():
    satirlar = satirlari_oku()
    hocalar = Hocalar()
    for s in satirlar:
        ad, kod2 = ders_adi_temizle(s["ad"])
        s["ad_temiz"], s["kod"] = ad, s["kod"] or kod2
        s["anahtar"] = ad_anahtar(ad)
        s["hoca_id"] = hocalar.anahtar(s["hoca"])

    dersler = [s for s in satirlar if not s["yer_tutucu"]]
    # yer tutucu satırlara yazılmış talepler ("Küçük İşletme Yönetimi dersini ... alıyor")
    ekler = []
    for s in satirlar:
        if s["yer_tutucu"] and any(not bosmu(s[k]) for k in ("bolumler", "sayi", "talep")):
            hedef = sade(s["bolumler"] + " " + s["talep"])
            hedef = re.sub(r"[^0-9a-z]", "", hedef)
            aday = sorted({d["anahtar"] for d in dersler if len(d["anahtar"]) > 8 and d["anahtar"] in hedef},
                          key=len, reverse=True)
            if aday:
                s["anahtar"], s["ad_temiz"] = aday[0], next(d["ad_temiz"] for d in dersler if d["anahtar"] == aday[0])
                s["kod"] = s["kod"] if s["kod"] and not s["kod"].startswith("20B23") else None
                s["yer_tutucu_eki"] = True
                ekler.append(s)
    dersler += ekler

    # union-find: ders kodu + normalize ad
    ebeveyn = list(range(len(dersler)))
    def bul(x):
        while ebeveyn[x] != x:
            ebeveyn[x] = ebeveyn[ebeveyn[x]]; x = ebeveyn[x]
        return x
    ilk = {}
    for i, s in enumerate(dersler):
        for k in ([f"K:{s['kod']}"] if s["kod"] else []) + [f"N:{s['anahtar']}"]:
            if k in ilk:
                ebeveyn[bul(i)] = bul(ilk[k])
            else:
                ilk[k] = i
    gruplar = collections.defaultdict(list)
    for i, s in enumerate(dersler):
        gruplar[bul(i)].append(s)

    # farklı hoca = farklı sınav (aynı gün/saat/derslikte işlenen şubeler hariç)
    sinavlar = []
    for grup in gruplar.values():
        alt = collections.defaultdict(list)
        for s in grup:
            alt[s["hoca_id"]].append(s)
        hocasiz = alt.pop(None, [])
        parcalar = [v for v in alt.values()]
        birlesen = []
        while parcalar:
            p = parcalar.pop(0)
            imza = {(sade(x["gun"]), x["bas"], x["derslik"]) for x in p if x["derslik"] and sade(x["derslik"]) != "online"}
            for q in list(parcalar):
                if imza & {(sade(x["gun"]), x["bas"], x["derslik"]) for x in q}:
                    p = p + q; parcalar.remove(q)
                    p[0].setdefault("_hoca_uyari", True)
            birlesen.append(p)
        if not birlesen:
            birlesen = [hocasiz]
        else:
            birlesen[max(range(len(birlesen)), key=lambda i: len(birlesen[i]))] += hocasiz
        for p in birlesen:
            sinavlar.append({"satirlar": p, "bolunmus": len(birlesen) > 1})

    # sistem sayıları
    sistem = []
    ws = openpyxl.load_workbook(SISTEM, data_only=True).active
    PROG = [("SİYASET BİLİMİ VE ULUSLARARASI İLİŞKİLER (İNGİLİZCE)", "PSIR"), ("SİYASET", "SBUİ"),
            ("İŞLETME (İNGİLİZCE)", "BUS"), ("İŞLETME", "İŞL"), ("EKONOMİ", "ECF"),
            ("ULUSLARARASI TİCARET", "UTF"), ("LOJİSTİK", "LOJ"), ("HAVACILIK", "HAV"),
            ("BİLİŞİM", "YBS"), ("ELEKTRONİK TİCARET", "ETİC")]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[1]:
            continue
        prog = next(k for p, k in PROG if p in tr_upper(r[3]))
        sistem.append({"kod": kod_temizle(r[1]), "ad": metin(r[2]), "anahtar": ad_anahtar(ders_adi_temizle(r[2])[0]),
                       "bolum": prog, "sayi": int(str(r[5]).split("/")[0]), "hoca": metin(r[4]), "kullanildi": False})

    # ---------------------------------------------------------- 3) sınav alanlarını doldur
    ciktilar = []
    for sv in sinavlar:
        P = sv["satirlar"]
        notlar, uyari = [], False
        blok = [s for s in P if not s["havuz"]]
        kohort = {(s["bolum"], s["yy"]) for s in P if not s["havuz"]}
        kohort |= {(s["bolum"], None) for s in P if s["havuz"]}
        ana_kohort = sorted({f"{s['bolum']}-{s['yy']}" for s in P if not s["havuz"]})  # arka arkaya sınav kuralı için
        ingilizce = any(b in INGILIZCE for b, _ in kohort)
        yylar = sorted({s["yy"] for s in blok})
        # bölüm listesinden eksik kalan bölümler
        listelenen = set()
        for s in P:
            if s["bolumler"]:
                listelenen |= bolum_listesi(s["bolumler"], ingilizce)
        mevcut = {b for b, _ in kohort}
        oneriler = []
        ad_tahmin = collections.Counter(s["ad_temiz"] for s in (blok or P)).most_common(1)[0][0]
        for b in sorted(listelenen - mevcut, key=BOLUM_SIRA.index):
            yy = yylar[0] if yylar else None
            if ONAYLI_BOLUM_EKLERI == "hepsi" or (ad_tahmin, b) in ONAYLI_BOLUM_EKLERI:
                kohort.add((b, yy))
                notlar.append(f"{b}-{yy} hocanın 'dersi alan bölümler' bilgisinden eklendi (onaylı)")
            else:
                oneriler.append(f"{b}-{yy}")
                notlar.append(f"ONAY BEKLİYOR: hoca {b} bölümünü de yazmış, {b}-{yy} çakışma grubuna eklenmedi")
                uyari = True
        anahtar0 = ad_anahtar(ad_tahmin)
        if anahtar0 in TUM_BOLUMLER_ALIR:
            yy = TUM_BOLUMLER_ALIR[anahtar0]
            kohort |= {(b, yy) for b in TUM_BOLUM_KODLARI}
            notlar.append(f"bütün bölümler alıyor kabul edildi (tüm bölümlerin {yy}. yarıyılı)")
        mevcut = {b for b, _ in kohort}
        kohort_str = sorted(
            {f"{b}({HAVUZ_YARIYIL[b]}. yy)" if b.startswith("HAVUZ") else f"{b}-{yy}" for b, yy in kohort},
            key=lambda k: (BOLUM_SIRA.index(re.split(r"[-(]", k)[0]), k))

        # ad, kod, hoca
        ad_say = collections.Counter(s["ad_temiz"] for s in (blok or P))
        ad = max(ad_say, key=lambda a: (ad_say[a], len(a)))
        kodlar = []
        for s in sorted(P, key=lambda s: s["satir"]):
            if s["kod"] and s["kod"] not in kodlar:
                kodlar.append(s["kod"])
        hoca_idler = [s["hoca_id"] for s in P if s["hoca_id"] is not None]
        hoca_id = collections.Counter(hoca_idler).most_common(1)[0][0] if hoca_idler else None
        hoca = hocalar.kanonik(hoca_id)
        if sv["bolunmus"]:
            ad = f"{ad} ({hoca})"
            notlar.append("aynı ad farklı öğretim elemanı → ayrı sınav")
        if len(set(hoca_idler)) > 1:
            digerleri = sorted({hocalar.kanonik(h) for h in hoca_idler} - {hoca})
            notlar.append(f"aynı saat/derslikteki şubede farklı hoca yazılmış ({', '.join(digerleri)}) → tek sınav sayıldı (kullanıcı onayı)")
        if not hoca:
            notlar.append("öğretim elemanı yok"); uyari = True

        # sınav var mı
        sinav_var = "Evet"
        ortak = any(s.get("ortak") for s in P)
        if ortak and not ingilizce:
            sinav_var = "Hayır"
            notlar.append("TR ortak ders: sınavı Ortak Dersler Bölümü programlıyor (dosya notu)")
        talepler = [s for s in P if any(not bosmu(s[k]) for k in ("sayi", "izin", "talep", "gozetmen", "sure", "ek", "bolumler"))]
        for s in P:
            if re.search(r"yapılmayacak|ödev|odev", tr_lower(s["talep"] + " " + s["ek"])):
                sinav_var = "Hayır"; notlar.append("talepte sınav yapılmayacak yazıyor")

        # öğrenci sayısı (talep > sistem)
        sayilar = []
        for s in talepler:
            v = s["sayi"]
            if isinstance(v, (int, float)):
                sayilar.append(int(v))
            elif not bosmu(v):
                m = re.search(r"(\d+)\s*öğrenci", str(v))
                if m: sayilar.append(int(m.group(1)))
        eslesen = [x for x in sistem if (x["kod"] and x["kod"] in kodlar) or x["anahtar"] == ad_anahtar(ad.split(" (")[0])
                   or any(x["anahtar"] == s["anahtar"] or x["anahtar"] == SISTEM_AD_ESI.get(s["anahtar"]) for s in P)
                   or (len(x["anahtar"]) > 10 and any(levenshtein(x["anahtar"], s["anahtar"]) <= 2 for s in P))]
        ic = [x for x in eslesen if x["bolum"] in mevcut | listelenen]
        dis = [x for x in eslesen if x["bolum"] not in mevcut | listelenen and not any(b.startswith("HAVUZ") for b in mevcut)]
        havuz_dis = [x for x in eslesen if x["bolum"] not in mevcut | listelenen and any(b.startswith("HAVUZ") for b in mevcut)]
        if dis:  # sayıya dahil edilir; çakışma grubuna eklenmez (kullanıcıya soruldu)
            ic += dis
        for x in eslesen:
            x["kullanildi"] = True
        sistem_toplam = sum(x["sayi"] for x in ic + havuz_dis)
        if sayilar:
            if len(sayilar) > 1 and len(set(sayilar)) == 1:
                ogr, kaynak = sayilar[0], "Talep"
                notlar.append(f"aynı sayı {len(sayilar)} satırda tekrar ediyor, bir kez sayıldı")
            elif len(sayilar) > 1:
                ogr, kaynak = sum(sayilar), f"Talep ({'+'.join(map(str, sayilar))})"
            else:
                ogr, kaynak = sayilar[0], "Talep"
        elif eslesen:
            ogr = sistem_toplam
            parca = [f"{x['bolum']} {x['sayi']}" for x in ic + havuz_dis]
            kaynak = "Sistem" + (f" ({' + '.join(parca)})" if len(parca) > 1 else "")
        else:
            ogr, kaynak = None, "Eksik"
            if sinav_var == "Evet":
                notlar.append("öğrenci sayısı yok (ne talepte ne sistemde)"); uyari = True
        sistem_ek = sorted({x["bolum"] for x in dis + havuz_dis} - mevcut, key=BOLUM_SIRA.index)
        if sistem_ek and SISTEM_BOLUM_EKLE and sinav_var == "Evet":
            yy = yylar[0] if yylar else (6 if "HAVUZ2" in mevcut else 8)
            eklenen = []
            for b in sistem_ek:
                kohort.add((b, yy)); eklenen.append(f"{b}-{yy}")
                kacar = sum(x["sayi"] for x in dis + havuz_dis if x["bolum"] == b)
                oneriler.append(f"{b}-{yy} ({kacar} öğr.)")
            notlar.append("sistemde bu dersi alan başka bölüm öğrencisi var → çakışma grubuna eklendi: " +
                          ", ".join(f"{e} ({sum(x['sayi'] for x in dis + havuz_dis if x['bolum'] == e.split('-')[0])} öğr.)"
                                    for e in eklenen))
            kohort_str = sorted(
                {f"{b}({HAVUZ_YARIYIL[b]}. yy)" if b.startswith("HAVUZ") else f"{b}-{yy2}" for b, yy2 in kohort},
                key=lambda k: (BOLUM_SIRA.index(re.split(r"[-(]", k)[0]), k))
        if ogr == 0 and sinav_var == "Evet":
            sinav_var = "Hayır"; notlar.append("öğrenci sayısı 0")

        # bölüm bazında öğrenci sayısı (bölünen sınavda hangi bölüm hangi derslikte)
        bolum_sayilari, bolum_kaynak = collections.Counter(), ""
        talep_parca = [(s["bolum"], int(s["sayi"])) for s in talepler
                       if not s["havuz"] and isinstance(s["sayi"], (int, float)) and s["sayi"] > 0]
        if len(talep_parca) > 1 and len({v for _, v in talep_parca}) > 1:
            for b, v in talep_parca: bolum_sayilari[b] += v
            bolum_kaynak = "talep"
        elif ic + havuz_dis:
            for x in ic + havuz_dis: bolum_sayilari[x["bolum"]] += x["sayi"]
            bolum_kaynak = "sistem"
        if ogr:
            if not bolum_sayilari:
                bolumler_ = sorted({b for b, _ in kohort if not b.startswith("HAVUZ")} & {s["bolum"] for s in blok}) or \
                            sorted({b for b, _ in kohort if not b.startswith("HAVUZ")})
                bolum_sayilari = collections.Counter({b: ogr / len(bolumler_) for b in bolumler_}) if bolumler_ else collections.Counter()
                bolum_kaynak = "eşit bölüştürme (tahmin)" if len(bolumler_) > 1 else "tek bölüm"
            toplam_b = sum(bolum_sayilari.values())
            if toplam_b and toplam_b != ogr:  # talep toplamına ölçekle (en büyük kalan yöntemi)
                ham = {b: v * ogr / toplam_b for b, v in bolum_sayilari.items()}
                tam = {b: int(v) for b, v in ham.items()}
                for b in sorted(ham, key=lambda b: ham[b] - tam[b], reverse=True)[:ogr - sum(tam.values())]:
                    tam[b] += 1
                bolum_sayilari = collections.Counter(tam)
        bolum_sayilari = {b: int(v) for b, v in sorted(bolum_sayilari.items(), key=lambda kv: BOLUM_SIRA.index(kv[0])) if v}

        # izin, gün/saat talebi, süre, gözetmen
        izin = sorted({g for s in talepler for g in gunleri_bul(s["izin"])}, key=GUNLER.index)
        gunler, saatler, haric, tarihler, ders_saatinde = set(), set(), set(), [], False
        talep_metinleri = []
        for s in talepler:
            kaynak_metin = s["talep"]
            if bosmu(kaynak_metin) and re.search(tr_lower(AY_ADI) + r"|pazartesi|salı|çarşamba|perşembe|cuma", tr_lower(s["bolumler"])):
                kaynak_metin = s["bolumler"]  # talep yanlış sütuna yazılmış
            p = talep_ayristir(kaynak_metin, s["gun"], s["bas"], s["bit"])
            gunler |= set(p["gunler"]); saatler |= set(p["saatler"]); haric |= set(p["haric"])
            tarihler += p["tarih"]; ders_saatinde |= p["ders_saatinde"]
        sureler = [sure_dk(s["sure"]) for s in talepler if sure_dk(s["sure"])]
        sure = max(sureler) if sureler else None
        if sure and sure <= SURE_YUVARLA_UST:
            sure = DERSLIK_SAAT_SINIRI
        if ad in SURE_ISTISNA:
            sure = SURE_ISTISNA[ad]
            notlar.append(f"süre {sure} dk (kullanıcının sonradan bildirdiği değişiklik)")
        gz = [gozetmen(s["gozetmen"]) for s in talepler if gozetmen(s["gozetmen"])]
        hoca_gozet = "Var" if "Var" in gz else ("Yok" if gz else None)
        if ogr is None:
            gozet = hoca_gozet
        else:
            gozet = "Var" if ogr >= GOZETMEN_ESIGI else "Yok"
            if hoca_gozet == "Var" and gozet == "Yok" and sinav_var == "Evet" and \
                    any(k in sade(hoca).upper() for k in GOZETMEN_ONAYLI_HOCA):
                gozet = "Var"
                notlar.append(f"{ogr} öğrenci ama hoca istediği için gözetmen verilecek (kullanıcı onayı)")
            elif hoca_gozet == "Var" and gozet == "Yok" and sinav_var == "Evet":
                gozet = "SORULACAK"
                notlar.append(f"SORU: hoca gözetmen 'Var' dedi ama {ogr} öğrenci (<{GOZETMEN_ESIGI})"); uyari = True
            elif hoca_gozet and hoca_gozet != gozet and sinav_var == "Evet":
                notlar.append(f"hoca gözetmen '{hoca_gozet}' demişti; {ogr} öğrenci → kural gereği '{gozet}'")
        if gozet == "Yok" and sinav_var == "Evet":
            notlar.append("dersin hocası girer")
        if not talepler and sinav_var == "Evet":
            notlar.append("hoca talep göndermemiş")

        if set(gunler) & set(izin):
            notlar.append(f"gün talebi ({', '.join(sorted(set(gunler) & set(izin), key=GUNLER.index))}) izin günüyle çakışıyor")
            uyari = True
        if ders_saatinde:
            notlar.append("talep: ders gün/saatinde")
        if tarihler:
            notlar.append("tarihli talep: " + ", ".join(sorted(set(tarihler))))
        if sure and sure > DERSLIK_SAAT_SINIRI:
            notlar.append(f"süre {sure} dk > 50 dk → derslik 2 oturum kapatılır")

        derslik = ""
        for s in talepler:
            ek = sade(s["ek"] + " " + s["bolumler"] + " " + s["talep"])
            if "kutuphane lab" in ek:
                derslik = "Kütüphane Lab"
                notlar.append("hocanın istediği Kütüphane Lab kullanılacak (kullanıcı onayı)")
            elif re.search(r"\blab\b", ek) and not re.search(r"gerek yok|sinifta", ek):
                derslik = "LAB"
            m = re.search(r"-\s*(\d{1,2}[a-z]\d{1,2})\b", ek)
            if m and not derslik:
                derslik = m.group(1).upper()
                notlar.append(f"hocanın istediği {derslik} dersliği (kullanıcı onayı)")
            if "7. kat" in tr_lower(s["talep"]):
                derslik = "7. kat (hoca belirler)"
                notlar.append("7. katta hocanın istediği derslik kullanılacak (kullanıcı onayı)")
        if derslik == "LAB" and ogr and ogr > LAB_KAPASITE:
            if anahtar0 in LAB_TASMA_ONAYLI:
                notlar.append(f"{ogr} öğrenci tek LAB'a sığar (kullanıcı onayı)")
            else:
                notlar.append(f"LAB kapasitesi {LAB_KAPASITE}, öğrenci {ogr}"); uyari = True
        for s in talepler:
            tt = tr_lower(s["talep"] + " " + s["bolumler"] + " " + s["gozetmen"] + " " + s["ek"])
            if "aynı saat" in tt:
                notlar.append("hoca: karşılık gelen TR/İng dersle aynı saatte ve yakın derslikte olsun")
            if "büyük sınıf" in tt: notlar.append("hoca: büyük sınıf istiyor")
            if "2 sınıfa" in tt: notlar.append("hoca: 2 dersliğe bölünsün")
            if "üstten" in tt: notlar.append("hoca: dersi üstten alan öğrenciler var, perşembe sabahı çakışmasız olsun")
            if "gözükmüyor" in tt: notlar.append("hoca: SBUİ İng programında görünmüyor")

        ham = []
        for s in talepler:
            parca = [f"{k}: {s[a]}" for k, a in (("Bölümler", "bolumler"), ("Sayı", "sayi"), ("İzin", "izin"),
                                                 ("Talep", "talep"), ("Gözetmen", "gozetmen"), ("Süre", "sure"),
                                                 ("Ek", "ek")) if not bosmu(s[a])]
            satir = f"[{s['bolum']}{'' if s['havuz'] else '-' + str(s['yy'])}] " + "; ".join(parca)
            if satir not in ham:
                ham.append(satir)

        ciktilar.append({
            "sira": min(s["satir"] for s in (blok or P)),
            "Ders Kodu": " / ".join(kodlar),
            "Ders": ad,
            "Program-Yarıyıl (Çakışma Grubu)": "; ".join(kohort_str),
            "Öğrenci Sayısı": ogr,
            "Sayı Kaynağı": kaynak,
            "Derslik": derslik,
            "Sınav Var mı?": sinav_var,
            "İzin Günü (sınav konulamaz)": ", ".join(izin),
            "Gün Talebi": ", ".join(sorted(gunler, key=GUNLER.index)),
            "Saat Talebi": ", ".join(f"{h:02d}:00" for h in sorted(saatler)) +
                           (f" (hariç {', '.join(f'{h:02d}:00' for h in sorted(haric))})" if haric else ""),
            "Süre": f"{sure} dk" if sure else "",
            "Gözetmen İhtiyacı": gozet or "",
            "Öğretim Elemanı": hoca,
            "Not (otomatik)": "; ".join(dict.fromkeys(notlar)),
            "Talep (kaynak metin)": "\n".join(ham),
            "_uyari": uyari, "_oneriler": oneriler, "_izin": izin, "_sinav_var": sinav_var, "_ana_kohort": ana_kohort,
            "_derslik_sabit": bool(derslik and derslik != "LAB"),
            "_ders_slot": sorted({(g, h) for s in P if s["bas"] is not None and s["bit"]
                                  for g in gunleri_bul(s["gun"]) for h in range(s["bas"], s["bit"]) if 9 <= h <= 17},
                                 key=lambda t: (GUNLER.index(t[0]), t[1])),
            "_hoca_id": hoca_id, "_gunler": sorted(gunler, key=GUNLER.index), "_saatler": sorted(saatler),
            "_tarihli": bool(tarihler), "_ayni_saat": any("aynı saat" in tr_lower(s["talep"] + " " + s["bolumler"]) for s in talepler),
            "_yy": yylar, "_ing": ingilizce,
            "_kaynak_satirlar": sorted(s["satir"] for s in P),
            "_satir_bolum": {str(s["satir"]): s["bolum"] for s in P},
            "_bolum_sayilari": bolum_sayilari, "_bolum_kaynak": bolum_kaynak,
        })

    ciktilar.sort(key=lambda c: c["sira"])
    # "aynı saat ve yakın derslik" çiftleri: aynı hoca, biri TR biri İng, aynı yarıyıl
    for a in ciktilar:
        if not a["_ayni_saat"]:
            continue
        for b in ciktilar:
            if b is not a and b["_ayni_saat"] and b["_hoca_id"] == a["_hoca_id"] and b["_ing"] != a["_ing"] and b["_yy"] == a["_yy"]:
                a["_ayni_saat_es"] = b["Ders"]
                a["Not (otomatik)"] = a["Not (otomatik)"].replace(
                    "hoca: karşılık gelen TR/İng dersle aynı saatte ve yakın derslikte olsun",
                    f"hoca: '{b['Ders']}' ile AYNI SAATTE ve yakın derslikte olsun")
    # aynı hocanın aynı tarih/saate istediği sınavlar
    for a in ciktilar:
        if not a["_tarihli"]:
            continue
        es = [b["Ders"] for b in ciktilar if b is not a and b["_tarihli"] and b["_hoca_id"] == a["_hoca_id"]
              and b["_gunler"] == a["_gunler"] and b["_saatler"] == a["_saatler"]]
        if es:
            a["Not (otomatik)"] += f"; aynı hoca aynı saatte '{', '.join(es)}' sınavını da yapacak (aynı derslikte birlikte, kullanıcı onayı)"
            a["_birlikte"] = es
    # kullanıcının "tek derslikte birlikte" dediği sınavlar
    for k1, k2 in BIRLIKTE_SINAVLAR:
        a = next((c for c in ciktilar if ad_anahtar(c["Ders"]) == k1), None)
        b = next((c for c in ciktilar if ad_anahtar(c["Ders"]) == k2), None)
        if not (a and b):
            print("UYARI: birlikte sınav çifti bulunamadı:", k1, k2)
            continue
        toplam = (a["Öğrenci Sayısı"] or 0) + (b["Öğrenci Sayısı"] or 0)
        for x, y in ((a, b), (b, a)):
            x["_birlikte"] = [y["Ders"]]
            x["Not (otomatik)"] = "; ".join(filter(None, [x["Not (otomatik)"],
                f"'{y['Ders']}' ile aynı saatte aynı derslikte birlikte (kullanıcı onayı; toplam {toplam} öğrenci)"]))
            if toplam >= GOZETMEN_ESIGI and x["Gözetmen İhtiyacı"] != "Var":
                x["Gözetmen İhtiyacı"] = "Var"
                x["Not (otomatik)"] = x["Not (otomatik)"].replace("; dersin hocası girer", "") + \
                    f"; ortak derslikte {toplam} öğrenci → gözetmen"
    # akademik izin günü hocaya aittir: bir dersinde yazdıysa bütün sınavlarına uygula
    hoca_izin = collections.defaultdict(set)
    for c in ciktilar:
        if c["_hoca_id"] is not None:
            hoca_izin[c["_hoca_id"]].update(c["_izin"])
    for c in ciktilar:
        tum = sorted(hoca_izin.get(c["_hoca_id"], set()), key=GUNLER.index)
        if tum and tum != c["_izin"]:
            eksik = [g for g in tum if g not in c["_izin"]]
            c["İzin Günü (sınav konulamaz)"] = ", ".join(tum)
            c["Not (otomatik)"] = "; ".join(filter(None, [c["Not (otomatik)"],
                f"izin günü ({', '.join(eksik)}) hocanın diğer ders talebinden alındı"]))
            c["_izin"] = tum
            cakisan = set(c["_gunler"]) & set(tum)
            if cakisan:
                c["Not (otomatik)"] += f"; gün talebi izin günüyle çakışıyor ({', '.join(sorted(cakisan, key=GUNLER.index))})"
                c["_uyari"] = True
    # izin günü bilinmeyen hoca: dersin normal gün/saati tercih edilir
    if IZINSIZ_HOCA_DERS_SAATI:
        for c in ciktilar:
            if c["_sinav_var"] != "Evet" or c["_izin"] or c["_gunler"] or c["_derslik_sabit"] or not c["_ders_slot"]:
                continue
            gunler = sorted({g for g, _ in c["_ders_slot"]}, key=GUNLER.index)
            c["Gün Talebi"] = ", ".join(gunler)
            if len(gunler) == 1:
                c["Saat Talebi"] = ", ".join(f"{h:02d}:00" for _, h in c["_ders_slot"])
            else:
                c["Saat Talebi"] = " | ".join(
                    f"{g}: " + ", ".join(f"{h:02d}:00" for g2, h in c["_ders_slot"] if g2 == g) for g in gunler)
            c["_gunler"], c["_saatler"] = gunler, sorted({h for _, h in c["_ders_slot"]})
            c["_tercih_ders_saati"] = True
            c["Not (otomatik)"] = "; ".join(filter(None, [c["Not (otomatik)"],
                "izin günü bilinmiyor → dersin normal gün/saatine konmaya çalışılacak (tercih)"]))
    kullanilmayan = [x for x in sistem if not x["kullanildi"]]
    return ciktilar, kullanilmayan, hocalar


def yaz(ciktilar):
    wb = openpyxl.load_workbook(SABLON)
    ws = wb["Dersler"]
    basliklar = [c.value for c in ws[1]]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    ince = Side(style="thin", color="FFBFBFBF")
    SARI = PatternFill("solid", fgColor="FFFFF2CC")
    GRI = PatternFill("solid", fgColor="FFEDEDED")
    for i, c in enumerate(ciktilar, start=2):
        for j, b in enumerate(basliklar, start=1):
            hucre = ws.cell(i, j, c.get(b) if c.get(b) not in ("", None) else None)
            hucre.alignment = Alignment(wrap_text=True, vertical="top")
            hucre.border = Border(left=ince, right=ince, top=ince, bottom=ince)
            if c["Sınav Var mı?"] == "Hayır":
                hucre.fill = GRI
            elif c["_uyari"]:
                hucre.fill = SARI
    son = openpyxl.utils.get_column_letter(len(basliklar))
    ws.auto_filter.ref = f"A1:{son}{len(ciktilar) + 1}"
    ws.freeze_panes = "C2"
    wb.save(SABLON)


if __name__ == "__main__":
    ciktilar, kullanilmayan, hocalar = main()
    with open(JSON_CIKTI, "w", encoding="utf-8") as f:
        json.dump({"sinavlar": ciktilar,
                   "sistemde_olup_programda_olmayan": kullanilmayan}, f, ensure_ascii=False, indent=1, default=str)
    if "--yaz" in sys.argv:
        yaz(ciktilar)
        print("Dersler sekmesi yazıldı:", len(ciktilar), "satır")
    else:
        print("Kuru çalıştırma:", len(ciktilar), "sınav (yazmak için --yaz)")
