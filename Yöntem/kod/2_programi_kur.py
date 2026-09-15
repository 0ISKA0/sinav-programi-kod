# -*- coding: utf-8 -*-
"""
Adım 2 — Sınav programını kurar, gözetmen atar, bağımsız kontrol yapar, şablona yazar.

Girdi : Yöntem/kod/dersler.json                      (1_dersleri_yukle.py çıktısı)
        Şablon: "Derslik Kısıtlamaları" (Müsait mi?), "Gözetmenler"
Çıktı : Şablon: "Sınav Programı", "Derslik Kısıtlamaları" (dolu hücreler), "Gözetmenler" (yük sütunları)
        Yöntem/kod/program.json, Yöntem/kod/rapor.txt

Kullanım: python 2_programi_kur.py [--yaz] [--deneme N] [--tohum S]
"""
import collections, itertools, json, os, random, re, sys
import openpyxl

try:
    from donem_ayarlari import *   # dönem bilgileri ve şifreler (GitHub'a yüklenmez)
except ModuleNotFoundError:
    raise SystemExit("donem_ayarlari.py bulunamadı: 'donem_ayarlari_ornek.py' dosyasını 'donem_ayarlari.py' adıyla kopyalayıp doldurun.")
from openpyxl.styles import Alignment, Border, PatternFill, Side

KOK = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SABLON = os.path.join(KOK, SABLON_DOSYASI)
DERSLER = os.path.join(KOK, "Yöntem", "kod", "dersler.json")
PROGRAM_JSON = os.path.join(KOK, "Yöntem", "kod", "program.json")
RAPOR = os.path.join(KOK, "Yöntem", "kod", "rapor.txt")

GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
TARIH = dict(SINAV_TARIHLERI)
SAATLER = list(range(9, 18))

# ---------------------------------------------------------------- genel kurallar (Genel Kısıtlar sekmesi)
BUYUK_SINAV = 60            # > 60 öğrenci -> 3. kat tercih
GOZETMEN_ESIGI = 30
KUCUK_TEK_DERSLIK = ("6D4", "6D5")   # hocanın gireceği 25-29 kişilik sınavlar
ARDISIK_SINAV_SINIRI = 2    # bir program-yarıyıl grubu için arka arkaya en fazla
KAPASITE_TOLERANS = 6       # derslik kapasiteleri yaklaşık: 6 kişiye kadar aşılabilir (cezalı)
GOZETMEN_ARDISIK_SINIRI = 2

# dönem kuralları (HOCA_EN_GEC_BASLANGIC, TEK_DERSLIK_TERCIH, GOZETMEN_… vb.) donem_ayarlari.py içinde

# ---------------------------------------------------------------- tercih ağırlıkları
W_TALEP_TAM, W_TALEP_GUN, W_TALEP_SAAT = 60, 45, 15     # hocanın açık talebi en güçlü tercih
W_DERS_SAATI_TAM, W_DERS_SAATI_GUN = 25, 8               # izin günü bilinmeyen hoca: dersin gün/saati
W_UST_KAT = 90                                           # >60 öğrenci: 3. kat hoca talebinden önce gelir (kullanıcı kararı)
W_KAT_KARISIK = 10                                       # 3. + 6. kat karışık bölünme cezası
W_ALT_KAT_BUYUK = 30                                     # >60 öğrenci tamamen 6. katta: ek ceza
W_GUN_YUKU, W_AYNI_GUN_KOHORT = 0.3, 8
W_DERSLIK_SAYISI = 4                                     # her ek derslik = ek gözetmen: az derslik tercih
W_KAPASITE_ASIMI = 0.4                                   # kapasite üstü her öğrenci için                   # günlere denge; öğrencinin aynı gün ikinci sınavı (oransal)


# ================================================================ veri
def kat(oda):
    if oda.startswith("6"): return 6
    if oda.startswith("3"): return 3
    return oda  # LAB / dış derslik


def derslikleri_oku(wb):
    ws = wb["Derslik Kısıtlamaları"]
    h = [c.value for c in ws[1]]
    I = {k: i for i, k in enumerate(h)}
    kap, musait = {}, set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        oda = r[I["Sınıf"]]
        kap.setdefault(oda, r[I["Kapasite"]] or 40)
        gun = r[I["Tarih"]].split()[1]
        saat = int(str(r[I["Sınav Başlangıç Saati"]])[:2])
        m = str(r[I["Müsait mi?"]] or "müsait").strip().lower()
        if m.startswith("müsait") or m in ("evet", "e"):
            musait.add((oda, gun, saat))
    return kap, musait


def gozetmenleri_oku(wb):
    ws = wb["Gözetmenler"]
    h = [c.value for c in ws[1]]
    I = {k: i for i, k in enumerate(h)}
    kisiler = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        ad = r[I["Gözetmen"]]
        if not ad or str(ad).startswith("Gözetmen kuralları"):
            continue
        kisit = str(r[I["Kısıt"]] or "")
        yasak = set()
        if re.search(r"olmas[ıi]n", kisit.lower()):
            yasak = {int(x) for x in re.findall(r"\d{1,2}", kisit) if 9 <= int(x) <= 17}
        izin_metni = str(r[I["İzin Günü"]] or "").lower()
        izin = set(GUNLER) if "tüm" in izin_metni or "tum" in izin_metni else {g for g in GUNLER if g.lower() in izin_metni}
        cins = str(r[I["Cinsiyet"]] or "")[:1].upper()
        kisiler.append({"ad": ad, "cins": cins, "izin": izin, "yasak_saat": yasak})
    return kisiler


def sinavlari_oku():
    ham = json.load(open(DERSLER, encoding="utf-8"))["sinavlar"]
    sinavlar = []
    for c in ham:
        if c["Sınav Var mı?"] != "Evet":
            continue
        kohort = [k for k in c["Program-Yarıyıl (Çakışma Grubu)"].split("; ") if k]
        sure = int(c["Süre"].split()[0]) if c["Süre"] else 50
        s = {
            "ad": c["Ders"], "kod": c["Ders Kodu"], "kohort": set(kohort), "kohort_str": c["Program-Yarıyıl (Çakışma Grubu)"],
            "ana": set(c["_ana_kohort"]), "hoca": c["Öğretim Elemanı"], "hoca_id": c["_hoca_id"],
            "n": c["Öğrenci Sayısı"] or 0, "sure": sure, "dur": 2 if sure > 50 else 1,
            "gozetmen": c["Gözetmen İhtiyacı"] == "Var", "izin": set(c["_izin"]),
            "derslik": c["Derslik"] or "", "ing": any(k.startswith(("PSIR", "BUS", "ECF")) for k in kohort),
            "birlikte": c.get("_birlikte") or [], "es": c.get("_ayni_saat_es"),
            "bolum_sayilari": c.get("_bolum_sayilari") or {},
            "pref": {}, "pref_gun": {}, "sert": None, "tercih_turu": None,
        }
        gunler, saatler = c["_gunler"], c["_saatler"]
        if c.get("_tercih_ders_saati"):
            s["tercih_turu"] = "ders saati"
            for g, h in c["_ders_slot"]:
                s["pref"][(g, h)] = W_DERS_SAATI_TAM
            for g in {g for g, _ in c["_ders_slot"]}:
                s["pref_gun"][g] = W_DERS_SAATI_GUN
        elif gunler or saatler:
            s["tercih_turu"] = "talep"
            for g in (gunler or GUNLER):
                for h in (saatler or SAATLER):
                    s["pref"][(g, h)] = W_TALEP_TAM if (gunler and saatler) else (W_TALEP_GUN if gunler else W_TALEP_SAAT)
            if gunler and saatler:
                for g in gunler:
                    s["pref_gun"][g] = W_TALEP_GUN
        if (c.get("_tarihli") or s["derslik"].startswith(SERT_TALEPLI_DERSLIK)) and gunler and saatler:
            s["sert"] = {(g, h) for g in gunler for h in saatler}
        s["en_gec"] = min([v for k, v in HOCA_EN_GEC_BASLANGIC.items() if k in s["hoca"].upper()] or [17])
        sinavlar.append(s)
    return sinavlar


def birimleri_kur(sinavlar):
    """Birlikte (aynı derslik) ve eş (aynı saat) sınavları tek yerleştirme birimi yapar."""
    ada = {s["ad"]: s for s in sinavlar}
    ebeveyn = {s["ad"]: s["ad"] for s in sinavlar}
    def bul(x):
        while ebeveyn[x] != x:
            x = ebeveyn[x]
        return x
    for s in sinavlar:
        for d in s["birlikte"] + ([s["es"]] if s["es"] else []):
            if d in ada:
                ebeveyn[bul(s["ad"])] = bul(d)
    gruplar = collections.defaultdict(list)
    for s in sinavlar:
        gruplar[bul(s["ad"])].append(s)
    birimler = []
    for i, uyeler in enumerate(gruplar.values()):
        # derslik grupları: birlikte olanlar aynı derslik grubunda
        dg = {}
        for s in uyeler:
            anahtar = min([s["ad"]] + [d for d in s["birlikte"] if d in ada])
            dg.setdefault(anahtar, []).append(s)
        izin = set().union(*(s["izin"] for s in uyeler))
        serts = [s["sert"] for s in uyeler if s["sert"]]
        sert = set.intersection(*serts) if serts else None
        birimler.append({
            "id": i, "uyeler": uyeler, "derslik_gruplari": list(dg.values()),
            "kohort": set().union(*(s["kohort"] for s in uyeler)),
            "ana": set().union(*(s["ana"] for s in uyeler)),
            "hocalar": {s["hoca_id"] for s in uyeler if s["hoca_id"] is not None},
            "dur": max(s["dur"] for s in uyeler), "izin": izin, "sert": sert,
            "en_gec": min(s["en_gec"] for s in uyeler),
            "n": sum(s["n"] for s in uyeler),
        })
    return birimler


# ================================================================ derslik seçimi
ALT_KAT = ["6A1", "6A2", "6D4", "6D5", "6D1", "6D3"]
UST_KAT = ["3A1", "3A3", "3A5", "3A6", "3A4"]


def en_iyi_kume(adaylar, kap, n, tolerans=KAPASITE_TOLERANS):
    """Kapasitesi (toleransla) n'i karşılayan en az derslikli küme; eşitlikte aşımı ve boşluğu en az olan."""
    for k in range(1, len(adaylar) + 1):
        en = None
        for kume in itertools.combinations(adaylar, k):
            t = sum(kap[o] for o in kume)
            if t + tolerans * k < n:
                continue
            anahtar = (max(0, n - t), abs(t - n))
            if en is None or anahtar < en[0]:
                en = (anahtar, kume)
        if en:
            return list(en[1])
    return None


def bolum_dagit(dg, odalar, kap):
    """Bölünen sınavda bölümleri dersliklere atar (Genel Kısıt: her bölüm mümkünse tek derslikte).
    Amaç sırası: bölünen bölüm sayısı en az -> bölünenler en az dersliğe yayılsın -> sıkı yerleşim.
    -> {(ders, oda): {bölüm: sayı}}"""
    parcalar = [(s["ad"], b, v) for s in dg for b, v in (s["bolum_sayilari"] or {"?": s["n"]}).items()]
    parcalar.sort(key=lambda p: (-p[2], p[1]))
    kap0 = {o: kap[o] + (KAPASITE_TOLERANS if kap[o] < 999 else 0) for o in odalar}
    nominal = {o: kap[o] for o in odalar}
    en = {"bol": None, "atama": None}
    sayac = [0]

    KUCUK_PARCA = 5

    def bol_yerlestir(kalan, bekleyen):
        """Tek dersliğe konamayan parçaları en az dersliğe, küçük artık parça bırakmadan yayar."""
        atama, kalan = [], dict(kalan)
        oda_sayisi = kucuk = 0
        for ad, b, v in sorted(bekleyen, key=lambda p: -p[2]):
            bos = [o for o in odalar if kalan[o] > 0]
            en_k = None
            for k in range(1, len(bos) + 1):
                for kume in itertools.combinations(bos, k):
                    t = sum(kalan[o] for o in kume)
                    if t < v:
                        continue
                    kume = sorted(kume, key=lambda o: -kalan[o])
                    # önce normal kapasiteye kadar doldur, esnekliği (tolerans) yalnız gerekirse kullan
                    tol = {o: kap0[o] - nominal[o] for o in kume}
                    koyulan = {o: min(max(0, kalan[o] - tol[o]), 10 ** 6) for o in kume}
                    kal = v
                    for o in kume:
                        koyulan[o] = min(koyulan[o], kal); kal -= koyulan[o]
                    for o in kume:
                        if kal <= 0: break
                        ek = min(kalan[o] - koyulan[o], kal); koyulan[o] += ek; kal -= ek
                    parca = [(o, koyulan[o]) for o in kume if koyulan[o] > 0]
                    if kal > 0 or len(parca) < k:
                        continue
                    ceza = (k, sum(1 for _, x in parca if x < KUCUK_PARCA), t - v)
                    if en_k is None or ceza < en_k[0]:
                        en_k = (ceza, parca)
                if en_k:
                    break
            if en_k is None:  # kapasite toleransı (ör. LAB 42/40)
                o = max(odalar, key=lambda o: kap0[o]); atama.append((ad, o, b, v)); oda_sayisi += 1
                continue
            for o, koy in en_k[1]:
                atama.append((ad, o, b, koy)); kalan[o] -= koy
            oda_sayisi += len(en_k[1]); kucuk += en_k[0][1]
        return atama, (kucuk, oda_sayisi)

    def dfs(i, kalan, secim, bekleyen):
        sayac[0] += 1
        if en["bol"] is not None and len(bekleyen) > en["bol"][0] or sayac[0] > 50000:
            return
        if i == len(parcalar):
            ek, (kucuk, oda_sayisi) = bol_yerlestir(kalan, bekleyen)
            dolu = collections.Counter()
            for _, o, _, v in secim + ek:
                dolu[o] += v
            asim = sum(max(0, dolu[o] - nominal[o]) for o in odalar)
            puan = (len(bekleyen), kucuk, oda_sayisi, asim)
            if en["bol"] is None or puan < en["bol"]:
                en["bol"], en["atama"] = puan, secim + ek
            return
        ad, b, v = parcalar[i]
        for o in sorted([o for o in odalar if kalan[o] >= v], key=lambda o: kalan[o] - v):
            kalan[o] -= v
            dfs(i + 1, kalan, secim + [(ad, o, b, v)], bekleyen)
            kalan[o] += v
        dfs(i + 1, kalan, secim, bekleyen + [parcalar[i]])

    dfs(0, dict(kap0), [], [])
    atama = collections.defaultdict(collections.Counter)
    for ad, o, b, v in en["atama"]:
        atama[(ad, o)][b] += v
    return atama


def dagit(odalar, kap, n):
    """Öğrencileri dersliklere kapasite oranında dağıtır."""
    if len(odalar) == 1:
        return [(odalar[0], n)]
    top = sum(kap[o] for o in odalar)
    say = [int(n * kap[o] / top) for o in odalar]
    i = 0
    while sum(say) < n:
        if say[i % len(say)] < kap[odalar[i % len(say)]]:
            say[i % len(say)] += 1
        i += 1
    return list(zip(odalar, say))


class Durum:
    def __init__(self, kap, musait, arz):
        self.kap, self.musait, self.arz = kap, musait, arz
        self.goz = collections.Counter()  # (gün, saat) -> gereken gözetmen
        self.oda = set()                  # (oda, gün, saat)
        self.kohort = set()               # (kohort, gün, saat)
        self.hoca = set()                 # (hoca_id, gün, saat)
        self.ana = collections.defaultdict(dict)   # (kohort, gün) -> {saat: birim_id}
        self.gun_yuk = collections.Counter()
        self.yer = {}                     # birim_id -> yerleşim

    def serbest(self, oda, g, span):
        dis = kat(oda) not in (6, 3) and oda != "LAB"
        return all(((oda, g, h) in self.musait or dis) and (oda, g, h) not in self.oda for h in span)


def derslik_grubu_ayir(dg, g, span, D, kullanilan):
    """Bir derslik grubu (birlikte yapılan sınavlar) için derslik seçer -> (odalar, ceza) ya da None."""
    n = sum(s["n"] for s in dg)
    goz = any(s["gozetmen"] for s in dg)
    sabit = next((s["derslik"] for s in dg if s["derslik"]), "")
    serbest = lambda o: o not in kullanilan and D.serbest(o, g, span)
    if sabit:
        return ([sabit], 0) if serbest(sabit) else None
    if any(s["ad"] in TEK_DERSLIK_TERCIH for s in dg):
        tek = sorted([o for o in ALT_KAT + UST_KAT if serbest(o) and D.kap[o] >= n], key=lambda o: D.kap[o])
        if tek:
            return [tek[0]], -8
    alt = [o for o in ALT_KAT if serbest(o)]
    ust = [o for o in UST_KAT if serbest(o)]
    if n > BUYUK_SINAV:
        k = en_iyi_kume(ust, D.kap, n)
        if k:
            return k, -W_UST_KAT + W_DERSLIK_SAYISI * (len(k) - 1) + W_KAPASITE_ASIMI * max(0, n - sum(D.kap[o] for o in k))
        if ust:
            ust_kap = sum(D.kap[o] + KAPASITE_TOLERANS for o in ust)
            k2 = en_iyi_kume(alt, D.kap, n - ust_kap)
            if k2:
                return ust + k2, -(W_UST_KAT * min(1, ust_kap / n) - W_KAT_KARISIK) + W_DERSLIK_SAYISI * (len(ust) + len(k2) - 1)
        k = en_iyi_kume(alt, D.kap, n)
        return (k, W_ALT_KAT_BUYUK + W_DERSLIK_SAYISI * len(k)) if k else None
    if not goz:  # hoca girer: tek derslik
        if n >= 25:
            aday = [o for o in KUCUK_TEK_DERSLIK if serbest(o)]
            if aday:
                return [aday[0]], 0
        aday = sorted([o for o in alt if D.kap[o] + KAPASITE_TOLERANS >= n], key=lambda o: (D.kap[o] < n, D.kap[o]))
        if aday:
            return [aday[0]], W_KAPASITE_ASIMI * max(0, n - D.kap[aday[0]])
        aday = [o for o in ust if D.kap[o] + KAPASITE_TOLERANS >= n]
        return ([min(aday, key=lambda o: D.kap[o])], 10) if aday else None
    k = en_iyi_kume(alt, D.kap, n)
    if k:
        return k, W_DERSLIK_SAYISI * (len(k) - 1) + W_KAPASITE_ASIMI * max(0, n - sum(D.kap[o] for o in k))
    aday = [o for o in ust if D.kap[o] >= n]
    if aday:
        return [min(aday, key=lambda o: D.kap[o])], 8
    if ust:
        kalan = n - sum(D.kap[o] for o in ust)
        k2 = en_iyi_kume(alt, D.kap, kalan)
        if k2:
            return ust + k2, 12
    return None


# ================================================================ yerleştirme
def izinli_mi(b, g, h):
    if g in b["izin"] or h + b["dur"] - 1 > 17 or h > b["en_gec"]:
        return False
    return b["sert"] is None or (g, h) in b["sert"]


def dene(b, g, h, D):
    """Birimi (g,h)'ye koymayı dener -> (puan, derslikler) ya da None. Durumu değiştirmez."""
    if not izinli_mi(b, g, h):
        return None
    span = list(range(h, h + b["dur"]))
    for hh in span:
        if any((k, g, hh) in D.kohort for k in b["kohort"]) or any((x, g, hh) in D.hoca for x in b["hocalar"]):
            return None
    for k in b["ana"]:
        saatler = dict(D.ana.get((k, g), {}))
        for hh in span:
            saatler[hh] = b["id"]
        sol, sag = min(span), max(span)
        while sol - 1 in saatler: sol -= 1
        while sag + 1 in saatler: sag += 1
        if len({saatler[x] for x in range(sol, sag + 1)}) > ARDISIK_SINAV_SINIRI:
            return None
    kullanilan, secim, ceza = set(), [], 0
    for dg in sorted(b["derslik_gruplari"], key=lambda d: -sum(s["n"] for s in d)):
        r = derslik_grubu_ayir(dg, g, span, D, kullanilan)
        if r is None:
            return None
        odalar, c = r
        kullanilan.update(odalar); secim.append((dg, odalar)); ceza += c
    if len(b["derslik_gruplari"]) > 1 and len({kat(o) for _, od in secim for o in od}) > 1:
        ceza += 4
    gerek = gozetmen_ihtiyaci(secim)
    if gerek and any(D.goz[(g, hh)] + gerek > D.arz.get((g, hh), 0) for hh in span):
        return None
    puan = -ceza
    for s in b["uyeler"]:
        if (g, h) in s["pref"]:
            puan += s["pref"][(g, h)]
        elif g in s["pref_gun"]:
            puan += s["pref_gun"][g]
    puan -= W_GUN_YUKU * D.gun_yuk[g]
    if b["ana"]:
        puan -= W_AYNI_GUN_KOHORT * sum(1 for k in b["ana"] if D.ana.get((k, g))) / len(b["ana"])
    return puan, secim


def gozetmen_ihtiyaci(secim):
    return sum((2 if kat(o) == 3 else 1) for dg, odalar in secim if any(s["gozetmen"] for s in dg) for o in odalar)


def koy(b, g, h, secim, D):
    span = range(h, h + b["dur"])
    gerek = gozetmen_ihtiyaci(secim)
    for hh in span:
        D.goz[(g, hh)] += gerek
        for k in b["kohort"]: D.kohort.add((k, g, hh))
        for x in b["hocalar"]: D.hoca.add((x, g, hh))
        for k in b["ana"]: D.ana[(k, g)][hh] = b["id"]
        for _, odalar in secim:
            for o in odalar: D.oda.add((o, g, hh))
    D.gun_yuk[g] += 1
    D.yer[b["id"]] = (g, h, secim)


def kaldir(b, D):
    g, h, secim = D.yer.pop(b["id"])
    gerek = gozetmen_ihtiyaci(secim)
    for hh in range(h, h + b["dur"]):
        D.goz[(g, hh)] -= gerek
        for k in b["kohort"]: D.kohort.discard((k, g, hh))
        for x in b["hocalar"]: D.hoca.discard((x, g, hh))
        for k in b["ana"]: D.ana[(k, g)].pop(hh, None)
        for _, odalar in secim:
            for o in odalar: D.oda.discard((o, g, hh))
    D.gun_yuk[g] -= 1


def en_iyi_yer(b, D, rnd):
    en = None
    for g in GUNLER:
        for h in SAATLER:
            r = dene(b, g, h, D)
            if r and (en is None or r[0] + rnd.random() * 0.5 > en[0]):
                en = (r[0] + rnd.random() * 0.5, g, h, r[1])
    return en


def yerlestir(birimler, kap, musait, arz, rnd):
    D = Durum(kap, musait, arz)
    def zorluk(b):
        izinli = sum(1 for g in GUNLER for h in SAATLER if izinli_mi(b, g, h))
        buyuk = b["n"] > BUYUK_SINAV
        return (0 if b["sert"] else 1, 0 if buyuk else 1, -b["n"] if buyuk else izinli, -len(b["kohort"]), -b["n"], rnd.random())
    sira = sorted(birimler, key=zorluk)
    kalan = []
    for b in sira:
        e = en_iyi_yer(b, D, rnd)
        if e:
            koy(b, e[1], e[2], e[3], D)
        else:
            kalan.append(b)
    # onarım: engelleyen en fazla 2 birimi kaldırıp yeniden yerleştir
    ada = {b["id"]: b for b in birimler}
    for b in list(kalan):
        basarili = False
        for g in GUNLER:
            for h in SAATLER:
                if not izinli_mi(b, g, h) or basarili:
                    continue
                span = range(h, h + b["dur"])
                engel = engelleyenler(b, g, h, D, ada)
                if not engel or len(engel) > 3:
                    continue
                yedek = {bid: D.yer[bid] for bid in engel}
                for bid in engel: kaldir(ada[bid], D)
                r = dene(b, g, h, D)
                if r:
                    koy(b, g, h, r[1], D)
                    tasindi = []
                    for bid in engel:
                        e = en_iyi_yer(ada[bid], D, rnd)
                        if not e: break
                        koy(ada[bid], e[1], e[2], e[3], D); tasindi.append(bid)
                    if len(tasindi) == len(engel):
                        basarili = True; kalan.remove(b); continue
                    for bid in tasindi: kaldir(ada[bid], D)
                    kaldir(b, D)
                for bid, (g2, h2, sec) in yedek.items(): koy(ada[bid], g2, h2, sec, D)
    # iyileştirme: her birimi daha iyi puanlı yere taşımayı dene
    for _ in range(3):
        degisti = False
        for b in rnd.sample(birimler, len(birimler)):
            if b["id"] not in D.yer:
                continue
            g0, h0, s0 = D.yer[b["id"]]
            kaldir(b, D)
            eski = dene(b, g0, h0, D)
            e = en_iyi_yer(b, D, rnd)
            if e and eski and e[0] > eski[0] + 1:
                koy(b, e[1], e[2], e[3], D); degisti = True
            else:
                koy(b, g0, h0, eski[1] if eski else s0, D)
        if not degisti:
            break
    talepleri_onar(birimler, D, rnd, ada)
    toplam = toplam_puan(birimler, D)
    return D, kalan, toplam


def engelleyenler(b, g, h, D, ada):
    span = set(range(h, h + b["dur"]))
    genis = set(range(h - 2, h + b["dur"] + 2))
    out = set()
    for bid, (g2, h2, secim) in D.yer.items():
        if g2 != g:
            continue
        s2 = set(range(h2, h2 + ada[bid]["dur"]))
        if s2 & span and (ada[bid]["kohort"] & b["kohort"] or ada[bid]["hocalar"] & b["hocalar"]):
            out.add(bid)
        elif s2 & genis and ada[bid]["ana"] & b["ana"]:
            out.add(bid)
        elif s2 & span and b["n"] > BUYUK_SINAV and any(kat(o) == 3 for _, od in secim for o in od):
            out.add(bid)
    return out


def talepleri_onar(birimler, D, rnd, ada):
    """Açık hoca talebi karşılanmayan birimler için engelleyenleri taşıyarak talebi karşılamayı dener."""
    for b in birimler:
        if b["id"] not in D.yer:
            continue
        g0, h0, s0 = D.yer[b["id"]]
        talepli = [s for s in b["uyeler"] if s["tercih_turu"] == "talep"]
        if not talepli or all((g0, h0) in s["pref"] or g0 in s["pref_gun"] for s in talepli):
            continue
        hedefler = sorted({x for s in talepli for x in s["pref"]}, key=lambda x: rnd.random())
        once = toplam_puan(birimler, D)
        for g, h in hedefler:
            if not izinli_mi(b, g, h):
                continue
            kaldir(b, D)
            engel = engelleyenler(b, g, h, D, ada)
            if len(engel) > 3:
                koy(b, g0, h0, s0, D); continue
            yedek = {bid: D.yer[bid] for bid in engel}
            for bid in engel: kaldir(ada[bid], D)
            r = dene(b, g, h, D)
            ok = False
            if r:
                koy(b, g, h, r[1], D)
                tasindi = []
                for bid in engel:
                    e = en_iyi_yer(ada[bid], D, rnd)
                    if not e: break
                    koy(ada[bid], e[1], e[2], e[3], D); tasindi.append(bid)
                if len(tasindi) == len(engel) and toplam_puan(birimler, D) > once:
                    ok = True
                else:
                    for bid in tasindi: kaldir(ada[bid], D)
                    kaldir(b, D)
            if ok:
                break
            for bid, (g2, h2, sec) in yedek.items(): koy(ada[bid], g2, h2, sec, D)
            koy(b, g0, h0, s0, D)


def toplam_puan(birimler, D):
    return sum((dene_puan(b, D) or 0) for b in birimler if b["id"] in D.yer)


def dene_puan(b, D):
    g, h, secim = D.yer[b["id"]]
    kaldir(b, D)
    r = dene(b, g, h, D)
    koy(b, g, h, secim, D)
    return r[0] if r else None


# ================================================================ gözetmen atama
def gorevleri_cikar(birimler, D):
    gorevler = []
    ada = {b["id"]: b for b in birimler}
    for bid, (g, h, secim) in D.yer.items():
        b = ada[bid]
        for dg, odalar in secim:
            if not any(s["gozetmen"] for s in dg):
                continue
            for o in odalar:
                adet = 2 if kat(o) == 3 else 1
                for i in range(adet):
                    gorevler.append({"birim": bid, "gun": g, "span": list(range(h, h + b["dur"])), "oda": o,
                                     "ing": any(s["ing"] for s in dg), "ciftli": adet == 2, "sira": i,
                                     "dersler": [s["ad"] for s in dg], "kisi": None})
    return gorevler


def gozetmen_ata(gorevler, kisiler, rnd, kilitli=frozenset()):
    """kilitli: önceden atanmış (senaryoda korunacak) görev indeksleri; dengelemede taşınmaz."""
    yuk = collections.Counter(); ing = collections.Counter()
    takvim = collections.defaultdict(dict)  # ad -> {(gün,saat): görev_idx}
    def uygun(k, gi):
        gv = gorevler[gi]
        if gv["gun"] in k["izin"] or set(gv["span"]) & k["yasak_saat"]:
            return False
        if any((gv["gun"], hh) in takvim[k["ad"]] for hh in gv["span"]):
            return False
        if k["ad"] not in GOZETMEN_ARDISIK_MUAF:
            saat = {h: v for (g, h), v in takvim[k["ad"]].items() if g == gv["gun"]}
            for hh in gv["span"]: saat[hh] = gi
            sol, sag = min(gv["span"]), max(gv["span"])
            while sol - 1 in saat: sol -= 1
            while sag + 1 in saat: sag += 1
            if len({saat[x] for x in range(sol, sag + 1)}) > GOZETMEN_ARDISIK_SINIRI:
                return False
        return True
    def maliyet(k, gi):
        gv = gorevler[gi]; ad = k["ad"]
        gunler = {g for g, _ in takvim[ad]}
        c = 10 * yuk[ad] + (8 * ing[ad] if gv["ing"] else 0)
        if ad in GOZETMEN_AZ_GUN: c += -18 if gv["gun"] in gunler else 0
        if ad in GOZETMEN_COK_GUN: c += -14 if gv["gun"] not in gunler else 0
        if any((gv["gun"], hh) in takvim[ad] for hh in (min(gv["span"]) - 1, max(gv["span"]) + 1)): c += 3
        if gv["ciftli"]:
            es = [gorevler[j]["kisi"] for j, x in enumerate(gorevler) if j != gi and x["birim"] == gv["birim"]
                  and x["oda"] == gv["oda"] and x["kisi"]]
            if es and k["cins"] and any(kk["ad"] == es[0] and kk["cins"] == k["cins"] for kk in kisiler):
                c += 15
        return c + rnd.random()
    def ata(k, gi):
        gorevler[gi]["kisi"] = k["ad"]; yuk[k["ad"]] += 1; ing[k["ad"]] += gorevler[gi]["ing"]
        for hh in gorevler[gi]["span"]: takvim[k["ad"]][(gorevler[gi]["gun"], hh)] = gi
    def birak(k, gi):
        gorevler[gi]["kisi"] = None; yuk[k["ad"]] -= 1; ing[k["ad"]] -= gorevler[gi]["ing"]
        for hh in gorevler[gi]["span"]: takvim[k["ad"]].pop((gorevler[gi]["gun"], hh), None)
    for gi in sorted(kilitli):
        ata(next(k for k in kisiler if k["ad"] == gorevler[gi]["kisi"]), gi)
    sira = sorted((i for i in range(len(gorevler)) if i not in kilitli),
                  key=lambda i: (sum(1 for k in kisiler if uygun(k, i)), gorevler[i]["sira"], rnd.random()))
    atanamayan = []
    for gi in sira:
        aday = [k for k in kisiler if uygun(k, gi)]
        if not aday:
            atanamayan.append(gi); continue
        ata(min(aday, key=lambda k: maliyet(k, gi)), gi)
    kad = {k["ad"]: k for k in kisiler}
    # denge: en yüklüden en yüksüze görev taşı
    for _ in range(400):
        mx = max(kisiler, key=lambda k: yuk[k["ad"]]); mn = min(kisiler, key=lambda k: yuk[k["ad"]])
        if yuk[mx["ad"]] - yuk[mn["ad"]] <= 1:
            break
        tasindi = False
        for gi in rnd.sample(range(len(gorevler)), len(gorevler)):
            if gorevler[gi]["kisi"] != mx["ad"] or gi in kilitli:
                continue
            birak(mx, gi)
            if uygun(mn, gi):
                ata(mn, gi); tasindi = True; break
            ata(mx, gi)
        if not tasindi:
            break
    # İngilizce dengesi: takas
    for _ in range(200):
        mx = max(kisiler, key=lambda k: ing[k["ad"]]); mn = min(kisiler, key=lambda k: ing[k["ad"]])
        if ing[mx["ad"]] - ing[mn["ad"]] <= 1:
            break
        ok = False
        for a in [i for i, x in enumerate(gorevler) if x["kisi"] == mx["ad"] and x["ing"] and i not in kilitli]:
            for bb in [i for i, x in enumerate(gorevler) if x["kisi"] == mn["ad"] and not x["ing"] and i not in kilitli]:
                birak(mx, a); birak(mn, bb)
                if uygun(mn, a) and uygun(mx, bb):
                    ata(mn, a); ata(mx, bb); ok = True; break
                ata(mx, a); ata(mn, bb)
            if ok: break
        if not ok:
            break
    return atanamayan


# ================================================================ bağımsız kontrol
def kontrol(satirlar, gorevler, kisiler, kap, musait, sinavlar):
    ihlal = []
    ada = {s["ad"]: s for s in sinavlar}
    zaman = collections.defaultdict(list)   # (gün,saat) -> satırlar
    for r in satirlar:
        for hh in r["span"]:
            zaman[(r["gun"], hh)].append(r)
    for (g, hh), rs in sorted(zaman.items()):
        dersler = {r["ders"] for r in rs}
        # kohort
        say = collections.Counter(k for d in dersler for k in ada[d]["kohort"])
        for k, v in say.items():
            if v > 1:
                ds = [d for d in dersler if k in ada[d]["kohort"]]
                birlikte = all(set(ds) - {d} <= set(ada[d]["birlikte"]) | ({ada[d]["es"]} if ada[d]["es"] else set()) for d in ds)
                if not birlikte:
                    ihlal.append(f"ÇAKIŞMA {g} {hh}:00 {k}: {ds}")
        # hoca
        hs = collections.defaultdict(set)
        for d in dersler:
            if ada[d]["hoca_id"] is not None: hs[ada[d]["hoca_id"]].add(d)
        for x, ds in hs.items():
            if len(ds) > 1 and not all(set(ds) - {d} <= set(ada[d]["birlikte"]) | ({ada[d]["es"]} if ada[d]["es"] else set()) for d in ds):
                ihlal.append(f"HOCA {g} {hh}:00 {ada[next(iter(ds))]['hoca']}: {sorted(ds)}")
        # derslik
        od = collections.defaultdict(set)
        for r in rs: od[r["oda"]].add(r["ders"])
        for o, ds in od.items():
            if len(ds) > 1 and not all(set(ds) - {d} <= set(ada[d]["birlikte"]) for d in ds):
                ihlal.append(f"DERSLİK {g} {hh}:00 {o}: {sorted(ds)}")
            if kat(o) in (3, 6) or o == "LAB":
                if (o, g, hh) not in musait:
                    ihlal.append(f"MÜSAİT DEĞİL {g} {hh}:00 {o}")
    # kapasite, izin, en geç saat, ardışık
    oda_top = collections.defaultdict(int)
    for r in satirlar:
        oda_top[(r["oda"], r["gun"], r["span"][0])] += r["kisi"]
        s = ada[r["ders"]]
        if r["gun"] in s["izin"]:
            ihlal.append(f"İZİN GÜNÜ {s['ad']} {r['gun']}")
        if r["span"][0] > s["en_gec"]:
            ihlal.append(f"EN GEÇ SAAT {s['ad']} {r['span'][0]}:00")
    for (o, g, h), n in oda_top.items():
        if o in kap and n > kap[o]:
            tolerans = n <= kap[o] + KAPASITE_TOLERANS or (o in KUCUK_TEK_DERSLIK and n <= 29)
            if not tolerans:
                ihlal.append(f"KAPASİTE {o} {g} {h}:00 {n}>{kap[o]}")
    toplam = collections.Counter()
    for r in satirlar:
        toplam[r["ders"]] += r["kisi"]
        if r["kisi"] == 0:
            ihlal.append(f"BOŞ DERSLİK {r['ders']} {r['oda']}")
    for d, v in toplam.items():
        if v != ada[d]["n"]:
            ihlal.append(f"ÖĞRENCİ TOPLAMI {d}: {v} != {ada[d]['n']}")
    ana_saat = collections.defaultdict(dict)
    for r in satirlar:
        for k in ada[r["ders"]]["ana"]:
            for hh in r["span"]:
                ana_saat[(k, r["gun"])][hh] = r["ders"]
    for (k, g), sa in ana_saat.items():
        hs = sorted(sa)
        zincir = [hs[0]] if hs else []
        for a, bq in zip(hs, hs[1:]):
            zincir = zincir + [bq] if bq == a + 1 else [bq]
            if len({sa[x] for x in zincir}) > ARDISIK_SINAV_SINIRI:
                ihlal.append(f"ARKA ARKAYA 3 SINAV {k} {g}: {[sa[x] for x in zincir]}")
    # gözetmen
    kad = {k["ad"]: k for k in kisiler}
    gt = collections.defaultdict(dict)
    for gv in gorevler:
        if not gv["kisi"]:
            ihlal.append(f"GÖZETMEN ATANAMADI {gv['gun']} {gv['span'][0]}:00 {gv['oda']} {gv['dersler']}"); continue
        k = kad[gv["kisi"]]
        if gv["gun"] in k["izin"]: ihlal.append(f"GÖZETMEN İZİN {k['ad']} {gv['gun']}")
        if set(gv["span"]) & k["yasak_saat"]: ihlal.append(f"GÖZETMEN SAAT {k['ad']} {gv['gun']} {gv['span']}")
        for hh in gv["span"]:
            if (gv["gun"], hh) in gt[k["ad"]] and gt[k["ad"]][(gv["gun"], hh)] is not gv:
                ihlal.append(f"GÖZETMEN İKİ YERDE {k['ad']} {gv['gun']} {hh}:00")
            gt[k["ad"]][(gv["gun"], hh)] = gv
    for ad, t in gt.items():
        if ad in GOZETMEN_ARDISIK_MUAF: continue
        for g in GUNLER:
            hs = sorted(h for gg, h in t if gg == g)
            zincir = []
            for i, h in enumerate(hs):
                zincir = zincir + [h] if zincir and h == zincir[-1] + 1 else [h]
                if len({id(t[(g, x)]) for x in zincir}) > GOZETMEN_ARDISIK_SINIRI:
                    ihlal.append(f"GÖZETMEN ARKA ARKAYA 3 {ad} {g}")
    return sorted(set(ihlal))


# ================================================================ ana akış
def main():
    args = sys.argv[1:]
    deneme = int(args[args.index("--deneme") + 1]) if "--deneme" in args else 30
    tohum = int(args[args.index("--tohum") + 1]) if "--tohum" in args else 2026
    wb = openpyxl.load_workbook(SABLON)
    kap, musait = derslikleri_oku(wb)
    kisiler = gozetmenleri_oku(wb)
    sinavlar = sinavlari_oku()
    birimler = birimleri_kur(sinavlar)

    arz = {(g, h): sum(1 for k in kisiler if g not in k["izin"] and h not in k["yasak_saat"])
           for g in GUNLER for h in SAATLER}
    en = None
    for t in range(deneme):
        rnd = random.Random(tohum + t)
        D, kalan, puan = yerlestir(birimler, kap, musait, arz, rnd)
        gorevler = gorevleri_cikar(birimler, D)
        atanamayan = gozetmen_ata(gorevler, kisiler, random.Random(tohum + t))
        anahtar = (len(kalan), len(atanamayan), -puan)
        if en is None or anahtar < en[0]:
            en = (anahtar, D, kalan, puan, tohum + t, gorevler, atanamayan)
    _, D, kalan, puan, secilen, gorevler, atanamayan = en

    satirlar = satirlari_kur(birimler, D, gorevler, kap)
    ihlaller = kontrol(satirlar, gorevler, kisiler, kap, musait, sinavlar)
    return dict(D=D, kalan=kalan, puan=puan, tohum=secilen, birimler=birimler, sinavlar=sinavlar,
                satirlar=satirlar, gorevler=gorevler, atanamayan=atanamayan, kisiler=kisiler,
                kap=kap, musait=musait, ihlaller=ihlaller, wb=wb)


def satirlari_kur(birimler, D, gorevler, kap):
    ada_b = {b["id"]: b for b in birimler}
    satirlar = []
    for bid, (g, h, secim) in D.yer.items():
        b = ada_b[bid]
        for dg, odalar in secim:
            atama = bolum_dagit(dg, odalar, {**kap, **{x: 999 for x in odalar if x not in kap}})
            for s in dg:
                for o in odalar:
                    bol = dict(atama.get((s["ad"], o), {}))
                    kisi = sum(bol.values())
                    if kisi == 0 and len(dg) > 1:
                        continue  # birlikte yapılan sınavda bu dersten bu derslikte öğrenci yok
                    gz = [gv["kisi"] for gv in gorevler if gv["birim"] == bid and gv["oda"] == o and s["ad"] in gv["dersler"]]
                    satirlar.append({"gun": g, "span": list(range(h, h + s["dur"])), "oda": o, "ders": s["ad"],
                                     "kisi": kisi, "bolumler": bol, "gozetmenler": gz, "sinav": s, "odalar": odalar})
    return satirlar


def bitis(h, sure):
    dk = h * 60 + sure
    return f"{dk // 60:02d}:{dk % 60:02d}"


def rapor_uret(R):
    D, sinavlar, satirlar, gorevler, kisiler = R["D"], R["sinavlar"], R["satirlar"], R["gorevler"], R["kisiler"]
    yer = {}
    for r in satirlar:
        yer.setdefault(r["ders"], (r["gun"], r["span"][0], set()))[2].add(r["oda"])
    satir = []
    yaz = satir.append
    yaz(f"SINAV PROGRAMI KONTROL RAPORU — {DONEM_ETIKETI} (v2)")
    yaz(f"Yerleşen sınav: {len(yer)} / {len(sinavlar)} | yerleşmeyen: {[u['ad'] for b in R['kalan'] for u in b['uyeler']] or 'yok'}")
    yaz(f"Sert kısıt ihlali: {len(R['ihlaller'])}")
    for i in R["ihlaller"]:
        yaz("  ! " + i)
    # talepler
    tam, gun, yok = [], [], []
    for s in sinavlar:
        if not s["tercih_turu"] or s["ad"] not in yer:
            continue
        g, h, _ = yer[s["ad"]]
        istenen = sorted(s["pref"], key=lambda x: (GUNLER.index(x[0]), x[1]))
        ozet = ", ".join(sorted({x[0] for x in istenen}, key=GUNLER.index)) + " " + \
               "/".join(f"{x:02d}" for x in sorted({x[1] for x in istenen}))
        if (g, h) in s["pref"]:
            tam.append(s)
        elif g in s["pref_gun"]:
            gun.append((s, f"{g} {h:02d}:00 (istenen: {ozet})"))
        else:
            yok.append((s, f"{g} {h:02d}:00 (istenen: {ozet})"))
        s["_sonuc"] = "tam" if (g, h) in s["pref"] else ("gün" if g in s["pref_gun"] else "yok")
        s["_istenen"] = ozet
    for tur in ("talep", "ders saati"):
        T = [s for s in sinavlar if s["tercih_turu"] == tur and s["ad"] in yer]
        yaz(f"\n{'Hoca talebi' if tur == 'talep' else 'Ders gün/saati tercihi (izin günü bilinmeyen)'}: {len(T)} sınav — "
            f"tam {sum(1 for s in T if s['_sonuc'] == 'tam')}, sadece gün {sum(1 for s in T if s['_sonuc'] == 'gün')}, "
            f"karşılanamadı {sum(1 for s in T if s['_sonuc'] == 'yok')}")
        for s, t in gun:
            if s["tercih_turu"] == tur: yaz(f"  ~ gün tuttu, saat kaydı: {s['ad']} ({s['hoca']}) → {t}")
        for s, t in yok:
            if s["tercih_turu"] == tur: yaz(f"  ✗ karşılanamadı: {s['ad']} ({s['hoca']}) → {t}")
    # derslik tercihleri
    yaz("\nKalabalık sınavlar (>60) ve katlar:")
    for s in sorted(sinavlar, key=lambda s: -s["n"]):
        if s["n"] > BUYUK_SINAV and s["ad"] in yer:
            g, h, odalar = yer[s["ad"]]
            katlar = sorted({str(kat(o)) for o in odalar})
            yaz(f"  {s['ad']} ({s['n']}): {g} {h:02d}:00 → {', '.join(sorted(odalar))} "
                f"{'✓ 3. kat' if katlar == ['3'] else ('~ 3. + 6. kat' if '3' in katlar else '✗ 6. kat')}")
    karisik = [s["ad"] for s in sinavlar if s["ad"] in yer and s["n"] <= BUYUK_SINAV and len({kat(o) for o in yer[s['ad']][2]}) > 1]
    ust_kucuk = [f"{s['ad']} ({s['n']})" for s in sinavlar if s["ad"] in yer and s["n"] <= BUYUK_SINAV and any(kat(o) == 3 for o in yer[s['ad']][2])]
    yaz(f"60 ve altı olup 3. kata düşen: {ust_kucuk or 'yok'}")
    yaz(f"Farklı katlara bölünen (≤60): {karisik or 'yok'}")
    # hoca istekleri (not)
    yaz("\nÖzel istekler:")
    for ad in RAPOR_OZEL_ISTEKLER:
        if ad in yer:
            g, h, odalar = yer[ad]
            yaz(f"  {ad}: {g} {h:02d}:00 → {', '.join(sorted(odalar))}")
    # gün dağılımı
    gs = collections.Counter(v[0] for v in yer.values())
    yaz("\nGünlere göre sınav: " + ", ".join(f"{g} {gs[g]}" for g in GUNLER))
    # gözetmen
    yaz(f"\nGözetmenlik görevi: {len(gorevler)} | atanamayan: {len(R['atanamayan'])}")
    kad = {k["ad"]: k for k in kisiler}
    for k in kisiler:
        G = [g for g in gorevler if g["kisi"] == k["ad"]]
        gunler = collections.Counter(g["gun"] for g in G)
        yaz(f"  {k['ad']:<26} toplam {len(G):>2} | İngilizce {sum(g['ing'] for g in G)} | "
            + ", ".join(f"{g} {gunler[g]}" for g in GUNLER if gunler[g]))
    ciftler = collections.defaultdict(list)
    for g in gorevler:
        if g["ciftli"] and g["kisi"]:
            ciftler[(g["birim"], g["oda"], g["gun"])].append(kad[g["kisi"]]["cins"])
    ayni = [k for k, v in ciftler.items() if len(v) == 2 and v[0] == v[1]]
    yaz(f"3. kat 2 gözetmenli derslik: {len(ciftler)} | kadın+erkek olmayan: {len(ayni)}")
    return "\n".join(satir)


def excele_yaz(R):
    wb = R["wb"]
    ince = Side(style="thin", color="FFBFBFBF")
    kenar = Border(left=ince, right=ince, top=ince, bottom=ince)
    ZEBRA = PatternFill("solid", fgColor="FFF2F7FB")
    ws = wb["Sınav Programı"]
    bas = [c.value for c in ws[1]]
    if "Bu Derslikteki Bölümler" not in bas:
        bas.insert(bas.index("Bu Derslikteki Öğrenci") + 1, "Bu Derslikteki Bölümler")
        from copy import copy as _copy
        stil = ws.cell(1, 1)
        for j, b in enumerate(bas, start=1):
            c = ws.cell(1, j, b)
            c.font, c.fill, c.alignment, c.border = _copy(stil.font), _copy(stil.fill), _copy(stil.alignment), _copy(stil.border)
    GENISLIK = {"Tarih": 11, "Gün": 10, "Başlangıç": 9, "Bitiş": 8, "Derslik": 9, "Derslik Kapasitesi": 10, "Ders Kodu": 18,
                "Ders": 44, "Program-Yarıyıl": 28, "Öğretim Elemanı": 32, "Bu Derslikteki Öğrenci": 10,
                "Bu Derslikteki Bölümler": 22, "Sınavın Toplam Öğrencisi": 10, "Gözetmen İhtiyacı": 10, "Gözetmen": 30, "Not": 40}
    for j, b in enumerate(bas, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = GENISLIK.get(b, 12)
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)
    sirali = sorted(R["satirlar"], key=lambda r: (GUNLER.index(r["gun"]), r["span"][0], r["oda"], r["ders"]))
    for i, r in enumerate(sirali, start=2):
        s = r["sinav"]; h = r["span"][0]
        notlar = []
        if s.get("_sonuc") == "tam": notlar.append(f"{'talep' if s['tercih_turu'] == 'talep' else 'ders saati tercihi'} karşılandı")
        elif s.get("_sonuc") == "gün": notlar.append(f"gün tuttu, saat farklı (istenen: {s['_istenen']})")
        elif s.get("_sonuc") == "yok": notlar.append(f"{'talep' if s['tercih_turu'] == 'talep' else 'ders saati tercihi'} karşılanamadı (istenen: {s['_istenen']})")
        if s["birlikte"]: notlar.append("aynı derslikte birlikte: " + ", ".join(s["birlikte"]))
        if s["es"]: notlar.append(f"aynı saatte: {s['es']}")
        if len(r["odalar"]) > 1: notlar.append(f"{len(r['odalar'])} dersliğe bölündü")
        if kat(r["oda"]) == 3 and s["gozetmen"]: notlar.append("3. kat: 2 gözetmen")
        if s["dur"] == 2: notlar.append(f"{s['sure']} dk: derslik 2 oturum")
        deger = {
            "Tarih": TARIH[r["gun"]], "Gün": r["gun"], "Başlangıç": f"{h:02d}:00", "Bitiş": bitis(h, s["sure"]),
            "Derslik": r["oda"], "Derslik Kapasitesi": R["kap"].get(r["oda"]), "Ders Kodu": s["kod"] or None,
            "Ders": s["ad"], "Program-Yarıyıl": s["kohort_str"], "Öğretim Elemanı": s["hoca"],
            "Bu Derslikteki Öğrenci": r["kisi"], "Sınavın Toplam Öğrencisi": s["n"],
            "Bu Derslikteki Bölümler": ", ".join(f"{b} {v}" for b, v in r["bolumler"].items()),
            "Gözetmen İhtiyacı": "Var" if s["gozetmen"] else "Yok",
            "Gözetmen": "; ".join(x for x in r["gozetmenler"] if x) if s["gozetmen"] else "Dersin hocası",
            "Not": "; ".join(notlar) or None,
        }
        for j, b in enumerate(bas, start=1):
            c = ws.cell(i, j, deger.get(b))
            c.alignment = Alignment(wrap_text=True, vertical="top"); c.border = kenar
            if GUNLER.index(r["gun"]) % 2 == 0: c.fill = ZEBRA
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{openpyxl.utils.get_column_letter(len(bas))}{len(sirali) + 1}"

    ws = wb["Derslik Kısıtlamaları"]
    h = [c.value for c in ws[1]]
    I = {k: i + 1 for i, k in enumerate(h)}
    dolu = collections.defaultdict(list)
    for r in R["satirlar"]:
        for hh in r["span"]:
            dolu[(r["oda"], r["gun"], hh)].append(r)
    for row in range(2, ws.max_row + 1):
        oda = ws.cell(row, I["Sınıf"]).value
        g = ws.cell(row, I["Tarih"]).value.split()[1]
        hh = int(str(ws.cell(row, I["Sınav Başlangıç Saati"]).value)[:2])
        rs = dolu.get((oda, g, hh), [])
        ws.cell(row, I["Program"]).value = "; ".join(dict.fromkeys(
            (f"{r['ders']}: " if len(rs) > 1 else "") + ", ".join(r["bolumler"]) for r in rs)) or None
        ws.cell(row, I["Ders Adı"]).value = "; ".join(dict.fromkeys(r["ders"] for r in rs)) or None
        ws.cell(row, I["Öğretim Elemanı"]).value = "; ".join(dict.fromkeys(r["sinav"]["hoca"] for r in rs)) or None
        gz = [x for r in rs for x in (r["gozetmenler"] if r["sinav"]["gozetmen"] else ["Dersin hocası"])]
        ws.cell(row, I["Gözetmen"]).value = "; ".join(dict.fromkeys(x for x in gz if x)) or None

    ws = wb["Gözetmenler"]
    h = [c.value for c in ws[1]]
    I = {k: i + 1 for i, k in enumerate(h)}
    for row in range(2, ws.max_row + 1):
        ad = ws.cell(row, I["Gözetmen"]).value
        if not ad or ad.startswith("Gözetmen kuralları"):
            continue
        G = [g for g in R["gorevler"] if g["kisi"] == ad]
        gunler = collections.Counter(g["gun"] for g in G)
        ws.cell(row, I["Toplam Görev"]).value = len(G)
        ws.cell(row, I["İngilizce Sınav"]).value = sum(g["ing"] for g in G)
        ws.cell(row, I["Günlere Dağılım"]).value = ", ".join(f"{g} {gunler[g]}" for g in GUNLER if gunler[g])
    wb.save(SABLON)


def json_yaz(R):
    out = {"tohum": R["tohum"], "puan": R["puan"], "ihlaller": R["ihlaller"],
           "yerlesim": [{"ders": r["ders"], "gun": r["gun"], "baslangic": r["span"][0], "sure_oturum": len(r["span"]),
                         "derslik": r["oda"], "ogrenci": r["kisi"], "bolumler": r["bolumler"], "gozetmenler": r["gozetmenler"]} for r in R["satirlar"]]}
    json.dump(out, open(PROGRAM_JSON, "w", encoding="utf-8"), ensure_ascii=False, indent=1)


if __name__ == "__main__":
    R = main()
    metin = rapor_uret(R)
    open(RAPOR, "w", encoding="utf-8").write(metin)
    json_yaz(R)
    print(metin)
    if "--yaz" in sys.argv:
        excele_yaz(R)
        print("\nŞablona yazıldı.")
