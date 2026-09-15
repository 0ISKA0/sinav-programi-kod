# -*- coding: utf-8 -*-
"""
Adım 4 — Yayınlanmış programa sonradan gelen değişikliği EN AZ DEĞİŞİKLİKLE uygular (anlık tepki).

Mantık:
  1) Önceki program (program.json) yerinde sabitlenir: bütün sınavlar aynı gün/saat/derslikte.
  2) Değişen sınavlar (dersler.json'daki yeni hali ile farklı olanlar) kaldırılıp önce ESKİ yerlerinde
     denenir; olmuyorsa en yakın uygun saate taşınır (farklı gün ve saat farkı cezalı).
     Hâlâ yer yoksa engelleyen en fazla 2 sınav da en az kaymayla taşınır.
  3) Gözetmenler: değişmeyen görevler kişisiyle korunur; değişen görevlerde önce eski kişi denenir,
     olmazsa yalnızca o görevler yeniden atanır.
  4) Bağımsız kontrol, değişiklik raporu, şablon + bütün formlar (3_formlari_uret.py) yenilenir.

Değiştirilemez sınavlar (kullanıcı kuralı): başlamış/yapılmış ya da duyurusu kesinleşmiş sınavlar yerinden
oynatılmaz, gözetmeni değişmez, başka sınavlar geçmiş bir saate taşınmaz.
  --simdi "Salı 12:00"   : bu an ve öncesinde başlamış bütün sınavlar kilitlenir
  --sabit "Ders A;Ders B" : adı verilen sınavlar kilitlenir

Kullanım: python 4_senaryo.py "senaryo adı" [--simdi "Gün SS:00"] [--sabit "ad;ad"] [--onceki program.json] [--yaz]
"""
import collections, importlib.util, json, os, random, subprocess, sys, time

KOD = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("p", os.path.join(KOD, "2_programi_kur.py"))
p = importlib.util.module_from_spec(spec); spec.loader.exec_module(p)
import openpyxl

W_GUN_DEGISIMI, W_SAAT_FARKI = 25, 4


def onceki_yerlesim():
    args = sys.argv[1:]
    yol = args[args.index("--onceki") + 1] if "--onceki" in args else p.PROGRAM_JSON
    return json.load(open(yol, encoding="utf-8"))


def main():
    t0 = time.time()
    ad = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") else "senaryo"
    args = sys.argv[1:]
    simdi = None
    if "--simdi" in args:
        gun, saat = args[args.index("--simdi") + 1].split()
        simdi = (p.GUNLER.index(gun), int(saat[:2]))
    elle_sabit = set(filter(None, args[args.index("--sabit") + 1].split(";"))) if "--sabit" in args else set()

    def gecmis(g, h):
        """Bu saat başlamış ya da geçmiş mi?"""
        return simdi is not None and (p.GUNLER.index(g), h) <= simdi
    eski = onceki_yerlesim()
    wb = openpyxl.load_workbook(p.SABLON)
    kap, musait = p.derslikleri_oku(wb)
    kisiler = p.gozetmenleri_oku(wb)
    sinavlar = p.sinavlari_oku()
    birimler = p.birimleri_kur(sinavlar)
    arz = {(g, h): sum(1 for k in kisiler if g not in k["izin"] and h not in k["yasak_saat"]) for g in p.GUNLER for h in p.SAATLER}

    satir_ders = collections.defaultdict(list)
    for y in eski["yerlesim"]:
        satir_ders[y["ders"]].append(y)

    D = p.Durum(kap, musait, arz)
    degisen, sabit = [], []
    for b in birimler:
        uyeler = b["uyeler"]
        if any(s["ad"] not in satir_ders for s in uyeler):
            degisen.append(b); continue
        ilk = satir_ders[uyeler[0]["ad"]][0]
        eski_dur = max(satir_ders[s["ad"]][0]["sure_oturum"] for s in uyeler)
        if eski_dur != b["dur"]:
            degisen.append(b)
        sabit.append(b)
    # eski derslik seçimini yeniden kur (derslik sırası korunur)
    def eski_secim(b):
        secim = []
        for dg in b["derslik_gruplari"]:
            odalar = []
            for y in eski["yerlesim"]:
                if y["ders"] in {s["ad"] for s in dg} and y["derslik"] not in odalar:
                    odalar.append(y["derslik"])
            secim.append((dg, odalar))
        return secim
    konum = {}
    for b in sabit:
        y = satir_ders[b["uyeler"][0]["ad"]][0]
        konum[b["id"]] = (y["gun"], y["baslangic"], eski_secim(b))
    ada = {b["id"]: b for b in birimler}
    hareket = []
    # değiştirilemez (başlamış ya da kullanıcının kilitlediği) sınavlar
    donmus = {b["id"] for b in sabit if gecmis(*konum[b["id"]][:2]) or any(s["ad"] in elle_sabit for s in b["uyeler"])}
    for b in degisen:
        if b["id"] in donmus:
            g0, h0, _ = konum[b["id"]]
            for s in b["uyeler"]:
                s["dur"] = satir_ders[s["ad"]][0]["sure_oturum"]
                s["sure"] = 50 if s["dur"] == 1 else s["sure"]
            b["dur"] = max(s["dur"] for s in b["uyeler"])
            neden_k = "sınav başlamış/yapılmış" if gecmis(g0, h0) else "değiştirilemez olarak işaretli"
            hareket.append(f"UYGULANAMADI: {', '.join(s['ad'] for s in b['uyeler'])} ({g0} {h0:02d}:00) — {neden_k}; eski haliyle kaldı")
    degisen = [b for b in degisen if b["id"] not in donmus]
    degisen_id = {b["id"] for b in degisen}
    for b in sabit:
        if b["id"] not in degisen_id:
            g, h, sec = konum[b["id"]]
            p.koy(b, g, h, sec, D)

    rnd = random.Random(2026)

    def en_yakin(b, g0, h0):
        en = None
        for g in p.GUNLER:
            for h in p.SAATLER:
                if gecmis(g, h):
                    continue  # geçmiş saate taşınamaz
                r = p.dene(b, g, h, D)
                if not r:
                    continue
                puan = r[0] - (W_GUN_DEGISIMI if g != g0 else 0) - W_SAAT_FARKI * abs(h - h0)
                if en is None or puan > en[0]:
                    en = (puan, g, h, r[1])
        return en

    def neden(b, g0, h0):
        """Eski saatte uzatılamama sebebi (kullanıcıya açıklama)."""
        span = list(range(h0, h0 + b["dur"]))
        if span[-1] > 17:
            return f"{h0:02d}:00'da başlayan 90 dk sınav 18:00'den sonra biter"
        if span[-1] > b["en_gec"] or g0 in b["izin"]:
            return "hocanın saat/izin kısıtı"
        sebep = []
        for bid, (g2, h2, sec2) in D.yer.items():
            if g2 != g0 or not set(range(h2, h2 + ada[bid]["dur"])) & set(span):
                continue
            ortak_k = sorted(ada[bid]["kohort"] & b["kohort"])
            if ortak_k:
                sebep.append(f"{', '.join(ortak_k)} öğrencilerinin {ada[bid]['uyeler'][0]['ad']} sınavı {h2:02d}:00'da")
            elif ada[bid]["hocalar"] & b["hocalar"]:
                sebep.append(f"hocanın {ada[bid]['uyeler'][0]['ad']} sınavı {h2:02d}:00'da")
        if not sebep:
            for k in b["ana"]:
                saatler = dict(D.ana.get((k, g0), {}))
                for hh in span: saatler[hh] = b["id"]
                sol, sag = min(span), max(span)
                while sol - 1 in saatler: sol -= 1
                while sag + 1 in saatler: sag += 1
                if len({saatler[x] for x in range(sol, sag + 1)}) > p.ARDISIK_SINAV_SINIRI:
                    sebep.append(f"{k} öğrencilerinin aynı gün arka arkaya 3 sınavı olur"); break
        if not sebep:
            kullanilan, secim = set(), []
            for dg in sorted(b["derslik_gruplari"], key=lambda d: -sum(s["n"] for s in d)):
                r = p.derslik_grubu_ayir(dg, g0, span, D, kullanilan)
                if r is None:
                    bos_saat = [hh for hh in span if p.derslik_grubu_ayir(dg, g0, [hh], D, kullanilan) is None]
                    sebep.append(f"{', '.join(f'{x:02d}:00' for x in bos_saat)} oturumunda uygun derslik yok" if bos_saat else
                                 f"{span[0]:02d}:00–{span[0] + 1:02d}:30 boyunca kesintisiz boş kalan yeterli derslik yok")
                    break
                kullanilan.update(r[0]); secim.append((dg, r[0]))
            else:
                gerek = p.gozetmen_ihtiyaci(secim)
                for hh in span:
                    if D.goz[(g0, hh)] + gerek > D.arz.get((g0, hh), 0):
                        sebep.append(f"{hh:02d}:00'da gözetmen yetmiyor (o saatte {D.goz[(g0, hh)]} görev var, "
                                     f"{gerek} daha gerekiyor, müsait gözetmen {D.arz.get((g0, hh), 0)})")
        return "; ".join(dict.fromkeys(sebep)) or "tercih puanı (başka saat daha uygun)"

    for b in sorted(degisen, key=lambda b: -b["n"]):
        g0, h0, sec0 = konum.get(b["id"], (None, None, None))
        adlar = ", ".join(s["ad"] for s in b["uyeler"])
        r = p.dene(b, g0, h0, D) if g0 else None
        if r:
            # aynı saat tutuyor: eski derslikler hâlâ serbestse onları koru
            span = range(h0, h0 + b["dur"])
            eski_odalar_serbest = all(D.serbest(o, g0, span) for _, od in sec0 for o in od)
            secim = sec0 if eski_odalar_serbest else r[1]
            p.koy(b, g0, h0, secim, D)
            hareket.append(f"{adlar}: yerinde kaldı ({g0} {h0:02d}:00), {b['dur'] * 50 if b['dur'] == 1 else 90} dk'ya uzatıldı"
                           + ("" if eski_odalar_serbest else " — derslik değişti"))
            continue
        sebep = neden(b, g0, h0) if g0 else "yeni sınav"
        e = en_yakin(b, g0 or p.GUNLER[0], h0 or 12)
        if e:
            p.koy(b, e[1], e[2], e[3], D)
            hareket.append(f"{adlar}: {g0} {h0:02d}:00 → {e[1]} {e[2]:02d}:00 taşındı (sebep: {sebep})")
            continue
        # onarım: engelleyenleri de kaydır
        basarili = False
        for g in p.GUNLER:
            for h in p.SAATLER:
                if basarili or gecmis(g, h) or not p.izinli_mi(b, g, h):
                    continue
                engel = p.engelleyenler(b, g, h, D, ada)
                if not engel or len(engel) > 2 or engel & donmus:
                    continue  # değiştirilemez sınav yer açmak için taşınmaz
                yedek = {bid: D.yer[bid] for bid in engel}
                for bid in engel: p.kaldir(ada[bid], D)
                r = p.dene(b, g, h, D)
                if r:
                    p.koy(b, g, h, r[1], D)
                    tasinan = []
                    for bid in engel:
                        g2, h2, _ = yedek[bid]
                        e2 = en_yakin(ada[bid], g2, h2)
                        if not e2: break
                        p.koy(ada[bid], e2[1], e2[2], e2[3], D); tasinan.append((bid, g2, h2, e2[1], e2[2]))
                    if len(tasinan) == len(engel):
                        basarili = True
                        hareket.append(f"{adlar}: {g0} {h0:02d}:00 → {g} {h:02d}:00 taşındı")
                        for bid, g2, h2, g3, h3 in tasinan:
                            hareket.append(f"   yer açmak için: {', '.join(s['ad'] for s in ada[bid]['uyeler'])}: "
                                           f"{g2} {h2:02d}:00 → {g3} {h3:02d}:00")
                        continue
                    for bid, *_ in tasinan: p.kaldir(ada[bid], D)
                    p.kaldir(b, D)
                for bid, (g2, h2, s2) in yedek.items(): p.koy(ada[bid], g2, h2, s2, D)
        if not basarili:
            hareket.append(f"{adlar}: YERLEŞTİRİLEMEDİ")

    # ---- gözetmen sayısı azaldıysa: gereken gözetmen o saatteki müsait gözetmeni aşıyor mu?
    for _ in range(20):
        asan = sorted((k for k, v in D.goz.items() if v > D.arz.get(k, 0)), key=lambda k: (p.GUNLER.index(k[0]), k[1]))
        if not asan:
            break
        g, h = asan[0]
        adaylar = [ada[bid] for bid, (g2, h2, sec) in D.yer.items() if g2 == g and h2 <= h < h2 + ada[bid]["dur"]
                   and bid not in donmus and p.gozetmen_ihtiyaci(sec) > 0]
        if not adaylar:
            hareket.append(f"{g} {h:02d}:00: gözetmen yetmiyor ve kaydırılabilecek sınav yok"); break
        tasindi = False
        for b in sorted(adaylar, key=lambda b: b["n"]):
            g0, h0, s0 = D.yer[b["id"]]
            p.kaldir(b, D)
            e = en_yakin(b, g0, h0)
            if e:
                p.koy(b, e[1], e[2], e[3], D)
                hareket.append(f"{', '.join(s['ad'] for s in b['uyeler'])}: {g0} {h0:02d}:00 → {e[1]} {e[2]:02d}:00 taşındı "
                               f"(sebep: {g} {h:02d}:00'da gözetmen yetmiyor — {D.arz.get((g, h), 0)} müsait gözetmen)")
                tasindi = True; break
            p.koy(b, g0, h0, s0, D)
        if not tasindi:
            hareket.append(f"{g} {h:02d}:00: gözetmen yetmiyor, sınav kaydırılamadı"); break

    # ---- gözetmenler: korunan görevler kilitli
    gorevler = p.gorevleri_cikar(birimler, D)
    eski_goz = collections.defaultdict(list)   # (gün, saat, oda, ders) -> [kişi]
    for y in eski["yerlesim"]:
        for k in y["gozetmenler"]:
            if k and k not in eski_goz[(y["gun"], y["baslangic"], y["derslik"], y["ders"])]:
                eski_goz[(y["gun"], y["baslangic"], y["derslik"], y["ders"])].append(k)
    kilitli, aday_eski = set(), {}
    sayac = collections.Counter()
    kad = {k["ad"]: k for k in kisiler}
    musait_degil = []   # (görev, eski kişi): kişi artık bu görevi alamıyor (izin/saat kısıtı değişti)
    for i, gv in enumerate(gorevler):
        anahtar = (gv["gun"], gv["span"][0], gv["oda"])
        kisiler_eski = []
        for d in gv["dersler"]:
            kisiler_eski += [k for k in eski_goz.get(anahtar + (d,), []) if k not in kisiler_eski]
        j = sayac[anahtar]; sayac[anahtar] += 1
        if j < len(kisiler_eski):
            k0 = kisiler_eski[j]
            hala_uygun = k0 in kad and gv["gun"] not in kad[k0]["izin"] and not set(gv["span"]) & kad[k0]["yasak_saat"]
            if gv["birim"] in donmus:
                gv["kisi"] = k0; kilitli.add(i)       # başlamış sınavın gözetmeni değişmez
            elif not hala_uygun:
                musait_degil.append((i, k0))
            elif gv["birim"] in degisen_id:
                aday_eski[i] = k0                     # değişen sınav: eski kişi denenecek
            else:
                gv["kisi"] = k0; kilitli.add(i)
    # değişen görevlerde eski kişi uygunsa onu da kilitle
    dolu = collections.defaultdict(set)
    for i in kilitli:
        for hh in gorevler[i]["span"]:
            dolu[gorevler[i]["kisi"]].add((gorevler[i]["gun"], hh))
    goz_degisim = []
    for i, k in aday_eski.items():
        gv = gorevler[i]
        cak = any((gv["gun"], hh) in dolu[k] for hh in gv["span"]) or set(gv["span"]) & kad[k]["yasak_saat"]
        if not cak:
            gv["kisi"] = k; kilitli.add(i)
            for hh in gv["span"]: dolu[k].add((gv["gun"], hh))
        else:
            goz_degisim.append((i, k))
    atanamayan = p.gozetmen_ata(gorevler, kisiler, rnd, kilitli=frozenset(kilitli))
    for i, k in musait_degil:
        gv = gorevler[i]
        hareket.append(f"{', '.join(gv['dersler'])} ({gv['oda']} {gv['gun']} {gv['span'][0]:02d}:00): gözetmen {k} → "
                       f"{gv['kisi'] or 'ATANAMADI'} ({k} artık görev alamıyor)")
    for i, k in goz_degisim:
        goz_degisim_metin = f"{', '.join(gorevler[i]['dersler'])} ({gorevler[i]['oda']}): gözetmen {k} → {gorevler[i]['kisi']}"
        hareket.append(goz_degisim_metin)
    # arka arkaya kuralı yüzünden yeniden atanan (eski kişiden farklı) görevler
    for i, gv in enumerate(gorevler):
        if i not in kilitli and i not in dict(goz_degisim) and i not in dict(musait_degil) and gv["kisi"]:
            anahtar = (gv["gun"], gv["span"][0], gv["oda"])
            onceki = [k for d in gv["dersler"] for k in eski_goz.get(anahtar + (d,), [])]
            if gv["kisi"] not in onceki:
                hareket.append(f"{', '.join(gv['dersler'])} ({gv['oda']} {gv['gun']} {gv['span'][0]:02d}:00): gözetmen {gv['kisi']}")

    satirlar = p.satirlari_kur(birimler, D, gorevler, kap)
    ihlaller = p.kontrol(satirlar, gorevler, kisiler, kap, musait, sinavlar)
    R = dict(D=D, kalan=[b for b in birimler if b["id"] not in D.yer], puan=p.toplam_puan(birimler, D), tohum=eski.get("tohum"),
             birimler=birimler, sinavlar=sinavlar, satirlar=satirlar, gorevler=gorevler, atanamayan=atanamayan,
             kisiler=kisiler, kap=kap, musait=musait, ihlaller=ihlaller, wb=wb)

    # değişiklik özeti: eskiye göre farklı olan satırlar
    eski_imza = {(y["ders"], y["gun"], y["baslangic"], y["derslik"]) for y in eski["yerlesim"]}
    yeni_imza = {(r["ders"], r["gun"], r["span"][0], r["oda"]) for r in satirlar}
    etkilenen = sorted({x[0] for x in eski_imza ^ yeni_imza})
    R["donmus"] = sorted(s["ad"] for bid in donmus for s in ada[bid]["uyeler"])
    return R, hareket, etkilenen, time.time() - t0, ad


if __name__ == "__main__":
    R, hareket, etkilenen, sure, ad = main()
    print(f"SENARYO: {ad}")
    print("Değişiklikler:")
    for h in hareket:
        print("  •", h)
    print("Değiştirilemez (kilitli) sınav sayısı:", len(R["donmus"]))
    print("Yerinde değişen sınavlar (gün/saat/derslik):", etkilenen or "yok")
    print("Sert kısıt ihlali:", len(R["ihlaller"]), R["ihlaller"][:10])
    print("Atanamayan gözetmen görevi:", len(R["atanamayan"]))
    print(f"Hesap süresi: {sure:.1f} sn")
    if "--yaz" in sys.argv and not R["ihlaller"] and not R["kalan"] and not R["atanamayan"]:
        metin = p.rapor_uret(R)
        open(p.RAPOR, "w", encoding="utf-8").write(metin)
        p.json_yaz(R)
        p.excele_yaz(R)
        r = subprocess.run([sys.executable, os.path.join(KOD, "3_formlari_uret.py"), "--pdf"], capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env={**os.environ, "PYTHONIOENCODING": "utf-8"})
        print((r.stdout or "").strip() or (r.stderr or "").strip())
        print("Bütün dosyalar güncellendi.")
    elif "--yaz" in sys.argv:
        print("İhlal/yerleşmeyen var: dosyalar YAZILMADI.")
