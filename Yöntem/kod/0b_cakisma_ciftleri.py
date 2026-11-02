# -*- coding: utf-8 -*-
"""
Adım 0b — Ders Programı PDF'indeki derslere göre ÖĞRENCİ BAZLI sınav çakışma çiftlerini üretir.

İki dersin ortak öğrencisi varsa o iki dersin sınavı aynı gün ve saate konmamalıdır.
Girdi : Belgeler/Ders Programı.pdf (hangi dersler), Belgeler/altan ve üsten .xlsx (kim hangi dersi alıyor)
Çıktı : Çıktılar/Sınav Çakışma Çiftleri.xlsx   (öğrenci numaraları içerir; GitHub'a YÜKLENMEZ, Çıktılar/ .gitignore'da)
        Çıktılar/cakisma_ciftleri.csv           (yalnızca ders kodu çiftleri ve sayılar; zamanlayıcının girdisi)
Kod gösterimi: şimdilik PDF'teki ders kodu esas alınır (PDF ile OBS kodu farklıysa PDF kodu yazılır).

Kullanım: python 0b_cakisma_ciftleri.py
"""
import os, sys, csv, itertools
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kaynaklar import KOK, ders_listesi, tr_upper  # noqa: E402

CIKTI = os.path.join(KOK, "Çıktılar")
LISTE_ESIGI = 10  # bu sayıya kadar ortak öğrencinin numaraları satıra yazılır

BASLIK = PatternFill("solid", fgColor="1F3864")


def etiket(b):
    kod = " / ".join(b["kodlar"]) or "kodsuz"
    return f'{kod} {b["ad"]}'


def sayfa(wb, ad, baslik, satirlar, genislik):
    ws = wb.create_sheet(ad)
    ws.append(baslik)
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), BASLIK
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for s in satirlar:
        ws.append(s)
    for j, w in enumerate(genislik, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    return ws


def main():
    _, _, birimler = ders_listesi()
    birimler = [b for b in birimler if b["ogr"]]
    ciftler = []
    for a, b in itertools.combinations(birimler, 2):
        if a["tum"] & b["tum"]:  # aynı dersin farklı kodları
            continue
        ortak = a["ogr"] & b["ogr"]
        if ortak:
            ozel = ortak & (a["ozel"] | b["ozel"])
            ciftler.append((a, b, ortak, ozel))
    ciftler.sort(key=lambda c: (-len(c[2]), etiket(c[0])))

    satir = []
    for a, b, ortak, ozel in ciftler:
        liste = ", ".join(sorted(ortak)) if len(ortak) <= LISTE_ESIGI else f"{len(ortak)} öğrenci (liste uzun)"
        satir.append([etiket(a), ", ".join(a["hocalar"]), etiket(b), ", ".join(b["hocalar"]),
                      len(ortak), len(ozel), liste])
    sayisi = {id(b): [0, 0] for b in birimler}  # çakıştığı ders sayısı, en büyük ortak
    for a, b, ortak, _ in ciftler:
        for x in (a, b):
            sayisi[id(x)][0] += 1
            sayisi[id(x)][1] = max(sayisi[id(x)][1], len(ortak))
    ozet = [[etiket(b), ", ".join(b["hocalar"]) or "—", len(b["ogr"]), len(b["ozel"]), sayisi[id(b)][0], sayisi[id(b)][1]]
            for b in sorted(birimler, key=lambda b: (-sayisi[id(b)][0], etiket(b)))]

    os.makedirs(CIKTI, exist_ok=True)
    yol = os.path.join(CIKTI, "Sınav Çakışma Çiftleri.xlsx")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Açıklama")
    for s in ("Ortak öğrencisi olan ders çiftleri: bu derslerin sınavları aynı gün ve saate konmamalı.",
              "Kapsam: yalnızca Ders Programı PDF'inde yazan dersler; öğrenciler OBS ders kayıtlarından.",
              "Alttan/Üstten sütunu: ortak öğrencilerden o derslerden en az birini zorunlu/seçmeli dışında "
              "(alttan, devamlı alttan, üstten, yükseltme) alanların sayısı.",
              f"Ortak öğrenci sayısı {LISTE_ESIGI} ve altındaysa numaralar yazılır. Bu dosya öğrenci numarası içerir: paylaşmayın."):
        ws.append([s])
    ws.column_dimensions["A"].width = 140
    sayfa(wb, "Çakışma Çiftleri", ["Ders A", "Hoca A", "Ders B", "Hoca B", "Ortak öğrenci",
                                   "Bunlardan alttan/üstten alan", "Öğrenci No"], satir, (48, 26, 48, 26, 12, 16, 50))
    sayfa(wb, "Ders Özeti", ["Ders", "Hoca", "Kayıtlı öğrenci", "Alttan/üstten alan öğrenci",
                             "Çakıştığı ders sayısı", "En büyük ortak öğrenci"], ozet, (52, 30, 12, 14, 14, 14))
    try:
        wb.save(yol)
    except PermissionError:
        yol = yol.replace(".xlsx", " (yeni).xlsx"); wb.save(yol)
        print("UYARI: asıl dosya açık olduğu için yeni adla kaydedildi.")
    csv_yol = os.path.join(CIKTI, "cakisma_ciftleri.csv")
    with open(csv_yol, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["ders_a", "ders_b", "ortak_ogrenci", "alttan_ustten"])
        for a, b, ortak, ozel in ciftler:
            w.writerow([etiket(a), etiket(b), len(ortak), len(ozel)])
    print(f"{len(birimler)} ders, {len(ciftler)} çakışma çifti -> {yol}")
    print("En çok çakışan 5 çift:")
    for a, b, ortak, ozel in ciftler[:5]:
        print(f"  {len(ortak):4d} ortak ({len(ozel)} alttan/üstten): {etiket(a)}  <->  {etiket(b)}")


if __name__ == "__main__":
    main()
