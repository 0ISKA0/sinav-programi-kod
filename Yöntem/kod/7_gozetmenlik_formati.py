"""Gözetmen atamalarını eski örnek dosyanın biçiminde ('2025-26 Büt Gözetmenlikler.xlsx', tek sayfa 'Gzt') yazar.

Her gözetmen için 13 satırlık blok: ad (sarı), gün başlıkları, SAAT/TARİH + tarihler, 09:00-18:00 saat satırları, TOPLAM SINAV GÖREVİ.
Hücre: 'Ders (derslik)'. 90 dakikalık sınav iki saat satırını birleştirir.

Girdi: Sınav Çıktıları/Gözetmen Atamaları - Liste N ....xlsx ('Atamalar' sayfası; 5_yerlestir/gozetmen atama çıktısı)
Çalıştırma:  py -3 Yöntem\\kod\\7_gozetmenlik_formati.py
"""
import collections
import os
import unicodedata
import re
import sys
from copy import copy

import openpyxl

KOK = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
KLASOR = os.path.join(KOK, "Sınav Çıktıları")
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
GUN_BASLIK = ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
TARIHLER = ["02.11.2026", "03.11.2026", "04.11.2026", "05.11.2026", "06.11.2026"]


def ornek_dosya():
    nfc = lambda s: unicodedata.normalize("NFC", s)
    return os.path.join(KLASOR, next(x for x in os.listdir(KLASOR) if "Gözetmenlikler" in nfc(x) and nfc(x).startswith("2025-26")))


def atamalari_oku(yol):
    """[(kişi, gün_indeksi, başlangıç_saati, süre_saat, 'Ders (derslik)')] ve kişi sırası"""
    ws = openpyxl.load_workbook(yol)["Atamalar"]
    gorev, sira = collections.defaultdict(list), []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[1]:
            continue
        ders, gun, saat, oda, kisi = r[1], r[2], r[4], r[5], r[6]
        m = re.match(r"(\d\d):00-(\d\d):(\d\d)", saat)
        bas = int(saat[:2])
        sure = 2 if (int(m.group(2)) * 60 + int(m.group(3)) - bas * 60) > 60 else 1     # 90 dk ve 2 saatlik sınav iki saat satırı
        gorev[kisi].append((GUNLER.index(gun), bas, sure, ders, oda))
        if kisi not in sira:
            sira.append(kisi)
    tr = "abcçdefgğhıijklmnoöprsştuüvyz"
    sira.sort(key=lambda k: [tr.index(ch) if ch in tr else 99 for ch in k.lower().replace("i̇", "i")])
    return gorev, sira


def yaz(gorev, sira, derslik_haritasi, cikti):
    wb = openpyxl.load_workbook(ornek_dosya())
    ws = wb.active
    # örnek 1. bloğun biçimlerini al
    sablon = {(r, c): copy(ws.cell(r, c)._style) for r in range(1, 14) for c in range(1, 8)}
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    ws.delete_rows(1, ws.max_row)
    for i, kisi in enumerate(sira):
        b = 13 * i
        for r in range(1, 14):
            for c in range(1, 8):
                ws.cell(b + r, c)._style = copy(sablon[(r, c)])
        ws.cell(b + 1, 1).value = kisi
        for j, g in enumerate(GUNLER):
            ws.cell(b + 1, 2 + j).value = GUN_BASLIK[j]
            ws.cell(b + 2, 2 + j).value = TARIHLER[j]
        ws.cell(b + 2, 1).value = "SAAT/TARİH"
        for k in range(10):
            hucre = ws.cell(b + 3 + k, 1)
            hucre.value = __import__("datetime").time(9 + k, 0)
        toplam = 0
        for gun, bas, sure, ders, oda in sorted(gorev[kisi]):
            oda_ad = derslik_haritasi.get((ders, gun, bas), oda) if oda in ("(kendi sınavı)", "") else oda
            satir = b + 3 + (bas - 9)
            hucre = ws.cell(satir, 2 + gun)
            metin = f"{ders} ({oda_ad})"
            hucre.value = (hucre.value + "\n" if hucre.value else "") + metin
            if sure == 2:
                ws.merge_cells(start_row=satir, start_column=2 + gun, end_row=satir + 1, end_column=2 + gun)
            toplam += sure       # 90 dakikalık sınav 2 görev sayılır
        ws.merge_cells(start_row=b + 13, start_column=1, end_row=b + 13, end_column=6)
        ws.cell(b + 13, 1).value = f"TOPLAM SINAV GÖREVİ: {toplam}"
        ws.row_dimensions[b + 13].height = 14.55
        for k in range(3, 13):
            ws.row_dimensions[b + k].height = 28.8
    ws.title = "Gzt"
    try:
        wb.save(cikti)
    except PermissionError:
        cikti = cikti.replace(".xlsx", " (yeni).xlsx")
        wb.save(cikti)
    return cikti


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import importlib.util
    sp = importlib.util.spec_from_file_location("y", os.path.join(os.path.dirname(os.path.abspath(__file__)), "5_yerlestir.py"))
    y = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(y)
    sinavlar, izin, kt = y.sinavlari_kur()
    f = os.path.join(KLASOR, next(x for x in os.listdir(KLASOR) if unicodedata.normalize("NFC", x).startswith("Aşama 2 - Derslik")))
    yuk, ua = y.dosyadan_yukle(f, sinavlar)
    harita = {}
    for e in sinavlar:
        if id(e) in yuk:
            s, o, _ = yuk[id(e)]
            g, h = divmod(s, len(y.SAATLER))
            harita[(e["ad"].replace("Dept Elective: ", ""), g, h + y.SAATLER[0])] = ", ".join(sorted(o))
    for dosya in sorted(x for x in os.listdir(KLASOR) if unicodedata.normalize("NFC", x).startswith("Gözetmen Atamaları - Liste")):
        gorev, sira = atamalari_oku(os.path.join(KLASOR, dosya))
        ad = dosya.replace("Gözetmen Atamaları", "Gözetmenlikler")
        print("->", yaz(gorev, sira, harita, os.path.join(KLASOR, ad)), "|", len(sira), "kişi,", sum(len(v) for v in gorev.values()), "görev")
