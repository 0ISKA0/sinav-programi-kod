"""Komisyon paneli: günlük tüm sınavlar + gözetmen boş/dolu tablosu. Tek dosya, veri AES-GCM ile şifreli.

Girdi : 8_site_uret.py'nin kullandığı veriler + Sınav Çıktıları/Gözetmen Atamaları - Liste N ... .xlsx (Gözetmen Yükü sayfası)
Şifre : Yöntem/panel_sifre.txt (yoksa rastgele üretilir; GitHub'a yüklenmez)
Çıktı : Sınav Çıktıları/Panel/Komisyon Paneli (Liste N).html
Çalıştırma: py -3 Yöntem\\kod\\9_panel_uret.py [--liste 2]
"""
import base64
import datetime
import importlib.util
import json
import math
import os
import re
import secrets
import sys

import openpyxl
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

KOD = os.path.dirname(os.path.abspath(__file__))
KOK = os.path.dirname(os.path.dirname(KOD))
sys.path.insert(0, KOD)

sp = importlib.util.spec_from_file_location("s8", os.path.join(KOD, "8_site_uret.py"))
m8 = importlib.util.module_from_spec(sp)
sp.loader.exec_module(m8)

GOZETMENSIZ = {"Turkish Language I"}   # bilerek gözetmen atanmayan çok derslikli sınavlar (kullanıcı kararı, 2 Kasım); panelde uyarı vermez
KISA = ["Pzt", "Sal", "Çar", "Per", "Cum"]
ITER = 300_000
SIFRE_DOSYASI = os.path.join(KOK, "Yöntem", "panel_sifre.txt")


def sifre_al():
    if os.path.exists(SIFRE_DOSYASI):
        return open(SIFRE_DOSYASI, encoding="utf-8").read().strip()
    alfabe = "abcdefghjkmnpqrstuvwxyz23456789"
    s = "".join(secrets.choice(alfabe) for _ in range(10))
    open(SIFRE_DOSYASI, "w", encoding="utf-8").write(s + "\n")
    return s


def etiket(ders):
    for onek, kisa in (("Sağlık Bil. Fak. ", "SBF: "), ("Ortak Seçmeli: ", "Ortak: ")):
        if ders.startswith(onek):
            return kisa + ders[len(onek):]
    return ders


def veri_olustur(liste_no):
    L, sinavlar = m8.sinavlari_topla()
    dosya = next(x for x in os.listdir(m8.KLASOR) if m8.nfc(x).startswith(f"Gözetmen Atamaları - Liste {liste_no}"))
    yol = os.path.join(m8.KLASOR, dosya)
    atama = m8.atamalar_oku(yol)
    data = m8.veri_kur(L, sinavlar, atama)
    sure = {(m8.temiz(e["ad"]), g, h): dk for e, g, h, dk, o in L}

    # --- sınavlar ---
    sin = []
    for e in data["exams"]:
        g = m8.ISO.index(e["date"])
        odalar = []
        for r in e["rooms"]:
            ham = r["proctor"]
            kisiler = [] if ham == "Gerekmiyor" else [k.strip() for k in ham.split(",")]
            durum = "hoca" if any("(dersin hocası)" in k for k in kisiler) else ""
            kisiler = [k.replace(" (dersin hocası)", "") for k in kisiler]
            odalar.append({"oda": r["room"], "ogr": r["students"], "gozetmen": kisiler, "durum": durum, "bolum": r["departments"]})
        hic = not any(o["gozetmen"] for o in odalar)
        sin.append({"g": g, "bas": e["start"], "bit": e["end"], "ders": e["course"], "kod": e["courseCode"], "hoca": e["teacher"],
                    "programlar": e["programs"], "ogr": e["totalStudents"], "notlar": e["notes"], "odalar": odalar,
                    "uyari": bool(len(odalar) >= 2 and hic and not e["course"].endswith(")") and e["course"] not in GOZETMENSIZ)})
    # --- gözetmenler ---
    ws = openpyxl.load_workbook(yol)["Gözetmen Yükü"]
    kisiler = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[0]:
            continue
        izin = []
        m = re.match(r"izin:\s*([^;]*)", r[10] or "")
        if m:
            izin = [KISA.index(x.strip()) for x in m.group(1).split(",") if x.strip() in KISA]
        kisiler[r[0]] = {"ad": r[0], "toplam": r[1], "atanan": r[2], "kendi": r[3], "ing": r[4], "gunluk": list(r[5:10]),
                         "izin": izin, "not": r[10] or "", "mesgul": []}
    for (ders, g, h), gor in atama.items():
        dk = sure.get((ders, g, h), 60)
        n = math.ceil(dk / 60)
        for oda, k in gor:
            if k in kisiler:
                kisiler[k]["mesgul"].append({"g": g, "h": h, "n": n, "ders": etiket(ders), "oda": "kendi dersi" if oda == "(kendi sınavı)" else oda,
                                             "kendi": oda == "(kendi sınavı)"})
    return {"meta": {"baslik": m8.SITE_BASLIGI.replace("Sınav Programı", "Sınav Paneli"), "donem": m8.DONEM,
                     "uretim": datetime.datetime.now().strftime("%d.%m.%Y %H:%M")},
            "gunler": [{"ad": m8.GUNLER[i], "tarih": m8.TARIH[i], "iso": m8.ISO[i], "renk": m8.GUN_RENK[m8.GUNLER[i]]} for i in range(5)],
            "sinavlar": sorted(sin, key=lambda t: (t["g"], t["bas"], t["ders"])),
            "kisiler": list(kisiler.values())}, dosya


def sifrele(data, sifre):
    tuz, iv = os.urandom(16), os.urandom(12)
    anahtar = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=tuz, iterations=ITER).derive(sifre.encode("utf-8"))
    ct = AESGCM(anahtar).encrypt(iv, json.dumps(data, ensure_ascii=False).encode("utf-8"), None)
    b = lambda x: base64.b64encode(x).decode()
    return {"s": b(tuz), "v": b(iv), "c": b(ct), "i": ITER}


if __name__ == "__main__":
    liste = sys.argv[sys.argv.index("--liste") + 1] if "--liste" in sys.argv else "2"
    data, dosya = veri_olustur(liste)
    sifre = "".join(sifre_al().split()).lower()   # tarayıcıdaki normalleştirmeyle aynı (boşluksuz, küçük harf)
    sablon = open(os.path.join(KOK, "Yöntem", "panel_sablonu.html"), encoding="utf-8").read()
    assert sablon.count("__PAYLOAD__") == 1
    html = sablon.replace("__PAYLOAD__", json.dumps(sifrele(data, sifre)))
    klasor = os.path.join(m8.KLASOR, "Panel")
    os.makedirs(klasor, exist_ok=True)
    yol = os.path.join(klasor, f"Komisyon Paneli (Liste {liste}).html")
    open(yol, "w", encoding="utf-8").write(html)
    uyari = [s for s in data["sinavlar"] if s["uyari"]]
    print("->", yol)
    print(len(data["sinavlar"]), "sınav,", len(data["kisiler"]), "gözetmen | gözetmensiz çok derslikli:", [(s["ders"], s["bas"]) for s in uyari])
    print("Şifre dosyası:", SIFRE_DOSYASI)
