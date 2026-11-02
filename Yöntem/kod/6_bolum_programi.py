"""Bölüm bazlı sınav programı (bölüm çakışması dosyası biçiminde).

Şablon: Çıktılar/Sınav Programı/Güz Vize Bölümler Sınav Programı .xlsx (kullanıcının dosyası; her bölüm bir sekme,
Pazartesi-Cuma x saat tablosu, 1-4. sınıf renkleri). Şablona dokunulmaz; dolu hali "(dolu)" adıyla ayrı yazılır.
Program, kullanıcının Derslik Dosyası (6. Kat) dosyasından okunur (5_yerlestir.py --onayli ile aynı yükleme).

Çalıştırma:  py -3 Yöntem\\kod\\6_bolum_programi.py [--dosya <Derslik Dosyası.xlsx>]
"""
import argparse
import collections
import datetime
import importlib.util
import os
import re
import sys

import openpyxl
import kaynaklar as K
from openpyxl.styles import Alignment, Font, PatternFill

KOD = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, KOD)
_sp = importlib.util.spec_from_file_location("yerlestir", os.path.join(KOD, "5_yerlestir.py"))
y = importlib.util.module_from_spec(_sp)
_sp.loader.exec_module(y)

# sekme adı -> (program adı, kısa ad)
SEKMELER = {
    "SİYASET TR": ("Siyaset Bilimi ve Uluslararası İlişkiler", "SİYASET"),
    "SİYASET ING": ("Siyaset Bilimi ve Uluslararası İlişkiler (İngilizce)", "SİYASET İNG"),
    "İŞLETME TR": ("İşletme", "İŞLETME"),
    "İŞLETME ING": ("İşletme (İngilizce)", "İŞLETME İNG"),
    "UTF": ("Uluslararası Ticaret ve Finansman", "UTF"),
    "LOJİSTİK YÖN": ("Lojistik Yönetimi", "LOJİSTİK"),
    "HAVACILIK YÖN": ("Havacılık Yönetimi", "HAVACILIK"),
    "YBS": ("Yönetim Bilişim Sistemleri", "YBS"),
    "EKONOMİ": ("Ekonomi ve Finans (İngilizce)", "EKONOMİ"),
    "E TİCARET": ("Elektronik Ticaret ve Yönetimi", "E-TİCARET"),
}
KISA_AD = {p: k for p, k in SEKMELER.values()}
SINIF_DOLGU = {1: "FF0000", 2: "92D050", 3: "B4C6E7", 4: "F4B084"}      # şablondaki renk açıklaması
SAAT_ILK, GUN_SUTUN, ILK_SATIR = 9, "CDEFG", 5                         # C..G = Pzt..Cum, 5. satır = 09:00


def pdf_grubu(g):
    """'Ekonomi ve Finans (İng.) 7. yarıyıl' -> ('Ekonomi ve Finans (İngilizce)', 4. sınıf)"""
    m = re.match(r"(.*?)\s+(\d+)\.\s*yarıyıl$", g)
    if not m:                                                           # PDF'te yarıyıl yazmıyor: sınıf OBS'ten alınır
        return g.replace("(İng.)", "(İngilizce)"), None
    return m.group(1).replace("(İng.)", "(İngilizce)"), (int(m.group(2)) + 1) // 2


# Kullanıcının düzelttiği sınıflar: ders adı (başlangıç) -> sınıf (bütün bölümler için). Yeni düzeltme gelince buraya eklenir.
SINIF_DUZELT = {
    "Terörizm ve İstihbarat": 3, "Terrorism and Intelligence": 3,       # 21 Ekim: 3. sınıf (OBS 4 gösteriyor)
    "Davranışsal Finans": 3,
    "Managerial Leadership and Business Ethics": 4,
    "Current Issues in World Politics": 3,
    "Örgütsel Davranış Araştırmaları": 2,
    "Theory of International Trade": 3,
    "Innovation Management and Creativity": 4,
}


# Kullanıcının belirttiği bölümler (PDF/OBS'ten farklı): ders adı (başlangıç) -> program adları
ALAN_DUZELT = {
    "Sektör Kampüste: Dijital Muhasebe": {"Uluslararası Ticaret ve Finansman", "Yönetim Bilişim Sistemleri"},   # 22 Ekim: UTF ve YBS
    "Introduction to Economics-I": {"İşletme (İngilizce)", "Siyaset Bilimi ve Uluslararası İlişkiler (İngilizce)", "Ekonomi ve Finans (İngilizce)"},   # 23 Ekim: Bayraktar+Şimşek grupları tek sınav
    "International Finance": {"İşletme (İngilizce)", "Siyaset Bilimi ve Uluslararası İlişkiler (İngilizce)", "Ekonomi ve Finans (İngilizce)"},   # 23 Ekim: bütün İngilizce bölümler
}


def sinif_bilgisi(yerli):
    """Her sınav için ({bölüm adı: {sınıf}}, dersi alan bölümler) ve kontrol listesi.
    Bölüme dahil etme PDF'teki satırlara göre (prefix'li satırlar); satır yoksa OBS'te en az 3 öğrencisi olan bölümler.
    Sınıf: kullanıcı düzeltmesi > OBS (öğrencilerin sınıfı). PDF'teki ortak havuz yarıyılı farklıysa 'kontrol' listesine yazılır."""
    rows = K.pdf_oku()
    bilgi, kontrol = {}, []
    for e in yerli:
        parcalar = {K.ad_anahtar(re.sub(r"\s*\(.*?\)", "", p_)) for p_ in re.split(r"\s*/\s*", e["ad"])}
        kodlar = {re.sub(r"\[.*\]", "", k_) for k_ in e["tum"]}
        sur = [K.ad_anahtar(h.split("|")[0]) for h, _ in e["hocalar"]]
        eslesen = []
        for r in rows:
            if not ((r["kod"] and r["kod"] in kodlar) or K.ad_anahtar(r["ad"]) in parcalar):
                continue
            if r["hoca"] and sur and not any(s_ and s_ in K.ad_anahtar(r["hoca"]) for s_ in sur):
                continue
            eslesen.append(r)
        pdf_prog = set()
        for r in eslesen:
            if r["prefix"]:
                et = K.bolum_etiketi(r["prefix"], None).replace("(İng.)", "(İngilizce)")
                if et in KISA_AD:
                    pdf_prog.add(et)
        havuz = {(r["yy"] + 1) // 2 for r in eslesen if not r["prefix"] and r["yy"]}
        alan = pdf_prog or {p_ for p_, n in e["prog"].items() if n >= 3 and p_ in KISA_AD}
        alan = next((set(v) for k_, v in ALAN_DUZELT.items() if e["ad"].startswith(k_)), alan)
        duz = next((v for k_, v in SINIF_DUZELT.items() if k_.lower() in e["ad"].lower()), None)
        sinif = {}
        for p_ in alan:
            obs = e["sinif"].get(p_)
            sinif[p_] = {duz} if duz else ({obs} if obs else {0})
            if not duz and havuz and obs and obs not in havuz:
                kontrol.append((e["ad"], KISA_AD[p_], obs, "/".join(str(h) for h in sorted(havuz))))
        bilgi[id(e)] = (sinif, set(alan))
    return bilgi, kontrol


def dolgu(renk):
    return PatternFill("solid", fgColor=renk)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dosya", default=None, help="Derslik Dosyası (6. Kat); varsayılan _yedek_kullanici içindeki güncel kopya")
    a = ap.parse_args()
    sinavlar, izin, kt = y.sinavlari_kur()
    kaynak = a.dosya or os.path.join(y.CIKTI, "_yedek_kullanici", "Aşama 2 - Derslik Dosyası (6. Kat) - kullanıcı 20.10 17.48.xlsx")
    yuk, uyari = y.dosyadan_yukle(kaynak, sinavlar)
    yuk.update(y.ekstra_yukle(sinavlar, yuk))        # 3./4./11. kat sınavları (İktisada, SBM, Sosyolojide TK, YBS lab...) da programda görünsün
    for u in uyari:
        print("UYARI:", u)
    P = y.Program()
    for e in sinavlar:
        if id(e) in yuk:
            P.koy(e, *yuk[id(e)])
    yerli = [e for e in sinavlar if id(e) in P.yer]
    print(f"{len(yerli)} sınav yüklendi")

    sablon = next(os.path.join(y.CIKTI, x) for x in os.listdir(y.CIKTI) if "Vize B" in x and x.endswith(".xlsx") and not x.startswith("~"))
    wb = openpyxl.load_workbook(sablon)
    gunler = [datetime.datetime(2026, 11, 2 + i) for i in range(5)]

    bilgi, kontrol = sinif_bilgisi(yerli)
    cakisan = 0
    for ad, (prog, kisa) in SEKMELER.items():
        ws = wb[ad]
        for j, g in enumerate(gunler):                                  # tarihler (şablonda geçen yılın tarihleri var)
            ws[f"{GUN_SUTUN[j]}4"].value = g
        for r in range(ILK_SATIR, ILK_SATIR + 11):                      # önceki yıldan kalan hücreleri temizle
            for col in GUN_SUTUN:
                c = ws[f"{col}{r}"]
                c.value = None
                c.fill = PatternFill(fill_type=None)
        hucre = collections.defaultdict(list)                           # (satır, gün) -> [(sınav, sınıflar)]
        for e in yerli:
            sinif, alan = bilgi[id(e)]
            if prog not in sinif:
                continue
            s, _, _ = P.yer[id(e)]
            for x in P.slotlar(e, s):
                g, h = divmod(x, len(y.SAATLER))
                hucre[(ILK_SATIR + (h + y.SAATLER[0] - SAAT_ILK), g)].append((e, sorted(sinif[prog])))
        # Çarşamba 09-12: sınav yok
        for h in (9, 10, 11):
            c = ws[f"{GUN_SUTUN[2]}{ILK_SATIR + h - SAAT_ILK}"]
            c.fill = dolgu("D9D9D9")
            c.value = "Sınav yok"
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.font = Font(italic=True, color="7F7F7F")
        for (r, g), lst in hucre.items():
            c = ws[f"{GUN_SUTUN[g]}{r}"]
            satirlar, siniflar = [], []
            for e, sn in lst:
                sinif, alan = bilgi[id(e)]
                digerleri = sorted(KISA_AD[p] for p in alan if p != prog and p in KISA_AD)
                ortak = len(alan) >= 3
                metin = e["ad"].replace("Dept Elective: ", "") + (f"  ({'+'.join(str(k) for k in sn)}. sınıf)" if len(lst) > 1 else "")
                if digerleri:
                    metin += "\n" + ("ORTAK DERS, diğer bölümler: " if ortak else "Diğer bölüm: ") + ", ".join(digerleri)
                satirlar.append(metin)
                siniflar.append(set(sn))
            c.value = "\n\n".join(satirlar)
            ayni_sinif = len(lst) > 1 and any(siniflar[i] & siniflar[j] for i in range(len(lst)) for j in range(i + 1, len(lst)))
            if ayni_sinif:
                c.fill = dolgu("000000")                                 # aynı bölüm ve sınıfın iki dersi aynı saatte
                c.font = Font(bold=True, color="FFFFFF", size=9)
                cakisan += 1
            else:
                sn0 = min(min(s_) for s_ in siniflar)  # 0 = sınıf bilinmiyor (gri)
                c.fill = dolgu(SINIF_DOLGU.get(sn0, "D9D9D9"))
                c.font = Font(color="FFFFFF" if sn0 == 1 else "000000", size=9, bold=any(len(bilgi[id(e)][1]) >= 3 for e, _ in lst))
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            ws.row_dimensions[r].height = max(40, 13 * (sum(s_.count("\n") + 1 for s_ in satirlar) + 2 * (len(lst) - 1)) + 8)
        ws["H7"].value = "Kalın yazı = ortak ders (3+ bölüm). Her hücrede dersi alan diğer bölümler yazılı."
        ws["H8"].value = "Siyah hücre = aynı bölüm ve sınıfın iki dersi aynı saatte (kontrol edin)."
        ws["H9"].value = "Gri = Çarşamba 09-12 sınav yok. Renk = dersin sınıfı (1 kırmızı, 2 yeşil, 3 mavi, 4 turuncu)."
        for k in ("H7", "H8", "H9"):
            ws[k].font = Font(size=9, italic=True)
        for col in GUN_SUTUN:
            if (ws.column_dimensions[col].width or 0) < 26:
                ws.column_dimensions[col].width = 26
    if "Sınıf Kontrol" in wb.sheetnames:
        del wb["Sınıf Kontrol"]
    wk = wb.create_sheet("Sınıf Kontrol")
    wk.append(["Ders", "Bölüm", "Tabloda kullanılan sınıf (OBS)", "PDF ortak havuz yarıyılına göre sınıf", "Not"])
    for satir in sorted(set(kontrol)):
        wk.append(list(satir) + ["Farklı: doğrusunu söyleyin, SINIF_DUZELT'e eklenir"])
    for col, w_ in zip("ABCDE", (46, 22, 26, 30, 50)):
        wk.column_dimensions[col].width = w_
    for c in wk[1]:
        c.font = Font(bold=True)
    cikti = os.path.join(y.CIKTI, "Güz Vize Bölümler Sınav Programı (dolu).xlsx")
    try:
        wb.save(cikti)
    except PermissionError:
        cikti = cikti.replace(".xlsx", " (yeni).xlsx")
        wb.save(cikti)
    print("->", cikti, f"| siyah (aynı sınıf çakışması) hücre: {cakisan}")


if __name__ == "__main__":
    main()
