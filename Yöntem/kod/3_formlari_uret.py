# -*- coding: utf-8 -*-
"""
Adım 3 — Kurulan programdan resmi formları üretir.

1) Yayın formu  : talep dosyasının (bölüm/yarıyıl blokları) biçimi korunur; D–G sütunlarına sınav
                  tarihi, başlangıç, bitiş, derslik yazılır; talep sütunları gizlenir. Excel + PDF.
2) Derslik formu: üniversitenin "Boş Derslik Dosyası" — yalnızca 6. kat sekmesi doldurulur
                  (her hücre: ÖĞRETİM ELEMANI / DERS).

Girdi : Yöntem/kod/program.json, Yöntem/kod/dersler.json,
        Belgeler/Sınav Programı Talep Dosyası.xlsx, Belgeler/Boş Derslik Dosyası.xlsx
Çıktı : Çıktılar/…Yayın.xlsx, Çıktılar/…Yayın.pdf, Çıktılar/…Derslik Dosyası.xlsx
"""
import collections, json, os, re, subprocess, sys
from copy import copy
import openpyxl

try:
    from donem_ayarlari import *   # dönem bilgileri ve şifreler (GitHub'a yüklenmez)
except ModuleNotFoundError:
    raise SystemExit("donem_ayarlari.py bulunamadı: 'donem_ayarlari_ornek.py' dosyasını 'donem_ayarlari.py' adıyla kopyalayıp doldurun.")
from openpyxl.styles import Alignment
from openpyxl.worksheet.properties import PageSetupProperties

KOK = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
KOD = os.path.join(KOK, "Yöntem", "kod")
TALEP = os.path.join(KOK, "Belgeler", TALEP_DOSYASI)
BOS_DERSLIK = os.path.join(KOK, "Belgeler", BOS_DERSLIK_DOSYASI)
CIKTI = os.path.join(KOK, "Çıktılar")
YAYIN_XLSX = os.path.join(CIKTI, f"{DONEM_ETIKETI} Sınav Programı - Yayın.xlsx")
YAYIN_PDF = YAYIN_XLSX[:-5] + ".pdf"
DERSLIK_XLSX = os.path.join(CIKTI, f"{DONEM_ETIKETI} - Derslik Dosyası (6. Kat).xlsx")

BASLIK = YAYIN_BASLIGI
GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
TARIH = dict(SINAV_TARIHLERI)


def verileri_oku():
    program = json.load(open(os.path.join(KOD, "program.json"), encoding="utf-8"))
    dersler = json.load(open(os.path.join(KOD, "dersler.json"), encoding="utf-8"))["sinavlar"]
    yer = {}
    for y in program["yerlesim"]:
        d = yer.setdefault(y["ders"], {"gun": y["gun"], "bas": y["baslangic"], "dur": y["sure_oturum"], "odalar": [],
                                        "bolum_oda": collections.defaultdict(list), "oda_bolum": {}})
        if y["derslik"] not in d["odalar"]:
            d["odalar"].append(y["derslik"])
        d["oda_bolum"][y["derslik"]] = list(y.get("bolumler", {}))
        for b in y.get("bolumler", {}):
            d["bolum_oda"][b].append(y["derslik"])
    sira = lambda o: ({"3": 0, "6": 1}.get(o[0], 2), o)  # 3. kat önce, sonra 6. kat, sonra diğerleri
    for d in yer.values():
        d["odalar"].sort(key=sira)
        for b in d["bolum_oda"]:
            d["bolum_oda"][b].sort(key=sira)
    return program, dersler, yer


def sure_dk(c):
    return int(c["Süre"].split()[0]) if c.get("Süre") and int(c["Süre"].split()[0]) > 50 else 50


def bitis(h, dk):
    """Gerçek bitiş saati (kullanıcı kararı): 50 dk -> HH:50, 90 dk -> HH+1:30."""
    t = h * 60 + dk
    return f"{t // 60:02d}:{t % 60:02d}"


def hucre_yaz(ws, satir, sutun, deger):
    """Birleştirilmiş hücreye yazmadan önce birleştirmeyi çözer."""
    for m in list(ws.merged_cells.ranges):
        if m.min_row <= satir <= m.max_row and m.min_col <= sutun <= m.max_col and (m.min_row, m.min_col) != (satir, sutun):
            ws.unmerge_cells(str(m))
    ws.cell(satir, sutun).value = deger


# ================================================================ 1) yayın formu
def yayin_formu(dersler, yer):
    wb = openpyxl.load_workbook(TALEP)
    ws = wb.active
    son = ws.max_row
    havuz_bas = next((r for r in range(1, son + 1) if str(ws.cell(r, 1).value or "").startswith("Fakülte Seçmeli Havuzu")), son + 1)
    satir_ders = {}
    for c in dersler:
        for r in c["_kaynak_satirlar"]:
            satir_ders[r] = c

    yazilan, eksik = set(), []
    for r in range(3, son + 1):
        havuz = r >= havuz_bas
        s0 = 3 if havuz else 4                     # sınav günü sütunu (havuzda bir sola kayık)
        ad_sutun = 2 if havuz else 3
        a = str(ws.cell(r, 1).value or "")
        hucreler = [str(ws.cell(r, k).value or "") for k in range(1, 9)]
        tum = " ".join(hucreler)
        # başlık satırları
        if a.strip() in ("BÖLÜM", "DEPARTMENT") or hucreler[0 if havuz else 1].strip() in ("SınavKODU",) or \
                (havuz and a.strip() in ("SınavKODU",)):
            ing = a.strip() == "DEPARTMENT"
            etiket = ["DAY", "START TIME", "END TIME", "HALL"] if ing else ["SINAV GÜNÜ", "BAŞLANGIÇ", "BİTİŞ", "DERSLİK"]
            for i, e in enumerate(etiket):
                hucre_yaz(ws, r, s0 + i, e)
            if not ing:
                hucre_yaz(ws, r, s0 - 2, "DERS KODU"); hucre_yaz(ws, r, s0 - 1, "DERS ADI")
            continue
        if re.search(r"SınavSAATİ|DERSSAATİ", tum) or (hucreler[s0 - 1:s0 + 3] and set(filter(None, (h.strip() for h in hucreler))) <= {"BAŞLANGIÇ", "BİTİŞ"}):
            for k in range(1, 9):
                if str(ws.cell(r, k).value or "").strip() in ("SınavSAATİ", "DERSSAATİ", "BAŞLANGIÇ", "BİTİŞ"):
                    hucre_yaz(ws, r, k, None)
            continue
        c = satir_ders.get(r)
        ad = str(ws.cell(r, ad_sutun).value or "")
        yer_tutucu = bool(re.search(r"Alt Tarafta|Bakınız|Bakiniz", ad, re.I))
        if c is None or yer_tutucu:
            # ders olmayan veri satırı: dönem ders programı bilgisini temizle
            if r > 2 and not blok_basligi(ws, r):
                for i in range(4):
                    if ws.cell(r, s0 + i).value not in (None, ""):
                        hucre_yaz(ws, r, s0 + i, None)
            continue
        if c["Sınav Var mı?"] == "Evet" and c["Ders"] in yer:
            y = yer[c["Ders"]]
            bolum = c.get("_satir_bolum", {}).get(str(r))
            odalar = y["bolum_oda"].get(bolum) or y["odalar"]   # bu satırın bölümünün dersliği; havuz satırında hepsi
            degerler = [TARIH[y["gun"]], f"{y['bas']:02d}:00", bitis(y["bas"], sure_dk(c)), "/".join(odalar)]
            for i, v in enumerate(degerler):
                hucre_yaz(ws, r, s0 + i, v)
            yazilan.add(c["Ders"])
        else:
            metin = "ORTAK DERSLER" if "Ortak Dersler Bölümü" in (c["Not (otomatik)"] or "") else "Sınav yapılmayacak"
            for i in range(4):
                hucre_yaz(ws, r, s0 + i, None)
            hucre_yaz(ws, r, s0, metin)
            ws.merge_cells(start_row=r, start_column=s0, end_row=r, end_column=s0 + 3)
            ws.cell(r, s0).alignment = Alignment(horizontal="center", vertical="center")
        if c["Sınav Var mı?"] == "Evet" and c["Ders"] not in yer:
            eksik.append(c["Ders"])

    # havuz bloklarını diğer bloklarla aynı hizaya getir (A..G -> B..H)
    for r in range(havuz_bas + 1, son + 1):
        if blok_basligi(ws, r):
            continue
        degerler = [ws.cell(r, k).value for k in range(1, 8)]
        if not any(v not in (None, "") for v in degerler):
            continue
        for m in list(ws.merged_cells.ranges):
            if m.min_row <= r <= m.max_row and m.max_col <= 8:
                ws.unmerge_cells(str(m))
        stiller = [copy(ws.cell(r, k)._style) for k in range(1, 8)]
        ws.cell(r, 1).value = None
        for k, (v, st) in enumerate(zip(degerler, stiller), start=2):
            ws.cell(r, k).value = v
            ws.cell(r, k)._style = st
        if str(degerler[0] or "").strip() == "DERS KODU":
            ws.cell(r, 1).value = "HAVUZ"
    ws.column_dimensions["G"].width = 24
    ws.column_dimensions["D"].width = 11
    for r in range(3, son + 1):
        ws.cell(r, 7).alignment = Alignment(wrap_text=True, vertical=ws.cell(r, 7).alignment.vertical)
    hucre_yaz(ws, 1, 1, BASLIK)
    # talep sütunlarını gizle, baskı ayarı
    for k in range(9, ws.max_column + 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(k)].hidden = True
    ws.print_area = f"A1:H{son}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.page_margins.left = ws.page_margins.right = 0.25
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    wb.save(YAYIN_XLSX)
    return yazilan, eksik


def blok_basligi(ws, r):
    a = str(ws.cell(r, 1).value or "")
    return bool(re.search(r"YARIYIL|SEMESTER|Havuzu", a, re.I)) and not any(ws.cell(r, k).value for k in range(2, 9))


def pdf_uret():
    ps = (f"$x = New-Object -ComObject Excel.Application; $x.DisplayAlerts = $false; "
          f"$w = $x.Workbooks.Open('{YAYIN_XLSX}'); $w.ActiveSheet.ExportAsFixedFormat(0, '{YAYIN_PDF}'); "
          f"$w.Close($false); $x.Quit()")
    r = subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True)
    return os.path.exists(YAYIN_PDF), r.stderr.strip()


# ================================================================ 2) derslik formu (6. kat)
GUN_BASLIK = {"PAZARTESİ": "Pazartesi", "SALI": "Salı", "ÇARŞAMBA": "Çarşamba", "PERŞEMBE": "Perşembe", "CUMA": "Cuma"}


def derslik_formu(dersler, program):
    wb = openpyxl.load_workbook(BOS_DERSLIK)
    ws = wb["6.Kat"]
    ada = {c["Ders"]: c for c in dersler}
    harita = {}  # (oda, gün, saat) -> (satır, hoca sütunu)
    for r in range(1, ws.max_row + 1):
        m = re.match(r"\s*(6[AD]\d)\s*-", str(ws.cell(r, 2).value or ""))
        if not m:
            continue
        oda = m.group(1)
        gun_sutun = {GUN_BASLIK[str(ws.cell(r + 1, k).value).strip().upper().replace("I", "I")]: k
                     for k in range(3, 13) if str(ws.cell(r + 1, k).value or "").strip().upper() in GUN_BASLIK}
        for rr in range(r + 3, r + 16):
            s = re.match(r"(\d{2}):00-", str(ws.cell(rr, 2).value or ""))
            if s:
                for g, k in gun_sutun.items():
                    harita[(oda, g, int(s.group(1)))] = (rr, k)
    icerik = collections.defaultdict(lambda: {"hoca": [], "ders": []})
    oda_sayisi = collections.Counter(y["ders"] for y in program["yerlesim"])
    for y in program["yerlesim"]:
        if not y["derslik"].startswith("6"):
            continue
        c = ada[y["ders"]]
        etiket = y["ders"]
        if oda_sayisi[y["ders"]] > 1 and y.get("bolumler"):   # bölünen sınav: bu dersliğe giren bölümler
            etiket += " (" + "-".join(y["bolumler"]) + ")"
        for hh in range(y["baslangic"], y["baslangic"] + y["sure_oturum"]):
            e = icerik[(y["derslik"], y["gun"], hh)]
            if c["Öğretim Elemanı"] not in e["hoca"]: e["hoca"].append(c["Öğretim Elemanı"])
            if etiket not in e["ders"]: e["ders"].append(etiket)
    eksik = []
    for anahtar, e in icerik.items():
        if anahtar not in harita:
            eksik.append(anahtar); continue
        rr, k = harita[anahtar]
        for kk, v in ((k, " / ".join(e["hoca"])), (k + 1, " / ".join(e["ders"]))):
            hucre = ws.cell(rr, kk)
            hucre.value = v
            hucre.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wb.save(DERSLIK_XLSX)
    return len(icerik), eksik


# ================================================================ 3) gözetmenlik çizelgesi
GOZETMEN_XLSX = os.path.join(CIKTI, f"{DONEM_ETIKETI} - Gözetmenlik.xlsx")


def gozetmen_gorevleri(dersler, program):
    """-> {gözetmen: [(gün, saat, oturum, metin, ingilizce)]} ; birlikte yapılan sınavlar tek görev."""
    ada = {c["Ders"]: c for c in dersler}
    gorev = collections.defaultdict(dict)
    for y in program["yerlesim"]:
        c = ada[y["ders"]]
        ing = any(k.startswith(("PSIR", "BUS", "ECF")) for k in c["Program-Yarıyıl (Çakışma Grubu)"].split("; "))
        for ad in y["gozetmenler"]:
            if not ad:
                continue
            anahtar = (y["gun"], y["baslangic"], y["derslik"])
            e = gorev[ad].setdefault(anahtar, {"dersler": [], "oturum": y["sure_oturum"], "ing": False})
            if y["ders"] not in e["dersler"]:
                e["dersler"].append(y["ders"])
            e["ing"] |= ing
    return gorev


def gozetmen_cizelgesi(dersler, program):
    from openpyxl.styles import Border, Font, PatternFill, Side
    wb_s = openpyxl.load_workbook(os.path.join(KOK, SABLON_DOSYASI))
    kisiler = [r[0] for r in wb_s["Gözetmenler"].iter_rows(min_row=2, values_only=True)
               if r[0] and not str(r[0]).startswith("Gözetmen kuralları")]
    gorev = gozetmen_gorevleri(dersler, program)
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Gzt"
    ince = Side(style="thin"); kenar = Border(left=ince, right=ince, top=ince, bottom=ince)
    KALIN = Font(bold=True, size=11); SARI = PatternFill("solid", fgColor="FFFFC000")
    ORTA = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions["A"].width = 22
    for k in "BCDEF":
        ws.column_dimensions[k].width = 30
    satir = 1
    for ad in kisiler:
        G = gorev.get(ad, {})
        basliklar = [ad] + ["PAZARTESİ", "SALI", "ÇARŞAMBA", "PERŞEMBE", "CUMA"]
        tarihler = ["SAAT/TARİH"] + [TARIH[g] for g in GUNLER]
        for j, v in enumerate(basliklar, start=1):
            c = ws.cell(satir, j, v); c.font = KALIN; c.alignment = ORTA; c.border = kenar
        ws.cell(satir, 1).fill = SARI
        for j, v in enumerate(tarihler, start=1):
            c = ws.cell(satir + 1, j, v); c.font = KALIN; c.alignment = ORTA; c.border = kenar
        for i, h in enumerate(range(9, 19)):
            r = satir + 2 + i
            c = ws.cell(r, 1, f"{h:02d}:00"); c.font = KALIN; c.alignment = ORTA; c.border = kenar
            for j, g in enumerate(GUNLER, start=2):
                ws.cell(r, j).border = kenar; ws.cell(r, j).alignment = ORTA; ws.cell(r, j).font = KALIN
        for (g, h, oda), e in sorted(G.items()):
            metin = " / ".join(e["dersler"]) + f" ({oda})"
            for k in range(e["oturum"]):
                hucre = ws.cell(satir + 2 + (h - 9) + k, GUNLER.index(g) + 2)
                hucre.value = metin if k == 0 else metin + " — devam"
            ws.row_dimensions[satir + 2 + (h - 9)].height = 30
        r = satir + 12
        ws.cell(r, 1, f"TOPLAM SINAV GÖREVİ: {len(G)}").font = KALIN
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.cell(r, 1).alignment = ORTA
        for j in range(1, 7): ws.cell(r, j).border = kenar
        satir += 14
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    # özet
    oz = wb.create_sheet("Özet")
    bas = ["Gözetmen", "Toplam Görev", "İngilizce Sınav"] + GUNLER
    for j, v in enumerate(bas, start=1):
        c = oz.cell(1, j, v); c.font = KALIN; c.fill = SARI; c.border = kenar; c.alignment = ORTA
    for i, ad in enumerate(kisiler, start=2):
        G = gorev.get(ad, {})
        gun = collections.Counter(g for g, _, _ in G)
        for j, v in enumerate([ad, len(G), sum(e["ing"] for e in G.values())] + [gun[g] or None for g in GUNLER], start=1):
            c = oz.cell(i, j, v); c.border = kenar
    oz.column_dimensions["A"].width = 28
    wb.save(GOZETMEN_XLSX)
    return {ad: len(gorev.get(ad, {})) for ad in kisiler}


# ================================================================ 4) web sitesi
SITE_SABLON = os.path.join(KOK, "Yöntem", "site_sablonu.html")   # verisiz, şifresiz şablon
SITE_HTML = os.path.join(CIKTI, SITE_DOSYASI)
SITE_YAYIN_HTML = os.path.join(CIKTI, "GitHub Pages", "index.html")   # internete konacak sürüm
BOLUMLER = [("SBUİ", "Siyaset Bilimi ve Uluslararası İlişkiler", "Türkçe"),
            ("PSIR", "Political Science and International Relations", "English"),
            ("İŞL", "İşletme", "Türkçe"), ("BUS", "Business Administration", "English"),
            ("UTF", "Uluslararası Ticaret ve Finansman", "Türkçe"), ("LOJ", "Lojistik Yönetimi", "Türkçe"),
            ("HAV", "Havacılık Yönetimi", "Türkçe"), ("YBS", "Yönetim Bilişim Sistemleri", "Türkçe"),
            ("ECF", "Economics and Finance", "English"), ("ETİC", "Elektronik Ticaret ve Yönetimi", "Türkçe")]
ISO = {g: "-".join(reversed(t.split("."))) for g, t in SINAV_TARIHLERI.items()}
KAPASITE = {"6A1": 24, "6A2": 24, "6D1": 15, "6D3": 15, "6D4": 20, "6D5": 20,
            "3A1": 80, "3A3": 80, "3A4": 40, "3A5": 70, "3A6": 60, "LAB": 40}


BOLUM_RENK = {"SBUİ": "#3f8f2a", "PSIR": "#2f75b5", "İŞL": "#b7791f", "BUS": "#e46c0a", "UTF": "#2f5fb3",
              "LOJ": "#6f2f9f", "HAV": "#c00000", "YBS": "#1f3b73", "ECF": "#548235", "ETİC": "#0089b8", "HAVUZ2": "#6b7785",
              "HAVUZ4": "#6b7785"}
GUN_RENK = {"Pazartesi": "#1769aa", "Salı": "#0f8a6b", "Çarşamba": "#6f2f9f", "Perşembe": "#d9730d", "Cuma": "#c0392b"}


def site_ek_css():
    bolum = "\n".join(f'    .department-button[data-code="{k}"], .program-tag[data-dept="{k}"] {{ --c: {v}; }}'
                      for k, v in BOLUM_RENK.items())
    gun = "\n".join(f'    .day-heading[data-day="{k}"] {{ --d: {v}; }}' for k, v in GUN_RENK.items())
    serit = ", ".join(f"{v} {i * 10}% {(i + 1) * 10}%" for i, v in enumerate(list(BOLUM_RENK.values())[:10]))
    return f"""
  <style>
  /* ---- renk katmanı (3_formlari_uret.py ekler; yalnız ekranda) ---- */
  @media screen {{
    .site-header, .access-gate {{ background: linear-gradient(120deg, #0d253d 0%, #1769aa 55%, #16917f 100%); }}
    .site-header::after {{ content: ""; display: block; height: 6px; background: linear-gradient(90deg, {serit}); }}
    .access-submit {{ background: #1769aa; }}
    .access-submit:hover {{ background: #0f8a6b; }}

    #tab-schedule, #page-schedule {{ --accent: #1769aa; }}
    #tab-teachers, #page-teachers {{ --accent: #0f8a6b; }}
    #tab-proctors, #page-proctors {{ --accent: #d9730d; }}
    .top-nav .tab-button {{ color: var(--accent); background: color-mix(in srgb, var(--accent) 9%, white);
      border: 2px solid color-mix(in srgb, var(--accent) 22%, white); }}
    .top-nav .tab-button:hover {{ color: var(--accent); background: color-mix(in srgb, var(--accent) 18%, white); }}
    .top-nav .tab-button[aria-selected="true"] {{ color: #fff; background: var(--accent); border-color: var(--accent);
      box-shadow: 0 5px 16px color-mix(in srgb, var(--accent) 35%, transparent); }}

    .page .page-heading {{ padding-left: 14px; border-left: 6px solid var(--accent); }}
    .page .page-heading h2 {{ color: var(--accent); }}
    .page .summary-pill {{ color: var(--accent); background: color-mix(in srgb, var(--accent) 10%, white);
      border-color: color-mix(in srgb, var(--accent) 35%, white); }}
    .page .exam-card {{ border-top: 4px solid var(--accent); }}
    .page .exam-time {{ color: var(--accent); }}
    .page .person-button[aria-pressed="true"] {{ color: var(--accent); background: color-mix(in srgb, var(--accent) 12%, white); }}
    .page .person-button[aria-pressed="true"] .person-count {{ background: var(--accent); }}
    .page .room-departments {{ color: var(--accent); font-weight: 700; }}

    .department-button, .program-tag {{ --c: #6b7785; }}
    .department-button {{ border-left: 6px solid var(--c); }}
{bolum}
    .department-grid .department-button:hover {{ border-color: var(--c); background: color-mix(in srgb, var(--c) 7%, white); }}
    .department-grid .department-button[aria-pressed="true"] {{ border-color: var(--c); background: color-mix(in srgb, var(--c) 15%, white);
      box-shadow: 0 6px 18px color-mix(in srgb, var(--c) 25%, transparent); }}
    .department-button .department-code {{ color: #fff; background: var(--c); }}
    .department-button .department-name {{ color: color-mix(in srgb, var(--c) 70%, #102a43); }}

    .program-tags .program-tag {{ color: var(--c); background: color-mix(in srgb, var(--c) 12%, white);
      border: 1px solid color-mix(in srgb, var(--c) 35%, white); }}

    .day-heading {{ --d: #1769aa; }}
{gun}
    .day-group .day-heading {{ padding: 8px 14px; border-radius: 10px; color: #fff;
      background: linear-gradient(90deg, var(--d), color-mix(in srgb, var(--d) 45%, white)); }}
    .day-group .day-heading::after {{ display: none; }}

    #tab-unit, #page-unit {{ --accent: #6f2f9f; }}
    .unit-lock {{ max-width: 360px; padding: 20px; border: 1px solid var(--line); border-radius: 14px; background: #fff; }}
    .unit-lock .access-submit {{ background: var(--accent); }}
    .unit-summary {{ margin: 0 0 12px; color: var(--accent); font-weight: 750; }}
    .unit-table-wrap {{ overflow-x: auto; border: 1px solid var(--line); border-radius: 14px; background: #fff; }}
    .unit-table {{ width: 100%; border-collapse: collapse; font-size: .9rem; }}
    .unit-table th {{ padding: 10px; color: #fff; background: var(--accent); text-align: left; white-space: nowrap; }}
    .unit-table td {{ padding: 9px 10px; border-top: 1px solid var(--line); }}
    .unit-table tr:nth-child(even) td {{ background: color-mix(in srgb, var(--accent) 5%, white); }}
    .unit-table .num {{ text-align: center; }}
    .unit-table .strong {{ font-weight: 800; color: var(--accent); }}
  }}
  </style>
"""


NOSCRIPT_UYARI = (
    '      <noscript><p class="access-error">Bu dosya JavaScript\'i kapatan bir önizleme uygulamasında açıldı '
    '(WhatsApp, Drive, Dosyalar vb.). Lütfen dosyayı Chrome veya Safari gibi bir tarayıcıda açın; şifre ancak '
    'tarayıcıda çalışır.</p></noscript>')

JS_UYARI = (
    '      <p class="access-error" id="js-uyari" style="color:#79520c;background:#fff7df;padding:10px;border-radius:8px">'
    '<strong>Şifre çalışmıyorsa:</strong> dosya bir önizleme ekranında açılmış olabilir (WhatsApp, Dosyalar, Drive). '
    'Android: sağ üstteki menüden <em>Chrome ile aç</em>. iPhone: yerel dosya Safari\'de açılamaz; bağlantı (web adresi) '
    'üzerinden açın.</p>')

BIRIM_SIFRE = BIRIM_SIFRESI   # donem_ayarlari.py
BIRIM_BOLUMU = """    <section class="page" id="page-unit" role="tabpanel" aria-labelledby="tab-unit" hidden>
      <div class="page-heading">
        <div>
          <h2>Birim</h2>
          <p>Gözetmenlerin görev sayıları ve dağılımı. Bu bölüm birim şifresiyle açılır.</p>
        </div>
      </div>
      <form class="unit-lock" id="unit-form">
        <label class="access-label" for="unit-password">Birim şifresi</label>
        <input class="access-input" id="unit-password" type="password" autocapitalize="off" autocorrect="off" spellcheck="false" autocomplete="off">
        <button class="access-submit" type="submit">Göster</button>
        <p class="access-error" id="unit-error" role="alert" aria-live="polite"></p>
      </form>
      <div id="unit-content" hidden></div>
    </section>
"""
BIRIM_JS = r"""    (function () {
      const UNIT_PASSWORD = "__SIFRE__";
      const form = $("#unit-form"), input = $("#unit-password"), error = $("#unit-error"), content = $("#unit-content");
      const GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"];
      function tablo() {
        const examById = {};
        DATA.exams.forEach(function (exam) { examById[exam.id] = exam; });
        const rows = DATA.proctors.map(function (name) {
          const own = DATA.assignments.filter(function (a) { return a.proctor === name; });
          const days = GUNLER.map(function (g) { return own.filter(function (a) { return a.day === g; }).length; });
          const eng = own.filter(function (a) {
            return (examById[a.examId].programs || []).some(function (p) { return /^(PSIR|BUS|ECF)-/.test(p); });
          }).length;
          const third = own.filter(function (a) { return /^3/.test(a.room); }).length;
          return { name: name, total: own.length, eng: eng, third: third, days: days };
        });
        const toplam = rows.reduce(function (t, r) { return t + r.total; }, 0);
        return '<div class="unit-summary">' + rows.length + ' gözetmen · ' + toplam + ' görev</div>' +
          '<div class="unit-table-wrap"><table class="unit-table"><thead><tr><th>Gözetmen</th><th>Toplam görev</th>' +
          '<th>İngilizce</th><th>3. kat</th>' + GUNLER.map(function (g) { return '<th>' + g + '</th>'; }).join("") +
          '</tr></thead><tbody>' + rows.map(function (r) {
            return '<tr><td>' + esc(r.name) + '</td><td class="num strong">' + r.total + '</td><td class="num">' + r.eng +
              '</td><td class="num">' + r.third + '</td>' +
              r.days.map(function (d) { return '<td class="num">' + (d || "") + '</td>'; }).join("") + '</tr>';
          }).join("") + '</tbody></table></div>';
      }
      form.addEventListener("submit", function (event) {
        event.preventDefault();
        if (input.value.replace(/\s+/g, "") === UNIT_PASSWORD) {
          error.textContent = ""; input.value = "";
          content.innerHTML = tablo(); content.hidden = false; form.hidden = true;
        } else {
          error.textContent = "Şifre hatalı.";
        }
      });
    })();
""".replace("__SIFRE__", BIRIM_SIFRE)


def birim_tablosu_html(data):
    """Birim sekmesi tablosu (sunucu tarafında hazırlanır; internet sürümünde şifrelenerek gömülür)."""
    import html as _html
    ing = {e["id"]: any(re.match(r"^(PSIR|BUS|ECF)-", p) for p in e["programs"]) for e in data["exams"]}
    satirlar, toplam = [], 0
    for ad in data["proctors"]:
        own = [a for a in data["assignments"] if a["proctor"] == ad]
        toplam += len(own)
        gunler = [sum(1 for a in own if a["day"] == g) for g in GUNLER]
        satirlar.append(
            f'<tr><td>{_html.escape(ad)}</td><td class="num strong">{len(own)}</td>'
            f'<td class="num">{sum(ing[a["examId"]] for a in own)}</td>'
            f'<td class="num">{sum(1 for a in own if a["room"].startswith("3"))}</td>'
            + "".join(f'<td class="num">{d or ""}</td>' for d in gunler) + "</tr>")
    return (f'<div class="unit-summary">{len(data["proctors"])} gözetmen · {toplam} görev</div>'
            '<div class="unit-table-wrap"><table class="unit-table"><thead><tr><th>Gözetmen</th><th>Toplam görev</th>'
            '<th>İngilizce</th><th>3. kat</th>' + "".join(f"<th>{g}</th>" for g in GUNLER) +
            "</tr></thead><tbody>" + "".join(satirlar) + "</tbody></table></div>")


def sifrele(metin_, sifre, tur=250000):
    """AES-256-GCM; anahtar PBKDF2-SHA256 ile şifreden türetilir. Tarayıcıda Web Crypto ile çözülür.
    Dosyada şifrenin kendisi bulunmaz; yanlış şifrede çözme başarısız olur."""
    import base64, hashlib
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    tuz, iv = os.urandom(16), os.urandom(12)
    anahtar = hashlib.pbkdf2_hmac("sha256", sifre.encode(), tuz, tur, dklen=32)
    veri = AESGCM(anahtar).encrypt(iv, metin_.encode("utf-8"), None)
    b64 = lambda b: base64.b64encode(b).decode()
    return {"tuz": b64(tuz), "iv": b64(iv), "veri": b64(veri), "tur": tur}


# tarayıcıda paket çözme (internet sürümünde sayfanın ilk script'ine eklenir)
PAKET_COZ_JS = r"""    async function paketCoz(paket, sifre) {
      function b64(s) { return Uint8Array.from(atob(s), function (c) { return c.charCodeAt(0); }); }
      try {
        const temel = await crypto.subtle.importKey("raw", new TextEncoder().encode(sifre), "PBKDF2", false, ["deriveKey"]);
        const anahtar = await crypto.subtle.deriveKey(
          { name: "PBKDF2", salt: b64(paket.tuz), iterations: paket.tur, hash: "SHA-256" },
          temel, { name: "AES-GCM", length: 256 }, false, ["decrypt"]);
        const acik = await crypto.subtle.decrypt({ name: "AES-GCM", iv: b64(paket.iv) }, anahtar, b64(paket.veri));
        return new TextDecoder().decode(acik);
      } catch (hata) {
        return null;   // yanlış şifre
      }
    }
    function sifreCozDestekli() { return !!(window.crypto && crypto.subtle); }
    const SIFRE_DESTEK_YOK = "Bu tarayıcı şifreli sayfayı açamıyor. Sayfayı Chrome veya Safari'de, https ile başlayan bağlantıdan açın.";
"""

BIRIM_JS_SIFRELI = r"""    (function () {
      const PAKET = __PAKET__;
      const form = $("#unit-form"), input = $("#unit-password"), error = $("#unit-error"), content = $("#unit-content");
      form.addEventListener("submit", async function (event) {
        event.preventDefault();
        if (!sifreCozDestekli()) { error.textContent = SIFRE_DESTEK_YOK; return; }
        error.textContent = "Kontrol ediliyor…";
        const tablo = await paketCoz(PAKET, input.value.replace(/\s+/g, ""));
        if (tablo === null) { error.textContent = "Şifre hatalı."; return; }
        error.textContent = ""; input.value = "";
        content.innerHTML = tablo; content.hidden = false; form.hidden = true;
      });
    })();
"""

# internet sürümü: sayfa şifresi = bütün verinin şifresi
GIRIS_ESKI = """    accessForm.addEventListener("submit", function (event) {
      event.preventDefault();
      if (accessPassword.value.replace(/\\s+/g, "") === ACCESS_PASSWORD) {
        try { sessionStorage.setItem(ACCESS_SESSION_KEY, "granted"); } catch (error) {}
        accessError.textContent = "";
        accessPassword.value = "";
        unlockPage();
        return;
      }
      accessError.textContent = "Şifre hatalı. Lütfen tekrar deneyin.";
      accessPassword.select();
    });"""
GIRIS_SIFRELI = """    accessForm.addEventListener("submit", async function (event) {
      event.preventDefault();
      if (!sifreCozDestekli()) { accessError.textContent = SIFRE_DESTEK_YOK; return; }
      accessError.textContent = "Açılıyor…";
      const acik = await paketCoz(DATA_PAKET, accessPassword.value.replace(/\\s+/g, ""));
      if (acik === null) {
        accessError.textContent = "Şifre hatalı. Lütfen tekrar deneyin.";
        accessPassword.select();
        return;
      }
      DATA = JSON.parse(acik);
      accessError.textContent = "";
      accessPassword.value = "";
      initApp();
      unlockPage();
    });"""


def web_sitesi(dersler, program):
    ada = {c["Ders"]: c for c in dersler}
    grup = collections.defaultdict(list)
    for y in program["yerlesim"]:
        grup[y["ders"]].append(y)
    sirali = sorted(grup, key=lambda d: (GUNLER.index(grup[d][0]["gun"]), grup[d][0]["baslangic"], d))
    exams, assignments = [], []
    for i, d in enumerate(sirali, start=1):
        c, ys = ada[d], sorted(grup[d], key=lambda y: ({"3": 0, "6": 1}.get(y["derslik"][0], 2), y["derslik"]))
        g, h = ys[0]["gun"], ys[0]["baslangic"]
        notlar = []
        if len(ys) > 1: notlar.append(f"{len(ys)} dersliğe bölündü")
        if c.get("_birlikte"): notlar.append("aynı derslikte birlikte: " + ", ".join(c["_birlikte"]))
        if c.get("_ayni_saat_es"): notlar.append("aynı saatte: " + c["_ayni_saat_es"])
        if sure_dk(c) > 50: notlar.append(f"{sure_dk(c)} dakika")
        ortak = dict(id=f"exam-{i}", date=ISO[g], dateLabel=TARIH[g], day=g, start=f"{h:02d}:00", end=bitis(h, sure_dk(c)),
                     courseCode=c["Ders Kodu"] or "", course=d, programs=c["Program-Yarıyıl (Çakışma Grubu)"].split("; "),
                     teacher=c["Öğretim Elemanı"] or "", totalStudents=str(c["Öğrenci Sayısı"] or ""))
        rooms = []
        for y in ys:
            proctor = ", ".join(x for x in y["gozetmenler"] if x) if c["Gözetmen İhtiyacı"] == "Var" else "Gerekmiyor"
            rooms.append({"room": y["derslik"], "capacity": str(KAPASITE.get(y["derslik"], "")), "students": str(y["ogrenci"]),
                          "proctor": proctor, "departments": ", ".join(f"{b} {v}" for b, v in y.get("bolumler", {}).items())})
            for ad in y["gozetmenler"]:
                if ad:
                    assignments.append({**ortak, "examId": ortak["id"], "room": y["derslik"], "roomStudents": str(y["ogrenci"]),
                                        "proctor": ad, "note": ""})
        exams.append({**ortak, "notes": notlar, "rooms": rooms})
    teachers = sorted({e["teacher"] for e in exams if e["teacher"]})
    proctors = sorted({a["proctor"] for a in assignments})
    data = {"meta": {"institution": "İstanbul Esenyurt Üniversitesi", "faculty": "İşletme ve Yönetim Bilimleri Fakültesi",
                     "period": DONEM_SURESI_METNI, "examCount": len(exams), "assignmentCount": len(assignments),
                     "teacherCount": len(teachers), "proctorCount": len(proctors)},
            "departments": [{"code": k, "name": n, "language": l} for k, n, l in BOLUMLER],
            "exams": exams, "assignments": assignments, "teachers": teachers, "proctors": proctors}
    h = open(SITE_SABLON, encoding="utf-8").read()
    bas = h.index("const DATA = ") + len("const DATA = ")
    son = h.index(";\n", bas)
    h = h[:bas] + json.dumps(data, ensure_ascii=False) + h[son:]
    degis = [
        ("__SITE_BASLIGI__ | İşletme", f"{SITE_BASLIGI} | İşletme"),
        ("__SITE_BASLIGI__ · __DONEM_SURESI__", f"{SITE_BASLIGI} · {DONEM_SURESI_METNI}"),
        ('const ACCESS_PASSWORD = "__SITE_SIFRESI__";', f'const ACCESS_PASSWORD = "{SITE_SIFRESI}";'),
        ('''const proctor = room.proctor === "Gerekmiyor" ? "Gözetmen gerekmiyor" : room.proctor;''',
         '''const proctor = room.proctor === "Gerekmiyor" ? "Dersin hocası (gözetmen gerekmiyor)" : room.proctor;'''),
        ('''          '<div class="room-capacity"><span class="detail-label">Kapasite</span>' + capacityText + '</div>' +''',
         '''          '<div class="room-capacity"><span class="detail-label">Kapasite</span>' + capacityText + '</div>' +
          (room.departments ? '<div class="room-departments"><span class="detail-label">Bu derslikteki bölümler</span>' + esc(room.departments) + '</div>' : '') +'''),
    ]
    degis += [
        ("""        return '<span class="program-tag">' + esc(program) + '</span>';""",
         """        return '<span class="program-tag" data-dept="' + esc(program.split(/[-(]/)[0]) + '">' + esc(program) + '</span>';"""),
        ("""'<h4 class="day-heading">' + esc(group.dateLabel)""",
         """'<h4 class="day-heading" data-day="' + esc(group.day) + '">' + esc(group.dateLabel)"""),
        ("</head>", site_ek_css() + "</head>"),
        # telefon: şifredeki boşlukları temizle; JavaScript kapalıysa uyarı göster
        ('if (accessPassword.value === ACCESS_PASSWORD) {',
         'if (accessPassword.value.replace(/\\s+/g, "") === ACCESS_PASSWORD) {'),
        ('      <p class="access-error" id="access-error" role="alert" aria-live="polite"></p>',
         '      <p class="access-error" id="access-error" role="alert" aria-live="polite"></p>\n' + NOSCRIPT_UYARI + '\n' + JS_UYARI),
        # JavaScript çalışırsa uyarıyı gizle (önizleme uygulamaları <noscript> göstermeyebilir)
        ('const accessError = document.querySelector("#access-error");',
         'const accessError = document.querySelector("#access-error");\n'
         '    (function () { var u = document.getElementById("js-uyari"); if (u) u.hidden = true; })();'),
    ]
    for a, b in degis:
        assert h.count(a) == 1, a[:50]
        h = h.replace(a, b)

    # 1) tam sürüm (bilgisayarda kalır): şifreli Birim sekmesi dahil
    birim_degis = [
        ('data-page="proctors">Gözetmenler</button>',
         'data-page="proctors">Gözetmenler</button>\n      <button class="tab-button" id="tab-unit" role="tab" '
         'aria-controls="page-unit" aria-selected="false" data-page="unit">Birim</button>'),
        ("  </main>", BIRIM_BOLUMU + "  </main>"),
        ('const hashes = { schedule: "sinav-programi", teachers: "ogretim-uyeleri", proctors: "gozetmenler" };',
         'const hashes = { schedule: "sinav-programi", teachers: "ogretim-uyeleri", proctors: "gozetmenler", unit: "birim" };'),
        ('    $("#schedule-summary").textContent', BIRIM_JS + '    $("#schedule-summary").textContent'),
    ]
    tam = h
    for a, b in birim_degis:
        assert tam.count(a) == 1, a[:50]
        tam = tam.replace(a, b)
    open(SITE_HTML, "w", encoding="utf-8").write(tam)

    # 2) internet sürümü (GitHub Pages): BÜTÜN veri sayfa şifresiyle, Birim tablosu birim şifresiyle şifreli.
    #    Dosyada ne şifreler ne isimler açık yazılıdır; doğru şifre girilince tarayıcı çözer.
    tablo_html = birim_tablosu_html(data)
    birim_js = BIRIM_JS_SIFRELI.replace("__PAKET__", json.dumps(sifrele(tablo_html, BIRIM_SIFRESI)))
    veri_paketi = json.dumps(sifrele(json.dumps(data, ensure_ascii=False), SITE_SIFRESI))
    yayin = h
    bas = yayin.index("const DATA = "); son = yayin.index(";\n", bas)
    yayin = yayin[:bas] + "let DATA = null;   // şifre girilince DATA_PAKET çözülür" + yayin[son + 1:]
    son_script = yayin.rindex("  </script>")
    yayin = yayin[:son_script] + "    }\n" + yayin[son_script:]          # initApp kapanışı
    internet_degis = birim_degis[:3] + [
        ('    $("#schedule-summary").textContent', birim_js + "    function initApp() {\n" + '    $("#schedule-summary").textContent'),
        ("  <script>\n", "  <script>\n" + PAKET_COZ_JS, 1),
        (f'const ACCESS_PASSWORD = "{SITE_SIFRESI}";', f"const DATA_PAKET = {veri_paketi};"),
        ('if (sessionStorage.getItem(ACCESS_SESSION_KEY) === "granted") unlockPage();',
         "/* internet sürümü: her açılışta şifre sorulur */"),
        (GIRIS_ESKI, GIRIS_SIFRELI),
        ('<meta charset="utf-8">', '<meta charset="utf-8">\n  <meta name="robots" content="noindex, nofollow">'),
    ]
    for d in internet_degis:
        a, b = d[0], d[1]
        adet = d[2] if len(d) > 2 else None
        if adet is None:
            assert yayin.count(a) == 1, a[:60]
            yayin = yayin.replace(a, b)
        else:
            assert yayin.count(a) >= 1, a[:60]
            yayin = yayin.replace(a, b, adet)
    # dosyada açık yazılı şifre ve kişi adı kalmamalı (alt satırdaki geliştirici teşekkürü hariç)
    tesekkur = re.search(r'<div class="footer-inner">.*?</div>', yayin, re.S).group(0)
    kontrol = yayin.replace(tesekkur, "")
    for yasak in [SITE_SIFRESI, BIRIM_SIFRESI, "UNIT_PASSWORD", "Toplam görev"] + data["proctors"] + data["teachers"][:40]:
        assert yasak not in kontrol, f"internet sürümünde açık kaldı: {yasak}"
    os.makedirs(os.path.dirname(SITE_YAYIN_HTML), exist_ok=True)
    open(SITE_YAYIN_HTML, "w", encoding="utf-8").write(yayin)
    return len(exams), len(assignments)


if __name__ == "__main__":
    os.makedirs(CIKTI, exist_ok=True)
    program, dersler, yer = verileri_oku()
    yazilan, eksik = yayin_formu(dersler, yer)
    print(f"Yayın formu: {len(yazilan)} sınav yazıldı; programda olup formda bulunamayan: {eksik or 'yok'}")
    if "--pdf" in sys.argv:
        ok, hata = pdf_uret()
        print("PDF:", "üretildi" if ok else f"üretilemedi {hata}")
    n, eks = derslik_formu(dersler, program)
    print(f"Derslik formu (6. kat): {n} derslik-saat hücresi dolduruldu; formda karşılığı olmayan: {eks or 'yok'}")
    yuk = gozetmen_cizelgesi(dersler, program)
    print("Gözetmenlik çizelgesi:", sum(yuk.values()), "görev |", ", ".join(f"{k.split()[0]} {v}" for k, v in yuk.items()))
    ne, na = web_sitesi(dersler, program)
    print(f"Web sitesi: {ne} sınav, {na} gözetmen ataması → {os.path.basename(SITE_HTML)} (tam) + "
          f"GitHub Pages/index.html (internet sürümü, Birim sekmesiz)")
