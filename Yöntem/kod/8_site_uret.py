"""Sınav programı + gözetmenlik web sitesi (ŞİFRESİZ). Şablon: Yöntem/site_sablonu.html

Girdi : Sınav Çıktıları/Aşama 2 - Derslik Dosyası (6. Kat) ... .xlsx (6. kat programı)
        + okul dosyasındaki büyük salon/lab sınavları (aşağıda EKSTRA) + Sınav Çıktıları/Gözetmen Atamaları - Liste N ... .xlsx
Çıktı : Sınav Çıktıları/GitHub Pages/Liste N/index.html  (GitHub Pages'e index.html olarak konur)
Çalıştırma: py -3 Yöntem\\kod\\8_site_uret.py
"""
import ast
import collections
import importlib.util
import html
import json
import os
import re
import sys
import unicodedata

import openpyxl

KOD = os.path.dirname(os.path.abspath(__file__))
KOK = os.path.dirname(os.path.dirname(KOD))
KLASOR = os.path.join(KOK, "Sınav Çıktıları")
nfc = lambda s: unicodedata.normalize("NFC", s)
sys.path.insert(0, KOD)


def yukle(ad, dosya):
    sp = importlib.util.spec_from_file_location(ad, os.path.join(KOD, dosya))
    m = importlib.util.module_from_spec(sp)
    sp.loader.exec_module(m)
    return m


y = yukle("y5", "5_yerlestir.py")
b6 = yukle("b6", "6_bolum_programi.py")

GUNLER = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma"]
ISO = ["2026-11-02", "2026-11-03", "2026-11-04", "2026-11-05", "2026-11-06"]
TARIH = ["02.11.2026", "03.11.2026", "04.11.2026", "05.11.2026", "06.11.2026"]
SITE_BASLIGI = "2026-2027 Güz Dönemi Vize Sınav Programı"
DONEM = "2–6 Kasım 2026"
BOLUMLER = [("SBUİ", "Siyaset Bilimi ve Uluslararası İlişkiler", "Türkçe"),
            ("PSIR", "Political Science and International Relations", "English"),
            ("İŞL", "İşletme", "Türkçe"), ("BUS", "Business Administration", "English"),
            ("UTF", "Uluslararası Ticaret ve Finansman", "Türkçe"), ("LOJ", "Lojistik Yönetimi", "Türkçe"),
            ("HAV", "Havacılık Yönetimi", "Türkçe"), ("YBS", "Yönetim Bilişim Sistemleri", "Türkçe"),
            ("ECF", "Economics and Finance", "English"), ("ETİC", "Elektronik Ticaret ve Yönetimi", "Türkçe")]
BOLUM_RENK = {"SBUİ": "#3f8f2a", "PSIR": "#2f75b5", "İŞL": "#b7791f", "BUS": "#e46c0a", "UTF": "#2f5fb3",
              "LOJ": "#6f2f9f", "HAV": "#c00000", "YBS": "#1f3b73", "ECF": "#548235", "ETİC": "#0089b8"}
GUN_RENK = {"Pazartesi": "#1769aa", "Salı": "#0f8a6b", "Çarşamba": "#6f2f9f", "Perşembe": "#d9730d", "Cuma": "#c0392b"}
KOD_PROG = {"Siyaset Bilimi ve Uluslararası İlişkiler": "SBUİ", "Siyaset Bilimi ve Uluslararası İlişkiler (İngilizce)": "PSIR",
            "İşletme": "İŞL", "İşletme (İngilizce)": "BUS", "Uluslararası Ticaret ve Finansman": "UTF", "Lojistik Yönetimi": "LOJ",
            "Havacılık Yönetimi": "HAV", "Yönetim Bilişim Sistemleri": "YBS", "Ekonomi ve Finans (İngilizce)": "ECF",
            "Elektronik Ticaret ve Yönetimi": "ETİC"}
# salon sınav kapasiteleri: 6. kat dosyadan, 3./4. kat kullanıcının verdiği değerler
KAPASITE = {**y.DERSLIKLER, "3A1": 90, "3A3": 85, "3A5": 60, "3A6": 50, "4A3": 80, "4A4": 50}
# okul dosyasında (3./4./11. kat) yapılan ve 6.Kat sayfasında olmayan sınavlar: ad -> (gün, saat, süre_dk, {derslik: [(program, sayı)]})
EKSTRA = y.EKSTRA   # 5_yerlestir.py'de tanımlı (bölüm programı ve çakışma kontrolü de kullanır)


def bitis(h, dk):
    t = h * 60 + dk
    return f"{t // 60:02d}:{t % 60:02d}"


def temiz(a):
    return a.replace("Dept Elective: ", "").strip()


def atamalar_oku(yol):
    """(ders, gün, başlangıç saati) -> [(derslik, kişi)]"""
    ws = openpyxl.load_workbook(yol)["Atamalar"]
    sonuc = collections.defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r[1]:
            continue
        sonuc[(r[1], GUNLER.index(r[2]), int(r[4][:2]))].append((r[5], r[6]))
    return sonuc


def site_ek_css():
    kaynak = open(os.path.join(KOD, "3_formlari_uret.py"), encoding="utf-8").read()
    agac = ast.parse(kaynak)
    fn = next(n for n in agac.body if isinstance(n, ast.FunctionDef) and n.name == "site_ek_css")
    ns = {"BOLUM_RENK": BOLUM_RENK, "GUN_RENK": GUN_RENK}
    exec(compile(ast.Module([fn], []), "css", "exec"), ns)
    return ns["site_ek_css"]()


def sinavlari_topla():
    sinavlar, izin, kt = y.sinavlari_kur()
    f = os.path.join(KLASOR, next(x for x in os.listdir(KLASOR) if nfc(x).startswith("Aşama 2 - Derslik")))
    yuk, ua = y.dosyadan_yukle(f, sinavlar)
    L = []
    for e in sinavlar:
        ad = temiz(e["ad"])
        if id(e) in yuk:
            s, o, es = yuk[id(e)]
            g, h = divmod(s, len(y.SAATLER))
            L.append((e, g, h + y.SAATLER[0], y.SURE_DK.get(ad, 90 if e["sure"] > 50 else 50), o))
        else:
            for k, (g, h, dk, o) in EKSTRA.items():
                if ad.startswith(k):
                    L.append((e, g, h, dk, o))
    # Sektör Kampüste dersi öğrencisiz (n=0): programda tek satır olarak kalsın
    return L, sinavlar


def program_kodlari(e, bilgi):
    sinif, alan = bilgi[id(e)]
    out = []
    for p, sn in sinif.items():
        kod = KOD_PROG.get(p)
        if not kod:
            continue
        for k in sorted(sn):
            out.append(f"{kod}-{2 * k - 1}" if k else f"{kod}-?")
    return sorted(out)


def veri_kur(L, sinavlar, atama):
    yerli = [e for e, *_ in L]
    bilgi, _ = b6.sinif_bilgisi(yerli)
    sira = sorted(L, key=lambda t: (t[1], t[2], temiz(t[0]["ad"])))
    exams, assignments = [], []
    for i, (e, g, h, dk, o) in enumerate(sira, 1):
        ad = temiz(e["ad"])
        gorevler = atama.get((ad, g, h), [])
        ortak = dict(id=f"exam-{i}", date=ISO[g], dateLabel=TARIH[g], day=GUNLER[g], start=f"{h:02d}:00", end=bitis(h, dk),
                     courseCode=(e["kodlar"][0] if e["kodlar"] else ""), course=ad, programs=program_kodlari(e, bilgi),
                     teacher=", ".join(sorted({a for _, a in e["hocalar"]})), totalStudents=str(sum(n for l in o.values() for _, n in l) or ""))
        odalar = sorted(o, key=lambda r: (not r.startswith("3"), r))
        rooms = []
        for r in odalar:
            kisiler = [k for oda, k in gorevler if oda == r]
            kendi = [k for oda, k in gorevler if oda == "(kendi sınavı)"]
            if kendi and not kisiler and len(odalar) == 1:
                kisiler = [f"{k} (dersin hocası)" for k in kendi]
            bolumler = ", ".join(f"{KOD_PROG.get(p, p)} {n}" if n else f"{KOD_PROG.get(p, p)}" for p, n in o[r])
            rooms.append({"room": r, "capacity": str(KAPASITE.get(r, "")), "students": str(sum(n for _, n in o[r]) or ""),
                          "proctor": ", ".join(kisiler) if kisiler else "Gerekmiyor", "departments": bolumler})
            for k in [k for oda, k in gorevler if oda == r] + [k for oda, k in gorevler if oda == "(kendi sınavı)" and len(odalar) == 1]:
                assignments.append({**ortak, "examId": ortak["id"], "room": r, "roomStudents": str(sum(n for _, n in o[r])),
                                    "proctor": k, "note": ""})
        notlar = []
        if len(odalar) > 1:
            notlar.append(f"{len(odalar)} dersliğe bölündü")
        if dk > 50:
            notlar.append(f"{dk} dakika")
        exams.append({**ortak, "notes": notlar, "rooms": rooms})
    # İYBF'nin başka birimlere verdiği gözetmenlikler (Sağlık Bilimleri Fakültesi, üniversite ortak seçmeli dersleri): "başka birim" kaydı olarak girer
    n_iybf = len(exams)
    DIS = (("Sağlık Bil. Fak. ", "Sağlık Bilimleri Fakültesi", "Başka fakülte: İYBF'den verilen gözetmen"),
           ("Ortak Seçmeli: ", "Üniversite ortak seçmeli ders", "Üniversite ortak seçmeli: İYBF'den verilen gözetmen"))
    for (ad, g, h), gor in sorted(atama.items(), key=lambda t: (t[0][1], t[0][2], t[0][0])):
        dis = next((d for d in DIS if ad.startswith(d[0])), None)
        if not dis:
            continue
        no = len(exams) + 1
        ders = ad[len(dis[0]):]
        ortak = dict(id=f"exam-{no}", date=ISO[g], dateLabel=TARIH[g], day=GUNLER[g], start=f"{h:02d}:00", end=bitis(h, 60), courseCode="",
                     course=ders + f" ({dis[1]})", programs=[], teacher="", totalStudents="")
        odalar = []
        for oda, k in gor:
            if oda not in odalar:
                odalar.append(oda)
            assignments.append({**ortak, "examId": ortak["id"], "room": oda, "roomStudents": "", "proctor": k, "note": ""})
        exams.append({**ortak, "notes": [dis[2], "60 dakika"],
                      "rooms": [{"room": r, "capacity": "", "students": "", "proctor": ", ".join(k for o_, k in gor if o_ == r), "departments": dis[1]} for r in odalar]})
    exams.sort(key=lambda e: (e["date"], e["start"]))
    teachers = sorted({e["teacher"] for e in exams if e["teacher"]})
    proctors = sorted({a["proctor"] for a in assignments})
    return {"meta": {"institution": "İstanbul Esenyurt Üniversitesi", "faculty": "İşletme ve Yönetim Bilimleri Fakültesi",
                     "period": DONEM, "examCount": n_iybf, "assignmentCount": len(assignments),
                     "teacherCount": len(teachers), "proctorCount": len(proctors)},
            "departments": [{"code": k, "name": n, "language": l} for k, n, l in BOLUMLER],
            "exams": exams, "assignments": assignments, "teachers": teachers, "proctors": proctors}


def degisiklik_blogu():
    """Ana sayfada 'Son değişiklikler' kutusu (Yöntem/degisiklikler.json, en yeni üstte). Yalnızca sınav ve gözetmen değişiklikleri, teknik ayrıntı yok."""
    yol = os.path.join(KOK, "Yöntem", "degisiklikler.json")
    liste = json.load(open(yol, encoding="utf-8")) if os.path.exists(yol) else []
    if not liste:
        return ""
    def grup(girdiler):
        out, son = [], None
        for g in girdiler:
            if g["tarih"] != son:
                out.append(f'<h4>{html.escape(g["tarih"])}</h4>')
                son = g["tarih"]
            out.append(f'<p><span class="dg-tur dg-{"sinav" if g["tur"] == "Sınav" else "goz"}">{html.escape(g["tur"])}</span> {html.escape(g["metin"])}</p>')
        return "".join(out)
    ilk_tarih = liste[0]["tarih"]
    # en son güncelleme gününün değişiklikleri açık, öncekiler "Önceki değişiklikler" altında
    yeni = [g for g in liste if g["tarih"] == ilk_tarih]
    eski = [g for g in liste if g["tarih"] != ilk_tarih]
    return ('      <section class="degisiklik" aria-label="Son değişiklikler"><h3>Son değişiklikler</h3>' + grup(yeni) +
            (f'<details><summary>Önceki değişiklikler ({len(eski)})</summary>{grup(eski)}</details>' if eski else "") + '</section>\n')


DEGISIKLIK_CSS = ("<style>.degisiklik{margin:0 0 18px;padding:14px 18px;background:var(--surface,#fff);border:1px solid var(--line);border-left:5px solid var(--blue);"
                  "border-radius:12px;box-shadow:var(--shadow)}.degisiklik h3{margin:0 0 6px;font-size:1.02rem;color:var(--navy)}"
                  ".degisiklik h4{margin:10px 0 2px;font-size:.82rem;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}"
                  ".degisiklik p{margin:3px 0;font-size:.92rem;line-height:1.45}.degisiklik summary{cursor:pointer;margin-top:10px;color:var(--blue);font-weight:700;font-size:.9rem}"
                  ".dg-tur{display:inline-block;margin-right:6px;padding:1px 8px;border-radius:999px;font-size:.72rem;font-weight:800}"
                  ".dg-sinav{background:var(--blue-soft);color:var(--blue)}.dg-goz{background:#e8f6ee;color:#1d6b43}</style>")
PANEL_SURUM =__import__("time").strftime("%Y%m%d%H%M%S")   # panel adresine eklenir: tarayıcı eski paneli önbellekten göstermesin
PANEL_CSS = ("<style>#panel-frame{display:block;width:100%;height:calc(100vh - 170px);min-height:620px;border:1px solid var(--line,#d9e0e8);"
             "border-radius:12px;background:transparent}</style>")


def html_uret(data):
    h = open(os.path.join(KOK, "Yöntem", "site_sablonu.html"), encoding="utf-8").read()
    bas = h.index("const DATA = ") + len("const DATA = ")
    son = h.index(";\n", bas)
    h = h[:bas] + json.dumps(data, ensure_ascii=False) + h[son:]
    degis = [
        ("__SITE_BASLIGI__ | İşletme", f"{SITE_BASLIGI} | İşletme"),
        ("__SITE_BASLIGI__ · __DONEM_SURESI__", f"{SITE_BASLIGI} · {DONEM}"),
        # JavaScript çalışmazsa (WhatsApp/Dosyalar/Drive önizlemesi) görünen uyarı; script başlar başlamaz gizlenir
        ('const ACCESS_PASSWORD = "__SITE_SIFRESI__";',
         '(function () { var u = document.getElementById("js-uyari"); if (u) u.hidden = true; })();\n    const ACCESS_PASSWORD = "";'),
        ("  <main>\n", '  <p id="js-uyari" style="margin:0;padding:12px 16px;color:#79520c;background:#fff7df;text-align:center;font-weight:600">'
         'Bölümler ve program görünmüyorsa bu sayfa bir önizleme ekranında (WhatsApp, Dosyalar, Drive) açılmıştır. '
         'Lütfen tarayıcıdan şu adresi açın: <a href="https://iybf-sinav.github.io/">iybf-sinav.github.io</a></p>\n  <main>\n'),
        ('<body class="access-locked">', "<body>"),                      # şifre ekranı yok
        ('''const proctor = room.proctor === "Gerekmiyor" ? "Gözetmen gerekmiyor" : room.proctor;''',
         '''const proctor = room.proctor === "Gerekmiyor" ? "Dersin hocası (gözetmen gerekmiyor)" : room.proctor;'''),
        ('''          '<div class="room-capacity"><span class="detail-label">Kapasite</span>' + capacityText + '</div>' +''',
         '''          '<div class="room-capacity"><span class="detail-label">Kapasite</span>' + capacityText + '</div>' +
          (room.departments ? '<div class="room-departments"><span class="detail-label">Bu derslikteki bölümler</span>' + esc(room.departments) + '</div>' : '') +'''),
        ("""        return '<span class="program-tag">' + esc(program) + '</span>';""",
         """        return '<span class="program-tag" data-dept="' + esc(program.split(/[-(]/)[0]) + '">' + esc(program) + '</span>';"""),
        ("""'<h4 class="day-heading">' + esc(group.dateLabel)""",
         """'<h4 class="day-heading" data-day="' + esc(group.day) + '">' + esc(group.dateLabel)"""),
        ("</head>", site_ek_css() + PANEL_CSS + DEGISIKLIK_CSS + "</head>"),
        ('      <div class="department-grid" id="department-grid"></div>\n', degisiklik_blogu() + '      <div class="department-grid" id="department-grid"></div>\n'),
        # Panel sekmesi: şifreli komisyon paneli (panel/index.html, 9_panel_uret.py üretir) iframe olarak açılır
        ('''data-page="proctors">Gözetmenler</button>\n''',
         '''data-page="proctors">Gözetmenler</button>\n      <button class="tab-button" id="tab-panel" role="tab" aria-controls="page-panel" aria-selected="false" data-page="panel">Panel</button>\n'''),
        ("    </section>\n  </main>\n",
         '''    </section>\n\n    <section class="page" id="page-panel" role="tabpanel" aria-labelledby="tab-panel" hidden>\n'''
         '''      <iframe id="panel-frame" title="Panel" loading="lazy"></iframe>\n    </section>\n  </main>\n'''),
        ("grid-template-columns: repeat(3, 1fr); gap: 4px; }\n      .tab-button { padding: 10px 6px;", "grid-template-columns: repeat(4, 1fr); gap: 4px; }\n      .tab-button { padding: 10px 6px;"),
        ('''proctors: "gozetmenler" };''', '''proctors: "gozetmenler", panel: "panel" };'''),
        ('''      "gozetmenler": "proctors"\n''', '''      "gozetmenler": "proctors",\n      "panel": "panel"\n'''),
        ('''        page.hidden = page.id !== "page-" + pageName;\n      });\n''',
         '''        page.hidden = page.id !== "page-" + pageName;\n      });\n'''
         '''      if (pageName === "panel") { const f = document.getElementById("panel-frame"); if (!f.getAttribute("src")) f.setAttribute("src", "panel/?v=''' + PANEL_SURUM + '''"); }\n'''),
        ('<meta charset="utf-8">', '<meta charset="utf-8">\n  <meta name="robots" content="noindex, nofollow">'),
    ]
    for a, c in degis:
        assert h.count(a) == 1, a[:60]
        h = h.replace(a, c)
    return h


if __name__ == "__main__":
    L, sinavlar = sinavlari_topla()
    print(len(L), "sınav")
    for dosya in sorted(x for x in os.listdir(KLASOR) if nfc(x).startswith("Gözetmen Atamaları - Liste")):
        no = re.search(r"Liste (\d)", dosya).group(1)
        atama = atamalar_oku(os.path.join(KLASOR, dosya))
        data = veri_kur(L, sinavlar, atama)
        klasor = os.path.join(KLASOR, "GitHub Pages", f"Liste {no}")
        os.makedirs(klasor, exist_ok=True)
        yol = os.path.join(klasor, "index.html")
        open(yol, "w", encoding="utf-8").write(html_uret(data))
        print("->", yol, "|", data["meta"]["examCount"], "sınav,", data["meta"]["assignmentCount"], "görev,", data["meta"]["proctorCount"], "gözetmen,", data["meta"]["teacherCount"], "öğretim üyesi")
