"""Genereert OpdrachtZeker-financieel-model.xlsx.

Strategie: eerst GRATIS BETA, daarna abonnementen zodra er betalende mensen zijn.
Alle maandcellen zijn ECHTE Excel-formules die verwijzen naar blad 'Aannames'.
Wijzig je een aanname, dan herberekent het hele model bij het openen in Excel.
Onderaan draait een Python-mirror van dezelfde logica ter verificatie.
Afronding: half-up (gelijk aan Excel ROUND), zodat mirror == xlsx == dashboard.
"""

import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

OUT = "OpdrachtZeker-financieel-model.xlsx"


def rnd(x):
    """Half-up afronding (Excel ROUND), voor positieve en negatieve waarden."""
    return math.floor(x + 0.5) if x >= 0 else math.ceil(x - 0.5)


# ---------------- Aannames (base case) ----------------
A = {
    "startkapitaal": (75000, "EUR", "Eigen middelen + kleine lening"),
    "arpu": (230, "EUR/mnd", "Gemiddeld abonnement opdrachtgever (mix 60% x 149 / 30% x 299 / 10% x 499)"),
    "pro_prijs": (12, "EUR/mnd", "ZZP Pro-abonnement"),
    "conversie_pro": (0.04, "%", "Van nieuwe gratis ZZP-leden naar betalend Pro"),
    "churn": (0.03, "%/mnd", "Maandelijkse churn opdrachtgevers"),
    "churn_pro": (0.02, "%/mnd", "Maandelijkse churn ZZP Pro"),
    "cac": (600, "EUR", "Acquisitiekosten per opdrachtgever (marketing)"),
    "sales_new": (2, "per mnd", "Nieuwe opdrachtgevers via directe sales (vanaf mnd 7)"),
    "commissie": (15, "EUR/mnd", "Commissie-inkomsten per actieve opdrachtgever (vanaf mnd 4)"),
    "salaris": (5000, "EUR/mnd", "Oprichtersalaris 2 x 2.500 (bootstrapped)"),
    "sales_hire": (4500, "EUR/mnd", "Sales-hire (vanaf mnd 7)"),
    "cloud": (500, "EUR/mnd", "Cloud/AI basiskosten"),
    "cloud_klant": (2, "EUR/mnd", "Cloud/AI per actieve opdrachtgever"),
    "tooling": (400, "EUR/mnd", "Tooling/kantoor/verzekering"),
    "juridisch": (2500, "EUR", "Eenmalige juridische setup (maand 2)"),
    "mkt1": (1000, "EUR/mnd", "Marketingbudget maand 1-3"),
    "mkt2": (2000, "EUR/mnd", "Marketingbudget maand 4-6"),
    "mkt3": (3000, "EUR/mnd", "Marketingbudget maand 7-12"),
    "marge": (0.8, "%", "Bruto-marge (voor LTV)"),
    "beta_einde": (3, "maand", "Laatste gratis beta-maand (0 = meteen betaald)"),
    "beta_boost": (1.5, "x", "Groei-boost nieuwe klanten tijdens gratis beta"),
}
ORDER = ["startkapitaal", "arpu", "pro_prijs", "conversie_pro", "churn", "churn_pro", "cac",
         "sales_new", "commissie", "salaris", "sales_hire", "cloud", "cloud_klant", "tooling",
         "juridisch", "mkt1", "mkt2", "mkt3", "marge", "beta_einde", "beta_boost"]
AR = {key: i + 3 for i, key in enumerate(ORDER)}  # rijnummer op blad Aannames (kolom B)

NIEUWE_FREE = [50, 70, 90, 110, 140, 170, 200, 240, 280, 320, 370, 420]

EUR = '€ #,##0'
EUR_M = '€ #,##0;[Rood]-€ #,##0'
PCT = '0%'

wb = Workbook()

hdr_font = Font(bold=True, size=12)
title_font = Font(bold=True, size=14)
input_fill = PatternFill("solid", fgColor="FFF2CC")
hdr_fill = PatternFill("solid", fgColor="D9E1F2")


def style_header(ws, row, cols, fill=hdr_fill):
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.font = hdr_font
        cell.fill = fill


# ================= Blad 1: Leeswijzer =================
ws = wb.active
ws.title = "Leeswijzer"
ws["A1"] = "OpdrachtZeker — financieel model (12 maanden)"
ws["A1"].font = title_font
lines = [
    "",
    "Strategie: eerst GRATIS BETA, daarna abonnementen zodra er betalende mensen zijn.",
    "",
    "Hoe werkt dit model?",
    "",
    "1. Blad 'Aannames' is de ENIGE plek waar je invoer wijzigt (gele cellen).",
    "   - 'Laatste gratis beta-maand': tot en met deze maand is het product gratis (geen omzet).",
    "   - 'Groei-boost': gratis beta trekt sneller nieuwe klanten aan.",
    "2. Blad 'Model' rekent 12 maanden door: klanten, MRR, kosten, resultaat en kaspositie.",
    "   Alle maandcellen zijn formules en herberekenen automatisch in Excel/Sheets.",
    "3. Blad 'Scenario' vergelijkt pessimistisch / base / optimistisch, inclusief",
    "   eerste break-even maand en minimaal benodigd startkapitaal.",
    "4. Blad 'Fasen' koppelt de modeluitkomsten aan de fasen uit het businessplan.",
    "",
    "Let op: maand 1 = start van de verkoop (na validatiefase). Alle bedragen in euro's.",
    "Afronding is half-up (gelijk aan Excel ROUND), zodat xlsx, dashboard en mirror kloppen.",
    "Dit is een prognose- en beslissingsinstrument, geen garantie.",
]
for i, t in enumerate(lines, start=2):
    ws.cell(row=i, column=1, value=t)
ws.column_dimensions["A"].width = 100

# ================= Blad 2: Aannames =================
wa = wb.create_sheet("Aannames")
wa["A1"] = "Aannames (base case) — gele cellen mag je aanpassen"
wa["A1"].font = title_font
for j, h in enumerate(["Parameter", "Waarde", "Eenheid", "Toelichting"], start=1):
    wa.cell(row=2, column=j, value=h)
style_header(wa, 2, [1, 2, 3, 4])
for key in ORDER:
    r = AR[key]
    val, unit, desc = A[key]
    wa.cell(row=r, column=1, value=key.replace("_", " ").capitalize())
    c = wa.cell(row=r, column=2, value=val)
    c.fill = input_fill
    c.number_format = PCT if unit.startswith("%") else (EUR if "EUR" in unit else "0.0")
    wa.cell(row=r, column=3, value=unit)
    wa.cell(row=r, column=4, value=desc)
wa.column_dimensions["A"].width = 38
wa.column_dimensions["B"].width = 12
wa.column_dimensions["C"].width = 12
wa.column_dimensions["D"].width = 70


def aref(key):
    return f"Aannames!$B${AR[key]}"


# ================= Blad 3: Model =================
wm = wb.create_sheet("Model")
wm["A1"] = "Model — 12 maanden (maand 1 = start verkoop; maand 1 t/m beta = gratis)"
wm["A1"].font = title_font
wm["B2"] = "Formules op basis van blad Aannames"
wm["B2"].font = Font(italic=True)

MAAND, PFLAG, MKT, NEWFREE, NEWCO, SALESROW, CHURNCO, ACTCO = 3, 4, 5, 6, 7, 8, 9, 10
FREE, NEWPRO, CHURNPRO, ACTPRO = 11, 12, 13, 14
MRR_CO, MRR_PRO, COMM, OMZET = 15, 16, 17, 18
K_PERS, K_MKT, K_TECH, K_OVER, K_TOT = 19, 20, 21, 22, 23
RESULT, KAS, KAS_Z, RUNWAY, CACR, LTV, LTVCAC, POS = 24, 25, 26, 27, 28, 29, 30, 31
S_TITLE, S_MRR, S_BE, S_KAS, S_DIP, S_MIN = 33, 34, 35, 36, 37, 38

labels = {
    MAAND: "Maand", PFLAG: "Beta-fase? (1 = gratis, 0 = betaald)",
    MKT: "INPUT Marketingbudget", NEWFREE: "INPUT Nieuwe gratis ZZP-leden",
    NEWCO: "Nieuwe opdrachtgevers (marketing)", SALESROW: "INPUT Nieuwe opdrachtgevers (directe sales)",
    CHURNCO: "Churn opdrachtgevers", ACTCO: "Actieve opdrachtgevers (eind)",
    FREE: "Cumulatief gratis ZZP-leden", NEWPRO: "Nieuwe ZZP Pro", CHURNPRO: "Churn ZZP Pro",
    ACTPRO: "ZZP Pro abonnees (eind)", MRR_CO: "MRR opdrachtgevers", MRR_PRO: "MRR ZZP Pro",
    COMM: "Commissie-inkomsten", OMZET: "Totale omzet", K_PERS: "Kosten personeel",
    K_MKT: "Kosten marketing", K_TECH: "Kosten techniek", K_OVER: "Overige kosten",
    K_TOT: "Totale kosten", RESULT: "Maandresultaat", KAS: "Kaspositie (eind)",
    KAS_Z: "Kas zonder startkapitaal", RUNWAY: "Runway-indicatie (mnd)",
    CACR: "CAC gerealiseerd (marketing)", LTV: "LTV opdrachtgever", LTVCAC: "LTV / CAC",
    POS: "Positief resultaat? (1/0)",
}
for r, t in labels.items():
    c = wm.cell(row=r, column=1, value=t)
    if t.startswith("INPUT"):
        c.fill = input_fill
    if r in (MAAND, OMZET, K_TOT, RESULT, KAS):
        c.font = hdr_font

for i in range(1, 13):
    col = 2 + i  # C..N
    L = get_column_letter(col)
    P = get_column_letter(col - 1)
    wm.cell(row=MAAND, column=col, value=i).font = hdr_font
    wm.cell(row=MAAND, column=col).fill = hdr_fill
    wm.cell(row=PFLAG, column=col,
            value=f"=IF({L}${MAAND}<={aref('beta_einde')},1,0)")
    wm.cell(row=MKT, column=col,
            value=f"=IF({L}${MAAND}<=3,{aref('mkt1')},IF({L}${MAAND}<=6,{aref('mkt2')},{aref('mkt3')}))").number_format = EUR
    wm.cell(row=NEWFREE, column=col, value=NIEUWE_FREE[i - 1]).fill = input_fill
    wm.cell(row=NEWCO, column=col,
            value=f"=ROUND({L}{MKT}/{aref('cac')}*IF({L}${MAAND}<={aref('beta_einde')},{aref('beta_boost')},1),0)")
    wm.cell(row=SALESROW, column=col,
            value=f"=IF({L}${MAAND}>=7,{aref('sales_new')},0)").fill = input_fill
    wm.cell(row=CHURNCO, column=col,
            value=f"=IF({L}${MAAND}=1,0,ROUND({P}{ACTCO}*{aref('churn')},0))")
    wm.cell(row=ACTCO, column=col,
            value=(f"={L}{NEWCO}+{L}{SALESROW}-{L}{CHURNCO}" if i == 1
                   else f"={P}{ACTCO}+{L}{NEWCO}+{L}{SALESROW}-{L}{CHURNCO}"))
    wm.cell(row=FREE, column=col,
            value=(f"={L}{NEWFREE}" if i == 1 else f"={P}{FREE}+{L}{NEWFREE}"))
    wm.cell(row=NEWPRO, column=col, value=f"=ROUND({L}{NEWFREE}*{aref('conversie_pro')},0)")
    wm.cell(row=CHURNPRO, column=col,
            value=f"=IF({L}${MAAND}=1,0,ROUND({P}{ACTPRO}*{aref('churn_pro')},0))")
    wm.cell(row=ACTPRO, column=col,
            value=(f"={L}{NEWPRO}-{L}{CHURNPRO}" if i == 1
                   else f"={P}{ACTPRO}+{L}{NEWPRO}-{L}{CHURNPRO}"))
    wm.cell(row=MRR_CO, column=col,
            value=f"={L}{ACTCO}*{aref('arpu')}*(1-{L}{PFLAG})").number_format = EUR
    wm.cell(row=MRR_PRO, column=col,
            value=f"={L}{ACTPRO}*{aref('pro_prijs')}*(1-{L}{PFLAG})").number_format = EUR
    wm.cell(row=COMM, column=col,
            value=f"=IF({L}${MAAND}>=4,{L}{ACTCO}*{aref('commissie')}*(1-{L}{PFLAG}),0)").number_format = EUR
    wm.cell(row=OMZET, column=col,
            value=f"={L}{MRR_CO}+{L}{MRR_PRO}+{L}{COMM}").number_format = EUR
    wm.cell(row=K_PERS, column=col,
            value=f"={aref('salaris')}+IF({L}${MAAND}>=7,{aref('sales_hire')},0)").number_format = EUR
    wm.cell(row=K_MKT, column=col, value=f"={L}{MKT}").number_format = EUR
    wm.cell(row=K_TECH, column=col,
            value=f"={aref('cloud')}+{L}{ACTCO}*{aref('cloud_klant')}").number_format = EUR
    wm.cell(row=K_OVER, column=col,
            value=f"={aref('tooling')}+IF({L}${MAAND}=2,{aref('juridisch')},0)").number_format = EUR
    wm.cell(row=K_TOT, column=col,
            value=f"={L}{K_PERS}+{L}{K_MKT}+{L}{K_TECH}+{L}{K_OVER}").number_format = EUR
    wm.cell(row=RESULT, column=col, value=f"={L}{OMZET}-{L}{K_TOT}").number_format = EUR_M
    wm.cell(row=KAS, column=col,
            value=(f"={aref('startkapitaal')}+{L}{RESULT}" if i == 1
                   else f"={P}{KAS}+{L}{RESULT}")).number_format = EUR_M
    wm.cell(row=KAS_Z, column=col,
            value=(f"={L}{RESULT}" if i == 1 else f"={P}{KAS_Z}+{L}{RESULT}")).number_format = EUR_M
    wm.cell(row=RUNWAY, column=col,
            value=f'=IF({L}{RESULT}>=0,"break-even",ROUND({L}{KAS}/-{L}{RESULT},1))')
    wm.cell(row=CACR, column=col,
            value=f"=IFERROR(ROUND({L}{MKT}/({L}{NEWCO}+{L}{SALESROW}),0),0)").number_format = EUR
    wm.cell(row=LTV, column=col,
            value=f"=ROUND({aref('arpu')}*(1/{aref('churn')})*{aref('marge')},0)").number_format = EUR
    wm.cell(row=LTVCAC, column=col, value=f"=IFERROR(ROUND({L}{LTV}/{L}{CACR},1),0)").number_format = "0.0"
    wm.cell(row=POS, column=col, value=f"=IF({L}{RESULT}>=0,1,0)")

wm.cell(row=S_TITLE, column=1, value="Samenvatting").font = hdr_font
wm.cell(row=S_MRR, column=1, value="Omzet na 12 maanden")
wm.cell(row=S_MRR, column=2, value=f"=N{OMZET}").number_format = EUR
wm.cell(row=S_BE, column=1, value="Eerste maand met positief resultaat")
wm.cell(row=S_BE, column=2, value=f'=IFERROR(MATCH(1,C{POS}:N{POS},0),"geen binnen 12 mnd")')
wm.cell(row=S_KAS, column=1, value="Kaspositie eind maand 12")
wm.cell(row=S_KAS, column=2, value=f"=N{KAS}").number_format = EUR_M
wm.cell(row=S_DIP, column=1, value="Dieptepunt kas (zonder startkapitaal)")
wm.cell(row=S_DIP, column=2, value=f"=MIN(C{KAS_Z}:N{KAS_Z})").number_format = EUR_M
wm.cell(row=S_MIN, column=1, value="Minimaal startkapitaal nodig (base)")
wm.cell(row=S_MIN, column=2, value=f"=MAX(0,-MIN(C{KAS_Z}:N{KAS_Z}))").number_format = EUR

wm.column_dimensions["A"].width = 38
for i in range(3, 15):
    wm.column_dimensions[get_column_letter(i)].width = 11
wm.freeze_panes = "C5"

# ================= Blad 4: Scenario =================
ws4 = wb.create_sheet("Scenario")
ws4["A1"] = "Scenario's — zelfde logica, andere aannames"
ws4["A1"].font = title_font
for j, h in enumerate(["Parameter", "Pessimistisch", "Base", "Optimistisch"], start=1):
    ws4.cell(row=2, column=j, value=h)
style_header(ws4, 2, [1, 2, 3, 4])
scen = {
    "churn": (0.06, 0.03, 0.02, PCT, "Maandelijkse churn opdrachtgevers"),
    "cac": (1000, 600, 450, EUR, "CAC per opdrachtgever"),
    "arpu": (190, 230, 260, EUR, "Gemiddelde abonnementsprijs opdrachtgever"),
}
for i, (k, (p, b, o, fmt, desc)) in enumerate(scen.items()):
    r = 3 + i
    ws4.cell(row=r, column=1, value=desc)
    for j, v in enumerate((p, b, o)):
        c = ws4.cell(row=r, column=2 + j, value=v)
        c.number_format = fmt
        c.fill = input_fill

SC_ROWS = {"churn": 3, "cac": 4, "arpu": 5}
blocks = [("Pessimistisch", 8, "B"), ("Base", 20, "C"), ("Optimistisch", 32, "D")]

for name, b, sc in blocks:
    ws4.cell(row=b, column=1, value=f"Scenario: {name}").font = hdr_font
    style_header(ws4, b, [1])
    r_act, r_omz, r_kos, r_res, r_kas, r_pos = b + 1, b + 2, b + 3, b + 4, b + 5, b + 6
    for lbl, r in [("Actieve opdrachtgevers", r_act), ("Totale omzet", r_omz),
                   ("Totale kosten", r_kos), ("Maandresultaat", r_res),
                   ("Kas zonder startkapitaal", r_kas), ("Positief? (1/0)", r_pos)]:
        ws4.cell(row=r, column=1, value=lbl)
    ws4.cell(row=b + 8, column=1, value="Eerste break-even maand")
    ws4.cell(row=b + 8, column=2,
             value=f'=IFERROR(MATCH(1,C{r_pos}:N{r_pos},0),">12 maanden")')
    ws4.cell(row=b + 9, column=1, value="Minimaal startkapitaal nodig (EUR)")
    ws4.cell(row=b + 9, column=2,
             value=f"=MAX(0,-MIN(C{r_kas}:N{r_kas}))").number_format = EUR
    for i in range(1, 13):
        col = 2 + i
        L = get_column_letter(col)
        P = get_column_letter(col - 1)
        new = (f"ROUND(Model!{L}${MKT}/${sc}${SC_ROWS['cac']}"
               f"*IF(Model!{L}${MAAND}<={aref('beta_einde')},{aref('beta_boost')},1),0)"
               f"+Model!{L}${SALESROW}")
        ws4.cell(row=r_act, column=col,
                 value=(f"={new}" if i == 1
                        else f"={P}{r_act}+{new}-ROUND({P}{r_act}*${sc}${SC_ROWS['churn']},0)"))
        ws4.cell(row=r_omz, column=col,
                 value=(f"=({L}{r_act}*${sc}${SC_ROWS['arpu']}"
                        f"+IF(Model!{L}${MAAND}>=4,{L}{r_act}*{aref('commissie')},0))"
                        f"*(1-Model!{L}${PFLAG})+Model!{L}${MRR_PRO}")).number_format = EUR
        ws4.cell(row=r_kos, column=col,
                 value=(f"=Model!{L}${K_TOT}-Model!{L}${K_TECH}"
                        f"+{aref('cloud')}+{L}{r_act}*{aref('cloud_klant')}")).number_format = EUR
        ws4.cell(row=r_res, column=col,
                 value=f"={L}{r_omz}-{L}{r_kos}").number_format = EUR_M
        ws4.cell(row=r_kas, column=col,
                 value=(f"={L}{r_res}" if i == 1 else f"={P}{r_kas}+{L}{r_res}")).number_format = EUR_M
        ws4.cell(row=r_pos, column=col, value=f"=IF({L}{r_res}>=0,1,0)")

ws4.column_dimensions["A"].width = 34
ws4.column_dimensions["B"].width = 13
for i in range(3, 15):
    ws4.column_dimensions[get_column_letter(i)].width = 11

# ================= Blad 5: Fasen =================
wf = wb.create_sheet("Fasen")
wf["A1"] = "Fasen & KPI's — model tegen het businessplan"
wf["A1"].font = title_font
for j, h in enumerate(["Fase", "Periode", "Doelen (businessplan)", "Break-even-criterium",
                       "Model-prognose (base)"], start=1):
    wf.cell(row=3, column=j, value=h)
style_header(wf, 3, [1, 2, 3, 4, 5])
rows = [
    ("Fase 1 – NL product-market fit", "Maand 0–6",
     "Gratis beta (maand 1–3), daarna eerste betalende klanten; 10–25 opdrachtgevers",
     "Maandresultaat structureel >= 0", f"=Model!H{OMZET}"),
    ("Fase 2 – verbreding + eerste export", "Maand 6–18",
     "Abonnementen volledig, 100 opdrachtgevers, ~EUR 15k MRR-doel, UK (IR35) launch",
     "MRR >= totale kosten", f"=Model!N{OMZET}"),
    ("Fase 3 – platform & internationaal", "Maand 18–48",
     "Rules-as-code engine, API/resellers, BE/DE/ES",
     "EBITDA-positief per markt", "buiten modelvenster"),
]
for i, (f, p, d, be, mp) in enumerate(rows, start=4):
    wf.cell(row=i, column=1, value=f)
    wf.cell(row=i, column=2, value=p)
    wf.cell(row=i, column=3, value=d)
    wf.cell(row=i, column=4, value=be)
    c = wf.cell(row=i, column=5, value=mp)
    if mp.startswith("="):
        c.number_format = EUR_M

wf.cell(row=8, column=1, value="KPI-doelen (maand 12)").font = hdr_font
kpis = [
    ("CAC (marketing) <= EUR 600", f"=Model!N{CACR}"),
    ("LTV / CAC >= 3", f"=Model!N{LTVCAC}"),
    ("Maandelijkse churn <= 3%", "=Aannames!B7"),
    ("Omzet per maand (incl. ZZP + commissies)", f"=Model!N{OMZET}"),
    ("Minimaal startkapitaal (base)", "=Scenario!B29"),
    ("Minimaal startkapitaal (pessimistisch)", "=Scenario!B17"),
]
for i, (k, f) in enumerate(kpis, start=9):
    wf.cell(row=i, column=1, value=k)
    c = wf.cell(row=i, column=2, value=f)
    c.number_format = PCT if "churn" in k else ("0.0" if "LTV" in k else EUR)
wf.column_dimensions["A"].width = 44
wf.column_dimensions["B"].width = 16
wf.column_dimensions["C"].width = 60
wf.column_dimensions["D"].width = 32
wf.column_dimensions["E"].width = 22

wb.save(OUT)
print(f"Opgeslagen: {OUT}")

# ================= Python-mirror ter verificatie =================
def run(churn, cac, arpu, sales_new=2):
    v = {k: val for k, (val, _, _) in A.items()}
    act = pro = free = 0
    kas = kas_z = 0.0
    out = []
    for m in range(1, 13):
        mkt = v["mkt1"] if m <= 3 else (v["mkt2"] if m <= 6 else v["mkt3"])
        newco = rnd(mkt / cac * (v["beta_boost"] if m <= v["beta_einde"] else 1))
        sales = sales_new if m >= 7 else 0
        churnco = rnd(act * churn)
        act = act + newco + sales - churnco
        free += NIEUWE_FREE[m - 1]
        newpro = rnd(NIEUWE_FREE[m - 1] * v["conversie_pro"])
        churnpro = rnd(pro * v["churn_pro"])
        pro = pro + newpro - churnpro
        paid = 0 if m <= v["beta_einde"] else 1
        omzet = (act * arpu + pro * v["pro_prijs"] + (act * v["commissie"] if m >= 4 else 0)) * paid
        kpers = v["salaris"] + (v["sales_hire"] if m >= 7 else 0)
        ktech = v["cloud"] + act * v["cloud_klant"]
        kover = v["tooling"] + (v["juridisch"] if m == 2 else 0)
        ktot = kpers + mkt + ktech + kover
        res = omzet - ktot
        kas += res
        kas_z += res
        out.append((m, act, pro, rnd(omzet), rnd(ktot), rnd(res), rnd(kas), rnd(kas_z)))
    be = next((o[0] for o in out if o[5] >= 0), None)
    return out, be, max(0, -min(o[7] for o in out))


print("\n=== MIRROR: BASE (beta gratis t/m mnd 3, churn 3%, CAC 600, ARPU 230) ===")
print("mnd | klanten | pro | omzet | kosten | resultaat | kas")
out, be, need = run(0.03, 600, 230)
for o in out:
    print(f"{o[0]:3d} | {o[1]:7d} | {o[2]:3d} | {o[3]:6d} | {o[4]:6d} | {o[5]:9d} | {o[6]:6d}")
print(f"Eerste break-even maand: {be or '>12'} | Minimaal startkapitaal nodig: EUR {need:,.0f}")
for label, (ch, cac, arpu) in {"PESSIMISTISCH": (0.06, 1000, 190),
                               "OPTIMISTISCH": (0.02, 450, 260)}.items():
    _, b, n = run(ch, cac, arpu)
    print(f"{label}: break-even maand {b or '>12'} | min. startkapitaal EUR {n:,.0f}")
