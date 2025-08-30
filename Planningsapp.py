import streamlit as st
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import json


# -----------------------------
# Datum
# -----------------------------
vandaag = datetime.date.today().strftime("%d-%m-%Y")

# -----------------------------
# Excelbestand uploaden
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
if not uploaded_file:
    st.warning("Upload eerst een Excel-bestand om verder te gaan.")
    st.stop()

wb = load_workbook(BytesIO(uploaded_file.read()))
ws = wb["Blad1"]

# -----------------------------
# Hulpfunctie: plan blokken bij attractie
# -----------------------------
def plan_attractie_pos(attractie, studenten, student_bezet, gebruik_per_student_attractie, open_uren, dagplanning, max_per_student=6):
    planning = {}
    uren = sorted(open_uren)
    i = 0

    def max_consecutive_hours(urenlijst):
        """Bepaal de langste aaneengesloten reeks in een lijst uren."""
        if not urenlijst:
            return 0
        # sort + set maar 1x
        urenlijst = sorted(set(urenlijst))
        max_reeks = huidig = 1
        for idx in range(1, len(urenlijst)):
            if urenlijst[idx] == urenlijst[idx - 1] + 1:
                huidig += 1
                if huidig > max_reeks:
                    max_reeks = huidig
            else:
                huidig = 1
        return max_reeks

    def check_max4(student, nieuwe_blokuren):
        """Controleer of student met dit blok niet meer dan 4 opeenvolgend komt te staan."""
        geplande_uren = []
        # verzamel alle al geplande uren bij deze attractie
        if attractie in dagplanning:
            for pos in dagplanning[attractie]:
                geplande_uren.extend([u for u, naam in pos.items() if naam == student])
        # uren die we in deze ronde al gepland hebben
        geplande_uren.extend([u for u, naam in planning.items() if naam == student])
        alle_uren = geplande_uren + nieuwe_blokuren
        return max_consecutive_hours(alle_uren) <= 4

    # hoofdloop
    while i < len(uren):
        geplanned = False
        # probeer blokken in volgorde 3,4,2,1
        for blok in (3, 4, 2, 1):
            if i + blok > len(uren):
                continue
            blokuren = uren[i:i + blok]
            kandidaten = []
            # list comprehension met korte evaluatie
            for s in studenten:
                naam = s["naam"]
                if (
                    attractie in s["attracties"]
                    and all(u in s["uren_beschikbaar"] for u in blokuren)
                    and not any(u in student_bezet[naam] for u in blokuren)
                    and gebruik_per_student_attractie[naam] + blok <= max_per_student
                    and check_max4(naam, blokuren)
                ):
                    kandidaten.append(s)

            if kandidaten:
                # kies student met minst gebruik
                min_uren = min(gebruik_per_student_attractie[s["naam"]] for s in kandidaten)
                beste = [s for s in kandidaten if gebruik_per_student_attractie[s["naam"]] == min_uren]
                gekozen = random.choice(beste)
                naam = gekozen["naam"]
                for u in blokuren:
                    planning[u] = naam
                    student_bezet[naam].append(u)
                gebruik_per_student_attractie[naam] += blok
                i += blok
                geplanned = True
                break
        if not geplanned:
            # fallback: 1 uur
            u = uren[i]
            kandidaten_1 = []
            for s in studenten:
                naam = s["naam"]
                if (
                    attractie in s["attracties"]
                    and u in s["uren_beschikbaar"]
                    and u not in student_bezet[naam]
                    and gebruik_per_student_attractie[naam] < max_per_student
                    and check_max4(naam, [u])
                ):
                    kandidaten_1.append(s)

            if kandidaten_1:
                gekozen = random.choice(kandidaten_1)
                naam = gekozen["naam"]
                planning[u] = naam
                student_bezet[naam].append(u)
                gebruik_per_student_attractie[naam] += 1
            else:
                planning[u] = "NIEMAND"
            i += 1
    return planning

# -----------------------------
# Studenten inlezen
# -----------------------------
studenten = []
# cache headerwaarden van kolommen voor snelheid
attractie_headers = [ws.cell(1, kol).value for kol in range(14, 32)]
for rij in range(2, 500):
    naam = ws.cell(rij, 12).value
    if not naam:
        continue

    # beschikbaarheid: check kolommen 3 t/m 11
    uren_beschikbaar = []
    for kol in range(3, 12):
        val = ws.cell(rij, kol).value
        if val in (1, True, "WAAR", "X"):
            uren_beschikbaar.append(10 + (kol - 3))

    # attracties: koppel rechtstreeks aan headers
    attracties = []
    for idx, kol in enumerate(range(14, 32)):
        val = ws.cell(rij, kol).value
        if val in (1, True, "WAAR", "X"):
            attracties.append(attractie_headers[idx])

    raw_ag = ws[f'AG{rij}'].value
    try:
        aantal_attracties = int(raw_ag) if raw_ag is not None else len(attracties)
    except Exception:
        aantal_attracties = len(attracties)

    studenten.append({
        "naam": naam,
        "uren_beschikbaar": uren_beschikbaar,
        "attracties": attracties,
        "aantal_attracties": aantal_attracties,
        "is_pauzevlinder": False,
        "pv_number": None,
    })

# -----------------------------
# Openingsuren
# -----------------------------
open_uren = []
for kol in range(36, 45):
    uur_raw = ws.cell(1, kol).value
    vink = ws.cell(2, kol).value
    if vink in (1, True, "WAAR", "X"):
        if isinstance(uur_raw, int):
            uur = uur_raw
        else:
            uur = int(str(uur_raw).replace("u", "").strip())
        open_uren.append(uur)

if not open_uren:
    open_uren = list(range(10, 19))
else:
    open_uren = sorted(set(open_uren))

# -----------------------------
# Attracties & aantallen
# -----------------------------
aantallen = {}
attracties_te_plannen = []
for kol in range(47, 65):
    naam = ws.cell(1, kol).value
    if not naam:
        continue
    try:
        aantal = int(ws.cell(2, kol).value)
    except Exception:
        aantal = 0
    aantal = max(0, min(2, aantal))
    aantallen[naam] = aantal
    if aantal >= 1:
        attracties_te_plannen.append(naam)

# sorteer attracties op "kritiek" (weinig studenten beschikbaar eerst)
def kritieke_score(attr):
    return sum(1 for s in studenten if attr in s['attracties'])

attracties_te_plannen.sort(key=kritieke_score)
# -----------------------------
# Maak planning inclusief schuiven, swaps en extra regels
# -----------------------------
def maak_planning(studenten_local):
    # Pauzevlinders inlezen uit Excel (BN4:BN10)
    pauzevlinder_namen = []
    for rij in range(4, 11):
        naam = ws[f'BN{rij}'].value
        if naam:
            pauzevlinder_namen.append(str(naam).strip())

    required_hours = [12, 13, 14, 15, 16, 17]

    # Build quick lookup: naam -> student (wijzigt in-place zoals origineel)
    student_by_name = {s["naam"]: s for s in studenten_local}

    selected = []
    for idx, naam in enumerate(pauzevlinder_namen, start=1):
        s = student_by_name.get(naam)
        if s:
            s["is_pauzevlinder"] = True
            s["pv_number"] = idx
            # behoud dezelfde semantiek: verwijder required_hours uit uren_beschikbaar
            s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_hours]
            selected.append(s)

    student_bezet = {s["naam"]: [] for s in studenten_local}
    dagplanning = {}
    gebruik_per_attractie_student = {attr: {s["naam"]: 0 for s in studenten_local} for attr in attracties_te_plannen}

    # --- Eerste posities ---
    for attractie in attracties_te_plannen:
        dagplanning[attractie] = [plan_attractie_pos(
            attractie, studenten_local, student_bezet,
            gebruik_per_attractie_student[attractie], open_uren, dagplanning
        )]

    # --- Tweede posities ---
    for attractie in attracties_te_plannen:
        if aantallen.get(attractie, 1) >= 2:
            dagplanning[attractie].append(plan_attractie_pos(
                attractie, studenten_local, student_bezet,
                gebruik_per_attractie_student[attractie], open_uren, dagplanning
            ))

    # Kleine helpers (alleen voor checks in extra/swap-stappen)
    def _uren_student_bij_attr(naam, attr):
        # snel: comprehension + set om duplicaten te vermijden, daarna sorteren
        if attr not in dagplanning:
            return []
        uren = {u for pos in dagplanning[attr] for u, n in pos.items() if n == naam}
        return sorted(uren)

    def _max_consecutive(urenlijst):
        if not urenlijst:
            return 0
        urenlijst = sorted(set(urenlijst))
        maxr = cur = 1
        for i in range(1, len(urenlijst)):
            if urenlijst[i] == urenlijst[i-1] + 1:
                cur += 1
                if cur > maxr:
                    maxr = cur
            else:
                cur = 1
        return maxr

    def _ok_max4(naam, attr, extra_uren):
        alle = _uren_student_bij_attr(naam, attr) + list(extra_uren)
        return _max_consecutive(alle) <= 4

    # --- Iteratief schuiven en swaps (met 4u-bewaking in extra/regel-2) ---
    while True:
        wijziging = False

        # Bezet per uur (uren_bezet[u] = set van namen)
        uren_bezet = defaultdict(set)
        for posities in dagplanning.values():
            for pos in posities:
                for u, naam in pos.items():
                    if naam not in ("", "NIEMAND"):
                        uren_bezet[u].add(naam)
        for pv in selected:
            for u in required_hours:
                uren_bezet[u].add(pv["naam"])

        # Extra studenten: optimaler vullen door 1x over studenten te lopen
        extra_per_uur = defaultdict(list)
        for s in studenten_local:
            if s.get("is_pauzevlinder"):
                continue
            naam = s["naam"]
            for uur in s["uren_beschikbaar"]:
                # alleen toevoegen als nog niet bezet op dat uur
                if naam not in uren_bezet.get(uur, set()):
                    extra_per_uur[uur].append(naam)

        # --- Regel 1: geen extra zolang er lege plekken zijn (nu met 4u-check) ---
        for uur in open_uren:
            # verzamel alle lege plekken (alle posities waar niets/NIEMAND staat)
            lege_posities = []
            for attractie, posities in dagplanning.items():
                for pos in posities:
                    if pos.get(uur, "") in ("", "NIEMAND"):
                        lege_posities.append((attractie, pos))

            # zolang er extra's én lege plekken zijn, proberen vullen via directe plaatsing of swap
            tried_guard = 0
            while extra_per_uur[uur] and lege_posities:
                tried_guard += 1
                if tried_guard > 2000:
                    break

                extra_student = extra_per_uur[uur].pop(0)
                s_obj = student_by_name[extra_student]  # O(1) lookup
                geplaatst = False

                # 1) Directe plaatsing als mogelijk én geen >4-streak
                for attractie, pos in list(lege_posities):
                    if (attractie in s_obj["attracties"]
                        and uur in s_obj["uren_beschikbaar"]
                        and _ok_max4(extra_student, attractie, [uur])):
                        pos[uur] = extra_student
                        student_bezet[extra_student].append(uur)
                        gebruik_per_attractie_student[attractie][extra_student] += 1
                        uren_bezet[uur].add(extra_student)
                        # verwijder uit lege_posities
                        try:
                            lege_posities.remove((attractie, pos))
                        except ValueError:
                            pass
                        wijziging = True
                        geplaatst = True
                        break

                if geplaatst:
                    continue

                # 2) Swap: extra_student -> bezette plek A ; 'huidige' -> lege plek B
                for attractie, posities in dagplanning.items():
                    if geplaatst:
                        break
                    for pos in posities:
                        huidige = pos.get(uur, "")
                        if huidige in ("", "NIEMAND"):
                            continue
                        if (attractie in s_obj["attracties"]
                            and uur in s_obj["uren_beschikbaar"]
                            and _ok_max4(extra_student, attractie, [uur])):
                            # zoek lege plek B waar huidige heen kan, ook met 4u-check
                            for lege_attr, lege_pos in list(lege_posities):
                                h_obj = student_by_name[huidige]
                                if (lege_attr in h_obj["attracties"]
                                    and _ok_max4(huidige, lege_attr, [uur])):
                                    # voer swap uit
                                    pos[uur] = extra_student
                                    student_bezet[extra_student].append(uur)
                                    gebruik_per_attractie_student[attractie][extra_student] += 1
                                    uren_bezet[uur].add(extra_student)

                                    # zet de vorige persoon in lege_pos op dit uur
                                    lege_pos[uur] = huidige
                                    student_bezet[huidige].append(uur)
                                    gebruik_per_attractie_student[lege_attr][huidige] += 1
                                    uren_bezet[uur].add(huidige)

                                    try:
                                        lege_posities.remove((lege_attr, lege_pos))
                                    except ValueError:
                                        pass
                                    wijziging = True
                                    geplaatst = True
                                    break
                                # einde if
                            if geplaatst:
                                break
                        # einde if
                    # einde for posities
                # einde swap loop

                # 3) Niet gelukt -> terug naar extra lijst en stop voor dit uur (vermijd eindeloze rondjes)
                if not geplaatst:
                    extra_per_uur[uur].append(extra_student)
                    break  # laat een volgende iteratie het opnieuw proberen

        # --- Regel 2: minstens 1 persoon per attractie per uur (met 4u-check) ---
        for uur in open_uren:
            for attractie, posities in dagplanning.items():
                bezet = [p.get(uur, "") for p in posities if p.get(uur, "") not in ("", "NIEMAND")]
                if not bezet:  # nog helemaal leeg in dit uur voor deze attractie
                    kandidaat = None

                    # 2a) Neem iemand uit extra als die mag/kan en 4u niet overschrijdt
                    if extra_per_uur[uur]:
                        for i, naam in enumerate(list(extra_per_uur[uur])):
                            s_obj = student_by_name[naam]
                            if (attractie in s_obj["attracties"]
                                and uur in s_obj["uren_beschikbaar"]
                                and _ok_max4(naam, attractie, [uur])):
                                kandidaat = naam
                                del extra_per_uur[uur][i]
                                break

                    # 2b) Anders: vrije student (niet bezet op dit uur) die 4u niet overschrijdt
                    if not kandidaat:
                        for s in studenten_local:
                            if (uur in s["uren_beschikbaar"]
                                and attractie in s["attracties"]
                                and s["naam"] not in uren_bezet[uur]
                                and _ok_max4(s["naam"], attractie, [uur])):
                                kandidaat = s["naam"]
                                break

                    # 2c) Als nog geen kandidaat: probeer een milde swap
                    if not kandidaat:
                        for bron_attr, bron_posities in dagplanning.items():
                            if kandidaat:
                                break
                            bezet_bron = [p for p in bron_posities if p.get(uur, "") not in ("", "NIEMAND")]
                            if len(bezet_bron) >= 2 or (bron_attr != attractie and len(bezet_bron) >= 1):
                                for p in bezet_bron:
                                    naam = p[uur]
                                    s_obj = student_by_name[naam]
                                    if (attractie in s_obj["attracties"]
                                        and _ok_max4(naam, attractie, [uur])):
                                        p[uur] = ""  # laat bron even leeg; volgende iteraties vullen dit op
                                        kandidaat = naam
                                        wijziging = True
                                        break

                    if kandidaat:
                        posities[0][uur] = kandidaat
                        student_bezet[kandidaat].append(uur)
                        gebruik_per_attractie_student[attractie][kandidaat] += 1
                        uren_bezet[uur].add(kandidaat)
                        wijziging = True

        if not wijziging:
            break

    return dagplanning, extra_per_uur, selected
# -----------------------------
# Herhaal tot volledige planning
# -----------------------------
max_attempts = 2
dagplanning = {}
extra_per_uur = {}
selected = []

for attempt in range(max_attempts):
    # deepcopy is duur → gebruik json.loads/dumps (sneller voor dict/list-structuren)
    studenten_copy = json.loads(json.dumps(studenten))  
    dagplanning, extra_per_uur, selected = maak_planning(studenten_copy)

    # check sneller met any() ipv nested loops
    if not any(
        pos.get(u, "") == "NIEMAND" and extra_per_uur.get(u)
        for p in dagplanning.values()
        for pos in p
        for u in pos
    ):
        studenten = studenten_copy
        break


# -----------------------------
# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"

# Veel gebruikte objecten 1x aanmaken
header_font = Font(bold=True)
bold_font = Font(bold=True)
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

# Kleuren vooraf
header_fill = PatternFill(start_color="BDD7EE", fill_type="solid")
attr_fill   = PatternFill(start_color="E2EFDA", fill_type="solid")
pv_fill     = PatternFill(start_color="FFF2CC", fill_type="solid")
extra_fill  = PatternFill(start_color="FCE4D6", fill_type="solid")

# Header
ws_out.cell(1,1,vandaag).font = header_font
for col_idx, uur in enumerate(open_uren, start=2):  # open_uren is al gesorteerd
    c = ws_out.cell(1, col_idx, f"{uur}:00")
    c.font = header_font
    c.fill = header_fill
    c.alignment = center_align
    c.border = thin_border

rij_out = 2

# Attracties
for attractie, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attractie if len(posities) == 1 else f"{attractie} {idx}"
        c = ws_out.cell(rij_out, 1, naam_attr)
        c.font = bold_font
        c.fill = attr_fill
        c.border = thin_border

        for col_idx, uur in enumerate(open_uren, start=2):
            naam = planning.get(uur, "")
            if naam == "NIEMAND":
                naam = ""
            c = ws_out.cell(rij_out, col_idx, naam)
            c.alignment = center_align
            c.border = thin_border
        rij_out += 1

# Pauzevlinders
rij_out += 1
required_hours = [12,13,14,15,16,17]
for pv_idx, s in enumerate(selected, start=1):
    c = ws_out.cell(rij_out, 1, f"Pauzevlinder {pv_idx}")
    c.font = bold_font
    c.fill = pv_fill
    c.border = thin_border

    for col_idx, uur in enumerate(open_uren, start=2):
        naam = s["naam"] if uur in required_hours else ""
        c = ws_out.cell(rij_out, col_idx, naam)
        c.alignment = center_align
        c.border = thin_border
    rij_out += 1

# Extra studenten
rij_out += 1
max_extra = max((len(v) for v in extra_per_uur.values()), default=0)
for i in range(max_extra):
    c = ws_out.cell(rij_out, 1, "Extra")
    c.font = bold_font
    c.fill = extra_fill
    c.border = thin_border

    for col_idx, uur in enumerate(open_uren, start=2):
        naam = extra_per_uur.get(uur, [])
        naam = naam[i] if i < len(naam) else ""
        c = ws_out.cell(rij_out, col_idx, naam)
        c.alignment = center_align
        c.border = thin_border
    rij_out += 1

# Kolombreedte
for col in range(1, len(open_uren) + 2):
    ws_out.column_dimensions[get_column_letter(col)].width = 15

# Download in Streamlit
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.download_button(
    "Download planning",
    data=output,
    file_name=f"Planning_{datetime.datetime.now():%Y%m%d_%H%M%S}.xlsx"
)




#ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo






#DEEL 2
#oooooooooooooooooooo
#oooooooooooooooooooo

# -----------------------------
# DEEL 2: Pauzevlinder overzicht
# -----------------------------
ws_pauze = wb_out.create_sheet(title="Pauzevlinders")

light_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# -----------------------------
# Rij 1: Uren
# -----------------------------
uren_rij1 = []

# Halve uren 12:00 tot 14:30
u = 12
m = 0
while u < 15 or (u == 14 and m <= 30):
    uren_rij1.append(f"{u:02d}:{m:02d}")
    m += 30
    if m >= 60:
        u += 1
        m = 0

# Lege kolom tussen 14:30 en 15:30
uren_rij1.append("")  # lege kolom

# Start vanaf 15:30 met kwartier tot 17:15
u = 15
m = 30
while u < 17 or (u == 17 and m <= 15):
    uren_rij1.append(f"{u:02d}:{m:02d}")
    m += 15
    if m >= 60:
        u += 1
        m = 0

# Schrijf uren in rij 1, start in kolom B
for col_idx, uur in enumerate(uren_rij1, start=2):
    c = ws_pauze.cell(1, col_idx, uur)
    c.fill = light_fill
    c.alignment = center_align
    c.border = thin_border

# Zet cel A1 ook in licht kleurtje
a1 = ws_pauze.cell(1, 1, "")
a1.fill = light_fill
a1.border = thin_border

# -----------------------------
# Pauzevlinders en namen
# -----------------------------
rij_out = 2
for pv_idx, pv in enumerate(selected, start=1):
    # Titel: Pauzevlinder X
    title_cell = ws_pauze.cell(rij_out, 1, f"Pauzevlinder {pv_idx}")
    title_cell.font = Font(bold=True)
    title_cell.fill = light_fill
    title_cell.alignment = center_align
    title_cell.border = thin_border

    # Naam eronder (zelfde stijl en kleur)
    naam_cel = ws_pauze.cell(rij_out + 1, 1, pv["naam"])
    naam_cel.fill = light_fill
    naam_cel.alignment = center_align
    naam_cel.border = thin_border

    rij_out += 3  # lege rij ertussen

# -----------------------------
# Kolombreedte voor overzicht
# -----------------------------
for col in range(1, len(uren_rij1) + 2):
    ws_pauze.column_dimensions[get_column_letter(col)].width = 10

# Opslaan met dezelfde unieke naam

# Maak in-memory bestand
output = BytesIO()
wb_out.save(output)
output.seek(0)  # Zorg dat lezen vanaf begin kan



#oooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo


#DEEL 3
#oooooooooooooooooooo
#oooooooooooooooooooo

import random
from collections import defaultdict
from openpyxl.styles import Alignment, Border, Side, PatternFill
import datetime

# -----------------------------
# DEEL 3: Extra naam voor pauzevlinders die langer dan 6 uur werken
# -----------------------------

# Sheet referenties
ws_planning = wb_out["Planning"]
ws_pauze = wb_out["Pauzevlinders"]

# Pauzekolommen (B–G in Pauzevlinders sheet)
pauze_cols = list(range(2, 8))  # B(2) t/m G(7)

# Bouw lijst met pauzevlinder-rijen
pv_rows = []
for pv in selected:
    row_found = None
    for r in range(2, ws_pauze.max_row + 1):
        if str(ws_pauze.cell(r, 1).value).strip() == str(pv["naam"]).strip():
            row_found = r
            break
    if row_found is not None:
        pv_rows.append((pv, row_found))

# Bereken totaal uren per student in Werkblad "Planning"
student_totalen = defaultdict(int)
for row in ws_planning.iter_rows(min_row=2, values_only=True):
    for naam in row[1:]:
        if naam and str(naam).strip() != "":
            student_totalen[naam] += 1

# Loop door pauzevlinders in Werkblad "Pauzevlinders"
for row in range(2, ws_pauze.max_row+1, 3):  # elke pauzevlinder heeft 2 rijen + 1 lege rij
    naam_cel = ws_pauze.cell(row + 1, 1).value
    if not naam_cel:
        continue
    totaal_uren = student_totalen.get(naam_cel, 0)
    if totaal_uren > 6:
        # Kies random kolom tussen B en G (2 t/m 7)
        random_col = random.randint(2, 7)
        ws_pauze.cell(row + 1, random_col, naam_cel)
        # Opmaak toepassen
        ws_pauze.cell(row + 1, random_col).alignment = Alignment(horizontal="center", vertical="center")
        ws_pauze.cell(row + 1, random_col).border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin")
        )

# ---- Lege naamcellen inkleuren ----
naam_leeg_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
center_align = Alignment(horizontal="center", vertical="center")

for pv, pv_row in pv_rows:
    for col in pauze_cols:
        if ws_pauze.cell(pv_row, col).value in [None, ""]:
            ws_pauze.cell(pv_row, col).fill = naam_leeg_fill

# -----------------------------
# Opslaan in uniek bestand
# -----------------------------
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
planning_bestand = f"Planning_{timestamp}.xlsx"

# Maak in-memory bestand
output = BytesIO()
wb_out.save(output)
output.seek(0)  # Zorg dat lezen vanaf begin kan


ws_feedback = wb_out.create_sheet("Feedback")
def log_feedback(msg):
    """Voeg een regel toe in het feedback-werkblad."""
    next_row = ws_feedback.max_row + 1
    ws_feedback.cell(next_row, 1, msg)


log_feedback(f"✅ Alle pauzevlinders die >6u werken kregen een extra pauzeplek (B–G) in {planning_bestand}")




#ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo

# DEEL 4: Lange werkers (>6 uur) pauze inplannen – gegarandeerd
# -----------------------------

from openpyxl.styles import Alignment, Border, Side, PatternFill
import random  # <— toegevoegd voor willekeurige verdeling

thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
center_align = Alignment(horizontal="center", vertical="center")
# Zachtblauw, anders dan je titelkleuren; alleen voor naamcellen
naam_leeg_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")

# Alleen kolommen B..G
pauze_cols = list(range(2, 8))  # B(2),C(3),D(4),E(5),F(6),G(7)


def is_student_extra(naam):
    """Check of student in Planning bij 'extra' staat (kolom kan 'Extra' zijn of specifieke marker)."""
    for row in range(2, ws_planning.max_row + 1):
        if ws_planning.cell(row, 1).value:  # rij met attractienaam
            for col in range(2, ws_planning.max_column + 1):
                if str(ws_planning.cell(row, col).value).strip().lower() == str(naam).strip().lower():
                    header = str(ws_planning.cell(1, col).value).strip().lower()
                    if "extra" in header:  # of een andere logica afhankelijk van hoe 'extra' wordt aangeduid
                        return True
    return False


# ---- Helpers ----
def parse_header_uur(header):
    """Map headertekst (bv. '14u', '14:00', '14:30') naar het hele uur (14)."""
    if not header:
        return None
    s = str(header).strip()
    try:
        if "u" in s:
            # '14u' of '14u30' -> 14
            return int(s.split("u")[0])
        if ":" in s:
            # '14:00' of '14:30' -> 14 (halfuur koppelen aan het hele uur)
            uur, _min = s.split(":")
            return int(uur)
        return int(s)  # fallback
    except:
        return None

def normalize_attr(name):
    """Normaliseer attractienaam zodat 'X 2' telt als 'X'; trim & lower-case voor vergelijking."""
    if not name:
        return ""
    s = str(name).strip()
    parts = s.rsplit(" ", 1)
    if len(parts) == 2 and parts[1].isdigit():
        s = parts[0]
    return s.strip().lower()

# Build: pauzevlinder-rijen en capability-sets
pv_rows = []  # lijst: (pv_dict, naam_rij_index)
pv_cap_set = {}  # pv-naam -> set genormaliseerde attracties
for pv in selected:
    # zoek de rij waar de pv-naam in kolom A staat
    row_found = None
    for r in range(2, ws_pauze.max_row + 1):
        if str(ws_pauze.cell(r, 1).value).strip() == str(pv["naam"]).strip():
            row_found = r
            break
    if row_found is not None:
        pv_rows.append((pv, row_found))
        pv_cap_set[pv["naam"]] = {normalize_attr(a) for a in pv.get("attracties", [])}

# Lange werkers: namen-set voor snelle checks
lange_werkers = [s for s in studenten
                 if student_totalen.get(s["naam"], 0) > 6
                 and s["naam"] not in [pv["naam"] for pv in selected]]
lange_werkers_names = {s["naam"] for s in lange_werkers}

def get_student_work_hours(naam):
    """Welke uren werkt deze student echt (zoals te zien in werkblad 'Planning')?"""
    uren = set()
    for col in range(2, ws_planning.max_column + 1):
        header = ws_planning.cell(1, col).value
        uur = parse_header_uur(header)
        if uur is None:
            continue
        # check of student in deze kolom ergens staat
        for row in range(2, ws_planning.max_row + 1):
            if ws_planning.cell(row, col).value == naam:
                uren.add(uur)
                break
    return sorted(uren)

def vind_attractie_op_uur(naam, uur):
    """Geef attractienaam (exact zoals in Planning-kolom A) waar student staat op dit uur; None als niet gevonden."""
    for col in range(2, ws_planning.max_column + 1):
        header = ws_planning.cell(1, col).value
        col_uur = parse_header_uur(header)
        if col_uur != uur:
            continue
        for row in range(2, ws_planning.max_row + 1):
            if ws_planning.cell(row, col).value == naam:
                return ws_planning.cell(row, 1).value
    return None

def pv_kan_attr(pv, attr):
    """Check of pv attr kan (met normalisatie, zodat 'X 2' telt als 'X')."""
    base = normalize_attr(attr)
    return base in pv_cap_set.get(pv["naam"], set())

# Willekeurige volgorde van pauzeplekken (pv-rij x kolom) om lege cellen random te spreiden
slot_order = [(pv, pv_row, col) for (pv, pv_row) in pv_rows for col in pauze_cols]
random.shuffle(slot_order)  # <— kern om lege plekken later willekeurig verspreid te krijgen

def plaats_student(student, harde_mode=False):
    """
    Plaats student bij een geschikte pauzevlinder in B..G op een uur waar student effectief werkt.
    - Overschrijven alleen in harde_mode én alleen als de huidige inhoud géén lange werker is.
    - Volgorde van slots is willekeurig (slot_order) zodat lege plekken random verdeeld blijven.
    """
    naam = student["naam"]
    werk_uren = get_student_work_hours(naam)  # echte uren waarop student in 'Planning' staat

    # uren ook willekeurig proberen, voor extra spreiding
    for uur in random.sample(werk_uren, len(werk_uren)):
        attr = vind_attractie_op_uur(naam, uur)
        if not attr:
            continue

        # Probeer alle slots in random volgorde; filter op uur en pv-capability
        for (pv, pv_row, col) in slot_order:
            col_header = ws_pauze.cell(1, col).value
            col_uur = parse_header_uur(col_header)
            if col_uur != uur:
                continue
            if not pv_kan_attr(pv, attr) and not is_student_extra(naam):
                continue

            cel = ws_pauze.cell(pv_row, col)            # naamcel (rij met pv-naam)
            boven_cel = ws_pauze.cell(pv_row - 1, col)  # attractiecel (rij erboven)
            current_val = cel.value

            if current_val in [None, ""]:
                # Vrij: direct plaatsen
                boven_cel.value = attr
                boven_cel.alignment = center_align
                boven_cel.border = thin_border

                cel.value = naam
                cel.alignment = center_align
                cel.border = thin_border
                return True
            else:
                # Bezet: in harde modus enkel overschrijven als de huidige naam GEEN lange werker is
                if harde_mode:
                    occupant = str(current_val).strip()
                    if occupant not in lange_werkers_names:
                        boven_cel.value = attr
                        boven_cel.alignment = center_align
                        boven_cel.border = thin_border

                        cel.value = naam
                        cel.alignment = center_align
                        cel.border = thin_border
                        return True
        # volgende werk-uur proberen
    return False

# ---- Fase 1: zachte toewijzing (niet overschrijven) ----
niet_geplaatst = []
# Studenten in willekeurige volgorde proberen om vulling te spreiden
for s in random.sample(lange_werkers, len(lange_werkers)):
    if not plaats_student(s, harde_mode=False):
        niet_geplaatst.append(s)

# ---- Fase 2: iteratief, met gecontroleerd overschrijven van niet-lange-werkers ----
# we herhalen meerdere passes om iedereen >6u te kunnen plaatsen
max_passes = 12
for _ in range(max_passes):
    if not niet_geplaatst:
        break
    rest = []
    # Ook hier willekeurige volgorde voor extra spreiding
    for s in random.sample(niet_geplaatst, len(niet_geplaatst)):
        if not plaats_student(s, harde_mode=True):
            rest.append(s)
    # Als niets veranderde in een hele pass, stoppen we
    if len(rest) == len(niet_geplaatst):
        break
    niet_geplaatst = rest

# ---- Lege naamcellen inkleuren (alleen de NAAM-rij; bovenliggende attractie-rij NIET kleuren) ----
for pv, pv_row in pv_rows:
    for col in pauze_cols:
        if ws_pauze.cell(pv_row, col).value in [None, ""]:
            ws_pauze.cell(pv_row, col).fill = naam_leeg_fill


# Maak in-memory bestand
output = BytesIO()
wb_out.save(output)
output.seek(0)  # Zorg dat lezen vanaf begin kan


if niet_geplaatst:
    log_feedback(f"⚠️ Nog niet geplaatst (controleer of pv's deze attracties kunnen): {[s['naam'] for s in niet_geplaatst]}")
else:
    log_feedback("✅ Alle studenten die >6u werken kregen een pauzeplek (B–G gevuld waar mogelijk)")




#ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo


# -----------------------------
# DEEL 5: Kwartierpauzes namiddag (15:30-17:30)
# -----------------------------
from openpyxl.styles import Alignment, Border, Side, PatternFill
import random

# Styles
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
center_align = Alignment(horizontal="center", vertical="center")
pauze_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")  # lichtgroen

# Kolommen I..P
pauze_cols = list(range(9, 17))  # I(9) t/m P(16)

# Helper: map kolomheader naar uur
def parse_header_uur(header):
    if not header:
        return None
    s = str(header).strip()
    try:
        if "u" in s:
            return int(s.split("u")[0])
        if ":" in s:
            uur, _ = s.split(":")
            return int(uur)
        return int(s)
    except:
        return None

# Pauzevlinder-rijen en hun capabilities
pv_rows = []
pv_cap_set = {}
for pv in selected:
    row_found = None
    for r in range(2, ws_pauze.max_row + 1):
        if str(ws_pauze.cell(r, 1).value).strip() == str(pv["naam"]).strip():
            row_found = r
            break
    if row_found:
        pv_rows.append((pv, row_found))
        pv_cap_set[pv["naam"]] = {normalize_attr(a) for a in pv.get("attracties", [])}

# Studenten die minstens 1 uur werken
min_werkers = [s for s in studenten if student_totalen.get(s["naam"],0) >= 1]
min_werkers_names = {s["naam"] for s in min_werkers}

# Werkuren per student
def get_student_work_hours(naam):
    uren = set()
    for col in range(2, ws_planning.max_column + 1):
        header = ws_planning.cell(1, col).value
        uur = parse_header_uur(header)
        if uur is None:
            continue
        for row in range(2, ws_planning.max_row + 1):
            if ws_planning.cell(row, col).value == naam:
                uren.add(uur)
                break
    return sorted(uren)

def vind_attractie_op_uur(naam, uur):
    for col in range(2, ws_planning.max_column + 1):
        header = ws_planning.cell(1, col).value
        col_uur = parse_header_uur(header)
        if col_uur != uur:
            continue
        for row in range(2, ws_planning.max_row + 1):
            if ws_planning.cell(row, col).value == naam:
                return ws_planning.cell(row,1).value
    return None

def pv_kan_attr(pv, attr):
    base = normalize_attr(attr)
    return base in pv_cap_set.get(pv["naam"], set())

def plaats_student(student, harde_mode=False):
    naam = student["naam"]
    werk_uren = get_student_work_hours(naam)
    if not werk_uren:
        return False
    random.shuffle(werk_uren)
    for uur in werk_uren:
        attr = vind_attractie_op_uur(naam, uur)
        for (pv, pv_row) in random.sample(pv_rows, len(pv_rows)):
            # 🚩 uitzondering: "extra" mag altijd
            if attr and attr.strip().lower() != "extra" and not pv_kan_attr(pv, attr):
                continue
            for col in random.sample(pauze_cols, len(pauze_cols)):
                col_header = ws_pauze.cell(1, col).value
                col_uur = parse_header_uur(col_header)
                if col_uur != uur:
                    continue
                cel = ws_pauze.cell(pv_row, col)
                current_val = cel.value
                boven_cel = ws_pauze.cell(pv_row-1, col)
                if current_val in [None, ""]:
                    if attr:
                        boven_cel.value = attr
                        boven_cel.alignment = center_align
                        boven_cel.border = thin_border
                    cel.value = naam
                    cel.alignment = center_align
                    cel.border = thin_border
                    cel.fill = pauze_fill
                    return True
                elif harde_mode:
                    occupant = str(current_val).strip()
                    if occupant not in min_werkers_names:
                        if attr:
                            boven_cel.value = attr
                            boven_cel.alignment = center_align
                            boven_cel.border = thin_border
                        cel.value = naam
                        cel.alignment = center_align
                        cel.border = thin_border
                        cel.fill = pauze_fill
                        return True
    return False


def plaats_student_links_eerst(student):
    naam = student["naam"]
    werk_uren = get_student_work_hours(naam)
    if not werk_uren:
        return False
    for uur in werk_uren:
        attr = vind_attractie_op_uur(naam, uur)
        for (pv, pv_row) in pv_rows:
            # 🚩 uitzondering: "extra" mag altijd
            if attr and attr.strip().lower() != "extra" and not pv_kan_attr(pv, attr):
                continue
            for col in pauze_cols:  # van links naar rechts
                col_header = ws_pauze.cell(1, col).value
                col_uur = parse_header_uur(col_header)
                if col_uur != uur:
                    continue
                cel = ws_pauze.cell(pv_row, col)
                if cel.value in [None, ""]:
                    if attr:
                        boven_cel = ws_pauze.cell(pv_row-1, col)
                        boven_cel.value = attr
                        boven_cel.alignment = center_align
                        boven_cel.border = thin_border
                    cel.value = naam
                    cel.alignment = center_align
                    cel.border = thin_border
                    cel.fill = pauze_fill
                    return True
    return False


# -----------------------------
# Fase 1: pauzevlinders eerst in hun eigen rijen
for pv, pv_row in pv_rows:
    col = random.choice(pauze_cols)
    cel = ws_pauze.cell(pv_row, col)
    if cel.value in [None, ""]:
        cel.value = pv["naam"]
        cel.alignment = center_align
        cel.border = thin_border
        cel.fill = pauze_fill

# -----------------------------
# Fase 2: studenten die max 6 uur werken, links → rechts proberen
kortere_werkers = [s for s in min_werkers if student_totalen.get(s["naam"],0) <= 6 and s["naam"] not in [pv["naam"] for pv in selected]]
niet_geplaatst = []
for s in kortere_werkers:  # vaste volgorde
    if not plaats_student_links_eerst(s):
        niet_geplaatst.append(s)

# -----------------------------
# Fase 3: overige studenten (>6u of nog niet geplaatst) → random
overige = [s for s in min_werkers if s["naam"] not in [pv["naam"] for pv in selected] and s["naam"] not in [w["naam"] for w in kortere_werkers]]
niet_geplaatst_extra = []
for s in random.sample(overige, len(overige)):
    if not plaats_student(s, harde_mode=True):
        niet_geplaatst_extra.append(s)

# -----------------------------
# Lege naamcellen inkleuren
for pv, pv_row in pv_rows:
    for col in pauze_cols:
        if ws_pauze.cell(pv_row, col).value in [None, ""]:
            ws_pauze.cell(pv_row, col).fill = pauze_fill

# -----------------------------
# Opslaan in hetzelfde unieke bestand als DEEL 3
# -----------------------------

# Maak in-memory bestand
output = BytesIO()
wb_out.save(output)
output.seek(0)  # Zorg dat lezen vanaf begin kan


st.title("Planning Generator")

st.write("Upload je Excel-bestand om een planning te maken.")



st.download_button(
    label="Download Planning Excel",
    data=output,
    file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)





