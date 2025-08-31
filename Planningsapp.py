import streamlit as st
import copy
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime

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
        urenlijst = sorted(set(urenlijst))
        max_reeks = 1
        huidig = 1
        for idx in range(1, len(urenlijst)):
            if urenlijst[idx] == urenlijst[idx-1] + 1:
                huidig += 1
                max_reeks = max(max_reeks, huidig)
            else:
                huidig = 1
        return max_reeks

    def check_max4(student, nieuwe_blokuren):
        """Controleer of student met dit blok niet meer dan 4 opeenvolgend komt te staan."""
        # Verzamel alle al geplande uren bij deze attractie (alle posities + dit pos)
        geplande_uren = []
        if attractie in dagplanning:
            for pos in dagplanning[attractie]:
                for u, naam in pos.items():
                    if naam == student:
                        geplande_uren.append(u)
        # Ook de uren die we in deze ronde al gepland hebben in deze positie
        for u, naam in planning.items():
            if naam == student:
                geplande_uren.append(u)

        alle_uren = geplande_uren + nieuwe_blokuren
        return max_consecutive_hours(alle_uren) <= 4

    while i < len(uren):
        geplanned = False
        # probeer blokken in volgorde 3,4,2,1
        for blok in [3, 4, 2, 1]:
            if i + blok > len(uren):
                continue
            blokuren = uren[i:i+blok]
            kandidaten = [
                s for s in studenten
                if attractie in s["attracties"]
                and all(u in s["uren_beschikbaar"] for u in blokuren)
                and not any(u in student_bezet[s["naam"]] for u in blokuren)
                and gebruik_per_student_attractie[s["naam"]] + blok <= max_per_student
                and check_max4(s["naam"], blokuren)
            ]
            if kandidaten:
                # kies student met minst gebruik
                min_uren = min(gebruik_per_student_attractie[s["naam"]] for s in kandidaten)
                beste = [s for s in kandidaten if gebruik_per_student_attractie[s["naam"]] == min_uren]
                gekozen = random.choice(beste)
                for u in blokuren:
                    planning[u] = gekozen["naam"]
                    student_bezet[gekozen["naam"]].append(u)
                gebruik_per_student_attractie[gekozen["naam"]] += blok
                i += blok
                geplanned = True
                break
        if not geplanned:
            # fallback: 1 uur
            u = uren[i]
            kandidaten_1 = [
                s for s in studenten
                if attractie in s["attracties"]
                and u in s["uren_beschikbaar"]
                and u not in student_bezet[s["naam"]]
                and gebruik_per_student_attractie[s["naam"]] < max_per_student
                and check_max4(s["naam"], [u])
            ]
            if kandidaten_1:
                gekozen = random.choice(kandidaten_1)
                planning[u] = gekozen["naam"]
                student_bezet[gekozen["naam"]].append(u)
                gebruik_per_student_attractie[gekozen["naam"]] += 1
            else:
                planning[u] = "NIEMAND"
            i += 1
    return planning



# -----------------------------
# Studenten inlezen
# -----------------------------
studenten = []
for rij in range(2, 500):
    naam = ws.cell(rij, 12).value
    if not naam:
        continue
    uren_beschikbaar = [10 + (kol - 3) for kol in range(3, 12) if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]
    attracties = [ws.cell(1, kol).value for kol in range(14, 32) if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]
    raw_ag = ws[f'AG{rij}'].value
    try:
        aantal_attracties = int(raw_ag) if raw_ag is not None else len(attracties)
    except:
        aantal_attracties = len(attracties)
    studenten.append({
        "naam": naam,
        "uren_beschikbaar": uren_beschikbaar,
        "attracties": attracties,
        "aantal_attracties": aantal_attracties,
        "is_pauzevlinder": False,
        "pv_number": None
    })

# -----------------------------
# Openingsuren
# -----------------------------
open_uren = []
for kol in range(36, 45):
    uur_raw = ws.cell(1, kol).value
    vink = ws.cell(2, kol).value
    if vink in [1, True, "WAAR", "X"]:
        uur = int(str(uur_raw).replace("u","").strip()) if not isinstance(uur_raw,int) else uur_raw
        open_uren.append(uur)
if not open_uren:
    open_uren = list(range(10, 19))
open_uren = sorted(set(open_uren))

# -----------------------------
# Attracties & aantallen
# -----------------------------
aantallen = {}
attracties_te_plannen = []
for kol in range(47, 65):
    naam = ws.cell(1, kol).value
    if naam:
        try:
            aantal = int(ws.cell(2, kol).value)
        except:
            aantal = 0
        aantallen[naam] = max(0, min(2, aantal))
        if aantallen[naam] >= 1:
            attracties_te_plannen.append(naam)

def kritieke_score(attr):
    return sum(1 for s in studenten if attr in s['attracties'])
attracties_te_plannen.sort(key=kritieke_score)




@st.cache_data
def maak_planning(studenten_local):
    pauzevlinder_namen = [str(ws[f'BN{rij}'].value).strip() for rij in range(4, 11) if ws[f'BN{rij}'].value]
    required_hours = [12,13,14,15,16,17]

    selected = []
    for idx, naam in enumerate(pauzevlinder_namen, start=1):
        for s in studenten_local:
            if s['naam'] == naam:
                s['is_pauzevlinder'] = True
                s['pv_number'] = idx
                s['uren_beschikbaar'] = [u for u in s['uren_beschikbaar'] if u not in required_hours]
                selected.append(s)
                break

    student_bezet = {s['naam']: set() for s in studenten_local}
    gebruik_per_attractie_student = {attr: {s['naam']: 0 for s in studenten_local} for attr in attracties_te_plannen}
    dagplanning = {}

    # --- Eerste en tweede posities met hulpfunctie
    for attractie in attracties_te_plannen:
        dagplanning[attractie] = []
        eerste = plan_attractie_pos(attractie, studenten_local, student_bezet,
                                   gebruik_per_attractie_student[attractie], open_uren, dagplanning)
        dagplanning[attractie].append(eerste)
        if aantallen.get(attractie, 1) >= 2:
            tweede = plan_attractie_pos(attractie, studenten_local, student_bezet,
                                       gebruik_per_attractie_student[attractie], open_uren, dagplanning)
            dagplanning[attractie].append(tweede)

    # --- Extra helpers ---
    def _uren_student_bij_attr(naam, attr):
        uren = set()
        for pos in dagplanning.get(attr, []):
            for u, n in pos.items():
                if n == naam:
                    uren.add(u)
        return sorted(uren)

    def _max4_check(naam, attr, extra_uren):
        alle_uren = sorted(set(_uren_student_bij_attr(naam, attr) + list(extra_uren)))
        maxr = cur = 1
        for i in range(1, len(alle_uren)):
            cur = cur+1 if alle_uren[i]==alle_uren[i-1]+1 else 1
            maxr = max(maxr, cur)
        return maxr <= 4

    # --- Stap 1: Minstens 1 student per attractie per uur ---
    for uur in open_uren:
        for attractie, posities in dagplanning.items():
            bezet = any(pos.get(uur, "") not in ["", "NIEMAND"] for pos in posities)
            if not bezet:
                # eerst extra beschikbaar
                kandidaat = next((s for s in studenten_local
                                  if uur in s['uren_beschikbaar']
                                  and attractie in s['attracties']
                                  and uur not in student_bezet[s['naam']]
                                  and _max4_check(s['naam'], attractie, [uur])), None)
                if kandidaat:
                    posities[0][uur] = kandidaat['naam']
                    student_bezet[kandidaat['naam']].add(uur)
                    gebruik_per_attractie_student[attractie][kandidaat['naam']] += 1
                else:
                    # fallback NIEMAND
                    posities[0][uur] = "NIEMAND"

    # --- Stap 2: Vul vrije plekken met beschikbare studenten ---
    extra_per_uur = defaultdict(list)
    for uur in open_uren:
        for s in studenten_local:
            if uur in s['uren_beschikbaar'] and uur not in student_bezet[s['naam']] and not s.get('is_pauzevlinder'):
                extra_per_uur[uur].append(s['naam'])

    for uur in open_uren:
        for attractie, posities in dagplanning.items():
            for pos in posities:
                if pos.get(uur, "") in ["", "NIEMAND"]:
                    kandidaten = [s for s in studenten_local
                                  if uur in s['uren_beschikbaar']
                                  and attractie in s['attracties']
                                  and uur not in student_bezet[s['naam']]
                                  and _max4_check(s['naam'], attractie, [uur])]
                    if kandidaten:
                        min_gebruik = min(gebruik_per_attractie_student[attractie][s['naam']] for s in kandidaten)
                        beste = [s for s in kandidaten if gebruik_per_attractie_student[attractie][s['naam']] == min_gebruik]
                        gekozen = random.choice(beste)
                        pos[uur] = gekozen['naam']
                        student_bezet[gekozen['naam']].add(uur)
                        gebruik_per_attractie_student[attractie][gekozen['naam']] += 1
                        if gekozen['naam'] in extra_per_uur[uur]:
                            extra_per_uur[uur].remove(gekozen['naam'])
                    else:
                        pos[uur] = "NIEMAND"

    return dagplanning, extra_per_uur, selected


# -----------------------------
# Herhaal tot volledige planning (veilige versie)
# -----------------------------
max_attempts = 150
for attempt in range(max_attempts):
    studenten_copy = copy.deepcopy(studenten)
    dagplanning, extra_per_uur, selected = maak_planning(studenten_copy)

    # Check: elke attractie heeft minstens één student per uur
    geldig = True
    for posities in dagplanning.values():
        for pos in posities:
            for uur in open_uren:
                if pos.get(uur, '') in ['', 'NIEMAND']:
                    geldig = False
                    break
            if not geldig:
                break
        if not geldig:
            break

    if geldig:
        studenten = studenten_copy
        break




# -----------------------------
# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"
header_fill = PatternFill(start_color="BDD7EE", fill_type="solid")
attr_fill = PatternFill(start_color="E2EFDA", fill_type="solid")
pv_fill = PatternFill(start_color="FFF2CC", fill_type="solid")
extra_fill = PatternFill(start_color="FCE4D6", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Header
ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out=2
# Attracties
for attractie,posities in dagplanning.items():
    for idx,planning in enumerate(posities,start=1):
        naam_attr = attractie if len(posities)==1 else f"{attractie} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=attr_fill
        ws_out.cell(rij_out,1).border=thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            naam = planning.get(uur,"")
            if naam=="NIEMAND": naam=""
            ws_out.cell(rij_out,col_idx,naam).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out+=1

# Pauzevlinders
rij_out+=1
for pv_idx,s in enumerate(selected,start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out,col_idx,s["naam"] if uur in [12,13,14,15,16,17] else "").alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Extra studenten
rij_out+=1
max_extra=max(len(names) for names in extra_per_uur.values()) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=extra_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = extra_per_uur[uur][i] if i<len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out,col_idx,naam).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=15

# Download in Streamlit
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.download_button("Download planning", data=output, file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")


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




