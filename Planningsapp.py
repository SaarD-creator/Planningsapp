#looking perfect

#red spots, geen (ithink) extra's, geen lange blokken, wel nog extra bij max van 4 uur

#perfectttt, enkel nog probleempje als persoon met meeste attracties iets niet kan

# we zien de red spots al, gwn nog teveel bij extra


import streamlit as st
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
    # st.warning("Upload eerst een Excel-bestand om verder te gaan.")
    # st.stop()

wb = load_workbook(BytesIO(uploaded_file.read()))
ws = wb["Blad1"]

# -----------------------------

# Hulpfuncties
# -----------------------------
def max_consecutive_hours(urenlijst):
    if not urenlijst:
        return 0
    urenlijst = sorted(set(urenlijst))
    maxr = huidig = 1
    for i in range(1, len(urenlijst)):
        huidig = huidig + 1 if urenlijst[i] == urenlijst[i-1] + 1 else 1
        maxr = max(maxr, huidig)
    return maxr

def partition_run_lengths(L):
    """Flexibele blokken: prioritair 3 uur, dan 2,4,1 om shift te vullen."""
    blocks = [3,2,4,1]
    dp = [(10**9, [])]*(L+1)
    dp[0] = (0, [])
    for i in range(1, L+1):
        best = (10**9, [])
        for b in blocks:
            if i-b < 0:
                continue
            prev_ones, prev_blocks = dp[i-b]
            new_blocks = prev_blocks + [b]
            ones = prev_ones + (1 if b==1 else 0)
            if ones < best[0]:
                best = (ones, new_blocks)
        dp[i] = best
    return dp[L][1]

def contiguous_runs(sorted_hours):
    runs=[]
    if not sorted_hours:
        return runs
    run=[sorted_hours[0]]
    for h in sorted_hours[1:]:
        if h==run[-1]+1:
            run.append(h)
        else:
            runs.append(run)
            run=[h]
    runs.append(run)
    return runs

# Helpers die in meerdere delen gebruikt worden
def normalize_attr(name):
    """Normaliseer attractienaam zodat 'X 2' telt als 'X'; trim & lower-case voor vergelijking."""
    if not name:
        return ""
    s = str(name).strip()
    parts = s.rsplit(" ", 1)
    if len(parts) == 2 and parts[1].isdigit():
        s = parts[0]
    return s.strip().lower()

def parse_header_uur(header):
    """Map headertekst (bv. '14u', '14:00', '14:30') naar het hele uur (14)."""
    if not header:
        return None
    s = str(header).strip()
    try:
        if "u" in s:
            return int(s.split("u")[0])
        if ":" in s:
            uur, _min = s.split(":")
            return int(uur)
        return int(s)
    except:
        return None

# -----------------------------
# Studenten inlezen
# -----------------------------
studenten=[]
for rij in range(2,500):
    naam = ws.cell(rij,12).value
    if not naam:
        continue
    uren_beschikbaar=[10+(kol-3) for kol in range(3,12) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    attracties=[ws.cell(1,kol).value for kol in range(14,32) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    try:
        aantal_attracties=int(ws[f'AG{rij}'].value) if ws[f'AG{rij}'].value else len(attracties)
    except:
        aantal_attracties=len(attracties)
    studenten.append({
        "naam": naam,
        "uren_beschikbaar": sorted(uren_beschikbaar),
        "attracties":[a for a in attracties if a],
        "aantal_attracties":aantal_attracties,
        "is_pauzevlinder":False,
        "pv_number":None,
        "assigned_attracties":set(),
        "assigned_hours":[]
    })

# -----------------------------
# Openingsuren
# -----------------------------
open_uren=[int(str(ws.cell(1,kol).value).replace("u","").strip()) for kol in range(36,45) if ws.cell(2,kol).value in [1,True,"WAAR","X"]]
if not open_uren:
    open_uren=list(range(10,19))
open_uren=sorted(set(open_uren))

# -----------------------------

# Pauzevlinders
# -----------------------------
pauzevlinder_namen=[ws[f'BN{rij}'].value for rij in range(4,11) if ws[f'BN{rij}'].value]

def compute_pauze_hours(open_uren):
    if 10 in open_uren and 18 in open_uren:
        return [h for h in open_uren if 12 <= h <= 16]
    elif 10 in open_uren and 17 in open_uren:
        return [h for h in open_uren if 12 <= h <= 16]
    elif 12 in open_uren:
        return [h for h in open_uren if 13 <= h <= 18]
    else:
        return list(open_uren)

required_pauze_hours=compute_pauze_hours(open_uren)

for idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    for s in studenten:
        if s["naam"]==pvnaam:
            s["is_pauzevlinder"]=True
            s["pv_number"]=idx
            s["uren_beschikbaar"]=[u for u in s["uren_beschikbaar"] if u not in required_pauze_hours]
            break

# Maak 'selected' lijst van pauzevlinders (dicts met naam en attracties)
selected = [s for s in studenten if s.get("is_pauzevlinder")]

# -----------------------------
# Attracties & aantallen (raw)
# -----------------------------
aantallen_raw = {}
attracties_te_plannen = []
for kol in range(47, 65):  # AU-BL
    naam = ws.cell(1, kol).value
    if naam:
        try:
            aantal = int(ws.cell(2, kol).value)
        except:
            aantal = 0
        aantallen_raw[naam] = max(0, min(2, aantal))
        if aantallen_raw[naam] >= 1:
            attracties_te_plannen.append(naam)

# Priority order for second spots (column BA, rows 5-11)
second_priority_order = [
    ws["BA" + str(rij)].value for rij in range(5, 12)
    if ws["BA" + str(rij)].value
]

# -----------------------------
# Compute aantallen per hour + red spots
# -----------------------------
aantallen = {uur: {a: 1 for a in attracties_te_plannen} for uur in open_uren}
red_spots = {uur: set() for uur in open_uren}

for uur in open_uren:
    # Hoeveel studenten beschikbaar dit uur (excl. pauzevlinders op duty)
    student_count = sum(
        1 for s in studenten
        if uur in s["uren_beschikbaar"] and not (
            s["is_pauzevlinder"] and uur in required_pauze_hours
        )
    )
    # Hoeveel attracties minimaal bemand moeten worden
    base_spots = sum(1 for a in attracties_te_plannen if aantallen_raw[a] >= 1)
    extra_spots = student_count - base_spots

    # Allocate 2e plekken volgens prioriteit
    for attr in second_priority_order:
        if attr in aantallen_raw and aantallen_raw[attr] == 2:
            if extra_spots > 0:
                aantallen[uur][attr] = 2
                extra_spots -= 1
            else:
                red_spots[uur].add(attr)

# -----------------------------
# Studenten die effectief inzetbaar zijn
# -----------------------------
studenten_workend = [
    s for s in studenten if any(u in open_uren for u in s["uren_beschikbaar"])
]

# Sorteer attracties op "kritieke score" (hoeveel studenten ze kunnen doen)
def kritieke_score(attr, studenten_list):
    return sum(1 for s in studenten_list if attr in s["attracties"])

attracties_te_plannen.sort(key=lambda a: kritieke_score(a, studenten_workend))

# -----------------------------
# Assign per student
# -----------------------------
assigned_map = defaultdict(list)  # (uur, attr) -> list of student-names
per_hour_assigned_counts = {uur: {a: 0 for a in attracties_te_plannen} for uur in open_uren}
extra_assignments = defaultdict(list)

MAX_CONSEC = 4
MAX_PER_STUDENT_ATTR = 6

studenten_sorted = sorted(studenten_workend, key=lambda s: s["aantal_attracties"])

# -----------------------------
# Voorbereiden: expand naar posities per uur
# -----------------------------
positions_per_hour = {uur: [] for uur in open_uren}
for uur in open_uren:
    for attr in attracties_te_plannen:
        max_pos = aantallen[uur].get(attr, 1)
        for pos in range(1, max_pos+1):
            # sla rode posities over
            if attr in red_spots[uur] and pos == 2:
                continue
            positions_per_hour[uur].append((attr, pos))

# Mapping: welke student staat waar
assigned_map = defaultdict(list)  # (uur, attr) -> [namen]
occupied_positions = {uur: {} for uur in open_uren}  # (attr,pos) -> naam
extra_assignments = defaultdict(list)


# -----------------------------
# Hulpfunctie: kan blok geplaatst worden?
# -----------------------------
def can_place_block(student, block_hours, attr):
    for h in block_hours:
        # check of attractie beschikbaar is in dit uur
        if (attr, 1) not in positions_per_hour[h] and (attr, 2) not in positions_per_hour[h]:
            return False
        # alle posities bezet?
        taken1 = (attr,1) in occupied_positions[h]
        taken2 = (attr,2) in occupied_positions[h]
        if taken1 and taken2:
            return False
    return True

# -----------------------------
# Plaats blok
# -----------------------------
def place_block(student, block_hours, attr):
    for h in block_hours:
        # kies positie: eerst pos1, anders pos2
        if (attr,1) in positions_per_hour[h] and (attr,1) not in occupied_positions[h]:
            pos = 1
        else:
            pos = 2
        occupied_positions[h][(attr,pos)] = student["naam"]
        assigned_map[(h, attr)].append(student["naam"])
        student["assigned_hours"].append(h)
    student["assigned_attracties"].add(attr)


# =============================
# FLEXIBELE BLOKKEN & PLAATSLOGICA
# =============================

def _max_spots_for(attr, uur):
    """Houd rekening met red_spots: 2e plek dicht als het rood is."""
    max_spots = aantallen[uur].get(attr, 1)
    if attr in red_spots.get(uur, set()):
        max_spots = 1
    return max_spots

def _has_capacity(attr, uur):
    return per_hour_assigned_counts[uur][attr] < _max_spots_for(attr, uur)

def _try_place_block_on_attr(student, block_hours, attr):
    """Check capaciteit in alle uren en plaats dan in één keer, met max 4 uur per attractie per dag (positie 1 en 2 tellen samen)."""
    # Capaciteit check
    for h in block_hours:
        if not _has_capacity(attr, h):
            return False
    # Check max 4 unieke uren per attractie per dag (positie 1 en 2 tellen samen)
    # Verzamel alle uren waarop deze student al bij deze attractie staat
    uren_bij_attr = set()
    for h in student["assigned_hours"]:
        namen = assigned_map.get((h, attr), [])
        if student["naam"] in namen:
            uren_bij_attr.add(h)
    # Voeg de nieuwe uren toe
    nieuwe_uren = set(block_hours)
    totaal_uren = uren_bij_attr | nieuwe_uren
    if len(totaal_uren) > 4:
        return False
    # Plaatsen
    for h in block_hours:
        assigned_map[(h, attr)].append(student["naam"])
        per_hour_assigned_counts[h][attr] += 1
        student["assigned_hours"].append(h)
    student["assigned_attracties"].add(attr)
    return True

def _try_place_block_any_attr(student, block_hours):
    """Probeer dit blok te plaatsen op eender welke attractie die student kan."""
    # Eerst attracties die nu het minst keuze hebben (kritiek), zodat we schaarste slim benutten
    candidate_attrs = [a for a in attracties_te_plannen if a in student["attracties"]]
    candidate_attrs.sort(key=lambda a: sum(1 for s in studenten_workend if a in s["attracties"]))
    for attr in candidate_attrs:
        # vermijd dubbele toewijzing van hetzelfde attr als het niet per se moet
        if attr in student["assigned_attracties"]:
            continue
        if _try_place_block_on_attr(student, block_hours, attr):
            return True
    # Als niets lukte zonder herhaling, laat herhaling van attractie toe als dat nodig is
    for attr in candidate_attrs:
        if _try_place_block_on_attr(student, block_hours, attr):
            return True
    return False

def _place_block_with_fallback(student, hours_seq):
    """
    Probeer een reeks opeenvolgende uren te plaatsen.
    - Eerst als blok van 3, anders 2, anders 1.
    - Als niets lukt aan het begin van de reeks, schuif 1 uur op (dat uur gaat voorlopig naar extra),
      en probeer verder; tweede pass zal het later alsnog proberen op te vullen.
    Retourneert: lijst 'unplaced' uren die (voorlopig) niet geplaatst raakten.
    """
    if not hours_seq:
        return []

    # Probeer blok aan de voorkant, groot -> klein
    for size in [3, 2, 1]:
        if len(hours_seq) >= size:
            first_block = hours_seq[:size]
            if _try_place_block_any_attr(student, first_block):
                # Rest recursief
                return _place_block_with_fallback(student, hours_seq[size:])

    # Helemaal niks paste aan de voorkant: markeer eerste uur tijdelijk als 'unplaced' en schuif door
    return [hours_seq[0]] + _place_block_with_fallback(student, hours_seq[1:])



# -----------------------------
# Nieuwe assign_student
# -----------------------------
def assign_student(s):
    """
    Plaats één student in de planning volgens alle regels:
    - Alleen uren waar de student beschikbaar is én open_uren zijn.
    - Geen overlap met pauzevlinder-uren.
    - Alleen attracties die de student kan.
    - Eerst lange blokken proberen (3 uur), dan korter (2 of 1).
    - Blokken die niet passen, gaan voorlopig naar extra_assignments.
    """
    # Filter op effectieve inzetbare uren
    uren = sorted(u for u in s["uren_beschikbaar"] if u in open_uren)
    if s["is_pauzevlinder"]:
        # Verwijder uren waarin pauzevlinder moet werken
        uren = [u for u in uren if u not in required_pauze_hours]

    if not uren:
        return  # geen beschikbare uren

    # Vind aaneengesloten runs van uren
    runs = contiguous_runs(uren)

    for run in runs:
        # Plaats run met fallback (3->2->1), en schuif als het echt niet kan
        unplaced = _place_block_with_fallback(s, run)
        # Wat niet lukte, gaat voorlopig naar extra
        for h in unplaced:
            extra_assignments[h].append(s["naam"])



for s in studenten_sorted:
    assign_student(s)

# -----------------------------
# Post-processing: lege plekken opvullen door doorschuiven
# -----------------------------

def doorschuif_leegplek(uur, attr, pos_idx, student_naam, stap, max_stappen=5):
    if stap > max_stappen:
        return False
    namen = assigned_map.get((uur, attr), [])
    naam = namen[pos_idx-1] if pos_idx-1 < len(namen) else ""
    if naam:
        return False

    kandidaten = []
    for b_attr in attracties_te_plannen:
        b_namen = assigned_map.get((uur, b_attr), [])
        for b_pos, b_naam in enumerate(b_namen):
            if not b_naam or b_naam == student_naam:
                continue
            cand_student = next((s for s in studenten_workend if s["naam"] == b_naam), None)
            if not cand_student:
                continue
            # Mag deze student de lege attractie doen?
            if attr not in cand_student["attracties"]:
                continue
            # Mag de extra de vrijgekomen plek doen?
            extra_student = next((s for s in studenten_workend if s["naam"] == student_naam), None)
            if not extra_student:
                continue
            if b_attr not in extra_student["attracties"]:
                continue
            # Score: zo min mogelijk 1-uursblokken creëren
            uren_cand = sorted([u for u in cand_student["assigned_hours"] if u != uur] + [uur])
            uren_extra = sorted(extra_student["assigned_hours"] + [uur])
            def count_1u_blokken(uren):
                if not uren:
                    return 0
                runs = contiguous_runs(uren)
                return sum(1 for r in runs if len(r) == 1)
            score = count_1u_blokken(uren_cand) + count_1u_blokken(uren_extra)
            kandidaten.append((score, b_attr, b_pos, b_naam, cand_student))
    kandidaten.sort()

    for score, b_attr, b_pos, b_naam, cand_student in kandidaten:
        extra_student = next((s for s in studenten_workend if s["naam"] == student_naam), None)
        if not extra_student:
            continue
        # Voer de swap uit
        assigned_map[(uur, b_attr)][b_pos] = student_naam
        extra_student["assigned_hours"].append(uur)
        extra_student["assigned_attracties"].add(b_attr)
        per_hour_assigned_counts[uur][b_attr] += 0  # netto gelijk
        assigned_map[(uur, attr)].insert(pos_idx-1, b_naam)
        assigned_map[(uur, attr)] = assigned_map[(uur, attr)][:aantallen[uur].get(attr, 1)]
        cand_student["assigned_hours"].remove(uur)
        cand_student["assigned_attracties"].discard(b_attr)
        cand_student["assigned_hours"].append(uur)
        cand_student["assigned_attracties"].add(attr)
        per_hour_assigned_counts[uur][attr] += 0  # netto gelijk
        # Check of alles klopt (geen dubbele, geen restricties overtreden)
        # (optioneel: extra checks toevoegen)
        return True
    return False

max_iterations = 5
for _ in range(max_iterations):
    changes_made = False
    for uur in open_uren:
        for attr in attracties_te_plannen:
            max_pos = aantallen[uur].get(attr, 1)
            if attr in red_spots.get(uur, set()):
                max_pos = 1
            for pos_idx in range(1, max_pos+1):
                namen = assigned_map.get((uur, attr), [])
                naam = namen[pos_idx-1] if pos_idx-1 < len(namen) else ""
                if naam:
                    continue
                # Probeer voor alle extra's op dit uur
                extras_op_uur = list(extra_assignments[uur])  # kopie ivm mutatie
                for extra_naam in extras_op_uur:
                    extra_student = next((s for s in studenten_workend if s["naam"] == extra_naam), None)
                    if not extra_student:
                        continue
                    if attr in extra_student["attracties"]:
                        # Kan direct geplaatst worden, dus hoort niet bij dit scenario
                        continue
                    # Probeer doorschuiven
                    if doorschuif_leegplek(uur, attr, pos_idx, extra_naam, 1, max_iterations):
                        extra_assignments[uur].remove(extra_naam)
                        changes_made = True
                        break  # stop met deze plek, ga naar volgende lege plek
    if not changes_made:
        break



# -----------------------------

# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"

# Witte fill voor headers en attracties
white_fill = PatternFill(start_color="FFFFFF", fill_type="solid")
pv_fill = PatternFill(start_color="FFF2CC", fill_type="solid")
extra_fill = PatternFill(start_color="FCE4D6", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

# Unieke lichte kleuren voor studenten
studenten_namen = sorted({s["naam"] for s in studenten})
light_colors = [
    "FFEBEE", "FFF3E0", "FFFDE7", "E8F5E9", "E3F2FD", "E1F5FE", "F3E5F5", "FCE4EC", "F9FBE7", "E0F2F1",
    "F8BBD0", "B2EBF2", "B3E5FC", "DCEDC8", "F0F4C3", "FFECB3", "FFE0B2", "D7CCC8", "F5F5F5", "CFD8DC"
]
import itertools
student_kleuren = dict(zip(studenten_namen, itertools.cycle(light_colors)))

# Header
ws_out.cell(1, 1, vandaag).font = Font(bold=True)
ws_out.cell(1, 1).fill = white_fill
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1, col_idx, f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1, col_idx).fill = white_fill
    ws_out.cell(1, col_idx).alignment = center_align
    ws_out.cell(1, col_idx).border = thin_border

rij_out = 2
for attr in attracties_te_plannen:
    # FIX: correcte berekening max_pos
    max_pos = max(
        max(aantallen[uur].get(attr, 1) for uur in open_uren),
        max(per_hour_assigned_counts[uur].get(attr, 0) for uur in open_uren)
    )

    for pos_idx in range(1, max_pos + 1):
        naam_attr = attr if max_pos == 1 else f"{attr} {pos_idx}"
        ws_out.cell(rij_out, 1, naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out, 1).fill = white_fill
        ws_out.cell(rij_out, 1).border = thin_border

        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            # Red fill voor ongebruikte 2e plekken
            red_fill = PatternFill(start_color="D9D9D9", fill_type="solid")

            if attr in red_spots.get(uur, set()) and pos_idx == 2:
                ws_out.cell(rij_out, col_idx, "").fill = red_fill
                ws_out.cell(rij_out, col_idx).border = thin_border
            else:
                namen = assigned_map.get((uur, attr), [])
                naam = namen[pos_idx - 1] if pos_idx - 1 < len(namen) else ""
                ws_out.cell(rij_out, col_idx, naam).alignment = center_align
                ws_out.cell(rij_out, col_idx).border = thin_border
                if naam and naam in student_kleuren:
                    ws_out.cell(rij_out, col_idx).fill = PatternFill(start_color=student_kleuren[naam], fill_type="solid")

        rij_out += 1

# Pauzevlinders
rij_out += 1
for pv_idx, pvnaam in enumerate(pauzevlinder_namen, start=1):
    ws_out.cell(rij_out, 1, f"Pauzevlinder {pv_idx}").font = Font(bold=True)
    ws_out.cell(rij_out, 1).fill = pv_fill
    ws_out.cell(rij_out, 1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out, col_idx, pvnaam if uur in required_pauze_hours else "").alignment = center_align
        ws_out.cell(rij_out, col_idx).border = thin_border
    rij_out += 1

# Extra
rij_out += 1
ws_out.cell(rij_out, 1, "Extra").font = Font(bold=True)
ws_out.cell(rij_out, 1).fill = extra_fill
ws_out.cell(rij_out, 1).border = thin_border

for col_idx, uur in enumerate(sorted(open_uren), start=2):
    for r_offset, naam in enumerate(extra_assignments[uur]):
        ws_out.cell(rij_out + 1 + r_offset, col_idx, naam).alignment = center_align

# Kolombreedte
for col in range(1, len(open_uren) + 2):
    ws_out.column_dimensions[get_column_letter(col)].width = 18

# ---- student_totalen beschikbaar maken voor volgende delen ----
from collections import defaultdict
student_totalen = defaultdict(int)
for row in ws_out.iter_rows(min_row=2, values_only=True):
    for naam in row[1:]:
        if naam and str(naam).strip() != "":
            student_totalen[naam] += 1





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

# --- doorschuif debuglog naar feedback sheet ---
try:
    for regel in doorschuif_debuglog:
        log_feedback(regel)
except Exception as e:
    log_feedback(f"[DEBUGLOG ERROR] {e}")




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


output = BytesIO()
wb_out.save(output)
output.seek(0)
# st.success("Planning gegenereerd!")
st.download_button(
    "Download planning",
    data=output.getvalue(),
    file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)

