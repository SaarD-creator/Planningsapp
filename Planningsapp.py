#tussenstap (15/09) enkel nog te kort tijd tussen pauzes pauzevlinders 


# tussenstapje
# niet random, geen extra's, kleurtjes!


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
    st.warning("Upload eerst het Excelbestand met de gegevens om verder te gaan.")
    st.stop()

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
    elif 12 in open_uren and 18 in open_uren:
        return [h for h in open_uren if 13 <= h <= 17]
    elif 14 in open_uren and 18 in open_uren:
        return [h for h in open_uren if 15 <= h <= 17]
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

# Felle, maar lichte pastelkleuren (gelijkmatige felheid, veel variatie)
studenten_namen = sorted({s["naam"] for s in studenten})
# Pauzevlinders krijgen ook een kleur uit het schema
alle_namen = studenten_namen + [pv for pv in pauzevlinder_namen if pv not in studenten_namen]
# Unieke kleuren genereren: als er te weinig kleuren zijn, maak er meer met lichte variatie
base_colors = [
    "FFB3BA", "FFDFBA", "FFFFBA", "BAFFC9", "BAE1FF", "E0BBE4", "957DAD", "D291BC", "FEC8D8", "FFDFD3",
    "B5EAD7", "C7CEEA", "FFDAC1", "E2F0CB", "F6DFEB", "F9E2AE", "B6E2D3", "B6D0E2", "F6E2B3", "F7C5CC",
    "F7E6C5", "C5F7D6", "C5E6F7", "F7F6C5", "F7C5F7", "C5C5F7", "C5F7F7", "F7C5C5", "C5F7C5", "F7E2C5",
    "E2F7C5", "C5F7E2", "E2C5F7", "C5E2F7", "F7C5E2", "F7F7C5", "C5F7F7", "F7C5F7", "C5C5F7", "F7C5C5",
    "C5F7C5", "F7E2C5", "E2F7C5", "C5F7E2", "E2C5F7", "C5E2F7", "F7C5E2", "E2C5F7", "C5F7E2", "E2F7C5"
]
import colorsys
def pastel_variant(hex_color, variant):
    # hex_color: 'RRGGBB', variant: int
    r = int(hex_color[0:2], 16) / 255.0
    g = int(hex_color[2:4], 16) / 255.0
    b = int(hex_color[4:6], 16) / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    # kleine variatie in lichtheid en saturatie
    l = min(1, l + 0.03 * (variant % 3))
    s = max(0.3, s - 0.04 * (variant % 5))
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return f"{int(r2*255):02X}{int(g2*255):02X}{int(b2*255):02X}"

unique_colors = []
needed = len(alle_namen)
variant = 0
while len(unique_colors) < needed:
    for base in base_colors:
        if len(unique_colors) >= needed:
            break
        # voeg lichte variatie toe als nodig
        color = pastel_variant(base, variant) if variant > 0 else base
        if color not in unique_colors:
            unique_colors.append(color)
    variant += 1

student_kleuren = dict(zip(alle_namen, unique_colors))

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
            # Red spots nu wit maken
            if attr in red_spots.get(uur, set()) and pos_idx == 2:
                ws_out.cell(rij_out, col_idx, "").fill = white_fill
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
    ws_out.cell(rij_out, 1, f"Pauzevlinder {pv_idx}").font = Font(bold=True)  # tekst blijft zwart
    ws_out.cell(rij_out, 1).fill = white_fill  # cel wit
    ws_out.cell(rij_out, 1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = pvnaam if uur in required_pauze_hours else ""
        ws_out.cell(rij_out, col_idx, naam).alignment = center_align
        ws_out.cell(rij_out, col_idx).border = thin_border
        if naam and naam in student_kleuren:
            ws_out.cell(rij_out, col_idx).fill = PatternFill(start_color=student_kleuren[naam], fill_type="solid")
    rij_out += 1

# Extra's per rij
rij_out += 1
extras_flat = []
for uur in sorted(open_uren):
    for naam in extra_assignments[uur]:
        if naam not in extras_flat:
            extras_flat.append(naam)
for extra_idx, naam in enumerate(extras_flat, start=1):
    ws_out.cell(rij_out, 1, f"Extra {extra_idx}").font = Font(bold=True)
    ws_out.cell(rij_out, 1).fill = white_fill
    ws_out.cell(rij_out, 1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        # Toon naam alleen als deze extra op dit uur is ingepland
        cell_naam = naam if naam in extra_assignments[uur] else ""
        ws_out.cell(rij_out, col_idx, cell_naam).alignment = center_align
        ws_out.cell(rij_out, col_idx).border = thin_border
        if cell_naam and cell_naam in student_kleuren:
            ws_out.cell(rij_out, col_idx).fill = PatternFill(start_color=student_kleuren[cell_naam], fill_type="solid")
    rij_out += 1

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
# Gebruik compute_pauze_hours/open_uren als basis voor de pauzeplanning-urenrij
uren_rij1 = []
from datetime import datetime, timedelta
if required_pauze_hours:
    start_uur = min(required_pauze_hours)
    eind_uur = max(required_pauze_hours)
    tijd = datetime(2020,1,1,start_uur,0)
    # Laatste pauze mag een kwartier vóór het einde starten
    laatste_pauze = datetime(2020,1,1,eind_uur,30)
    while tijd <= laatste_pauze:
        uren_rij1.append(f"{tijd.hour}u" if tijd.minute==0 else f"{tijd.hour}u{tijd.minute:02d}")
        tijd += timedelta(minutes=15)
else:
    # fallback: gebruik open_uren
    for uur in sorted(open_uren):
        uren_rij1.append(f"{uur}u")

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

# Automatisch de breedte van kolom A instellen op basis van de langste tekst
max_len_colA = 0
for row in range(1, ws_pauze.max_row + 1):
    val = ws_pauze.cell(row, 1).value
    if val:
        max_len_colA = max(max_len_colA, len(str(val)))
# Voeg wat extra ruimte toe
ws_pauze.column_dimensions['A'].width = max(12, max_len_colA + 2)

for col in range(2, len(uren_rij1) + 2):
    ws_pauze.column_dimensions[get_column_letter(col)].width = 10

# Gebruik exact dezelfde open_uren en headers als in deel 1 voor de pauzeplanning
uren_rij1 = []
for uur in sorted(open_uren):
    # Zoek de originele header uit ws_out (de hoofdplanning)
    for col in range(2, ws_out.max_column + 1):
        header = ws_out.cell(1, col).value
        if header and str(header).startswith(str(uur)):
            uren_rij1.append(header)
            break

# Opslaan met dezelfde unieke naam

# Maak in-memory bestand
output = BytesIO()





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
# Dynamisch: alle kolommen waar in rij 1 een uur staat (bv. '13u45', '14u', ...)
pauze_cols = []
for col in range(2, ws_pauze.max_column + 1):
    header = ws_pauze.cell(1, col).value
    if header and ("u" in str(header)):
        pauze_cols.append(col)

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


# ---- OPTIMALISATIE: Meerdere planningen genereren en de beste kiezen ----
import copy
best_score = None
best_state = None
num_runs = 5
for _run in range(num_runs):
    # Maak een deep copy van de relevante werkbladen en variabelen
    ws_pauze_tmp = wb_out.copy_worksheet(ws_pauze)
    ws_pauze_tmp.title = f"Pauzevlinders_tmp_{_run}"
    # Reset alle naamcellen
    for pv, pv_row in pv_rows:
        for col in pauze_cols:
            ws_pauze_tmp.cell(pv_row, col).value = None
    # Herhaal de bestaande logica voor pauzeplanning, maar werk op ws_pauze_tmp
    # ...existing code for pauzeplanning, but use ws_pauze_tmp instead of ws_pauze...
    # (Voor deze patch: laat de bestaande logica staan, dit is een structuurvoorzet. Zie opmerking hieronder)
    # ---- Evalueer deze planning ----
    # 1. Iedereen een pauze?
    korte_pauze_ontvangers = set()
    for pv, pv_row in pv_rows:
        for col in pauze_cols:
            cel = ws_pauze_tmp.cell(pv_row, col)
            if cel.value and str(cel.value).strip() != "":
                # Check of het een korte pauze is (enkel blok, niet dubbel)
                idx = pauze_cols.index(col)
                is_lange = False
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze_tmp.cell(pv_row, next_col)
                    if cel_next.value == cel.value:
                        is_lange = True
                if idx > 0:
                    prev_col = pauze_cols[idx-1]
                    prev_cel = ws_pauze_tmp.cell(pv_row, prev_col)
                    if prev_cel.value == cel.value:
                        is_lange = True
                if not is_lange:
                    korte_pauze_ontvangers.add(str(cel.value).strip())
    alle_studenten = [s["naam"] for s in studenten if student_totalen.get(s["naam"], 0) > 0]
    iedereen_pauze = all(naam in korte_pauze_ontvangers for naam in alle_studenten)
    # 2. Eerlijkheid: verschil max-min korte pauzes per pauzevlinder
    from collections import Counter
    pv_korte_pauze_count = Counter()
    for pv, pv_row in pv_rows:
        for col in pauze_cols:
            cel = ws_pauze_tmp.cell(pv_row, col)
            if cel.value and str(cel.value).strip() != "":
                idx = pauze_cols.index(col)
                is_lange = False
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze_tmp.cell(pv_row, next_col)
                    if cel_next.value == cel.value:
                        is_lange = True
                if idx > 0:
                    prev_col = pauze_cols[idx-1]
                    prev_cel = ws_pauze_tmp.cell(pv_row, prev_col)
                    if prev_cel.value == cel.value:
                        is_lange = True
                if not is_lange:
                    pv_korte_pauze_count[pv["naam"]] += 1
    if pv_korte_pauze_count:
        eerlijkheid = max(pv_korte_pauze_count.values()) - min(pv_korte_pauze_count.values())
    else:
        eerlijkheid = 999
    # Score: eerst iedereen_pauze, dan eerlijkheid
    score = (iedereen_pauze, -eerlijkheid)
    if (best_score is None) or (score > best_score):
        best_score = score
        best_state = copy.deepcopy(ws_pauze_tmp)

# Na alle runs: kopieer best_state naar ws_pauze
if best_state is not None:

    for pv, pv_row in pv_rows:
        for col in pauze_cols:
            ws_pauze.cell(pv_row, col).value = best_state.cell(pv_row, col).value

# ---- Verwijder tijdelijke werkbladen ----
tmp_sheets = [ws for ws in wb_out.worksheets if ws.title.startswith("Pauzevlinders_tmp")]
for ws in tmp_sheets:
    wb_out.remove(ws)

# ---- Lege naamcellen inkleuren ----
naam_leeg_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
center_align = Alignment(horizontal="center", vertical="center")

for pv, pv_row in pv_rows:
    for col in pauze_cols:
        if ws_pauze.cell(pv_row, col).value in [None, ""]:
            ws_pauze.cell(pv_row, col).fill = naam_leeg_fill






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
# Dynamisch: alle kolommen waar in rij 1 een uur staat (bv. '13u45', '14u', ...)
pauze_cols = []
for col in range(2, ws_pauze.max_column + 1):
    header = ws_pauze.cell(1, col).value
    if header and ("u" in str(header)):
        pauze_cols.append(col)


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


def is_pauzevlinder(naam):
    """Is deze naam een pauzevlinder?"""
    return any(pv["naam"] == naam for pv in selected)



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

# ---- Pauze-restrictie: geen korte pauze in eerste 12 kwartieren van de pauzeplanning (tenzij <=6u open) ----
def get_verboden_korte_pauze_kolommen():
    """Geeft de kolomnummers van de eerste 12 kwartieren in ws_pauze (B t/m M)."""
    return list(range(2, 14))  # kolommen 2 t/m 13 (B t/m M)

def is_korte_pauze_toegestaan_col(col):
    if len(open_uren) <= 6:
        return True
    return col not in get_verboden_korte_pauze_kolommen()

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

# Iedereen met '-18' in de naam krijgt altijd een halfuurpauze
lange_werkers = [s for s in studenten
    if (
        student_totalen.get(s["naam"], 0) > 6
        or ("-18" in str(s["naam"]) and student_totalen.get(s["naam"], 0) > 0)
    )
    and s["naam"] not in [pv["naam"] for pv in selected]
]
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
    if base == "extra":
        return True
    return base in pv_cap_set.get(pv["naam"], set())

# Willekeurige volgorde van pauzeplekken (pv-rij x kolom) om lege cellen random te spreiden
slot_order = [(pv, pv_row, col) for (pv, pv_row) in pv_rows for col in pauze_cols]
random.shuffle(slot_order)  # <— kern om lege plekken later willekeurig te verspreiden

def plaats_student(student, harde_mode=False):
    """
    Plaats student bij een geschikte pauzevlinder in B..G op een uur waar student effectief werkt.
    - Overschrijven alleen in harde_mode én alleen als de huidige inhoud géén lange werker is.
    - Volgorde van slots is willekeurig (slot_order) zodat lege plekken random verdeeld blijven.
    """
    naam = student["naam"]
    werk_uren = get_student_work_hours(naam)  # echte uren waarop student in 'Planning' staat
    # Pauze mag niet in eerste of laatste werkuur vallen
    werk_uren_set = set(werk_uren)
    if len(werk_uren) > 2:
        verboden_uren = {werk_uren[0], werk_uren[-1]}
    else:
        verboden_uren = set(werk_uren)  # als maar 1 of 2 uur: geen pauze mogelijk

    # Sorteer alle pauzekolommen op volgorde
    pauze_cols_sorted = sorted(pauze_cols)
    # Zoek alle (uur, col) paren, filter verboden uren
    uur_col_pairs = []
    for col in pauze_cols_sorted:
        col_header = ws_pauze.cell(1, col).value
        col_uur = parse_header_uur(col_header)
        if col_uur is not None and col_uur in werk_uren_set and col_uur not in verboden_uren:
            uur_col_pairs.append((col_uur, col))

    # Houd bij of deze student al een lange/korte pauze heeft gekregen
    if not hasattr(plaats_student, "pauze_registry"):
        plaats_student.pauze_registry = {}
    reg = plaats_student.pauze_registry.setdefault(naam, {"lange": False, "korte": False})

    # Eerst: zoek alle mogelijke dubbele blokjes voor de lange pauze
    lange_pauze_opties = []
    for i in range(len(uur_col_pairs)-1):
        uur1, col1 = uur_col_pairs[i]
        uur2, col2 = uur_col_pairs[i+1]
        if col2 == col1 + 1:
            lange_pauze_opties.append((i, uur1, col1, uur2, col2))

    # Probeer alle opties voor de lange pauze (max 1x per student)
    if not reg["lange"] and not heeft_al_lange_pauze(naam):
        for optie in lange_pauze_opties:
            i, uur1, col1, uur2, col2 = optie
            attr1 = vind_attractie_op_uur(naam, uur1)
            attr2 = vind_attractie_op_uur(naam, uur2)
            if not attr1 or not attr2:
                continue
            for (pv, pv_row, _) in slot_order:
                if not pv_kan_attr(pv, attr1) and not is_student_extra(naam):
                    continue
                cel1 = ws_pauze.cell(pv_row, col1)
                cel2 = ws_pauze.cell(pv_row, col2)
                boven_cel1 = ws_pauze.cell(pv_row-1, col1)
                boven_cel2 = ws_pauze.cell(pv_row-1, col2)
                if cel1.value in [None, ""] and cel2.value in [None, ""] and not heeft_al_lange_pauze(naam):
                    # Vul beide blokjes voor lange pauze
                    boven_cel1.value = attr1
                    boven_cel1.alignment = center_align
                    boven_cel1.border = thin_border
                    boven_cel2.value = attr2
                    boven_cel2.alignment = center_align
                    boven_cel2.border = thin_border
                    cel1.value = naam
                    cel1.alignment = center_align
                    cel1.border = thin_border
                    cel2.value = naam
                    cel2.alignment = center_align
                    cel2.border = thin_border
                    reg["lange"] = True
                    # Nu: zoek een korte pauze, eerst exact 10 blokjes afstand, dan 11, 12, ... tot einde, daarna 9, 8, ... tot 1
                    if not reg["korte"]:
                        found = False
                        # Eerst exact 10 blokjes afstand
                        for min_blokjes in list(range(10, len(uur_col_pairs)-i)) + list(range(9, 0, -1)):
                            for j in range(i+min_blokjes, len(uur_col_pairs)):
                                uur_kort, col_kort = uur_col_pairs[j]
                                if not is_korte_pauze_toegestaan_col(col_kort):
                                    continue
                                attr_kort = vind_attractie_op_uur(naam, uur_kort)
                                if not attr_kort:
                                    continue
                                # Belangrijk: alleen op deze PV-rij plaatsen als de pauzevlinder deze attractie kan, behalve bij EXTRA
                                if (not pv_kan_attr(pv, attr_kort)) and (not is_student_extra(naam)):
                                    continue
                                # Alleen zoeken in dezelfde rij als de lange pauze (dus bij dezelfde pauzevlinder)
                                cel_kort = ws_pauze.cell(pv_row, col_kort)
                                boven_cel_kort = ws_pauze.cell(pv_row-1, col_kort)
                                if cel_kort.value in [None, ""]:
                                    boven_cel_kort.value = attr_kort
                                    boven_cel_kort.alignment = center_align
                                    boven_cel_kort.border = thin_border
                                    cel_kort.value = naam
                                    cel_kort.alignment = center_align
                                    cel_kort.border = thin_border
                                    reg["korte"] = True
                                    found = True
                                    return True
                                elif harde_mode:
                                    occupant = str(cel_kort.value).strip() if cel_kort.value else ""
                                    if occupant not in lange_werkers_names:
                                        boven_cel_kort.value = attr_kort
                                        boven_cel_kort.alignment = center_align
                                        boven_cel_kort.border = thin_border
                                        cel_kort.value = naam
                                        cel_kort.alignment = center_align
                                        cel_kort.border = thin_border
                                        reg["korte"] = True
                                        found = True
                                        return True
                            if found:
                                break
                    # Geen korte pauze gevonden, maar lange pauze is wel gezet
                    return True
                elif harde_mode:
                    occupant1 = str(cel1.value).strip() if cel1.value else ""
                    occupant2 = str(cel2.value).strip() if cel2.value else ""
                    if (occupant1 not in lange_werkers_names) and (occupant2 not in lange_werkers_names) and not heeft_al_lange_pauze(naam):
                        boven_cel1.value = attr1
                        boven_cel1.alignment = center_align
                        boven_cel1.border = thin_border
                        boven_cel2.value = attr2
                        boven_cel2.alignment = center_align
                        boven_cel2.border = thin_border
                        cel1.value = naam
                        cel1.alignment = center_align
                        cel1.border = thin_border
                        cel2.value = naam
                        cel2.alignment = center_align
                        cel2.border = thin_border
                        reg["lange"] = True
                        # Nu: zoek een korte pauze minstens 6 blokjes verderop
                        if not reg["korte"]:
                            for j in range(i+6, len(uur_col_pairs)):
                                uur_kort, col_kort = uur_col_pairs[j]
                                attr_kort = vind_attractie_op_uur(naam, uur_kort)
                                if not attr_kort:
                                    continue
                                for (pv2, pv_row2, _) in slot_order:
                                    if not pv_kan_attr(pv2, attr_kort) and not is_student_extra(naam):
                                        continue
                                    cel_kort = ws_pauze.cell(pv_row2, col_kort)
                                    boven_cel_kort = ws_pauze.cell(pv_row2-1, col_kort)
                                    if cel_kort.value in [None, ""]:
                                        boven_cel_kort.value = attr_kort
                                        boven_cel_kort.alignment = center_align
                                        boven_cel_kort.border = thin_border
                                        cel_kort.value = naam
                                        cel_kort.alignment = center_align
                                        cel_kort.border = thin_border
                                        reg["korte"] = True
                                        return True
                                    elif harde_mode:
                                        occupant = str(cel_kort.value).strip() if cel_kort.value else ""
                                        if occupant not in lange_werkers_names:
                                            boven_cel_kort.value = attr_kort
                                            boven_cel_kort.alignment = center_align
                                            boven_cel_kort.border = thin_border
                                            cel_kort.value = naam
                                            cel_kort.alignment = center_align
                                            cel_kort.border = thin_border
                                            reg["korte"] = True
                                            return True
                        return True
    # Als geen geldige combinatie gevonden, probeer fallback (oude logica)
    # Korte pauze alleen als nog niet toegekend
    for uur in random.sample(werk_uren, len(werk_uren)):
        if uur in verboden_uren:
            continue
        attr = vind_attractie_op_uur(naam, uur)
        if not attr:
            continue
        for (pv, pv_row, col) in slot_order:
            col_header = ws_pauze.cell(1, col).value
            col_uur = parse_header_uur(col_header)
            if col_uur != uur:
                continue
            if not is_korte_pauze_toegestaan_col(col):
                continue
            if not pv_kan_attr(pv, attr) and not is_student_extra(naam):
                continue
            cel = ws_pauze.cell(pv_row, col)
            boven_cel = ws_pauze.cell(pv_row - 1, col)
            current_val = cel.value
            if current_val in [None, ""]:
                if not reg["korte"]:
                    boven_cel.value = attr
                    boven_cel.alignment = center_align
                    boven_cel.border = thin_border
                    cel.value = naam
                    cel.alignment = center_align
                    cel.border = thin_border
                    reg["korte"] = True
                    return True
            else:
                if harde_mode:
                    occupant = str(current_val).strip()
                    if occupant not in lange_werkers_names:
                        if not reg["korte"]:
                            boven_cel.value = attr
                            boven_cel.alignment = center_align
                            boven_cel.border = thin_border
                            cel.value = naam
                            cel.alignment = center_align
                            cel.border = thin_border
                            reg["korte"] = True
                            return True
    return False

# ---- Fase 1: zachte toewijzing (niet overschrijven) ----
def heeft_al_lange_pauze(naam):
    # Check of student al een dubbele blok (lange pauze) heeft in het pauzeoverzicht
    for pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value == naam:
                # Check of volgende cel ook deze naam heeft (dubbele blok)
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze.cell(pv_row, next_col)
                    if cel_next.value == naam:
                        return True
    return False

lichtgroen_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")  # lange pauze
lichtpaars_fill = PatternFill(start_color="E6DAF7", end_color="E6DAF7", fill_type="solid")  # kwartierpauze
naam_leeg_fill = PatternFill(start_color="CCE5FF", end_color="CCE5FF", fill_type="solid")
lange_pauze_ontvangers = set()
# --- Verspreid lange pauzes van lange werkers net als bij pauzevlinders ---
niet_geplaatst = []
for s in random.sample(lange_werkers, len(lange_werkers)):
    naam = s["naam"]
    if naam in lange_pauze_ontvangers or heeft_al_lange_pauze(naam):
        if not plaats_student(s, harde_mode=False):
            niet_geplaatst.append(s)
        continue
    werk_uren = get_student_work_hours(naam)
    if len(werk_uren) <= 6:
        if not plaats_student(s, harde_mode=False):
            niet_geplaatst.append(s)
        continue
    # Alleen de eerste 11 kwartieren (indices 0 t/m 10) zijn toegestaan voor lange pauzes
    halve_uren = []  # lijst van (col1, col2, uur1, uur2, pv, pv_row)
    werk_uren_set = set(werk_uren)
    verboden_uren = {werk_uren[0], werk_uren[-1]} if len(werk_uren) > 2 else set(werk_uren)
    max_start_idx = min(10, len(pauze_cols)-2)  # idx 0 t/m 10 zijn halve uren binnen eerste 11 kwartieren
    for pv, pv_row in pv_rows:
        for idx in range(max_start_idx+1):
            col1 = pauze_cols[idx]
            col2 = pauze_cols[idx+1]
            col1_header = ws_pauze.cell(1, col1).value
            col2_header = ws_pauze.cell(1, col2).value
            # Alleen starten op heel of half uur
            try:
                min1 = int(str(col1_header).split('u')[1]) if 'u' in str(col1_header) and len(str(col1_header).split('u')) > 1 else 0
            except:
                min1 = 0
            if min1 not in (0, 30):
                continue
            uur1 = parse_header_uur(col1_header)
            uur2 = parse_header_uur(col2_header)
            if uur1 is None or uur2 is None:
                continue
            if uur1 not in werk_uren_set or uur2 not in werk_uren_set:
                continue
            if uur1 in verboden_uren or uur2 in verboden_uren:
                continue
            cel1 = ws_pauze.cell(pv_row, col1)
            cel2 = ws_pauze.cell(pv_row, col2)
            # Attractie moet kloppen
            attr1 = vind_attractie_op_uur(naam, uur1)
            attr2 = vind_attractie_op_uur(naam, uur2)
            if not attr1 or not attr2:
                continue
            if not pv_kan_attr(pv, attr1) and not is_student_extra(naam):
                continue
            if cel1.value in [None, ""] and cel2.value in [None, ""]:
                halve_uren.append((col1, col2, uur1, uur2, pv, pv_row))
    random.shuffle(halve_uren)
    # Fairness: keep a live counter of long breaks per pauzevlinder
    from collections import Counter
    if not hasattr(plaats_student, "pv_lange_pauze_count"):
        plaats_student.pv_lange_pauze_count = Counter()
    pv_lange_pauze_count = plaats_student.pv_lange_pauze_count
    for pv, _ in pv_rows:
        if pv["naam"] not in pv_lange_pauze_count:
            pv_lange_pauze_count[pv["naam"]] = 0
    geplaatst = False
    # Sorteer bij elke poging op actuele telling
    halve_uren_sorted = sorted(halve_uren, key=lambda x: pv_lange_pauze_count[x[4]["naam"]])
    for col1, col2, uur1, uur2, pv, pv_row in halve_uren_sorted:
        cel1 = ws_pauze.cell(pv_row, col1)
        cel2 = ws_pauze.cell(pv_row, col2)
        boven_cel1 = ws_pauze.cell(pv_row-1, col1)
        boven_cel2 = ws_pauze.cell(pv_row-1, col2)
        attr1 = vind_attractie_op_uur(naam, uur1)
        attr2 = vind_attractie_op_uur(naam, uur2)
        if cel1.value in [None, ""] and cel2.value in [None, ""] and not heeft_al_lange_pauze(naam):
            boven_cel1.value = attr1
            boven_cel1.alignment = center_align
            boven_cel1.border = thin_border
            boven_cel2.value = attr2
            boven_cel2.alignment = center_align
            boven_cel2.border = thin_border
            cel1.value = naam
            cel1.alignment = center_align
            cel1.border = thin_border
            cel1.fill = lichtgroen_fill
            cel2.value = naam
            cel2.alignment = center_align
            cel2.border = thin_border
            cel2.fill = lichtgroen_fill
            lange_pauze_ontvangers.add(naam)
            geplaatst = True
            # Fairness: niet meetellen als deze lange pauze (een van de twee blokken) een 'Extra' overname is
            if (normalize_attr(attr1) != 'extra') and (normalize_attr(attr2) != 'extra'):
                pv_lange_pauze_count[pv["naam"]] += 1
            break
    if not geplaatst:
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

# ---- Pauze-kleuren: lichtgroen voor lange pauze (dubbele blok), lichtpaars voor kwartierpauze (enkel blok) ----

lichtgroen_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")  # lange pauze
lichtpaars_fill = PatternFill(start_color="E6DAF7", end_color="E6DAF7", fill_type="solid")  # kwartierpauze

# Pauze kleuren invullen (lange en korte pauzes)
for pv, pv_row in pv_rows:
    for idx, col in enumerate(pauze_cols):
        cel = ws_pauze.cell(pv_row, col)
        if cel.value in [None, ""]:
            cel.fill = naam_leeg_fill
        else:
            # Check of dit een lange pauze is (dubbele blok: zelfde naam in 2 naast elkaar liggende cellen)
            is_langepauze = False
            # Kijk vooruit: als deze en de volgende cel dezelfde naam hebben, kleur beide groen
            if idx+1 < len(pauze_cols):
                next_col = pauze_cols[idx+1]
                cel_next = ws_pauze.cell(pv_row, next_col)
                if cel_next.value == cel.value and cel.value not in [None, ""]:
                    is_langepauze = True
                    cel.fill = lichtgroen_fill
                    cel_next.fill = lichtgroen_fill
                    continue  # sla de volgende cel over, die is al gekleurd
            # Kijk achteruit: als vorige cel al groen is door lange pauze, deze niet opnieuw kleuren
            if idx > 0:
                prev_col = pauze_cols[idx-1]
                prev_cel = ws_pauze.cell(pv_row, prev_col)
                if prev_cel.value == cel.value and cel.value not in [None, ""]:
                    # Deze cel is al als tweede helft van lange pauze gekleurd
                    continue
            # Anders: kwartierpauze
            cel.fill = lichtpaars_fill

# ---- Korte pauze voor pauzevlinders zelf toevoegen (eerst, met afstandscriterium) ----
def _pv_has_short_pause(naam, pv_row):
    for idx, col in enumerate(pauze_cols):
        if ws_pauze.cell(pv_row, col).value == naam:
            left_same = idx > 0 and ws_pauze.cell(pv_row, pauze_cols[idx-1]).value == naam
            right_same = idx+1 < len(pauze_cols) and ws_pauze.cell(pv_row, pauze_cols[idx+1]).value == naam
            if not left_same and not right_same:
                return True
    return False

def _pv_place_best_short(pv, pv_row, target_gap=10):
    """Plaats korte pauze op eigen rij met voorkeur: exact +10 blokken na lange pauze-einde,
    anders +11, +12, ...; als dat niet past, probeer +9, +8, ...; valt terug op globale laatste lange-pauze-einde als geen eigen lange pauze."""
    naam = pv["naam"]
    # Als er al een korte pauze staat, niets doen
    if _pv_has_short_pause(naam, pv_row):
        return False

    # Hulpfuncties
    def is_toegestaan_pv_col(col):
        if len(open_uren) <= 6:
            return True
        return col not in get_verboden_korte_pauze_kolommen()

    # Zoek eigen lange pauze-einde op deze rij
    lange_blok_einde = None
    i = 0
    while i < len(pauze_cols)-1:
        c1 = pauze_cols[i]
        c2 = pauze_cols[i+1]
        if ws_pauze.cell(pv_row, c1).value == naam and ws_pauze.cell(pv_row, c2).value == naam:
            lange_blok_einde = i+1
            # ga door; kies de laatste indien meerdere
            i += 2
        else:
            i += 1

    # Geen eigen lange pauze: kies een geldige plek op eigen rij (liefst vroegste index >= target_gap)
    if lange_blok_einde is None:
        werk_uren = get_student_work_hours(naam)
        if len(werk_uren) > 2:
            verboden_uren = {werk_uren[0], werk_uren[-1]}
        else:
            verboden_uren = set(werk_uren)
        candidates = []  # (prefer, idx, col, uur)
        for i, col in enumerate(pauze_cols):
            header = ws_pauze.cell(1, col).value
            uur = parse_header_uur(header)
            if uur not in werk_uren or uur in verboden_uren:
                continue
            if not is_toegestaan_pv_col(col):
                continue
            if ws_pauze.cell(pv_row, col).value not in [None, ""]:
                continue
            prefer = 1 if i >= target_gap else 0
            candidates.append((prefer, i, col, uur))
        if not candidates:
            return False
        # Kies met voorkeur voor index >= target_gap; binnen die groep de laatste (grootste index) om niet te vroeg te vallen
        candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        _pref, i, col, uur = candidates[0]
        attr = vind_attractie_op_uur(naam, uur)
        if not attr:
            return False
        # Voor PV-korte pauzes: laat het vakje erboven leeg
        ws_pauze.cell(pv_row-1, col).value = None
        ws_pauze.cell(pv_row-1, col).alignment = center_align
        ws_pauze.cell(pv_row-1, col).border = thin_border
        cel = ws_pauze.cell(pv_row, col)
        cel.value = naam
        cel.fill = lichtpaars_fill
        cel.alignment = center_align
        cel.border = thin_border
        return True
    else:
        anchor_idx = lange_blok_einde

    if anchor_idx is None or anchor_idx < 0:
        # Geen anchor beschikbaar: kies de eerste toegestane lege cel op eigen rij (zeldzaam)
        werk_uren = get_student_work_hours(naam)
        if len(werk_uren) > 2:
            verboden_uren = {werk_uren[0], werk_uren[-1]}
        else:
            verboden_uren = set(werk_uren)
        for i, col in enumerate(pauze_cols):
            header = ws_pauze.cell(1, col).value
            uur = parse_header_uur(header)
            if uur not in werk_uren or uur in verboden_uren:
                continue
            if is_toegestaan_pv_col(col) and ws_pauze.cell(pv_row, col).value in [None, ""]:
                # schrijf bovenliggende attr
                attr = vind_attractie_op_uur(naam, uur)
                if not attr:
                    continue
                # Voor PV-korte pauzes: laat het vakje erboven leeg
                ws_pauze.cell(pv_row-1, col).value = None
                ws_pauze.cell(pv_row-1, col).alignment = center_align
                ws_pauze.cell(pv_row-1, col).border = thin_border
                cel = ws_pauze.cell(pv_row, col)
                cel.value = naam
                cel.fill = lichtpaars_fill
                cel.alignment = center_align
                cel.border = thin_border
                return True
        return False

    # Eerst exact +10 blokken, dan +11, +12, ...
    werk_uren = get_student_work_hours(naam)
    if len(werk_uren) > 2:
        verboden_uren = {werk_uren[0], werk_uren[-1]}
    else:
        verboden_uren = set(werk_uren)
    for d in range(target_gap, len(pauze_cols)-anchor_idx):
        idx = anchor_idx + d
        if idx >= len(pauze_cols):
            break
        col = pauze_cols[idx]
        if not is_toegestaan_pv_col(col):
            continue
        header = ws_pauze.cell(1, col).value
        uur = parse_header_uur(header)
        if uur not in werk_uren or uur in verboden_uren:
            continue
        if ws_pauze.cell(pv_row, col).value in [None, ""]:
            # schrijf bovenliggende attr
            attr = vind_attractie_op_uur(naam, uur)
            if not attr:
                continue
            # Voor PV-korte pauzes: laat het vakje erboven leeg
            ws_pauze.cell(pv_row-1, col).value = None
            ws_pauze.cell(pv_row-1, col).alignment = center_align
            ws_pauze.cell(pv_row-1, col).border = thin_border
            cel = ws_pauze.cell(pv_row, col)
            cel.value = naam
            cel.fill = lichtpaars_fill
            cel.alignment = center_align
            cel.border = thin_border
            return True

    # Dan lagere alternatieven: +9, +8, ... +1
    for d in range(target_gap-1, 0, -1):
        idx = anchor_idx + d
        if 0 <= idx < len(pauze_cols):
            col = pauze_cols[idx]
            if not is_toegestaan_pv_col(col):
                continue
            header = ws_pauze.cell(1, col).value
            uur = parse_header_uur(header)
            if uur not in werk_uren or uur in verboden_uren:
                continue
            if ws_pauze.cell(pv_row, col).value in [None, ""]:
                attr = vind_attractie_op_uur(naam, uur)
                if not attr:
                    continue
                # Voor PV-korte pauzes: laat het vakje erboven leeg
                ws_pauze.cell(pv_row-1, col).value = None
                ws_pauze.cell(pv_row-1, col).alignment = center_align
                ws_pauze.cell(pv_row-1, col).border = thin_border
                cel = ws_pauze.cell(pv_row, col)
                cel.value = naam
                cel.fill = lichtpaars_fill
                cel.alignment = center_align
                cel.border = thin_border
                return True

    return False

for pv, pv_row in pv_rows:
    _pv_place_best_short(pv, pv_row, target_gap=8)


# ---- Korte pauze voor ALLE studenten (ook <=6u, behalve pauzevlinders en lange werkers) ----
# --- Houd bij wie al een korte pauze heeft gekregen ---
korte_pauze_ontvangers = set()
# Zoek alle namen die al een korte pauze hebben in het pauzeoverzicht
for pv, pv_row in pv_rows:
    for col in pauze_cols:
        cel = ws_pauze.cell(pv_row, col)
        if cel.value and str(cel.value).strip() != "":
            # Check of het een korte pauze is (enkel blok, niet dubbel)
            idx = pauze_cols.index(col)
            is_lange = False
            if idx+1 < len(pauze_cols):
                next_col = pauze_cols[idx+1]
                cel_next = ws_pauze.cell(pv_row, next_col)
                if cel_next.value == cel.value:
                    is_lange = True
            if idx > 0:
                prev_col = pauze_cols[idx-1]
                prev_cel = ws_pauze.cell(pv_row, prev_col)
                if prev_cel.value == cel.value:
                    is_lange = True
            if not is_lange:
                korte_pauze_ontvangers.add(str(cel.value).strip())



# ---- Korte pauze voor ALLE studenten (ook <=6u, behalve pauzevlinders en lange werkers) ----
# --- Houd bij wie al een korte pauze heeft gekregen ---
korte_pauze_ontvangers = set()
# Zoek alle namen die al een korte pauze hebben in het pauzeoverzicht
for pv, pv_row in pv_rows:
    for col in pauze_cols:
        cel = ws_pauze.cell(pv_row, col)
        if cel.value and str(cel.value).strip() != "":
            # Check of het een korte pauze is (enkel blok, niet dubbel)
            idx = pauze_cols.index(col)
            is_lange = False
            if idx+1 < len(pauze_cols):
                next_col = pauze_cols[idx+1]
                cel_next = ws_pauze.cell(pv_row, next_col)
                if cel_next.value == cel.value:
                    is_lange = True
            if idx > 0:
                prev_col = pauze_cols[idx-1]
                prev_cel = ws_pauze.cell(pv_row, prev_col)
                if prev_cel.value == cel.value:
                    is_lange = True
            if not is_lange:
                korte_pauze_ontvangers.add(str(cel.value).strip())


# Nieuwe logica: eerlijke verdeling van korte pauzes over pauzevlinders
from collections import Counter

# Tel per pauzevlinder het aantal korte pauzes dat al is toegekend (EXTRA niet meetellen)
pv_korte_pauze_count = Counter()
for pv, pv_row in pv_rows:
    for col in pauze_cols:
        cel = ws_pauze.cell(pv_row, col)
        if cel.value and str(cel.value).strip() != "":
            # Check of het een korte pauze is (enkel blok, niet dubbel)
            idx = pauze_cols.index(col)
            is_lange = False
            if idx+1 < len(pauze_cols):
                next_col = pauze_cols[idx+1]
                cel_next = ws_pauze.cell(pv_row, next_col)
                if cel_next.value == cel.value:
                    is_lange = True
            if idx > 0:
                prev_col = pauze_cols[idx-1]
                prev_cel = ws_pauze.cell(pv_row, prev_col)
                if prev_cel.value == cel.value:
                    is_lange = True
            if not is_lange:
                # Kijk naar bovenliggende attractie; tel niet als dit 'Extra' is of leeg (zoals bij PV zelf)
                attr_above = ws_pauze.cell(pv_row-1, col).value
                if attr_above and normalize_attr(attr_above) != 'extra':
                    pv_korte_pauze_count[pv["naam"]] += 1

niet_geplaatste_korte_pauze = []

# --- NIEUW: Eerst korte pauzes toewijzen aan iedereen met een LANGE pauze,
# in volgorde van wie het LAATST z'n lange pauze had ---

def _has_long_pause(naam):
    for _pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols[:-1]):
            if ws_pauze.cell(pv_row, col).value == naam and ws_pauze.cell(pv_row, pauze_cols[idx+1]).value == naam:
                return True
    return False

def _last_long_pause_end_index(naam):
    """Geef de hoogste kolomindex (in pauze_cols) die een tweede helft is van een dubbele blok voor deze naam; -1 indien geen lange pauze."""
    last_idx = -1
    for _pv, pv_row in pv_rows:
        for idx in range(len(pauze_cols)-1):
            c1 = pauze_cols[idx]
            c2 = pauze_cols[idx+1]
            if ws_pauze.cell(pv_row, c1).value == naam and ws_pauze.cell(pv_row, c2).value == naam:
                last_idx = max(last_idx, idx+1)
    return last_idx

def _has_short_pause(naam):
    for _pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            if ws_pauze.cell(pv_row, col).value == naam:
                # korte pauze: geen buur met dezelfde naam
                left_same = (idx > 0 and ws_pauze.cell(pv_row, pauze_cols[idx-1]).value == naam)
                right_same = (idx+1 < len(pauze_cols) and ws_pauze.cell(pv_row, pauze_cols[idx+1]).value == naam)
                if not left_same and not right_same:
                    return True
    return False

def _place_short_pause_for(naam):
    if _has_short_pause(naam):
        return True
    werk_uren = get_student_work_hours(naam)
    if not werk_uren:
        return False
    verboden_uren = {werk_uren[0], werk_uren[-1]} if len(werk_uren) > 2 else set(werk_uren)
    # Zoek anker = einde van eigen lange pauze (kolomindex in pauze_cols)
    def _last_long_end_index_for(naam):
        best = -1
        for _pv, _row in pv_rows:
            for idx in range(len(pauze_cols)-1):
                if ws_pauze.cell(_row, pauze_cols[idx]).value == naam and ws_pauze.cell(_row, pauze_cols[idx+1]).value == naam:
                    best = max(best, idx+1)
        return best
    anchor = _last_long_end_index_for(naam)

    # Helper om te checken en plaatsen op gewenste col
    def _try_place_at_col(col):
        header = ws_pauze.cell(1, col).value
        uur = parse_header_uur(header)
        if uur not in werk_uren or uur in verboden_uren:
            return False
        if not is_korte_pauze_toegestaan_col(col):
            return False
        attr = vind_attractie_op_uur(naam, uur)
        if not attr:
            return False
        # verzamel geldige (pv_row) volgens regels
        rows = []
        for pv, pv_row in pv_rows:
            if is_pauzevlinder(naam) and pv["naam"] != naam:
                continue
            if ws_pauze.cell(pv_row, col).value not in [None, ""]:
                continue
            if not pv_kan_attr(pv, attr) and not is_student_extra(naam):
                continue
            rows.append((pv, pv_row))
        if not rows:
            return False
        # fairness op pv-rijen
        rows.sort(key=lambda r: pv_korte_pauze_count[r[0]["naam"]])
        pv, pv_row = rows[0]
        # Voor PV-korte pauzes: laat het vakje erboven leeg
        if is_pauzevlinder(naam):
            ws_pauze.cell(pv_row-1, col).value = None
        else:
            ws_pauze.cell(pv_row-1, col).value = attr
        ws_pauze.cell(pv_row-1, col).alignment = center_align
        ws_pauze.cell(pv_row-1, col).border = thin_border
        cel = ws_pauze.cell(pv_row, col)
        cel.value = naam
        cel.alignment = center_align
        cel.border = thin_border
        cel.fill = lichtpaars_fill
        # Niet meetellen als dit een EXTRA overname is of als de pauze voor een PV zelf is
        if (not is_pauzevlinder(naam)) and (normalize_attr(attr) != 'extra'):
            pv_korte_pauze_count[pv["naam"]] += 1
        return True

    # Als anchor bestaat, probeer exact +10, dan groter, anders lagere alternatieven
    if anchor >= 0:
        # +10 en verder
        for d in range(10, len(pauze_cols)-anchor):
            if _try_place_at_col(pauze_cols[anchor + d]):
                return True
        # lager
        for d in range(9, 0, -1):
            idx = anchor + d
            if 0 <= idx < len(pauze_cols) and _try_place_at_col(pauze_cols[idx]):
                return True

    # Als geen anchor of niets gevonden: val terug op fairness, maar zonder links-bias
    # Kies uit alle geldige (pv_row, col) paren, sorteer op pv fairness en dan op kolomindex die het verst ligt van begin (rechts-bias)
    pairs = []
    for col in pauze_cols:
        if not is_korte_pauze_toegestaan_col(col):
            continue
        header = ws_pauze.cell(1, col).value
        uur = parse_header_uur(header)
        if uur not in werk_uren or uur in verboden_uren:
            continue
        attr = vind_attractie_op_uur(naam, uur)
        if not attr:
            continue
        for pv, pv_row in pv_rows:
            if is_pauzevlinder(naam) and pv["naam"] != naam:
                continue
            if ws_pauze.cell(pv_row, col).value not in [None, ""]:
                continue
            if not pv_kan_attr(pv, attr) and not is_student_extra(naam):
                continue
            pairs.append((pv, pv_row, col))
    if not pairs:
        return False
    pairs.sort(key=lambda x: (pv_korte_pauze_count[x[0]["naam"]], -pauze_cols.index(x[2])))
    pv, pv_row, col = pairs[0]
    attr = vind_attractie_op_uur(naam, parse_header_uur(ws_pauze.cell(1, col).value))
    ws_pauze.cell(pv_row-1, col).value = attr
    ws_pauze.cell(pv_row-1, col).alignment = center_align
    ws_pauze.cell(pv_row-1, col).border = thin_border
    cel = ws_pauze.cell(pv_row, col)
    cel.value = naam
    cel.alignment = center_align
    cel.border = thin_border
    cel.fill = lichtpaars_fill
    # Niet meetellen als dit een EXTRA overname is of als de pauze voor een PV zelf is
    if (not is_pauzevlinder(naam)) and (normalize_attr(attr) != 'extra'):
        pv_korte_pauze_count[pv["naam"]] += 1
    return True

# verzamel alle namen met een lange pauze en sorteer op laatste einde (desc)
names_with_long = []
alle_studenten_namen = {s["naam"] for s in studenten if student_totalen.get(s["naam"], 0) > 0}
for naam in alle_studenten_namen:
    if _has_long_pause(naam):
        end_idx = _last_long_pause_end_index(naam)
        names_with_long.append((end_idx, naam))
names_with_long.sort(reverse=True)  # laatste eerst

for _end, naam in names_with_long:
    _place_short_pause_for(naam)

# Zorg dat latere rondes deze personen overslaan: recompute korte_pauze_ontvangers nu
korte_pauze_ontvangers = set()
for pv, pv_row in pv_rows:
    for idx, col in enumerate(pauze_cols):
        cel = ws_pauze.cell(pv_row, col)
        if cel.value and str(cel.value).strip() != "":
            # korte pauze = enkel blok
            is_lange = False
            if idx+1 < len(pauze_cols):
                next_col = pauze_cols[idx+1]
                cel_next = ws_pauze.cell(pv_row, next_col)
                if cel_next.value == cel.value:
                    is_lange = True
            if idx > 0:
                prev_col = pauze_cols[idx-1]
                prev_cel = ws_pauze.cell(pv_row, prev_col)
                if prev_cel.value == cel.value:
                    is_lange = True
            if not is_lange:
                korte_pauze_ontvangers.add(str(cel.value).strip())

# Bepaal wie geen lange pauze heeft gekregen
studenten_zonder_lange_pauze = []
for s in studenten:
    naam = s["naam"]
    heeft_lange = False
    for pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value == naam:
                # Check of volgende cel ook deze naam heeft (dubbele blok)
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze.cell(pv_row, next_col)
                    if cel_next.value == naam:
                        heeft_lange = True
                        break
        if heeft_lange:
            break
    if not heeft_lange:
        studenten_zonder_lange_pauze.append(s)

# Eerst: korte pauze toewijzen aan studenten zonder lange pauze
def korte_pauze_toewijzen(studenten_lijst):
    for s in studenten_lijst:
        if s["naam"] in korte_pauze_ontvangers or _has_short_pause(s["naam"]):
            continue
        naam = s["naam"]
        werk_uren = get_student_work_hours(naam)
        if len(werk_uren) > 2:
            verboden_uren = {werk_uren[0], werk_uren[-1]}
        else:
            verboden_uren = set(werk_uren)
        pauze_cols_sorted = sorted(pauze_cols)
        geplaatst = False
        for uur in random.sample(werk_uren, len(werk_uren)):
            if uur in verboden_uren:
                continue
            attr = vind_attractie_op_uur(naam, uur)
            if not attr:
                continue
            geldige_slots = []
            for (pv, pv_row) in pv_rows:
                # Pauzevlinders: enkel op eigen rij
                if is_pauzevlinder(naam) and pv["naam"] != naam:
                    continue
                for col in pauze_cols:
                    col_header = ws_pauze.cell(1, col).value
                    col_uur = parse_header_uur(col_header)
                    if col_uur != uur:
                        continue
                    if not is_korte_pauze_toegestaan_col(col):
                        continue
                    if not pv_kan_attr(pv, attr) and not is_student_extra(naam):
                        continue
                    cel = ws_pauze.cell(pv_row, col)
                    if cel.value in [None, ""]:
                        geldige_slots.append((pv, pv_row, col))
            geldige_slots.sort(key=lambda slot: pv_korte_pauze_count[slot[0]["naam"]])
            for (pv, pv_row, col) in geldige_slots:
                boven_cel = ws_pauze.cell(pv_row - 1, col)
                # PV korte pauze: laat boven leeg
                boven_cel.value = None if is_pauzevlinder(naam) else attr
                boven_cel.alignment = center_align
                boven_cel.border = thin_border
                cel = ws_pauze.cell(pv_row, col)
                cel.value = naam
                cel.alignment = center_align
                cel.border = thin_border
                cel.fill = lichtpaars_fill
                korte_pauze_ontvangers.add(naam)
                # Niet meetellen als dit een EXTRA overname is of als de pauze voor een PV zelf is
                if (not is_pauzevlinder(naam)) and (normalize_attr(attr) != 'extra'):
                    pv_korte_pauze_count[pv["naam"]] += 1
                geplaatst = True
                break
            if geplaatst:
                break
        if not geplaatst:
            niet_geplaatste_korte_pauze.append(naam)

korte_pauze_toewijzen(studenten_zonder_lange_pauze)
# Daarna: de rest
korte_pauze_toewijzen([s for s in studenten if s not in studenten_zonder_lange_pauze])
korte_pauze_toewijzen([s for s in studenten if s not in studenten_zonder_lange_pauze])

# --- Iteratief wisselen: studenten zonder korte pauze proberen te ruilen met anderen (geen pauzevlinders) ---

def vind_korte_pauze_cell(naam):
    """Vind (pv_row, col) van de korte pauze van deze student, of None."""
    for pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value == naam:
                # Check of het een korte pauze is (enkel blok, niet dubbel)
                is_lange = False
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze.cell(pv_row, next_col)
                    if cel_next.value == naam:
                        is_lange = True
                if idx > 0:
                    prev_col = pauze_cols[idx-1]
                    prev_cel = ws_pauze.cell(pv_row, prev_col)
                    if prev_cel.value == naam:
                        is_lange = True
                if not is_lange:
                    return (pv_row, col)
    return None

def kan_student_korte_pauze_op_plek(naam, pv_row, col):
    """Check of student naam op deze plek een korte pauze mag hebben."""
    # Mag niet op pauzevlinder-rij
    if is_pauzevlinder(naam):
        return False
    # Moet werken op dit uur
    col_header = ws_pauze.cell(1, col).value
    col_uur = parse_header_uur(col_header)
    werk_uren = get_student_work_hours(naam)
    if col_uur not in werk_uren:
        return False
    # Niet in eerste/laatste werkuur
    if len(werk_uren) > 2:
        if col_uur == werk_uren[0] or col_uur == werk_uren[-1]:
            return False
    # Attractie moet kloppen
    attr = vind_attractie_op_uur(naam, col_uur)
    if not attr:
        return False
    # Pauzevlinder moet deze attractie kunnen
    pv = None
    for pv_obj, row in pv_rows:
        if row == pv_row:
            pv = pv_obj
            break
    if not pv:
        return False
    if not pv_kan_attr(pv, attr) and not is_student_extra(naam):
        return False
    # Kolom moet korte pauze toestaan
    if not is_korte_pauze_toegestaan_col(col):
        return False
    return True

# Verzamel actuele lijst van studenten zonder korte pauze
werkende_studenten = [s for s in studenten if student_totalen.get(s["naam"], 0) > 0 and not is_pauzevlinder(s["naam"])]
studenten_zonder_korte_pauze = []
for s in werkende_studenten:
    naam = s["naam"]
    heeft_korte = False
    for pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value == naam:
                # Check of GEEN dubbele blok (dus geen lange pauze)
                is_lange = False
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze.cell(pv_row, next_col)
                    if cel_next.value == naam:
                        is_lange = True
                if idx > 0:
                    prev_col = pauze_cols[idx-1]
                    prev_cel = ws_pauze.cell(pv_row, prev_col)
                    if prev_cel.value == naam:
                        is_lange = True
                if not is_lange:
                    heeft_korte = True
                    break
        if heeft_korte:
            break
    if not heeft_korte:
        studenten_zonder_korte_pauze.append(naam)

max_wissel_passes = 10
for _ in range(max_wissel_passes):
    if not studenten_zonder_korte_pauze:
        break
    verbeterd = False
    for naam_zonder in studenten_zonder_korte_pauze:
        # Probeer te ruilen met een student die wél een korte pauze heeft (geen pauzevlinder)
        for s in werkende_studenten:
            naam_met = s["naam"]
            if naam_met == naam_zonder:
                continue
            if naam_met in studenten_zonder_korte_pauze:
                continue
            # Vind de korte pauze van deze student
            plek = vind_korte_pauze_cell(naam_met)
            if not plek:
                continue
            pv_row, col = plek
            # Mag naam_zonder op deze plek een korte pauze hebben?
            if not kan_student_korte_pauze_op_plek(naam_zonder, pv_row, col):
                continue
            # Bepaal attractie voor naam_zonder op deze plek
            col_header = ws_pauze.cell(1, col).value
            col_uur = parse_header_uur(col_header)
            attr_zonder = vind_attractie_op_uur(naam_zonder, col_uur)
            if not attr_zonder:
                continue
            # Mag naam_met elders een korte pauze krijgen?
            # Zoek alternatieve plek voor naam_met
            gevonden = False
            for pv2, pv_row2 in pv_rows:
                if is_pauzevlinder(naam_met):
                    continue
                for col2 in pauze_cols:
                    if (pv_row2, col2) == (pv_row, col):
                        continue
                    cel2 = ws_pauze.cell(pv_row2, col2)
                    if cel2.value not in [None, ""]:
                        continue
                    if not kan_student_korte_pauze_op_plek(naam_met, pv_row2, col2):
                        continue
                    # Bepaal attractie voor naam_met op nieuwe plek
                    col2_header = ws_pauze.cell(1, col2).value
                    col2_uur = parse_header_uur(col2_header)
                    attr_met = vind_attractie_op_uur(naam_met, col2_uur)
                    if not attr_met:
                        continue
                    # Wissel uitvoeren
                    # 1. naam_met uit oude plek halen
                    ws_pauze.cell(pv_row, col).value = None
                    ws_pauze.cell(pv_row, col).fill = naam_leeg_fill
                    ws_pauze.cell(pv_row-1, col).value = None
                    # 2. naam_zonder op deze plek zetten
                    ws_pauze.cell(pv_row, col).value = naam_zonder
                    ws_pauze.cell(pv_row, col).fill = lichtpaars_fill
                    ws_pauze.cell(pv_row, col).alignment = center_align
                    ws_pauze.cell(pv_row, col).border = thin_border
                    ws_pauze.cell(pv_row-1, col).value = attr_zonder
                    ws_pauze.cell(pv_row-1, col).alignment = center_align
                    ws_pauze.cell(pv_row-1, col).border = thin_border
                    # 3. naam_met op nieuwe plek zetten
                    ws_pauze.cell(pv_row2, col2).value = naam_met
                    ws_pauze.cell(pv_row2, col2).fill = lichtpaars_fill
                    ws_pauze.cell(pv_row2, col2).alignment = center_align
                    ws_pauze.cell(pv_row2, col2).border = thin_border
                    ws_pauze.cell(pv_row2-1, col2).value = attr_met
                    ws_pauze.cell(pv_row2-1, col2).alignment = center_align
                    ws_pauze.cell(pv_row2-1, col2).border = thin_border
                    verbeterd = True
                    gevonden = True
                    break
                if gevonden:
                    break
            if verbeterd:
                break
        if verbeterd:
            break
    # Update lijst van studenten zonder korte pauze
    studenten_zonder_korte_pauze = []
    for s in werkende_studenten:
        naam = s["naam"]
        heeft_korte = False
        for pv, pv_row in pv_rows:
            for idx, col in enumerate(pauze_cols):
                cel = ws_pauze.cell(pv_row, col)
                if cel.value == naam:
                    # Check of GEEN dubbele blok (dus geen lange pauze)
                    is_lange = False
                    if idx+1 < len(pauze_cols):
                        next_col = pauze_cols[idx+1]
                        cel_next = ws_pauze.cell(pv_row, next_col)
                        if cel_next.value == naam:
                            is_lange = True
                    if idx > 0:
                        prev_col = pauze_cols[idx-1]
                        prev_cel = ws_pauze.cell(pv_row, prev_col)
                        if prev_cel.value == naam:
                            is_lange = True
                    if not is_lange:
                        heeft_korte = True
                        break
            if heeft_korte:
                break
        if not heeft_korte:
            studenten_zonder_korte_pauze.append(naam)
    if not verbeterd:
        break  # geen verbetering meer mogelijk

# Iteratieve optimalisatie: verschuif korte pauzes van "rijke" naar "arme" pauzevlinders
max_opt_passes = 10
for _ in range(max_opt_passes):
    # Zoek max en min aantal korte pauzes
    if not pv_korte_pauze_count:
        break
    max_pv = max(pv_korte_pauze_count, key=lambda k: pv_korte_pauze_count[k])
    min_pv = min(pv_korte_pauze_count, key=lambda k: pv_korte_pauze_count[k])
    if pv_korte_pauze_count[max_pv] - pv_korte_pauze_count[min_pv] <= 1:
        break  # verdeling is al redelijk
    # Zoek een korte pauze van max_pv die overgezet kan worden naar min_pv
    found = False
    for col in pauze_cols:
        pv_row_max = next((row for pv, row in pv_rows if pv["naam"] == max_pv), None)
        pv_row_min = next((row for pv, row in pv_rows if pv["naam"] == min_pv), None)
        if pv_row_max is None or pv_row_min is None:
            continue
        cel_max = ws_pauze.cell(pv_row_max, col)
        naam = cel_max.value
        if not naam or str(naam).strip() == "":
            continue
        # Check of het een korte pauze is (enkel blok, niet dubbel)
        idx = pauze_cols.index(col)
        is_lange = False
        if idx+1 < len(pauze_cols):
            next_col = pauze_cols[idx+1]
            cel_next = ws_pauze.cell(pv_row_max, next_col)
            if cel_next.value == cel_max.value:
                is_lange = True
        if idx > 0:
            prev_col = pauze_cols[idx-1]
            prev_cel = ws_pauze.cell(pv_row_max, prev_col)
            if prev_cel.value == cel_max.value:
                is_lange = True
        if is_lange:
            continue
        # Mag min_pv deze attractie overnemen?
        attr = ws_pauze.cell(pv_row_max-1, col).value
        if not pv_kan_attr(next(pv for pv, _ in pv_rows if pv["naam"] == min_pv), attr):
            continue
        # Is de cel bij min_pv vrij?
        cel_min = ws_pauze.cell(pv_row_min, col)
        if cel_min.value not in [None, ""]:
            continue
        # Wissel de korte pauze van max_pv naar min_pv
        cel_min.value = naam
        cel_min.alignment = center_align
        cel_min.border = thin_border
        cel_min.fill = lichtpaars_fill
        ws_pauze.cell(pv_row_min-1, col).value = attr
        ws_pauze.cell(pv_row_min-1, col).alignment = center_align
        ws_pauze.cell(pv_row_min-1, col).border = thin_border
        cel_max.value = None
        ws_pauze.cell(pv_row_max-1, col).value = None
        # Pas telling aan enkel als dit geen EXTRA-overname is
        if attr and normalize_attr(attr) != 'extra':
            pv_korte_pauze_count[max_pv] -= 1
            pv_korte_pauze_count[min_pv] += 1
        found = True
        break
    if not found:
        break  # geen verschuiving meer mogelijk



# --- Iteratieve optimalisatie: verdeel lange pauzes zo eerlijk mogelijk over pauzevlinders ---

max_opt_passes_lange = 10
from collections import Counter
for _ in range(max_opt_passes_lange):
    pass  # (oude optimalisatie-code is verwijderd, want niet meer nodig)

# --- Pauzevlinders met >6u: altijd lange pauze in eigen rij ---
import random
# --- Pauzevlinders met >6u: altijd lange pauze in eigen rij, gespreid over eerste drie pauzeuren ---
for pv, pv_row in pv_rows:
    naam = pv["naam"]
    werk_uren = get_student_work_hours(naam)
    if len(werk_uren) > 6:
        # Alleen de eerste 11 kwartieren (indices 0 t/m 10) zijn toegestaan voor lange pauzes
        if heeft_al_lange_pauze(naam):
            continue
        halve_uren = []  # lijst van (idx, col1, col2)
        max_start_idx = min(10, len(pauze_cols)-2)  # idx 0 t/m 10 zijn halve uren binnen eerste 11 kwartieren
        for idx in range(max_start_idx+1):
            col1 = pauze_cols[idx]
            col2 = pauze_cols[idx+1]
            col1_header = ws_pauze.cell(1, col1).value
            # Alleen starten op heel of half uur
            try:
                min1 = int(str(col1_header).split('u')[1]) if 'u' in str(col1_header) and len(str(col1_header).split('u')) > 1 else 0
            except:
                min1 = 0
            if min1 not in (0, 30):
                continue
            cel1 = ws_pauze.cell(pv_row, col1)
            cel2 = ws_pauze.cell(pv_row, col2)
            if cel1.value in [None, ""] and cel2.value in [None, ""]:
                halve_uren.append((idx, col1, col2))
        # Shuffle de halve uren
        random.shuffle(halve_uren)
        # Probeer in geshuffelde volgorde een lange pauze te plaatsen
        geplaatst = False
        for idx, col1, col2 in halve_uren:
            cel1 = ws_pauze.cell(pv_row, col1)
            cel2 = ws_pauze.cell(pv_row, col2)
            if cel1.value in [None, ""] and cel2.value in [None, ""] and not heeft_al_lange_pauze(naam):
                cel1.value = naam
                cel2.value = naam
                cel1.alignment = center_align
                cel2.alignment = center_align
                cel1.border = thin_border
                cel2.border = thin_border
                cel1.fill = lichtgroen_fill
                cel2.fill = lichtgroen_fill
                geplaatst = True
                break
        # Indien geen plek gevonden, doe niets (komt zelden voor)



output = BytesIO()

# -----------------------------
# ENFORCE: Korte pauzes minimaal 10 blokken uit elkaar (maximaliseer anders)
# -----------------------------

# We beschouwen alle pauzecellen (korte én beide helften van lange pauzes).
# We verplaatsen alleen korte pauzes (enkelblok) naar lege geschikte slots.

def _get_student_pause_cols(naam):
    cols = []
    for _pv, pv_row in pv_rows:
        for col in pauze_cols:
            if ws_pauze.cell(pv_row, col).value == naam:
                cols.append(col)
    return sorted(cols)

def _get_student_short_pause_positions(naam):
    pos = []  # lijst van (pv_row, col)
    for _pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value != naam:
                continue
            # controleer of GEEN deel van een dubbele blok (lange pauze)
            is_dubbel = False
            if idx+1 < len(pauze_cols):
                if ws_pauze.cell(pv_row, pauze_cols[idx+1]).value == naam:
                    is_dubbel = True
            if idx > 0:
                if ws_pauze.cell(pv_row, pauze_cols[idx-1]).value == naam:
                    is_dubbel = True
            if not is_dubbel:
                pos.append((pv_row, col))
    return pos

def _min_gap(cols):
    if len(cols) < 2:
        return 10**9
    cols = sorted(cols)
    mg = min(cols[i+1]-cols[i] for i in range(len(cols)-1))
    return mg

def _can_place_short_pause(naam, pv_row, col):
    # cel moet leeg zijn
    if ws_pauze.cell(pv_row, col).value not in [None, ""]:
        return False
    # kolom moet korte pauze toelaten
    if not is_korte_pauze_toegestaan_col(col):
        return False
    # student moet werken op dit uur en niet in eerste/laatste werkuur
    header = ws_pauze.cell(1, col).value
    uur = parse_header_uur(header)
    if uur is None:
        return False
    werk_uren = get_student_work_hours(naam)
    if uur not in werk_uren:
        return False
    if len(werk_uren) > 2:
        if uur == werk_uren[0] or uur == werk_uren[-1]:
            return False
    else:
        # bij 1-2 uren: geen korte pauze plannen
        return False
    # pauzevlinder-capability: pauzevlinder op pv_row moet attractie kunnen
    attr = vind_attractie_op_uur(naam, uur)
    if not attr:
        return False
    pv_obj = None
    for pv, row in pv_rows:
        if row == pv_row:
            pv_obj = pv
            break
    if pv_obj is None:
        return False
    if not pv_kan_attr(pv_obj, attr) and not is_student_extra(naam):
        return False
    return True

def _move_short_pause(naam, from_row, from_col, to_row, to_col):
    # leegmaken bron
    ws_pauze.cell(from_row, from_col).value = None
    ws_pauze.cell(from_row-1, from_col).value = None
    # invullen doel
    header = ws_pauze.cell(1, to_col).value
    uur = parse_header_uur(header)
    attr = vind_attractie_op_uur(naam, uur)
    # Voor PV korte pauze: laat boven leeg
    ws_pauze.cell(to_row-1, to_col).value = None if is_pauzevlinder(naam) else attr
    ws_pauze.cell(to_row-1, to_col).alignment = center_align
    ws_pauze.cell(to_row-1, to_col).border = thin_border
    ws_pauze.cell(to_row, to_col).value = naam
    ws_pauze.cell(to_row, to_col).alignment = center_align
    ws_pauze.cell(to_row, to_col).border = thin_border

def _recolor_pauze_sheet():
    # Kleur korte pauze paars, lange (dubbel) groen, leeg lichtblauw
    for _pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            val = cel.value
            if val in [None, ""]:
                cel.fill = naam_leeg_fill
                continue
            # lange pauze als naastliggende cel dezelfde naam heeft
            is_dubbel = False
            if idx+1 < len(pauze_cols):
                c2 = ws_pauze.cell(pv_row, pauze_cols[idx+1])
                if c2.value == val:
                    cel.fill = lichtgroen_fill
                    c2.fill = lichtgroen_fill
                    is_dubbel = True
                    continue
            if idx > 0:
                c1 = ws_pauze.cell(pv_row, pauze_cols[idx-1])
                if c1.value == val:
                    # onderdeel van reeds gekleurde lange pauze
                    continue
            # anders korte pauze
            cel.fill = lichtpaars_fill

def _enforce_min_gap_for_short_pauses(desired_gap=10, max_passes=5):
    changed = False
    for _ in range(max_passes):
        improved = False
        # itereren over alle studenten met minstens 1 korte pauze
        alle_namen = {s["naam"] for s in studenten if student_totalen.get(s["naam"], 0) > 0}
        for naam in alle_namen:
            short_pos = _get_student_short_pause_positions(naam)
            if not short_pos:
                continue
            all_cols = _get_student_pause_cols(naam)
            # bekijk elke korte pauze afzonderlijk
            for (from_row, from_col) in short_pos:
                # huidige minimale gap voor deze korte pauze
                other_cols = [c for c in all_cols if c != from_col]
                cur_gap = min((abs(from_col - c) for c in other_cols), default=10**9)
                if cur_gap >= desired_gap:
                    continue  # al goed
                # zoek beste lege slot
                best = None  # (best_gap, to_row, to_col)
                for _pv, pv_row in pv_rows:
                    # Pauzevlinders: korte pauze enkel op eigen rij verplaatsen/plaatsen
                    if is_pauzevlinder(naam) and _pv["naam"] != naam:
                        continue
                    for col in pauze_cols:
                        if not _can_place_short_pause(naam, pv_row, col):
                            continue
                        # gap als we hierheen verplaatsen
                        new_gap = min((abs(col - c) for c in other_cols), default=10**9)
                        if (best is None) or (new_gap > best[0]) or (new_gap == best[0] and col < best[2]):
                            best = (new_gap, pv_row, col)
                            # early exit als we desired halen
                            if new_gap >= desired_gap:
                                break
                    if best and best[0] >= desired_gap:
                        break
                if best and best[0] > cur_gap:
                    _move_short_pause(naam, from_row, from_col, best[1], best[2])
                    improved = True
                    changed = True
                    # update caches voor volgende iteraties
                    all_cols = _get_student_pause_cols(naam)
        if not improved:
            break
    # herkleuren na eventuele wijzigingen
    if changed:
        _recolor_pauze_sheet()
    return changed

# Voer de enforce stap uit met doelafstand 10 blokken
_enforce_min_gap_for_short_pauses(desired_gap=10, max_passes=6)

# Optionele samenvatting in Streamlit
# Debug samenvatting (globale minimale pauze-afstand) verwijderd om UI schoon te houden.
# Indien opnieuw nodig: functie _global_min_gap_summary() herstellen.

# --- FEEDBACK SHEET ---
ws_feedback = wb_out.create_sheet("Feedback")
row_fb = 1

# 1. Lange werkers (>6u) zonder lange pauze
lange_werkers_zonder_lange_pauze = set()

def _heeft_lange_pauze_naam(naam: str) -> bool:
    """Zoek in ws_pauze of deze persoon een dubbele blok (lange pauze) heeft."""
    for pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value == naam:
                # Check of volgende cel ook deze naam heeft (dubbele blok)
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze.cell(pv_row, next_col)
                    if cel_next.value == naam:
                        return True
    return False

# a) reguliere lange werkers
for s in lange_werkers:
    naam = s["naam"]
    if not _heeft_lange_pauze_naam(naam):
        lange_werkers_zonder_lange_pauze.add(naam)

# b) pauzevlinders die >6u werken meenemen
for pv, _pv_row in pv_rows:
    naam = pv["naam"]
    werk_uren = get_student_work_hours(naam) or []
    if len(werk_uren) > 6:
        if not _heeft_lange_pauze_naam(naam):
            lange_werkers_zonder_lange_pauze.add(naam)

ws_feedback.cell(row_fb, 1, "Lange werkers (>6u) zonder lange pauze:")
row_fb += 1
if lange_werkers_zonder_lange_pauze:
    for naam in sorted(lange_werkers_zonder_lange_pauze):
        ws_feedback.cell(row_fb, 1, naam)
        row_fb += 1
else:
    vinkje_cel = ws_feedback.cell(row_fb, 1, "✓")
    ws_feedback.cell(row_fb, 2, "Iedereen heeft een lange pauze gekregen.")
    from openpyxl.styles import PatternFill, Font
    vinkje_cel.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")  # opvallend groen
    vinkje_cel.font = Font(bold=True, color="006100")  # donkergroen
    row_fb += 1

# 2. Werkende studenten zonder korte pauze
werkende_studenten = [s for s in studenten if student_totalen.get(s["naam"], 0) > 0]
studenten_zonder_korte_pauze = []
for s in werkende_studenten:
    naam = s["naam"]
    # Zoek in ws_pauze of deze student een korte pauze heeft (enkel blok, niet dubbel)
    heeft_korte = False
    for pv, pv_row in pv_rows:
        for idx, col in enumerate(pauze_cols):
            cel = ws_pauze.cell(pv_row, col)
            if cel.value == naam:
                # Check of GEEN dubbele blok (dus geen lange pauze)
                is_lange = False
                if idx+1 < len(pauze_cols):
                    next_col = pauze_cols[idx+1]
                    cel_next = ws_pauze.cell(pv_row, next_col)
                    if cel_next.value == naam:
                        is_lange = True
                if idx > 0:
                    prev_col = pauze_cols[idx-1]
                    prev_cel = ws_pauze.cell(pv_row, prev_col)
                    if prev_cel.value == naam:
                        is_lange = True
                if not is_lange:
                    heeft_korte = True
                    break
        if heeft_korte:
            break
    if not heeft_korte:
        studenten_zonder_korte_pauze.append(naam)

ws_feedback.cell(row_fb, 1, "Werkende studenten zonder korte pauze:")
row_fb += 1
if studenten_zonder_korte_pauze:
    for naam in studenten_zonder_korte_pauze:
        ws_feedback.cell(row_fb, 1, naam)
        row_fb += 1
else:
    vinkje_cel = ws_feedback.cell(row_fb, 1, "✓")
    ws_feedback.cell(row_fb, 2, "Iedereen heeft een korte pauze gekregen.")
    vinkje_cel.fill = PatternFill(start_color="92D050", end_color="92D050", fill_type="solid")
    vinkje_cel.font = Font(bold=True, color="006100")
    row_fb += 1

wb_out.save(output)
output.seek(0)  # Zorg dat lezen vanaf begin kan




#ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo


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
