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
    st.warning("Upload eerst een Excel-bestand om verder te gaan.")
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
        return [h for h in open_uren if 12 <= h <= 17]
    elif 10 in open_uren and 17 in open_uren:
        return [h for h in open_uren if 12 <= h <= 17]
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
    """Check capaciteit in alle uren en plaats dan in één keer."""
    # Capaciteit check
    for h in block_hours:
        if not _has_capacity(attr, h):
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



# =============================
# FIXED TWEEDE PASS: gaten vullen + swaps
# =============================

# Mapping van naam → student dict
naam2student = {s["naam"]: s for s in studenten_workend}


def _has_capacity(attr, uur):
    """
    Controleer of er nog plek is op deze attractie en dit uur.
    Initialiseer per_hour_assigned_counts[uur][attr] als het nog niet bestaat.
    """
    if uur not in per_hour_assigned_counts:
        per_hour_assigned_counts[uur] = {}
    if attr not in per_hour_assigned_counts[uur]:
        per_hour_assigned_counts[uur][attr] = 0
    return per_hour_assigned_counts[uur][attr] < _max_spots_for(attr, uur)

def _fill_gaps_singletons():
    """
    Greedy: vul resterende lege plekken per uur met studenten uit extra,
    met blokjes van 1 uur.
    """
    for uur in sorted(open_uren):
        # Kopie, want we muteren lijst
        for naam in list(extra_assignments.get(uur, [])):
            s = naam2student.get(naam)
            if not s:
                continue
            candidate_attrs = [a for a in attracties_te_plannen if a in s["attracties"]]
            candidate_attrs.sort(key=lambda a: sum(1 for st in studenten_workend if a in st["attracties"]))
            placed = False
            for attr in candidate_attrs:
                if _has_capacity(attr, uur):
                    assigned_map.setdefault((uur, attr), []).append(naam)
                    per_hour_assigned_counts[uur][attr] += 1
                    s["assigned_hours"].append(uur)
                    s["assigned_attracties"].add(attr)
                    extra_assignments[uur].remove(naam)
                    placed = True
                    break
            # Als niet geplaatst, blijft in extra_assignments

def _attempt_simple_swaps():
    """
    Simpele 1-stap swap binnen hetzelfde uur:
    - Voor een student in extra op uur U die attr A kan (maar A zit vol),
      probeer een huidige ingeroosterde student te verplaatsen naar een andere attr op U
      waar nog plek is en die hij ook kan.
    """
    for uur in sorted(open_uren):
        for naam in list(extra_assignments.get(uur, [])):
            s = naam2student.get(naam)
            if not s:
                continue
            wanted_attrs = [a for a in attracties_te_plannen if a in s["attracties"]]

            swap_done = False
            for attr in wanted_attrs:
                # Alleen interessant als attr vol zit
                if not _has_capacity(attr, uur):
                    incumbents = list(assigned_map.get((uur, attr), []))
                    for inc_naam in incumbents:
                        inc = naam2student.get(inc_naam)
                        if not inc:
                            continue
                        # Zoek alternatieve attr voor incumbent met ruimte
                        for alt_attr in inc["attracties"]:
                            if alt_attr == attr:
                                continue
                            if _has_capacity(alt_attr, uur):
                                # Verplaats incumbent
                                assigned_map[(uur, attr)].remove(inc_naam)
                                per_hour_assigned_counts[uur][attr] -= 1

                                assigned_map.setdefault((uur, alt_attr), []).append(inc_naam)
                                per_hour_assigned_counts[uur][alt_attr] += 1
                                inc["assigned_attracties"].add(alt_attr)

                                # Plaats extra-student op vrijgekomen plek
                                assigned_map.setdefault((uur, attr), []).append(naam)
                                per_hour_assigned_counts[uur][attr] += 1
                                s["assigned_hours"].append(uur)
                                s["assigned_attracties"].add(attr)
                                extra_assignments[uur].remove(naam)

                                swap_done = True
                                break
                        if swap_done:
                            break
                if swap_done:
                    break

# =============================
# Uitvoeren van de resolutiefase
# =============================
_fill_gaps_singletons()
_attempt_simple_swaps()
# Nog een keer greedy vullen, soms ontstaan er nieuwe gaten na swaps
_fill_gaps_singletons()


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
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)

# Header
ws_out.cell(1, 1, vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1, col_idx, f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1, col_idx).fill = header_fill
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
        ws_out.cell(rij_out, 1).fill = attr_fill
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

output = BytesIO()
wb_out.save(output)
output.seek(0)
st.success("Planning gegenereerd!")
st.download_button(
    "Download planning",
    data=output.getvalue(),
    file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
)


