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
# Upload / lees Excel
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
if not uploaded_file:
    st.warning("Upload eerst een Excel-bestand om verder te gaan.")
    st.stop()

wb = load_workbook(BytesIO(uploaded_file.read()))
ws = wb["Blad1"]

# -----------------------------
# Hulpfuncties (hergebruikt)
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
    """
    DP partition voor lengte L. ### AANPASSING:
    - blokgroottes preferentie: 3, 2, 4, dan 1 (1 is laatste redmiddel)
    - objective: minimaliseer eerst aantal 1-blokjes, dan maximaliseer aantal (3 of 2)-blokken,
      daarna minimaliseer aantal blokken.
    Retourneert lijst van blokgroottes die optellen tot L.
    """
    blocks = [3,2,4,1]  # voorkeurvolgorde zoals gevraagd
    # dp[i] = tuple (num_ones, -num_3_or_2_blocks, num_blocks, partition_list)
    big = 10**9
    dp = [(big, big, big, [])] * (L+1)
    dp[0] = (0, 0, 0, [])
    for i in range(1, L+1):
        best = (big, big, big, [])
        for b in blocks:
            if i-b < 0:
                continue
            prev = dp[i-b]
            num_ones = prev[0] + (1 if b == 1 else 0)
            num_32 = prev[1] + ( -1 if b in (2,3) else 0 )  # we store negative so that smaller is better
            num_blocks = prev[2] + 1
            part = prev[3] + [b]
            cand = (num_ones, num_32, num_blocks, part)
            # lexicografische vergelijking; we willen min num_ones, min num_32 (dat is negatie: dus min -> max 3/2),
            # min num_blocks
            if (cand[0], cand[1], cand[2]) < (best[0], best[1], best[2]):
                best = cand
        dp[i] = best
    return dp[L][3]

def contiguous_runs(sorted_hours):
    runs = []
    if not sorted_hours:
        return runs
    run = [sorted_hours[0]]
    for h in sorted_hours[1:]:
        if h == run[-1] + 1:
            run.append(h)
        else:
            runs.append(run)
            run = [h]
    runs.append(run)
    return runs

# -----------------------------
# Studenten inlezen (zoals vroeger, maar we bewaren aantal_attracties uit AG)
# -----------------------------
studenten = []
for rij in range(2,500):
    naam = ws.cell(rij,12).value
    if not naam:
        continue
    uren_beschikbaar = [10+(kol-3) for kol in range(3,12) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    attracties = [ws.cell(1,kol).value for kol in range(14,32) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    try:
        aantal_attracties = int(ws[f'AG{rij}'].value) if ws[f'AG{rij}'].value else len(attracties)
    except:
        aantal_attracties = len(attracties)
    studenten.append({
        "naam": naam,
        "uren_beschikbaar": sorted(uren_beschikbaar),
        "attracties": [a for a in attracties if a],
        "aantal_attracties": aantal_attracties,
        "is_pauzevlinder": False,
        "pv_number": None,
        "assigned_attracties": set(),
        "assigned_hours": []
    })

# -----------------------------
# Openingsuren bepalen (zoals voorheen)
# -----------------------------
open_uren = [int(str(ws.cell(1,kol).value).replace("u","").strip()) for kol in range(36,45) if ws.cell(2,kol).value in [1,True,"WAAR","X"]]
if not open_uren:
    open_uren = list(range(10,19))
open_uren = sorted(set(open_uren))

# -----------------------------
# Pauzevlinders lezen en hun bezetting bepalen (volgens jouw regels)
# -----------------------------
pauzevlinder_namen = [ws[f'BN{rij}'].value for rij in range(4,11) if ws[f'BN{rij}'].value]

def compute_pauze_hours(open_uren):
    mn = min(open_uren)
    mx = max(open_uren)
    if 10 in open_uren and 18 in open_uren:
        req = [h for h in open_uren if 12 <= h <= 17]
    elif 12 in open_uren and 18 in open_uren:
        req = [h for h in open_uren if 13 <= h <= 18]
    elif min(open_uren) >= 14:
        req = list(open_uren)
    else:
        req = [h for h in open_uren if 12 <= h <= 17]
    return sorted(req)

required_pauze_hours = compute_pauze_hours(open_uren)

for idx, pvnaam in enumerate(pauzevlinder_namen, start=1):
    if not pvnaam:
        continue
    for s in studenten:
        if s["naam"] == pvnaam:
            s["is_pauzevlinder"] = True
            s["pv_number"] = idx
            s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_pauze_hours]
            break

# -----------------------------
# Attracties en aantallen inlezen (en prioriteren)
# -----------------------------
aantallen = {}
attracties_te_plannen = []
for kol in range(47,65):
    naam = ws.cell(1,kol).value
    if naam:
        try:
            aantal = int(ws.cell(2,kol).value)
        except:
            aantal = 0
        aantallen[naam] = max(0, min(2, aantal))
        if aantallen[naam] >= 1:
            attracties_te_plannen.append(naam)

def kritieke_score(attr, studenten_list):
    return sum(1 for s in studenten_list if attr in s["attracties"])

studenten_workend = [s for s in studenten if any(u in open_uren for u in s["uren_beschikbaar"])]

# Sorteer attracties op basis van schaarste: weinigste kandidaten eerst
attracties_te_plannen.sort(key=lambda a: kritieke_score(a, studenten_workend))

# -----------------------------
# BA kolom: tweede-posities prioriteit (BA rijen 5..12)
# -----------------------------
ba_prioriteit = [ws[f'BA{r}'].value for r in range(5,13) if ws[f'BA{r}'].value]

# -----------------------------
# Per-uur: beschikbare studenten (lijst) EN per-uur beschikbare plekken (vanuit excel of berekend)
# -----------------------------
col_hour_map = {}
for kol in range(36,45):
    val_flag = ws.cell(2,kol).value
    if val_flag in [1,True,"WAAR","X"]:
        hdr = ws.cell(1,kol).value
        try:
            hour = int(str(hdr).replace("u","").strip())
        except:
            continue
        col_hour_map[kol] = hour

hour_count_from_sheet = {}
for kol,hr in col_hour_map.items():
    v = ws.cell(3,kol).value
    if isinstance(v, (int, float)):
        hour_count_from_sheet[hr] = int(v)

per_hour_students_list = {}
for uur in open_uren:
    names = [s for s in studenten_workend if uur in s["uren_beschikbaar"]]
    per_hour_students_list[uur] = sorted(names, key=lambda s: s["aantal_attracties"])

per_hour_total_workers = {}
for uur in open_uren:
    if uur in hour_count_from_sheet:
        totaal = hour_count_from_sheet[uur]
    else:
        totaal = len(per_hour_students_list[uur])
    num_pv = sum(1 for nv in pauzevlinder_namen if nv and uur in required_pauze_hours
                 and any(s["naam"]==nv for s in studenten))
    beschikbaar = max(0, totaal - num_pv)
    per_hour_total_workers[uur] = beschikbaar

# -----------------------------
# Per-uur: positions per attractie (aanpassen: haal second posities weg als er te weinig mensen zijn)
# -----------------------------
per_hour_positions = {}
for uur in open_uren:
    positions = {attr: aantallen.get(attr, 0) for attr in attracties_te_plannen}
    totaal_posities = sum(positions.values())
    beschikbaar = per_hour_total_workers.get(uur, 0)
    if totaal_posities <= beschikbaar:
        per_hour_positions[uur] = positions
        continue

    for attr_to_remove in reversed(ba_prioriteit):
        if totaal_posities <= beschikbaar:
            break
        if attr_to_remove in positions and positions[attr_to_remove] >= 2:
            positions[attr_to_remove] = 1
            totaal_posities -= 1

    if totaal_posities > beschikbaar:
        fallback_order = sorted(attracties_te_plannen, key=lambda a: kritieke_score(a, studenten_workend), reverse=True)
        for a in fallback_order:
            if totaal_posities <= beschikbaar:
                break
            if positions.get(a,0) >= 2:
                positions[a] = 1
                totaal_posities -=1

    per_hour_positions[uur] = positions

# -----------------------------
# Maak per-uur attractielijst (geordend op schaarste)
# -----------------------------
per_hour_attractielijsten = {}
for uur in open_uren:
    sorted_attrs = sorted(attracties_te_plannen, key=lambda a: kritieke_score(a, studenten_workend))
    lst = []
    for a in sorted_attrs:
        npos = per_hour_positions[uur].get(a, 0)
        for p in range(npos):
            lst.append(a)
    per_hour_attractielijsten[uur] = lst

# -----------------------------
# ASSIGNMENT: per student (min attracties eerst), split shifts in blokken en toewijzen
# -----------------------------
assigned_map = defaultdict(list)
per_hour_assigned_counts = {uur: {a: 0 for a in attracties_te_plannen} for uur in open_uren}

MAX_CONSEC = 4
MAX_PER_STUDENT_ATTR = 6

# Sorteer studenten op aantal attracties (min first) ### (start met de minst-capable student)
studenten_sorted = sorted(studenten_workend, key=lambda s: s["aantal_attracties"])

def current_student_hours(s):
    return sorted(s["assigned_hours"])

# ### AANPASSING in assignment loop:
# - we gebruiken partition_run_lengths met nieuwe voorkeur (3,2,4,1)
# - bij het zoeken naar een attractie voor een block gebruiken we de per_hour_attractielijsten
#   van het eerste uur van het block (dus we vullen 'naast' eerste attracties)
# - extra constraint: als er al een blok aanwezig is op die attractie voor één van de uren in block
#   (dus assigned_map[(h,attr)] niet leeg), dan skip attr (geen tweede blok op dezelfde attractie)
for s in studenten_sorted:
    uren = [u for u in s["uren_beschikbaar"] if u in open_uren and u not in s["assigned_hours"]]
    if not uren:
        continue
    uren = sorted(uren)
    runs = contiguous_runs(uren)
    for run in runs:
        L = len(run)
        if L == 0:
            continue
        block_sizes = partition_run_lengths(L)
        start_idx = 0
        for b in block_sizes:
            block_hours = run[start_idx:start_idx+b]
            start_idx += b
            assigned = False

            # we pick attracties in order van de attractielijst van het eerste uur (zo vullen we "eerste attracties")
            uur0 = block_hours[0]
            candidate_attrs = per_hour_attractielijsten.get(uur0, [])
            # remove duplicates but keep order
            seen = set()
            cand_ordered = []
            for a in candidate_attrs:
                if a not in seen:
                    seen.add(a)
                    cand_ordered.append(a)

            for attr in cand_ordered:
                if attr not in s["attracties"]:
                    continue
                if attr in s["assigned_attracties"]:
                    continue  # student al eens op die attractie gezet, skip zoals gevraagd

                # check: geen tweede block op zelfde attractie als er al iets staat in een van de uren
                any_block_existing = any(bool(assigned_map.get((h,attr))) for h in block_hours)
                if any_block_existing:
                    # volgens jouw regel: als er al een blok staat, kan er niet nog een blok komen bij dezelfde attractie
                    continue

                # check max per student per attr
                already_on_attr = sum(1 for h in s["assigned_hours"] if (s["naam"] in assigned_map.get((h,attr), [])))
                if already_on_attr + len(block_hours) > MAX_PER_STUDENT_ATTR:
                    continue

                # check capacity for all uren in block
                ruimte = True
                for h in block_hours:
                    allowed = per_hour_positions.get(h, {}).get(attr, 0)
                    used = per_hour_assigned_counts[h].get(attr, 0)
                    if used >= allowed:
                        ruimte = False
                        break
                if not ruimte:
                    continue

                # check max consecutive if we add block
                hypothetische = sorted(set(current_student_hours(s) + block_hours))
                if max_consecutive_hours(hypothetische) > MAX_CONSEC:
                    continue

                # assign block
                for h in block_hours:
                    assigned_map[(h, attr)].append(s["naam"])
                    per_hour_assigned_counts[h][attr] += 1
                    s["assigned_hours"].append(h)
                s["assigned_attracties"].add(attr)
                assigned = True
                break

            if not assigned:
                # kon geen passende attractie vinden -> laat onvervuld (later kun je hier heuristieken toevoegen)
                continue

# -----------------------------
# Excel output (maak sheet met planning)
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

ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1,col_idx).fill = header_fill
    ws_out.cell(1,col_idx).alignment = center_align
    ws_out.cell(1,col_idx).border = thin_border

rij_out = 2
for attr in attracties_te_plannen:
    max_pos = max(per_hour_positions[h].get(attr,0) for h in open_uren)
    max_pos = max(max_pos, aantallen.get(attr,1))
    for pos_idx in range(1, max_pos+1):
        naam_attr = attr if max_pos == 1 else f"{attr} {pos_idx}"
        ws_out.cell(rij_out,1,naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out,1).fill = attr_fill
        ws_out.cell(rij_out,1).border = thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            assigned_list = assigned_map.get((uur,attr), [])
            naam = assigned_list[pos_idx-1] if pos_idx-1 < len(assigned_list) else ""
            if naam == "NIEMAND":
                naam = ""
            ws_out.cell(rij_out,col_idx,naam).alignment = center_align
            ws_out.cell(rij_out,col_idx).border = thin_border
        rij_out += 1

rij_out += 1
for pv_idx, pvnaam in enumerate(pauzevlinder_namen, start=1):
    if not pvnaam:
        continue
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = pv_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out,col_idx,pvnaam if uur in required_pauze_hours else "").alignment = center_align
        ws_out.cell(rij_out,col_idx).border = thin_border
    rij_out += 1

for col in range(1, len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width = 18

output = BytesIO()
wb_out.save(output)
output.seek(0)
st.success("Basisplanning (blok-voor-blok toewijzing) gegenereerd!")
st.download_button("Download planning", data=output.getvalue(),
                   file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
