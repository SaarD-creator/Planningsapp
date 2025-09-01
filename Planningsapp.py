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

# partitioneer een continugedragen run van lengte L in blokken
# we gebruiken DP om 1-uurs-blokjes te vermijden waar mogelijk:
def partition_run_lengths(L):
    # blokgroottes toegestaan
    blocks = [3,4,2,1]
    # dp[i] = (num_ones, num_blocks, partition_list) voor lengte i
    dp = [(10**9, 10**9, [])] * (L+1)
    dp[0] = (0, 0, [])
    for i in range(1, L+1):
        best = (10**9, 10**9, [])
        for b in blocks:
            if i-b < 0:
                continue
            prev = dp[i-b]
            num_ones = prev[0] + (1 if b==1 else 0)
            num_blocks = prev[1] + 1
            part = prev[2] + [b]
            cand = (num_ones, num_blocks, part)
            # minimaliseer eerst aantal 1's, dan aantal blocks
            if (cand[0], cand[1]) < (best[0], best[1]):
                best = cand
        dp[i] = best
    return dp[L][2]

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
# zoek kolommen 36..44 waar in rij2 een "1"/True staat
open_uren = [int(str(ws.cell(1,kol).value).replace("u","").strip()) for kol in range(36,45) if ws.cell(2,kol).value in [1,True,"WAAR","X"]]
if not open_uren:
    open_uren = list(range(10,19))
open_uren = sorted(set(open_uren))

# -----------------------------
# Pauzevlinders lezen en hun bezetting bepalen (volgens jouw regels)
# -----------------------------
# Pauzevlinder namen in kolom BN, rijen 4..10 (zoals in jouw code)
pauzevlinder_namen = [ws[f'BN{rij}'].value for rij in range(4,11) if ws[f'BN{rij}'].value]

# Kies required_hours op basis van open_uren volgens jouw regels:
def compute_pauze_hours(open_uren):
    mn = min(open_uren)
    mx = max(open_uren)
    # cases described in de opdracht:
    # - als dag 10..18 (oftewel open_uren bevat 10 en 18): bezet 12..17
    # - als dag 12..18: bezet 13..18
    # - als dag 14..18: bezet hele uren (dus alle open_uren)
    # fallback: probeer intersectie met [12..17]
    if 10 in open_uren and 18 in open_uren:
        req = [h for h in open_uren if 12 <= h <= 17]
    elif 12 in open_uren and 18 in open_uren:
        req = [h for h in open_uren if 13 <= h <= 18]
    elif min(open_uren) >= 14:
        req = list(open_uren)
    else:
        # fallback, volgende beste: 12..17 intersect open_uren
        req = [h for h in open_uren if 12 <= h <= 17]
    return sorted(req)

required_pauze_hours = compute_pauze_hours(open_uren)

# Markeer in studenten en verwijder die uren uit hun beschikbaarheid
for idx, pvnaam in enumerate(pauzevlinder_namen, start=1):
    if not pvnaam:
        continue
    for s in studenten:
        if s["naam"] == pvnaam:
            s["is_pauzevlinder"] = True
            s["pv_number"] = idx
            # verwijder de pauze-uren uit hun beschikbaarheid
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

# kritieke score: hoeveel studenten (van degenen die werken vandaag) kunnen deze attractie?
def kritieke_score(attr, studenten_list):
    return sum(1 for s in studenten_list if attr in s["attracties"])

# Filter studenten die vandaag werken (hebben minstens 1 uur beschikbaar binnen open_uren)
studenten_workend = [s for s in studenten if any(u in open_uren for u in s["uren_beschikbaar"])]

# Sorteer attracties op basis van schaarste: weinigste kandidaten eerst
attracties_te_plannen.sort(key=lambda a: kritieke_score(a, studenten_workend))

# -----------------------------
# BA kolom: tweede-posities prioriteit (BA rijen 5..12)
# -----------------------------
ba_prioriteit = [ws[f'BA{r}'].value for r in range(5,13) if ws[f'BA{r}'].value]
# verwijderen van lege cellen al gedaan; ba_prioriteit is top-to-bottom order
# we gaan later bottom-first (reversed) items weghalen indien nodig per uur

# -----------------------------
# Per-uur: beschikbare studenten (lijst) EN per-uur beschikbare plekken (vanuit excel of berekend)
# -----------------------------
# Probeer aantal studenten per uur uit excel te lezen (rij 3 in dezelfde kolommen als open_uren),
# anders fallback: tel uit studenten_workend
# Maak mapping col->hour voor kolommen 36..44
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

# Attempt to read counts from row 3
hour_count_from_sheet = {}
for kol,hr in col_hour_map.items():
    v = ws.cell(3,kol).value
    if isinstance(v, (int, float)):
        hour_count_from_sheet[hr] = int(v)

# Fallback / compute if not found
per_hour_students_list = {}
for uur in open_uren:
    # studenten beschikbaar op dit uur (excl pauze-uren we've already removed from their availability)
    names = [s for s in studenten_workend if uur in s["uren_beschikbaar"]]
    per_hour_students_list[uur] = sorted(names, key=lambda s: s["aantal_attracties"])  # sorted student-objects

# Per-uur: hoeveel studenten totaal werken (either from sheet or from computed), en trek pauzevlinders af
per_hour_total_workers = {}
for uur in open_uren:
    if uur in hour_count_from_sheet:
        totaal = hour_count_from_sheet[uur]
    else:
        totaal = len(per_hour_students_list[uur])
    # bepaal hoeveel pauzevlinders op dit uur normaal zouden zijn (we hebben ze al uit uren_beschikbaar gehaald,
    # maar als sheet gaf al reductie, we moeten eerlijk zijn: we halen gewoon pauzevlinders die *normaal* op dat uur zouden zitten)
    # eenvoud: tel hoeveel van de pauzevlinder-namen oorspronkelijk in studenten lijst hadden dit uur (zonder wijziging)
    # maar we hebben gewijzigd uren_beschikbaar; eenvoudiger: zoek pauzevlinders -> als required_pauze_hours bevat uur, haal 1 per pauzevlinder
    num_pv = sum(1 for nv in pauzevlinder_namen if nv and uur in required_pauze_hours
                 and any(s["naam"]==nv for s in studenten))
    beschikbaar = max(0, totaal - num_pv)
    per_hour_total_workers[uur] = beschikbaar

# -----------------------------
# Per-uur: positions per attractie (aanpassen: haal second posities weg als er te weinig mensen zijn)
# -----------------------------
# Start met baseline positions = aantallen (1 of 2)
# Voor elk uur maken we een positions dict attr->posities_allowed
per_hour_positions = {}
for uur in open_uren:
    positions = {attr: aantallen.get(attr, 0) for attr in attracties_te_plannen}
    totaal_posities = sum(positions.values())
    beschikbaar = per_hour_total_workers.get(uur, 0)
    if totaal_posities <= beschikbaar:
        per_hour_positions[uur] = positions
        continue

    # teveel posities: verwijder second positions volgens BA (laatste in kolom eerst)
    # we iterate reversed(ba_prioriteit) (bottom-up)
    for attr_to_remove in reversed(ba_prioriteit):
        if totaal_posities <= beschikbaar:
            break
        if attr_to_remove in positions and positions[attr_to_remove] >= 2:
            positions[attr_to_remove] = 1
            totaal_posities -= 1

    # fallback: als nog teveel, verwijder second posities van attracties met de meeste kandidaten (dus minst kritisch)
    if totaal_posities > beschikbaar:
        # order by kritieke_score descending (meeste kandidaten eerst -> we remove their second pos first)
        fallback_order = sorted(attracties_te_plannen, key=lambda a: kritieke_score(a, studenten_workend), reverse=True)
        for a in fallback_order:
            if totaal_posities <= beschikbaar:
                break
            if positions.get(a,0) >= 2:
                positions[a] = 1
                totaal_posities -=1

    # als nog teveel (we hebben geen second posities meer) dan kan het zijn dat we gewoon te weinig studenten hebben;
    # we laten dit zo (later zullen posities ongevuld blijven)
    per_hour_positions[uur] = positions

# -----------------------------
# Maak per-uur attractielijst (geordend op schaarste)
# -----------------------------
per_hour_attractielijsten = {}
for uur in open_uren:
    # we maken een lijst met attracties (repeated per positie) gesorteerd op kritieke_score (weinig kandidaten eerst)
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
# Data-structs voor toewijzingen
# assigned_map[(uur, attr)] = list of student-names assigned on that attr at that hour (volgt volgorde van toevoegen)
assigned_map = defaultdict(list)

# per hour assigned counts per attr
per_hour_assigned_counts = {uur: {a: 0 for a in attracties_te_plannen} for uur in open_uren}

# maximale opeenvolgende uren limit (zoals eerder)
MAX_CONSEC = 4
MAX_PER_STUDENT_ATTR = 6  # fallback limiter

# Sorteer studenten op aantal attracties (min first)
studenten_sorted = sorted(studenten_workend, key=lambda s: s["aantal_attracties"])

# Helper: hoeveel uren heeft student al (voor max_consecutive check)
def current_student_hours(s):
    return sorted(s["assigned_hours"])

# Enkel de eerste student toewijzen
eerste_student = studenten[0]
beschikbare_attracties = [a for a in attracties if not a['bezet']]

for i in range(min(eerste_student['aantal'], len(beschikbare_attracties))):
    attractie = beschikbare_attracties[i]
    planning.append({
        "student": eerste_student['naam'],
        "attractie": attractie['naam'],
        "shift": eerste_student['shift']
    })
    # Markeer attractie als bezet
    attractie['bezet'] = True

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

# Header : datum + uren
ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1,col_idx).fill = header_fill
    ws_out.cell(1,col_idx).alignment = center_align
    ws_out.cell(1,col_idx).border = thin_border

# Voor uitvoer: maak voor elke attractie rijen gelijk aan max aantal posities (1 of 2)
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
            # bepaal welke naam op deze pos_idx staat: we nemen assigned_map[(uur,attr)] list en vullen per pos index
            assigned_list = assigned_map.get((uur,attr), [])
            naam = assigned_list[pos_idx-1] if pos_idx-1 < len(assigned_list) else ""
            if naam == "NIEMAND":
                naam = ""
            ws_out.cell(rij_out,col_idx,naam).alignment = center_align
            ws_out.cell(rij_out,col_idx).border = thin_border
        rij_out += 1

# Pauzevlinders (toon wie het zijn en in welke uren ze pauze hebben)
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

# Kolombreedte
for col in range(1, len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width = 18

# Download button
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.success("Basisplanning (blok-voor-blok toewijzing) gegenereerd!")
st.download_button("Download planning", data=output.getvalue(),
                   file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

