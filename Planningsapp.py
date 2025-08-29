import streamlit as st
import pandas as pd
import copy
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import datetime
from io import BytesIO

# -----------------------------
# Datum van vandaag
# -----------------------------
vandaag = datetime.date.today().strftime("%d-%m-%Y")

# -----------------------------
# Excelbestand uploaden
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
if not uploaded_file:
    st.warning("Upload eerst een Excel-bestand om verder te gaan.")
    st.stop()
wb = load_workbook(uploaded_file)
ws = wb["Blad1"]

# -----------------------------
# Studenten inlezen
# -----------------------------
studenten = []
for rij in range(2, 500):
    naam = ws.cell(rij, 12).value
    if not naam:
        continue

    uren_beschikbaar = []
    for kol in range(3, 12):
        val = ws.cell(rij, kol).value
        if val in [1, True, "WAAR", "X"]:
            uren_beschikbaar.append(10 + (kol - 3))

    attracties = []
    for kol in range(14, 32):
        val = ws.cell(rij, kol).value
        if val in [1, True, "WAAR", "X"]:
            attr_naam = ws.cell(1, kol).value
            if attr_naam:
                attracties.append(attr_naam)

    raw_ag = ws['AG' + str(rij)].value
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
# Openingsuren inlezen (AJ:AR)
# -----------------------------
open_uren = []
for kol in range(36, 45):
    uur_raw = ws.cell(1, kol).value
    vink = ws.cell(2, kol).value
    if vink in [1, True, "WAAR", "X"]:
        uur = int(str(uur_raw).replace("u", "").strip()) if not isinstance(uur_raw, int) else uur_raw
        open_uren.append(uur)
if not open_uren:
    open_uren = list(range(10, 19))
open_uren = sorted(set(open_uren))

# -----------------------------
# Pauzevlinders kiezen (BN4-BN10)
# -----------------------------
required_hours = list(range(12, 18))
num_pauzevlinders = 0
bn_range = range(4, 11)
for idx in bn_range:
    val = ws['BN' + str(idx)].value
    if val and str(val).strip() != "":
        num_pauzevlinders += 1

# Pauzevlinder kandidaten: voldoende uren en minstens 8 attracties
candidates = [s for s in studenten if all(u in s['uren_beschikbaar'] for u in required_hours) and s['aantal_attracties'] >= 8]
selected_pauzevlinders = candidates[:num_pauzevlinders] if candidates else []

for idx, s in enumerate(selected_pauzevlinders, start=1):
    s['is_pauzevlinder'] = True
    s['pv_number'] = idx
    s['uren_beschikbaar'] = [u for u in s['uren_beschikbaar'] if u not in required_hours]

# -----------------------------
# Attracties en aantal posities (rij 1 = naam, rij 2 = aantal)
# -----------------------------
attr_posities = {}
for kol in range(47, 64):  # AU-BL
    attr_name = ws.cell(1, kol).value
    aantal = ws.cell(2, kol).value
    try:
        aantal = int(aantal)
    except:
        aantal = 1
    if attr_name:
        attr_posities[attr_name] = aantal

# -----------------------------
# Functie om attracties te plannen
# -----------------------------
def plan_attractie(attr, studenten, student_bezet, gebruik_per_student, open_uren):
    planning = {}
    blokken = [3, 4, 2, 1]
    i = 0
    uren = sorted(open_uren)
    while i < len(uren):
        geplanned = False
        for blok in blokken:
            if i + blok > len(uren):
                continue
            blokuren = uren[i:i+blok]
            kandidaten = []
            for s in studenten:
                naam = s['naam']
                if attr not in s['attracties']:
                    continue
                if any(u in student_bezet[naam] for u in blokuren):
                    continue
                if not all(u in s['uren_beschikbaar'] for u in blokuren):
                    continue
                kandidaten.append(s)
            if kandidaten:
                min_count = min([gebruik_per_student[s['naam']][attr] for s in kandidaten])
                beste = [s for s in kandidaten if gebruik_per_student[s['naam']][attr] == min_count]
                gekozen = random.choice(beste)
                for u in blokuren:
                    planning[u] = gekozen['naam']
                    student_bezet[gekozen['naam']].append(u)
                gebruik_per_student[gekozen['naam']][attr] += blok
                i += blok
                geplanned = True
                break
        if not geplanned:
            u = uren[i]
            kandidaten = [s for s in studenten if attr in s['attracties'] and u in s['uren_beschikbaar'] and u not in student_bezet[s['naam']]]
            if kandidaten:
                gekozen = random.choice(kandidaten)
                planning[u] = gekozen['naam']
                student_bezet[gekozen['naam']].append(u)
                gebruik_per_student[gekozen['naam']][attr] += 1
            else:
                planning[u] = ""
            i += 1
    return planning

# -----------------------------
# Dagplanning maken
# -----------------------------
dagplanning = {}
student_bezet = defaultdict(list)
gebruik_per_student = defaultdict(lambda: defaultdict(int))
extra_per_uur = defaultdict(list)

for attr, pos_count in attr_posities.items():
    dagplanning[attr] = []
    for pos in range(pos_count):
        pl = plan_attractie(attr, studenten, student_bezet, gebruik_per_student, open_uren)
        # Minstens 1 student per attractie per positie
        for uur in open_uren:
            if pl.get(uur, "") == "":
                kandidaten = [s for s in studenten if attr in s['attracties'] and uur in s['uren_beschikbaar'] and uur not in student_bezet[s['naam']]]
                if kandidaten:
                    gekozen = random.choice(kandidaten)
                    pl[uur] = gekozen['naam']
                    student_bezet[gekozen['naam']].append(uur)
                    gebruik_per_student[gekozen['naam']][attr] += 1
        dagplanning[attr].append(pl)

# -----------------------------
# Extra-studenten bepalen
# -----------------------------
for uur in open_uren:
    for s in studenten:
        if uur not in student_bezet[s['naam']] and not s['is_pauzevlinder']:
            extra_per_uur[uur].append(s['naam'])

# -----------------------------
# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"

header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
attr_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
pv_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
extra_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
empty_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

ws_feedback = wb_out.create_sheet("Feedback")
def log_feedback(msg):
    next_row = ws_feedback.max_row + 1
    ws_feedback.cell(next_row, 1, msg)

# Header
ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1,col_idx).fill = header_fill
    ws_out.cell(1,col_idx).alignment = center_align
    ws_out.cell(1,col_idx).border = thin_border

rij_out = 2

# Attracties
for attractie, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attractie if len(posities)==1 else f"{attractie} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out,1).fill = attr_fill
        ws_out.cell(rij_out,1).border = thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            naam = planning.get(uur,"")
            ws_out.cell(rij_out,col_idx,naam if naam else "").fill = empty_fill if naam=="" else None
            ws_out.cell(rij_out,col_idx).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out += 1

# Scheidingsrij
rij_out += 1

# Pauzevlinders
for pv_idx in range(1,len(selected_pauzevlinders)+1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = pv_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam=""
        for s in selected_pauzevlinders:
            if s['pv_number']==pv_idx and uur in required_hours:
                naam=s['naam']
                break
        ws_out.cell(rij_out,col_idx,naam)
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out +=1

# Extra
rij_out += 1
max_extra = max([len(lst) for lst in extra_per_uur.values()]) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = extra_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = extra_per_uur[uur][i] if i < len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out,col_idx,naam)
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out += 1

# Kolombreedtes
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width = 15

# Opslaan in-memory
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.download_button("Download planning", data=output, file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
log_feedback(f"Aantal studenten ingeladen: {len(studenten)}")

