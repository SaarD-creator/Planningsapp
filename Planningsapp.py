import streamlit as st
import pandas as pd
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import datetime
from io import BytesIO

# Datum van vandaag
vandaag = datetime.date.today().strftime("%d-%m-%Y")

# -----------------------------
# Excelbestand openen
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

    uren_beschikbaar = [10 + (kol - 3) for kol in range(3, 12)
                        if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]

    attracties = [ws.cell(1, kol).value for kol in range(14, 32)
                  if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]

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
# Openingsuren
# -----------------------------
open_uren = [int(str(ws.cell(1, kol).value).replace("u", "").strip())
             for kol in range(36, 45)
             if ws.cell(2, kol).value in [1, True, "WAAR", "X"]]

if not open_uren:
    open_uren = list(range(10, 19))
open_uren = sorted(set(open_uren))

# -----------------------------
# Attractie capaciteit (AU2:BL2)
# -----------------------------
attractie_capaciteit = {}
for kol in range(47, 64):  # AU=47, BL=64
    attr = ws.cell(1, kol).value
    try:
        cap = int(ws.cell(2, kol).value)
    except:
        cap = 1
    if attr:
        attractie_capaciteit[attr] = cap

# -----------------------------
# Pauzevlinders kiezen (BN4-BN10)
# -----------------------------
required_hours = list(range(12, 18))  # 12u-18u
selected_pauzevlinders = []

for rij in range(4, 11):  # BN4-BN10
    naam = ws['BN' + str(rij)].value
    if naam:
        for s in studenten:
            if s['naam'] == naam:
                s['is_pauzevlinder'] = True
                s['pv_number'] = len(selected_pauzevlinders) + 1
                s['uren_beschikbaar'] = [u for u in s['uren_beschikbaar'] if u not in required_hours]
                selected_pauzevlinders.append(s)
                break

# -----------------------------
# Planning dictionary
# -----------------------------
dagplanning = {attr: [dict() for _ in range(cap)] for attr, cap in attractie_capaciteit.items()}
student_bezet = {s['naam']: [] for s in studenten}
gebruik_per_attractie_student = {attr: {s['naam']: 0 for s in studenten} for attr in attractie_capaciteit}

# -----------------------------
# Attracties inplannen
# -----------------------------
blokvolgorde = [3, 4, 2, 1]  # blokkenvolgorde

for attr, posities in dagplanning.items():
    for idx, planning in enumerate(posities):
        uren = sorted(open_uren)
        i = 0
        while i < len(uren):
            geplanned = False
            for blok in blokvolgorde:
                if i + blok > len(uren):
                    continue
                blokuren = uren[i:i+blok]
                kandidaten = [s for s in studenten
                              if attr in s['attracties']
                              and all(u in s['uren_beschikbaar'] for u in blokuren)
                              and not any(u in student_bezet[s['naam']] for u in blokuren)
                              and gebruik_per_attractie_student[attr][s['naam']] + blok <= 5]
                if kandidaten:
                    min_uren = min(gebruik_per_attractie_student[attr][s['naam']] for s in kandidaten)
                    beste = [s for s in kandidaten if gebruik_per_attractie_student[attr][s['naam']] == min_uren]
                    gekozen = random.choice(beste)
                    for u in blokuren:
                        planning[u] = gekozen['naam']
                        student_bezet[gekozen['naam']].append(u)
                    gebruik_per_attractie_student[attr][gekozen['naam']] += blok
                    i += blok
                    geplanned = True
                    break
            if not geplanned:
                u = uren[i]
                kandidaten_1 = [s for s in studenten
                                if attr in s['attracties']
                                and u in s['uren_beschikbaar']
                                and u not in student_bezet[s['naam']]
                                and gebruik_per_attractie_student[attr][s['naam']] < 5]
                if kandidaten_1:
                    gekozen = random.choice(kandidaten_1)
                    planning[u] = gekozen['naam']
                    student_bezet[gekozen['naam']].append(u)
                    gebruik_per_attractie_student[attr][gekozen['naam']] += 1
                i += 1

# -----------------------------
# Extra planning (studenten zonder taak in dat uur)
# -----------------------------
extra_per_uur = defaultdict(list)
for s in studenten:
    for u in open_uren:
        if u not in student_bezet[s['naam']]:
            extra_per_uur[u].append(s['naam'])

# -----------------------------
# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"

# Fills en styles
header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
attr_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
pv_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
extra_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
empty_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Feedback
ws_feedback = wb_out.create_sheet("Feedback")
def log_feedback(msg):
    next_row = ws_feedback.max_row + 1
    ws_feedback.cell(next_row, 1, msg)

# Header
ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1, col_idx, f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1, col_idx).fill = header_fill
    ws_out.cell(1, col_idx).alignment = center_align
    ws_out.cell(1, col_idx).border = thin_border

# Attracties
rij_out = 2
for attr, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attr if len(posities) == 1 else f"{attr} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out,1).fill = attr_fill
        ws_out.cell(rij_out,1).border = thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            naam = planning.get(uur,"")
            cell = ws_out.cell(rij_out, col_idx, naam)
            cell.alignment = center_align
            cell.border = thin_border
            cell.fill = empty_fill if naam=="" else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        rij_out += 1

# Pauzevlinders
rij_out += 1
for pv_idx, s in enumerate(selected_pauzevlinders, start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = pv_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = s['naam'] if uur in required_hours else ""
        cell = ws_out.cell(rij_out, col_idx, naam)
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = pv_fill if uur in required_hours else empty_fill
    rij_out += 1

# Extra
rij_out += 1
max_extra = max(len(names) for names in extra_per_uur.values()) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = extra_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = extra_per_uur[uur][i] if i < len(extra_per_uur[uur]) else ""
        cell = ws_out.cell(rij_out, col_idx, naam)
        cell.alignment = center_align
        cell.border = thin_border
        cell.fill = extra_fill if naam else empty_fill
    rij_out += 1

# Kolombreedte
for col in range(1, len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width = 15

# Opslaan in geheugen
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
planning_bestand = f"Planning_{timestamp}.xlsx"
output = BytesIO()
wb_out.save(output)
output.seek(0)

st.download_button("Download Planning", data=output, file_name=planning_bestand)
log_feedback(f"Aantal studenten ingeladen: {len(studenten)}")
