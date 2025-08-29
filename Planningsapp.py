import streamlit as st
import pandas as pd
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime

# -----------------------------
# Configuratie
# -----------------------------
vandaag = datetime.date.today().strftime("%d-%m-%Y")
blokvolgorde = [3, 4, 2, 1]  # volgorde van blokken
max_blokken = 5  # normaal
max_blokken_no_alt = 6  # indien absoluut geen optie

# -----------------------------
# Streamlit bestand upload
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
            uren_beschikbaar.append(10 + kol - 3)

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
# Attracties en aantal plaatsen
# -----------------------------
attracties_info = {}
for kol in range(47, 64):  # AU-BL
    attr_naam = ws.cell(1, kol).value
    try:
        max_personen = int(ws.cell(2, kol).value)
    except:
        max_personen = 0
    if max_personen > 0:
        attracties_info[attr_naam] = max_personen

# -----------------------------
# Pauzevlinders
# -----------------------------
raw_bn2 = ws['BN2'].value
try:
    num_pauzevlinders = int(float(str(raw_bn2).replace(",", ".").strip())) if raw_bn2 else 0
except:
    num_pauzevlinders = 0

required_hours = [12, 13, 14, 15, 16, 17]  # 12u-18u
candidates = [s for s in studenten if all(u in s['uren_beschikbaar'] for u in required_hours) and s['aantal_attracties'] >= 8]
selected_pauzevlinders = random.sample(candidates, min(num_pauzevlinders, len(candidates))) if num_pauzevlinders > 0 else []
for idx, s in enumerate(selected_pauzevlinders, start=1):
    s["is_pauzevlinder"] = True
    s["pv_number"] = idx
    s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_hours]

# -----------------------------
# Planning voorbereiden
# -----------------------------
dagplanning = {attr: [{} for _ in range(attracties_info[attr])] for attr in attracties_info}
student_bezet = {s["naam"]: [] for s in studenten}
gebruik_per_attractie_student = {attr: {s["naam"]: 0 for s in studenten} for attr in attracties_info}

# Functie om studenten bij attracties in te plannen
def plan_student(student, dagplanning, student_bezet, gebruik_per_attractie_student, open_uren):
    for blok in blokvolgorde:
        for attr, posities in dagplanning.items():
            for idx, pos in enumerate(posities):
                # check welke uren beschikbaar zijn
                uren_toevoegen = [u for u in open_uren if u in student["uren_beschikbaar"] and u not in student_bezet[student["naam"]]]
                if len(uren_toevoegen) >= blok:
                    for u in uren_toevoegen[:blok]:
                        dagplanning[attr][idx][u] = student["naam"]
                        student_bezet[student["naam"]].append(u)
                        gebruik_per_attractie_student[attr][student["naam"]] += 1
                    break
            break  # 1 attractie per student per ronde

# Sorteer studenten op aantal attracties (minimaal eerst)
studenten_sorted = sorted([s for s in studenten if not s["is_pauzevlinder"]], key=lambda x: x["aantal_attracties"])
for student in studenten_sorted:
    plan_student(student, dagplanning, student_bezet, gebruik_per_attractie_student, open_uren)

# -----------------------------
# Extra studenten
# -----------------------------
extra_per_uur = defaultdict(list)
for s in studenten:
    for u in s["uren_beschikbaar"]:
        if u not in student_bezet[s["naam"]]:
            extra_per_uur[u].append(s["naam"])

# -----------------------------
# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"

# Kleuren en opmaak
header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
attr_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
pv_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
extra_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
leeg_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

ws_feedback = wb_out.create_sheet("Feedback")
def log_feedback(msg):
    next_row = ws_feedback.max_row + 1
    ws_feedback.cell(next_row, 1, msg)

# Header
ws_out.cell(1, 1, vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out=2

# Attracties
for attr, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attr if len(posities)==1 else f"{attr} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=attr_fill
        ws_out.cell(rij_out,1).border=thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            naam = planning.get(uur,"")
            ws_out.cell(rij_out,col_idx,naam)
            if naam == "":
                ws_out.cell(rij_out,col_idx).fill = leeg_fill
            ws_out.cell(rij_out,col_idx).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out+=1

# Scheidingsrij
rij_out+=1

# Pauzevlinders
for pv_idx, s in enumerate(selected_pauzevlinders, start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out,col_idx, s["naam"] if uur in required_hours else "")
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Extra
rij_out+=1
max_extra = max(len(names) for names in extra_per_uur.values()) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=extra_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = extra_per_uur[uur][i] if i < len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out,col_idx,naam)
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=15

# In-memory bestand
output = BytesIO()
wb_out.save(output)
output.seek(0)

st.download_button("Download Planning", output, file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
log_feedback(f"Aantal studenten ingeladen: {len(studenten)}")
