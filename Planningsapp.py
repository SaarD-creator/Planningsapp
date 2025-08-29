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
if uploaded_file is None:
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

    uren_beschikbaar = [10 + kol - 3 for kol in range(3, 12) 
                        if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]

    attracties = [ws.cell(1, kol).value for kol in range(14, 32) 
                  if ws.cell(rij, kol).value in [1, True, "WAAR", "X"] and ws.cell(1, kol).value]

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
        uur = int(str(uur_raw).replace("u", "").strip()) if not isinstance(uur_raw, int) else uur_raw
        open_uren.append(uur)
open_uren = sorted(set(open_uren)) if open_uren else list(range(10, 19))

# -----------------------------
# Pauzevlinders selecteren (BN4-BN10)
# -----------------------------
required_hours = [12, 13, 14, 15, 16, 17]
pauzevlinders_namen = []
for rij in range(4, 11):
    naam = ws['BN' + str(rij)].value
    if naam:
        pauzevlinders_namen.append(naam.strip())

selected_pauzevlinders = []
for pv_num, naam in enumerate(pauzevlinders_namen, start=1):
    for s in studenten:
        if s["naam"] == naam:
            s["is_pauzevlinder"] = True
            s["pv_number"] = pv_num
            s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_hours]
            selected_pauzevlinders.append(s)
            break

# -----------------------------
# Attracties inlezen
# -----------------------------
attracties_te_plannen = []
aantallen = {}
for kol in range(47, 65):
    naam = ws.cell(1, kol).value
    if not naam:
        continue
    raw = ws.cell(2, kol).value
    try:
        aantal = int(raw) if raw is not None else 0
    except:
        aantal = 0
    aantallen[naam] = max(0, min(2, aantal))
    if aantallen[naam] >= 1:
        attracties_te_plannen.append(naam)

# -----------------------------
# Initialiseer gebruik_per_attractie_student
# -----------------------------
gebruik_per_attractie_student = {attr: {s["naam"]: 0 for s in studenten} for attr in attracties_te_plannen}

# -----------------------------
# Plan 1 student op attractie
# -----------------------------
def plan_student(student, dagplanning, student_bezet, gebruik_per_attractie_student, open_uren):
    blokvolgorde = [3,4,2,1]  # aangepaste volgorde
    beschikbare_uren = sorted(student["uren_beschikbaar"])
    i = 0
    while i < len(beschikbare_uren):
        geplanned = False
        for blok in blokvolgorde:
            if i + blok > len(beschikbare_uren):
                continue
            blokuren = beschikbare_uren[i:i+blok]
            mogelijke_attracties = [a for a in student["attracties"] if a in dagplanning]

            # Kies attractie met minst aantal geplande studenten
            min_count = float('inf')
            gekozen_attr = None
            for a in mogelijke_attracties:
                count = sum(gebruik_per_attractie_student[a].values())
                if count < min_count:
                    min_count = count
                    gekozen_attr = a

            if gekozen_attr is None:
                continue

            # Check of student max 5 uur bij attractie overschrijdt
            if gebruik_per_attractie_student[gekozen_attr][student['naam']] + blok > 5:
                continue

            # Plan student
            for u in blokuren:
                dagplanning[gekozen_attr][u] = student['naam']
                student_bezet[student['naam']].append(u)
                gebruik_per_attractie_student[gekozen_attr][student['naam']] += 1
            i += blok
            geplanned = True
            break

        if not geplanned:
            i += 1

# -----------------------------
# Dagplanning initialiseren
# -----------------------------
dagplanning = {attr: {u: "NIEMAND" for u in open_uren} for attr in attracties_te_plannen}
student_bezet = {s["naam"]: [] for s in studenten}

# -----------------------------
# Studenten sorteren op aantal attracties (minimaal eerst)
# -----------------------------
studenten_te_plannen = [s for s in studenten if not s["is_pauzevlinder"]]
studenten_te_plannen.sort(key=lambda x: x["aantal_attracties"])

# -----------------------------
# Planning uitvoeren
# -----------------------------
for student in studenten_te_plannen:
    plan_student(student, dagplanning, student_bezet, gebruik_per_attractie_student, open_uren)

# -----------------------------
# Streamlit / Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"
header_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
attr_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
pv_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Header
ws_out.cell(1, 1, vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out=2
# Attracties
for attractie, uren in dagplanning.items():
    ws_out.cell(rij_out,1,attractie).font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=attr_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out, col_idx, uren.get(uur,""))
        ws_out.cell(rij_out, col_idx).alignment=center_align
        ws_out.cell(rij_out, col_idx).border=thin_border
    rij_out += 1

# Pauzevlinders
rij_out += 1
for pv in selected_pauzevlinders:
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv['pv_number']}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out,col_idx, pv['naam'] if uur in required_hours else "")
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out +=1

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=15

# Save in-memory
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.download_button("Download planning", data=output, file_name=f"Planning_{vandaag}.xlsx")
st.success(f"Aantal studenten ingeladen: {len(studenten)}")

st.title("Planning Generator")

st.write("Upload je Excel-bestand om een planning te maken.")



st.download_button(
    label="Download Planning Excel",
    data=output,
    file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

