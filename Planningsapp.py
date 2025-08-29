import streamlit as st
import pandas as pd
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import datetime
from io import BytesIO

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
            uur = 10 + (kol - 3)
            uren_beschikbaar.append(uur)

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
required_hours = [12,13,14,15,16,17]
selected_pauzevlinders = []

for rij in range(4, 11):  # BN4-BN10
    naam = ws.cell(rij, 14).value  # Pas kolom aan als nodig
    if naam:
        for s in studenten:
            if s['naam'] == naam:
                s['is_pauzevlinder'] = True
                s['pv_number'] = len(selected_pauzevlinders)+1
                s['uren_beschikbaar'] = [u for u in s['uren_beschikbaar'] if u not in required_hours]
                selected_pauzevlinders.append(s)
                break

# -----------------------------
# Planning functies
# -----------------------------
def plan_attractie(attr, studenten, student_bezet, gebruik_per_student, open_uren):
    planning = {}
    uren = sorted(open_uren)
    blokken_volgorde = [3,4,2,1]  # volgorde van blokken
    i = 0
    while i < len(uren):
        geplanned = False
        for blok in blokken_volgorde:
            if i + blok > len(uren):
                continue
            blokuren = uren[i:i+blok]
            kandidaten = []
            for s in studenten:
                naam = s["naam"]
                if gebruik_per_student[naam][attr] + blok > 5:
                    continue
                if attr not in s["attracties"]:
                    continue
                if not all(u in s["uren_beschikbaar"] for u in blokuren):
                    continue
                if any(u in student_bezet[naam] for u in blokuren):
                    continue
                kandidaten.append(s)
            if kandidaten:
                min_count = min(gebruik_per_student[s["naam"]][attr] for s in kandidaten)
                beste = [s for s in kandidaten if gebruik_per_student[s["naam"]][attr]==min_count]
                gekozen = random.choice(beste)
                for u in blokuren:
                    planning[u] = gekozen["naam"]
                    student_bezet[gekozen["naam"]].append(u)
                gebruik_per_student[gekozen["naam"]][attr] += blok
                i += blok
                geplanned = True
                break
        if not geplanned:
            u = uren[i]
            kandidaten = [s for s in studenten if attr in s["attracties"] and u in s["uren_beschikbaar"] and u not in student_bezet[s["naam"]] and gebruik_per_student[s["naam"]][attr]<5]
            if kandidaten:
                gekozen = random.choice(kandidaten)
                planning[u] = gekozen["naam"]
                student_bezet[gekozen["naam"]].append(u)
                gebruik_per_student[gekozen["naam"]][attr] += 1
            else:
                planning[u] = ""  # lege cel ipv NIEMAND
            i += 1
    return planning

# -----------------------------
# Dagplanning opstellen
# -----------------------------
dagplanning = {}
student_bezet = defaultdict(list)
gebruik_per_student = defaultdict(lambda: defaultdict(int))
for s in studenten:
    for attr in s['attracties']:
        gebruik_per_student[s['naam']][attr] = 0

attractie_namen = sorted({a for s in studenten for a in s['attracties']})
for attr in attractie_namen:
    dagplanning[attr] = [plan_attractie(attr, studenten, student_bezet, gebruik_per_student, open_uren)]

# -----------------------------
# Extra per uur (studenten zonder blok)
# -----------------------------
extra_per_uur = defaultdict(list)
for s in studenten:
    for u in s['uren_beschikbaar']:
        if u not in student_bezet[s['naam']]:
            extra_per_uur[u].append(s['naam'])

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
for col_idx, uur in enumerate(open_uren, start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill = header_fill
    ws_out.cell(1,col_idx).alignment = center_align
    ws_out.cell(1,col_idx).border = thin_border

rij_out = 2
# Attracties
for attr, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attr if len(posities)==1 else f"{attr} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out,1).fill = attr_fill
        ws_out.cell(rij_out,1).border = thin_border
        for col_idx, uur in enumerate(open_uren, start=2):
            naam = planning.get(uur, "")
            if naam=="":
                ws_out.cell(rij_out,col_idx,"").fill = empty_fill
            ws_out.cell(rij_out,col_idx,naam)
            ws_out.cell(rij_out,col_idx).alignment = center_align
            ws_out.cell(rij_out,col_idx).border = thin_border
        rij_out += 1

rij_out += 1
# Pauzevlinders
for pv in selected_pauzevlinders:
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv['pv_number']}").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = pv_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(open_uren, start=2):
        ws_out.cell(rij_out,col_idx, pv['naam'] if uur in required_hours else "")
        ws_out.cell(rij_out,col_idx).alignment = center_align
        ws_out.cell(rij_out,col_idx).border = thin_border
    rij_out += 1

rij_out += 1
# Extra studenten
max_extra = max([len(namen) for namen in extra_per_uur.values()]) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = extra_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(open_uren, start=2):
        naam = extra_per_uur[uur][i] if i < len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out,col_idx,naam)
        ws_out.cell(rij_out,col_idx).alignment = center_align
        ws_out.cell(rij_out,col_idx).border = thin_border
    rij_out += 1

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width = 15

# Bestand opslaan
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
planning_bestand = f"Planning_{timestamp}.xlsx"
output = BytesIO()
wb_out.save(output)
output.seek(0)

st.download_button("Download planning", output, file_name=planning_bestand)
log_feedback(f"Aantal studenten ingeladen: {len(studenten)}")

st.title("Planning Generator")

st.write("Upload je Excel-bestand om een planning te maken.")



st.download_button(
    label="Download Planning Excel",
    data=output,
    file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

