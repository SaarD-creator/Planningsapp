import streamlit as st
import pandas as pd
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime
import random

# -----------------------------
# Setup
# -----------------------------
vandaag = datetime.date.today().strftime("%d-%m-%Y")
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
    uren_beschikbaar = [10 + (k-3) for k in range(3, 12) if ws.cell(rij, k).value in [1, True, "WAAR", "X"]]
    attracties = [ws.cell(1, k).value for k in range(14, 32) if ws.cell(rij, k).value in [1, True, "WAAR", "X"]]
    try:
        aantal_attracties = int(ws['AG'+str(rij)].value) if ws['AG'+str(rij)].value is not None else len(attracties)
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
# Openingsuren inlezen
# -----------------------------
open_uren = [int(str(ws.cell(1, k).value).replace("u","").strip()) for k in range(36, 45) if ws.cell(2, k).value in [1, True, "WAAR", "X"]]
if not open_uren:
    open_uren = list(range(10,19))
open_uren = sorted(open_uren)

# -----------------------------
# Attracties met aantal plaatsen inlezen (AU2:BL2)
# -----------------------------
attractie_capaciteit = {}
for kol in range(47, 64):  # AU=47 tot BL=63
    attr = ws.cell(1, kol).value
    cap = ws.cell(2, kol).value
    if attr and cap and cap > 0:
        attractie_capaciteit[attr] = int(cap)

# -----------------------------
# Pauzevlinders selecteren (BN4-BN10)
# -----------------------------
required_hours = list(range(12,18))
pv_cells = [ws['BN'+str(r)].value for r in range(4,11)]
selected_pauzevlinders = [s for s in studenten if s["naam"] in pv_cells]
for idx, s in enumerate(selected_pauzevlinders, start=1):
    s["is_pauzevlinder"] = True
    s["pv_number"] = idx
    s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_hours]

# -----------------------------
# Dagplanning initialiseren
# -----------------------------
dagplanning = {}
for attr, cap in attractie_capaciteit.items():
    dagplanning[attr] = [defaultdict(lambda: "") for _ in range(cap)]  # één dict per positie

student_bezet = {s["naam"]: [] for s in studenten}
gebruik_per_student = {s["naam"]: 0 for s in studenten}

# -----------------------------
# Attracties plannen
# -----------------------------
blok_volgorde = [3,4,2,1]

for attr, posities in dagplanning.items():
    for pos_idx, pos in enumerate(posities):
        for u in open_uren:
            if pos[u] != "":
                continue  # al ingevuld
            kandidaten = [s for s in studenten if attr in s["attracties"] and u in s["uren_beschikbaar"]]
            if not kandidaten:
                continue
            # Sorteer op minst aantal attracties
            kandidaten.sort(key=lambda x: x["aantal_attracties"])
            for blok in blok_volgorde:
                uren_bloc = [hr for hr in range(u, u+blok) if hr in open_uren]
                if all(pos[hr]=="" for hr in uren_bloc):
                    for hr in uren_bloc:
                        pos[hr] = kandidaten[0]["naam"]
                        student_bezet[kandidaten[0]["naam"]].append(hr)
                        gebruik_per_student[kandidaten[0]["naam"]] += len(uren_bloc)
                    break

# -----------------------------
# Extra per uur invullen
# -----------------------------
extra_per_uur = defaultdict(list)
for u in open_uren:
    for s in studenten:
        if s["is_pauzevlinder"] or u not in s["uren_beschikbaar"]:
            continue
        if u not in student_bezet[s["naam"]]:
            extra_per_uur[u].append(s["naam"])
            student_bezet[s["naam"]].append(u)

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

# Header
ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(open_uren, start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out = 2
# Attracties
for attr, posities in dagplanning.items():
    for idx, pos in enumerate(posities, start=1):
        naam_attr = attr if len(posities)==1 else f"{attr} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out,1).fill = attr_fill
        ws_out.cell(rij_out,1).border = thin_border
        for col_idx, uur in enumerate(open_uren, start=2):
            naam = pos[uur]
            ws_out.cell(rij_out,col_idx,naam)
            if naam=="":
                ws_out.cell(rij_out,col_idx).fill=empty_fill
            ws_out.cell(rij_out,col_idx).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out+=1

# Pauzevlinders
rij_out+=1
for pv_idx, s in enumerate(selected_pauzevlinders, start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(open_uren, start=2):
        ws_out.cell(rij_out,col_idx,s["naam"] if uur in required_hours else "")
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Extra
rij_out+=1
max_extra = max(len(lst) for lst in extra_per_uur.values()) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=extra_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(open_uren, start=2):
        naam = extra_per_uur[uur][i] if i<len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out,col_idx,naam)
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=15

timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
planning_bestand = f"Planning_{timestamp}.xlsx"
output = BytesIO()
wb_out.save(output)
output.seek(0)

st.download_button("Download planning", output, file_name=planning_bestand)
st.success(f"Aantal studenten ingeladen: {len(studenten)}")

