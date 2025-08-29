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

    uren_beschikbaar = [10 + kol - 3 for kol in range(3, 12) 
                        if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]
    
    attracties = [ws.cell(1, kol).value for kol in range(14, 32) 
                  if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]
    
    try:
        aantal_attracties = int(ws['AG' + str(rij)].value)
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
open_uren = sorted(set(
    int(str(ws.cell(1, kol).value).replace("u","").strip()) if not isinstance(ws.cell(1, kol).value,int) else ws.cell(1, kol).value
    for kol in range(36, 45) if ws.cell(2, kol).value in [1, True, "WAAR", "X"]
))
if not open_uren:
    open_uren = list(range(10, 19))

# -----------------------------
# Pauzevlinders bepalen
# -----------------------------
required_hours = [12,13,14,15,16,17]
pv_cell_value = ws['BN2'].value
try:
    num_pauzevlinders = int(float(str(pv_cell_value).replace(",",".").strip()))
except:
    num_pauzevlinders = 0

candidates = [s for s in studenten if all(u in s['uren_beschikbaar'] for u in required_hours)
              and s['aantal_attracties'] >= 8]
selected_pauzevlinders = random.sample(candidates, min(num_pauzevlinders, len(candidates))) if num_pauzevlinders > 0 else []

for idx, s in enumerate(selected_pauzevlinders, start=1):
    s["is_pauzevlinder"] = True
    s["pv_number"] = idx
    s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_hours]

# -----------------------------
# Attracties en posities inlezen (AU-BL)
# -----------------------------
attr_posities = {}
for kol in range(47, 64):  # AU=47, BL=63
    attr_naam = ws.cell(1, kol).value
    if not attr_naam:
        continue
    try:
        pos_count = int(ws.cell(2, kol).value)
    except:
        pos_count = 1
    attr_posities[attr_naam] = pos_count

# -----------------------------
# Functie voor attractieplanning
# -----------------------------
def plan_attractie(attr, studenten, student_bezet, gebruik_per_student, open_uren):
    planning = {}
    blokken = [3,4,2,1]
    uren = sorted(open_uren)
    i = 0
    while i < len(uren):
        geplanned = False
        for blok in blokken:
            if i + blok > len(uren):
                continue
            blokuren = uren[i:i+blok]
            kandidaten = [s for s in studenten 
                          if attr in s['attracties'] 
                          and all(u in s['uren_beschikbaar'] for u in blokuren)
                          and not any(u in student_bezet[s['naam']] for u in blokuren)
                          and gebruik_per_student[s['naam']][attr]+blok <=5]
            if kandidaten:
                min_usage = min(gebruik_per_student[s['naam']][attr] for s in kandidaten)
                beste = [s for s in kandidaten if gebruik_per_student[s['naam']][attr]==min_usage]
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
            kandidaten = [s for s in studenten 
                          if attr in s['attracties'] 
                          and u in s['uren_beschikbaar'] 
                          and u not in student_bezet[s['naam']]
                          and gebruik_per_student[s['naam']][attr]<5]
            if kandidaten:
                gekozen = random.choice(kandidaten)
                planning[u] = gekozen['naam']
                student_bezet[gekozen['naam']].append(u)
                gebruik_per_student[gekozen['naam']][attr] +=1
            else:
                planning[u] = ""  # leeg vakje
            i+=1
    return planning

# -----------------------------
# Dagplanning maken
# -----------------------------
dagplanning = {}
student_bezet = defaultdict(list)
gebruik_per_student = defaultdict(lambda: defaultdict(int))

for attr, pos_count in attr_posities.items():
    dagplanning[attr] = []
    for pos in range(pos_count):
        pl = plan_attractie(attr, studenten, student_bezet, gebruik_per_student, open_uren)
        # Minstens 1 persoon op eerste of tweede plaats
        if pos < 2 and all(v=="" for v in pl.values()):
            kandidaten = [s for s in studenten if attr in s['attracties']]
            if kandidaten:
                gekozen = random.choice(kandidaten)
                for uur in open_uren:
                    if uur in gekozen['uren_beschikbaar']:
                        pl[uur] = gekozen['naam']
                        student_bezet[gekozen['naam']].append(uur)
                        gebruik_per_student[gekozen['naam']][attr] +=1
                        break
        dagplanning[attr].append(pl)

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
empty_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Header
ws_out.cell(1,1,vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

# Attracties
rij_out = 2
for attr, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attr if len(posities)==1 else f"{attr} {idx}"
        ws_out.cell(rij_out,1,naam_attr).font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=attr_fill
        ws_out.cell(rij_out,1).border=thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            naam = planning.get(uur,"")
            ws_out.cell(rij_out,col_idx,naam)
            ws_out.cell(rij_out,col_idx).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
            if naam == "":
                ws_out.cell(rij_out,col_idx).fill=empty_fill
        rij_out+=1

# Pauzevlinders tussen 12u-18u
rij_out+=1
for pv_idx, pv in enumerate(selected_pauzevlinders, start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out,col_idx,pv['naam'] if uur in required_hours else "")
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Extra (optioneel: leeg in dit voorbeeld)
extra_per_uur = defaultdict(list)
rij_out+=1
max_extra = max(len(names) for names in extra_per_uur.values()) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out,1,"Extra").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=extra_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        naam = extra_per_uur[uur][i] if i<len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out,col_idx,naam)
        ws_out.cell(rij_out,col_idx).alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width = 15

# Opslaan in-memory
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.download_button("Download planning", data=output, file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
