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
    uren_beschikbaar = [10 + (kol - 3) for kol in range(3, 12)
                        if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]
    attracties = [ws.cell(1, kol).value for kol in range(14, 32)
                  if ws.cell(rij, kol).value in [1, True, "WAAR", "X"]]
    raw_ag = ws['AG' + str(rij)].value
    try:
        aantal_attracties = int(raw_ag) if raw_ag else len(attracties)
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
# Pauzevlinders kiezen
# -----------------------------
raw_bn2 = ws['BN2'].value
try:
    num_pauzevlinders = int(float(str(raw_bn2).replace(",", ".").strip())) if raw_bn2 else 0
except:
    num_pauzevlinders = 0

required_hours = [12, 13, 14, 15, 16, 17]
candidates = [s for s in studenten if all(u in s['uren_beschikbaar'] for u in required_hours)
              and s['aantal_attracties'] >= 8]

selected_pauzevlinders = random.sample(candidates, min(num_pauzevlinders, len(candidates))) if num_pauzevlinders > 0 else []

for idx, s in enumerate(selected_pauzevlinders, start=1):
    s["is_pauzevlinder"] = True
    s["pv_number"] = idx
    s["uren_beschikbaar"] = [u for u in s["uren_beschikbaar"] if u not in required_hours]

# -----------------------------
# Attractiecapaciteit inlezen (rij 2, kol AU-BL), namen in rij 1
# -----------------------------
attractie_capaciteit = {}
for kol in range(47, 64):  # AU=47 tot BL=63
    naam = ws.cell(1, kol).value
    raw_cap = ws.cell(2, kol).value
    try:
        cap = int(raw_cap) if raw_cap else 1
    except:
        cap = 1
    if naam:
        attractie_capaciteit[naam] = cap

# -----------------------------
# Planning voorbereiden
# -----------------------------
dagplanning = {attr: [dict() for _ in range(cap)] for attr, cap in attractie_capaciteit.items()}
student_bezet = {s['naam']: [] for s in studenten}
gebruik_per_attractie_student = {attr: {s['naam']: 0 for s in studenten} for attr in attractie_capaciteit}
blokvolgorde = [3, 4, 2, 1]

# -----------------------------
# Attracties plannen
# -----------------------------
for attr, posities in dagplanning.items():
    for plek_idx, planning in enumerate(posities):
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
# Extra planning (studenten die nog niets te doen hebben)
# -----------------------------
extra_per_uur = defaultdict(list)
for u in open_uren:
    for s in studenten:
        if u in s['uren_beschikbaar'] and u not in student_bezet[s['naam']]:
            extra_per_uur[u].append(s['naam'])
            student_bezet[s['naam']].append(u)

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
empty_fill = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")

center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Header
ws_out.cell(1, 1, vandaag).font = Font(bold=True)
for col_idx, uur in enumerate(open_uren, start=2):
    ws_out.cell(1, col_idx, f"{uur}:00").font = Font(bold=True)
    ws_out.cell(1, col_idx).fill = header_fill
    ws_out.cell(1, col_idx).alignment = center_align
    ws_out.cell(1, col_idx).border = thin_border

rij_out = 2
# Attracties
for attr, posities in dagplanning.items():
    for idx, planning in enumerate(posities, start=1):
        naam_attr = attr if len(posities) == 1 else f"{attr} {idx}"
        ws_out.cell(rij_out, 1, naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out, 1).fill = attr_fill
        ws_out.cell(rij_out, 1).border = thin_border
        for col_idx, uur in enumerate(open_uren, start=2):
            naam = planning.get(uur, "")
            ws_out.cell(rij_out, col_idx, naam)
            ws_out.cell(rij_out, col_idx).alignment = center_align
            ws_out.cell(rij_out, col_idx).border = thin_border
            if naam == "":
                ws_out.cell(rij_out, col_idx).fill = empty_fill
        rij_out += 1

# Scheidingsrij
rij_out += 1

# Pauzevlinders
for pv_idx, pv in enumerate(selected_pauzevlinders, start=1):
    ws_out.cell(rij_out, 1, f"Pauzevlinder {pv_idx}").font = Font(bold=True)
    ws_out.cell(rij_out, 1).fill = pv_fill
    ws_out.cell(rij_out, 1).border = thin_border
    for col_idx, uur in enumerate(open_uren, start=2):
        pv_name = pv["naam"] if uur in required_hours else ""
        ws_out.cell(rij_out, col_idx, pv_name)
        ws_out.cell(rij_out, col_idx).alignment = center_align
        ws_out.cell(rij_out, col_idx).border = thin_border
        if pv_name == "":
            ws_out.cell(rij_out, col_idx).fill = empty_fill
    rij_out += 1

# Extra
rij_out += 1
max_extra = max(len(names) for names in extra_per_uur.values()) if extra_per_uur else 0
for i in range(max_extra):
    ws_out.cell(rij_out, 1, "Extra").font = Font(bold=True)
    ws_out.cell(rij_out, 1).fill = extra_fill
    ws_out.cell(rij_out, 1).border = thin_border
    for col_idx, uur in enumerate(open_uren, start=2):
        naam = extra_per_uur[uur][i] if i < len(extra_per_uur[uur]) else ""
        ws_out.cell(rij_out, col_idx, naam)
        ws_out.cell(rij_out, col_idx).alignment = center_align
        ws_out.cell(rij_out, col_idx).border = thin_border
        if naam == "":
            ws_out.cell(rij_out, col_idx).fill = empty_fill
    rij_out += 1

# Kolombreedte
for col in range(1, len(open_uren) + 2):
    ws_out.column_dimensions[get_column_letter(col)].width = 15

# Excel in-memory maken
output = BytesIO()
wb_out.save(output)
output.seek(0)

st.download_button("Download Planning", data=output, file_name=f"Planning_{vandaag}.xlsx")
st.success("Planning succesvol aangemaakt!")

