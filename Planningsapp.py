# beste tot nu toe met logische indeling


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
# Excelbestand uploaden
# -----------------------------
uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
if not uploaded_file:
    st.warning("Upload eerst een Excel-bestand om verder te gaan.")
    st.stop()

wb = load_workbook(uploaded_file)
ws = wb["Blad1"]

# -----------------------------
# Hulpfuncties
# -----------------------------
def max_consecutive_hours(hours):
    """Bepaal de langste aaneengesloten reeks uren in een lijst."""
    if not hours:
        return 0
    hours = sorted(set(hours))
    max_run = run = 1
    for i in range(1, len(hours)):
        if hours[i] == hours[i-1] + 1:
            run += 1
            max_run = max(max_run, run)
        else:
            run = 1
    return max_run

def contiguous_runs(sorted_hours):
    runs=[]
    if not sorted_hours:
        return runs
    run=[sorted_hours[0]]
    for h in sorted_hours[1:]:
        if h==run[-1]+1:
            run.append(h)
        else:
            runs.append(run)
            run=[h]
    runs.append(run)
    return runs

def partition_run_lengths(L):
    """Blokken van 3-4-2-1 als voorkeur, ondergeschikt aan andere eisen"""
    blocks = [3,4,2,1]
    result = []
    i = 0
    while L > 0:
        for b in blocks:
            if b <= L:
                result.append(b)
                L -= b
                break
    return result

# -----------------------------
# Studenten inlezen
# -----------------------------
studenten=[]
for rij in range(2,500):
    naam = ws.cell(rij,12).value
    if not naam:
        continue
    uren_beschikbaar=[10+(kol-3) for kol in range(3,12) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    attracties=[ws.cell(1,kol).value for kol in range(14,32) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    try:
        aantal_attracties=int(ws[f'AG{rij}'].value) if ws[f'AG{rij}'].value else len(attracties)
    except:
        aantal_attracties=len(attracties)
    studenten.append({
        "naam": naam,
        "uren_beschikbaar": sorted(uren_beschikbaar),
        "attracties":[a for a in attracties if a],
        "aantal_attracties":aantal_attracties,
        "is_pauzevlinder":False,
        "pv_number":None,
        "assigned_attracties":set(),
        "assigned_hours":[],
        "kan": []
    })

# -----------------------------
# Openingsuren
# -----------------------------
open_uren=[int(str(ws.cell(1,kol).value).replace("u","").strip()) for kol in range(36,45) if ws.cell(2,kol).value in [1,True,"WAAR","X"]]
if not open_uren:
    open_uren=list(range(10,19))
open_uren=sorted(set(open_uren))

# -----------------------------
# Pauzevlinders
# -----------------------------
pauzevlinder_namen=[ws[f'BN{rij}'].value for rij in range(4,11) if ws[f'BN{rij}'].value]

def compute_pauze_hours(open_uren):
    if 10 in open_uren and 18 in open_uren:
        return [h for h in open_uren if 12 <= h <= 17]
    elif 10 in open_uren and 17 in open_uren:
        return [h for h in open_uren if 12 <= h <= 17]
    elif 12 in open_uren:
        return [h for h in open_uren if 13 <= h <= 18]
    else:
        return list(open_uren)

required_pauze_hours=compute_pauze_hours(open_uren)

for idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    for s in studenten:
        if s["naam"]==pvnaam:
            s["is_pauzevlinder"]=True
            s["pv_number"]=idx
            s["uren_beschikbaar"]=[u for u in s["uren_beschikbaar"] if u not in required_pauze_hours]
            break

# -----------------------------
# Attracties & aantallen
# -----------------------------
aantallen={}
attracties_te_plannen=[]
for kol in range(47,65):
    naam=ws.cell(1,kol).value
    if naam:
        try: aantal=int(ws.cell(2,kol).value)
        except: aantal=0
        aantallen[naam]=max(0,min(2,aantal))
        if aantallen[naam]>=1:
            attracties_te_plannen.append(naam)

def kritieke_score(attr,studenten_list):
    return sum(1 for s in studenten_list if attr in s["attracties"])

studenten_workend=[s for s in studenten if any(u in open_uren for u in s["uren_beschikbaar"])]
attracties_te_plannen.sort(key=lambda a: kritieke_score(a,studenten_workend))

# Vul 'kan' lijst per student
for s in studenten_workend:
    s["kan"] = [a for a in s["attracties"] if a in attracties_te_plannen]

# -----------------------------
# Rode vakjes per uur berekenen
# -----------------------------
def maak_rode_vakjes(studenten, aantallen, open_uren):
    rode_vakjes = {h: set() for h in open_uren}
    for h in open_uren:
        actieve_students = [s for s in studenten if h in s["uren_beschikbaar"] and not s["is_pauzevlinder"]]
        n_students = len(actieve_students)
        attracties_2_posities = [a for a in aantallen if aantallen[a] >= 2]
        n_attracties = len(attracties_2_posities)
        if n_students <= n_attracties:
            for a in attracties_2_posities:
                rode_vakjes[h].add((a,2))
    return rode_vakjes

rode_vakjes = maak_rode_vakjes(studenten_workend, aantallen, open_uren)

# -----------------------------
# Initialiseer toewijzingsstructuren
# -----------------------------
assigned_map = {}  # (uur, attractie, positie) -> student
per_hour_assigned_counts = {uur: {a: 0 for a in aantallen} for uur in open_uren}
extra_assignments = {uur: [] for uur in open_uren}

# -----------------------------
# CONSTANTEN
# -----------------------------
MAX_CONSEC = 4
BLOK_VOLGORDE = [3,4,2,1]

# -----------------------------
# Functie om student in te plannen
# -----------------------------
def assign_student(student):
    # Sorteer uren op beschikbaarheid
    uren_lijst = sorted(student["uren_beschikbaar"])
    runs = contiguous_runs(uren_lijst)
    for run in runs:
        run_len = len(run)
        blokken = partition_run_lengths(run_len)
        start_idx = 0
        for b in blokken:
            uren_blok = run[start_idx:start_idx+b]
            start_idx += b
            placed=False
            # Probeer attracties van de student
            for attr in student["kan"]:
                ruimte = True
                for u in uren_blok:
                    pos = per_hour_assigned_counts[u][attr] + 1
                    if pos > aantallen[attr] or (attr,pos) in rode_vakjes[u]:
                        ruimte = False
                        break
                if not ruimte:
                    continue
                # Plaats student
                for u in uren_blok:
                    pos = per_hour_assigned_counts[u][attr] + 1
                    assigned_map[(u,attr,pos)] = student["naam"]
                    per_hour_assigned_counts[u][attr] +=1
                    student["assigned_hours"].append(u)
                student["assigned_attracties"].add(attr)
                placed=True
                break
            if not placed:
                # Zet uren bij extra
                for u in uren_blok:
                    extra_assignments[u].append(student["naam"])
                    student["assigned_hours"].append(u)

# -----------------------------
# Alle studenten inplannen
# -----------------------------
for s in sorted(studenten_workend, key=lambda x: x["aantal_attracties"]):
    assign_student(s)





# -----------------------------
# Excel output
# -----------------------------
wb_out=Workbook()
ws_out=wb_out.active
ws_out.title="Planning"

header_fill=PatternFill(start_color="BDD7EE",fill_type="solid")
attr_fill=PatternFill(start_color="E2EFDA",fill_type="solid")
pv_fill=PatternFill(start_color="FFF2CC",fill_type="solid")
extra_fill=PatternFill(start_color="FCE4D6",fill_type="solid")
center_align=Alignment(horizontal="center",vertical="center")
thin_border=Border(left=Side(style="thin"),right=Side(style="thin"),
                   top=Side(style="thin"),bottom=Side(style="thin"))

# Header
ws_out.cell(1,1,vandaag).font=Font(bold=True)
for col_idx,uur in enumerate(sorted(open_uren),start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out=2
for attr in attracties_te_plannen:
    max_pos=max(aantallen.get(attr,1), max(per_hour_assigned_counts[h].get(attr,0) for h in open_uren))
    for pos_idx in range(1,max_pos+1):
        naam_attr=attr if max_pos==1 else f"{attr} {pos_idx}"
        ws_out.cell(rij_out,1,naam_attr).font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=attr_fill
        ws_out.cell(rij_out,1).border=thin_border
        for col_idx,uur in enumerate(sorted(open_uren),start=2):
            naam=assigned_map.get((uur,attr),[])
            naam=naam[pos_idx-1] if pos_idx-1<len(naam) else ""
            ws_out.cell(rij_out,col_idx,naam).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out+=1

# Pauzevlinders
rij_out+=1
for pv_idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx,uur in enumerate(sorted(open_uren),start=2):
        ws_out.cell(rij_out,col_idx,pvnaam if uur in required_pauze_hours else "").alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Extra
rij_out+=1
ws_out.cell(rij_out,1,"Extra").font=Font(bold=True)
ws_out.cell(rij_out,1).fill=extra_fill
ws_out.cell(rij_out,1).border=thin_border

for col_idx,uur in enumerate(sorted(open_uren),start=2):
    for r_offset, naam in enumerate(extra_assignments[uur]):
        ws_out.cell(rij_out+1+r_offset,col_idx,naam).alignment=center_align

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=18

output=BytesIO()
wb_out.save(output)
output.seek(0)
st.success("Planning gegenereerd!")
st.download_button("Download planning",data=output.getvalue(),
                   file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
