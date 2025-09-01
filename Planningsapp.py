import streamlit as st
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
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

wb = load_workbook(BytesIO(uploaded_file.read()))
ws = wb["Blad1"]

# -----------------------------
# Hulpfuncties
# -----------------------------
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
    """Blokken volgorde 3-4-2-1"""
    blocks = [3,4,2,1]
    dp = [(10**9, [])]*(L+1)
    dp[0] = (0, [])
    for i in range(1,L+1):
        best = (10**9, [])
        for b in blocks:
            if i-b < 0:
                continue
            prev_ones, prev_blocks = dp[i-b]
            ones = prev_ones + (1 if b==1 else 0)
            if ones < best[0]:
                best = (ones, prev_blocks+[b])
        dp[i] = best
    return dp[L][1]

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
        "assigned_hours":[]
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

# -----------------------------
# Bereken taboe/rode vakjes per uur (samengevoegd)
# -----------------------------
rode_vakjes_per_uur = defaultdict(set)
for uur in open_uren:
    beschikbaar = sum(1 for s in studenten_workend if uur in s["uren_beschikbaar"] and not s["is_pauzevlinder"])
    benodigd = 0
    # tel voor elke attractie hoeveel studenten minimaal nodig zijn (1 voor enkel, 2 als max>=2)
    for attr in attracties_te_plannen:
        plekken = aantallen.get(attr,1)
        benodigd += min(2, plekken)
    # als er niet genoeg studenten zijn voor alle benodigde plekken, markeer tweede posities als taboe
    if beschikbaar < benodigd:
        for attr in attracties_te_plannen:
            if aantallen.get(attr,0) >= 2:
                rode_vakjes_per_uur[uur].add(attr)

# -----------------------------
# Assign studenten
# -----------------------------
assigned_map = defaultdict(list)  # (uur, attr) -> lijst van student-namen
per_hour_assigned_counts = {uur:{a:0 for a in attracties_te_plannen} for uur in open_uren}
extra_assignments = defaultdict(list)
studenten_sorted = sorted(studenten_workend, key=lambda s:s["aantal_attracties"])

def assign_student(s):
    uren = [u for u in s["uren_beschikbaar"] if u in open_uren]
    runs = contiguous_runs(uren)
    for run in runs:
        L = len(run)
        if L == 0: continue
        blocks = partition_run_lengths(L)  # blokken 3-4-2-1
        start_idx = 0
        for b in blocks:
            block_hours = run[start_idx:start_idx+b]
            start_idx += b
            placed = False
            # probeer attracties in kritieke volgorde
            for attr in attracties_te_plannen:
                if attr not in s["attracties"]: continue
                if attr in s["assigned_attracties"]: continue
                # check of attractie in taboe-uren zit
                if any(attr in rode_vakjes_per_uur.get(h,set()) for h in block_hours):
                    continue
                # check of er nog plek is
                ruimte = True
                for h in block_hours:
                    if per_hour_assigned_counts[h][attr] >= aantallen.get(attr,1):
                        ruimte = False
                        break
                if ruimte:
                    for h in block_hours:
                        assigned_map[(h,attr)].append(s["naam"])
                        per_hour_assigned_counts[h][attr] += 1
                        s["assigned_hours"].append(h)
                    s["assigned_attracties"].add(attr)
                    placed = True
                    break
            # als nog steeds niet geplaatst, extra assignments op vrije, niet-taboe plekken
            if not placed:
                for h in block_hours:
                    vrije_attr = [a for a in attracties_te_plannen if a not in rode_vakjes_per_uur.get(h,set())]
                    if vrije_attr:
                        extra_assignments[h].append(s["naam"])

# studenten toewijzen
for s in studenten_sorted:
    assign_student(s)




from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = "Planning"

# Styles
header_fill = PatternFill(start_color="BDD7EE", fill_type="solid")
attr_fill   = PatternFill(start_color="E2EFDA", fill_type="solid")
pv_fill     = PatternFill(start_color="FFF2CC", fill_type="solid")
extra_fill  = PatternFill(start_color="FCE4D6", fill_type="solid")
rood_fill   = PatternFill(start_color="FFC7CE", fill_type="solid")
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))

# Header
ws_out.cell(1,1,vandaag).font=Font(bold=True)
for col_idx, uur in enumerate(sorted(open_uren), start=2):
    cel = ws_out.cell(1,col_idx, f"{uur}:00")
    cel.font = Font(bold=True)
    cel.fill = header_fill
    cel.alignment = center_align
    cel.border = thin_border

rij_out = 2
# Attracties invullen
for attr in attracties_te_plannen:
    max_pos = max(aantallen.get(attr,1), max(per_hour_assigned_counts[h].get(attr,0) for h in open_uren))
    for pos_idx in range(1, max_pos+1):
        naam_attr = attr if max_pos==1 else f"{attr} {pos_idx}"
        ws_out.cell(rij_out,1,naam_attr).font = Font(bold=True)
        ws_out.cell(rij_out,1).fill = attr_fill
        ws_out.cell(rij_out,1).border = thin_border
        for col_idx, uur in enumerate(sorted(open_uren), start=2):
            cel = ws_out.cell(rij_out,col_idx)
            # Rode vakjes respecteren
            if pos_idx == 2 and attr in rode_vakjes_per_uur.get(uur,set()):
                cel.value = ""  # geen student
                cel.fill = rood_fill
            else:
                namen = assigned_map.get((uur,attr),[])
                cel.value = namen[pos_idx-1] if pos_idx-1<len(namen) else ""
            cel.alignment = center_align
            cel.border = thin_border
        rij_out += 1

# Pauzevlinders
rij_out += 1
for pv_idx, pvnaam in enumerate(pauzevlinder_namen,start=1):
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font = Font(bold=True)
    ws_out.cell(rij_out,1).fill = pv_fill
    ws_out.cell(rij_out,1).border = thin_border
    for col_idx, uur in enumerate(sorted(open_uren), start=2):
        ws_out.cell(rij_out,col_idx, pvnaam if uur in required_pauze_hours else "").alignment = center_align
        ws_out.cell(rij_out,col_idx).border = thin_border
    rij_out += 1

# Extra
rij_out += 1
ws_out.cell(rij_out,1,"Extra").font = Font(bold=True)
ws_out.cell(rij_out,1).fill = extra_fill
ws_out.cell(rij_out,1).border = thin_border

for col_idx, uur in enumerate(sorted(open_uren), start=2):
    for r_offset, naam in enumerate(extra_assignments[uur]):
        ws_out.cell(rij_out+1+r_offset, col_idx, naam).alignment = center_align

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=18

# Output voor Streamlit
output = BytesIO()
wb_out.save(output)
output.seek(0)
st.success("Planning gegenereerd!")
st.download_button("Download planning", data=output.getvalue(),
                   file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
