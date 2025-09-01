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

wb = load_workbook(BytesIO(uploaded_file.read()))
ws = wb["Blad1"]

# -----------------------------
# Hulpfuncties
# -----------------------------
def max_consecutive_hours(urenlijst):
    if not urenlijst:
        return 0
    urenlijst = sorted(set(urenlijst))
    maxr = huidig = 1
    for i in range(1, len(urenlijst)):
        huidig = huidig + 1 if urenlijst[i] == urenlijst[i-1] + 1 else 1
        maxr = max(maxr, huidig)
    return maxr

def flexible_partition(L):
    """Verdeel L uren in flexibele blokken: voorkeur 3>2>4>1 uur, zo min mogelijk 1-uur blokken."""
    blocks = []
    while L > 0:
        if L >= 3:
            blocks.append(3)
            L -= 3
        elif L == 2:
            blocks.append(2)
            L -= 2
        elif L == 1:
            blocks.append(1)
            L -= 1
        elif L == 4:  # soms beter als 4-uur blok
            blocks.append(4)
            L -= 4
    return blocks

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
        "attracties": [a for a in attracties if a],
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
        return [h for h in open_uren if 12 <= h <= 16]
    elif min(open_uren) >= 12:
        return list(open_uren)
    else:
        return [h for h in open_uren if 12 <= h <= 17]

required_pauze_hours=compute_pauze_hours(open_uren)

for idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    if not pvnaam: continue
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
# Planning initiëren
# -----------------------------
assigned_map = defaultdict(list)  # (uur, attr) -> lijst van student-namen
per_hour_assigned_counts = {uur: {a:0 for a in attracties_te_plannen} for uur in open_uren}
MAX_CONSEC = 4
MAX_PER_STUDENT_ATTR = 6
extra_assignments = defaultdict(list)

# Sorteer studenten: eerst wie minst attracties kan
studenten_sorted = sorted(studenten_workend, key=lambda s:s["aantal_attracties"])

def assign_student_block(s, block_hours):
    uur0 = block_hours[0]
    candidate_attrs = [a for a in attracties_te_plannen if a in s["attracties"] and a not in s["assigned_attracties"]]
    for attr in candidate_attrs:
        ruimte = True
        for h in block_hours:
            if per_hour_assigned_counts[h][attr] >= aantallen[attr]:
                ruimte = False
                break
        if not ruimte:
            continue
        hypothetische = sorted(set(s["assigned_hours"] + block_hours))
        if max_consecutive_hours(hypothetische) > MAX_CONSEC:
            continue
        # toewijzen
        for h in block_hours:
            assigned_map[(h,attr)].append(s["naam"])
            per_hour_assigned_counts[h][attr] +=1
            s["assigned_hours"].append(h)
        s["assigned_attracties"].add(attr)
        return True
    return False

# -----------------------------
# Vul de planning student per student
# -----------------------------
for s in studenten_sorted:
    uren = [u for u in s["uren_beschikbaar"] if u in open_uren and u not in s["assigned_hours"]]
    if not uren:
        continue
    runs = contiguous_runs(sorted(uren))
    for run in runs:
        L = len(run)
        if L==0: continue
        blocks = flexible_partition(L)
        start_idx = 0
        for b in blocks:
            block_hours = run[start_idx:start_idx+b]
            start_idx += b
            assigned = assign_student_block(s, block_hours)
            if not assigned:
                # Als het echt niet kan, zet bij extra
                for h in block_hours:
                    extra_assignments[h].append(s["naam"])

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

# header
ws_out.cell(1,1,vandaag).font=Font(bold=True)
for col_idx,uur in enumerate(sorted(open_uren),start=2):
    ws_out.cell(1,col_idx,f"{uur}:00").font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out=2
# normale attracties
for attr in attracties_te_plannen:
    max_pos=max(per_hour_assigned_counts[h].get(attr,0) for h in open_uren)
    max_pos=max(max_pos,aantallen.get(attr,1))
    for pos_idx in range(1,max_pos+1):
        naam_attr=attr if max_pos==1 else f"{attr} {pos_idx}"
        ws_out.cell(rij_out,1,naam_attr).font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=attr_fill
        ws_out.cell(rij_out,1).border=thin_border
        for col_idx,uur in enumerate(sorted(open_uren),start=2):
            assigned_list=assigned_map.get((uur,attr),[])
            naam=assigned_list[pos_idx-1] if pos_idx-1<len(assigned_list) else ""
            ws_out.cell(rij_out,col_idx,naam).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out+=1

# pauzevlinders
rij_out+=1
for pv_idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    if not pvnaam: continue
    ws_out.cell(rij_out,1,f"Pauzevlinder {pv_idx}").font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx,uur in enumerate(sorted(open_uren),start=2):
        ws_out.cell(rij_out,col_idx,pvnaam if uur in required_pauze_hours else "").alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# extra-studenten
rij_out+=1
for uur in sorted(open_uren):
    for naam in extra_assignments.get(uur,[]):
        ws_out.cell(rij_out,1,"Extra").font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=extra_fill
        ws_out.cell(rij_out,1).border=thin_border
        ws_out.cell(rij_out,sorted(open_uren).index(uur)+2,naam).alignment=center_align
        ws_out.cell(rij_out,sorted(open_uren).index(uur)+2).border=thin_border
        rij_out+=1

# kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=18

output=BytesIO()
wb_out.save(output)
output.seek(0)
st.success("Planning gegenereerd!")
st.download_button("Download planning",data=output.getvalue(),
                   file_name=f"Planning_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
