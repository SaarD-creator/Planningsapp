import streamlit as st
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime

vandaag = datetime.date.today().strftime("%d-%m-%Y")

# -----------------------------
# Upload / lees Excel
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

def partition_run_lengths(L):
    blocks = [3,2,4,1]
    big = 10**9
    dp = [(big, big, big, [])]*(L+1)
    dp[0] = (0,0,0,[])
    for i in range(1,L+1):
        best = (big,big,big,[])
        for b in blocks:
            if i-b < 0: continue
            prev = dp[i-b]
            num_ones = prev[0] + (1 if b==1 else 0)
            num_32 = prev[1] + (-1 if b in (2,3) else 0)
            num_blocks = prev[2] + 1
            part = prev[3] + [b]
            cand = (num_ones,num_32,num_blocks,part)
            if (cand[0],cand[1],cand[2]) < (best[0],best[1],best[2]):
                best = cand
        dp[i] = best
    return dp[L][3]

def contiguous_runs(sorted_hours):
    runs=[]
    if not sorted_hours: return runs
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
    if not naam: continue
    uren_beschikbaar = [10+(kol-3) for kol in range(3,12) if ws.cell(rij,kol).value in [1,True,'WAAR','X']]
    attracties = [ws.cell(1,kol).value for kol in range(14,32) if ws.cell(rij,kol).value in [1,True,'WAAR','X']]
    try:
        aantal_attracties = int(ws[f'AG{rij}'].value) if ws[f'AG{rij}'].value else len(attracties)
    except:
        aantal_attracties = len(attracties)
    studenten.append({
        'naam': naam,
        'uren_beschikbaar': sorted(uren_beschikbaar),
        'attracties': [a for a in attracties if a],
        'aantal_attracties': aantal_attracties,
        'is_pauzevlinder': False,
        'pv_number': None,
        'assigned_attracties': set(),
        'assigned_hours': []
    })

# -----------------------------
# Openingsuren
# -----------------------------
open_uren = [int(str(ws.cell(1,kol).value).replace('u','').strip()) for kol in range(36,45) if ws.cell(2,kol).value in [1,True,'WAAR','X']]
if not open_uren: open_uren=list(range(10,19))
open_uren = sorted(set(open_uren))

# -----------------------------
# Pauzevlinders
# -----------------------------
pauzevlinder_namen = [ws[f'BN{rij}'].value for rij in range(4,11) if ws[f'BN{rij}'].value]

def compute_pauze_hours(open_uren):
    if 10 in open_uren and 18 in open_uren:
        return [h for h in open_uren if 12 <= h <= 17]
    elif 12 in open_uren and 18 in open_uren:
        return [h for h in open_uren if 13 <= h <= 18]
    elif min(open_uren) >= 14:
        return list(open_uren)
    else:
        return [h for h in open_uren if 12 <= h <= 17]

required_pauze_hours = compute_pauze_hours(open_uren)

for idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    if not pvnaam: continue
    for s in studenten:
        if s['naam']==pvnaam:
            s['is_pauzevlinder']=True
            s['pv_number']=idx
            s['uren_beschikbaar'] = [u for u in s['uren_beschikbaar'] if u not in required_pauze_hours]
            break

# -----------------------------
# Attracties & aantallen
# -----------------------------
aantallen = {}
attracties_te_plannen=[]
for kol in range(47,65):
    naam = ws.cell(1,kol).value
    if naam:
        try: aantal=int(ws.cell(2,kol).value)
        except: aantal=0
        aantallen[naam]=max(0,min(2,aantal))
        if aantallen[naam]>=1: attracties_te_plannen.append(naam)

def kritieke_score(attr, studenten_list):
    return sum(1 for s in studenten_list if attr in s['attracties'])

studenten_workend = [s for s in studenten if any(u in open_uren for u in s['uren_beschikbaar'])]
attracties_te_plannen.sort(key=lambda a: kritieke_score(a, studenten_workend))

# -----------------------------
# Tweede posities volgorde uit kolom BA rij 5-11
# -----------------------------
tweede_pos_volgorde = [ws[f'BA{rij}'].value for rij in range(5,12) if ws[f'BA{rij}'].value]

# -----------------------------
# Planning per student
# -----------------------------
assigned_map = defaultdict(list)  # (uur, attr) -> list of student-names
per_hour_assigned_counts = {uur: {a:0 for a in attracties_te_plannen} for uur in open_uren}
MAX_CONSEC = 4
MAX_PER_STUDENT_ATTR = 6
extra_assignments = defaultdict(list)

studenten_sorted = sorted(studenten_workend, key=lambda s: s['aantal_attracties'])

def assign_student(s):
    uren = [u for u in s['uren_beschikbaar'] if u in open_uren and u not in s['assigned_hours']]
    if not uren: return
    uren = sorted(uren)
    runs = contiguous_runs(uren)
    for run in runs:
        L = len(run)
        if L==0: continue
        block_sizes = partition_run_lengths(L)
        start_idx = 0
        for b in block_sizes:
            block_hours = run[start_idx:start_idx+b]
            start_idx += b
            assigned=False
            # probeer attracties in volgorde
            candidate_attrs = [a for a in attracties_te_plannen if a in s['attracties'] and a not in s['assigned_attracties']]
            for attr in candidate_attrs:
                ruimte = True
                for h in block_hours:
                    allowed = aantallen.get(attr,1)
                    used = per_hour_assigned_counts[h].get(attr,0)
                    if used >= allowed:
                        ruimte=False
                        break
                hypothetisch = sorted(set(s['assigned_hours'] + block_hours))
                if max_consecutive_hours(hypothetisch) > MAX_CONSEC:
                    ruimte=False
                if not ruimte: continue
                # alles ok -> toewijzen
                for h in block_hours:
                    assigned_map[(h,attr)].append(s['naam'])
                    per_hour_assigned_counts[h][attr]+=1
                    s['assigned_hours'].append(h)
                s['assigned_attracties'].add(attr)
                assigned=True
                break
            if not assigned:
                # fallback: zet in extra
                for h in block_hours:
                    extra_assignments[h].append(s['naam'])

for s in studenten_sorted:
    assign_student(s)

# -----------------------------
# Excel output
# -----------------------------
wb_out = Workbook()
ws_out = wb_out.active
ws_out.title = 'Planning'

header_fill = PatternFill(start_color='BDD7EE',fill_type='solid')
attr_fill = PatternFill(start_color='E2EFDA',fill_type='solid')
pv_fill = PatternFill(start_color='FFF2CC',fill_type='solid')
extra_fill = PatternFill(start_color='FCE4D6',fill_type='solid')
center_align = Alignment(horizontal='center',vertical='center')
thin_border=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))

# Header
ws_out.cell(1,1,vandaag).font=Font(bold=True)
for col_idx,uur in enumerate(sorted(open_uren),start=2):
    ws_out.cell(1,col_idx,f'{uur}:00').font=Font(bold=True)
    ws_out.cell(1,col_idx).fill=header_fill
    ws_out.cell(1,col_idx).alignment=center_align
    ws_out.cell(1,col_idx).border=thin_border

rij_out = 2
for attr in attracties_te_plannen:
    max_pos = max(aantallen.get(attr,1), max(per_hour_assigned_counts[h].get(attr,0) for h in open_uren))
    for pos_idx in range(1,max_pos+1):
        naam_attr = attr if max_pos==1 else f'{attr} {pos_idx}'
        ws_out.cell(rij_out,1,naam_attr).font=Font(bold=True)
        ws_out.cell(rij_out,1).fill=attr_fill
        ws_out.cell(rij_out,1).border=thin_border
        for col_idx,uur in enumerate(sorted(open_uren),start=2):
            assigned_list = assigned_map.get((uur,attr),[])
            naam = assigned_list[pos_idx-1] if pos_idx-1<len(assigned_list) else ''
            ws_out.cell(rij_out,col_idx,naam).alignment=center_align
            ws_out.cell(rij_out,col_idx).border=thin_border
        rij_out+=1

# Pauzevlinders
rij_out+=1
for pv_idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    if not pvnaam: continue
    ws_out.cell(rij_out,1,f'Pauzevlinder {pv_idx}').font=Font(bold=True)
    ws_out.cell(rij_out,1).fill=pv_fill
    ws_out.cell(rij_out,1).border=thin_border
    for col_idx,uur in enumerate(sorted(open_uren),start=2):
        ws_out.cell(rij_out,col_idx,pvnaam if uur in required_pauze_hours else '').alignment=center_align
        ws_out.cell(rij_out,col_idx).border=thin_border
    rij_out+=1

# Extra studenten
rij_out+=1
ws_out.cell(rij_out,1,'Extra').font=Font(bold=True)
ws_out.cell(rij_out,1).fill=extra_fill
ws_out.cell(rij_out,1).border=thin_border
for col_idx, uur in enumerate(sorted(open_uren),start=2):
    for r_offset,n in enumerate(extra_assignments.get(uur,[])):
        ws_out.cell(rij_out+1+r_offset,col_idx,n).alignment=center_align

# Kolombreedte
for col in range(1,len(open_uren)+2):
    ws_out.column_dimensions[get_column_letter(col)].width=18

output = BytesIO()
wb_out.save(output)
output.seek(0)
st.success('Planning gegenereerd!')
st.download_button('Download planning', data=output.getvalue(), file_name=f'Planning_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
