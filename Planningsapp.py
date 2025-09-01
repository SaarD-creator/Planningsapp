import streamlit as st
import random
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import datetime

vandaag = datetime.date.today().strftime("%d-%m-%Y")
uploaded_file = st.file_uploader("Upload Excel bestand", type=["xlsx"])
if not uploaded_file:
    st.warning("Upload eerst een Excel-bestand om verder te gaan.")
    st.stop()

wb = load_workbook(BytesIO(uploaded_file.read()))
ws = wb["Blad1"]

def max_consecutive_hours(urenlijst):
    if not urenlijst: return 0
    urenlijst = sorted(set(urenlijst))
    maxr = huidig = 1
    for i in range(1,len(urenlijst)):
        huidig = huidig+1 if urenlijst[i]==urenlijst[i-1]+1 else 1
        maxr = max(maxr,huidig)
    return maxr

def partition_run_lengths(L):
    blocks = [3,2,4,1]
    dp = [(float('inf'), [])]*(L+1)
    dp[0] = (0,[])
    for i in range(1,L+1):
        best=(float('inf'),[])
        for b in blocks:
            if i-b<0: continue
            prev = dp[i-b]
            cand = (prev[0]+1, prev[1]+[b])
            if cand[0]<best[0]: best=cand
        dp[i]=best
    return dp[L][1]

def contiguous_runs(sorted_hours):
    runs=[]
    if not sorted_hours: return runs
    run=[sorted_hours[0]]
    for h in sorted_hours[1:]:
        if h==run[-1]+1: run.append(h)
        else: runs.append(run); run=[h]
    runs.append(run)
    return runs

# Studenten inlezen
studenten=[]
for rij in range(2,500):
    naam = ws.cell(rij,12).value
    if not naam: continue
    uren_beschikbaar=[10+(kol-3) for kol in range(3,12) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    attracties=[ws.cell(1,kol).value for kol in range(14,32) if ws.cell(rij,kol).value in [1,True,"WAAR","X"]]
    try:
        aantal_attracties=int(ws[f'AG{rij}'].value) if ws[f'AG{rij}'].value else len(attracties)
    except:
        aantal_attracties=len(attracties)
    studenten.append({"naam":naam,"uren_beschikbaar":sorted(uren_beschikbaar),"attracties":[a for a in attracties if a],"aantal_attracties":aantal_attracties,"is_pauzevlinder":False,"pv_number":None,"assigned_attracties":set(),"assigned_hours":[]})

# Openingsuren
open_uren=[int(str(ws.cell(1,kol).value).replace("u","").strip()) for kol in range(36,45) if ws.cell(2,kol).value in [1,True,"WAAR","X"]]
if not open_uren: open_uren=list(range(10,19))
open_uren=sorted(set(open_uren))

# Pauzevlinders
pauzevlinder_namen=[ws[f'BN{rij}'].value for rij in range(4,11) if ws[f'BN{rij}'].value]
def compute_pauze_hours(open_uren):
    if 10 in open_uren and 18 in open_uren: return [h for h in open_uren if 12<=h<=17]
    elif 12 in open_uren and 18 in open_uren: return [h for h in open_uren if 13<=h<=18]
    elif min(open_uren)>=14: return list(open_uren)
    else: return [h for h in open_uren if 12<=h<=17]
required_pauze_hours=compute_pauze_hours(open_uren)
for idx,pvnaam in enumerate(pauzevlinder_namen,start=1):
    if not pvnaam: continue
    for s in studenten:
        if s["naam"]==pvnaam:
            s["is_pauzevlinder"]=True
            s["pv_number"]=idx
            s["uren_beschikbaar"]=[u for u in s["uren_beschikbaar"] if u not in required_pauze_hours]
            break

# Attracties & aantallen
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

def kritieke_score(attr): return sum(1 for s in studenten if attr in s["attracties"])
attracties_te_plannen.sort(key=kritieke_score)

# -----------------------------
# Bereken per uur beschikbare studenten
beschikbaar_per_uur = {u:[s for s in studenten if u in s["uren_beschikbaar"] and not s["is_pauzevlinder"]] for u in open_uren}

# -----------------------------
# Per student planning
assigned_map = defaultdict(list)
per_hour_assigned_counts = {uur:{a:0 for a in attracties_te_plannen} for uur in open_uren}
MAX_CONSEC=4
MAX_PER_STUDENT_ATTR=6
studenten_sorted=sorted(studenten,key=lambda s:s["aantal_attracties"])
extra_assignments=[]

for s in studenten_sorted:
    uren = [u for u in s["uren_beschikbaar"] if u in open_uren]
    if not uren: continue
    uren = sorted(uren)
    runs = contiguous_runs(uren)
    for run in runs:
        L = len(run)
        block_sizes = partition_run_lengths(L)
        start_idx = 0
        for b in block_sizes:
            block_hours = run[start_idx:start_idx+b]
            start_idx += b
            geplaatst=False
            for attr in attracties_te_plannen:
                if attr not in s["attracties"]: continue
                if any(s["naam"] in assigned_map[(h,attr)] for h in block_hours): continue
                # Check capaciteit op basis van beschikbaar studenten per uur
                ruimte = all(per_hour_assigned_counts[h][attr]<min(aantallen[attr],len(beschikbaar_per_uur[h])) for h in block_hours)
                if not ruimte: continue
                hypothetische = sorted(set(s["assigned_hours"] + block_hours))
                if max_consecutive_hours(hypothetische)>MAX_CONSEC: continue
                # Toewijzen
                for h in block_hours:
                    assigned_map[(h,attr)].append(s["naam"])
                    per_hour_assigned_counts[h][attr]+=1
                    s["assigned_hours"].append(h)
                s["assigned_attracties"].add(attr)
                geplaatst=True
                break
            if not geplaatst:
                extra_assignments.append({"naam":s["naam"],"uren":block_hours})

# -----------------------------
# Excel-output (zelfde layout als eerder)
wb_out=Workbook()
ws_out=wb_out.active
ws_out.title="Planning"
header_fill=PatternFill(start_color="BDD7EE",fill_type="solid")
attr_fill=PatternFill(start_color="E2EFDA",fill_type="solid")
pv_fill=PatternFill(start_color="FFF2CC",fill_type="solid")
extra_fill=PatternFill(start_color="FCE4D6",fill_type="solid")
center_align=Alignment(horizontal="center",vertical="center")
thin_border=Border(left=Side(style="thin"),right=Side(style
