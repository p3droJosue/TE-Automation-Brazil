"""Compare the Brazilian DB exports against the golden templates, line by line.

Run from anywhere:  python analysis/compare_db_vs_golden.py
Prints, for every template block, how many DB loss categories land on the
expected cell with the exact golden value, plus a production check.
The MAP / CAT tables are the working hypothesis documented in
analysis/BRAZIL_ANALYSIS.md.
"""
import openpyxl, os
R=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def load_pivot(fn):
    ws=openpyxl.load_workbook(f"{R}/databases/{fn}",data_only=True).active
    rows=list(ws.iter_rows(values_only=True)); lines=rows[0][1:]
    d={}
    for row in rows[2:]:
        if row[0] is None or str(row[0]).startswith("Applied") or row[0]=="Total": continue
        for ln,v in zip(lines,row[1:]):
            if v is not None: d.setdefault(ln,{})[row[0].strip()]=v
    return d
APP=load_pivot("Detalhe True Efficiency - APP.xlsx"); ITA=load_pivot("Detalhe True Efficiency - ITA.xlsx"); MES=load_pivot("Detalle True Efficiency  - MES.xlsx")
# category -> (col,row) in DEFAULT layout
CAT={
 "Dark hours: line available to run if needed":("C",19),"Dark hrs: line available to run if needed":("C",19),
 "Planned maintenance hours":("C",22),"Planned sanitation hours":("C",23),"Planned startup & shutdown hours":("C",24),
 "Planned changeovers hours":("C",25),"Planned network outages hours":("C",26),"Team meetings hours":("C",27),
 "Other scheduled downtime hours":("C",28),"Maintenance overrun hours":("C",30),"Sanitation overrun hours":("C",31),
 "Changeover overrun hours":("C",32),"Raw material shortages hours":("C",35),"Quality issues hours":("C",36),
 "Other hours":("C",39),"Other Hours":("C",39),
 "Dark hour Loss Tree - RM shortage":("G",20),"Dark hour Loss Tree - FG warehouse":("G",21),
 "Dark hour Loss Tree - Holiday & Religious reasons":("G",22),"Dark hour Loss Tree - CAPEX & Test (R&D)":("G",23),
 "Dark hour Loss Tree - Staffling":("G",24),"Dark hour Loss Tree - Inventory":("G",27),"Dark hour Loss Tree - Other":("G",28),
 "Dark hour Loss Tree - Corn Cook Hours":("G",29),  # TC only (row 30 after +1)
 "Processing Breakdown - Technician Failure":("G",32),"Processing Breakdown - Operational Failure":("G",33),
 "Packaging Breakdown - Technician Failure":("G",36),"Packaging Breakdown - Operational Failure":("G",37),
 "Utilities - Internal":("G",40),"Utilities - External":("G",41),
 # MES-only coarse buckets (guesses)
 "Packaging Breakdown":("G",36),"Utilities hours":("G",41),"Processing Breakdown":("G",32),
}
OFF={"DEFAULT":0,"TC":1,"PC":10}
# (plant file key, sheet, cons, source, db line)
MAP=[
 ("Amacoco","Other",0,APP,"TBA19"),("Amacoco","Other",1,APP,"TBA21"),("Amacoco","Other",2,APP,"TBA22"),
 ("CIC","PC",0,APP,"PC60"),("CIC","TC",0,APP,"DORITOS_TC2"),("CIC","Extruded",0,APP,"CLEXTRAL III"),("CIC","Extruded",1,APP,"CLEXTRAL VII"),
 ("CIC","Other",0,APP,"STIKSY"),("CIC","Other",1,APP,"PINGO CIC"),("CIC","Other",2,APP,"BACON CIC"),("CIC","Other",3,APP,"POPCORNERS"),
 ("Feira","Other",0,APP,"BOSCH"),("Feira","Other",1,APP,"PLF"),("Feira","Other",2,APP,"MASIPACK"),("Feira","Other",3,APP,"SIG 1"),("Feira","Other",4,APP,"SIG 2"),
 ("ITU","PC",0,MES,"PC50"),("ITU","PC",1,APP,"KETTLE"),("ITU","TC",0,MES,"DORITOS_TC1"),("ITU","Extruded",0,MES,"CLEXTRAL I"),("ITU","Extruded",1,MES,"CLEXTRAL II"),
 ("ITU","Other",0,APP,"CRUNCHY ITU"),("ITU","Other",1,APP,"PINGO ITU"),
 ("Lucky","Extruded",0,ITA,"FOFURA"),("Lucky","Other",0,ITA,"TORCIDA"),
 ("Recife","Extruded",0,APP,"CLEXTRAL VI"),("Recife","Extruded",1,APP,"SCHAAF_1"),("Recife","Extruded",2,APP,"SCHAAF_2"),("Recife","Other",0,APP,"TORCIDA REC"),
 ("Sete","PC",0,MES,"PC37"),("Sete","TC",0,MES,"Doritos 3"),("Sete","TC",1,MES,"Doritos 4"),("Sete","Extruded",0,MES,"CLX IV"),("Sete","Extruded",1,MES,"CLX V"),
 ("Sorocaba","Extruded",0,APP,"SCHAAF_3"),("Sorocaba","Other",0,APP,"TODDY_2"),("Sorocaba","Other",1,APP,"TODDY_1"),("Sorocaba","Other",2,APP,"TODDY_3"),
 ("Sorocaba","Other",3,APP,"TODDY_4"),("Sorocaba","Other",4,APP,"SPEED_2"),("Sorocaba","Other",5,APP,"SPEED_3"),("Sorocaba","Other",6,APP,"SPEED_4"),
 ("Sorocaba","Stax",0,APP,"CANISTER_1"),
]
from openpyxl.utils import column_index_from_string as CI
G=f"{R}/templates/TE_completed_templates"; files=os.listdir(G); cache={}
tot_exact=tot=0
for key,sheet,cons,src,line in MAP:
    fn=[f for f in files if key in f][0]
    if fn not in cache: cache[fn]=openpyxl.load_workbook(f"{G}/{fn}",data_only=True)
    ws=cache[fn][sheet]; kind="PC" if sheet=="PC" else "TC" if sheet=="TC" else "DEFAULT"
    srcname="APP" if src is APP else "ITA" if src is ITA else "MES"
    hdr=ws.cell(6 if kind=="PC" else 4, 2+8*cons).value
    ex=0; n=0; bad=[]; used=set()
    for cat,v in src.get(line,{}).items():
        if cat not in CAT: 
            if cat not in ("Productive time",): bad.append(f"[unmapped {cat}={v:.4g}]")
            continue
        col,row=CAT[cat]; r=row+OFF[kind]; c=CI(col)+8*cons; used.add((r,c))
        gv=ws.cell(r,c).value; gv=gv if isinstance(gv,(int,float)) else 0
        n+=1
        if abs(gv-v)<1e-3: ex+=1
        else: bad.append(f"{cat[:28]}:{v:.4g}->{gv:.4g}")
    tot+=n; tot_exact+=ex
    print(f"{key:9}{sheet:9}{cons} {hdr!s:20} {srcname} {line:12} exact {ex}/{n} | "+"; ".join(bad))
print("TOTAL exact", tot_exact, "/", tot)
print("\n=== production")
ws=openpyxl.load_workbook(f"{R}/databases/Detalle True Efficiency.xlsx",data_only=True).active
prod={}
for row in ws.iter_rows(min_row=2,values_only=True):
    if row[1] and isinstance(row[4],(int,float)): prod.setdefault(row[1],[]).append((row[2],row[3],row[4]))
for key,sheet,cons,src,line in MAP:
    fn=[f for f in files if key in f][0]; ws=cache[fn][sheet]
    r0,c0=(88,3) if sheet=="PC" else (69,3) if sheet=="TC" else (60,3)
    gv=[ws.cell(r0+i,c0+8*cons).value for i in range(30 if sheet=="PC" else 20 if sheet=="TC" else 12)]
    gv=[v for v in gv if isinstance(v,(int,float)) and v]
    dv=[p[2] for p in prod.get(line,[])]
    m=sum(1 for d in dv if any(abs(d-g)<0.5 for g in gv))
    print(f"{key:9}{sheet:9}{cons} {line:12} DB n={len(dv)} sum={sum(dv):>11,.0f} | golden n={len(gv)} sum={sum(gv):>11,.0f} | DB rows matching a golden cell exactly: {m}")
