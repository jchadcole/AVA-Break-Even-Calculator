"""Regenerates ava-break-even-check.xlsx, the formula-based validation workbook for index.html.

Run: pip install openpyxl && python tools/build_workbook.py
Keep its inputs, defaults and formulas in step with the model in index.html.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
wb = Workbook()
ws = wb.active; ws.title = "Calculator"
blue = Font(color="0000FF"); bold = Font(bold=True); hdr = Font(bold=True, color="FFFFFF")
hfill = PatternFill("solid", fgColor="0F6E78"); infill = PatternFill("solid", fgColor="EAF4F5")
ws.column_dimensions["A"].width = 52; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 16; ws.column_dimensions["D"].width = 16; ws.column_dimensions["E"].width = 60
def head(r, text):
    ws.cell(r, 1, text).font = hdr
    for c in range(1, 6): ws.cell(r, c).fill = hfill
ws["A1"] = "AVA containment break-even: validation workbook"; ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Blue cells are inputs. Every other number is a formula you can inspect. Enter the same inputs as the web page and compare."
head(4, "Supplied values (monthly basis)")
inputs = [
 ("Inbound calls per month", 96000, "#,##0", "calls"),
 ("Fully loaded agent cost (per agent, per month)", 7000, "$#,##0", "cost"),
 ("Agent occupancy rate", 0.80, "0.0%", "occ"),
 ("Average call time (sec)", 600, "#,##0", "aht"),
 ("Paid hours per agent per month (one FTE)", 173, "#,##0.0", "hours"),
 ("% of agent time spent on these calls", 0.75, "0.0%", "share"),
 ("% of calls that go directly to an agent", 0.10, "0.0%", "direct"),
 ("% of calls contained in IVR", 0.20, "0.0%", "ivr"),
 ("% of calls contained in AVA (target, % of ALL calls)", 0.50, "0.0%", "target"),
 ("IVR cost per call (only calls that never reach AVA)", 0.50, "$0.00", "ivrcost"),
 ("AVA cost per call (replaces the IVR fee)", 1.20, "$0.00", "avacost"),
 ("Charge AVA on escalated calls too (1 = yes, 0 = no)", 1, "0", "charge"),
 ("Round FTEs up to whole people (1 = yes, 0 = fractional)", 0, "0", "whole"),
 ("AVA platform fee per month", 0, "$#,##0", "platform"),
 ("AVA implementation (one-time)", 0, "$#,##0", "impl"),
 ("Spread implementation over (months)", 36, "0", "months"),
]
R = {}
r = 5
for label, val, fmt, key in inputs:
    ws.cell(r, 1, label); c = ws.cell(r, 2, val); c.font = blue; c.fill = infill; c.number_format = fmt
    R[key] = f"$B${r}"; r += 1
r += 1
head(r, "Calculated values"); r += 1
calc = [
 ("Agent calls per hour", f"=3600/{R['aht']}*{R['occ']}", "0.00", "perhour", "3600 ÷ call time × occupancy"),
 ("Handle hours per FTE per month", f"={R['hours']}*{R['share']}*{R['occ']}", "#,##0.00", "fteh", "paid hours × % on these calls × occupancy"),
 ("Calls per FTE per month", f"=B{{fteh}}*3600/{R['aht']}", "#,##0.0", "peragent", "handle hours per FTE × 3600 ÷ call time"),
 ("Agent cost per call", f"={R['cost']}/B{{peragent}}", "$0.0000", "cpc", "agent cost ÷ calls per FTE per month"),
 ("% of calls reaching AVA", f"=1-{R['direct']}-{R['ivr']}", "0.0%", "reach", "100% − direct − IVR contained"),
 ("% IVR → AVA → Agent (at target)", f"=B{{reach}}-{R['target']}", "0.0%", "esc", "reaching AVA − AVA contained"),
 ("Check: four paths add to 100%", f"={R['direct']}+{R['ivr']}+{R['target']}+B{{esc}}", "0.0%", "sum", "must equal 100.0%"),
 ("AVA fixed cost per month", f"={R['platform']}+{R['impl']}/{R['months']}", "$#,##0", "fixed", "platform fee + implementation ÷ months"),
]
for label, f, fmt, key, note in calc:
    R[key] = r; r += 1
r -= len(calc)
for label, f, fmt, key, note in calc:
    ws.cell(r, 1, label); c = ws.cell(r, 2, f.format(**{k: v for k, v in R.items() if isinstance(v, int)})); c.number_format = fmt; ws.cell(r, 5, note).font = Font(italic=True, color="666666"); r += 1
r += 1
head(r, "Monthly cost comparison"); ws.cell(r, 2, "Without AVA").font = hdr; ws.cell(r, 3, "At target").font = hdr; ws.cell(r, 4, "Change").font = hdr; r += 1
def ag(hours_expr):
    return f"=IF({R['whole']}=1,ROUNDUP({hours_expr}/$B${R['fteh']}-0.000000001,0),{hours_expr}/$B${R['fteh']})"
rows = [
 ("Agent-handled calls", f"=({R['direct']}+$B${R['reach']})*{R['calls']}", f"=({R['direct']}+$B${R['esc']})*{R['calls']}", "#,##0", "agcalls"),
 ("Agent handle hours", None, None, "#,##0.0", "hh"),
 ("Agent FTEs required", None, None, "#,##0.00", "agents"),
 ("Agent labor", None, None, "$#,##0", "labor"),
 ("IVR cost", f"=(1-{R['direct']})*{R['calls']}*{R['ivrcost']}", f"={R['ivr']}*{R['calls']}*{R['ivrcost']}", "$#,##0", "ivrc"),
 ("AVA per-call cost", "=0", f"=IF({R['charge']}=1,$B${R['reach']},{R['target']})*{R['calls']}*{R['avacost']}", "$#,##0", "avac"),
 ("AVA fixed cost", "=0", f"=$B${R['fixed']}", "$#,##0", "fixc"),
 ("Total monthly cost", None, None, "$#,##0", "total"),
]
start = r
for i, (label, b, c, fmt, key) in enumerate(rows): R[key] = start + i
for i, (label, b, c, fmt, key) in enumerate(rows):
    rr = start + i; ws.cell(rr, 1, label)
    if key == "hh": b = f"=B{R['agcalls']}*{R['aht']}/3600"; c = f"=C{R['agcalls']}*{R['aht']}/3600"
    if key == "agents": b = ag(f"B{R['hh']}"); c = ag(f"C{R['hh']}")
    if key == "labor": b = f"=B{R['agents']}*{R['cost']}"; c = f"=C{R['agents']}*{R['cost']}"
    if key == "total": b = f"=SUM(B{R['labor']}:B{R['fixc']})"; c = f"=SUM(C{R['labor']}:C{R['fixc']})"
    ws.cell(rr, 2, b).number_format = fmt; ws.cell(rr, 3, c).number_format = fmt
    ws.cell(rr, 4, f"=C{rr}-B{rr}").number_format = fmt
    if key == "total":
        for col in range(1, 5): ws.cell(rr, col).font = bold
r = start + len(rows) + 1
ws.cell(r, 1, "Monthly savings at target").font = bold; ws.cell(r, 2, f"=B{R['total']}-C{R['total']}").number_format = "$#,##0"; R["sav"] = r; r += 1
ws.cell(r, 1, "Annual savings at target"); ws.cell(r, 2, f"=B{R['sav']}*12").number_format = "$#,##0"; r += 2
head(r, "Hours automated → FTEs saved → dollars (at target)"); r += 1
note = Font(italic=True, color="666666")
chain = [
 ("Agent hours automated per month", f"={R['target']}*{R['calls']}*{R['aht']}/3600", "#,##0.0", "hauto", "AVA-contained calls × call time ÷ 3600"),
 ("FTEs saved (fractional)", f"=B{{hauto}}/B{R['fteh']}", "#,##0.00", "ftefrac", "hours automated ÷ handle hours per FTE"),
 ("FTEs saved (as staffed)", f"=B{R['agents']}-C{R['agents']}", "#,##0.00", "ftesaved", "FTEs without AVA − FTEs at target (whole people if rounding is on)"),
 ("Labor saved", f"=B{{ftesaved}}*{R['cost']}", "$#,##0", "lsaved", "FTEs saved × fully loaded cost"),
 ("IVR fees avoided", f"=B{R['ivrc']}-C{R['ivrc']}", "$#,##0", "ivrsaved", "calls reaching AVA × IVR cost per call (they pay the AVA fee instead)"),
 ("AVA cost (per-call + fixed)", f"=C{R['avac']}+C{R['fixc']}", "$#,##0", "avatot", ""),
 ("Net monthly savings", f"=B{{lsaved}}+B{{ivrsaved}}-B{{avatot}}", "$#,##0", "net", "labor saved + IVR fees avoided − AVA cost; must equal monthly savings above"),
]
for label, f, fmt, key, n in chain: R[key] = r; r += 1
r -= len(chain)
for label, f, fmt, key, n in chain:
    ws.cell(r, 1, label); ws.cell(r, 2, f.format(**{k: v for k, v in R.items() if isinstance(v, int)})).number_format = fmt; ws.cell(r, 5, n).font = note; r += 1
r += 1
head(r, "Break-even AVA containment (% of all calls)"); r += 1
ws.cell(r, 1, "First 0.1% step at or above break-even (works with or without rounding)")
ws.cell(r, 2, "=_xlfn.MINIFS(Ladder!A:A,Ladder!H:H,\">=0\")").number_format = "0.0%"; ws.cell(r, 5, "lowest containment in the Ladder sheet with savings ≥ 0").font = Font(italic=True, color="666666"); R["be"] = r; r += 1
ws.cell(r, 1, "Same, as % of calls reaching AVA"); ws.cell(r, 2, f"=B{R['be']}/B{R['reach']}").number_format = "0.0%"; r += 1
ws.cell(r, 1, "Closed form, fractional agents (cross-check)")
ws.cell(r, 2, f"=MAX(0,IF({R['charge']}=1,(B{R['reach']}*{R['calls']}*({R['avacost']}-{R['ivrcost']})+B{R['fixed']})/({R['calls']}*B{R['cpc']}),(B{R['fixed']}-B{R['reach']}*{R['calls']}*{R['ivrcost']})/({R['calls']}*(B{R['cpc']}-{R['avacost']}))))").number_format = "0.00%"
ws.cell(r, 5, "charged: (AVA calls × AVA cost + fixed) ÷ (calls × agent cost per call); not charged: fixed ÷ (calls × (agent cost per call − AVA cost))").font = Font(italic=True, color="666666"); r += 1

# Ladder
L = wb.create_sheet("Ladder")
heads = ["AVA contained (% all calls)", "% of AVA calls", "IVR → AVA → Agent", "Agent calls", "FTEs needed", "AVA per-call cost", "Total monthly cost", "Monthly savings", "Agent hours automated", "FTEs saved"]
for i, h in enumerate(heads): c = L.cell(1, i + 1, h); c.font = hdr; c.fill = hfill; L.column_dimensions[get_column_letter(i + 1)].width = 20
C = "Calculator!"
reach = f"{C}$B${R['reach']}"
for k in range(0, 1001):
    rr = k + 2
    L.cell(rr, 1, f"=IF({k}/1000<={reach}+0.0000001,{k}/1000,\"\")").number_format = "0.0%"
    L.cell(rr, 2, f"=IF(A{rr}=\"\",\"\",A{rr}/{reach})").number_format = "0.0%"
    L.cell(rr, 3, f"=IF(A{rr}=\"\",\"\",{reach}-A{rr})").number_format = "0.0%"
    L.cell(rr, 4, f"=IF(A{rr}=\"\",\"\",({C}{R['direct']}+C{rr})*{C}{R['calls']})").number_format = "#,##0"
    L.cell(rr, 5, f"=IF(A{rr}=\"\",\"\",IF({C}{R['whole']}=1,ROUNDUP(D{rr}*{C}{R['aht']}/3600/{C}$B${R['fteh']}-0.000000001,0),D{rr}*{C}{R['aht']}/3600/{C}$B${R['fteh']}))").number_format = "#,##0.00"
    L.cell(rr, 9, f"=IF(A{rr}=\"\",\"\",A{rr}*{C}{R['calls']}*{C}{R['aht']}/3600)").number_format = "#,##0"
    L.cell(rr, 10, f"=IF(A{rr}=\"\",\"\",{C}$B${R['agents']}-E{rr})").number_format = "#,##0.00"
    L.cell(rr, 6, f"=IF(A{rr}=\"\",\"\",IF({C}{R['charge']}=1,{reach},A{rr})*{C}{R['calls']}*{C}{R['avacost']})").number_format = "$#,##0"
    L.cell(rr, 7, f"=IF(A{rr}=\"\",\"\",E{rr}*{C}{R['cost']}+{C}$C${R['ivrc']}+F{rr}+{C}$B${R['fixed']})").number_format = "$#,##0"
    L.cell(rr, 8, f"=IF(A{rr}=\"\",\"\",{C}$B${R['total']}-G{rr})").number_format = "$#,##0"
L.freeze_panes = "A2"

# Checks
K = wb.create_sheet("Checks")
for i, w in enumerate([58, 18, 18, 12]): K.column_dimensions[get_column_letter(i + 1)].width = w
for i, h in enumerate(["Sanity check", "Expected", "Workbook", "Pass?"]):
    c = K.cell(1, i + 1, h); c.font = hdr; c.fill = hfill
def lad(col, cexpr):  # lookup a ladder column at containment cexpr
    return f"INDEX(Ladder!{col}:{col},MATCH(ROUND({cexpr},3),Ladder!A:A,0))"
checks = [
 ("Four paths add to 100%", "=1", f"={C}B{R['sum']}", "0.0%"),
 ("At 0% AVA: FTEs equal FTEs without AVA", f"={C}B{R['agents']}", f"={lad('E', '0')}", "#,##0.0"),
 ("At 0% AVA: savings = IVR fees avoided − (AVA per-call cost + fixed) when AVA is charged on escalations", f"={reach}*{C}{R['calls']}*{C}{R['ivrcost']}-(IF({C}{R['charge']}=1,{reach}*{C}{R['calls']}*{C}{R['avacost']},0)+{C}B{R['fixed']})", f"={lad('H', '0')}", "$#,##0"),
 ("At break-even: savings are zero or slightly positive", "≥ 0", f"={lad('H', C + 'B' + str(R['be']))}", "$#,##0"),
 ("One step (0.1%) below break-even: savings are negative", "< 0", f"=IF({C}B{R['be']}=0,\"n/a\",{lad('H', C + 'B' + str(R['be']) + '-0.001')})", "$#,##0"),
 ("At maximum AVA (all calls reaching AVA contained): only direct calls need FTEs", f"=IF({C}{R['whole']}=1,ROUNDUP({C}{R['direct']}*{C}{R['calls']}*{C}{R['aht']}/3600/{C}B{R['fteh']}-0.000000001,0),{C}{R['direct']}*{C}{R['calls']}*{C}{R['aht']}/3600/{C}B{R['fteh']})", f"=INDEX(Ladder!E:E,MATCH(ROUND({reach},3),Ladder!A:A,0))", "#,##0.0"),
 ("Hours chain: labor saved + IVR avoided − AVA cost = total cost difference", f"={C}B{R['sav']}", f"={C}B{R['net']}", "$#,##0"),
 ("Fractional FTEs: FTEs saved = hours automated ÷ handle hours per FTE", f"={C}B{R['ftefrac']}", f"=IF({C}{R['whole']}=1,\"turn rounding off\",{C}B{R['ftesaved']})", "#,##0.00"),
 ("Fractional agents: labor = agent calls × agent cost per call", f"={C}B{R['agcalls']}*{C}B{R['cpc']}", f"=IF({C}{R['whole']}=1,\"turn rounding off\",{C}B{R['labor']})", "$#,##0"),
]
for i, (label, exp, got, fmt) in enumerate(checks):
    rr = i + 2; K.cell(rr, 1, label)
    K.cell(rr, 2, exp).number_format = fmt; K.cell(rr, 3, got).number_format = fmt
    if exp == "≥ 0": K.cell(rr, 4, f"=IF(C{rr}>=0,\"PASS\",\"FAIL\")")
    elif exp == "< 0": K.cell(rr, 4, f"=IF(C{rr}=\"n/a\",\"n/a\",IF(C{rr}<0,\"PASS\",\"FAIL\"))")
    elif "rounding" in label or "Fractional" in label: K.cell(rr, 4, f"=IF(ISNUMBER(C{rr}),IF(ABS(B{rr}-C{rr})<0.01,\"PASS\",\"FAIL\"),\"n/a\")")
    else: K.cell(rr, 4, f"=IF(ABS(B{rr}-C{rr})<0.01,\"PASS\",\"FAIL\")")
K.cell(len(checks) + 3, 1, "The 0%-AVA savings check assumes 0% AVA leaves agents unchanged, which holds because both sides then route the same calls to agents.").font = Font(italic=True, color="666666")
import os
wb.save(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "ava-break-even-check.xlsx"))
