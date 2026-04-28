import re
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd
base = Path("C:/Users/gupta/SumoCallibration")
xlsx = base / "Streamlit_Callibration" / "LocationID.xlsx"
xmlf = base / "Streamlit_Callibration" / "Network.net.xml"
print("=== Excel Summary ===")
xl = pd.ExcelFile(xlsx)
print("Sheets:", list(xl.sheet_names))
df = pd.read_excel(xlsx, sheet_name="Locations")
print("LocationsColumns:", list(df.columns))
print("LocationsFirst15Rows:")
print(df.head(15).fillna("").to_dict(orient="records"))
mask = df["LocationName"].astype(str).str.contains("Thapathali|Pulchowk South", case=False, na=False)
sel = df[mask].copy()
print("MatchedRowsCount:", len(sel))
print("MatchedRows:")
print(sel.fillna("").to_dict(orient="records"))
cand_cols = [c for c in df.columns if re.search(r"junction|node", str(c), re.I)]
if not cand_cols:
    cand_cols = [c for c in df.columns if re.search(r"id", str(c), re.I)]
print("CandidateIdColumns:", cand_cols)
row_junctions = []
for _, r in sel.iterrows():
    jid = ""
    for c in cand_cols:
        v = r.get(c, "")
        if pd.notna(v) and str(v).strip() != "":
            jid = str(v).strip()
            break
    row_junctions.append((r.to_dict(), jid))
print("=== Network Summary ===")
tree = ET.parse(xmlf)
root = tree.getroot()
edges = []
for e in root.findall("edge"):
    func = e.attrib.get("function", "normal")
    if func == "internal":
        continue
    edges.append(e.attrib)
conns = [c.attrib for c in root.findall("connection")]
juncs = {j.attrib.get("id"): j.attrib for j in root.findall("junction")}
for row, jid in row_junctions:
    print("Location:", row.get("LocationName", ""), "| JunctionID:", jid if jid else "(not found)")
    if not jid or jid not in juncs:
        print("  Junction: not found in Network.net.xml")
        continue
    j = juncs[jid]
    inc_lanes = [x for x in j.get("incLanes", "").split(" ") if x]
    incoming = [e.get("id") for e in edges if e.get("to") == jid]
    outgoing = [e.get("id") for e in edges if e.get("from") == jid]
    print("  JunctionXY:", {"x": j.get("x"), "y": j.get("y")})
    print("  incLanes:", inc_lanes[:12] + (["..."] if len(inc_lanes) > 12 else []))
    print("  IncomingNormalEdgesCount:", len(incoming), "OutgoingNormalEdgesCount:", len(outgoing))
    print("  IncomingNormalEdgesSample:", incoming[:8])
    print("  OutgoingNormalEdgesSample:", outgoing[:8])
    incoming_set = set(incoming)
    inc_lane_set = set(inc_lanes)
    touch = []
    for c in conns:
        fe = c.get("from")
        fl = c.get("fromLane")
        full_lane = f"{fe}_{fl}" if fe is not None and fl is not None else None
        if fe in incoming_set or full_lane in inc_lane_set:
            touch.append({k: c.get(k) for k in ["from", "to", "fromLane", "toLane", "via", "dir", "state"] if c.get(k) is not None})
    print("  ConnectionsTouchingIncoming (up to 8):", touch[:8])
