import json
import xml.etree.ElementTree as ET
from pathlib import Path
import pandas as pd

base = Path(r"C:/Users/gupta/SumoCallibration")
xlsx = base / "Streamlit_Callibration" / "LocationID.xlsx"
xmlf = base / "Streamlit_Callibration" / "Network.net.xml"

cols_to_check = ["LocationID", "RenameLocationID", "Network ID", "LocationName"]

def norm(v):
    if pd.isna(v):
        return None
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    s = str(v).strip()
    if s == "" or s.lower() == "nan":
        return None
    return s

df = pd.read_excel(xlsx, sheet_name="Locations")
rows = []
for rec in df.to_dict(orient="records"):
    out = {}
    for k, v in rec.items():
        nv = None if pd.isna(v) else v
        if isinstance(nv, float) and nv.is_integer():
            nv = int(nv)
        out[k] = nv
    rows.append(out)

tree = ET.parse(xmlf)
root = tree.getroot()
junctions = {}
for j in root.findall("junction"):
    jid = j.attrib.get("id")
    if jid is not None:
        y = j.attrib.get("y")
        try:
            yv = float(y) if y is not None else None
        except Exception:
            yv = None
        junctions[jid] = {"y": yv}

junction_ids = set(junctions.keys())
match_rows = []
matched_points = []

for idx, rec in enumerate(rows):
    matched = []
    for c in cols_to_check:
        val = norm(rec.get(c))
        if val is not None and val in junction_ids:
            matched.append({"column": c, "junctionId": val})
            yv = junctions[val]["y"]
            if yv is not None:
                matched_points.append({
                    "rowIndex": idx,
                    "locationName": norm(rec.get("LocationName")),
                    "junctionId": val,
                    "y": yv
                })
    match_rows.append({
        "rowIndex": idx,
        "locationName": norm(rec.get("LocationName")),
        "matches": matched
    })

min_info = None
max_info = None
if matched_points:
    min_y = min(p["y"] for p in matched_points)
    max_y = max(p["y"] for p in matched_points)
    min_pts = [p for p in matched_points if p["y"] == min_y]
    max_pts = [p for p in matched_points if p["y"] == max_y]

    def pack(pts, yval):
        names = sorted({p["locationName"] for p in pts if p.get("locationName") is not None})
        jids = sorted({p["junctionId"] for p in pts})
        return {"y": yval, "locationNames": names, "junctionIds": jids}

    min_info = pack(min_pts, min_y)
    max_info = pack(max_pts, max_y)

result = {
    "locationsRows": rows,
    "rowMatches": match_rows,
    "matchedYExtremes": {"min": min_info, "max": max_info}
}

print(json.dumps(result, ensure_ascii=False, separators=(",", ":"), default=str))
