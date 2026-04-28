import pandas as pd
from pathlib import Path
from datetime import date

out_path = Path(r"ProcessedVideoOutput/EdgeWise_DirectionalCounts_5MinIntervalColumns_Datewise.xlsx")
sheet = "Motorcycle_5MinCols"
cols_needed = ["Date","Direction","09:00:00-09:05:00","09:05:00-09:10:00","08:00:00-08:05:00","SourceWorkbook"]

df = pd.read_excel(out_path, sheet_name=sheet)
df.columns = [str(c).strip() for c in df.columns]
if "Date" in df.columns:
    d = pd.to_datetime(df["Date"], errors="coerce").dt.date
else:
    d = pd.Series([None]*len(df))
loc_series = df["Location"] if "Location" in df.columns else pd.Series([None]*len(df))
dir_series = df["Direction"] if "Direction" in df.columns else pd.Series([None]*len(df))
mask = (
    (loc_series.astype(str).str.strip() == "Patan Dhoka Road")
    & (dir_series.astype(str).str.contains("South-to-North", case=False, na=False))
    & (d == date(2026,2,24))
)
res = df.loc[mask, [c for c in cols_needed if c in df.columns]].copy()
print("=== Output workbook rows (filtered) ===")
if res.empty:
    print("<no rows found>")
else:
    print(res.to_string(index=False))

src_files = [
    Path(r"ProcessedVideoOutput/Patan Dhoka Road/2026-02-24/9/approach_directional_09.00.00-09.05.00_R__0_0__0_20260416_001318.xlsx"),
    Path(r"ProcessedVideoOutput/Patan Dhoka Road/2026-02-24/9/approach_directional_09.05.00-09.10.00_R__0_0__0_20260416_002701.xlsx"),
]

def extract_motorcycle_s2n(path: Path):
    sdf = pd.read_excel(path, sheet_name="Direction Counts")
    sdf.columns = [str(c).strip() for c in sdf.columns]
    dir_col = next((c for c in sdf.columns if c.lower() == "direction"), None)
    moto_col = next((c for c in sdf.columns if c.lower() == "motorcycle"), None)
    if dir_col and moto_col:
        m = sdf[dir_col].astype(str).str.contains("South-to-North", case=False, na=False)
        vals = pd.to_numeric(sdf.loc[m, moto_col], errors="coerce").dropna()
        if len(vals):
            return int(vals.iloc[0])
    veh_col = next((c for c in sdf.columns if "vehicle" in c.lower()), None)
    cnt_col = next((c for c in sdf.columns if "count" in c.lower() or "volume" in c.lower()), None)
    if dir_col and veh_col and cnt_col:
        m = sdf[dir_col].astype(str).str.contains("South-to-North", case=False, na=False) & sdf[veh_col].astype(str).str.contains("Motorcycle", case=False, na=False)
        vals = pd.to_numeric(sdf.loc[m, cnt_col], errors="coerce").dropna()
        if len(vals):
            return int(vals.iloc[0])
    if dir_col:
        m = sdf[dir_col].astype(str).str.contains("South-to-North", case=False, na=False)
        row = sdf.loc[m]
        if not row.empty:
            for c in sdf.columns:
                if "motorcycle" in c.lower() or "bike" in c.lower():
                    vals = pd.to_numeric(row[c], errors="coerce").dropna()
                    if len(vals):
                        return int(vals.iloc[0])
            nums = pd.to_numeric(row.iloc[0], errors="coerce").dropna()
            if len(nums):
                return int(nums.iloc[0])
    return None

print("=== Source workbook South-to-North motorcycle counts (Direction Counts) ===")
source_counts = []
for p in src_files:
    val = extract_motorcycle_s2n(p)
    source_counts.append(val)
    print(f"{p}: {val}")

out_900 = None
out_905 = None
if not res.empty:
    out_900 = pd.to_numeric(res.iloc[0].get("09:00:00-09:05:00"), errors="coerce")
    out_905 = pd.to_numeric(res.iloc[0].get("09:05:00-09:10:00"), errors="coerce")
    out_900 = None if pd.isna(out_900) else int(out_900)
    out_905 = None if pd.isna(out_905) else int(out_905)

print("=== Match check ===")
print(f"output_09:00:00-09:05:00={out_900}, expected=247, exact_match={out_900 == 247}")
print(f"output_09:05:00-09:10:00={out_905}, expected=241, exact_match={out_905 == 241}")
print(f"source_file_1_exact_247={source_counts[0] == 247 if len(source_counts)>0 else False}")
print(f"source_file_2_exact_241={source_counts[1] == 241 if len(source_counts)>1 else False}")
print(f"output_counts_exactly_match_source_247_241={(out_900 == 247) and (out_905 == 241)}")
