from pathlib import Path
import subprocess
import xml.etree.ElementTree as ET
import os

ROOT = Path(r"C:\Users\gupta\SumoCallibration")
BASE = ROOT / "project" / "results" / "cadyts_corridor_20260222_20260225"
CAL = BASE / "calibration"
VAL = BASE / "validation"
SUMO_HOME = Path(os.environ.get("SUMO_HOME", r"C:\Program Files (x86)\Eclipse\Sumo"))
CADYTS = SUMO_HOME / "tools" / "assign" / "cadytsIterate.py"
JAVA_HOME = BASE / "tools" / "jdk-21.0.10+7-jre"
JAR = BASE / "cadyts.jar"
PY = Path(r"c:/Users/gupta/SumoCallibration/.venv_separate/Scripts/python.exe")
SCRIPT_ANALYZE = ROOT / "project" / "scripts" / "analyze_calibration.py"
JUNCTION_MAP = ROOT / "project" / "scripts" / "data" / "junction_id.xlsx"
SIG_CAL = ROOT / "project" / "results" / "corridor_calib_20260222_validate_20260225" / "inputs" / "signal_timings_20260222_0900_1000.xlsx"
TRAFFIC_CAL = ROOT / "project" / "results" / "corridor_calib_20260222_validate_20260225" / "inputs" / "directional_traffic_20260222_0900_1000.xlsx"
TRAFFIC_VAL = ROOT / "project" / "results" / "corridor_calib_20260222_validate_20260225" / "inputs" / "directional_traffic_20260225_0900_0920.xlsx"


def run(cmd, cwd=None, env=None):
    print("[run]", " ".join(str(c) for c in cmd))
    subprocess.run([str(c) for c in cmd], cwd=str(cwd) if cwd else None, check=True, env=env)


def build_sampled_alts(folder: Path, sample=0.15):
    flows_file = folder / "routes.rou.xml"
    trips_file = folder / "trips_cadyts_sampled.xml"
    routes_file = folder / "routes_cadyts_sampled.rou.xml"
    alt_file = folder / "routes_cadyts_sampled.rou.alt.xml"
    alt_linefix = folder / "routes_cadyts_sampled_linefix.rou.alt.xml"

    root = ET.parse(flows_file).getroot()
    r = ET.Element("routes")
    ET.SubElement(r, "vType", {
        "id": "passenger", "accel": "2.6", "decel": "4.5", "sigma": "0.5",
        "length": "5.0", "minGap": "2.5", "maxSpeed": "16.67"
    })

    for f in root.findall("flow"):
        fid = f.get("id", "flow")
        frm = f.get("from")
        to = f.get("to")
        via = f.get("via")
        n = int(float(f.get("number", "0")))
        n = max(1, int(round(n * sample))) if n > 0 else 0
        if not frm or not to or n <= 0:
            continue
        for i in range(n):
            dep = str(round((i + 0.5) * 3600.0 / max(1, n), 3))
            attrs = {"id": f"{fid}.{i}", "type": "passenger", "depart": dep, "from": frm, "to": to}
            if via:
                attrs["via"] = via
            ET.SubElement(r, "trip", attrs)

    ET.indent(r)
    trips_file.write_text(ET.tostring(r, encoding="unicode"), encoding="utf-8")

    run(["duarouter", "-n", "Network.net.xml", "-r", trips_file.name, "-o", routes_file.name, "--ignore-errors", "true"], cwd=folder)

    alt_root = ET.parse(alt_file).getroot()
    for v in alt_root.findall("vehicle"):
        if v.get("line") is None:
            v.set("line", "")
    ET.indent(alt_root)
    alt_linefix.write_text(ET.tostring(alt_root, encoding="unicode"), encoding="utf-8")
    return alt_linefix


def run_cadyts(folder: Path, choiceset: Path):
    env = os.environ.copy()
    env["JAVA_HOME"] = str(JAVA_HOME)
    env["PATH"] = str(JAVA_HOME / "bin") + os.pathsep + env.get("PATH", "")
    lck = folder / "calibration-log.txt.lck"
    if lck.exists():
        lck.unlink()
    run([
        PY, CADYTS,
        "-n", "Network.net.xml",
        "-r", choiceset.name,
        "-d", "detector_output.xml",
        "--classpath", str(JAR),
        "-l", "4",
        "-S", "1.0",
        "-F", "2",
        "-P", "1",
        "-W", "cadyts_eval_sampled"
    ], cwd=folder, env=env)


def latest_cadyts_route(folder: Path):
    it_dirs = sorted([p for p in folder.iterdir() if p.is_dir() and p.name.isdigit()], key=lambda p: int(p.name))
    if not it_dirs:
        raise RuntimeError(f"No CADYTS iteration folders in {folder}")
    last = it_dirs[-1]
    cal_files = sorted(last.glob("*.cal.xml"))
    if not cal_files:
        raise RuntimeError(f"No calibrated route file in {last}")
    return cal_files[0]


def eval_with_sumo_and_analyze(folder: Path, traffic_xlsx: Path, out_prefix: str):
    route_file = latest_cadyts_route(folder)
    run([
        "sumo",
        "-n", "Network.net.xml",
        "-r", str(route_file),
        "-a", "traffic_lights.add.xml,detectors.add.xml",
        "--tripinfo-output", f"tripinfo_{out_prefix}.xml",
        "--queue-output", f"queue_{out_prefix}.xml",
        "--edgedata-output", f"edge_data_{out_prefix}.xml",
        "--begin", "0",
        "--end", "3600",
    ], cwd=folder)

    report = folder / f"{out_prefix}_report.txt"
    demand = folder / f"demand_comparison_{out_prefix}.csv"
    run([
        PY, SCRIPT_ANALYZE,
        "--traffic", str(traffic_xlsx),
        "--junction-map", str(JUNCTION_MAP),
        "--signal-timings", str(SIG_CAL),
        "--routes", str(route_file),
        "--traffic-lights", str(folder / "traffic_lights.add.xml"),
        "--tripinfo", str(folder / f"tripinfo_{out_prefix}.xml"),
        "--edge-data", str(folder / f"edge_data_{out_prefix}.xml"),
        "--queue", str(folder / f"queue_{out_prefix}.xml"),
        "--detector", str(folder / "detector_output.xml"),
        "--demand-report-csv", str(demand),
        "--report", str(report),
    ], cwd=folder)
    return report


def pick_metrics(report_file: Path):
    lines = report_file.read_text(encoding="utf-8", errors="ignore").splitlines()
    keys = [
        "- Total observed demand:",
        "- Total generated demand:",
        "- Relative difference:",
        "- GEH < 5 share:",
        "- Completed trips:",
        "- Mean trip duration:",
        "- Mean waiting time:",
        "- Mean speed relative (edge weighted):",
        "- Edges with teleports:",
        "- Maximum queue length:",
        "- Realism classification:",
    ]
    out = []
    for k in keys:
        for ln in lines:
            if ln.strip().startswith(k):
                out.append(ln.strip())
                break
    return out


cal_choice = build_sampled_alts(CAL)
val_choice = build_sampled_alts(VAL)
run_cadyts(CAL, cal_choice)
run_cadyts(VAL, val_choice)

cal_report = eval_with_sumo_and_analyze(CAL, TRAFFIC_CAL, "cadyts_final_calibration")
val_report = eval_with_sumo_and_analyze(VAL, TRAFFIC_VAL, "cadyts_final_validation")

final = BASE / "cadyts_final_report.txt"
text = []
text.append("CADYTS Corridor Calibration/Validation Report (Separate Folder)")
text.append("============================================================")
text.append("")
text.append("Method")
text.append("- CADYTS using sampled trip-based route alternatives derived from corridor flows")
text.append("- 4 CADYTS iterations, prep=1, freeze=2, demandscale=1.0")
text.append("")
text.append("Calibration (2026-02-22 09:00-10:00)")
text.extend(pick_metrics(cal_report))
text.append("")
text.append("Validation (2026-02-25 09:00-09:20)")
text.extend(pick_metrics(val_report))
text.append("")
text.append(f"Calibration folder: {CAL}")
text.append(f"Validation folder: {VAL}")
text.append(f"Calibration report: {cal_report}")
text.append(f"Validation report: {val_report}")
final.write_text("\n".join(text), encoding="utf-8")
print("Wrote", final)
