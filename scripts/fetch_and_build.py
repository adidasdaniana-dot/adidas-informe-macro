"""
Fetch Adidas macro report Excel from Google Drive, parse it, and surgically
update var CHARTS in docs/index.html. Never regenerates the full HTML.
Reads GOOGLE_SHEETS_CREDENTIALS (service account JSON) and SHEET_FILE_ID from env vars.
Skips update if data hasn't changed (SHA-256 hash comparison).
"""

import hashlib
import io
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_DIR = Path(__file__).parent
REPO_ROOT = SCRIPT_DIR.parent
CONFIG_FILE = REPO_ROOT / "config" / "sheet_config.json"
DOCS_DIR = REPO_ROOT / "docs"
HASH_FILE = DOCS_DIR / ".data_hash"
INDEX_HTML = DOCS_DIR / "index.html"

MONTHS_ES = {
    "01": "ene", "02": "feb", "03": "mar", "04": "abr",
    "05": "may", "06": "jun", "07": "jul", "08": "ago",
    "09": "sep", "10": "oct", "11": "nov", "12": "dic",
}
MONTHS_ES_INV = {v: k for k, v in MONTHS_ES.items()}


# ── helpers ──────────────────────────────────────────────────────────────────

def _pct_to_float(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().replace("%", "").replace(",", ".").replace(" ", "")
    if s in ("", "#DIV/0!", "#REF!", "nan"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _currency_to_float(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().replace("$", "").replace(",", "").replace(" ", "")
    if s in ("", "nan"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _format_date(val):
    if isinstance(val, datetime):
        m = val.strftime("%m")
        y = val.strftime("%y")
        return f"{MONTHS_ES.get(m, m)}-{y}"
    s = str(val).strip()
    parts = re.split(r"[/\-]", s)
    if len(parts) == 3:
        m = parts[0].zfill(2)
        y = parts[2][-2:]
        return f"{MONTHS_ES.get(m, m)}-{y}"
    return s


def _pct_scale(data):
    """If all non-null values look like decimals (|val| < 0.5), multiply by 100."""
    non_null = [v for v in data if v is not None]
    if not non_null:
        return data
    if max(abs(v) for v in non_null) < 0.5:
        return [round(v * 100, 2) if v is not None else None for v in data]
    return [round(v, 4) if v is not None else None for v in data]


def _js_val(v):
    """Format a Python value as JavaScript literal."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, float):
        # Avoid Python's -0.0
        return str(v) if v != 0 else "0"
    return str(v)


def _js_list(lst):
    return "[" + ",".join(_js_val(v) for v in lst) + "]"


def _js_str_list(lst):
    return "[" + ",".join(f'"{v}"' for v in lst) + "]"


# ── Google Drive download ─────────────────────────────────────────────────────

def download_excel() -> io.BytesIO:
    creds_json = os.environ["GOOGLE_SHEETS_CREDENTIALS"]
    creds_dict = json.loads(creds_json)
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    file_id = os.environ["SHEET_FILE_ID"]

    service = build("drive", "v3", credentials=credentials)
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf


# ── Excel parsers ─────────────────────────────────────────────────────────────

def parse_ipc(xl: pd.ExcelFile) -> list:
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["data"], header=4)

    date_col = "Date"
    nacional_col = next((c for c in df.columns if "nivel nacional(%)" in str(c).lower()), None)
    calzado_col = next((c for c in df.columns if "prendas y calzado" in str(c).lower()), None)

    if not nacional_col:
        print("WARNING: Could not find national inflation column", file=sys.stderr)
        return []

    results = []
    for _, row in df.iterrows():
        date_val = row.get(date_col)
        if pd.isna(date_val) or str(date_val).strip() == "":
            continue
        infl = _pct_to_float(row.get(nacional_col))
        if infl is None:
            continue
        calzado = _pct_to_float(row.get(calzado_col)) if calzado_col else None
        results.append({
            "date": _format_date(date_val),
            "inflacion_nacional": infl,
            "inflacion_calzado": calzado,
        })
    return results


def parse_devaluation(xl: pd.ExcelFile) -> list:
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["data"], header=4)

    date_col = "Date.3"
    dev_col = next((c for c in df.columns if "devaluation" in str(c).lower()), None)

    if date_col not in df.columns or not dev_col:
        print("WARNING: Devaluation columns not found", file=sys.stderr)
        return []

    results = []
    for _, row in df.iterrows():
        date_val = row.get(date_col)
        dev = _pct_to_float(row.get(dev_col))
        if pd.isna(date_val) or str(date_val).strip() == "" or dev is None:
            continue
        results.append({
            "date": _format_date(date_val),
            "devaluacion": dev,
        })
    return results


def parse_pbi(xl: pd.ExcelFile) -> list:
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["pbi"], header=0)

    results = []
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if not label or label in ("Row Labels", "Grand Total"):
            continue
        try:
            year = int(float(label))
            if not (2000 <= year <= 2035):
                continue
        except ValueError:
            continue
        actual = _pct_to_float(row.iloc[1])
        proj = _pct_to_float(row.iloc[2]) if len(row) > 2 else None
        results.append({"year": str(year), "actual": actual, "proyeccion": proj})
    return results


def parse_fx_rate(xl: pd.ExcelFile) -> list:
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["fx_rate"], header=0)

    results = []
    current_year = None
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if not label or label in ("Row Labels", "Grand Total"):
            continue
        if re.match(r"^\d{4}$", label):
            current_year = label[-2:]
            continue
        if current_year is None:
            continue
        month = label.lower()
        month_num = MONTHS_ES_INV.get(month[:3])
        if not month_num:
            continue
        period = f"{month[:3]}-{current_year}"
        actual = _currency_to_float(row.iloc[1]) if len(row) > 1 else None
        proj = _currency_to_float(row.iloc[2]) if len(row) > 2 else None
        results.append({"period": period, "actual": actual, "proyeccion": proj})
    return results


def parse_inflacion_discriminada(xl: pd.ExcelFile) -> list:
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["inflacion_discriminada"], header=0)

    results = []
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if "Total" not in label:
            continue
        year = label.replace(" Total", "").strip()
        calzado = _pct_to_float(row.iloc[1]) if len(row) > 1 else None
        app = _pct_to_float(row.iloc[2]) if len(row) > 2 else None
        infl_general = _pct_to_float(row.iloc[3]) if len(row) > 3 else None
        results.append({"year": year, "calzado": calzado, "app": app, "infl_general": infl_general})
    return results


# ── Build CHARTS JS block ─────────────────────────────────────────────────────

def build_charts_js(ipc, devaluation, pbi, fx, disc) -> str:
    """Build the var CHARTS = {...}; JS block to inject into index.html."""

    ipc_dates = [r["date"] for r in ipc]
    ipc_nacional = _pct_scale([r["inflacion_nacional"] for r in ipc])
    ipc_calzado = _pct_scale([r["inflacion_calzado"] for r in ipc])

    dev_dates = [r["date"] for r in devaluation]
    dev_vals = _pct_scale([r["devaluacion"] for r in devaluation])

    pbi_years = [r["year"] for r in pbi]
    pbi_actual = _pct_scale([r["actual"] for r in pbi])
    pbi_proj = _pct_scale([r["proyeccion"] for r in pbi])

    # Overlap projection with last actual for smooth line connection
    if fx:
        last_actual_idx = max(
            (i for i, r in enumerate(fx) if r["actual"] is not None), default=None
        )
        if last_actual_idx is not None and last_actual_idx + 1 < len(fx):
            if fx[last_actual_idx + 1]["proyeccion"] is not None:
                fx[last_actual_idx]["proyeccion"] = fx[last_actual_idx]["actual"]

    fx_periods = [r["period"] for r in fx]
    fx_actual = [round(r["actual"], 2) if r["actual"] is not None else None for r in fx]
    fx_proj = [round(r["proyeccion"], 2) if r["proyeccion"] is not None else None for r in fx]

    disc_years = [r["year"] for r in disc]
    disc_calzado = _pct_scale([r["calzado"] for r in disc])
    disc_app = _pct_scale([r["app"] for r in disc])
    disc_general = _pct_scale([r.get("infl_general") for r in disc])

    lines = ["var CHARTS = {"]

    lines.append(f"  ipc_nacional: {{")
    lines.append(f"    title: 'IPC — Inflación Mensual Nacional', source: 'INDEC',")
    lines.append(f"    type: 'line', yLabel: '%', granularity: 'month',")
    lines.append(f"    labels: {_js_str_list(ipc_dates)},")
    lines.append(f"    datasets: [")
    lines.append(f'      {{label:"Inflación (%)", data:{_js_list(ipc_nacional)}, color:"#000000", dash:false}}')
    lines.append(f"    ]")
    lines.append(f"  }},")

    lines.append(f"  ipc_comparativo: {{")
    lines.append(f"    title: 'Prendas y Calzado vs. Inflación General', source: 'INDEC',")
    lines.append(f"    type: 'line', yLabel: '%', granularity: 'month',")
    lines.append(f"    labels: {_js_str_list(ipc_dates)},")
    lines.append(f"    datasets: [")
    lines.append(f'      {{label:"Nacional (%)", data:{_js_list(ipc_nacional)}, color:"#000000", dash:false}},')
    lines.append(f'      {{label:"Prendas y Calzado (%)", data:{_js_list(ipc_calzado)}, color:"#767677", dash:false}}')
    lines.append(f"    ]")
    lines.append(f"  }},")

    lines.append(f"  pbi: {{")
    lines.append(f"    title: 'Crecimiento del PBI Argentina', source: 'Banco Mundial / REM-BCRA',")
    lines.append(f"    type: 'bar', yLabel: '%', granularity: 'year',")
    lines.append(f"    labels: {_js_str_list(pbi_years)},")
    lines.append(f"    datasets: [")
    lines.append(f'      {{label:"Real (%)", data:{_js_list(pbi_actual)}, color:"#000000", dash:false}},')
    lines.append(f'      {{label:"Proyección (%)", data:{_js_list(pbi_proj)}, color:"#aaaaaa", dash:true}}')
    lines.append(f"    ]")
    lines.append(f"  }},")

    lines.append(f"  fx_rate: {{")
    lines.append(f"    title: 'Tipo de Cambio ARS / USD', source: 'BCRA / REM-BCRA',")
    lines.append(f"    type: 'line', yLabel: 'ARS', granularity: 'month',")
    lines.append(f"    labels: {_js_str_list(fx_periods)},")
    lines.append(f"    datasets: [")
    lines.append(f'      {{label:"ARS/USD (real)", data:{_js_list(fx_actual)}, color:"#000000", dash:false}},')
    lines.append(f'      {{label:"ARS/USD (proyección)", data:{_js_list(fx_proj)}, color:"#aaaaaa", dash:true}}')
    lines.append(f"    ]")
    lines.append(f"  }},")

    lines.append(f"  devaluacion: {{")
    lines.append(f"    title: 'Devaluación Mensual del Peso', source: 'BCRA / REM-BCRA',")
    lines.append(f"    type: 'bar_conditional', yLabel: '%', granularity: 'month',")
    lines.append(f"    labels: {_js_str_list(dev_dates)},")
    lines.append(f"    datasets: [")
    lines.append(f'      {{label:"Devaluación (%)", data:{_js_list(dev_vals)}, color:"CONDITIONAL", dash:false}}')
    lines.append(f"    ]")
    lines.append(f"  }},")

    lines.append(f"  discriminada: {{")
    lines.append(f"    title: 'Inflación Acumulada Anual — Calzado y APP', source: 'INDEC',")
    lines.append(f"    type: 'bar', yLabel: '%', granularity: 'year',")
    lines.append(f"    labels: {_js_str_list(disc_years)},")
    lines.append(f"    datasets: [")
    lines.append(f'      {{label:"Calzado (%)", data:{_js_list(disc_calzado)}, color:"#000000", dash:false}},')
    lines.append(f'      {{label:"Indumentaria / APP (%)", data:{_js_list(disc_app)}, color:"#767677", dash:false}},')
    lines.append(f'      {{label:"Inflación General (%)", data:{_js_list(disc_general)}, color:"#c0c0c0", dash:false}}')
    lines.append(f"    ]")
    lines.append(f"  }},")

    # pbi_sectorial is static (CEPAL data, not in Excel)
    lines.append(f'  pbi_sectorial: {{')
    lines.append(f'    title: \'PBI por Sector de Actividad — Q4 2025 vs. LY\', source: \'CEPAL\',')
    lines.append(f"    type: 'bar_h_conditional', yLabel: '%', granularity: 'sector',")
    lines.append(f'    labels: ["Intermediación financiera", "Agricultura, ganadería y caza", "Pesca", "Explotación de minas", "Impuestos netos de subsidios", "Electricidad, gas y agua", "Transporte y comunicaciones", "Act. inmobiliarias y empresarial", "Otras actividades servicios", "Construcción", "Enseñanza", "Servicios sociales y de salud", "Hogares privados", "Hoteles y restaurantes", "Adm. pública y defensa", "Comercio mayorista y minorista", "Industria manufacturera"],')
    lines.append(f'    datasets: [')
    lines.append(f'      {{label:"Var. % interanual", data:[17.2, 16.1, 10.6, 8.1, 5.5, 4.7, 2.2, 2.0, 1.1, 0.9, 0.6, 0.3, 0.0, -0.7, -1.1, -2.2, -5.0], color:"CONDITIONAL_H", dash:false}}')
    lines.append(f'    ]')
    lines.append(f'  }}')

    lines.append("};")
    return "\n".join(lines)


# ── Hash + change detection ───────────────────────────────────────────────────

def compute_hash(ipc, devaluation, pbi, fx, disc) -> str:
    blob = json.dumps(
        {"ipc": ipc, "dev": devaluation, "pbi": pbi, "fx": fx, "disc": disc},
        sort_keys=True, ensure_ascii=False, default=str
    )
    return hashlib.sha256(blob.encode()).hexdigest()


def has_changed(new_hash: str) -> bool:
    if not HASH_FILE.exists():
        return True
    return HASH_FILE.read_text().strip() != new_hash


def save_hash(new_hash: str):
    HASH_FILE.write_text(new_hash)


# ── Surgical HTML update ───────────────────────────────────────────────────────

def update_html(charts_js: str):
    html = INDEX_HTML.read_text(encoding="utf-8")

    # Replace var CHARTS = {...}; block
    new_html = re.sub(
        r"var CHARTS = \{.*?\};",
        charts_js,
        html,
        count=1,
        flags=re.DOTALL,
    )
    if new_html == html:
        print("WARNING: Could not find var CHARTS block to replace", file=sys.stderr)

    # Update header date
    today = datetime.utcnow()
    MONTHS_ES_LONG = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
    }
    date_str = f"{today.day} de {MONTHS_ES_LONG[today.month]}, {today.year}"
    new_html = re.sub(
        r'(<span class="header-updated-date">)[^<]*(</span>)',
        rf'\g<1>{date_str}\g<2>',
        new_html,
    )

    INDEX_HTML.write_text(new_html, encoding="utf-8")
    print(f"Report updated successfully. Date set to: {date_str}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS", "").strip()
    file_id = os.environ.get("SHEET_FILE_ID", "").strip()
    if not creds_json or not file_id:
        print("Secrets GOOGLE_SHEETS_CREDENTIALS / SHEET_FILE_ID not configured. Skipping update.")
        sys.exit(0)

    print("Downloading Excel from Google Drive...")
    file_bytes = download_excel()

    print("Parsing sheets...")
    xl = pd.ExcelFile(file_bytes)
    ipc = parse_ipc(xl)
    devaluation = parse_devaluation(xl)
    pbi = parse_pbi(xl)
    fx = parse_fx_rate(xl)
    disc = parse_inflacion_discriminada(xl)

    print(f"  IPC rows: {len(ipc)}")
    print(f"  Devaluation rows: {len(devaluation)}")
    print(f"  PBI rows: {len(pbi)}")
    print(f"  FX Rate rows: {len(fx)}")
    print(f"  Discriminada rows: {len(disc)}")

    new_hash = compute_hash(ipc, devaluation, pbi, fx, disc)

    if not has_changed(new_hash):
        print("No data changes detected. Skipping rebuild.")
        return

    print("Data changed. Updating HTML surgically...")
    charts_js = build_charts_js(ipc, devaluation, pbi, fx, disc)
    update_html(charts_js)
    save_hash(new_hash)


if __name__ == "__main__":
    main()
