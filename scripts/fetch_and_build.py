"""
Fetch Adidas macro report Excel from Google Drive, parse it, and generate docs/index.html.
Reads GOOGLE_SHEETS_CREDENTIALS (service account JSON) and SHEET_FILE_ID from env vars.
Skips rebuild if data hasn't changed (SHA-256 hash comparison).
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
from jinja2 import Environment, FileSystemLoader

SCRIPT_DIR = Path(__file__).parent
REPO_ROOT = SCRIPT_DIR.parent
CONFIG_FILE = REPO_ROOT / "config" / "sheet_config.json"
DOCS_DIR = REPO_ROOT / "docs"
HASH_FILE = DOCS_DIR / ".data_hash"
TEMPLATE_DIR = SCRIPT_DIR / "templates"

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


# ── Google Drive download ─────────────────────────────────────────────────────

def download_excel() -> io.BytesIO:
    creds_json = os.environ["GOOGLE_SHEETS_CREDENTIALS"]
    creds_dict = json.loads(creds_json)
    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    config = json.loads(CONFIG_FILE.read_text())
    file_id = os.environ["SHEET_FILE_ID"]  # required GitHub Secret

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
    """Parse IPC monthly data from Data sheet (columns Date, national %, clothing %)."""
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["data"], header=4)

    # Pandas renames duplicate columns: Date, Date.1, Date.2, Date.3
    date_col = "Date"
    nacional_col = next((c for c in df.columns if "nivel nacional(%)" in c), None)
    calzado_col = next((c for c in df.columns if "prendas y calzado" in c and "nacional" in c), None)

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
    """Parse monthly peso devaluation from Data sheet (4th Date block, Devaluation %)."""
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["data"], header=4)

    # 4th Date column becomes 'Date.3'
    date_col = "Date.3"
    dev_col = next((c for c in df.columns if "Devaluation" in c or "devaluation" in c), None)

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
    """Parse annual GDP growth from PBI pivot sheet."""
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["pbi"], header=0)

    results = []
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if not label or label in ("Row Labels", "Grand Total"):
            continue
        # Year rows are 4-digit numbers
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
    """Parse FX rate from hierarchical FX Rate pivot sheet."""
    config = json.loads(CONFIG_FILE.read_text())
    df = pd.read_excel(xl, sheet_name=config["sheet_names"]["fx_rate"], header=0)

    results = []
    current_year = None
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if not label or label in ("Row Labels", "Grand Total"):
            continue
        # Detect year rows
        if re.match(r"^\d{4}$", label):
            current_year = label[-2:]
            continue
        if current_year is None:
            continue
        month = label.lower()
        # Map Spanish month abbreviations
        month_num = MONTHS_ES_INV.get(month[:3])
        if not month_num:
            continue
        period = f"{month[:3]}-{current_year}"
        actual = _currency_to_float(row.iloc[1]) if len(row) > 1 else None
        proj = _currency_to_float(row.iloc[2]) if len(row) > 2 else None
        results.append({"period": period, "actual": actual, "proyeccion": proj})
    return results


def parse_inflacion_discriminada(xl: pd.ExcelFile) -> list:
    """Parse annual calzado/APP totals from Inflacion discriminada pivot sheet."""
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
        results.append({"year": year, "calzado": calzado, "app": app})
    return results


# ── Chart data builder ────────────────────────────────────────────────────────

def _latest_nonnull(data_list, key):
    for item in reversed(data_list):
        if item.get(key) is not None:
            return item
    return None


def _compute_insights(ipc, devaluation, pbi, fx, disc):
    insights = {}

    if ipc:
        latest = _latest_nonnull(ipc, "inflacion_nacional")
        peak = max(ipc, key=lambda r: r["inflacion_nacional"] or -999)
        trough = min(ipc, key=lambda r: r["inflacion_nacional"] if r["inflacion_nacional"] is not None else 999)
        insights["ipc_nacional"] = (
            f"Pico: <strong>{peak['inflacion_nacional']}%</strong> ({peak['date']}). "
            f"Mínimo: <strong>{trough['inflacion_nacional']}%</strong> ({trough['date']}). "
            f"Último dato: <strong>{latest['inflacion_nacional']}%</strong> ({latest['date']})."
        )
        insights["ipc_comparativo"] = (
            "La inflación en prendas y calzado es más volátil que el índice general. "
            "Los valores negativos indican deflación puntual en esa categoría."
        )

    if pbi:
        latest_actual = _latest_nonnull(pbi, "actual")
        latest_proj = _latest_nonnull(pbi, "proyeccion")
        insights["pbi"] = (
            f"Último dato real: <strong>{latest_actual['actual']}%</strong> ({latest_actual['year']}). "
            f"Proyección más reciente: <strong>{latest_proj['proyeccion']}%</strong> ({latest_proj['year']}). "
            "Fuente: Banco Mundial / REM-BCRA."
        )

    if fx:
        latest_actual = _latest_nonnull(fx, "actual")
        latest_proj = _latest_nonnull(fx, "proyeccion")
        insights["fx_rate"] = (
            f"Tipo de cambio actual: <strong>${latest_actual['actual']:,.2f}</strong> ({latest_actual['period']}). "
            f"Proyección: <strong>${latest_proj['proyeccion']:,.0f}</strong> ({latest_proj['period']}). "
            "Fuente: REM-BCRA."
        )

    if devaluation:
        latest = _latest_nonnull(devaluation, "devaluacion")
        peak = max(devaluation, key=lambda r: r["devaluacion"] or -999)
        insights["devaluacion"] = (
            f"Máximo mensual: <strong>{peak['devaluacion']}%</strong> ({peak['date']}). "
            f"Último dato: <strong>{latest['devaluacion']}%</strong> ({latest['date']}). "
            "Valores negativos indican apreciación del peso."
        )

    if disc:
        latest = disc[-1] if disc else None
        insights["discriminada"] = (
            f"Año más reciente ({latest['year']}): calzado <strong>{latest['calzado']}%</strong>, "
            f"indumentaria <strong>{latest['app']}%</strong>. "
            "Suma de inflaciones mensuales por categoría."
        )

    return insights


def build_chart_payload(ipc, devaluation, pbi, fx, disc):
    insights = _compute_insights(ipc, devaluation, pbi, fx, disc)

    # Overlap projection with last actual for smooth line connection
    if fx:
        last_actual_idx = max(
            (i for i, r in enumerate(fx) if r["actual"] is not None), default=None
        )
        if last_actual_idx is not None and last_actual_idx + 1 < len(fx):
            if fx[last_actual_idx + 1]["proyeccion"] is None:
                pass
            else:
                fx[last_actual_idx]["proyeccion"] = fx[last_actual_idx]["actual"]

    return [
        {
            "section": "Contexto Macroeconómico",
            "subtitle": "Argentina · Datos INDEC / Banco Mundial / REM-BCRA",
            "charts": [
                {
                    "id": "ipc_nacional",
                    "title": "IPC — Inflación Mensual Nacional",
                    "source": "INDEC",
                    "labels": [r["date"] for r in ipc],
                    "datasets": [
                        {
                            "label": "Inflación (%)",
                            "data": [r["inflacion_nacional"] for r in ipc],
                            "color": "#000000",
                            "dash": False,
                        }
                    ],
                    "type": "line",
                    "yLabel": "%",
                    "insight": insights.get("ipc_nacional", ""),
                },
                {
                    "id": "pbi",
                    "title": "Crecimiento del PBI Argentina",
                    "source": "Banco Mundial / REM-BCRA",
                    "labels": [r["year"] for r in pbi],
                    "datasets": [
                        {
                            "label": "Real (%)",
                            "data": [r["actual"] for r in pbi],
                            "color": "#000000",
                            "dash": False,
                        },
                        {
                            "label": "Proyección (%)",
                            "data": [r["proyeccion"] for r in pbi],
                            "color": "#aaaaaa",
                            "dash": True,
                        },
                    ],
                    "type": "bar",
                    "yLabel": "%",
                    "insight": insights.get("pbi", ""),
                },
                {
                    "id": "fx_rate",
                    "title": "Tipo de Cambio ARS / USD",
                    "source": "BCRA / REM-BCRA",
                    "labels": [r["period"] for r in fx],
                    "datasets": [
                        {
                            "label": "ARS/USD (real)",
                            "data": [r["actual"] for r in fx],
                            "color": "#000000",
                            "dash": False,
                        },
                        {
                            "label": "ARS/USD (proyección)",
                            "data": [r["proyeccion"] for r in fx],
                            "color": "#aaaaaa",
                            "dash": True,
                        },
                    ],
                    "type": "line",
                    "yLabel": "ARS",
                    "insight": insights.get("fx_rate", ""),
                },
                {
                    "id": "devaluacion",
                    "title": "Devaluación Mensual del Peso",
                    "source": "BCRA / REM-BCRA",
                    "labels": [r["date"] for r in devaluation],
                    "datasets": [
                        {
                            "label": "Devaluación (%)",
                            "data": [r["devaluacion"] for r in devaluation],
                            "color": "CONDITIONAL",
                            "dash": False,
                        }
                    ],
                    "type": "bar_conditional",
                    "yLabel": "%",
                    "insight": insights.get("devaluacion", ""),
                },
            ],
        },
        {
            "section": "Sector Calzado e Indumentaria",
            "subtitle": "Impacto en categorías clave de adidas",
            "charts": [
                {
                    "id": "ipc_comparativo",
                    "title": "Prendas y Calzado vs. Inflación General",
                    "source": "INDEC",
                    "labels": [r["date"] for r in ipc],
                    "datasets": [
                        {
                            "label": "Nacional (%)",
                            "data": [r["inflacion_nacional"] for r in ipc],
                            "color": "#000000",
                            "dash": False,
                        },
                        {
                            "label": "Prendas y Calzado (%)",
                            "data": [r["inflacion_calzado"] for r in ipc],
                            "color": "#767677",
                            "dash": False,
                        },
                    ],
                    "type": "line",
                    "yLabel": "%",
                    "insight": insights.get("ipc_comparativo", ""),
                },
                {
                    "id": "discriminada",
                    "title": "Inflación Acumulada Anual — Calzado y APP",
                    "source": "INDEC (discriminada por región)",
                    "labels": [r["year"] for r in disc],
                    "datasets": [
                        {
                            "label": "Calzado (%)",
                            "data": [r["calzado"] for r in disc],
                            "color": "#000000",
                            "dash": False,
                        },
                        {
                            "label": "Indumentaria / APP (%)",
                            "data": [r["app"] for r in disc],
                            "color": "#767677",
                            "dash": False,
                        },
                    ],
                    "type": "bar",
                    "yLabel": "%",
                    "insight": insights.get("discriminada", ""),
                },
            ],
        },
    ]


# ── Hash + change detection ───────────────────────────────────────────────────

def compute_hash(sections: list) -> str:
    blob = json.dumps(sections, sort_keys=True, ensure_ascii=False, default=str)
    return hashlib.sha256(blob.encode()).hexdigest()


def has_changed(new_hash: str) -> bool:
    if not HASH_FILE.exists():
        return True
    return HASH_FILE.read_text().strip() != new_hash


def save_hash(new_hash: str):
    HASH_FILE.write_text(new_hash)


# ── Render HTML ───────────────────────────────────────────────────────────────

def render_html(sections: list) -> str:
    env = Environment(loader=FileSystemLoader(str(TEMPLATE_DIR)), autoescape=False)
    tmpl = env.get_template("report.html.j2")
    return tmpl.render(
        sections=sections,
        chart_data_json=json.dumps(sections, ensure_ascii=False),
        generated_at=datetime.utcnow().strftime("%-d de %B, %Y — %H:%M UTC"),
    )


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

    sections = build_chart_payload(ipc, devaluation, pbi, fx, disc)
    new_hash = compute_hash(sections)

    if not has_changed(new_hash):
        print("No data changes detected. Skipping rebuild.")
        return

    print("Data changed. Rebuilding HTML...")
    html = render_html(sections)

    DOCS_DIR.mkdir(exist_ok=True)
    (DOCS_DIR / "index.html").write_text(html, encoding="utf-8")
    save_hash(new_hash)
    print("Report rebuilt successfully.")


if __name__ == "__main__":
    main()
