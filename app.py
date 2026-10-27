import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import sqlite3
import json
import hashlib
import re
from pathlib import Path
from datetime import date, datetime
from io import BytesIO

# ============================================================
# APPLUS+ | CONTROL CORPORATIVO DE VENCIMIENTOS
# Lee Excel/CSV con estructuras diferentes, conserva los archivos
# originales y consolida los vencimientos reconocidos.
# ============================================================

st.set_page_config(
    page_title="Control corporativo de vencimientos | Applus+",
    page_icon="📅",
    layout="wide"
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
ORIGINALS_DIR = DATA_DIR / "originales"
DATA_DIR.mkdir(exist_ok=True)
ORIGINALS_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "vencimientos.db"

CATEGORIES = [
    "EMO", "Licencia de conducir", "SOAT", "RTV",
    "SCTR", "Seguro vehicular", "Contrato", "Mantenimiento",
    "Calibración", "Certificado / permiso", "Otro"
]

# Palabras clave usadas para reconocer columnas; se comparan sin tildes.
ALIASES = {
    "due": [
        "fecha vencimiento", "fecha de vencimiento", "vencimiento",
        "vence", "vigente hasta", "vigencia hasta", "fecha caducidad",
        "fecha de caducidad", "caducidad", "fecha fin", "fin de vigencia",
        "expiry date", "expiration date", "valid until", "fecha termino",
        "fecha término", "proxima revision", "próxima revisión",
        "proximo mantenimiento", "próximo mantenimiento"
    ],
    "issue": [
        "fecha emision", "fecha de emision", "fecha emisión",
        "fecha de emisión", "fecha inicio", "fecha de inicio",
        "fecha expedicion", "fecha de expedicion", "fecha expedición",
        "fecha de expedición", "fecha examen", "fecha emo"
    ],
    "project": ["proyecto", "project", "obra", "servicio", "contrato/proyecto"],
    "person": [
        "nombres", "nombre completo", "trabajador", "colaborador",
        "empleado", "conductor", "responsable del equipo", "titular",
        "persona", "razon social", "razón social"
    ],
    "identifier": [
        "codigo", "código", "dni", "documento", "placa", "matricula",
        "matrícula", "id", "nro", "numero", "número", "serie",
        "codigo equipo", "código equipo", "unidad", "activo"
    ],
    "responsible": [
        "responsable", "encargado", "administrador", "asistente",
        "responsable seguimiento", "custodio"
    ],
    "email": ["correo", "email", "e-mail", "correo responsable", "mail"],
    "type": [
        "tipo documento", "documento", "tipo de control", "control",
        "certificado", "tipo"
    ],
    "status": ["estado", "estado seguimiento", "situacion", "situación"]
}

def norm(value):
    value = str(value).strip().lower()
    replacements = {
        "á":"a", "é":"e", "í":"i", "ó":"o", "ú":"u",
        "ü":"u", "ñ":"n"
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    value = re.sub(r"[_\-./]+", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value

def find_column(columns, key):
    aliases = [norm(x) for x in ALIASES[key]]
    normalized = {col: norm(col) for col in columns}
    # Prefer exact header matches.
    for col, ncol in normalized.items():
        if ncol in aliases:
            return col
    # Then match headers that contain a known phrase.
    for col, ncol in normalized.items():
        if any(alias in ncol for alias in aliases if len(alias) >= 4):
            return col
    return None

def infer_category(filename, columns):
    text = norm(filename + " " + " ".join(map(str, columns)))
    rules = [
        (["emo", "examen medico", "examen ocupacional"], "EMO"),
        (["licencia de conducir", "brevete", "licencia conducir"], "Licencia de conducir"),
        (["soat"], "SOAT"),
        (["rtv", "revision tecnica", "revisión técnica"], "RTV"),
        (["sctr"], "SCTR"),
        (["seguro vehicular", "poliza vehicular", "póliza vehicular"], "Seguro vehicular"),
        (["contrato", "fecha fin contrato"], "Contrato"),
        (["calibracion", "calibración"], "Calibración"),
        (["mantenimiento", "proximo servicio", "próximo servicio"], "Mantenimiento"),
        (["certificado", "permiso", "autorizacion", "autorización"], "Certificado / permiso"),
    ]
    for needles, category in rules:
        if any(norm(n) in text for n in needles):
            return category
    return "Otro"

def read_workbook(uploaded_file):
    name = uploaded_file.name
    content = uploaded_file.getvalue()
    if name.lower().endswith(".csv"):
        try:
            return {"CSV": pd.read_csv(BytesIO(content), dtype=str, keep_default_na=False)}
        except UnicodeDecodeError:
            return {"CSV": pd.read_csv(BytesIO(content), dtype=str, keep_default_na=False, encoding="latin-1")}
    excel = pd.ExcelFile(BytesIO(content))
    return {
        sheet: pd.read_excel(BytesIO(content), sheet_name=sheet, dtype=object)
        for sheet in excel.sheet_names
    }

def safe_json_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    return str(value).strip()

def parse_dates(series):
    """Parsea fechas variadas: Excel serial, datetime, ISO, DD/MM/AAAA y formatos comunes.

    En fechas ambiguas como 04/05/2026 se prioriza día/mes/año, habitual en Perú.
    """
    from datetime import timedelta

    formats = (
        "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y",
        "%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%Y/%m/%d",
        "%Y.%m.%d", "%m/%d/%Y", "%m-%d-%Y",
        "%d/%m/%Y %H:%M:%S", "%d-%m-%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S",
        "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y",
    )

    def parse_one(value):
        if value is None:
            return None
        try:
            if pd.isna(value):
                return None
        except (TypeError, ValueError):
            pass
        if isinstance(value, (pd.Timestamp, datetime, date)):
            return value.date() if isinstance(value, (pd.Timestamp, datetime)) else value
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            number = float(value)
            # Excel's common 1900 date system serial range.
            if 1 <= number <= 100000:
                try:
                    return (datetime(1899, 12, 30) + timedelta(days=number)).date()
                except (OverflowError, ValueError):
                    return None
            return None
        text = str(value).strip()
        if not text or text.lower() in {"nan", "nat", "none", "null", "-", "n/a", "s/f", "sin fecha"}:
            return None
        # Normalize common time suffixes and Spanish month accents.
        text = re.sub(r"\s+", " ", text)
        for fmt in formats:
            try:
                return datetime.strptime(text, fmt).date()
            except (ValueError, TypeError):
                continue
        # General fallback handles timezone-bearing ISO strings and additional variants.
        try:
            parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
            if not pd.isna(parsed):
                return parsed.date()
        except (TypeError, ValueError, OverflowError):
            pass
        return None

    return series.apply(parse_one)


def clean_project_name(value):
    """Standardize whitespace and missing labels so projects are easy to identify."""
    if value is None:
        return "Proyecto sin identificar"
    try:
        if pd.isna(value):
            return "Proyecto sin identificar"
    except (TypeError, ValueError):
        pass
    name = re.sub(r"\s+", " ", str(value)).strip()
    if not name or name.lower() in {"nan", "none", "null", "n/a", "-", "sin proyecto", "no aplica"}:
        return "Proyecto sin identificar"
    return name

def db_connect():
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_key TEXT NOT NULL,
            source_file TEXT NOT NULL,
            source_sheet TEXT NOT NULL,
            source_row INTEGER NOT NULL,
            source_hash TEXT NOT NULL,
            uploaded_by TEXT,
            category TEXT,
            document_type TEXT,
            project TEXT,
            person_asset TEXT,
            identifier TEXT,
            issue_date TEXT,
            due_date TEXT,
            responsible TEXT,
            email TEXT,
            source_status TEXT,
            raw_json TEXT,
            imported_at TEXT NOT NULL
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS sources (
            source_key TEXT PRIMARY KEY,
            source_file TEXT NOT NULL,
            source_hash TEXT NOT NULL,
            uploaded_by TEXT,
            sheets TEXT,
            imported_at TEXT NOT NULL
        )
    """)
    conn.commit()
    return conn

@st.cache_data(ttl=60, show_spinner=False)
def load_records():
    with db_connect() as conn:
        return pd.read_sql_query("SELECT * FROM records ORDER BY source_file, source_sheet, source_row", conn)

@st.cache_data(ttl=60, show_spinner=False)
def load_sources():
    with db_connect() as conn:
        return pd.read_sql_query("SELECT * FROM sources ORDER BY imported_at DESC", conn)

def import_sheet(
    source_file, sheet_name, frame, content_hash, uploader, category,
    due_col, issue_col, project_col, person_col, identifier_col,
    responsible_col, email_col, type_col, status_col
):
    now = datetime.now().isoformat(timespec="seconds")
    source_key = f"{source_file}::{sheet_name}"
    records = []
    df = frame.copy()
    df.columns = [str(c).strip() for c in df.columns]
    due_dates = parse_dates(df[due_col]) if due_col and due_col in df.columns else pd.Series([None] * len(df), index=df.index)
    issue_dates = parse_dates(df[issue_col]) if issue_col and issue_col in df.columns else pd.Series([None] * len(df), index=df.index)

    for idx, row_values in enumerate(df.itertuples(index=False, name=None)):
        row = dict(zip(df.columns, row_values))
        # Ignore entirely empty rows.
        if all(pd.isna(v) or str(v).strip() == "" for v in row.values()):
            continue
        def val(col):
            if not col or col not in df.columns:
                return None
            v = row[col]
            if pd.isna(v):
                return None
            return str(v).strip()
        due = due_dates.iloc[idx] if idx < len(due_dates) else None
        issue = issue_dates.iloc[idx] if idx < len(issue_dates) else None
        raw = {str(k): safe_json_value(v) for k, v in row.items()}
        records.append((
            source_key, source_file, sheet_name, int(idx) + 2, content_hash, uploader,
            category, val(type_col) or category, clean_project_name(val(project_col)), val(person_col),
            val(identifier_col), issue.isoformat() if due_or_date(issue) else None,
            due.isoformat() if due_or_date(due) else None, val(responsible_col),
            val(email_col), val(status_col), json.dumps(raw, ensure_ascii=False), now
        ))

    # Preparar y validar todos los registros ANTES de reemplazar la base activa.
    if not records:
        raise ValueError("La hoja seleccionada no contiene registros válidos. Se conservó la base anterior.")
    with db_connect() as conn:
        # Una sola transacción: si falla cualquier INSERT, SQLite revierte los DELETE.
        conn.execute("DELETE FROM records")
        conn.execute("DELETE FROM sources")
        conn.executemany("""
            INSERT INTO records (
                source_key, source_file, source_sheet, source_row, source_hash,
                uploaded_by, category, document_type, project, person_asset,
                identifier, issue_date, due_date, responsible, email, source_status,
                raw_json, imported_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, records)
        conn.execute("""
            INSERT INTO sources (source_key, source_file, source_hash, uploaded_by, sheets, imported_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (source_key, source_file, content_hash, uploader, sheet_name, now))
        conn.commit()
    return len(records)

def due_or_date(value):
    return value is not None and not pd.isna(value)

def calc_status(value, today, thresholds):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return "Fecha por revisar"
    try:
        due = pd.to_datetime(value).date()
    except Exception:
        return "Fecha por revisar"
    days = (due - today).days
    if days < 0:
        return "Vencido"
    if days == 0:
        return "Vence hoy"
    if thresholds and days <= max(thresholds):
        if days in thresholds:
            return f"Vence en {days} días"
        return "Próximo a vencer"
    return "Vigente según fecha"

def options_for(df, column, include_all=True):
    values = sorted([str(x) for x in df[column].dropna().unique() if str(x).strip()])
    return (["Todos"] + values) if include_all else values

st.markdown("""
<style>
:root { --applus-navy:#173B57; --applus-blue:#245B7A; --applus-light:#F3F6F8; }
.block-container {padding-top:1.4rem; padding-bottom:2rem; max-width:1500px;}
[data-testid="stHeader"] {background:rgba(255,255,255,.92);}
.applus-hero {background:linear-gradient(110deg,#173B57 0%,#245B7A 72%,#397895 100%); padding:1.35rem 1.6rem; border-radius:14px; color:white; margin-bottom:1rem;}
.applus-hero h1 {font-size:1.8rem; font-weight:700; margin:0 0 .35rem 0; color:white;}
.applus-hero p {margin:0; color:#E7EFF4; font-size:.98rem;}
[data-testid="stMetric"] {border-left:4px solid #245B7A; background:#FFFFFF; border:1px solid #E1E7EC; padding:1rem 1.1rem; border-radius:12px; box-shadow:0 2px 8px rgba(20,50,70,.04);}
[data-testid="stMetricLabel"] p {color:#536575; font-weight:600;}
[data-testid="stTabs"] button {font-weight:600;}
[data-testid="stSidebar"] {background:#F5F7F9;}
div.stButton > button, div.stDownloadButton > button {border-radius:8px; font-weight:600;}
</style>
<div class="applus-hero"><h1>APPLUS+ | Control documental</h1><p>Seguimiento de EMO, vencimientos y cumplimiento por proyecto</p></div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("Configuración")
    today = st.date_input("Fecha de evaluación", value=date.today())
    warning_days = st.multiselect("Anticipación de alertas (días)", [1, 7, 15], default=[1, 7, 15], help="Selecciona los avisos que deseas considerar: 1, 7 y/o 15 días antes del vencimiento.")
    st.caption("Las alertas se calculan por fecha. El sistema no valida aptitud médica ni renueva documentos.")
    st.divider()
    st.caption("La base consolidada se guarda en la carpeta `data`. En alojamientos con disco efímero, configura almacenamiento persistente antes del uso compartido real.")

# Cargar y preparar la base consolidada para el panel y los reportes.
# Esta lectura sustituye las antiguas pestañas de carga/archivos/registros;
# no añade esas secciones a la interfaz.
try:
    db_records = load_records()
except Exception as exc:
    st.error("No se pudo leer la base de datos de vencimientos.")
    st.caption(f"Detalle técnico: {type(exc).__name__}")
    db_records = pd.DataFrame()

try:
    all_sources = load_sources()
except Exception as exc:
    st.warning("No se pudo leer el historial de archivos de origen; el panel y los reportes seguirán disponibles.")
    st.caption(f"Detalle técnico del historial: {type(exc).__name__}")
    all_sources = pd.DataFrame(columns=["source_key", "source_file", "source_hash", "uploaded_by", "sheets", "imported_at"])

if db_records.empty:
    all_records = pd.DataFrame(columns=[
        "Categoria", "Proyecto", "Tipo_control", "Persona_activo", "Identificador",
        "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
        "Responsable", "Correo"
    ])
else:
    all_records = pd.DataFrame({
        "Categoria": db_records.get("category", pd.Series(index=db_records.index, dtype=object)).fillna("Otro").astype(str),
        "Proyecto": db_records.get("project", pd.Series(index=db_records.index, dtype=object)).apply(clean_project_name),
        "Tipo_control": db_records.get("document_type", pd.Series(index=db_records.index, dtype=object)).fillna("Sin tipo").astype(str),
        "Persona_activo": db_records.get("person_asset", pd.Series(index=db_records.index, dtype=object)).fillna("").astype(str),
        "Identificador": db_records.get("identifier", pd.Series(index=db_records.index, dtype=object)).fillna("").astype(str),
        "Fecha_emision": pd.to_datetime(db_records.get("issue_date", pd.Series(index=db_records.index, dtype=object)), errors="coerce").dt.date,
        "Fecha_vencimiento": pd.to_datetime(db_records.get("due_date", pd.Series(index=db_records.index, dtype=object)), errors="coerce").dt.date,
        "Responsable": db_records.get("responsible", pd.Series(index=db_records.index, dtype=object)).fillna("").astype(str),
        "Correo": db_records.get("email", pd.Series(index=db_records.index, dtype=object)).fillna("").astype(str),
    })
    def days_remaining(value):
        """Calculate remaining days safely for Python dates, pandas Timestamps and NaT."""
        if value is None or pd.isna(value):
            return None
        try:
            due_date = pd.Timestamp(value).date()
            return (due_date - today).days
        except (TypeError, ValueError, OverflowError):
            return None

    all_records["Dias_restantes"] = all_records["Fecha_vencimiento"].apply(days_remaining)
    all_records["Estado_seguimiento"] = all_records["Fecha_vencimiento"].apply(
        lambda value: calc_status(value, today, warning_days)
    )

# Diagnóstico temporal para identificar en qué etapa se pierden los datos.
# Está dentro de un expander para no recargar el panel ejecutivo.
with st.expander("🔎 Diagnóstico de datos (abrir si los gráficos aparecen vacíos)", expanded=False):
    st.write("Registros cargados desde SQLite:", len(db_records))
    st.write("Registros preparados para el dashboard:", len(all_records))
    if not db_records.empty:
        diagnostic_columns = [c for c in ["project", "due_date", "issue_date", "source_file", "source_sheet"] if c in db_records.columns]
        st.write("Muestra directa de SQLite:")
        st.dataframe(db_records[diagnostic_columns].head(20), use_container_width=True, hide_index=True)
        st.write("Distribución de proyectos guardados en SQLite:")
        if "project" in db_records.columns:
            st.dataframe(
                db_records["project"].fillna("(vacío)").astype(str).value_counts(dropna=False).head(30).rename_axis("Proyecto guardado").reset_index(name="Cantidad"),
                use_container_width=True, hide_index=True
            )
    if not all_records.empty:
        st.write("Muestra de fechas y estados calculados:")
        st.dataframe(
            all_records[[c for c in ["Proyecto", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento"] if c in all_records.columns]].head(20),
            use_container_width=True, hide_index=True
        )
        st.write("Distribución de estados:")
        st.dataframe(
            all_records["Estado_seguimiento"].fillna("(vacío)").astype(str).value_counts(dropna=False).rename_axis("Estado").reset_index(name="Cantidad"),
            use_container_width=True, hide_index=True
        )
        invalid_dates = int(pd.to_datetime(all_records["Fecha_vencimiento"], errors="coerce").isna().sum())
        st.write(f"Fechas de vencimiento no interpretadas: {invalid_dates} de {len(all_records)}")
    else:
        st.warning("No hay registros consolidados en esta base SQLite. Si ya habías importado un archivo, puede que Streamlit Cloud esté usando una base local distinta o que la importación no se haya completado.")

tab_dashboard, tab_export = st.tabs(["📊 Panel general", "📤 Reportes"])

def project_expiry_summary(frame, threshold):
    """Compact summary: counts and percentages are calculated over all rows per project."""
    if frame.empty:
        return pd.DataFrame(columns=[
            "Proyecto", "Total documentos", "Vencidos", "% vencidos",
            "Vencen hoy", "Próximos a vencer", "% próximos a vencer", "Sin fecha válida", "Vigentes", "% cumplimiento"
        ])
    data = frame.copy()
    days = pd.to_numeric(data["Dias_restantes"], errors="coerce")
    data["_vencido"] = days < 0
    data["_hoy"] = days == 0
    data["_proximo"] = days.gt(0) & days.le(threshold)
    summary = data.groupby("Proyecto", dropna=False).agg(
        **{
            "Total documentos": ("Proyecto", "size"),
            "Vencidos": ("_vencido", "sum"),
            "Vencen hoy": ("_hoy", "sum"),
            "Próximos a vencer": ("_proximo", "sum"),
        }
    ).reset_index()
    denominator = summary["Total documentos"].replace(0, pd.NA)
    summary["% vencidos"] = (summary["Vencidos"] / denominator * 100).fillna(0).round(1)
    summary["% próximos a vencer"] = (summary["Próximos a vencer"] / denominator * 100).fillna(0).round(1)
    summary["Sin fecha válida"] = data.groupby("Proyecto", dropna=False)["Dias_restantes"].apply(lambda s: pd.to_numeric(s, errors="coerce").isna().sum()).values
    summary["Vigentes"] = (summary["Total documentos"] - summary["Vencidos"] - summary["Vencen hoy"] - summary["Próximos a vencer"] - summary["Sin fecha válida"]).clip(lower=0)
    valid_denominator = (summary["Total documentos"] - summary["Sin fecha válida"]).replace(0, pd.NA)
    summary["% cumplimiento"] = (summary["Vigentes"] / valid_denominator * 100).fillna(0).round(1)
    return summary.sort_values(["Vencidos", "Próximos a vencer", "Total documentos"], ascending=False).reset_index(drop=True)


def build_filtered_excel(detail, summary, today):
    """Create a clean Excel workbook with a project summary and only the filtered base."""
    output = BytesIO()
    detail_cols = [
        "Proyecto", "Categoria", "Tipo_control", "Persona_activo", "Identificador",
        "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento", "Responsable", "Correo"
    ]
    available_cols = [col for col in detail_cols if col in detail.columns]
    detail_export = detail[available_cols].copy().rename(columns={
        "Categoria": "Categoría", "Tipo_control": "Tipo de control",
        "Persona_activo": "Persona / activo", "Identificador": "Código / placa",
        "Fecha_vencimiento": "Fecha de vencimiento", "Dias_restantes": "Días restantes",
        "Estado_seguimiento": "Estado", "Responsable": "Responsable", "Correo": "Correo"
    })
    summary_export = summary.rename(columns={
        "Proyecto": "Proyecto", "Total documentos": "Total documentos",
        "Vencidos": "Vencidos", "% vencidos": "% vencidos",
        "Vencen hoy": "Vencen hoy", "Próximos a vencer": "Próximos a vencer",
        "% próximos a vencer": "% próximos a vencer"
    }).copy()
    with pd.ExcelWriter(output, engine="openpyxl", date_format="DD/MM/YYYY", datetime_format="DD/MM/YYYY") as writer:
        summary_export.to_excel(writer, index=False, sheet_name="Resumen por proyecto")
        detail_export.to_excel(writer, index=False, sheet_name="Base filtrada")
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.chart import BarChart, Reference
        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="174A7E")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.row_dimensions[1].height = 30
            for col_cells in ws.columns:
                letter = col_cells[0].column_letter
                max_len = max((len(str(c.value or "")) for c in col_cells[:200]), default=10)
                ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 34)
        summary_ws = wb["Resumen por proyecto"]
        for row in range(2, summary_ws.max_row + 1):
            summary_ws.cell(row, 4).number_format = '0.0"%"'
            summary_ws.cell(row, 7).number_format = '0.0"%"'
        detail_ws = wb["Base filtrada"]
        for col_idx, header in enumerate([c.value for c in detail_ws[1]], start=1):
            if header == "Fecha de vencimiento":
                for row in range(2, detail_ws.max_row + 1):
                    detail_ws.cell(row, col_idx).number_format = "DD/MM/YYYY"
        if summary_ws.max_row > 1:
            chart = BarChart()
            chart.type = "bar"
            chart.style = 10
            chart.title = "Vencidos y próximos a vencer por proyecto"
            chart.y_axis.title = "Proyecto"
            chart.x_axis.title = "Cantidad de documentos"
            chart.height = max(7, min(14, summary_ws.max_row * 0.35))
            chart.width = 18
            # Columns C and F are the two action counts in the summary.
            data = Reference(summary_ws, min_col=3, max_col=3, min_row=1, max_row=summary_ws.max_row)
            cats = Reference(summary_ws, min_col=1, min_row=2, max_row=summary_ws.max_row)
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart2 = BarChart()
            chart2.type = "bar"
            chart2.style = 11
            chart2.title = "Próximos a vencer por proyecto"
            chart2.y_axis.title = "Proyecto"
            chart2.x_axis.title = "Cantidad"
            chart2.height = chart.height
            chart2.width = 18
            data2 = Reference(summary_ws, min_col=6, max_col=6, min_row=1, max_row=summary_ws.max_row)
            chart2.add_data(data2, titles_from_data=True)
            chart2.set_categories(cats)
            summary_ws.add_chart(chart, "I2")
            summary_ws.add_chart(chart2, "I20")
    return output.getvalue()


with tab_dashboard:
    with st.expander("＋ Cargar o actualizar base Excel / CSV", expanded=all_records.empty):
        st.caption("Carga una base de datos para incorporarla al seguimiento. Revisa la hoja y las columnas detectadas antes de importar.")
        uploaded_file = st.file_uploader("Selecciona un archivo Excel o CSV", type=["xlsx", "xls", "csv"], key="applus_upload")
        if uploaded_file is not None:
            try:
                sheets = read_workbook(uploaded_file)
                sheet_name = st.selectbox("Hoja a importar", list(sheets.keys()), key="applus_sheet")
                upload_df = sheets[sheet_name].copy()
                upload_df.columns = [str(c).strip() for c in upload_df.columns]
                if upload_df.empty or len(upload_df.columns) == 0:
                    st.warning("El archivo no contiene filas o columnas para importar.")
                else:
                    st.markdown("**Vista previa**")
                    st.dataframe(upload_df.head(5), use_container_width=True, hide_index=True)
                    cols = ["(No asignar)"] + list(upload_df.columns)
                    def guess(key):
                        found = find_column(upload_df.columns, key)
                        return found if found in upload_df.columns else "(No asignar)"
                    map1, map2, map3 = st.columns(3)
                    with map1:
                        up_category = st.selectbox("Categoría", CATEGORIES, index=CATEGORIES.index(infer_category(uploaded_file.name, upload_df.columns)) if infer_category(uploaded_file.name, upload_df.columns) in CATEGORIES else len(CATEGORIES)-1, key="up_category")
                        up_due = st.selectbox("Fecha de vencimiento (obligatoria para alertas)", cols, index=cols.index(guess("due")), key="up_due")
                        up_issue = st.selectbox("Fecha de emisión", cols, index=cols.index(guess("issue")), key="up_issue")
                        up_project = st.selectbox("Proyecto", cols, index=cols.index(guess("project")), key="up_project")
                    with map2:
                        up_person = st.selectbox("Persona / activo", cols, index=cols.index(guess("person")), key="up_person")
                        up_identifier = st.selectbox("DNI / código / placa", cols, index=cols.index(guess("identifier")), key="up_identifier")
                        up_type = st.selectbox("Tipo de documento", cols, index=cols.index(guess("type")), key="up_type")
                    with map3:
                        up_responsible = st.selectbox("Responsable", cols, index=cols.index(guess("responsible")), key="up_responsible")
                        up_email = st.selectbox("Correo", cols, index=cols.index(guess("email")), key="up_email")
                        up_status = st.selectbox("Estado de origen", cols, index=cols.index(guess("status")), key="up_status")
                    if up_due == "(No asignar)":
                        st.warning("Selecciona la columna de fecha de vencimiento para activar las alertas por fecha.")
                    if st.button("Importar hoja seleccionada", type="primary", key="applus_import"):
                        if up_due == "(No asignar)":
                            st.error("No se importó el archivo: falta asignar la fecha de vencimiento.")
                        else:
                            def mapped(value): return None if value == "(No asignar)" else value
                            file_bytes = uploaded_file.getvalue()
                            file_hash = hashlib.sha256(file_bytes).hexdigest()
                            saved_name = re.sub(r"[^A-Za-z0-9._ -]", "_", uploaded_file.name).strip() or "archivo_importado"
                            # Conservar una copia exacta del archivo original, identificada por hash.
                            # Así no se sobrescribe otro archivo con el mismo nombre.
                            suffix = Path(saved_name).suffix
                            stem = Path(saved_name).stem
                            original_path = ORIGINALS_DIR / f"{stem}_{file_hash[:12]}{suffix}"
                            if not original_path.exists():
                                with open(original_path, "wb") as original_file:
                                    original_file.write(file_bytes)
                            # La nueva importación reemplaza TODA la base activa, incluso si
                            # el archivo coincide con uno importado anteriormente.
                            try:
                                count = import_sheet(saved_name, sheet_name, upload_df, file_hash, "Usuario de la aplicación", up_category, mapped(up_due), mapped(up_issue), mapped(up_project), mapped(up_person), mapped(up_identifier), mapped(up_responsible), mapped(up_email), mapped(up_type), mapped(up_status))
                            except Exception as import_exc:
                                st.error(f"No se reemplazó la base: {import_exc}")
                            else:
                                load_records.clear()
                                load_sources.clear()
                                st.cache_data.clear()
                                st.success(f"Base reemplazada: {count} filas activas de {saved_name} · {sheet_name}. El archivo original se conservó en data/originales.")
                                st.rerun()
            except Exception as exc:
                st.error("No se pudo leer el archivo. Verifica que esté en formato Excel o CSV y que no esté dañado.")
                st.caption(f"Detalle técnico: {type(exc).__name__}")
    st.subheader("Indicadores y vencimientos por proyecto")
    if all_records.empty:
        st.info("Aún no hay registros consolidados. No hay registros en la base de datos configurada. Los datos deben estar cargados previamente en el almacenamiento de la aplicación.")
    else:
        threshold = max(warning_days or [30])
        st.caption(f"Próximos a vencer: documentos que vencen dentro de los próximos {threshold} días. Los porcentajes se calculan sobre el total de registros de cada proyecto.")
        fc1, fc2 = st.columns(2)
        dash_categories = options_for(all_records, "Categoria")
        dash_category = fc1.selectbox("Categoría", dash_categories, key="dash_category")
        dash_projects = options_for(all_records, "Proyecto")
        dash_project = fc2.selectbox("Identificar proyecto", dash_projects, key="dash_project", help="Los nombres se limpian de espacios extra y los proyectos sin nombre se muestran como 'Proyecto sin identificar'.")

        dash = all_records.copy()
        if dash_category != "Todos":
            dash = dash[dash["Categoria"].astype(str) == dash_category]
        if dash_project != "Todos":
            dash = dash[dash["Proyecto"].astype(str) == dash_project]

        if dash.empty:
            st.warning("No hay registros para los filtros seleccionados.")
        else:
            summary = project_expiry_summary(dash, threshold)
            total_docs = int(summary["Total documentos"].sum())
            expired_total = int(summary["Vencidos"].sum())
            soon_total = int(summary["Próximos a vencer"].sum())
            due_today_total = int(summary["Vencen hoy"].sum())
            valid_days = pd.to_numeric(dash["Dias_restantes"], errors="coerce")
            valid_mask = valid_days.notna()
            compliance = float((valid_days[valid_mask] > threshold).sum() / valid_mask.sum() * 100) if valid_mask.any() else 0.0
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total documentos", total_docs)
            m2.metric("Cumplimiento", f"{compliance:.1f}%")
            m3.metric("Vencidos", expired_total, f"{(expired_total / total_docs * 100 if total_docs else 0):.1f}% del total")
            m4.metric("Próximos a vencer", soon_total)
            m5.metric("Vencen hoy", due_today_total)

            left_visual, right_visual = st.columns(2)
            with left_visual:
                gauge = go.Figure(go.Indicator(
                    mode="gauge+number", value=compliance,
                    number={"suffix": "%", "font": {"color": "#173B57"}},
                    title={"text": "Cumplimiento documental"},
                    gauge={"axis": {"range": [0, 100]}, "bar": {"color": "#245B7A"},
                           "steps": [{"range": [0, 60], "color": "#F9D7D9"},
                                     {"range": [60, 85], "color": "#FFF0BF"},
                                     {"range": [85, 100], "color": "#D7F0DE"}]}
                ))
                gauge.update_layout(height=280, margin=dict(l=15, r=15, t=55, b=15))
                st.plotly_chart(gauge, use_container_width=True, key="compliance_gauge")
            with right_visual:
                status_counts = pd.DataFrame({
                    "Estado": ["Vencido", "Vence hoy", "Próximo a vencer", "Vigente", "Sin fecha válida"],
                    "Documentos": [expired_total, due_today_total, soon_total,
                                   int((valid_days > threshold).sum()), int((~valid_mask).sum())]
                })
                status_counts = status_counts[status_counts["Documentos"] > 0]
                donut = px.pie(status_counts, names="Estado", values="Documentos", hole=0.65,
                               title="Distribución documental", color="Estado",
                               color_discrete_map={"Vencido": "#C53030", "Vence hoy": "#E37D27",
                                                   "Próximo a vencer": "#E8BC48", "Vigente": "#27855D",
                                                   "Sin fecha válida": "#8394A4"})
                donut.update_layout(height=280, margin=dict(l=10, r=10, t=55, b=15))
                st.plotly_chart(donut, use_container_width=True, key="document_status_donut")

            # Gráfico 1: porcentaje de documentos vencidos por proyecto.
            project_chart = summary.copy()
            project_chart["% vencidos"] = pd.to_numeric(project_chart["% vencidos"], errors="coerce").fillna(0)
            fig_project = px.bar(
                project_chart.sort_values("% vencidos"),
                x="% vencidos", y="Proyecto", orientation="h", text="% vencidos",
                color="% vencidos", color_continuous_scale="Reds",
                title="% de documentos vencidos por proyecto"
            )
            fig_project.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
            fig_project.update_layout(height=500, xaxis_title="% vencidos", yaxis_title="", coloraxis_colorbar_title="% vencidos",
                                      margin=dict(l=10, r=30, t=60, b=20))
            st.plotly_chart(fig_project, use_container_width=True, key="chart_expired_percentage_project")

            # Gráfico 2: porcentaje de documentos vencidos por categoría.
            cat_total = dash.groupby("Categoria", dropna=False).size().reset_index(name="Total")
            cat_expired = (
                dash[dash["Estado_seguimiento"].astype(str) == "Vencido"]
                .groupby("Categoria", dropna=False).size().reset_index(name="Vencidos")
            )
            cat_chart = cat_total.merge(cat_expired, on="Categoria", how="left")
            cat_chart["Vencidos"] = cat_chart["Vencidos"].fillna(0)
            cat_chart["Porcentaje"] = (cat_chart["Vencidos"] / cat_chart["Total"].replace(0, pd.NA) * 100).fillna(0).round(1)
            fig_cat = px.bar(
                cat_chart.sort_values("Porcentaje", ascending=False), x="Categoria", y="Porcentaje",
                text="Porcentaje", color="Porcentaje", color_continuous_scale="Oranges",
                title="% de documentos vencidos por categoría"
            )
            fig_cat.update_traces(texttemplate="%{text:.1f}%", textposition="outside", cliponaxis=False)
            fig_cat.update_layout(height=500, xaxis_title="Categoría", yaxis_title="% vencidos",
                                  showlegend=False, margin=dict(l=10, r=20, t=60, b=25))
            st.plotly_chart(fig_cat, use_container_width=True, key="chart_expired_percentage_category")

            st.markdown("#### Resumen por proyecto")
            summary_display = summary.copy()
            summary_display["% vencidos"] = summary_display["% vencidos"].map(lambda x: f"{x:.1f}%")
            summary_display["% próximos a vencer"] = summary_display["% próximos a vencer"].map(lambda x: f"{x:.1f}%")
            st.dataframe(summary_display, use_container_width=True, hide_index=True, height=min(500, 45 * (len(summary_display) + 1)))

            st.markdown("#### Top 10 proyectos críticos")
            critical = summary.sort_values(["Vencidos", "Próximos a vencer", "Vencen hoy"], ascending=False).head(10).copy()
            critical["Semáforo"] = critical.apply(
                lambda r: "🔴 Vencido" if r["Vencidos"] > 0 else
                ("🟡 Próximo a vencer" if r["Próximos a vencer"] > 0 or r["Vencen hoy"] > 0 else "🟢 Vigente"), axis=1)
            st.dataframe(critical[["Proyecto", "Total documentos", "Vencidos", "Vencen hoy", "Próximos a vencer", "Semáforo"]],
                         use_container_width=True, hide_index=True)
            st.markdown("#### Mapa de calor de riesgos")
            heat = summary.head(20).set_index("Proyecto")[["Vencidos", "Vencen hoy", "Próximos a vencer"]]
            if not heat.empty:
                heat_fig = px.imshow(heat, aspect="auto", text_auto=True, color_continuous_scale="YlOrRd",
                                     labels={"x": "Riesgo", "y": "Proyecto", "color": "Documentos"})
                heat_fig.update_layout(height=max(300, min(700, len(heat) * 33 + 100)),
                                       margin=dict(l=10, r=15, t=30, b=20))
                st.plotly_chart(heat_fig, use_container_width=True, key="project_risk_heatmap")



        st.caption("El panel es interactivo: modifica los filtros para recalcular los indicadores, gráficos y tabla resumen.")

with tab_export:
    st.subheader("Informes ejecutivos y base consolidada")
    if all_records.empty:
        st.info("No hay registros disponibles para generar reportes.")
    else:
        report_filters_1, report_filters_2 = st.columns(2)
        report_cat_options = options_for(all_records, "Categoria")
        report_cat = report_filters_1.selectbox("Filtrar informe por categoría", report_cat_options, key="report_cat")
        report_project_options = options_for(all_records, "Proyecto")
        report_project = report_filters_2.selectbox("Identificar proyecto / área", report_project_options, key="report_project", help="Selecciona un proyecto para ver solo sus vencimientos e indicadores.")
        report_data = all_records.copy()
        if report_cat != "Todos":
            report_data = report_data[report_data["Categoria"].astype(str) == report_cat]
        if report_project != "Todos":
            report_data = report_data[report_data["Proyecto"].astype(str) == report_project]

        st.markdown("#### Resumen del reporte")
        st.caption("Los gráficos ejecutivos de vencimientos están disponibles en el Panel general. El reporte descargable contiene resumen, acciones prioritarias y riesgo por proyecto.")

        export_cols = [
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row", "uploaded_by"
        ]
        # Algunas bases antiguas o importaciones no incluyen metadatos de origen.
        # Completar columnas ausentes evita que el reporte falle con KeyError.
        export_defaults = {
            "Categoria": "Otro", "Tipo_control": "Sin tipo", "Proyecto": "Proyecto sin identificar",
            "Persona_activo": "", "Identificador": "", "Fecha_emision": None,
            "Fecha_vencimiento": None, "Dias_restantes": None, "Estado_seguimiento": "Sin fecha válida",
            "Responsable": "", "Correo": "", "Archivo": "", "Hoja": "",
            "source_row": None, "uploaded_by": ""
        }
        report_data = report_data.copy()
        for column in export_cols:
            if column not in report_data.columns:
                report_data[column] = export_defaults.get(column, "")
        report = report_data.reindex(columns=export_cols).rename(columns={
            "source_row": "Fila_origen", "uploaded_by": "Cargado_por"
        }).copy()
        # Exportación ejecutiva: solo indicadores y documentos que requieren atención.
        days_numeric = pd.to_numeric(report["Dias_restantes"], errors="coerce")
        priority_report = report.loc[days_numeric.notna() & (days_numeric <= threshold)].copy()
        priority_report["Dias_restantes"] = pd.to_numeric(priority_report["Dias_restantes"], errors="coerce")
        priority_report = priority_report.sort_values("Dias_restantes", ascending=True, na_position="last")

        report_days = pd.to_numeric(report["Dias_restantes"], errors="coerce")
        report_total = int(len(report))
        report_expired = int((report_days < 0).sum())
        report_soon = int(((report_days >= 0) & (report_days <= threshold)).sum())
        valid_report_days = report_days.notna()
        report_compliance = (
            float((report_days[valid_report_days] > threshold).sum() / valid_report_days.sum() * 100)
            if valid_report_days.any() else 0.0
        )
        executive = pd.DataFrame({
            "Indicador": ["Total documentos", "Vencidos", "Próximos a vencer", "Cumplimiento"],
            "Valor": [report_total, report_expired, report_soon, f"{report_compliance:.1f}%"]
        })

        risk_data = report.copy()
        risk_data["Dias_restantes"] = pd.to_numeric(risk_data["Dias_restantes"], errors="coerce")
        risk_data["_vencido"] = risk_data["Dias_restantes"] < 0
        risk_data["_hoy"] = risk_data["Dias_restantes"] == 0
        risk_data["_proximo"] = risk_data["Dias_restantes"].gt(0) & risk_data["Dias_restantes"].le(threshold)
        risk_summary = risk_data.groupby("Proyecto", dropna=False).agg(
            **{"Total documentos": ("Proyecto", "size"), "Vencidos": ("_vencido", "sum"),
               "Vencen hoy": ("_hoy", "sum"), "Próximos a vencer": ("_proximo", "sum")}
        ).reset_index().sort_values(["Vencidos", "Próximos a vencer", "Total documentos"], ascending=False)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            executive.to_excel(writer, index=False, sheet_name="Resumen_Ejecutivo", startrow=1)
            priority_report.to_excel(writer, index=False, sheet_name="Acciones_Prioritarias")
            risk_summary.to_excel(writer, index=False, sheet_name="Riesgo_por_Proyecto")

            from openpyxl.chart import BarChart, Reference
            from openpyxl.styles import Font, PatternFill, Alignment
            workbook = writer.book
            red_fill = PatternFill(fill_type="solid", start_color="FFFFB3B3", end_color="FFFFB3B3")
            yellow_fill = PatternFill(fill_type="solid", start_color="FFFFF2CC", end_color="FFFFF2CC")
            header_fill = PatternFill(fill_type="solid", fgColor="174A7E")

            ws_exec = workbook["Resumen_Ejecutivo"]
            ws_exec["A1"] = "RESUMEN EJECUTIVO"
            ws_exec["A1"].font = Font(bold=True, size=16, color="173B57")
            ws_exec.merge_cells("A1:B1")
            for cell in ws_exec[2]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
            ws_exec.column_dimensions["A"].width = 30
            ws_exec.column_dimensions["B"].width = 22
            ws_exec.freeze_panes = "A3"
            ws_exec.sheet_view.showGridLines = False

            for sheet_name in ["Acciones_Prioritarias", "Riesgo_por_Proyecto"]:
                ws = workbook[sheet_name]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                ws.sheet_view.showGridLines = False
                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                for col_cells in ws.columns:
                    letter = col_cells[0].column_letter
                    max_len = max((len(str(c.value or "")) for c in col_cells[:200]), default=10)
                    ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 34)

            ws_priority = workbook["Acciones_Prioritarias"]
            day_col = next((cell.column for cell in ws_priority[1] if cell.value == "Dias_restantes"), None)
            if day_col:
                for row_idx in range(2, ws_priority.max_row + 1):
                    value = ws_priority.cell(row_idx, day_col).value
                    if value is None:
                        continue
                    try:
                        days = float(value)
                    except (TypeError, ValueError):
                        continue
                    fill = red_fill if days < 0 else yellow_fill if days <= 15 else None
                    if fill:
                        for cell in ws_priority[row_idx]:
                            cell.fill = fill

            ws_risk = workbook["Riesgo_por_Proyecto"]
            if ws_risk.max_row > 1:
                chart = BarChart()
                chart.type = "bar"
                chart.style = 10
                chart.title = "Riesgo documental por proyecto"
                chart.y_axis.title = "Proyecto"
                chart.x_axis.title = "Cantidad de documentos"
                data = Reference(ws_risk, min_col=3, max_col=3, min_row=1, max_row=ws_risk.max_row)
                cats = Reference(ws_risk, min_col=1, min_row=2, max_row=ws_risk.max_row)
                chart.add_data(data, titles_from_data=True)
                upcoming = Reference(ws_risk, min_col=5, max_col=5, min_row=1, max_row=ws_risk.max_row)
                chart.add_data(upcoming, titles_from_data=True)
                chart.set_categories(cats)
                chart.height = max(7, min(14, ws_risk.max_row * 0.35))
                chart.width = 18
                ws_risk.add_chart(chart, "G2")

        st.download_button(
            "⬇️ Descargar informe ejecutivo Excel con gráficos",
            data=output.getvalue(),
            file_name=f"informe_corporativo_vencimientos_{today.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("El Excel contiene tres hojas: Resumen Ejecutivo, Acciones Prioritarias y Riesgo por Proyecto. Las acciones incluyen documentos vencidos o que vencen dentro del umbral seleccionado; el semáforo marca vencidos en rojo y próximos en amarillo.")

st.divider()
st.caption(
    "Prototipo administrativo. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el acceso a la aplicación, la retención de archivos y el almacenamiento. "
    "No cargues diagnósticos ni historias clínicas. Los archivos originales se guardan en data/originales."
)
