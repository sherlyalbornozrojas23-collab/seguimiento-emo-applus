import streamlit as st
import pandas as pd
import plotly.express as px
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
    """Parse dates from Excel/CSV robustly, prioritizing DD/MM/YYYY."""
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

        # Excel may store dates as serial numbers (days since 1899-12-30).
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            number = float(value)
            if 20000 <= number <= 80000:
                try:
                    return (pd.Timestamp("1899-12-30") + pd.to_timedelta(number, unit="D")).date()
                except (TypeError, ValueError, OverflowError):
                    return None

        text = str(value).strip()
        if not text or text.lower() in {"nan", "nat", "none", "null", "-", "n/a"}:
            return None

        # Explicit formats avoid ambiguous day/month parsing errors.
        formats = (
            "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y",
            "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d.%m.%y",
            "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
        )
        for fmt in formats:
            try:
                return datetime.strptime(text, fmt).date()
            except (ValueError, TypeError):
                continue

        try:
            parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
            if not pd.isna(parsed):
                return parsed.date()
        except (TypeError, ValueError, OverflowError):
            pass
        return None

    return series.apply(parse_one)

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

def load_records():
    with db_connect() as conn:
        return pd.read_sql_query("SELECT * FROM records ORDER BY source_file, source_sheet, source_row", conn)

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

    for idx, row in df.iterrows():
        # Ignore entirely empty rows.
        if all(pd.isna(v) or str(v).strip() == "" for v in row.tolist()):
            continue
        def val(col):
            if not col or col not in df.columns:
                return None
            v = row[col]
            if pd.isna(v):
                return None
            return str(v).strip()
        due = due_dates.loc[idx] if idx in due_dates.index else None
        issue = issue_dates.loc[idx] if idx in issue_dates.index else None
        raw = {str(k): safe_json_value(v) for k, v in row.to_dict().items()}
        records.append((
            source_key, source_file, sheet_name, int(idx) + 2, content_hash, uploader,
            category, val(type_col) or category, val(project_col), val(person_col),
            val(identifier_col), issue.isoformat() if due_or_date(issue) else None,
            due.isoformat() if due_or_date(due) else None, val(responsible_col),
            val(email_col), val(status_col), json.dumps(raw, ensure_ascii=False), now
        ))

    with db_connect() as conn:
        # Re-importing a sheet replaces its previously imported rows, not other sheets/files.
        conn.execute("DELETE FROM records WHERE source_key = ?", (source_key,))
        conn.execute("DELETE FROM sources WHERE source_key = ?", (source_key,))
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

st.title("📅 Sistema corporativo de seguimiento de vencimientos")
st.caption("Plataforma para consolidar vencimientos de distintas bases Excel, sin exigir una plantilla única ni modificar los archivos originales.")

with st.sidebar:
    st.header("Configuración")
    today = st.date_input("Fecha de evaluación", value=date.today())
    warning_days = st.multiselect("Anticipación de alertas (días)", [7, 15, 30, 45, 60, 90], default=[7, 15, 30])
    uploader = st.text_input("Nombre del asistente que carga", placeholder="Ej. Sherly / Administración")
    st.caption("Las alertas se calculan por fecha. El sistema no valida aptitud médica ni renueva documentos.")
    st.divider()
    st.caption("La base consolidada se guarda en la carpeta `data`. En alojamientos con disco efímero, configura almacenamiento persistente antes del uso compartido real.")

tab_dashboard, tab_upload, tab_sources, tab_records, tab_export = st.tabs([
    "📊 Panel general", "📥 Cargar Excel", "🗂️ Archivos cargados",
    "🔎 Registros", "📤 Reportes"
])

with tab_upload:
    st.subheader("Cargar una o varias bases de datos")
    st.write("Puedes seleccionar varios archivos Excel o CSV. Se leerán todas las hojas; cada hoja se puede mapear a sus propias columnas.")
    uploads = st.file_uploader(
        "Selecciona archivos .xlsx, .xlsm o .csv",
        type=["xlsx", "xlsm", "csv"],
        accept_multiple_files=True,
        key="multi_upload"
    )
    if uploads:
        for uploaded in uploads:
            file_bytes = uploaded.getvalue()
            file_hash = hashlib.sha256(file_bytes).hexdigest()
            st.markdown(f"#### 📄 {uploaded.name}")
            try:
                sheets = read_workbook(uploaded)
            except Exception as e:
                st.error(f"No se pudo leer {uploaded.name}: {e}")
                continue

            for sheet_name, frame in sheets.items():
                if frame is None or frame.empty:
                    st.warning(f"La hoja '{sheet_name}' está vacía.")
                    continue
                frame = frame.copy()
                frame.columns = [str(c).strip() for c in frame.columns]
                cols = list(frame.columns)
                inferred_category = infer_category(uploaded.name, cols)
                with st.expander(f"Hoja: {sheet_name} · {len(frame)} filas · categoría sugerida: {inferred_category}", expanded=(len(sheets) == 1)):
                    st.dataframe(frame.head(8), use_container_width=True, hide_index=True)
                    st.caption("Confirma o ajusta las columnas antes de importar. Si una columna no aplica, selecciona «No disponible».")

                    def choose(label, key, found, allow_none=True):
                        choices = (["— No disponible —"] if allow_none else []) + cols
                        default = found if found in cols else ("— No disponible —" if allow_none else cols[0])
                        return st.selectbox(label, choices, index=choices.index(default), key=f"{uploaded.name}_{sheet_name}_{key}")

                    category = st.selectbox(
                        "Categoría del control",
                        CATEGORIES,
                        index=CATEGORIES.index(inferred_category) if inferred_category in CATEGORIES else CATEGORIES.index("Otro"),
                        key=f"{uploaded.name}_{sheet_name}_category"
                    )
                    due_found = find_column(cols, "due")
                    issue_found = find_column(cols, "issue")
                    project_found = find_column(cols, "project")
                    person_found = find_column(cols, "person")
                    id_found = find_column(cols, "identifier")
                    responsible_found = find_column(cols, "responsible")
                    email_found = find_column(cols, "email")
                    type_found = find_column(cols, "type")
                    status_found = find_column(cols, "status")

                    due_col = choose("Columna de fecha de vencimiento (obligatoria para alertas)", "due", due_found)
                    issue_col = choose("Columna de fecha de emisión / inicio (opcional)", "issue", issue_found)
                    project_col = choose("Columna de proyecto / área (opcional)", "project", project_found)
                    person_col = choose("Columna de persona / activo (opcional)", "person", person_found)
                    id_col = choose("Columna de DNI / placa / código (opcional)", "identifier", id_found)
                    responsible_col = choose("Columna de responsable (opcional)", "responsible", responsible_found)
                    email_col = choose("Columna de correo (opcional)", "email", email_found)
                    type_col = choose("Columna de tipo de documento (opcional)", "type", type_found)
                    status_col = choose("Columna de estado existente (opcional)", "status", status_found)

                    if not uploader.strip():
                        st.warning("Escribe el nombre del asistente en la barra lateral para identificar quién realizó la carga.")
                    if due_col == "— No disponible —":
                        st.warning("No se ha asignado una columna de vencimiento. Los registros se importarán como «Fecha por revisar» y no generarán alertas de fecha hasta corregir la asignación.")
                    if st.button(f"Importar hoja '{sheet_name}'", key=f"import_{uploaded.name}_{sheet_name}"):
                        saved_path = ORIGINALS_DIR / uploaded.name
                        # Preserve original bytes for audit/reference. Same-name files are replaced on disk.
                        saved_path.write_bytes(file_bytes)
                        def selected(col):
                            return None if col == "— No disponible —" else col
                        n = import_sheet(
                            uploaded.name, sheet_name, frame, file_hash, uploader.strip(),
                            category, selected(due_col), selected(issue_col), selected(project_col),
                            selected(person_col), selected(id_col), selected(responsible_col),
                            selected(email_col), selected(type_col), selected(status_col)
                        )
                        st.success(f"Se importaron {n} registros de {uploaded.name} / {sheet_name}.")
                        st.rerun()

with db_connect() as conn:
    all_records = pd.read_sql_query("SELECT * FROM records ORDER BY imported_at DESC", conn)
    all_sources = pd.read_sql_query("SELECT * FROM sources ORDER BY imported_at DESC", conn)

if not all_records.empty:
    all_records["Fecha_vencimiento"] = pd.to_datetime(all_records["due_date"], errors="coerce").dt.date
    all_records["Fecha_emision"] = pd.to_datetime(all_records["issue_date"], errors="coerce").dt.date
    all_records["Dias_restantes"] = all_records["Fecha_vencimiento"].apply(
        lambda d: (d - today).days if pd.notna(d) else None
    )
    all_records["Estado_seguimiento"] = all_records["Fecha_vencimiento"].apply(
        lambda d: calc_status(d.isoformat() if pd.notna(d) else None, today, warning_days)
    )
    all_records["Proyecto"] = all_records["project"].fillna("").replace("", "No especificado")
    all_records["Categoria"] = all_records["category"].fillna("Otro")
    all_records["Tipo_control"] = all_records["document_type"].fillna("").astype(str).str.strip()
    empty_type = all_records["Tipo_control"].eq("")
    all_records.loc[empty_type, "Tipo_control"] = all_records.loc[empty_type, "Categoria"].astype(str)
    all_records["Persona_activo"] = all_records["person_asset"].fillna("")
    all_records["Identificador"] = all_records["identifier"].fillna("")
    all_records["Responsable"] = all_records["responsible"].fillna("")
    all_records["Correo"] = all_records["email"].fillna("")
    all_records["Archivo"] = all_records["source_file"].fillna("")
    all_records["Hoja"] = all_records["source_sheet"].fillna("")
else:
    all_records = pd.DataFrame(columns=[
        "id", "source_file", "source_sheet", "source_row", "category",
        "document_type", "project", "person_asset", "identifier", "responsible",
        "email", "due_date", "Fecha_vencimiento", "Dias_restantes",
        "Estado_seguimiento", "Proyecto", "Categoria", "Tipo_control",
        "Persona_activo", "Identificador", "Responsable", "Correo", "Archivo", "Hoja"
    ])

def project_expiry_summary(frame, threshold):
    """Compact summary: counts and percentages are calculated over all rows per project."""
    if frame.empty:
        return pd.DataFrame(columns=[
            "Proyecto", "Total documentos", "Vencidos", "% vencidos",
            "Vencen hoy", "Próximos a vencer", "% próximos a vencer"
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
    st.subheader("Vencimientos por proyecto")
    if all_records.empty:
        st.info("Aún no hay registros consolidados. Entra en «Cargar Excel» para importar la primera base.")
    else:
        threshold = max(warning_days or [30])
        st.caption(f"Próximos a vencer: documentos que vencen dentro de los próximos {threshold} días. Los porcentajes se calculan sobre el total de registros de cada proyecto.")
        fc1, fc2 = st.columns(2)
        dash_categories = options_for(all_records, "Categoria")
        dash_category = fc1.selectbox("Categoría", dash_categories, key="dash_category")
        dash_projects = options_for(all_records, "Proyecto")
        dash_project = fc2.selectbox("Proyecto", dash_projects, key="dash_project")

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
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Documentos filtrados", total_docs)
            m2.metric("Vencidos", expired_total, f"{(expired_total / total_docs * 100 if total_docs else 0):.1f}% del total")
            m3.metric("Próximos a vencer", soon_total, f"{(soon_total / total_docs * 100 if total_docs else 0):.1f}% del total")
            m4.metric("Vencen hoy", due_today_total)

            chart_data = summary.melt(
                id_vars="Proyecto", value_vars=["Vencidos", "Próximos a vencer"],
                var_name="Estado", value_name="Cantidad"
            )
            fig_project = px.bar(
                chart_data, y="Proyecto", x="Cantidad", color="Estado", barmode="group",
                text="Cantidad", orientation="h", title="Cantidad de vencidos y próximos a vencer por proyecto",
                hover_data={"Cantidad": True, "Proyecto": True, "Estado": True}
            )
            fig_project.update_layout(height=max(330, min(650, 38 * len(summary) + 130)), xaxis_title="Cantidad de documentos", yaxis_title="", legend_title="", margin=dict(l=10, r=20, t=60, b=20))
            fig_project.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(fig_project, use_container_width=True, key="chart_project_expiry_summary")

            st.markdown("#### Resumen por proyecto")
            summary_display = summary.copy()
            summary_display["% vencidos"] = summary_display["% vencidos"].map(lambda x: f"{x:.1f}%")
            summary_display["% próximos a vencer"] = summary_display["% próximos a vencer"].map(lambda x: f"{x:.1f}%")
            st.dataframe(summary_display, use_container_width=True, hide_index=True, height=min(500, 45 * (len(summary_display) + 1)))

            st.markdown("#### Descargar base filtrada")
            st.caption("La descarga respeta los filtros de categoría y proyecto seleccionados. El Excel contiene una hoja resumen por proyecto y otra con el detalle de los registros filtrados.")
            detail_cols = [
                "Proyecto", "Categoria", "Tipo_control", "Persona_activo", "Identificador",
                "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento", "Responsable", "Correo"
            ]
            detail = dash[detail_cols].copy()
            excel_bytes = build_filtered_excel(dash, summary, today)
            csv_detail = detail.rename(columns={
                "Categoria": "Categoría", "Tipo_control": "Tipo de control",
                "Persona_activo": "Persona / activo", "Identificador": "Código / placa",
                "Fecha_vencimiento": "Fecha de vencimiento", "Dias_restantes": "Días restantes",
                "Estado_seguimiento": "Estado"
            })
            download1, download2 = st.columns(2)
            with download1:
                st.download_button(
                    "⬇️ Descargar Excel filtrado (resumen + base)", data=excel_bytes,
                    file_name=f"vencimientos_filtrados_{today.isoformat()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_dashboard_excel", use_container_width=True
                )
            with download2:
                st.download_button(
                    "⬇️ Descargar base filtrada CSV", data=csv_detail.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"base_filtrada_{today.isoformat()}.csv", mime="text/csv",
                    key="download_dashboard_csv", use_container_width=True
                )
            with st.expander("Ver detalle de registros filtrados"):
                detail_display = dash[[
                    "Proyecto", "Categoria", "Tipo_control", "Persona_activo", "Identificador",
                    "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento"
                ]].rename(columns={
                    "Categoria": "Categoría", "Tipo_control": "Tipo de control",
                    "Persona_activo": "Persona / activo", "Identificador": "Código / placa",
                    "Fecha_vencimiento": "Fecha de vencimiento", "Dias_restantes": "Días restantes",
                    "Estado_seguimiento": "Estado"
                })
                st.dataframe(detail_display, use_container_width=True, hide_index=True)

        st.caption("El panel es interactivo: modifica los filtros para recalcular las cantidades, porcentajes, gráfico y archivos descargables.")

with tab_sources:
    st.subheader("Archivos y hojas importados")
    if all_sources.empty:
        st.info("Todavía no se han importado archivos.")
    else:
        sources_view = all_sources.rename(columns={
            "source_file": "Archivo", "source_hash": "Huella del archivo",
            "uploaded_by": "Cargado por", "sheets": "Hoja",
            "imported_at": "Fecha de importación"
        })
        st.dataframe(sources_view[["Archivo", "Hoja", "Cargado por", "Fecha de importación", "Huella del archivo"]], use_container_width=True, hide_index=True)
        st.caption("La huella permite distinguir versiones del archivo. Al importar de nuevo una hoja con el mismo nombre, sus registros anteriores se reemplazan por la versión nueva.")

with tab_records:
    st.subheader("Buscar registros consolidados")
    if all_records.empty:
        st.info("No hay registros para mostrar.")
    else:
        f1, f2, f3 = st.columns(3)
        cat_options = options_for(all_records, "Categoria")
        selected_cat = f1.selectbox("Categoría", cat_options)
        project_options = options_for(all_records, "Proyecto")
        selected_project = f2.selectbox("Proyecto / área", project_options)
        search = f3.text_input("Buscar por persona, placa, código, archivo o responsable").strip().lower()
        view = all_records.copy()
        if selected_cat != "Todos":
            view = view[view["Categoria"].astype(str) == selected_cat]
        if selected_project != "Todos":
            view = view[view["Proyecto"].astype(str) == selected_project]
        if search:
            search_cols = ["Persona_activo", "Identificador", "Archivo", "Responsable", "Tipo_control", "Proyecto"]
            mask = False
            for col in search_cols:
                mask = mask | view[col].astype(str).str.lower().str.contains(search, na=False)
            view = view[mask]
        display = view[[
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row"
        ]].rename(columns={"source_row": "Fila origen"})
        st.dataframe(display, use_container_width=True, hide_index=True)

with tab_export:
    st.subheader("Informes ejecutivos y base consolidada")
    if all_records.empty:
        st.info("Importa al menos una hoja para habilitar reportes.")
    else:
        report_filters_1, report_filters_2 = st.columns(2)
        report_cat_options = options_for(all_records, "Categoria")
        report_cat = report_filters_1.selectbox("Filtrar informe por categoría", report_cat_options, key="report_cat")
        report_project_options = options_for(all_records, "Proyecto")
        report_project = report_filters_2.selectbox("Filtrar informe por proyecto / área", report_project_options, key="report_project")
        report_data = all_records.copy()
        if report_cat != "Todos":
            report_data = report_data[report_data["Categoria"].astype(str) == report_cat]
        if report_project != "Todos":
            report_data = report_data[report_data["Proyecto"].astype(str) == report_project]

        status_summary = report_data.groupby("Estado_seguimiento").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        cat_summary = report_data.groupby("Categoria").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        project_summary = report_data.groupby("Proyecto").size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.markdown("#### Vista gráfica del informe")
        report_chart_1, report_chart_2 = st.columns(2)
        with report_chart_1:
            fig_report_status = px.bar(status_summary, x="Estado_seguimiento", y="Cantidad", color="Estado_seguimiento", text="Cantidad", title="Resumen por estado")
            fig_report_status.update_layout(xaxis_title="Estado", yaxis_title="Cantidad", xaxis_tickangle=-25, height=360)
            fig_report_status.update_traces(textposition="outside")
            st.plotly_chart(fig_report_status, use_container_width=True, key="report_status_chart")
        with report_chart_2:
            fig_report_cat = px.bar(cat_summary, x="Categoria", y="Cantidad", color="Categoria", text="Cantidad", title="Resumen por categoría")
            fig_report_cat.update_layout(xaxis_title="Categoría", yaxis_title="Cantidad", xaxis_tickangle=-25, height=360, showlegend=False)
            fig_report_cat.update_traces(textposition="outside")
            st.plotly_chart(fig_report_cat, use_container_width=True, key="report_category_chart")

        export_cols = [
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row", "uploaded_by"
        ]
        report = report_data[export_cols].rename(columns={
            "source_row": "Fila_origen", "uploaded_by": "Cargado_por"
        }).copy()
        priority_report = report[
            report["Estado_seguimiento"].isin(["Vencido", "Vence hoy", "Próximo a vencer"])
            | report["Estado_seguimiento"].str.startswith("Vence en", na=False)
        ]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Base_consolidada")
            priority_report.to_excel(writer, index=False, sheet_name="Acciones_prioritarias")
            cat_summary.to_excel(writer, index=False, sheet_name="Resumen_categoria")
            status_summary.to_excel(writer, index=False, sheet_name="Resumen_estado")
            project_summary.to_excel(writer, index=False, sheet_name="Resumen_proyecto")
            if not all_sources.empty:
                all_sources.to_excel(writer, index=False, sheet_name="Archivos_importados")
            # Add bar charts to the Excel workbook so the report includes visual summaries.
            from openpyxl.chart import BarChart, Reference
            from openpyxl.styles import Font, PatternFill, Alignment
            workbook = writer.book
            for sheet_name, title in [("Resumen_categoria", "Vencimientos por categoría"), ("Resumen_estado", "Vencimientos por estado"), ("Resumen_proyecto", "Vencimientos por proyecto")]:
                ws = workbook[sheet_name]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for cell in ws[1]:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="174A7E")
                    cell.alignment = Alignment(horizontal="center")
                for col_cells in ws.columns:
                    letter = col_cells[0].column_letter
                    max_len = max((len(str(c.value or "")) for c in col_cells), default=10)
                    ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 36)
                if ws.max_row > 1:
                    chart = BarChart()
                    chart.type = "bar"
                    chart.style = 10
                    chart.title = title
                    chart.y_axis.title = "Cantidad"
                    chart.x_axis.title = "Grupo"
                    data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
                    cats = Reference(ws, min_col=1, min_row=2, max_row=ws.max_row)
                    chart.add_data(data, titles_from_data=True)
                    chart.set_categories(cats)
                    chart.height = 7
                    chart.width = 15
                    ws.add_chart(chart, "D2")
            for sheet_name in ["Base_consolidada", "Acciones_prioritarias", "Archivos_importados"]:
                if sheet_name in workbook.sheetnames:
                    ws = workbook[sheet_name]
                    ws.freeze_panes = "A2"
                    ws.auto_filter.ref = ws.dimensions
                    for cell in ws[1]:
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill("solid", fgColor="174A7E")
                    for col_cells in ws.columns:
                        letter = col_cells[0].column_letter
                        max_len = max((len(str(c.value or "")) for c in col_cells[:100]), default=10)
                        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 36)

        st.download_button(
            "⬇️ Descargar informe ejecutivo Excel con gráficos",
            data=output.getvalue(),
            file_name=f"informe_corporativo_vencimientos_{today.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.caption("El Excel incluye la base consolidada, acciones prioritarias, resúmenes y gráficos de barras. Los gráficos del panel web responden a los filtros; los del Excel reflejan los datos incluidos al descargarlo.")

st.divider()
st.caption(
    "Prototipo administrativo. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el acceso a la aplicación, la retención de archivos y el almacenamiento. "
    "No cargues diagnósticos ni historias clínicas. Los archivos originales se guardan en data/originales."
)
