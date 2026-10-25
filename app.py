import streamlit as st
import pandas as pd
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
    # Handles Excel dates, ISO dates and common day/month/year formats.
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce").dt.date
    try:
        parsed = pd.to_datetime(series, errors="coerce", dayfirst=True, format="mixed")
    except (TypeError, ValueError):
        parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return parsed.dt.date

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
    all_records["Tipo_control"] = all_records["document_type"].fillna("").replace("", all_records["Categoria"])
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

with tab_dashboard:
    st.subheader("Resumen corporativo")
    if all_records.empty:
        st.info("Aún no hay registros consolidados. Entra en «Cargar Excel» para importar la primera base.")
    else:
        total = len(all_records)
        expired = int((all_records["Dias_restantes"].apply(lambda x: pd.notna(x) and x < 0)).sum())
        today_count = int((all_records["Dias_restantes"] == 0).sum())
        soon = int((all_records["Dias_restantes"].apply(lambda x: pd.notna(x) and 0 < x <= max(warning_days or [30]))).sum())
        unknown = int(all_records["Fecha_vencimiento"].isna().sum())
        a, b, c, d, e = st.columns(5)
        a.metric("Registros", total)
        b.metric("Vencidos", expired)
        c.metric("Vencen hoy", today_count)
        d.metric("Próximos a vencer", soon)
        e.metric("Fecha por revisar", unknown)

        st.markdown("#### Vencimientos por categoría")
        by_category = all_records.groupby("Categoria", dropna=False).size().reset_index(name="Cantidad").sort_values("Cantidad", ascending=False)
        st.dataframe(by_category, use_container_width=True, hide_index=True)

        st.markdown("#### Acciones prioritarias")
        priority = all_records[
            all_records["Dias_restantes"].apply(
                lambda x: pd.notna(x) and (x < 0 or x == 0 or (x > 0 and x <= max(warning_days or [30])))
            )
        ].sort_values("Dias_restantes", na_position="last")
        if priority.empty:
            st.success("No hay vencimientos prioritarios según la anticipación configurada.")
        else:
            st.dataframe(
                priority[["Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
                          "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
                          "Responsable", "Correo", "Archivo", "Hoja", "source_row"]],
                use_container_width=True, hide_index=True
            )
        st.caption("Los estados son cálculos administrativos basados en fechas. Confirma las obligaciones y periodicidades con el área responsable.")

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
    st.subheader("Exportar reportes")
    if all_records.empty:
        st.info("Importa al menos una hoja para habilitar reportes.")
    else:
        export_cols = [
            "Categoria", "Tipo_control", "Proyecto", "Persona_activo", "Identificador",
            "Fecha_emision", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento",
            "Responsable", "Correo", "Archivo", "Hoja", "source_row", "uploaded_by"
        ]
        report = all_records[export_cols].rename(columns={
            "source_row": "Fila_origen", "uploaded_by": "Cargado_por"
        }).copy()
        priority_report = report[
            report["Estado_seguimiento"].isin(["Vencido", "Vence hoy", "Próximo a vencer"])
            | report["Estado_seguimiento"].str.startswith("Vence en", na=False)
        ]
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            report.to_excel(writer, index=False, sheet_name="Todos_los_vencimientos")
            priority_report.to_excel(writer, index=False, sheet_name="Acciones_prioritarias")
            if not all_sources.empty:
                all_sources.to_excel(writer, index=False, sheet_name="Archivos_importados")
        st.download_button(
            "⬇️ Descargar reporte corporativo Excel",
            data=output.getvalue(),
            file_name=f"reporte_corporativo_vencimientos_{today.isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.divider()
st.caption(
    "Prototipo administrativo. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el acceso a la aplicación, la retención de archivos y el almacenamiento. "
    "No cargues diagnósticos ni historias clínicas. Los archivos originales se guardan en data/originales."
)
