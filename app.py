import json
import re
import sqlite3
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Seguimiento de vencimientos APPLUS+", page_icon="📅", layout="wide")

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
ORIGINALS_DIR = DATA_DIR / "originales"
DATA_DIR.mkdir(exist_ok=True)
ORIGINALS_DIR.mkdir(exist_ok=True)
DB_PATH = DATA_DIR / "vencimientos.db"

ALERT_OPTIONS = [1, 7, 15]

ALIASES = {
    "due": [
        "fecha vencimiento", "fecha de vencimiento", "vencimiento", "vence",
        "vigente hasta", "vigencia hasta", "fecha caducidad", "fecha de caducidad",
        "caducidad", "fecha fin", "fin de vigencia", "expiracion", "expiration",
        "expiry date", "expiration date", "valid until", "fecha termino", "fecha término",
        "proxima revision", "próxima revisión", "proximo mantenimiento", "próximo mantenimiento"
    ],
    "issue": [
        "fecha emision", "fecha de emision", "fecha emisión", "fecha de emisión",
        "fecha inicio", "fecha de inicio", "fecha expedicion", "fecha de expedicion",
        "fecha expedición", "fecha de expedición", "fecha examen", "fecha emo"
    ],
    "project": ["proyecto", "project", "obra", "servicio", "contrato/proyecto"],
    "person": [
        "nombres", "nombre completo", "trabajador", "colaborador", "empleado",
        "conductor", "responsable del equipo", "titular", "persona", "razon social",
        "razón social"
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
    "type": ["tipo documento", "documento", "tipo de control", "control", "certificado", "tipo"],
    "status": ["estado", "estado seguimiento", "situacion", "situación"]
}


def norm(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    replacements = {"á": "a", "é": "e", "í": "i", "ó": "o", "ú": "u", "ü": "u", "ñ": "n"}
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"[_\-./]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def find_column(columns, key):
    aliases = [norm(x) for x in ALIASES.get(key, [])]
    normalized = {c: norm(c) for c in columns}
    for col, ncol in normalized.items():
        if ncol in aliases:
            return col
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


def safe_json_value(value):
    if value is None or pd.isna(value):
        return None
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.isoformat()
    return str(value).strip()


def parse_excel_numeric_date(value):
    if value is None:
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # Excel serial dates in 1900 system.
        if pd.notna(value):
            try:
                return pd.to_datetime("1899-12-30") + pd.to_timedelta(float(value), unit="D")
            except Exception:
                pass
    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none", "null", "-", "n/a", "s/f", "sin fecha"}:
        return pd.NaT
    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def parse_date_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        return value
    if isinstance(value, datetime):
        return pd.Timestamp(value)
    if isinstance(value, date):
        return pd.Timestamp(value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return parse_excel_numeric_date(value)
    text = str(value).strip()
    if not text or text.lower() in {"nan", "nat", "none", "null", "-", "n/a", "s/f", "sin fecha"}:
        return pd.NaT
    text = re.sub(r"\s+", " ", text)
    # Rule for Peru: prefer day-first when ambiguous.
    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if pd.isna(parsed):
        parsed = pd.to_datetime(text, errors="coerce")
    return parsed


def parse_dates(series):
    return series.apply(parse_date_value)


def clean_project_name(value):
    if value is None:
        return "Proyecto sin identificar"
    try:
        if pd.isna(value):
            return "Proyecto sin identificar"
    except Exception:
        pass
    cleaned = re.sub(r"\s+", " ", str(value)).strip()
    if not cleaned or cleaned.lower() in {"nan", "none", "null", "n/a", "-", "sin proyecto", "no aplica"}:
        return "Proyecto sin identificar"
    return cleaned


def db_connect():
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.execute(
        """
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
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS sources (
            source_key TEXT PRIMARY KEY,
            source_file TEXT NOT NULL,
            source_hash TEXT NOT NULL,
            uploaded_by TEXT,
            sheets TEXT,
            imported_at TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn


def load_records():
    try:
        with db_connect() as conn:
            df = pd.read_sql_query("SELECT * FROM records ORDER BY source_file, source_sheet, source_row", conn)
    except Exception:
        return pd.DataFrame()
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()


def load_sources():
    try:
        with db_connect() as conn:
            df = pd.read_sql_query("SELECT * FROM sources ORDER BY imported_at DESC", conn)
    except Exception:
        return pd.DataFrame()
    return df if isinstance(df, pd.DataFrame) else pd.DataFrame()


def normalize_records(frame):
    if frame is None or frame.empty:
        return pd.DataFrame()

    df = frame.copy()
    required = [
        "category", "document_type", "project", "person_asset", "identifier",
        "issue_date", "due_date", "responsible", "email", "source_status",
        "source_file", "source_sheet", "source_row", "uploaded_by"
    ]
    for col in required:
        if col not in df.columns:
            df[col] = pd.NA

    df["Categoria"] = df["category"].fillna("Otro").astype(str)
    df["Tipo_control"] = df["document_type"].fillna("Sin tipo").astype(str)
    df["Proyecto"] = df["project"].fillna("Proyecto sin identificar").map(clean_project_name).astype(str)
    df["Persona_activo"] = df["person_asset"].fillna("Sin dato").astype(str)
    df["Identificador"] = df["identifier"].fillna("Sin dato").astype(str)
    df["Responsable"] = df["responsible"].fillna("Sin responsable").astype(str)
    df["Correo"] = df["email"].fillna("Sin correo").astype(str)
    df["Archivo"] = df["source_file"].fillna("Sin archivo").astype(str)
    df["Hoja"] = df["source_sheet"].fillna("Sin hoja").astype(str)
    df["Fecha_emision"] = pd.to_datetime(df["issue_date"], errors="coerce")
    df["Fecha_vencimiento"] = pd.to_datetime(df["due_date"], errors="coerce")
    df["Dias_restantes"] = (df["Fecha_vencimiento"] - pd.Timestamp(date.today())).dt.days
    df["Estado_seguimiento"] = df["source_status"].fillna("Fecha pendiente").astype(str)
    df["Estado_seguimiento"] = df.apply(
        lambda row: calc_status(row["Fecha_vencimiento"], date.today(), ALERT_OPTIONS)
        if pd.notna(row["Fecha_vencimiento"]) else "Fecha pendiente",
        axis=1,
    )

    df["source_row"] = pd.to_numeric(df["source_row"], errors="coerce").fillna(0).astype(int)
    df["uploaded_by"] = df["uploaded_by"].fillna("Sistema").astype(str)
    df["Dias_restantes"] = pd.to_numeric(df["Dias_restantes"], errors="coerce").fillna(0)
    return df


def read_workbook(uploaded_file):
    if uploaded_file is None:
        raise ValueError("No se recibió ningún archivo.")
    name = uploaded_file.name.lower()
    content = uploaded_file.getvalue()
    if name.endswith(".csv"):
        for encoding in ("utf-8-sig", "latin-1", "cp1252"):
            try:
                return {"CSV": pd.read_csv(BytesIO(content), dtype=str, keep_default_na=False, encoding=encoding)}
            except Exception:
                continue
        raise ValueError(f"No se pudo leer el CSV '{uploaded_file.name}'. Verifica encoding o estructura.")

    try:
        excel = pd.ExcelFile(BytesIO(content))
    except Exception as exc:
        raise ValueError(f"No se pudo abrir el archivo Excel '{uploaded_file.name}'.") from exc

    sheets = {}
    for sheet in excel.sheet_names:
        try:
            sheets[sheet] = pd.read_excel(BytesIO(content), sheet_name=sheet, dtype=object)
        except Exception as exc:
            raise ValueError(f"La hoja '{sheet}' no se pudo leer: {exc}") from exc
    return sheets


def import_sheet(source_file, sheet_name, frame, content_hash, uploader, category, due_col, issue_col, project_col, person_col, identifier_col, responsible_col, email_col, type_col, status_col):
    now = datetime.now().isoformat(timespec="seconds")
    source_key = f"{source_file}::{sheet_name}"
    records = []
    df = frame.copy()
    df.columns = [str(c).strip() for c in df.columns]

    due_dates = parse_dates(df[due_col]) if due_col and due_col in df.columns else pd.Series([pd.NaT] * len(df), index=df.index)
    issue_dates = parse_dates(df[issue_col]) if issue_col and issue_col in df.columns else pd.Series([pd.NaT] * len(df), index=df.index)

    for idx, row in df.iterrows():
        if all(pd.isna(v) or str(v).strip() == "" for v in row.tolist()):
            continue

        def val(col_name):
            if not col_name or col_name not in df.columns:
                return None
            value = row[col_name]
            if pd.isna(value):
                return None
            return str(value).strip()

        due = due_dates.loc[idx] if idx in due_dates.index else pd.NaT
        issue = issue_dates.loc[idx] if idx in issue_dates.index else pd.NaT

        raw = {str(k): safe_json_value(v) for k, v in row.to_dict().items()}
        records.append(
            (
                source_key,
                source_file,
                sheet_name,
                int(idx) + 2,
                content_hash,
                uploader,
                category,
                val(type_col) or category,
                clean_project_name(val(project_col)),
                val(person_col),
                val(identifier_col),
                issue.isoformat() if pd.notna(issue) else None,
                due.isoformat() if pd.notna(due) else None,
                val(responsible_col),
                val(email_col),
                val(status_col),
                json.dumps(raw, ensure_ascii=False),
                now,
            )
        )

    with db_connect() as conn:
        conn.execute("DELETE FROM records WHERE source_key = ?", (source_key,))
        conn.execute("DELETE FROM sources WHERE source_key = ?", (source_key,))
        conn.executemany(
            """
            INSERT INTO records (
                source_key, source_file, source_sheet, source_row, source_hash,
                uploaded_by, category, document_type, project, person_asset,
                identifier, issue_date, due_date, responsible, email, source_status,
                raw_json, imported_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            records,
        )
        conn.execute(
            """
            INSERT INTO sources (source_key, source_file, source_hash, uploaded_by, sheets, imported_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (source_key, source_file, content_hash, uploader, sheet_name, now),
        )
        conn.commit()
    return len(records)


def calc_status(value, today, thresholds):
    if value is None or pd.isna(value) or str(value).strip() == "":
        return "Fecha pendiente"
    try:
        due = pd.Timestamp(value).date() if not isinstance(value, date) else value
    except Exception:
        return "Fecha pendiente"

    today_date = pd.Timestamp(today).date() if not isinstance(today, date) else today
    days = (due - today_date).days

    if days < 0:
        return "Vencido"
    if days == 0:
        return "Vence hoy"
    if thresholds and days <= max(thresholds):
        if days in thresholds:
            return f"Vence en {days} días"
        return "Próximo a vencer"
    return "Vigente"


def options_for(df, column, include_all=True):
    if df.empty or column not in df.columns:
        return ["Todos"] if include_all else []
    values = sorted({str(v).strip() for v in df[column].dropna().unique() if str(v).strip()})
    return (["Todos"] + values) if include_all else values


def summary_by_project(frame, threshold):
    if frame.empty:
        return pd.DataFrame(columns=["Proyecto", "Total documentos", "Vencidos", "% vencidos", "Vencen hoy", "Próximos a vencer", "% próximos a vencer"])

    data = frame.copy()
    days = pd.to_numeric(data["Dias_restantes"], errors="coerce")
    data["_vencido"] = days < 0
    data["_hoy"] = days == 0
    data["_proximo"] = (days > 0) & (days <= threshold)

    summary = data.groupby("Proyecto", dropna=False).agg(
        Total_documentos=("Proyecto", "size"),
        Vencidos=("_vencido", "sum"),
        Vencen_hoy=("_hoy", "sum"),
        Proximos_a_vencer=("_proximo", "sum"),
    ).reset_index()

    summary.rename(columns={"Total_documentos": "Total documentos", "Vencen_hoy": "Vencen hoy", "Proximos_a_vencer": "Próximos a vencer"}, inplace=True)
    denom = summary["Total documentos"].replace(0, pd.NA)
    summary["% vencidos"] = (summary["Vencidos"] / denom * 100).fillna(0).round(1)
    summary["% próximos a vencer"] = (summary["Próximos a vencer"] / denom * 100).fillna(0).round(1)
    return summary.sort_values(["Vencidos", "Próximos a vencer", "Total documentos"], ascending=False).reset_index(drop=True)


def export_filtered_excel(detail, summary):
    output = BytesIO()
    detail_cols = [
        "Proyecto", "Categoria", "Tipo_control", "Persona_activo", "Identificador",
        "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento", "Responsable", "Correo"
    ]
    detail_export = detail[[c for c in detail_cols if c in detail.columns]].copy()
    detail_export = detail_export.rename(columns={
        "Categoria": "Categoría",
        "Tipo_control": "Tipo de control",
        "Persona_activo": "Persona / activo",
        "Identificador": "Código / placa",
        "Fecha_vencimiento": "Fecha de vencimiento",
        "Dias_restantes": "Días restantes",
        "Estado_seguimiento": "Estado",
        "Responsable": "Responsable",
        "Correo": "Correo",
    })

    summary_export = summary.copy()
    with pd.ExcelWriter(output, engine="openpyxl", date_format="DD/MM/YYYY", datetime_format="DD/MM/YYYY") as writer:
        summary_export.to_excel(writer, index=False, sheet_name="Resumen por proyecto")
        detail_export.to_excel(writer, index=False, sheet_name="Base filtrada")
    return output.getvalue()


all_records = normalize_records(load_records())
all_sources = load_sources()

with st.sidebar:
    st.header("Configuración")
    today = st.date_input("Fecha de evaluación", value=date.today())
    warning_days = st.multiselect(
        "Anticipación de alertas (días)",
        options=ALERT_OPTIONS,
        default=ALERT_OPTIONS,
        help="Selecciona la alerta: 1, 7 y/o 15 días."
    )
    st.caption("La lógica de vencimiento usa la misma clasificación en todos los indicadores y reportes.")

    st.divider()
    st.caption("La base de datos SQLite está en la carpeta data. En Streamlit Cloud, usa almacenamiento persistente para que la información no se pierda entre reinicios.")


tab_panel, tab_reports = st.tabs(["Panel general", "Reportes"])

with tab_panel:
    st.subheader("Resumen general")

    if all_records.empty:
        st.info("No hay registros disponibles en la base de datos. La aplicación está operativa, pero aún no existen documentos evaluados.")
    else:
        threshold = max(warning_days or [15])

        project_filter = st.selectbox("Proyecto", options_for(all_records, "Proyecto"), index=0)
        category_filter = st.selectbox("Categoría / tipo de documento", options_for(all_records, "Categoria"), index=0)

        filtered = all_records.copy()
        if project_filter != "Todos":
            filtered = filtered[filtered["Proyecto"].astype(str) == project_filter]
        if category_filter != "Todos":
            filtered = filtered[filtered["Categoria"].astype(str) == category_filter]

        if filtered.empty:
            st.warning("No existen registros para los filtros seleccionados.")
        else:
            total_docs = len(filtered)
            vencidos = int((filtered["Estado_seguimiento"] == "Vencido").sum())
            proximos = int(filtered["Estado_seguimiento"].str.contains("Próximo a vencer|Vence en|Vence hoy", regex=True, na=False).sum())
            vigentes = int((filtered["Estado_seguimiento"] == "Vigente").sum())

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total de registros evaluados", total_docs)
            col2.metric("Documentos vencidos", vencidos)
            col3.metric("Documentos próximos a vencer", proximos)
            col4.metric("Documentos vigentes", vigentes)

            status_summary = filtered["Estado_seguimiento"].value_counts().reset_index()
            status_summary.columns = ["Estado", "Cantidad"]
            status_summary = status_summary.sort_values("Cantidad", ascending=False)

            fig_status = px.bar(
                status_summary,
                x="Estado",
                y="Cantidad",
                color="Estado",
                text="Cantidad",
                title="Cuantificación por estado"
            )
            fig_status.update_layout(xaxis_title="Estado", yaxis_title="Cantidad", height=360)
            fig_status.update_traces(textposition="outside")
            st.plotly_chart(fig_status, use_container_width=True)

            summary = summary_by_project(filtered, threshold)
            if not summary.empty:
                fig_project = px.bar(
                    summary.melt(id_vars="Proyecto", value_vars=["Vencidos", "Próximos a vencer"], var_name="Estado", value_name="Cantidad"),
                    x="Cantidad",
                    y="Proyecto",
                    color="Estado",
                    barmode="group",
                    orientation="h",
                    text="Cantidad",
                    title="Vencidos y próximos a vencer por proyecto"
                )
                fig_project.update_layout(height=max(350, 45 * len(summary) + 120), margin=dict(l=10, r=20, t=60, b=10))
                fig_project.update_traces(textposition="outside")
                st.plotly_chart(fig_project, use_container_width=True)

                st.markdown("#### Tabla resumen por proyecto")
                summary_display = summary.copy()
                summary_display["% vencidos"] = summary_display["% vencidos"].map(lambda x: f"{x:.1f}%")
                summary_display["% próximos a vencer"] = summary_display["% próximos a vencer"].map(lambda x: f"{x:.1f}%")
                st.dataframe(summary_display, use_container_width=True, hide_index=True)

            st.markdown("#### Documentos por revisar / atención")
            attention = filtered[["Proyecto", "Categoria", "Tipo_control", "Persona_activo", "Identificador", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento"]].copy()
            attention["Fecha_vencimiento"] = pd.to_datetime(attention["Fecha_vencimiento"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(attention, use_container_width=True, hide_index=True)

with tab_reports:
    st.subheader("Reportes")

    if all_records.empty:
        st.info("No hay registros para generar reportes.")
    else:
        project_filter = st.selectbox("Proyecto", options_for(all_records, "Proyecto"), index=0, key="report_project")
        category_filter = st.selectbox("Categoría / tipo de documento", options_for(all_records, "Categoria"), index=0, key="report_category")

        report = all_records.copy()
        if project_filter != "Todos":
            report = report[report["Proyecto"].astype(str) == project_filter]
        if category_filter != "Todos":
            report = report[report["Categoria"].astype(str) == category_filter]

        if report.empty:
            st.warning("No existen datos con los filtros seleccionados.")
        else:
            columns_to_show = [
                "Proyecto", "Categoria", "Tipo_control", "Persona_activo", "Identificador",
                "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento", "Responsable", "Correo"
            ]
            output = report[columns_to_show].copy()
            output["Fecha_vencimiento"] = pd.to_datetime(output["Fecha_vencimiento"], errors="coerce").dt.strftime("%d/%m/%Y")
            st.dataframe(output, use_container_width=True, hide_index=True)

            csv_data = output.to_csv(index=False, encoding="utf-8-sig")
            excel_data = export_filtered_excel(report, summary_by_project(report, max(warning_days or [15])))

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Descargar CSV",
                    data=csv_data,
                    file_name="reporte_vencimientos.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with col2:
                st.download_button(
                    label="Descargar Excel",
                    data=excel_data,
                    file_name="reporte_vencimientos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

st.caption("Sistema administrativo para seguimiento de documentos y vencimientos. Los datos se muestran y reportan según la fecha de evaluación y los filtros seleccionados.")
