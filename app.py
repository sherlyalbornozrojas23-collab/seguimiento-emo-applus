import streamlit as st
import pandas as pd
from datetime import date, datetime
from io import BytesIO

st.set_page_config(page_title="Control EMO | Applus+", page_icon="🩺", layout="wide")

st.title("🩺 Sistema de Seguimiento de EMO")
st.caption("Prototipo administrativo para optimizar el seguimiento de exámenes médicos ocupacionales. No determina aptitud médica.")

st.info(
    "Privacidad: utiliza datos ficticios para probar esta aplicación. No cargues diagnósticos, historias clínicas "
    "ni resultados médicos detallados. Antes de usar datos reales, valida con TI, SST y el responsable de privacidad "
    "los permisos, el almacenamiento y las políticas corporativas de Applus+."
)

REQUIRED = ["Codigo", "Proyecto", "Puesto", "Fecha_vencimiento", "Responsable", "Correo_responsable", "Estado_aptitud"]
TEMPLATE = pd.DataFrame([
    {"Codigo":"T001","Proyecto":"ILLA","Puesto":"Ejemplo","Fecha_vencimiento":"2026-11-15","Responsable":"Administración","Correo_responsable":"responsable@empresa.com","Estado_aptitud":"Validar con SST"},
    {"Codigo":"T002","Proyecto":"ATIPAQ","Puesto":"Ejemplo","Fecha_vencimiento":"2026-11-30","Responsable":"RR. HH.","Correo_responsable":"rrhh@empresa.com","Estado_aptitud":"Validar con SST"},
    {"Codigo":"T003","Proyecto":"Bateas","Puesto":"Ejemplo","Fecha_vencimiento":"2026-10-20","Responsable":"SST","Correo_responsable":"sst@empresa.com","Estado_aptitud":"Validar con SST"},
])

with st.sidebar:
    st.header("Configuración")
    today = st.date_input("Fecha de evaluación", value=date.today())
    warning_days = st.multiselect("Anticipación de alertas (días)", [7, 15, 30, 45, 60], default=[7, 15, 30])
    st.caption("Las alertas se muestran en la app. El envío de correos no está habilitado en este prototipo.")

st.subheader("1. Cargar la base de seguimiento")
uploaded = st.file_uploader("Carga un archivo Excel (.xlsx) o CSV con el registro de EMO", type=["xlsx", "csv"])

if uploaded:
    try:
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded, dtype={"Codigo": str})
        else:
            df = pd.read_excel(uploaded, dtype={"Codigo": str})
    except Exception as e:
        st.error(f"No se pudo leer el archivo. Revisa el formato. Detalle: {e}")
        st.stop()
else:
    st.warning("Estás viendo datos de demostración. Descarga la plantilla, reemplázala con datos autorizados y vuelve a cargarla.")
    df = TEMPLATE.copy()

missing = [c for c in REQUIRED if c not in df.columns]
if missing:
    st.error("Faltan estas columnas obligatorias: " + ", ".join(missing))
    st.markdown("Columnas requeridas: `"+ "`, `".join(REQUIRED) + "`.")
    st.stop()

df = df.copy()
df["Fecha_vencimiento"] = pd.to_datetime(df["Fecha_vencimiento"], errors="coerce").dt.date
df["Dias_restantes"] = df["Fecha_vencimiento"].apply(lambda d: (d - today).days if pd.notna(d) else None)

def status(days):
    if pd.isna(days):
        return "Fecha por revisar"
    if days < 0:
        return "Vencido: requiere revisión"
    if days == 0:
        return "Vence hoy"
    if days in warning_days:
        return f"Alerta: vence en {days} días"
    if any(days <= threshold for threshold in warning_days):
        return "Próximo a vencer"
    return "Vigente según fecha"

df["Estado_seguimiento"] = df["Dias_restantes"].apply(status)

total = len(df)
expired = int((df["Dias_restantes"].apply(lambda x: pd.notna(x) and x < 0)).sum())
today_count = int((df["Dias_restantes"] == 0).sum())
soon = int((df["Dias_restantes"].apply(lambda x: pd.notna(x) and 0 < x <= max(warning_days or [30]))).sum())
unknown = int(df["Fecha_vencimiento"].isna().sum())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Registros", total)
c2.metric("Vencidos", expired)
c3.metric("Vencen hoy", today_count)
c4.metric("Próximos a vencer", soon)
c5.metric("Fecha por revisar", unknown)

st.subheader("2. Buscar y filtrar")
f1, f2, f3 = st.columns(3)
projects = ["Todos"] + sorted([str(x) for x in df["Proyecto"].dropna().unique()])
selected_project = f1.selectbox("Proyecto", projects)
statuses = ["Todos"] + sorted(df["Estado_seguimiento"].dropna().unique().tolist())
selected_status = f2.selectbox("Estado de seguimiento", statuses)
search = f3.text_input("Buscar por código o responsable").strip().lower()

view = df.copy()
if selected_project != "Todos":
    view = view[view["Proyecto"].astype(str) == selected_project]
if selected_status != "Todos":
    view = view[view["Estado_seguimiento"] == selected_status]
if search:
    view = view[
        view["Codigo"].astype(str).str.lower().str.contains(search, na=False)
        | view["Responsable"].astype(str).str.lower().str.contains(search, na=False)
    ]

display_cols = REQUIRED + ["Dias_restantes", "Estado_seguimiento"]
st.dataframe(view[display_cols], use_container_width=True, hide_index=True)

st.subheader("3. Lista de acciones prioritarias")
priority = df[
    df["Dias_restantes"].apply(
        lambda x: pd.notna(x) and (x < 0 or x == 0 or (x > 0 and x <= max(warning_days or [30])))
    )
].sort_values("Dias_restantes", na_position="last")

if priority.empty:
    st.success("No hay vencimientos prioritarios según la configuración actual.")
else:
    st.dataframe(priority[["Codigo", "Proyecto", "Responsable", "Correo_responsable", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento"]], use_container_width=True, hide_index=True)
    st.caption("Acción sugerida: verificar el estado con el área autorizada y coordinar renovación cuando corresponda. La app no valida aptitud.")

st.subheader("4. Exportar reporte")
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
    df[display_cols].to_excel(writer, index=False, sheet_name="Seguimiento_EMO")
    priority[["Codigo", "Proyecto", "Responsable", "Correo_responsable", "Fecha_vencimiento", "Dias_restantes", "Estado_seguimiento"]].to_excel(writer, index=False, sheet_name="Acciones_prioritarias")
st.download_button(
    "⬇️ Descargar reporte Excel",
    data=out.getvalue(),
    file_name=f"reporte_seguimiento_emo_{today.isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("5. Formato de plantilla")
template_bytes = BytesIO()
with pd.ExcelWriter(template_bytes, engine="openpyxl") as writer:
    TEMPLATE.to_excel(writer, index=False, sheet_name="Base_EMO")
st.download_button(
    "⬇️ Descargar plantilla Excel",
    data=template_bytes.getvalue(),
    file_name="plantilla_seguimiento_emo.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.caption("Versión piloto: los estados son indicadores administrativos basados en fechas. Deben confirmarse con SST/medicina ocupacional y las políticas de la empresa.")
