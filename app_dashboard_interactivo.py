import io
import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- Configuración visual de la aplicación ---
st.set_page_config(
    page_title="Dashboard Comercial | GlobalTech",
    page_icon="📊",
    layout="wide"
)

# --- 1. Autenticación con Google Drive ---
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

@st.cache_resource
def get_drive_service():
    """Inicializa y reutiliza el cliente autenticado de Google Drive."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

# --- 2. Descarga del archivo Excel en memoria ---
@st.cache_data(ttl=60)
def load_data_from_drive(file_id):
    """Descarga el Excel directamente de Google Drive con soporte para todas las unidades."""
    service = get_drive_service()
    
    # supportsAllDrives=True resuelve restricciones en unidades corporativas o compartidas
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    # Lee la pestaña 'Ventas' (o la primera hoja por defecto)
    df = pd.read_excel(fh, sheet_name=0, engine='openpyxl')
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'])
    return df

# --- 3. Cabecera y botón de refresco ---
col_title, col_btn = st.columns([5, 1])
with col_title:
    st.title("📈 Panel de Ventas en Vivo")
    st.caption("Conectado en tiempo real a Google Drive (base_datos_ventas.xlsx)")

with col_btn:
    st.write("")
    if st.button("🔄 Refrescar", help="Limpia la memoria caché para consultar la versión más reciente del Excel"):
        st.cache_data.clear()
        st.rerun()

# --- 4. Obtener ID y descargar datos ---
FILE_ID = st.secrets["drive_settings"]["file_id"]

try:
    with st.spinner("Cargando datos desde Google Drive..."):
        df = load_data_from_drive(FILE_ID)
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.info("Asegúrate de que el archivo tenga acceso de 'Lector' para cualquier persona con el enlace o para el bot.")
    st.stop()

# --- 5. Filtros en la barra lateral ---
st.sidebar.header("🔍 Filtros de Consulta")

# Filtro Ciudad
ciudades = ["Todas"] + sorted(df["ciudad"].dropna().unique().tolist()) if "ciudad" in df.columns else ["Todas"]
ciudad_sel = st.sidebar.selectbox("Ciudad:", ciudades)

# Filtro Categoría
categorias = ["Todas"] + sorted(df["categoria"].dropna().unique().tolist()) if "categoria" in df.columns else ["Todas"]
cat_sel = st.sidebar.selectbox("Categoría:", categorias)

# Filtro Estado
estados = ["Todos"] + sorted(df["estado"].dropna().unique().tolist()) if "estado" in df.columns else ["Todos"]
estado_sel = st.sidebar.selectbox("Estado de orden:", estados)

# Aplicación de filtros
df_f = df.copy()
if ciudad_sel != "Todas" and "ciudad" in df_f.columns:
    df_f = df_f[df_f["ciudad"] == ciudad_sel]
if cat_sel != "Todas" and "categoria" in df_f.columns:
    df_f = df_f[df_f["categoria"] == cat_sel]
if estado_sel != "Todos" and "estado" in df_f.columns:
    df_f = df_f[df_f["estado"] == estado_sel]

# --- 6. Indicadores Clave (KPIs) ---
k1, k2, k3, k4 = st.columns(4)

total_ventas = df_f["total_venta"].sum() if "total_venta" in df_f.columns else 0
total_unidades = df_f["cantidad"].sum() if "cantidad" in df_f.columns else 0
ticket_medio = df_f["total_venta"].mean() if ("total_venta" in df_f.columns and len(df_f) > 0) else 0

k1.metric("Ingresos Totales", f"${total_ventas:,.2f}")
k2.metric("Unidades Vendidas", f"{int(total_unidades):,}")
k3.metric("Ticket Promedio", f"${ticket_medio:,.2f}")
k4.metric("Nº de Órdenes", len(df_f))

st.markdown("---")

# --- 7. Gráficos ---
c_g1, c_g2 = st.columns(2)

with c_g1:
    st.subheader("Ventas por Categoría")
    if "categoria" in df_f.columns and "total_venta" in df_f.columns:
        ventas_cat = df_f.groupby("categoria")["total_venta"].sum()
        st.bar_chart(ventas_cat)
    else:
        st.info("Sin datos para agrupar por categoría.")

with c_g2:
    st.subheader("Ventas por Ciudad")
    if "ciudad" in df_f.columns and "total_venta" in df_f.columns:
        ventas_ciudad = df_f.groupby("ciudad")["total_venta"].sum()
        st.bar_chart(ventas_ciudad)
    else:
        st.info("Sin datos para agrupar por ciudad.")

# --- 8. Detalle tabular ---
with st.expander("📄 Ver registros detallados", expanded=False):
    if "fecha" in df_f.columns:
        st.dataframe(df_f.sort_values(by="fecha", ascending=False), use_container_width=True, hide_index=True)
    else:
        st.dataframe(df_f, use_container_width=True, hide_index=True)
