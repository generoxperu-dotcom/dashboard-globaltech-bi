import io
import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="Dashboard Ventas | Google Drive", page_icon="📊", layout="wide")

# Autenticación con Google Drive usando los secretos
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

@st.cache_resource
def get_drive_service():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

@st.cache_data(ttl=60)
def load_data_from_drive(file_id):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    df = pd.read_excel(fh, sheet_name="Ventas", engine='openpyxl')
    df['fecha'] = pd.to_datetime(df['fecha'])
    return df

# Cabecera y botón de refresco
col_title, col_btn = st.columns([5, 1])
with col_title:
    st.title("📈 Panel de Ventas en Vivo")
    st.caption("Conectado directamente al Excel de Google Drive")

with col_btn:
    st.write("")
    if st.button("🔄 Refrescar"):
        st.cache_data.clear()
        st.rerun()

FILE_ID = st.secrets["drive_settings"]["file_id"]

try:
    with st.spinner("Cargando base de datos..."):
        df = load_data_from_drive(FILE_ID)
except Exception as e:
    st.error(f"Error al conectar con Drive: {e}")
    st.stop()

# Filtros laterales
st.sidebar.header("Filtros")
ciudades = ["Todas"] + sorted(df["ciudad"].unique().tolist())
ciudad_sel = st.sidebar.selectbox("Ciudad:", ciudades)

categorias = ["Todas"] + sorted(df["categoria"].unique().tolist())
cat_sel = st.sidebar.selectbox("Categoría:", categorias)

df_filtrado = df.copy()
if ciudad_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["ciudad"] == ciudad_sel]
if cat_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado["categoria"] == cat_sel]

# Indicadores clave (KPIs)
k1, k2, k3 = st.columns(3)
k1.metric("Ingresos Totales", f"${df_filtrado['total_venta'].sum():,.2f}")
k2.metric("Unidades Vendidas", f"{df_filtrado['cantidad'].sum():,}")
k3.metric("Nº Transacciones", len(df_filtrado))

st.markdown("---")

# Visualizaciones
c_g1, c_g2 = st.columns(2)
with c_g1:
    st.subheader("Ventas por Categoría")
    st.bar_chart(df_filtrado.groupby("categoria")["total_venta"].sum())

with c_g2:
    st.subheader("Ventas por Ciudad")
    st.bar_chart(df_filtrado.groupby("ciudad")["total_venta"].sum())

with st.expander("Ver tabla completa de registros"):
    st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
