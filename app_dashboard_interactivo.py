import io
import streamlit as st
import pandas as pd
import numpy as np
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- Configuración visual de la aplicación ---
st.set_page_config(
    page_title="Dashboard Comercial | GlobalTech",
    page_icon="📊",
    layout="wide"
)

# --- Función para limpiar y transformar texto monetario a números ---
def clean_currency_series(series):
    """Limpia textos con $, espacios, comas y puntos a valores float."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors='coerce').fillna(0.0)
    
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace('$', '', regex=False)
        .str.replace(' ', '', regex=False)
        .str.replace('\xa0', '', regex=False)
    )
    
    def parse_val(v):
        if pd.isna(v) or v in ('nan', 'None', '', 'null'):
            return 0.0
        # Caso formato con coma y punto (ej. 1,250.00 o 1.250,00)
        if '.' in v and ',' in v:
            if v.rfind('.') > v.rfind(','):
                v = v.replace(',', '')
            else:
                v = v.replace('.', '').replace(',', '.')
        # Caso decimal con coma (ej. 6,50)
        elif ',' in v:
            parts = v.split(',')
            if len(parts) == 2 and len(parts[1]) in (1, 2):
                v = v.replace(',', '.')
            else:
                v = v.replace(',', '')
        try:
            return float(v)
        except Exception:
            return 0.0
            
    return cleaned.apply(parse_val)

# --- 1. Autenticación con Google Drive ---
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

@st.cache_resource
def get_drive_service():
    """Inicializa el cliente autenticado de Google Drive."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

# --- 2. Descarga y procesamiento de datos ---
@st.cache_data(ttl=60)
def load_data_from_drive(file_id):
    """Descarga el Excel en memoria y procesa los números y fórmulas."""
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    df = pd.read_excel(fh, sheet_name=0, engine='openpyxl')
    
    # Estandarizar nombres de columnas a minúsculas sin espacios
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    # Parsear fechas
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        
    # Limpiar columnas numéricas
    if 'cantidad' in df.columns:
        df['cantidad'] = clean_currency_series(df['cantidad'])
    else:
        df['cantidad'] = 0.0
        
    if 'precio_unitario' in df.columns:
        df['precio_unitario'] = clean_currency_series(df['precio_unitario'])
    else:
        df['precio_unitario'] = 0.0
        
    if 'total_venta' in df.columns:
        df['total_venta'] = clean_currency_series(df['total_venta'])
    else:
        df['total_venta'] = 0.0
        
    # Resolver fórmulas de Excel: Si total_venta quedó en 0 o NaN, multiplicamos cantidad * precio
    mask_calc = (df['total_venta'] == 0) | df['total_venta'].isna()
    df.loc[mask_calc, 'total_venta'] = df.loc[mask_calc, 'cantidad'] * df.loc[mask_calc, 'precio_unitario']
    
    return df

# --- 3. Cabecera y botón de refresco ---
col_title, col_btn = st.columns([5, 1])
with col_title:
    st.title("📈 Panel de Ventas en Vivo")
    st.caption("Conectado en tiempo real a Google Drive (base_datos_ventas.xlsx)")

with col_btn:
    st.write("")
    if st.button("🔄 Refrescar", help="Descarga los datos más recientes del archivo de Drive"):
        st.cache_data.clear()
        st.rerun()

# --- 4. Obtener ID y descargar datos ---
FILE_ID = st.secrets["drive_settings"]["file_id"]

try:
    with st.spinner("Conectando con Google Drive..."):
        df = load_data_from_drive(FILE_ID)
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

# --- 5. Filtros en la barra lateral ---
st.sidebar.header("🔍 Filtros de Consulta")

ciudades = ["Todas"] + sorted(df["ciudad"].dropna().unique().tolist()) if "ciudad" in df.columns else ["Todas"]
ciudad_sel = st.sidebar.selectbox("Ciudad:", ciudades)

categorias = ["Todas"] + sorted(df["categoria"].dropna().unique().tolist()) if "categoria" in df.columns else ["Todas"]
cat_sel = st.sidebar.selectbox("Categoría:", categorias)

estados = ["Todos"] + sorted(df["estado"].dropna().unique().tolist()) if "estado" in df.columns else ["Todos"]
estado_sel = st.sidebar.selectbox("Estado de orden:", estados)

# Aplicar filtros
df_f = df.copy()
if ciudad_sel != "Todas" and "ciudad" in df_f.columns:
    df_f = df_f[df_f["ciudad"] == ciudad_sel]
if cat_sel != "Todas" and "categoria" in df_f.columns:
    df_f = df_f[df_f["categoria"] == cat_sel]
if estado_sel != "Todos" and "estado" in df_f.columns:
    df_f = df_f[df_f["estado"] == estado_sel]

# --- 6. Indicadores Clave (KPIs) ---
k1, k2, k3, k4 = st.columns(4)

total_ventas = float(df_f["total_venta"].sum())
total_unidades = int(df_f["cantidad"].sum())
num_ordenes = len(df_f)
ticket_medio = (total_ventas / num_ordenes) if num_ordenes > 0 else 0.0

k1.metric("Ingresos Totales", f"${total_ventas:,.2f}")
k2.metric("Unidades Vendidas", f"{total_unidades:,}")
k3.metric("Ticket Promedio", f"${ticket_medio:,.2f}")
k4.metric("Nº de Órdenes", num_ordenes)

st.markdown("---")

# --- 7. Gráficos ---
c_g1, c_g2 = st.columns(2)

with c_g1:
    st.subheader("Ventas por Categoría")
    if "categoria" in df_f.columns:
        ventas_cat = df_f.groupby("categoria")["total_venta"].sum()
        st.bar_chart(ventas_cat)

with c_g2:
    st.subheader("Ventas por Ciudad")
    if "ciudad" in df_f.columns:
        ventas_ciudad = df_f.groupby("ciudad")["total_venta"].sum()
        st.bar_chart(ventas_ciudad)

# --- 8. Detalle tabular de registros ---
with st.expander("📄 Ver registros detallados", expanded=False):
    st.dataframe(df_f, use_container_width=True, hide_index=True)
