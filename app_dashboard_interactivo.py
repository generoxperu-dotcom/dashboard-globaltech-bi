import io
import streamlit as st
import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

# --- Configuración visual ---
st.set_page_config(
    page_title="Dashboard Comercial | GlobalTech",
    page_icon="📊",
    layout="wide"
)

# --- Función de limpieza numérica ---
def clean_val(v):
    if pd.isna(v) or v is None:
        return 0.0
    s = str(v).strip().replace('$', '').replace(' ', '').replace('\xa0', '')
    if not s or s.lower() in ('nan', 'none', 'null'):
        return 0.0
    if '.' in s and ',' in s:
        if s.rfind('.') > s.rfind(','):
            s = s.replace(',', '')
        else:
            s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        parts = s.split(',')
        if len(parts) == 2 and len(parts[1]) in (1, 2):
            s = s.replace(',', '.')
        else:
            s = s.replace(',', '')
    try:
        return float(s)
    except Exception:
        return 0.0

# --- 1. Autenticación con Google Drive (Permiso de lectura y escritura) ---
SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def get_drive_service():
    """Inicializa el cliente de Google Drive."""
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

# --- 2. Descargar datos desde Google Drive ---
@st.cache_data(ttl=60)
def load_data_from_drive(file_id):
    """Descarga el Excel en memoria."""
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    df = pd.read_excel(fh, sheet_name=0, engine='openpyxl')
    df.columns = [str(c).strip().lower() for c in df.columns]
    
    if 'fecha' in df.columns:
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce').dt.strftime('%Y-%m-%d')
        
    if 'cantidad' in df.columns:
        df['cantidad'] = df['cantidad'].apply(clean_val).astype(int)
    else:
        df['cantidad'] = 0
        
    if 'precio_unitario' in df.columns:
        df['precio_unitario'] = df['precio_unitario'].apply(clean_val)
    else:
        df['precio_unitario'] = 0.0
        
    if 'total_venta' in df.columns:
        df['total_venta'] = df['total_venta'].apply(clean_val)
    else:
        df['total_venta'] = 0.0
        
    # Recalcular total si viene en 0
    mask = (df['total_venta'] == 0)
    df.loc[mask, 'total_venta'] = df.loc[mask, 'cantidad'] * df.loc[mask, 'precio_unitario']
    
    return df

# --- 3. Subir y guardar datos en Google Drive ---
def save_data_to_drive(file_id, df_to_save):
    """Sobrescribe el archivo Excel en Google Drive con la versión editada."""
    service = get_drive_service()
    buffer = io.BytesIO()
    
    # Asegurar que total_venta esté recalculado antes de guardar
    df_to_save['total_venta'] = df_to_save['cantidad'] * df_to_save['precio_unitario']
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_to_save.to_excel(writer, sheet_name='Ventas', index=False)
    buffer.seek(0)
    
    media = MediaIoBaseUpload(
        buffer,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    service.files().update(
        fileId=file_id,
        media_body=media,
        supportsAllDrives=True
    ).execute()

# --- 4. Encabezado ---
col_title, col_btn = st.columns([5, 1])
with col_title:
    st.title("📈 Panel de Ventas en Vivo")
    st.caption("Conectado en tiempo real con Google Drive (base_datos_ventas.xlsx)")

with col_btn:
    st.write("")
    if st.button("🔄 Refrescar", help="Descarga los datos más recientes de Drive"):
        st.cache_data.clear()
        st.rerun()

FILE_ID = st.secrets["drive_settings"]["file_id"]

try:
    with st.spinner("Cargando datos desde Google Drive..."):
        df = load_data_from_drive(FILE_ID)
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

# --- 5. Pestañas: Dashboard Visual vs Editor Interactivo ---
tab_dash, tab_edit = st.tabs(["📊 Dashboard y Reportes", "✏️ Editor de Base de Datos"])

with tab_dash:
    # Filtros laterales
    st.sidebar.header("🔍 Filtros de Consulta")

    ciudades = ["Todas"] + sorted([c for c in df["ciudad"].dropna().unique().tolist() if str(c).strip()]) if "ciudad" in df.columns else ["Todas"]
    ciudad_sel = st.sidebar.selectbox("Ciudad:", ciudades)

    categorias = ["Todas"] + sorted([c for c in df["categoria"].dropna().unique().tolist() if str(c).strip()]) if "categoria" in df.columns else ["Todas"]
    cat_sel = st.sidebar.selectbox("Categoría:", categorias)

    estados = ["Todos"] + sorted([c for c in df["estado"].dropna().unique().tolist() if str(c).strip()]) if "estado" in df.columns else ["Todos"]
    estado_sel = st.sidebar.selectbox("Estado de orden:", estados)

    # Filtrar datos
    df_f = df.copy()
    if ciudad_sel != "Todas" and "ciudad" in df_f.columns:
        df_f = df_f[df_f["ciudad"] == ciudad_sel]
    if cat_sel != "Todas" and "categoria" in df_f.columns:
        df_f = df_f[df_f["categoria"] == cat_sel]
    if estado_sel != "Todos" and "estado" in df_f.columns:
        df_f = df_f[df_f["estado"] == estado_sel]

    # KPIs
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

    # Gráficos
    c_g1, c_g2 = st.columns(2)
    with c_g1:
        st.subheader("Ventas por Categoría")
        if "categoria" in df_f.columns and len(df_f) > 0:
            ventas_cat = df_f.groupby("categoria")["total_venta"].sum()
            st.bar_chart(ventas_cat)
        else:
            st.info("Sin datos para mostrar.")

    with c_g2:
        st.subheader("Ventas por Ciudad")
        if "ciudad" in df_f.columns and len(df_f) > 0:
            ventas_ciudad = df_f.groupby("ciudad")["total_venta"].sum()
            st.bar_chart(ventas_ciudad)
        else:
            st.info("Sin datos para mostrar.")

with tab_edit:
    st.subheader("📝 Edición directa en la Hoja de Cálculo")
    st.info("Haz doble clic sobre cualquier celda para editarla. Al terminar, presiona el botón verde para guardar en Google Drive.")
    
    # Tabla editable interactiva
    df_editado = st.data_editor(
        df,
        num_rows="dynamic",  # Permite agregar y borrar filas
        use_container_width=True,
        hide_index=True,
        disabled=["total_venta"]  # total_venta se autocalcula (cantidad * precio)
    )
    
    col_save, _ = st.columns([2, 5])
    with col_save:
        if st.button("💾 Guardar cambios en Google Drive", type="primary"):
            try:
                with st.spinner("Guardando en Google Drive..."):
                    save_data_to_drive(FILE_ID, df_editado)
                st.success("¡Base de datos actualizada exitosamente en Google Drive!")
                st.cache_data.clear()
                st.rerun()
            except Exception as e:
                st.error(f"Error al guardar: {e}")
                st.info("Verifica que el bot tenga rol de 'Editor' en Google Drive.")
