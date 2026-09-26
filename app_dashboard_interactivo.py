import io
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import plotly.express as px
import plotly.graph_objects as go

# --- 1. CONFIGURACIÓN VISUAL GENERAL ---
st.set_page_config(
    page_title="Management Dashboard | GlobalTech",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. INYECCIÓN DE ESTILO CSS (Management Dashboard - Yellow & Clean UI) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    
    .stApp {
        background-color: #F5F7FA;
    }

    .block-container {
        padding-top: 1.8rem;
        padding-bottom: 2.5rem;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Banner Superior */
    .dashboard-header {
        background: #FFFFFF;
        padding: 22px 28px;
        border-radius: 20px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 24px;
        border: 1px solid #ECEFF2;
        box-shadow: 0px 4px 20px rgba(0, 0, 0, 0.03);
    }
    .dashboard-title-box h1 {
        font-size: 26px;
        font-weight: 800;
        color: #1A1D20 !important;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .dashboard-title-box span {
        color: #FFB800 !important;
    }
    .dashboard-subtitle {
        color: #7A828A !important;
        font-size: 13px;
        font-weight: 500;
        margin-top: 4px;
    }

    /* Tarjetas KPI */
    .kpi-card {
        background: #FFFFFF;
        padding: 18px 20px;
        border-radius: 20px;
        border: 1px solid #ECEFF2;
        box-shadow: 0px 4px 18px rgba(0, 0, 0, 0.03);
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 15px;
    }
    .kpi-info {
        display: flex;
        flex-direction: column;
    }
    .kpi-label {
        font-size: 12px;
        font-weight: 600;
        color: #8C94A0 !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 24px;
        font-weight: 800;
        color: #1A1D20 !important;
        margin-top: 4px;
    }
    .kpi-badge {
        padding: 8px 16px;
        border-radius: 30px;
        font-size: 14px;
        font-weight: 700;
    }
    .badge-yellow { background: #FFB800; color: #FFFFFF; }
    .badge-dark { background: #1E2229; color: #FFFFFF; }
    .badge-orange { background: #FF7A00; color: #FFFFFF; }
    .badge-purple { background: #6C5CE7; color: #FFFFFF; }

    /* Tarjetas de Gráficos */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF !important;
        border-radius: 22px !important;
        border: 1px solid #ECEFF2 !important;
        box-shadow: 0px 4px 18px rgba(0, 0, 0, 0.03) !important;
        padding: 12px !important;
    }

    button[data-baseweb="tab"] {
        font-weight: 700;
        color: #7A828A;
    }
    button[aria-selected="true"] {
        color: #1A1D20 !important;
        border-bottom-color: #FFB800 !important;
    }

    div.stButton > button:first-child {
        background: #FFB800;
        color: #FFFFFF;
        font-weight: 700;
        border: none;
        border-radius: 14px;
        padding: 10px 24px;
        transition: all 0.2s ease;
    }
    div.stButton > button:first-child:hover {
        background: #E5A600;
        color: #FFFFFF;
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(255, 184, 0, 0.35);
    }
</style>
""", unsafe_allow_html=True)

PALETA_COLORES = ["#FFB800", "#1E2229", "#FF7A00", "#6C5CE7", "#00B894", "#E17055"]

# --- 3. FUNCIONES DE LIMPIEZA Y AUTENTICACIÓN ---
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

SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def get_drive_service():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds)

@st.cache_data(ttl=60)
def load_data_from_drive(file_id):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    try:
        df = pd.read_excel(fh, sheet_name="Ventas", engine='openpyxl')
    except Exception:
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
        
    mask = (df['total_venta'] == 0)
    df.loc[mask, 'total_venta'] = df.loc[mask, 'cantidad'] * df.loc[mask, 'precio_unitario']
    
    return df

def save_data_to_drive(file_id, df_to_save):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.seek(0)
    
    try:
        wb = openpyxl.load_workbook(fh)
    except Exception:
        wb = openpyxl.Workbook()

    if "Ventas" in wb.sheetnames:
        ws = wb["Ventas"]
        ws.delete_rows(1, ws.max_row + 10)
    else:
        ws = wb.active
        ws.title = "Ventas"

    cols = ["id_transaccion", "fecha", "cliente", "ciudad", "categoria", "producto", "cantidad", "precio_unitario", "total_venta", "estado"]
    ws.append(cols)

    header_fill = PatternFill(start_color="1E2229", end_color="1E2229", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFB800")
    header_alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

    thin_border = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    for row_idx, (_, row) in enumerate(df_to_save.iterrows(), start=2):
        cant = int(clean_val(row.get('cantidad', 0)))
        precio = float(clean_val(row.get('precio_unitario', 0.0)))
        formula_total = f"=G{row_idx}*H{row_idx}"
        
        row_values = [
            str(row.get('id_transaccion', f"TRX-{row_idx-1:04d}")),
            str(row.get('fecha', '')),
            str(row.get('cliente', '')),
            str(row.get('ciudad', '')),
            str(row.get('categoria', '')),
            str(row.get('producto', '')),
            cant,
            precio,
            formula_total,
            str(row.get('estado', 'Completado'))
        ]
        ws.append(row_values)

        ws.cell(row=row_idx, column=7).number_format = '#,##0'
        ws.cell(row=row_idx, column=8).number_format = '$#,##0.00'
        ws.cell(row=row_idx, column=9).number_format = '$#,##0.00'

        for c_i in range(1, len(cols) + 1):
            c_cell = ws.cell(row=row_idx, column=c_i)
            c_cell.border = thin_border
            if c_i in (1, 2, 10):
                c_cell.alignment = Alignment(horizontal="center")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max(len(str(c.value or '')) for c in col)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)

    media = MediaIoBaseUpload(
        out_buf,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )
    service.files().update(
        fileId=file_id,
        media_body=media,
        supportsAllDrives=True
    ).execute()

# --- 4. CARGA DE DATOS ---
FILE_ID = st.secrets["drive_settings"]["file_id"]

try:
    with st.spinner("Conectando con Google Drive..."):
        df = load_data_from_drive(FILE_ID)
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

# --- 5. ENCABEZADO SUPERIOR TIPO BANNER ---
st.markdown("""
<div class="dashboard-header">
    <div class="dashboard-title-box">
        <h1>MANAGEMENT <span>Dashboard.</span></h1>
        <div class="dashboard-subtitle">GlobalTech BI • Executive Performance & Commercial Intelligence</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Pestañas principales
tab_dash, tab_edit = st.tabs(["📊 Dashboard Ejecutivo", "✏️ Editor de Datos"])

with tab_dash:
    # --- FILTROS EN SIDEBAR ---
    st.sidebar.markdown("<h3 style='color:#FFFFFF; font-weight:800;'>Filtros de Control</h3>", unsafe_allow_html=True)
    
    ciudades = ["Todas"] + sorted([c for c in df["ciudad"].dropna().unique().tolist() if str(c).strip()]) if "ciudad" in df.columns else ["Todas"]
    ciudad_sel = st.sidebar.selectbox("Ciudad:", ciudades)

    categorias = ["Todas"] + sorted([c for c in df["categoria"].dropna().unique().tolist() if str(c).strip()]) if "categoria" in df.columns else ["Todas"]
    cat_sel = st.sidebar.selectbox("Categoría:", categorias)

    estados = ["Todos"] + sorted([c for c in df["estado"].dropna().unique().tolist() if str(c).strip()]) if "estado" in df.columns else ["Todos"]
    estado_sel = st.sidebar.selectbox("Estado de Orden:", estados)

    st.sidebar.markdown("---")
    if st.sidebar.button("🔄 Actualizar Datos", help="Recarga el libro en memoria desde Google Drive"):
        st.cache_data.clear()
        st.rerun()

    # Filtrar
    df_f = df.copy()
    if ciudad_sel != "Todas" and "ciudad" in df_f.columns:
        df_f = df_f[df_f["ciudad"] == ciudad_sel]
    if cat_sel != "Todas" and "categoria" in df_f.columns:
        df_f = df_f[df_f["categoria"] == cat_sel]
    if estado_sel != "Todos" and "estado" in df_f.columns:
        df_f = df_f[df_f["estado"] == estado_sel]

    # Cálculos KPIs
    total_ventas = float(df_f["total_venta"].sum())
    total_unidades = int(df_f["cantidad"].sum())
    num_ordenes = len(df_f)
    ticket_medio = (total_ventas / num_ordenes) if num_ordenes > 0 else 0.0

    # Tarjetas KPI
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-info">
                <span class="kpi-label">Ingresos Totales</span>
                <span class="kpi-value">${total_ventas:,.2f}</span>
            </div>
            <div class="kpi-badge badge-yellow">Ventas</div>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-info">
                <span class="kpi-label">Unidades Totales</span>
                <span class="kpi-value">{total_unidades:,}</span>
            </div>
            <div class="kpi-badge badge-dark">Items</div>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-info">
                <span class="kpi-label">Ticket Promedio</span>
                <span class="kpi-value">${ticket_medio:,.2f}</span>
            </div>
            <div class="kpi-badge badge-orange">Promedio</div>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-info">
                <span class="kpi-label">Total Órdenes</span>
                <span class="kpi-value">{num_ordenes:,}</span>
            </div>
            <div class="kpi-badge badge-purple">Órdenes</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")

    # ========================================================
    # --- FUNCIÓN DE ESTILO (TEXTOS Y EJES 100% VISIBLES) ---
    # ========================================================
    def estilizar_figura(fig, titulo=""):
        fig.update_layout(
            template="plotly_white",
            title=dict(
                text=f"<b>{titulo}</b>",
                font=dict(size=14, color="#1A1D20", family="Plus Jakarta Sans"),
                x=0.02,
                y=0.95
            ),
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            margin=dict(l=25, r=35, t=55, b=35),
            font=dict(family="Plus Jakarta Sans", color="#1A1D20", size=12),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=11, color="#1A1D20")
            )
        )
        # Forzar color visible oscuro en ejes
        fig.update_xaxes(
            tickfont=dict(color="#1A1D20", size=11, family="Plus Jakarta Sans"),
            title_font=dict(color="#1A1D20", size=12, family="Plus Jakarta Sans"),
            gridcolor="#F0F2F5",
            zeroline=False,
            showline=True,
            linecolor="#E2E8F0"
        )
        fig.update_yaxes(
            tickfont=dict(color="#1A1D20", size=11, family="Plus Jakarta Sans"),
            title_font=dict(color="#1A1D20", size=12, family="Plus Jakarta Sans"),
            gridcolor="#F0F2F5",
            zeroline=False,
            showline=True,
            linecolor="#E2E8F0"
        )
        return fig

    # ========================================================
    # --- MATRIZ 3X2 DE GRÁFICOS PERSONALIZADOS ---
    # ========================================================

    # --- FILA 1 ---
    r1_c1, r1_c2 = st.columns(2)

    with r1_c1:
        with st.container(border=True):
            if len(df_f) > 0:
                # 1. DISPERSIÓN MULTIDIMENSIONAL
                fig_disp = px.scatter(
                    df_f,
                    x="cantidad",
                    y="precio_unitario",
                    size="total_venta",
                    color="categoria" if "categoria" in df_f.columns else None,
                    hover_name="producto" if "producto" in df_f.columns else None,
                    hover_data=["cliente", "total_venta"] if "cliente" in df_f.columns else ["total_venta"],
                    color_discrete_sequence=PALETA_COLORES,
                    size_max=32,
                    labels={
                        "cantidad": "Unidades por Orden",
                        "precio_unitario": "Precio Unitario ($)",
                        "categoria": "Categoría",
                        "total_venta": "Monto Total ($)"
                    }
                )
                fig_disp = estilizar_figura(fig_disp, "1. Dispersión Multidimensional (Precio vs Volumen)")
                st.plotly_chart(fig_disp, width="stretch", theme=None)
            else:
                st.info("Sin registros.")

    with r1_c2:
        with st.container(border=True):
            # 2. FACTURACIÓN POR CATEGORÍA CON NÚMEROS VISIBLES
            if "categoria" in df_f.columns and len(df_f) > 0:
                v_cat = df_f.groupby("categoria", as_index=False)["total_venta"].sum().sort_values("total_venta", ascending=True)
                fig_cat = px.bar(
                    v_cat,
                    x="total_venta",
                    y="categoria",
                    orientation="h",
                    text="total_venta",
                    color_discrete_sequence=["#FFB800"],
                    labels={"total_venta": "Total Facturado ($)", "categoria": "Categoría"}
                )
                fig_cat = estilizar_figura(fig_cat, "2. Facturación por Categoría de Producto")
                # Etiquetas numéricas visibles sobre cada barra
                fig_cat.update_traces(
                    texttemplate='$%{text:,.2f}',
                    textposition='outside',
                    textfont=dict(color="#1A1D20", size=12, family="Plus Jakarta Sans")
                )
                # Margen extra a la derecha para que el número no se corte
                max_val = v_cat["total_venta"].max()
                fig_cat.update_xaxes(range=[0, max_val * 1.25])
                st.plotly_chart(fig_cat, width="stretch", theme=None)
            else:
                st.info("Sin registros.")

    # --- FILA 2 ---
    r2_c1, r2_c2 = st.columns(2)

    with r2_c1:
        with st.container(border=True):
            # 3. TENDENCIA TEMPORAL CON VALORES EN LOS PUNTOS
            if "fecha" in df_f.columns and len(df_f) > 0:
                v_tiempo = df_f.groupby("fecha", as_index=False)["total_venta"].sum().sort_values("fecha")
                fig_line = px.line(
                    v_tiempo,
                    x="fecha",
                    y="total_venta",
                    text="total_venta",
                    color_discrete_sequence=["#FF7A00"],
                    markers=True,
                    labels={"total_venta": "Ingresos ($)", "fecha": "Fecha"}
                )
                fig_line = estilizar_figura(fig_line, "3. Tendencia Histórica de Ventas")
                fig_line.update_traces(
                    line=dict(width=3, shape="spline"),
                    marker=dict(size=7, color="#1E2229"),
                    texttemplate='$%{text:,.0f}',
                    textposition='top center',
                    textfont=dict(color="#1A1D20", size=10, family="Plus Jakarta Sans")
                )
                st.plotly_chart(fig_line, width="stretch", theme=None)
            else:
                st.info("Sin registros de fecha.")

    with r2_c2:
        with st.container(border=True):
            # 4. DISTRIBUCIÓN POR ESTADO CON PORCENTAJES VISIBLES
            if "estado" in df_f.columns and len(df_f) > 0:
                v_est = df_f.groupby("estado", as_index=False)["total_venta"].sum()
                fig_donut = px.pie(
                    v_est,
                    values="total_venta",
                    names="estado",
                    hole=0.62,
                    color_discrete_sequence=["#FFB800", "#1E2229", "#FF7A00", "#6C5CE7"]
                )
                fig_donut = estilizar_figura(fig_donut, "4. Distribución por Estado de Orden")
                fig_donut.update_traces(
                    textinfo="label+percent",
                    textfont=dict(size=12, color="#FFFFFF"),
                    insidetextorientation="horizontal"
                )
                st.plotly_chart(fig_donut, width="stretch", theme=None)
            else:
                st.info("Sin registros de estado.")

    # --- FILA 3 ---
    r3_c1, r3_c2 = st.columns(2)

    with r3_c1:
        with st.container(border=True):
            # 5. TOP CIUDADES CON CIFRAS NUMÉRICAS
            if "ciudad" in df_f.columns and len(df_f) > 0:
                v_ciu = df_f.groupby("ciudad", as_index=False)["total_venta"].sum().sort_values("total_venta", ascending=False).head(5)
                fig_ciu = px.bar(
                    v_ciu,
                    x="ciudad",
                    y="total_venta",
                    text="total_venta",
                    color="ciudad",
                    color_discrete_sequence=PALETA_COLORES,
                    labels={"total_venta": "Ingresos ($)", "ciudad": "Ciudad"}
                )
                fig_ciu = estilizar_figura(fig_ciu, "5. Top 5 Ciudades por Desempeño")
                fig_ciu.update_traces(
                    texttemplate='$%{text:,.2f}',
                    textposition='outside',
                    textfont=dict(color="#1A1D20", size=11, family="Plus Jakarta Sans")
                )
                max_c = v_ciu["total_venta"].max()
                fig_ciu.update_yaxes(range=[0, max_c * 1.22])
                fig_ciu.update_layout(showlegend=False)
                st.plotly_chart(fig_ciu, width="stretch", theme=None)
            else:
                st.info("Sin registros de ciudad.")

    with r3_c2:
        with st.container(border=True):
            # 6. VOLUMEN DE UNIDADES
            if "categoria" in df_f.columns and len(df_f) > 0:
                v_comb = df_f.groupby("categoria", as_index=False)["cantidad"].sum().sort_values("cantidad", ascending=False)
                fig_comb = px.bar(
                    v_comb,
                    x="categoria",
                    y="cantidad",
                    text="cantidad",
                    color_discrete_sequence=["#1E2229"],
                    labels={"cantidad": "Unidades Vendidas", "categoria": "Categoría"}
                )
                fig_comb = estilizar_figura(fig_comb, "6. Volumen de Unidades por Categoría")
                fig_comb.update_traces(
                    texttemplate='%{text:,} und',
                    textposition='outside',
                    textfont=dict(color="#1A1D20", size=11, family="Plus Jakarta Sans")
                )
                max_u = v_comb["cantidad"].max()
                fig_comb.update_yaxes(range=[0, max_u * 1.22])
                st.plotly_chart(fig_comb, width="stretch", theme=None)
            else:
                st.info("Sin registros.")

with tab_edit:
    st.markdown("""
    <div style='background: white; padding: 20px; border-radius: 18px; border: 1px solid #ECEFF2; margin-bottom: 20px;'>
        <h3 style='margin:0; color:#1A1D20; font-weight:800;'>Editor Directo en Base de Datos</h3>
        <p style='color:#7A828A; font-size:13px; margin-top:4px;'>Modifica valores en la tabla. Las fórmulas se calculan automáticamente al guardar.</p>
    </div>
    """, unsafe_allow_html=True)
    
    df_editado = st.data_editor(
        df,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        disabled=["total_venta"]
    )
    
    if st.button("💾 Guardar cambios en Google Drive", type="primary"):
        try:
            with st.spinner("Guardando en Google Drive..."):
                save_data_to_drive(FILE_ID, df_editado)
            st.success("¡Base de datos actualizada correctamente en Google Drive!")
            st.cache_data.clear()
            st.rerun()
        except Exception as e:
            st.error(f"Error al guardar: {e}")
