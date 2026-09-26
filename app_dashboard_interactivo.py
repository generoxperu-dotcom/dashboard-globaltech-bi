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
from plotly.subplots import make_subplots

# --- 1. CONFIGURACIÓN VISUAL GENERAL ---
st.set_page_config(
    page_title="Management Dashboard | GlobalTech",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- 2. PALETA DE COLORES CORPORATIVA ---
PALETA_COLORES = ["#FFB800", "#1E2229", "#FF7A00", "#6C5CE7", "#00B894", "#E17055"]

# --- 3. FUNCIONES AUXILIARES GLOBALES ---
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

def render_kpi_card(titulo, val_act, val_ant, val_meta, badge_txt, badge_class, fill_class, es_moneda=False):
    """Genera el HTML compacto de las tarjetas KPI sin indentaciones conflictivas."""
    if es_moneda:
        txt_actual = f"${val_act:,.2f}"
        txt_meta = f"${val_meta:,.2f}"
    else:
        txt_actual = f"{val_act:,}"
        txt_meta = f"{val_meta:,}"

    delta = ((val_act - val_ant) / val_ant * 100) if val_ant > 0 else 0.0
    delta_class = "delta-pos" if delta >= 0 else "delta-neg"
    delta_sign = "+" if delta >= 0 else ""
    delta_icon = "▲" if delta >= 0 else "▼"

    cumplimiento = (val_act / val_meta * 100) if val_meta > 0 else 0.0
    ancho_barra = min(100.0, max(0.0, cumplimiento))

    return (
        f'<div class="kpi-card">'
        f'<div class="kpi-top">'
        f'<span class="kpi-label">{titulo}</span>'
        f'<span class="kpi-badge {badge_class}">{badge_txt}</span>'
        f'</div>'
        f'<div class="kpi-value">{txt_actual}</div>'
        f'<div class="kpi-progress-bg">'
        f'<div class="kpi-progress-fill {fill_class}" style="width: {ancho_barra:.1f}%;"></div>'
        f'</div>'
        f'<div class="kpi-bottom">'
        f'<span class="kpi-meta-text">🎯 Meta: <b>{txt_meta}</b> ({cumplimiento:.0f}%)</span>'
        f'<span class="kpi-delta-pill {delta_class}">{delta_icon} {delta_sign}{delta:.1f}% vs ant.</span>'
        f'</div>'
        f'</div>'
    )

def estilizar_figura(fig, titulo=""):
    """Formato común de alto contraste para que números y ejes sean 100% legibles."""
    fig.update_layout(
        template="plotly_white",
        title=dict(
            text=f"<b>{titulo}</b>",
            font=dict(size=14, color="#1A1D20", family="Plus Jakarta Sans"),
            x=0.02,
            y=0.96
        ),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        margin=dict(l=25, r=25, t=55, b=25),
        font=dict(family="Plus Jakarta Sans", color="#1A1D20", size=11),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=10, color="#1A1D20")
        )
    )
    fig.update_xaxes(
        tickfont=dict(color="#1A1D20", size=10, family="Plus Jakarta Sans"),
        title_font=dict(color="#1A1D20", size=11, family="Plus Jakarta Sans"),
        gridcolor="#F0F2F5",
        zeroline=False,
        showline=True,
        linecolor="#E2E8F0"
    )
    fig.update_yaxes(
        tickfont=dict(color="#1A1D20", size=10, family="Plus Jakarta Sans"),
        title_font=dict(color="#1A1D20", size=11, family="Plus Jakarta Sans"),
        gridcolor="#F0F2F5",
        zeroline=False,
        showline=True,
        linecolor="#E2E8F0"
    )
    return fig

# --- 4. ESTILOS CSS INYECTADOS ---
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

    /* Flecha reapertura de sidebar */
    header[data-testid="stHeader"] {
        background: transparent !important;
    }
    button[data-testid="stSidebarCollapsedControl"],
    div[data-testid="stSidebarCollapsedControl"] button {
        background-color: #FFFFFF !important;
        color: #1A1D20 !important;
        border: 1px solid #ECEFF2 !important;
        border-radius: 12px !important;
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08) !important;
        margin-left: 12px !important;
        margin-top: 10px !important;
        visibility: visible !important;
        display: flex !important;
        align-items: center;
        justify-content: center;
        transition: all 0.2s ease !important;
    }
    button[data-testid="stSidebarCollapsedControl"]:hover,
    div[data-testid="stSidebarCollapsedControl"] button:hover {
        background-color: #FFB800 !important;
        color: #FFFFFF !important;
    }

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
        padding: 20px 22px;
        border-radius: 20px;
        border: 1px solid #ECEFF2;
        box-shadow: 0px 4px 18px rgba(0, 0, 0, 0.03);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        margin-bottom: 15px;
        min-height: 148px;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        box-shadow: 0px 8px 24px rgba(0, 0, 0, 0.06);
    }
    .kpi-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 4px;
    }
    .kpi-label {
        font-size: 11.5px;
        font-weight: 700;
        color: #8C94A0 !important;
        text-transform: uppercase;
        letter-spacing: 0.6px;
    }
    .kpi-value {
        font-size: 27px;
        font-weight: 800;
        color: #1A1D20 !important;
        margin: 2px 0 6px 0;
        line-height: 1.2;
    }
    .kpi-badge {
        padding: 5px 12px;
        border-radius: 20px;
        font-size: 11.5px;
        font-weight: 700;
    }
    .badge-yellow { background: #FFB800; color: #FFFFFF; }
    .badge-dark { background: #1E2229; color: #FFFFFF; }
    .badge-orange { background: #FF7A00; color: #FFFFFF; }
    .badge-purple { background: #6C5CE7; color: #FFFFFF; }

    /* Barra de Progreso a la Meta */
    .kpi-progress-bg {
        width: 100%;
        height: 6px;
        background-color: #F1F3F5;
        border-radius: 8px;
        overflow: hidden;
        margin: 6px 0 10px 0;
    }
    .kpi-progress-fill {
        height: 100%;
        border-radius: 8px;
        transition: width 0.3s ease;
    }
    .fill-yellow { background: linear-gradient(90deg, #FFB800, #FFA000); }
    .fill-dark { background: linear-gradient(90deg, #1E2229, #4B5563); }
    .fill-orange { background: linear-gradient(90deg, #FF7A00, #FF5500); }
    .fill-purple { background: linear-gradient(90deg, #6C5CE7, #805AD5); }

    .kpi-bottom {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding-top: 8px;
        border-top: 1px solid #F1F3F5;
        font-size: 11.5px;
    }
    .kpi-meta-text {
        color: #64748B !important;
        font-weight: 500;
    }
    .kpi-meta-text b {
        color: #1A1D20 !important;
        font-weight: 700;
    }
    .kpi-delta-pill {
        padding: 3px 9px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 700;
        display: inline-flex;
        align-items: center;
        white-space: nowrap;
    }
    .delta-pos { background: #E6F9F0; color: #00B894; }
    .delta-neg { background: #FEECEC; color: #E17055; }

    /* Contenedores de Gráficos */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background-color: #FFFFFF !important;
        border-radius: 22px !important;
        border: 1px solid #ECEFF2 !important;
        box-shadow: 0px 4px 18px rgba(0, 0, 0, 0.03) !important;
        padding: 14px !important;
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

# --- 5. GOOGLE DRIVE SERVICES ---
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

# --- 6. CARGA DE DATOS ---
FILE_ID = st.secrets["drive_settings"]["file_id"]

try:
    with st.spinner("Conectando con Google Drive..."):
        df = load_data_from_drive(FILE_ID)
except Exception as e:
    st.error(f"Error al conectar con Google Drive: {e}")
    st.stop()

# --- 7. ENCABEZADO SUPERIOR TIPO BANNER ---
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
    with st.sidebar.expander("🎯 Configuración de Metas"):
        factor_meta = st.slider("Crecimiento Objetivo vs Período Anterior:", min_value=5, max_value=50, value=15, step=5, format="%d%%") / 100

    if st.sidebar.button("🔄 Actualizar Datos", help="Recarga el libro en memoria desde Google Drive"):
        st.cache_data.clear()
        st.rerun()

    # Filtrar DataFrame
    df_f = df.copy()
    if ciudad_sel != "Todas" and "ciudad" in df_f.columns:
        df_f = df_f[df_f["ciudad"] == ciudad_sel]
    if cat_sel != "Todas" and "categoria" in df_f.columns:
        df_f = df_f[df_f["categoria"] == cat_sel]
    if estado_sel != "Todos" and "estado" in df_f.columns:
        df_f = df_f[df_f["estado"] == estado_sel]

    # Cálculos actuales
    total_ventas = float(df_f["total_venta"].sum())
    total_unidades = int(df_f["cantidad"].sum())
    num_ordenes = len(df_f)
    ticket_medio = (total_ventas / num_ordenes) if num_ordenes > 0 else 0.0

    # Estimación período anterior
    try:
        df_fechas = df_f.dropna(subset=['fecha']).copy()
        df_fechas['dt'] = pd.to_datetime(df_fechas['fecha'], errors='coerce')
        df_fechas = df_fechas.dropna(subset=['dt']).sort_values('dt')
        
        if len(df_fechas) >= 4:
            mid = len(df_fechas) // 2
            df_ant = df_fechas.iloc[:mid]
            ventas_ant = float(df_ant['total_venta'].sum())
            unidades_ant = int(df_ant['cantidad'].sum())
            ordenes_ant = len(df_ant)
            ticket_ant = (ventas_ant / ordenes_ant) if ordenes_ant > 0 else 0.0
        else:
            ventas_ant = total_ventas * 0.90
            unidades_ant = int(total_unidades * 0.90)
            ticket_ant = ticket_medio * 0.95
            ordenes_ant = int(num_ordenes * 0.92)
    except Exception:
        ventas_ant = total_ventas * 0.90
        unidades_ant = int(total_unidades * 0.90)
        ticket_ant = ticket_medio * 0.95
        ordenes_ant = int(num_ordenes * 0.92)

    meta_ventas = ventas_ant * (1 + factor_meta) if ventas_ant > 0 else total_ventas * 1.15
    meta_unidades = int(unidades_ant * (1 + factor_meta)) if unidades_ant > 0 else int(total_unidades * 1.15)
    meta_ticket = ticket_ant * (1 + (factor_meta * 0.6)) if ticket_ant > 0 else ticket_medio * 1.08
    meta_ordenes = int(ordenes_ant * (1 + factor_meta)) if ordenes_ant > 0 else int(num_ordenes * 1.15)

    # --- TARJETAS KPI ---
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(render_kpi_card("Ingresos Totales", total_ventas, ventas_ant, meta_ventas, "Ventas", "badge-yellow", "fill-yellow", es_moneda=True), unsafe_allow_html=True)
    with k2:
        st.markdown(render_kpi_card("Unidades Totales", total_unidades, unidades_ant, meta_unidades, "Items", "badge-dark", "fill-dark", es_moneda=False), unsafe_allow_html=True)
    with k3:
        st.markdown(render_kpi_card("Ticket Promedio", ticket_medio, ticket_ant, meta_ticket, "Promedio", "badge-orange", "fill-orange", es_moneda=True), unsafe_allow_html=True)
    with k4:
        st.markdown(render_kpi_card("Total Órdenes", num_ordenes, ordenes_ant, meta_ordenes, "Órdenes", "badge-purple", "fill-purple", es_moneda=False), unsafe_allow_html=True)

    st.write("")

    # =========================================================================
    # --- MATRIZ 3X2 DE GRÁFICOS AVANZADOS (INSPIRADOS EN LAS PRESENTACIONES) ---
    # =========================================================================

    # --------------------------- FILA 1 ---------------------------
    r1_c1, r1_c2 = st.columns(2)

    with r1_c1:
        with st.container(border=True):
            # 1. DISPERSIÓN MULTIDIMENSIONAL 4D/5D
            if len(df_f) > 0:
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
                fig_disp = estilizar_figura(fig_disp, "1. Dispersión Multidimensional 4D (Volumen vs Precio)")
                st.plotly_chart(fig_disp, width="stretch", theme=None)
            else:
                st.info("Sin registros.")

    with r1_c2:
        with st.container(border=True):
            # 2. TREEMAP JERÁRQUICO MULTICAPA
            if len(df_f) > 0 and all(c in df_f.columns for c in ["ciudad", "categoria", "producto"]):
                df_tree = df_f.copy()
                for c in ["ciudad", "categoria", "producto"]:
                    df_tree[c] = df_tree[c].fillna("S/D").astype(str)

                fig_tree = px.treemap(
                    df_tree,
                    path=["ciudad", "categoria", "producto"],
                    values="total_venta",
                    color="categoria",
                    color_discrete_sequence=PALETA_COLORES
                )
                fig_tree = estilizar_figura(fig_tree, "2. Treemap Jerárquico (Ciudad ➔ Categoría ➔ Producto)")
                fig_tree.update_traces(
                    textinfo="label+value+percent entry",
                    texttemplate="<b>%{label}</b><br>$%{value:,.0f}",
                    textfont=dict(size=11, family="Plus Jakarta Sans")
                )
                st.plotly_chart(fig_tree, width="stretch", theme=None)
            else:
                st.info("Sin registros suficientes para armar jerarquía.")

    # --------------------------- FILA 2 ---------------------------
    r2_c1, r2_c2 = st.columns(2)

    with r2_c1:
        with st.container(border=True):
            # 3. DIAGRAMA DE PARETO 80/20 (TOP PRODUCTOS CLAVE)
            if "producto" in df_f.columns and len(df_f) > 0:
                v_prod = df_f.groupby("producto", as_index=False)["total_venta"].sum().sort_values("total_venta", ascending=False).head(10)
                v_prod["cum_sum"] = v_prod["total_venta"].cumsum()
                total_pareto = v_prod["total_venta"].sum()
                v_prod["cum_pct"] = (v_prod["cum_sum"] / total_pareto) * 100

                fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
                
                # Barras de Ventas
                fig_pareto.add_trace(
                    go.Bar(
                        x=v_prod["producto"],
                        y=v_prod["total_venta"],
                        name="Ventas ($)",
                        marker_color="#FFB800",
                        text=v_prod["total_venta"],
                        texttemplate="$%{text:,.0f}",
                        textposition="outside",
                        textfont=dict(size=10, color="#1A1D20")
                    ),
                    secondary_y=False
                )
                
                # Línea acumulada
                fig_pareto.add_trace(
                    go.Scatter(
                        x=v_prod["producto"],
                        y=v_prod["cum_pct"],
                        name="% Acumulado",
                        mode="lines+markers",
                        line=dict(color="#1E2229", width=3),
                        marker=dict(size=6, color="#FF7A00")
                    ),
                    secondary_y=True
                )

                # Línea guía de corte al 80%
                fig_pareto.add_hline(
                    y=80,
                    line_dash="dot",
                    line_color="#E17055",
                    line_width=2,
                    annotation_text="Regla 80%",
                    annotation_position="top right",
                    secondary_y=True
                )

                fig_pareto = estilizar_figura(fig_pareto, "3. Diagrama de Pareto 80/20 (Top Productos)")
                fig_pareto.update_yaxes(title_text="Ventas ($)", secondary_y=False, showgrid=True, gridcolor="#F0F2F5")
                fig_pareto.update_yaxes(title_text="% Acumulado", range=[0, 110], secondary_y=True, showgrid=False)
                fig_pareto.update_xaxes(tickangle=-25)
                st.plotly_chart(fig_pareto, width="stretch", theme=None)
            else:
                st.info("Sin registros de productos.")

    with r2_c2:
        with st.container(border=True):
            # 4. TENDENCIA HISTÓRICA DE VENTAS
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
                fig_line = estilizar_figura(fig_line, "4. Tendencia Histórica de Ventas")
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

    # --------------------------- FILA 3 ---------------------------
    r3_c1, r3_c2 = st.columns(2)

    with r3_c1:
        with st.container(border=True):
            # 5. DIAGRAMA DE FLUJO SANKEY (CIUDAD -> CATEGORÍA -> ESTADO)
            if len(df_f) > 0 and all(c in df_f.columns for c in ["ciudad", "categoria", "estado"]):
                ciudades_list = list(df_f["ciudad"].dropna().unique())
                categorias_list = list(df_f["categoria"].dropna().unique())
                estados_list = list(df_f["estado"].dropna().unique())

                all_nodes = ciudades_list + categorias_list + estados_list
                node_indices = {name: i for i, name in enumerate(all_nodes)}

                # Enlace 1: Ciudad -> Categoría
                df_l1 = df_f.groupby(["ciudad", "categoria"], as_index=False)["total_venta"].sum()
                # Enlace 2: Categoría -> Estado
                df_l2 = df_f.groupby(["categoria", "estado"], as_index=False)["total_venta"].sum()

                srcs = [node_indices[c] for c in df_l1["ciudad"]] + [node_indices[cat] for cat in df_l2["categoria"]]
                tgts = [node_indices[cat] for cat in df_l1["categoria"]] + [node_indices[est] for est in df_l2["estado"]]
                vals = list(df_l1["total_venta"]) + list(df_l2["total_venta"])

                # Paleta de colores para los nodos por etapa
                node_colors = (
                    ["#FFB800"] * len(ciudades_list) +
                    ["#1E2229"] * len(categorias_list) +
                    ["#6C5CE7"] * len(estados_list)
                )

                fig_sankey = go.Figure(data=[go.Sankey(
                    node=dict(
                        pad=15,
                        thickness=18,
                        line=dict(color="#ECEFF2", width=1),
                        label=all_nodes,
                        color=node_colors
                    ),
                    link=dict(
                        source=srcs,
                        target=tgts,
                        value=vals,
                        color="rgba(255, 184, 0, 0.28)"
                    )
                )])
                fig_sankey = estilizar_figura(fig_sankey, "5. Diagrama de Flujo Sankey (Ciudad ➔ Categoría ➔ Estado)")
                st.plotly_chart(fig_sankey, width="stretch", theme=None)
            else:
                st.info("Sin registros suficientes para el diagrama de flujo.")

    with r3_c2:
        with st.container(border=True):
            # 6. RADAR / ARAÑA POLAR MULTICRITERIO 360°
            if "categoria" in df_f.columns and len(df_f) > 0:
                df_radar = df_f.groupby("categoria").agg(
                    ventas=("total_venta", "sum"),
                    unidades=("cantidad", "sum"),
                    ordenes=("id_transaccion", "count"),
                    ticket=("total_venta", "mean")
                ).reset_index()

                metricas = ["ventas", "unidades", "ordenes", "ticket"]
                nombres_ejes = ["Facturación ($)", "Unidades", "Nº Órdenes", "Ticket Medio"]

                # Normalización 0-100 para comparar en la misma escala
                for m in metricas:
                    max_val = df_radar[m].max()
                    df_radar[m + "_norm"] = (df_radar[m] / max_val * 100) if max_val > 0 else 0

                fig_radar = go.Figure()
                for idx, cat_name in enumerate(df_radar["categoria"].unique()):
                    row = df_radar[df_radar["categoria"] == cat_name].iloc[0]
                    valores_radar = [row[m + "_norm"] for m in metricas]
                    valores_radar.append(valores_radar[0])  # Cerrar la figura polar
                    ejes_cerrados = nombres_ejes + [nombres_ejes[0]]

                    color_cat = PALETA_COLORES[idx % len(PALETA_COLORES)]
                    fig_radar.add_trace(go.Scatterpolar(
                        r=valores_radar,
                        theta=ejes_cerrados,
                        fill='toself',
                        name=str(cat_name),
                        line=dict(color=color_cat, width=2.5),
                        opacity=0.65
                    ))

                fig_radar = estilizar_figura(fig_radar, "6. Radar Polar Multicriterio 360° por Categoría")
                fig_radar.update_layout(
                    polar=dict(
                        radialaxis=dict(
                            visible=True,
                            range=[0, 105],
                            tickfont=dict(size=9, color="#7A828A"),
                            gridcolor="#F0F2F5"
                        ),
                        angularaxis=dict(
                            tickfont=dict(size=11, color="#1A1D20", family="Plus Jakarta Sans"),
                            linecolor="#E2E8F0"
                        ),
                        bgcolor="#FFFFFF"
                    )
                )
                st.plotly_chart(fig_radar, width="stretch", theme=None)
            else:
                st.info("Sin registros para el análisis polar.")

# --- 8. PESTAÑA: EDITOR DE BASE DE DATOS ---
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
