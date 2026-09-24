import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# ------------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL DE LA PÁGINA
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="GlobalTech | Executive Interactive Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos visuales para tarjetas y tipografía
st.markdown("""
    <style>
    .main { background-color: #F8FAFC; }
    div[data-testid="stMetricValue"] { font-size: 1.8rem; font-weight: 700; color: #0F172A; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        height: 42px;
        padding-left: 18px;
        padding-right: 18px;
        background-color: #FFFFFF;
        border-radius: 6px;
        border: 1px solid #E2E8F0;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1E3A8A !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 1. CARGA DE DATOS DESDE EXCEL
# ------------------------------------------------------------------------------
@st.cache_data
def cargar_datos():
    archivo = 'Dataset_Integrado_Final_Sesion4.xlsx'
    df = pd.read_excel(archivo, sheet_name='Consolidado_Multiarea')
    df['Fecha_Operacion'] = pd.to_datetime(df['Fecha_Operacion'])
    return df

try:
    df_raw = cargar_datos()
except Exception as e:
    st.error(f"Error al cargar el archivo de datos: {e}")
    st.stop()

# ------------------------------------------------------------------------------
# 2. BARRA LATERAL: FILTROS INTERACTIVOS DINÁMICOS
# ------------------------------------------------------------------------------
st.sidebar.title("Filtros Directivos")
st.sidebar.markdown("Segmentación de cartera en tiempo real")

# Filtro de Países
paises_disp = sorted(df_raw['Pais_Sede'].dropna().unique().tolist())
paises_sel = st.sidebar.multiselect("País / Filial", paises_disp, default=paises_disp)

# Filtro de Canales
canales_disp = sorted(df_raw['Canal'].dropna().unique().tolist())
canales_sel = st.sidebar.multiselect("Canal Comercial", canales_disp, default=canales_disp)

# Filtro de Tiers
tiers_disp = sorted(df_raw['Tier_Estrategico'].dropna().unique().tolist())
tiers_sel = st.sidebar.multiselect("Tier de Cliente", tiers_disp, default=tiers_disp)

# Filtro de Estado de Cobranza
estados_disp = sorted(df_raw['Estado_Cobranza'].dropna().unique().tolist())
estados_sel = st.sidebar.multiselect("Estado de Cobranza", estados_disp, default=estados_disp)

# Aplicar filtros
df_filtrado = df_raw.copy()
if paises_sel:
    df_filtrado = df_filtrado[df_filtrado['Pais_Sede'].isin(paises_sel)]
if canales_sel:
    df_filtrado = df_filtrado[df_filtrado['Canal'].isin(canales_sel)]
if tiers_sel:
    df_filtrado = df_filtrado[df_filtrado['Tier_Estrategico'].isin(tiers_sel)]
if estados_sel:
    df_filtrado = df_filtrado[df_filtrado['Estado_Cobranza'].isin(estados_sel)]

# ------------------------------------------------------------------------------
# 3. CABECERA & TARJETAS KPI EJECUTIVAS
# ------------------------------------------------------------------------------
st.title("📊 GlobalTech | Cuadro de Mando Integrado C-Level")
st.markdown(f"**Base analizada:** {len(df_filtrado)} de {len(df_raw)} transacciones registradas | **Moneda:** USD")

if df_filtrado.empty:
    st.warning("⚠️ No existen transacciones con la combinación de filtros seleccionada.")
    st.stop()

total_fact = df_filtrado['Monto_Facturado_USD'].sum()
total_cobr = df_filtrado['Monto_Neto_Cobrado_USD'].sum()
total_mora = df_filtrado[df_filtrado['Estado_Cobranza'] != 'Cobrado / Al Día']['Monto_Facturado_USD'].sum()
margen_neto = df_filtrado['Margen_Operativo_USD'].sum()
pct_mora = (total_mora / total_fact * 100) if total_fact > 0 else 0
pct_cobrado = (total_cobr / total_fact * 100) if total_fact > 0 else 0
margen_pct = (margen_neto / total_fact * 100) if total_fact > 0 else 0

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Facturación Total", f"${total_fact:,.2f}", delta=f"{len(df_filtrado)} pedidos")
with col2:
    st.metric("Cartera Cobrada", f"${total_cobr:,.2f}", delta=f"{pct_cobrado:.1f}% Cobrado", delta_color="normal")
with col3:
    st.metric("Cartera en Mora", f"${total_mora:,.2f}", delta=f"-{pct_mora:.1f}% en Riesgo", delta_color="inverse")
with col4:
    st.metric("Margen Operativo", f"{margen_pct:.1f}%", delta=f"${margen_neto:,.2f} neto")

st.markdown("---")

# ------------------------------------------------------------------------------
# 4. PESTAÑAS: GRÁFICOS, TABLAS DINÁMICAS Y AUDITORÍA
# ------------------------------------------------------------------------------
tab_graficos, tab_pivot, tab_datos = st.tabs([
    "📈 Gráficos Analíticos", 
    "🗂️ Constructor de Tablas Dinámicas", 
    "📋 Detalle de Transacciones"
])

# Paleta corporativa para cobranzas
colores_cobranza = {
    'Cobrado / Al Día': '#10B981',
    'Mora Leve (<30d)': '#F59E0B',
    'Mora Crítica (>30d)': '#EF4444'
}

with tab_graficos:
    c_g1, c_g2 = st.columns(2)
    with c_g1:
        # Gráfico 1: Barras Apiladas
        fig1 = px.bar(
            df_filtrado, x='Sector_Industria', y='Monto_Facturado_USD', color='Estado_Cobranza',
            title='<b>1. Facturación y Cobranza por Industria</b>',
            barmode='stack', color_discrete_map=colores_cobranza,
            labels={'Sector_Industria': 'Industria', 'Monto_Facturado_USD': 'Facturación (USD)'}
        )
        fig1.update_layout(height=400, plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
        st.plotly_chart(fig1, use_container_width=True)

    with c_g2:
        # Gráfico 2: Dona por Tier
        df_tier = df_filtrado.groupby('Tier_Estrategico')['Monto_Facturado_USD'].sum().reset_index()
        fig2 = px.pie(
            df_tier, values='Monto_Facturado_USD', names='Tier_Estrategico', hole=0.45,
            title='<b>2. Participación por Tier de Cliente</b>',
            color_discrete_sequence=['#1E3A8A', '#0D9488', '#F59E0B']
        )
        fig2.update_traces(textposition='inside', textinfo='percent+label')
        fig2.update_layout(height=400, plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
        st.plotly_chart(fig2, use_container_width=True)

    c_g3, c_g4 = st.columns(2)
    with c_g3:
        # Gráfico 3: Boxplot Días de Mora
        fig3 = px.box(
            df_filtrado, x='Sector_Industria', y='Dias_Mora', color='Estado_Cobranza',
            title='<b>3. Días de Mora por Sector Industrial</b>',
            color_discrete_map=colores_cobranza,
            labels={'Dias_Mora': 'Días de Retraso', 'Sector_Industria': 'Sector'}
        )
        fig3.update_layout(height=400, plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
        st.plotly_chart(fig3, use_container_width=True)

    with c_g4:
        # Gráfico 4: Cartera por KAM
        df_kam = df_filtrado.groupby(['Ejecutivo_KAM', 'Estado_Cobranza'])['Monto_Facturado_USD'].sum().reset_index()
        fig4 = px.bar(
            df_kam, y='Ejecutivo_KAM', x='Monto_Facturado_USD', color='Estado_Cobranza', orientation='h',
            title='<b>4. Gestión de Cobranza por Ejecutivo Comercial (KAM)</b>',
            color_discrete_map=colores_cobranza,
            labels={'Monto_Facturado_USD': 'Monto (USD)', 'Ejecutivo_KAM': 'Key Account Manager'}
        )
        fig4.update_layout(height=400, plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF')
        st.plotly_chart(fig4, use_container_width=True)

with tab_pivot:
    st.subheader("Configuración de la Tabla Dinámica")
    
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        opciones_filas = ['Sector_Industria', 'Pais_Sede', 'Ejecutivo_KAM', 'Canal', 'Metodo_Pago', 'Tier_Estrategico']
        p_fila = st.selectbox("Filas:", opciones_filas, index=0)
    with col_p2:
        opciones_columnas = ['Estado_Cobranza', 'Tier_Estrategico', 'Canal', 'Pais_Sede']
        p_col = st.selectbox("Columnas:", opciones_columnas, index=0)
    with col_p3:
        mapa_metricas = {
            'Suma Facturación ($ USD)': ('Monto_Facturado_USD', 'sum', '${:,.2f}'),
            'Suma Margen Neto ($ USD)': ('Margen_Operativo_USD', 'sum', '${:,.2f}'),
            'Promedio Días de Mora': ('Dias_Mora', 'mean', '{:.1f} días'),
            'Cantidad de Transacciones': ('ID_Transaccion', 'count', '{:,.0f}')
        }
        p_metrica = st.selectbox("Métrica:", list(mapa_metricas.keys()), index=0)

    if p_fila == p_col:
        st.warning("⚠️ Selecciona campos diferentes para Filas y Columnas.")
    else:
        col_valor, func_agg, fmt = mapa_metricas[p_metrica]
        
        # Generación de la tabla dinámica con totales generales
        tabla_dinamica = pd.pivot_table(
            df_filtrado,
            index=p_fila,
            columns=p_col,
            values=col_valor,
            aggfunc=func_agg,
            fill_value=0,
            margins=True,
            margins_name='Total General'
        )
        
        # Renderizado estilizado en Streamlit
        st.markdown(f"#### Matriz Dinámica: **{p_fila}** vs **{p_col}** ({p_metrica})")
        st.dataframe(
            tabla_dinamica.style.format(fmt),
            use_container_width=True
        )

with tab_datos:
    st.subheader("Auditoría de Transacciones Filtradas")
    cols_mostrar = [
        'ID_Transaccion', 'Fecha_Operacion', 'Cliente_Nombre', 'Pais_Sede',
        'Sector_Industria', 'Monto_Facturado_USD', 'Estado_Cobranza', 'Dias_Mora', 'Ejecutivo_KAM'
    ]
    st.dataframe(df_filtrado[cols_mostrar], use_container_width=True, hide_index=True)
    
    csv_bytes = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Descargar Base Filtrada (.CSV)",
        data=csv_bytes,
        file_name="Reporte_Consolidado_GlobalTech.csv",
        mime="text/csv"
    )
