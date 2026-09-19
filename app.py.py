import streamlit as st
import pandas as pd
import plotly.express as px

# 1. Configuración de la página
st.set_page_config(
    page_title="Dashboard Centro de Costos",
    page_icon="📊",
    layout="wide"
)

# 2. Carga y caché de datos
@st.cache_data
def load_data():
    df = pd.read_excel('Reporte CECO.xlsx', sheet_name='REPORTE_CENTRO_COSTOS_AVALOS_AS')
    # Limpieza de nulos en columnas clave
    df['CENTRO_DE_COSTOS_NOMBRE'] = df['CENTRO_DE_COSTOS_NOMBRE'].fillna('Sin CECO')
    df['Sucursal'] = df['Sucursal'].fillna('Sin Sucursal')
    df['Comportamiento_del_Gasto'] = df['Comportamiento_del_Gasto'].fillna('Otros')
    return df

df = load_data()

# 3. Menú de navegación en la barra lateral
st.sidebar.title("📌 Navegación")
pagina = st.sidebar.radio(
    "Seleccionar Página:",
    ["1. Resumen General", "2. Detalle por CECO", "3. Análisis Mensual y Tendencia"]
)

# --- FILTROS GLOBALES EN SIDEBAR ---
st.sidebar.markdown("---")
st.sidebar.header("🔍 Filtros Globales")

anios = st.sidebar.multiselect("Año:", options=sorted(df['Año'].dropna().unique()), default=sorted(df['Año'].dropna().unique()))
sucursales = st.sidebar.multiselect("Sucursal:", options=sorted(df['Sucursal'].unique()), default=sorted(df['Sucursal'].unique()))

# Aplicar filtros base
df_filtered = df[(df['Año'].isin(anios)) & (df['Sucursal'].isin(sucursales))]

# ==========================================
# PÁGINA 1: RESUMEN GENERAL
# ==========================================
if pagina == "1. Resumen General":
    st.title("📊 Resumen Ejecutivo - Centro de Costos")
    st.markdown("---")

    # KPIs Principales
    col1, col2, col3, col4 = st.columns(4)
    total_saldo = df_filtered['Saldo'].sum()
    total_asientos = df_filtered['Asiento'].nunique()
    total_cecos = df_filtered['CENTRO_DE_COSTOS_NOMBRE'].nunique()
    promedio_asiento = df_filtered['Saldo'].mean() if len(df_filtered) > 0 else 0

    col1.metric("Gasto Total", f"S/ {total_saldo:,.2f}")
    col2.metric("N° Registros/Asientos", f"{total_asientos:,}")
    col3.metric("Centros de Costo", f"{total_cecos}")
    col4.metric("Gasto Promedio por Reg.", f"S/ {promedio_asiento:,.2f}")

    st.markdown("---")

    # Gráficos de Resumen
    col_left, col_right = st.columns(2)

    with col_left:
        # Costo por Comportamiento del Gasto
        fig_comp = px.pie(
            df_filtered, 
            names='Comportamiento_del_Gasto', 
            values='Saldo',
            title="Distribución por Comportamiento del Gasto",
            hole=0.4
        )
        st.plotly_chart(fig_comp, use_container_width=True)

    with col_right:
        # Costo por Sucursal
        df_suc = df_filtered.groupby('Sucursal')['Saldo'].sum().reset_index()
        fig_suc = px.bar(
            df_suc, 
            x='Sucursal', 
            y='Saldo', 
            text_auto='.2s',
            title="Gasto Total por Sucursal",
            color='Sucursal'
        )
        st.plotly_chart(fig_suc, use_container_width=True)

# ==========================================
# PÁGINA 2: DETALLE POR CENTRO DE COSTO
# ==========================================
elif pagina == "2. Detalle por CECO":
    st.title("📌 Detalle por Centro de Costos")
    st.markdown("---")

    # Filtro específico de CECO
    ceco_sel = st.selectbox("Seleccionar Centro de Costos Específico:", options=["Todos"] + sorted(list(df_filtered['CENTRO_DE_COSTOS_NOMBRE'].unique())))

    if ceco_sel != "Todos":
        df_ceco = df_filtered[df_filtered['CENTRO_DE_COSTOS_NOMBRE'] == ceco_sel]
    else:
        df_ceco = df_filtered

    # Gráfico de Top Centros de Costo
    df_top_ceco = df_ceco.groupby('CENTRO_DE_COSTOS_NOMBRE')['Saldo'].sum().reset_index().sort_values(by='Saldo', ascending=False).head(10)
    fig_top = px.bar(
        df_top_ceco, 
        y='CENTRO_DE_COSTOS_NOMBRE', 
        x='Saldo', 
        orientation='h',
        title="Top 10 Centros de Costos con Mayor Gasto",
        color='Saldo',
        color_continuous_scale='Reds'
    )
    fig_top.update_layout(yaxis={'categoryorder':'total ascending'})
    st.plotly_chart(fig_top, use_container_width=True)

    # Tabla Dinámica / Matriz
    st.subheader("📋 Resumen por Categoría y Nombre de Cuenta")
    tabla_resumen = df_ceco.groupby(['Categoría', 'Nombre de la Cuenta'])['Saldo'].sum().reset_index()
    tabla_resumen['Saldo'] = tabla_resumen['Saldo'].apply(lambda x: f"S/ {x:,.2f}")
    st.dataframe(tabla_resumen, use_container_width=True)

# ==========================================
# PÁGINA 3: ANÁLISIS MENSUAL Y TENDENCIA
# ==========================================
elif pagina == "3. Análisis Mensual y Tendencia":
    st.title("📈 Evolución y Tendencia de Costos")
    st.markdown("---")

    # Evolución por Mes
    df_mes = df_filtered.groupby(['Mes', 'Comportamiento_del_Gasto'])['Saldo'].sum().reset_index()
    fig_mes = px.line(
        df_mes, 
        x='Mes', 
        y='Saldo', 
        color='Comportamiento_del_Gasto',
        markers=True,
        title="Evolución Mensual del Gasto por Comportamiento"
    )
    fig_mes.update_xaxes(dtick=1)
    st.plotly_chart(fig_mes, use_container_width=True)

    # Detalle de Transacciones
    st.subheader("🔎 Explorer de Asientos y Transacciones")
    st.dataframe(
        df_filtered[['Fecha_Contabilizacion', 'Asiento', 'CENTRO_DE_COSTOS_NOMBRE', 'Nombre de la Cuenta', 'Comentario_Asiento', 'Saldo']],
        use_container_width=True
    )