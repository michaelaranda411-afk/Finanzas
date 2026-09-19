import datetime
import os
from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="Control Presupuestal FP&A - Junior",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "bitacora_alertas" not in st.session_state:
    st.session_state.bitacora_alertas = []

st.markdown(
    """
    <style>
    .main { background-color: #f8fafc; }
    .header-banner {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: #ffffff; padding: 20px 24px; border-radius: 14px;
        margin-bottom: 20px; border-left: 6px solid #a81121;
        box-shadow: 0 4px 14px rgba(0,0,0,0.06);
    }
    .header-banner h1 { color: #ffffff !important; font-size: 24px; font-weight: 800; margin: 0; }
    .header-banner p { color: #94a3b8; margin: 4px 0 0 0; font-size: 13px; }
    
    .kpi-card {
        background-color: #ffffff; border-radius: 12px; padding: 16px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04); border: 1px solid #e2e8f0; border-top: 4px solid #a81121;
    }
    .kpi-title { font-size: 11px; color: #64748b; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 22px; color: #0f172a; font-weight: 800; margin-top: 4px; }
    
    .badge-green { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; background-color: #dcfce7; color: #15803d; }
    .badge-red { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; background-color: #fee2e2; color: #b91c1c; }
    </style>
""",
    unsafe_allow_html=True,
)


# 2. CARGA Y PREPARACIÓN DE DATOS CON RUTA ABSOLUTA Y DINÁMICA
@st.cache_data
def cargar_datos():
    BASE_DIR = Path(__file__).resolve().parent
    ruta_excel = BASE_DIR / "Reporte CECO.xlsx"

    if not ruta_excel.exists():
        ruta_excel = Path(
            r"C:\Users\JUNIOR FP&A\Desktop\Finanzas\POWER BI\Reporte CECO.xlsx"
        )
    if not ruta_excel.exists():
        ruta_excel = Path(r"C:\Users\JUNIOR FP&A\Reporte CECO.xlsx")

    df = pd.read_excel(
        ruta_excel, sheet_name="REPORTE_CENTRO_COSTOS_AVALOS_AS"
    )

    df["Centro de Costos"] = df["CENTRO_DE_COSTOS_NOMBRE"].fillna("Sin CECO")
    df["Sucursal"] = df["Sucursal"].fillna("Sin Sucursal")
    df["Comportamiento"] = df["Comportamiento_del_Gasto"].fillna("Otros")
    df["Clasificación"] = (
        df["Clasificación"].fillna("Gastos Varios")
        if "Clasificación" in df.columns
        else "Gastos Varios"
    )
    df["Nombre de Cuenta"] = df["Nombre de la Cuenta"].fillna("Sin Cuenta")
    df["Glosa / Comentario"] = df["Comentario_Asiento"].fillna("Sin Glosa")

    if "Código de Cuenta" in df.columns:
        df["Código de Cuenta"] = df["Código de Cuenta"].astype(str)
        df["Cuenta Contable"] = (
            df["Código de Cuenta"] + " - " + df["Nombre de Cuenta"]
        )
    else:
        df["Cuenta Contable"] = df["Nombre de Cuenta"]

    def mapear_geo_sucursal(suc):
        s = str(suc).upper().strip()
        if "MTRU" in s or "TRU" in s:
            return "Trujillo (La Libertad)", -8.1116, -79.0286
        elif "LIMA" in s or "LICI" in s or "VECO" in s:
            return "Lima Metropolitano", -12.0464, -77.0428
        elif "CHIC" in s:
            return "Chiclayo (Lambayeque)", -6.7714, -79.8409
        elif "AREQ" in s:
            return "Arequipa", -16.4090, -71.5375
        elif "PIU" in s:
            return "Piura", -5.1945, -80.6328
        elif "CUSC" in s:
            return "Cusco", -13.5319, -71.9675
        elif "HUAN" in s:
            return "Huancayo", -12.0651, -75.2049
        else:
            return "Sede Central (Lima)", -12.0464, -77.0428

    res_geo = df["Sucursal"].apply(mapear_geo_sucursal)
    df["Ciudad_Nombre"] = [r[0] for r in res_geo]
    df["Latitud"] = [r[1] for r in res_geo]
    df["Longitud"] = [r[2] for r in res_geo]

    def asignar_mapa_oficial(row):
        ceco = str(row["Centro de Costos"]).upper()
        if "GERENCIA" in ceco:
            return "Procesos Estratégicos", "Gerencia General"
        elif "NUEVOS NEGOCIOS" in ceco:
            return "Procesos Estratégicos", "Nuevos Negocios"
        elif "COMPRAS" in ceco or "IMPORTACION" in ceco:
            if "IMPORTACION" in ceco:
                return "Procesos Operativos", "Importaciones"
            return "Procesos Operativos", "Compras Locales"
        elif "LOGISTICA" in ceco:
            return "Procesos Operativos", "Logística"
        elif "ALMACEN" in ceco or "TRANSPORTE" in ceco:
            if "TRANSPORTE" in ceco:
                return "Procesos Operativos", "Transporte Interno"
            return "Procesos Operativos", "Gestión de Almacenes"
        elif "FERRETERIA" in ceco:
            return "Procesos Operativos", "Ventas Ferretería"
        elif "CORPORATIVO" in ceco:
            return "Procesos Operativos", "Ventas Corporativo"
        elif "LICITACION" in ceco:
            return "Procesos Operativos", "Licitaciones"
        elif "VENTA" in ceco:
            return "Procesos Operativos", "Ventas Generales"
        elif "DISTRIBUC" in ceco:
            return "Procesos Operativos", "Distribución"
        elif "MARKETING" in ceco:
            return "Procesos Operativos", "Marketing"
        elif "RRHH" in ceco or "RECURSOS HUMANOS" in ceco:
            return "Procesos de Soporte", "Recursos Humanos"
        elif "TESORERIA" in ceco:
            return "Procesos de Soporte", "Tesorería"
        elif "CONTABILIDAD" in ceco:
            return "Procesos de Soporte", "Contabilidad"
        elif "FINANZAS" in ceco or "CONTROL INTERNO" in ceco:
            return "Procesos de Soporte", "Control Interno & Finanzas"
        elif "TI" in ceco or "SISTEMAS" in ceco:
            return "Procesos de Soporte", "TI & Sistemas"
        elif "CREDITO" in ceco or "COBRANZA" in ceco:
            return "Procesos de Soporte", "Créditos y Cobranzas"
        elif "ADMINISTRACION" in ceco:
            return "Procesos de Soporte", "Administración"
        else:
            return "Procesos de Soporte", "Soporte General"

    res = df.apply(asignar_mapa_oficial, axis=1)
    df["Macroproceso"] = [r[0] for r in res]
    df["Subproceso / Área"] = [r[1] for r in res]

    return df


df = cargar_datos()

COLOR_ROJO = "#a81121"
COLOR_NAVY = "#1e293b"
COLOR_AZUL = "#2563eb"
PALETA_ARMONICA = [
    "#a81121",
    "#1e293b",
    "#2563eb",
    "#0d9488",
    "#d97706",
    "#64748b",
    "#0284c7",
]

# 3. SIDEBAR Y FILTROS GLOBALES
with st.sidebar:
    st.markdown(
        """
        <div style="background-color: #a81121; color: white; padding: 16px; border-radius: 12px; margin-bottom: 15px;">
            <h3 style="margin:0; font-size: 16px; color: white;">🔴 CONTROL PRESUPUESTAL</h3>
            <p style="margin:0; font-size: 12px; color: #fecaca; font-weight: bold;">Junior FP&A Control System</p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    if st.button("🖨️ Imprimir Pantalla a PDF"):
        components.html("<script>window.parent.print();</script>", height=0)

    st.markdown("---")
    st.markdown("#### 🎛️ Filtros Globales")

    if "Año" in df.columns:
        anios_tot = sorted(list(df["Año"].dropna().unique()))
        anios_sel = st.multiselect(
            "📅 Año:", options=anios_tot, default=anios_tot
        )
    else:
        anios_sel = []

    meses_totales = (
        sorted(list(df["Mes"].unique())) if "Mes" in df.columns else []
    )
    meses_sel = (
        st.multiselect("📅 Meses:", options=meses_totales, default=meses_totales)
        if meses_totales
        else []
    )

    suc_opt = sorted(list(df["Sucursal"].unique()))
    suc_sel = st.multiselect("🏢 Sucursales:", options=suc_opt, default=suc_opt)

# Aplicación de filtros
df_f = df[df["Sucursal"].isin(suc_sel)].copy()
if anios_sel and "Año" in df_f.columns:
    df_f = df_f[df_f["Año"].isin(anios_sel)]
if meses_sel and "Mes" in df_f.columns:
    df_f = df_f[df_f["Mes"].isin(meses_sel)]

# NAVEGACIÓN
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Resumen General",
    "🗺️ Mapa de Procesos",
    "🏛️ Organigrama por Áreas",
    "🗺️ Mapa Sucursales Perú",
    "🔍 Detalle por CECO / Asientos",
    "🚨 Alertas por Motivo",
])

# TAB 1: RESUMEN GENERAL
with tab1:
    st.markdown(
        '<div class="header-banner"><h1>📊 Resumen Ejecutivo - Control Presupuestal FP&A</h1><p>Panel Ejecutivo de Control Presupuestal y Análisis Financiero | <b>Junior FP&A</b></p></div>',
        unsafe_allow_html=True,
    )

    gasto_total = df_f["Saldo"].sum()
    cecos_cant = df_f["Centro de Costos"].nunique()
    registros = len(df_f)
    promedio = gasto_total / max(registros, 1)

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Gasto Acumulado Total</div><div class="kpi-value">S/ {gasto_total:,.2f}</div></div>',
            unsafe_allow_html=True,
        )
    with k2:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Centros de Costos</div><div class="kpi-value">{cecos_cant}</div></div>',
            unsafe_allow_html=True,
        )
    with k3:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Transacciones</div><div class="kpi-value">{registros:,}</div></div>',
            unsafe_allow_html=True,
        )
    with k4:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-title">Gasto Promedio por Reg.</div><div class="kpi-value">S/ {promedio:,.2f}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    col1, col2 = st.columns([1.3, 1])
    with col1:
        st.subheader("📈 Tendencia y Evolución Mensual")
        if "Mes" in df_f.columns:
            df_mes = df_f.groupby("Mes")["Saldo"].sum().reset_index()
            fig_m = px.bar(
                df_mes,
                x="Mes",
                y="Saldo",
                text_auto=".2s",
                color_discrete_sequence=[COLOR_ROJO],
            )
            fig_m.add_trace(
                go.Scatter(
                    x=df_mes["Mes"],
                    y=df_mes["Saldo"],
                    mode="lines+markers",
                    name="Tendencia",
                    line=dict(color=COLOR_NAVY, width=3),
                )
            )
            fig_m.update_layout(
                plot_bgcolor="white", height=350, showlegend=False
            )
            st.plotly_chart(fig_m, use_container_width=True)

    with col2:
        st.subheader("🍕 Comportamiento del Gasto")
        df_comp = (
            df_f.groupby("Comportamiento")["Saldo"].sum().reset_index()
        )
        fig_p = px.pie(
            df_comp,
            names="Comportamiento",
            values="Saldo",
            hole=0.45,
            color_discrete_sequence=PALETA_ARMONICA,
        )
        fig_p.update_layout(height=350)
        st.plotly_chart(fig_p, use_container_width=True)

# TAB 2: MAPA DE PROCESOS
with tab2:
    st.markdown(
        '<div class="header-banner"><h1>🗺️ Mapa de Procesos de Negocio</h1><p>Diseño corporativo estructurado por áreas tácticas y operativas.</p></div>',
        unsafe_allow_html=True,
    )

    def calc_val(cadena):
        val = df_f[
            df_f["Subproceso / Área"].str.contains(
                cadena, case=False, na=False
            )
        ]["Saldo"].sum()
        return f"S/ {val:,.2f}" if val > 0 else "S/ 0.00"

    html_mapa = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{ font-family: 'Inter', system-ui, sans-serif; background-color: #f8fafc; margin: 0; padding: 5px; }}
        .process-container {{ background: #ffffff; border-radius: 16px; border: 1px solid #e2e8f0; margin-bottom: 24px; box-shadow: 0 4px 12px rgba(0,0,0,0.03); overflow: hidden; }}
        .process-title {{ background: #1e293b; color: #ffffff; padding: 12px 20px; font-weight: 800; font-size: 13px; letter-spacing: 1px; text-transform: uppercase; }}
        .process-title.red {{ background: linear-gradient(90deg, #a81121 0%, #c2182b 100%); }}
        .process-body {{ padding: 20px; }}
        .grid-2 {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; }}
        .grid-4 {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }}
        .pill-card {{
            background: #ffffff; border: 1px solid #cbd5e1; border-radius: 12px; padding: 14px 16px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.02); transition: all 0.2s ease; border-left: 5px solid #2563eb;
        }}
        .pill-card:hover {{ transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,0.08); border-color: #94a3b8; }}
        .pill-title {{ font-size: 12px; font-weight: 700; color: #475569; }}
        .pill-val {{ font-size: 15px; font-weight: 800; color: #0f172a; margin-top: 6px; }}
        .column-header {{ background: #a81121; color: white; text-align: center; font-weight: 800; padding: 10px; border-radius: 10px; font-size: 12px; margin-bottom: 14px; text-transform: uppercase; letter-spacing: 0.5px; box-shadow: 0 2px 6px rgba(168,17,33,0.15); }}
        .marketing-banner {{ background: #334155; color: white; text-align: center; font-weight: 800; padding: 10px; border-radius: 10px; margin: 20px 0 14px 0; font-size: 12px; letter-spacing: 0.5px; }}
    </style>
    </head>
    <body>

    <div class="process-container">
        <div class="process-title">⚡ PROCESOS ESTRATÉGICOS</div>
        <div class="process-body">
            <div class="grid-2">
                <div class="pill-card" style="border-left-color: #1e293b;">
                    <div class="pill-title">010101 | Gerencia General</div>
                    <div class="pill-val">{calc_val('Gerencia General')}</div>
                </div>
                <div class="pill-card" style="border-left-color: #1e293b;">
                    <div class="pill-title">010102 | Nuevos Negocios</div>
                    <div class="pill-val">{calc_val('Nuevos Negocios')}</div>
                </div>
            </div>
        </div>
    </div>

    <div class="process-container">
        <div class="process-title red">⚙️ PROCESOS OPERATIVOS</div>
        <div class="process-body">
            <div class="grid-4">
                <div>
                    <div class="column-header">LOGÍSTICA</div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030101 Gastos Generales</div><div class="pill-val">{calc_val('Logística')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030102 Compras Locales</div><div class="pill-val">{calc_val('Compras Locales')}</div></div>
                    <div class="pill-card"><div class="pill-title">030103 Importaciones</div><div class="pill-val">{calc_val('Importaciones')}</div></div>
                </div>

                <div>
                    <div class="column-header">ALMACENES</div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030201 Gastos Generales</div><div class="pill-val">{calc_val('Gestión de Almacenes')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030202 Gestión Almacenes</div><div class="pill-val">{calc_val('Gestión de Almacenes')}</div></div>
                    <div class="pill-card"><div class="pill-title">030203 Transporte Interno</div><div class="pill-val">{calc_val('Transporte Interno')}</div></div>
                </div>

                <div>
                    <div class="column-header">VENTAS</div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030301 Gastos Generales</div><div class="pill-val">{calc_val('Ventas Generales')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030302 Ferretería</div><div class="pill-val">{calc_val('Ventas Ferretería')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">030303 Corporativo</div><div class="pill-val">{calc_val('Ventas Corporativo')}</div></div>
                    <div class="pill-card"><div class="pill-title">030304 Licitaciones</div><div class="pill-val">{calc_val('Licitaciones')}</div></div>
                </div>

                <div>
                    <div class="column-header">DISTRIBUCIÓN</div>
                    <div class="pill-card"><div class="pill-title">030401 Gastos Generales</div><div class="pill-val">{calc_val('Distribución')}</div></div>
                </div>
            </div>

            <div class="marketing-banner">MARKETING</div>
            <div style="max-width: 360px; margin: 0 auto;">
                <div class="pill-card" style="text-align:center;"><div class="pill-title">30501 Gastos Generales Marketing</div><div class="pill-val">{calc_val('Marketing')}</div></div>
            </div>
        </div>
    </div>

    <div class="process-container">
        <div class="process-title">🛠️ PROCESOS DE SOPORTE</div>
        <div class="process-body">
            <div class="grid-4" style="margin-bottom: 12px;">
                <div class="pill-card"><div class="pill-title">040101 Soporte General</div><div class="pill-val">{calc_val('Soporte General')}</div></div>
                <div class="pill-card"><div class="pill-title">040103 Tesorería</div><div class="pill-val">{calc_val('Tesorería')}</div></div>
                <div class="pill-card"><div class="pill-title">040105 Administración</div><div class="pill-val">{calc_val('Administración')}</div></div>
                <div class="pill-card"><div class="pill-title">040107 TI / Sistemas</div><div class="pill-val">{calc_val('TI & Sistemas')}</div></div>
            </div>
            <div class="grid-4">
                <div class="pill-card"><div class="pill-title">040102 Recursos Humanos</div><div class="pill-val">{calc_val('Recursos Humanos')}</div></div>
                <div class="pill-card"><div class="pill-title">040104 Contabilidad</div><div class="pill-val">{calc_val('Contabilidad')}</div></div>
                <div class="pill-card"><div class="pill-title">040106 Finanzas/Control</div><div class="pill-val">{calc_val('Control Interno & Finanzas')}</div></div>
                <div class="pill-card"><div class="pill-title">040108 Créditos y Cobranza</div><div class="pill-val">{calc_val('Créditos y Cobranzas')}</div></div>
            </div>
        </div>
    </div>

    </body>
    </html>
    """
    components.html(html_mapa, height=820, scrolling=True)

# TAB 3: ORGANIGRAMA
with tab3:
    st.markdown(
        '<div class="header-banner"><h1>🏛️ Organigrama Estructural por Áreas Orgánicas</h1><p>Jerarquía organizacional con montos acumulados por departamento.</p></div>',
        unsafe_allow_html=True,
    )

    def get_monto_area(*areas):
        val = df_f[
            df_f["Subproceso / Área"]
            .astype(str)
            .str.contains("|".join(areas), case=False, na=False)
        ]["Saldo"].sum()
        return f"S/ {val:,.2f}"

    monto_total = f"S/ {df_f['Saldo'].sum():,.2f}"

    html_org_tree = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        body {{ font-family: 'Inter', system-ui, -apple-system, sans-serif; background-color: #f8fafc; padding: 10px; margin: 0; }}
        .tree {{ width: 100%; display: flex; flex-direction: column; align-items: center; }}
        .root-node {{
            background: linear-gradient(135deg, #d97706 0%, #ea580c 100%);
            color: white; border-radius: 12px; padding: 12px 28px; text-align: center;
            box-shadow: 0 4px 12px rgba(217, 119, 6, 0.3); font-weight: 800; min-width: 220px;
        }}
        .root-node .role {{ font-size: 14px; text-transform: uppercase; letter-spacing: 1px; }}
        .root-node .sub {{ font-size: 11px; opacity: 0.95; margin-top: 2px; font-weight: 600; }}
        .root-node .amount {{ font-size: 13px; font-weight: 800; background: rgba(0,0,0,0.22); border-radius: 6px; padding: 2px 10px; margin-top: 5px; display: inline-block; }}
        .stem {{ width: 2px; height: 24px; background-color: #94a3b8; }}
        .branches {{ display: flex; justify-content: center; width: 100%; position: relative; }}
        .branches::before {{ content: ''; position: absolute; top: 0; left: 12%; right: 12%; height: 2px; background-color: #94a3b8; }}
        .branch-col {{ display: flex; flex-direction: column; align-items: center; flex: 1; padding: 0 6px; position: relative; }}
        .branch-col::before {{ content: ''; position: absolute; top: 0; width: 2px; height: 20px; background-color: #94a3b8; }}
        .area-card {{ margin-top: 20px; background: #1e293b; color: white; border-radius: 10px; padding: 10px 12px; text-align: center; width: 92%; box-shadow: 0 3px 8px rgba(0,0,0,0.12); }}
        .area-card.red-card {{ background: linear-gradient(135deg, #a81121 0%, #800d19 100%); }}
        .area-card .title {{ font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; }}
        .area-card .subtitle {{ font-size: 10px; color: #cbd5e1; margin-top: 2px; }}
        .area-card .monto-head {{ font-size: 12px; font-weight: 800; color: #f8fafc; margin-top: 5px; border-top: 1px solid rgba(255,255,255,0.25); padding-top: 4px; display: block; }}
        .sub-stem {{ width: 2px; height: 16px; background-color: #cbd5e1; }}
        .sub-cards {{ display: flex; flex-direction: column; gap: 8px; width: 92%; margin-top: 2px; }}
        .child-card {{
            background: #ffffff; color: #0f172a; border-radius: 8px; padding: 8px 10px;
            text-align: left; border-left: 4px solid #a81121; border-top: 1px solid #e2e8f0;
            border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0;
            box-shadow: 0 2px 5px rgba(0,0,0,0.04); display: flex; justify-content: space-between; align-items: center;
        }}
        .child-card .dept-name {{ font-size: 11px; font-weight: 700; color: #334155; }}
        .child-card .dept-val {{ font-size: 11px; font-weight: 800; color: #0f172a; background: #f1f5f9; padding: 2px 6px; border-radius: 4px; }}
    </style>
    </head>
    <body>
    <div class="tree">
        <div class="root-node">
            <div class="role">DIRECCIÓN GENERAL</div>
            <div class="sub">Gerencia & Administración</div>
            <div class="amount">{monto_total}</div>
        </div>
        <div class="stem"></div>
        <div class="branches">
            <div class="branch-col">
                <div class="area-card">
                    <div class="title">GERENCIA ESTRATÉGICA</div>
                    <div class="subtitle">Planeación & Expansión</div>
                    <div class="monto-head">{get_monto_area('Gerencia General', 'Nuevos Negocios')}</div>
                </div>
                <div class="sub-stem"></div>
                <div class="sub-cards">
                    <div class="child-card"><span class="dept-name">Gerencia General</span><span class="dept-val">{get_monto_area('Gerencia General')}</span></div>
                    <div class="child-card"><span class="dept-name">Nuevos Negocios</span><span class="dept-val">{get_monto_area('Nuevos Negocios')}</span></div>
                </div>
            </div>

            <div class="branch-col">
                <div class="area-card red-card">
                    <div class="title">GERENCIA OPERACIONES</div>
                    <div class="subtitle">Cadena de Suministro</div>
                    <div class="monto-head">{get_monto_area('Logística', 'Compras Locales', 'Importaciones', 'Gestión de Almacenes', 'Transporte Interno', 'Distribución')}</div>
                </div>
                <div class="sub-stem"></div>
                <div class="sub-cards">
                    <div class="child-card"><span class="dept-name">Logística & Compras</span><span class="dept-val">{get_monto_area('Logística', 'Compras Locales', 'Importaciones')}</span></div>
                    <div class="child-card"><span class="dept-name">Gestión Almacenes</span><span class="dept-val">{get_monto_area('Gestión de Almacenes', 'Transporte Interno')}</span></div>
                    <div class="child-card"><span class="dept-name">Distribución y Flota</span><span class="dept-val">{get_monto_area('Distribución')}</span></div>
                </div>
            </div>

            <div class="branch-col">
                <div class="area-card">
                    <div class="title">GERENCIA COMERCIAL</div>
                    <div class="subtitle">Ventas & Marketing</div>
                    <div class="monto-head">{get_monto_area('Ventas', 'Ferretería', 'Corporativo', 'Licitaciones', 'Marketing')}</div>
                </div>
                <div class="sub-stem"></div>
                <div class="sub-cards">
                    <div class="child-card"><span class="dept-name">Ferretería & Corp.</span><span class="dept-val">{get_monto_area('Ventas Ferretería', 'Ventas Corporativo')}</span></div>
                    <div class="child-card"><span class="dept-name">Licitaciones</span><span class="dept-val">{get_monto_area('Licitaciones')}</span></div>
                    <div class="child-card"><span class="dept-name">Marketing Digital</span><span class="dept-val">{get_monto_area('Marketing')}</span></div>
                </div>
            </div>

            <div class="branch-col">
                <div class="area-card">
                    <div class="title">G. ADMIN. & FINANZAS</div>
                    <div class="subtitle">Soporte Corporativo</div>
                    <div class="monto-head">{get_monto_area('Recursos Humanos', 'Tesorería', 'Contabilidad', 'Control Interno & Finanzas', 'TI & Sistemas', 'Créditos y Cobranzas', 'Administración', 'Soporte General')}</div>
                </div>
                <div class="sub-stem"></div>
                <div class="sub-cards">
                    <div class="child-card"><span class="dept-name">Recursos Humanos</span><span class="dept-val">{get_monto_area('Recursos Humanos')}</span></div>
                    <div class="child-card"><span class="dept-name">Finanzas & Contabilidad</span><span class="dept-val">{get_monto_area('Tesorería', 'Contabilidad', 'Control Interno & Finanzas')}</span></div>
                    <div class="child-card"><span class="dept-name">TI, Sistemas & Créditos</span><span class="dept-val">{get_monto_area('TI & Sistemas', 'Créditos y Cobranzas', 'Administración')}</span></div>
                </div>
            </div>
        </div>
    </div>
    </body>
    </html>
    """
    components.html(html_org_tree, height=450, scrolling=True)

    st.markdown("---")
    st.markdown("### 🖱️ Explorador Interactivo de Centros de Costo por Área")
    col_sel1, col_sel2 = st.columns(2)
    with col_sel1:
        macro_sel = st.selectbox(
            "1. Nivel Superior / Rama:",
            sorted(list(df_f["Macroproceso"].unique())),
        )
    df_sub_opt = df_f[df_f["Macroproceso"] == macro_sel]
    with col_sel2:
        area_sel = st.selectbox(
            "2. Área del Organigrama:",
            sorted(list(df_sub_opt["Subproceso / Área"].unique())),
        )

    df_cecos = (
        df_f[df_f["Subproceso / Área"] == area_sel]
        .groupby("Centro de Costos")["Saldo"]
        .agg(["sum", "count"])
        .reset_index()
    )
    df_cecos.columns = [
        "Centro de Costos",
        "Monto Presupuestado (S/)",
        "Cant. Registros",
    ]
    df_cecos = df_cecos.sort_values(
        by="Monto Presupuestado (S/)", ascending=False
    )

    cg1, cg2 = st.columns([1.3, 1])
    with cg1:
        fig_bar_ceco = px.bar(
            df_cecos,
            x="Monto Presupuestado (S/)",
            y="Centro de Costos",
            orientation="h",
            text_auto=".2s",
            color_discrete_sequence=[COLOR_ROJO],
            title=f"CECOs en: {area_sel}",
        )
        fig_bar_ceco.update_layout(
            height=300, plot_bgcolor="white", yaxis=dict(autorange="reversed")
        )
        st.plotly_chart(fig_bar_ceco, use_container_width=True)
    with cg2:
        st.dataframe(
            df_cecos.style.format(
                {"Monto Presupuestado (S/)": "S/ {:,.2f}"}
            ),
            use_container_width=True,
            height=280,
            hide_index=True,
        )

# TAB 4: SUCURSALES
with tab4:
    st.markdown(
        '<div class="header-banner"><h1>🗺️ Mapa de Perú y Distribución por Sucursales</h1><p>Geolocalización exacta de las sedes con montos ejecutados.</p></div>',
        unsafe_allow_html=True,
    )

    col_m1, col_m2 = st.columns([1.3, 1])

    df_geo_suc = (
        df_f.groupby(["Sucursal", "Ciudad_Nombre", "Latitud", "Longitud"])[
            "Saldo"
        ]
        .sum()
        .reset_index()
    )

    df_geo_suc["size_mapa"] = df_geo_suc["Saldo"].abs()

    with col_m1:
        st.subheader("📍 Geolocalización de Sucursales en Perú")

        fig_map_peru = px.scatter_geo(
            df_geo_suc,
            lat="Latitud",
            lon="Longitud",
            size="size_mapa",
            hover_name="Sucursal",
            text="Sucursal",
            color="Saldo",
            color_continuous_scale="Reds",
            projection="mercator",
            hover_data={
                "size_mapa": False,
                "Saldo": ":,.2f",
            },
            title="Sucursales Activas en el Mapa de Perú",
        )

        fig_map_peru.update_geos(
            center=dict(lat=-9.1899, lon=-75.0152),
            projection_scale=4.5,
            showland=True,
            landcolor="#f1f5f9",
            showocean=True,
            oceancolor="#e2e8f0",
            showcountries=True,
            countrycolor="#cbd5e1",
        )
        fig_map_peru.update_traces(
            textposition="top center", marker=dict(sizemin=10)
        )
        fig_map_peru.update_layout(
            height=480, margin={"r": 0, "t": 40, "l": 0, "b": 0}
        )
        st.plotly_chart(fig_map_peru, use_container_width=True)

    with col_m2:
        st.subheader("🏢 Gasto por Sucursal / Sede")
        df_suc = (
            df_f.groupby("Sucursal")["Saldo"]
            .sum()
            .reset_index()
            .sort_values(by="Saldo", ascending=True)
        )
        fig_suc = px.bar(
            df_suc,
            y="Sucursal",
            x="Saldo",
            orientation="h",
            text_auto=".2s",
            color_discrete_sequence=[COLOR_AZUL],
        )
        fig_suc.update_layout(height=480, plot_bgcolor="white")
        st.plotly_chart(fig_suc, use_container_width=True)

# TAB 5: DETALLE CECO Y ASIENTOS
with tab5:
    st.markdown(
        '<div class="header-banner"><h1>🔍 Detalle por Centro de Costos y Buscador de Asientos</h1></div>',
        unsafe_allow_html=True,
    )
    ceco_sel_det = st.selectbox(
        "Seleccionar Centro de Costos Específico:",
        options=["Todos"] + sorted(list(df_f["Centro de Costos"].unique())),
    )
    df_det = (
        df_f
        if ceco_sel_det == "Todos"
        else df_f[df_f["Centro de Costos"] == ceco_sel_det]
    )

    q = st.text_input(
        "🔎 Buscar Palabra Clave en Glosa / Asiento:",
        placeholder="Ej. Mantenimiento, 100234...",
    )
    if q:
        df_det = df_det[
            df_det["Glosa / Comentario"].str.contains(q, case=False, na=False)
        ]

    st.info(
        f"Registros filtrados: {len(df_det):,} | Monto Total: S/ {df_det['Saldo'].sum():,.2f}"
    )
    cols_ver = [
        c
        for c in [
            "Mes",
            "Asiento",
            "Centro de Costos",
            "Cuenta Contable",
            "Glosa / Comentario",
            "Saldo",
        ]
        if c in df_det.columns
    ]
    st.dataframe(
        df_det[cols_ver].style.format({"Saldo": "S/ {:,.2f}"}),
        use_container_width=True,
        height=450,
        hide_index=True,
    )

# TAB 6: ALERTAS Y MOTIVOS
with tab6:
    st.markdown(
        '<div class="header-banner"><h1>🚨 Centro de Alertas y Agrupación por Motivo</h1>'
        '<p>Clasificación inteligente de Reversiones, Extornos, Sin CECO y Ajustes.</p></div>',
        unsafe_allow_html=True
    )

    def clasificar_motivo_completo(row):
        glosa = str(row.get("Glosa / Comentario", "")).upper()
        ceco = str(row.get("Centro de Costos", "")).upper()
        saldo = row.get("Saldo", 0)
        
        kw_rev = ["REV", "REVERSION", "EXTORNO", "ANULAC", "ANULACION", "CONTRASIENTO", "DEVOLUCION"]
        kw_prov = ["PROV", "PROVISION", "ESTIMAC"]
        kw_recl = ["CORREC", "RECLASIF", "AJUSTE", "REGULARIZAC"]
        
        if saldo < 0 or any(k in glosa for k in kw_rev):
            return "1. Reversión / Extorno / Anulación"
        elif ceco == "SIN CECO":
            return "2. Pendiente Asignación CECO"
        elif any(k in glosa for k in kw_prov):
            return "3. Ajuste de Provisión"
        elif any(k in glosa for k in kw_recl):
            return "4. Reclasificación Contable"
        elif any(k in ceco for k in ["GASTOS FINANCIEROS", "FINANCIEROS", "ELUC"]):
            return "5. Gastos Financieros / Corporativos"
        else:
            return "6. Operación Normal"

    df_f['Motivo_Alerta'] = df_f.apply(clasificar_motivo_completo, axis=1)

    df_alertas_resumen = df_f.groupby('Motivo_Alerta')['Saldo'].agg(['sum', 'count']).reset_index()
    df_alertas_resumen.columns = ['Clasificación de Motivo', 'Monto Total (S/)', 'Cantidad Asientos']

    col_g1, col_g2 = st.columns(2)

    with col_g1:
        fig_alertas_monto = px.bar(
            df_alertas_resumen, 
            x="Clasificación de Motivo", 
            y="Monto Total (S/)", 
            color="Clasificación de Motivo", 
            text_auto='.2s',
            color_discrete_sequence=PALETA_ARMONICA,
            title="💰 Distribución de Montos (S/)"
        )
        fig_alertas_monto.update_layout(height=400, plot_bgcolor="white", showlegend=False)
        fig_alertas_monto.update_xaxes(tickangle=-30)
        st.plotly_chart(fig_alertas_monto, use_container_width=True)

    with col_g2:
        fig_alertas_cant = px.bar(
            df_alertas_resumen, 
            x="Clasificación de Motivo", 
            y="Cantidad Asientos", 
            color="Clasificación de Motivo", 
            text_auto=True,
            color_discrete_sequence=PALETA_ARMONICA,
            title="📊 Cantidad de Registros / Asientos"
        )
        fig_alertas_cant.update_layout(height=400, plot_bgcolor="white", showlegend=False)
        fig_alertas_cant.update_xaxes(tickangle=-30)
        st.plotly_chart(fig_alertas_cant, use_container_width=True)

    st.markdown("---")
    st.subheader("📋 Detalle de Transacciones por Alerta")
    motivo_sel = st.selectbox("Filtrar por tipo de Alerta:", options=sorted(list(df_f['Motivo_Alerta'].unique())))
    df_motivo_det = df_f[df_f['Motivo_Alerta'] == motivo_sel]

    cols_ver_al = [c for c in ['Mes', 'Asiento', 'Centro de Costos', 'Cuenta Contable', 'Glosa / Comentario', 'Saldo'] if c in df_motivo_det.columns]
    st.dataframe(df_motivo_det[cols_ver_al].style.format({'Saldo': 'S/ {:,.2f}'}), use_container_width=True, height=350, hide_index=True)
