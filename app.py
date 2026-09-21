import datetime
import io
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

    # Nombres de archivo posibles (con espacio o con guion bajo)
    posibles_nombres = ["Reporte CECO.xlsx", "Reporte_CECO.xlsx"]
    posibles_rutas = [BASE_DIR / n for n in posibles_nombres]
    posibles_rutas += [
        Path(r"C:\Users\JUNIOR FP&A\Desktop\Finanzas\POWER BI") / n
        for n in posibles_nombres
    ]
    posibles_rutas += [Path(r"C:\Users\JUNIOR FP&A") / n for n in posibles_nombres]

    ruta_excel = None
    for r in posibles_rutas:
        if r.exists():
            ruta_excel = r
            break

    if ruta_excel is None:
        st.error(
            "❌ No se encontró el archivo Excel. Verifica que 'Reporte CECO.xlsx' "
            f"esté en: {BASE_DIR}"
        )
        st.stop()

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
        elif "AREQ" in s or "SUR" in s:
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


def hex_a_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


RGB_ROJO = hex_a_rgb(COLOR_ROJO)
RGB_NAVY = hex_a_rgb(COLOR_NAVY)
RGB_AZUL = hex_a_rgb(COLOR_AZUL)

# --- FUNCIONES COMPARTIDAS (usadas por la pestaña de Alertas y por el PDF) ---

def clasificar_motivo_vectorizado(df_input):
    """Versión vectorizada (rápida) de la clasificación por motivo.
    Evita .apply(axis=1), que en tablas de 25,000+ filas puede tardar
    varios segundos o minutos según el equipo."""
    glosa = df_input["Glosa / Comentario"].astype(str).str.upper()
    ceco = df_input["Centro de Costos"].astype(str).str.upper()
    saldo = df_input["Saldo"]

    kw_rev = "REV|REVERSION|EXTORNO|ANULAC|ANULACION|CONTRASIENTO|DEVOLUCION"
    kw_prov = "PROV|PROVISION|ESTIMAC"
    kw_recl = "CORREC|RECLASIF|AJUSTE|REGULARIZAC"
    kw_fin = "GASTOS FINANCIEROS|FINANCIEROS|ELUC"

    condiciones = [
        (saldo < 0) | glosa.str.contains(kw_rev, regex=True, na=False),
        ceco.eq("SIN CECO"),
        glosa.str.contains(kw_prov, regex=True, na=False),
        glosa.str.contains(kw_recl, regex=True, na=False),
        ceco.str.contains(kw_fin, regex=True, na=False),
    ]
    resultados = [
        "1. Reversión / Extorno / Anulación",
        "2. Pendiente Asignación CECO",
        "3. Ajuste de Provisión",
        "4. Reclasificación Contable",
        "5. Gastos Financieros / Corporativos",
    ]
    return np.select(condiciones, resultados, default="6. Operación Normal")


KEYWORDS_RESIDUAL = [
    "REDONDEO", "AJUSTE", "VARIOS", "DIFERENCIA", "DIVERSOS",
    "PENDIENTE", "SUSCRIPCIONES Y COTIZAC", "RECONOCIMIENTO FALTANTES",
]


def detectar_alertas_ia(df_input, umbral_residual=1000, umbral_z=2.0):
    df_cuentas_mes = (
        df_input.groupby(["Cuenta Contable", "Centro de Costos", "Mes"])["Saldo"]
        .sum()
        .reset_index()
    )
    stats_cuenta = (
        df_cuentas_mes.groupby("Cuenta Contable")["Saldo"]
        .agg(["mean", "std", "count"])
        .reset_index()
    )
    stats_cuenta.columns = ["Cuenta Contable", "Promedio_Historico", "Desviacion_Std", "Cant_Registros"]
    stats_cuenta["Desviacion_Std"] = stats_cuenta["Desviacion_Std"].fillna(0)

    df_analisis = df_cuentas_mes.merge(stats_cuenta, on="Cuenta Contable", how="left")
    df_analisis["Z_Score"] = np.where(
        df_analisis["Desviacion_Std"] > 0,
        (df_analisis["Saldo"] - df_analisis["Promedio_Historico"]) / df_analisis["Desviacion_Std"],
        0.0,
    )
    df_analisis["Es_Cuenta_Residual"] = df_analisis["Cuenta Contable"].str.upper().str.contains(
        "|".join(KEYWORDS_RESIDUAL), regex=True, na=False
    )

    def evaluar_alerta(row):
        motivos = []
        if row["Es_Cuenta_Residual"] and abs(row["Saldo"]) > umbral_residual:
            motivos.append(
                f"Cuenta de naturaleza residual (ajuste/redondeo/partidas varias) con monto "
                f"atipicamente alto: S/ {row['Saldo']:,.2f}. Por logica contable deberia ser marginal."
            )
        if row["Cant_Registros"] >= 3 and abs(row["Z_Score"]) > umbral_z:
            motivos.append(
                f"Desviacion estadistica: este mes se aleja {row['Z_Score']:.1f} desviaciones "
                f"estandar del promedio historico de la cuenta (S/ {row['Promedio_Historico']:,.2f})."
            )
        return " | ".join(motivos) if motivos else None

    df_analisis["Motivo_Alerta_IA"] = df_analisis.apply(evaluar_alerta, axis=1)
    df_alertas = df_analisis[df_analisis["Motivo_Alerta_IA"].notna()].copy()
    return df_alertas.sort_values("Saldo", key=lambda s: s.abs(), ascending=False)


def limpiar_txt(texto):
    """Sanitiza texto para el PDF (fuentes core de fpdf2 solo soportan latin-1, sin emojis)."""
    return str(texto).encode("latin-1", "ignore").decode("latin-1")


def generar_reporte_pdf(df_f, anios_sel, meses_sel, suc_sel):
    from fpdf import FPDF

    ANCHO, ALTO = 297, 210  # A4 apaisado ("horizontal", tipo pantalla)
    MARGEN = 10
    ANCHO_UTIL = ANCHO - 2 * MARGEN

    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=14)
    pdf.set_margins(MARGEN, MARGEN, MARGEN)

    # ---------- Helpers internos de estilo ----------
    def banner(titulo, subtitulo=""):
        pdf.set_fill_color(*RGB_NAVY)
        pdf.rect(0, 0, ANCHO, 22, style="F")
        pdf.set_fill_color(*RGB_ROJO)
        pdf.rect(0, 0, 2.2, 22, style="F")
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("helvetica", "B", 15)
        pdf.set_xy(MARGEN, 5)
        pdf.cell(0, 8, limpiar_txt(titulo))
        if subtitulo:
            pdf.set_font("helvetica", "", 9)
            pdf.set_xy(MARGEN, 13)
            pdf.set_text_color(180, 190, 205)
            pdf.cell(0, 5, limpiar_txt(subtitulo))
        pdf.set_text_color(0, 0, 0)
        pdf.set_y(28)

    def titulo_seccion(texto):
        pdf.set_fill_color(*RGB_ROJO)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("helvetica", "B", 11)
        pdf.set_x(MARGEN)
        pdf.cell(ANCHO_UTIL, 7, limpiar_txt(texto), fill=True, ln=True)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(1)

    def kpi_box(x, y, w, h, titulo_kpi, valor_kpi):
        pdf.set_draw_color(226, 232, 240)
        pdf.rect(x, y, w, h, style="D")
        pdf.set_fill_color(*RGB_ROJO)
        pdf.rect(x, y, w, 1.3, style="F")
        pdf.set_xy(x + 3, y + 4)
        pdf.set_font("helvetica", "B", 7.5)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(w - 6, 4, limpiar_txt(titulo_kpi.upper()))
        pdf.set_xy(x + 3, y + 10)
        pdf.set_font("helvetica", "B", 13)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(w - 6, 6, limpiar_txt(valor_kpi))
        pdf.set_text_color(0, 0, 0)

    def tabla(headers, rows, col_widths, x=None):
        x0 = MARGEN if x is None else x
        pdf.set_x(x0)
        pdf.set_font("helvetica", "B", 8)
        pdf.set_fill_color(226, 232, 240)
        for h, w in zip(headers, col_widths):
            pdf.cell(w, 6, limpiar_txt(h), border=1, align="C", fill=True)
        pdf.ln()
        pdf.set_font("helvetica", "", 8)
        for row in rows:
            pdf.set_x(x0)
            for val, w in zip(row, col_widths):
                align = "R" if isinstance(val, (int, float)) else "L"
                texto = f"{val:,.2f}" if isinstance(val, (int, float)) else str(val)
                pdf.cell(w, 5.5, limpiar_txt(texto), border=1, align=align)
            pdf.ln()
        pdf.ln(2)

    def fig_a_png(fig, w_mm, h_mm):
        """Renderiza una figura Plotly a PNG (bytes) via kaleido, con el mismo
        estilo de la app, lista para insertarse en el PDF."""
        fig2 = go.Figure(fig)
        fig2.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(l=35, r=20, t=45, b=35),
        )
        return fig2.to_image(format="png", width=int(w_mm * 7.5), height=int(h_mm * 7.5), scale=1)

    def insertar_grafico(fig, x, y, w, h):
        try:
            img = fig_a_png(fig, w, h)
            pdf.image(io.BytesIO(img), x=x, y=y, w=w, h=h)
        except Exception:
            pdf.set_xy(x, y + h / 2 - 5)
            pdf.set_font("helvetica", "I", 8)
            pdf.set_text_color(150, 150, 150)
            pdf.multi_cell(w, 5, limpiar_txt(
                "Grafico no disponible en este equipo. Instala el motor de "
                "renderizado con:  pip install -U kaleido"
            ), align="C")
            pdf.set_text_color(0, 0, 0)

    # ================= 1. PORTADA + RESUMEN GENERAL =================
    pdf.add_page()
    banner(
        "Reporte Ejecutivo - Control Presupuestal FP&A",
        f"Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Junior FP&A",
    )

    pdf.set_font("helvetica", "", 8.5)
    filtro_txt = (
        f"Filtros aplicados  ->  Anios: {anios_sel if anios_sel else 'Todos'}   |   "
        f"Meses: {meses_sel if meses_sel else 'Todos'}   |   "
        f"Sucursales: {len(suc_sel)} seleccionadas"
    )
    pdf.set_x(MARGEN)
    pdf.cell(0, 5, limpiar_txt(filtro_txt), ln=True)
    pdf.ln(2)

    gasto_total = df_f["Saldo"].sum()
    cecos_cant = df_f["Centro de Costos"].nunique()
    registros = len(df_f)
    promedio = gasto_total / max(registros, 1)

    y_kpi = pdf.get_y()
    ancho_kpi = (ANCHO_UTIL - 3 * 4) / 4
    kpis = [
        ("Gasto Acumulado Total", f"S/ {gasto_total:,.2f}"),
        ("Centros de Costos", f"{cecos_cant}"),
        ("Transacciones", f"{registros:,}"),
        ("Gasto Promedio / Registro", f"S/ {promedio:,.2f}"),
    ]
    for i, (t, v) in enumerate(kpis):
        kpi_box(MARGEN + i * (ancho_kpi + 4), y_kpi, ancho_kpi, 18, t, v)

    y_charts = y_kpi + 24
    w_izq = ANCHO_UTIL * 0.58
    w_der = ANCHO_UTIL * 0.38
    h_charts = 80

    if "Mes" in df_f.columns:
        df_mes_pdf = df_f.groupby("Mes")["Saldo"].sum().reset_index().sort_values("Mes")
        fig_m = px.bar(
            df_mes_pdf, x="Mes", y="Saldo", text_auto=".2s",
            color_discrete_sequence=[COLOR_ROJO], title="Tendencia y Evolucion Mensual",
        )
        fig_m.add_trace(go.Scatter(
            x=df_mes_pdf["Mes"], y=df_mes_pdf["Saldo"], mode="lines+markers",
            line=dict(color=COLOR_NAVY, width=3),
        ))
        fig_m.update_layout(showlegend=False)
        insertar_grafico(fig_m, MARGEN, y_charts, w_izq, h_charts)

    df_comp_pdf = df_f.groupby("Comportamiento")["Saldo"].sum().reset_index()
    fig_p = px.pie(
        df_comp_pdf, names="Comportamiento", values="Saldo", hole=0.45,
        color_discrete_sequence=PALETA_ARMONICA, title="Comportamiento del Gasto",
    )
    insertar_grafico(fig_p, MARGEN + w_izq + 4, y_charts, w_der, h_charts)

    # ================= 2. MAPA DE PROCESOS (treemap jerarquico) =================
    pdf.add_page()
    banner("Mapa de Procesos de Negocio", "Distribucion jerarquica del gasto por Macroproceso, Area y Centro de Costo")
    df_treemap = (
        df_f.assign(Saldo_Abs=df_f["Saldo"].abs())
        .groupby(["Macroproceso", "Subproceso / Área", "Centro de Costos"])["Saldo_Abs"]
        .sum().reset_index()
    )
    fig_tree = px.treemap(
        df_treemap, path=["Macroproceso", "Subproceso / Área", "Centro de Costos"],
        values="Saldo_Abs", color="Macroproceso", color_discrete_sequence=PALETA_ARMONICA,
    )
    fig_tree.update_traces(textinfo="label+value")
    insertar_grafico(fig_tree, MARGEN, pdf.get_y(), ANCHO_UTIL, 155)

    # ================= 3. ORGANIGRAMA POR AREAS =================
    pdf.add_page()
    banner("Organigrama por Areas", "Monto ejecutado por Macroproceso y por Area / Subproceso")
    y0 = pdf.get_y()
    df_macro_pdf = df_f.groupby("Macroproceso")["Saldo"].sum().reset_index().sort_values("Saldo")
    fig_macro = px.bar(
        df_macro_pdf, x="Saldo", y="Macroproceso", orientation="h",
        text_auto=".2s", color_discrete_sequence=[COLOR_NAVY], title="Monto por Macroproceso",
    )
    insertar_grafico(fig_macro, MARGEN, y0, ANCHO_UTIL * 0.48, 80)

    df_area_pdf = (
        df_f.groupby("Subproceso / Área")["Saldo"].sum().reset_index()
        .sort_values("Saldo", key=lambda s: s.abs(), ascending=False).head(10)
        .sort_values("Saldo")
    )
    fig_area = px.bar(
        df_area_pdf, x="Saldo", y="Subproceso / Área", orientation="h",
        text_auto=".2s", color_discrete_sequence=[COLOR_ROJO], title="Top 10 Areas / Subprocesos",
    )
    insertar_grafico(fig_area, MARGEN + ANCHO_UTIL * 0.52, y0, ANCHO_UTIL * 0.48, 80)

    # ================= 4. MAPA SUCURSALES PERU =================
    pdf.add_page()
    banner("Mapa de Peru y Distribucion por Sucursales", "Geolocalizacion de sedes con montos ejecutados")
    y0 = pdf.get_y()
    df_geo_suc = (
        df_f.groupby(["Sucursal", "Ciudad_Nombre", "Latitud", "Longitud"])["Saldo"].sum().reset_index()
    )
    df_geo_suc["size_mapa"] = df_geo_suc["Saldo"].abs()
    fig_map = px.scatter_geo(
        df_geo_suc, lat="Latitud", lon="Longitud", size="size_mapa", hover_name="Sucursal",
        text="Sucursal", color="Saldo", color_continuous_scale="Reds", projection="mercator",
    )
    fig_map.update_geos(
        center=dict(lat=-9.1899, lon=-75.0152), projection_scale=4.5,
        showland=True, landcolor="#f1f5f9", showocean=True, oceancolor="#e2e8f0",
        showcountries=True, countrycolor="#cbd5e1",
    )
    fig_map.update_traces(textposition="top center", marker=dict(sizemin=8))
    insertar_grafico(fig_map, MARGEN, y0, ANCHO_UTIL * 0.58, 150)

    df_suc_bar = df_f.groupby("Sucursal")["Saldo"].sum().reset_index().sort_values("Saldo")
    fig_suc = px.bar(
        df_suc_bar, x="Saldo", y="Sucursal", orientation="h", text_auto=".2s",
        color_discrete_sequence=[COLOR_AZUL], title="Gasto por Sucursal",
    )
    insertar_grafico(fig_suc, MARGEN + ANCHO_UTIL * 0.60, y0, ANCHO_UTIL * 0.40, 150)

    # ================= 5. TOP 15 ASIENTOS =================
    pdf.add_page()
    banner("Top 15 Asientos por Monto Absoluto")
    titulo_seccion("Detalle de las transacciones de mayor impacto")
    df_top_asientos = df_f.reindex(df_f["Saldo"].abs().sort_values(ascending=False).index).head(15)
    cols_disp = [
        (
            str(r.get("Mes", "")),
            str(r.get("Centro de Costos", ""))[:28],
            str(r.get("Glosa / Comentario", ""))[:45],
            r.get("Saldo", 0),
        )
        for _, r in df_top_asientos.iterrows()
    ]
    tabla(["Mes", "Centro de Costo", "Glosa", "Monto (S/)"], cols_disp, [18, 70, 110, 40])

    # ================= 6. CENTRO DE ALERTAS =================
    pdf.add_page()
    banner("Centro de Alertas", "Por motivo del asiento y por comportamiento estadistico de la cuenta (IA)")
    titulo_seccion("Alertas por Motivo del Asiento")

    df_f_pdf = df_f.copy()
    df_f_pdf["Motivo_Alerta"] = clasificar_motivo_vectorizado(df_f_pdf)
    df_motivo_pdf = df_f_pdf.groupby("Motivo_Alerta")["Saldo"].agg(["sum", "count"]).reset_index()
    df_motivo_pdf.columns = ["Motivo", "Monto", "Cantidad"]

    y0 = pdf.get_y()
    fig_motivo_monto = px.bar(
        df_motivo_pdf, x="Motivo", y="Monto", color="Motivo",
        text_auto=".2s", color_discrete_sequence=PALETA_ARMONICA, title="Monto por Motivo (S/)",
    )
    fig_motivo_monto.update_layout(showlegend=False, xaxis_tickangle=-20)
    insertar_grafico(fig_motivo_monto, MARGEN, y0, ANCHO_UTIL * 0.48, 70)

    fig_motivo_cant = px.bar(
        df_motivo_pdf, x="Motivo", y="Cantidad", color="Motivo",
        text_auto=True, color_discrete_sequence=PALETA_ARMONICA, title="Cantidad de Asientos",
    )
    fig_motivo_cant.update_layout(showlegend=False, xaxis_tickangle=-20)
    insertar_grafico(fig_motivo_cant, MARGEN + ANCHO_UTIL * 0.52, y0, ANCHO_UTIL * 0.48, 70)

    # --- Alertas IA (pagina nueva, incluye tabla de detalle) ---
    pdf.add_page()
    banner("Centro de Alertas (continuacion)", "Cuentas con comportamiento atipico (motor IA, umbrales por defecto)")
    titulo_seccion("Alertas IA - Cuentas Atipicas")
    df_alertas_ia_pdf = detectar_alertas_ia(df_f, umbral_residual=1000, umbral_z=2.0)

    if len(df_alertas_ia_pdf) > 0:
        y0 = pdf.get_y()
        top_ia = df_alertas_ia_pdf.head(12).copy()
        top_ia["Etiqueta"] = top_ia["Cuenta Contable"].str[:35] + " (Mes " + top_ia["Mes"].astype(str) + ")"
        fig_ia = px.bar(
            top_ia.sort_values("Saldo"), x="Saldo", y="Etiqueta", orientation="h",
            text_auto=".2s", color_discrete_sequence=[COLOR_ROJO], title="Top Cuentas Atipicas",
        )
        insertar_grafico(fig_ia, MARGEN, y0, ANCHO_UTIL * 0.55, 90)

        resumen_cuenta = (
            df_alertas_ia_pdf.groupby("Cuenta Contable")["Saldo"].sum().reset_index()
            .sort_values("Saldo", key=lambda s: s.abs(), ascending=False).head(8)
        )
        fig_resumen = px.pie(
            resumen_cuenta, names="Cuenta Contable", values="Saldo", hole=0.45,
            color_discrete_sequence=PALETA_ARMONICA, title="Distribucion por Cuenta",
        )
        insertar_grafico(fig_resumen, MARGEN + ANCHO_UTIL * 0.58, y0, ANCHO_UTIL * 0.42, 90)
        pdf.set_y(y0 + 94)

        titulo_seccion("Detalle preciso (Top 15)")
        filas_ia = [
            (str(row["Mes"]), str(row["Cuenta Contable"])[:40], row["Saldo"], row["Z_Score"])
            for _, row in df_alertas_ia_pdf.head(15).iterrows()
        ]
        tabla(["Mes", "Cuenta Contable", "Monto (S/)", "Z-Score"], filas_ia, [20, 140, 45, 30])
    else:
        pdf.set_font("helvetica", "", 9)
        pdf.cell(0, 6, limpiar_txt("No se detectaron cuentas atipicas con los umbrales por defecto."), ln=True)

    return bytes(pdf.output())

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

    generar_pdf_click = st.button("📄 Generar Reporte PDF Ejecutivo", type="primary")
    st.caption(
        "✅ Recomendado: 1 clic, incluye las 6 pestañas con gráficos, KPIs y tablas "
        "ya maquetados — no necesitas ajustar escala ni imprimir pestaña por pestaña."
    )

    with st.expander("🖨️ Alternativa: imprimir pantalla actual"):
        st.caption(
            "Imprime solo la pestaña visible en este momento (limitación del navegador). "
            "Úsalo solo si necesitas una captura puntual de una vista específica."
        )
        if st.button("Imprimir pestaña actual"):
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

# Generación del PDF ejecutivo (se hace aquí porque necesita df_f ya filtrado)
if generar_pdf_click:
    _t0 = datetime.datetime.now()
    try:
        with st.spinner(f"Generando PDF ejecutivo ({len(df_f):,} registros a procesar)..."):
            st.session_state["pdf_bytes"] = generar_reporte_pdf(df_f, anios_sel, meses_sel, suc_sel)
        _segundos = (datetime.datetime.now() - _t0).total_seconds()
        with st.sidebar:
            st.success(f"✅ PDF generado en {_segundos:.1f} segundos. Descárgalo abajo 👇")
    except ModuleNotFoundError as e:
        with st.sidebar:
            st.error(
                f"❌ Falta una librería para generar el PDF: {e.name}. "
                f"Instálala en tu entorno con:\n\n"
                f"pip install -U fpdf2 kaleido"
            )

if "pdf_bytes" in st.session_state:
    with st.sidebar:
        st.download_button(
            label="⬇️ Descargar Reporte PDF",
            data=st.session_state["pdf_bytes"],
            file_name=f"Reporte_Ejecutivo_FPA_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
        )

# NAVEGACIÓN
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📊 Resumen General",
    "🗺️ Mapa de Procesos",
    "🏛️ Organigrama por Áreas",
    "🗺️ Mapa Sucursales Perú",
    "🔍 Detalle por CECO / Asientos",
    "🚨 Centro de Alertas",
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

    comp_sel_tab2 = st.radio(
        "🎚️ Filtrar por Comportamiento del Gasto:",
        options=["Todos", "Fijos", "Variables"],
        horizontal=True,
        key="comp_filter_tab2",
    )

    df_tab2 = (
        df_f
        if comp_sel_tab2 == "Todos"
        else df_f[df_f["Comportamiento"] == comp_sel_tab2]
    )

    # IMPORTANTE: coincidencia EXACTA por nombre real de Centro de Costos
    # (antes se agrupaba por "Subproceso / Área", lo que hacía que dos CECOs
    # distintos como "Almacen-Gastos Generales" y "Almacen-Gestión de Almacenes"
    # cayeran en la misma categoría y mostraran el mismo monto duplicado).
    def calc_val(nombre_ceco_exacto):
        val = df_tab2[df_tab2["Centro de Costos"] == nombre_ceco_exacto][
            "Saldo"
        ].sum()
        return f"S/ {val:,.2f}"

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
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Logistica-Gastos Generales</div><div class="pill-val">{calc_val('Logistica-Gastos Generales')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Logistica-Compras Locales</div><div class="pill-val">{calc_val('Logistica-Compras Locales')}</div></div>
                    <div class="pill-card"><div class="pill-title">Logistica-Importaciones</div><div class="pill-val">{calc_val('Logistica-Importaciones')}</div></div>
                </div>

                <div>
                    <div class="column-header">ALMACENES</div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Almacen-Gastos Generales</div><div class="pill-val">{calc_val('Almacen-Gastos Generales')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Almacen-Gestión de Almacenes</div><div class="pill-val">{calc_val('Almacen-Gestión de Almacenes')}</div></div>
                    <div class="pill-card"><div class="pill-title">Almacen-Transporte Interno</div><div class="pill-val">{calc_val('Almacen-Transporte Interno')}</div></div>
                </div>

                <div>
                    <div class="column-header">VENTAS</div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Ventas-Gastos Generales</div><div class="pill-val">{calc_val('Ventas-Gastos Generales')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Ventas-Ferreteria</div><div class="pill-val">{calc_val('Ventas-Ferreteria')}</div></div>
                    <div class="pill-card" style="margin-bottom:12px;"><div class="pill-title">Ventas-Corporativo</div><div class="pill-val">{calc_val('Ventas-Corporativo')}</div></div>
                    <div class="pill-card"><div class="pill-title">Ventas-Licitaciones</div><div class="pill-val">{calc_val('Ventas-Licitaciones')}</div></div>
                </div>

                <div>
                    <div class="column-header">DISTRIBUCIÓN</div>
                    <div class="pill-card"><div class="pill-title">Distribución-Gastos Generales</div><div class="pill-val">{calc_val('Distribución-Gastos Generales')}</div></div>
                </div>
            </div>

            <div class="marketing-banner">MARKETING</div>
            <div style="max-width: 360px; margin: 0 auto;">
                <div class="pill-card" style="text-align:center;"><div class="pill-title">Marketing-Gastos Generales</div><div class="pill-val">{calc_val('Marketing-Gastos Generales')}</div></div>
            </div>
        </div>
    </div>

    <div class="process-container">
        <div class="process-title">🛠️ PROCESOS DE SOPORTE</div>
        <div class="process-body">
            <div class="grid-4" style="margin-bottom: 12px;">
                <div class="pill-card"><div class="pill-title">Soporte-Gastos Generales</div><div class="pill-val">{calc_val('Soporte-Gastos Generales')}</div></div>
                <div class="pill-card"><div class="pill-title">Soporte-Tesoreria</div><div class="pill-val">{calc_val('Soporte-Tesoreria')}</div></div>
                <div class="pill-card"><div class="pill-title">Soporte-Administración</div><div class="pill-val">{calc_val('Soporte-Administración')}</div></div>
                <div class="pill-card"><div class="pill-title">Soporte-TI</div><div class="pill-val">{calc_val('Soporte-TI')}</div></div>
            </div>
            <div class="grid-4" style="margin-bottom: 12px;">
                <div class="pill-card"><div class="pill-title">Soporte-Recursos Humanos</div><div class="pill-val">{calc_val('Soporte-Recursos Humanos')}</div></div>
                <div class="pill-card"><div class="pill-title">Soporte-Contabilidad</div><div class="pill-val">{calc_val('Soporte-Contabilidad')}</div></div>
                <div class="pill-card"><div class="pill-title">Control Interno / Finanzas</div><div class="pill-val">{calc_val('Control Interno / Finanzas')}</div></div>
                <div class="pill-card"><div class="pill-title">Creditos y Cobranza</div><div class="pill-val">{calc_val('Creditos y Cobranza')}</div></div>
            </div>
            <div class="grid-2">
                <div class="pill-card"><div class="pill-title">Gastos Financieros</div><div class="pill-val">{calc_val('Gastos Financieros')}</div></div>
                <div class="pill-card"><div class="pill-title">ELUC</div><div class="pill-val">{calc_val('ELUC')}</div></div>
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

# TAB 6: CENTRO DE ALERTAS (unificado: por motivo + por comportamiento de cuenta)
with tab6:
    st.markdown(
        '<div class="header-banner"><h1>🚨 Centro de Alertas</h1>'
        '<p>Dos niveles de análisis: por motivo del asiento contable, y por comportamiento '
        'estadístico/lógico de la cuenta (IA).</p></div>',
        unsafe_allow_html=True
    )

    vista_alertas = st.radio(
        "Tipo de análisis:",
        options=["📋 Por Motivo del Asiento", "🤖 Por Comportamiento de Cuenta (IA)"],
        horizontal=True,
        key="vista_alertas_tab6",
    )
    st.markdown("---")

    # ============================================================
    # VISTA 1: POR MOTIVO DEL ASIENTO (glosa, saldo, CECO)
    # ============================================================
    if vista_alertas == "📋 Por Motivo del Asiento":
        st.caption(
            "Clasifica cada asiento según palabras clave en la glosa, el saldo y el Centro "
            "de Costo: Reversiones/Extornos, Pendiente CECO, Provisión, Reclasificación, "
            "Gastos Financieros u Operación Normal."
        )

        df_f['Motivo_Alerta'] = clasificar_motivo_vectorizado(df_f)

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

    # ============================================================
    # VISTA 2: POR COMPORTAMIENTO DE CUENTA (IA)
    # ============================================================
    else:
        st.caption(
            "Combina dos criterios: (1) **naturaleza contable** — cuentas de ajuste, redondeo "
            "o partidas varias que por lógica deberían tener montos marginales — y "
            "(2) **desviación estadística** — el monto del mes comparado contra el promedio "
            "histórico propio de esa cuenta (Z-score)."
        )

        col_u1, col_u2 = st.columns(2)
        with col_u1:
            umbral_residual = st.slider(
                "💰 Umbral (S/) para cuentas de ajuste/redondeo:",
                min_value=100, max_value=10000, value=1000, step=100,
                help="Si una cuenta de naturaleza residual (redondeo, ajuste, partidas varias) "
                     "supera este monto en un mes, se marca como alerta.",
            )
        with col_u2:
            umbral_z = st.slider(
                "📐 Sensibilidad estadística (Z-score):",
                min_value=1.0, max_value=4.0, value=2.0, step=0.1,
                help="Cuántas desviaciones estándar debe alejarse un mes del promedio histórico "
                     "de su propia cuenta para considerarse atípico. Menor valor = más sensible.",
            )

        df_alertas_ia = detectar_alertas_ia(df_f, umbral_residual=umbral_residual, umbral_z=umbral_z)

        st.markdown("---")
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Alertas Detectadas</div>'
                f'<div class="kpi-value">{len(df_alertas_ia):,}</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Monto Total en Riesgo</div>'
                f'<div class="kpi-value">S/ {df_alertas_ia["Saldo"].sum():,.2f}</div></div>',
                unsafe_allow_html=True,
            )
        with k3:
            cuenta_critica = (
                df_alertas_ia.iloc[0]["Cuenta Contable"] if len(df_alertas_ia) > 0 else "—"
            )
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Cuenta Más Crítica</div>'
                f'<div class="kpi-value" style="font-size:14px;">{cuenta_critica}</div></div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")
        col_g1, col_g2 = st.columns([1.3, 1])
        with col_g1:
            st.subheader("📈 Top Cuentas con Montos Atípicos")
            if len(df_alertas_ia) > 0:
                top_alertas = df_alertas_ia.head(12).copy()
                top_alertas["Etiqueta"] = (
                    top_alertas["Cuenta Contable"] + " (Mes " + top_alertas["Mes"].astype(str) + ")"
                )
                fig_ia = px.bar(
                    top_alertas,
                    x="Saldo",
                    y="Etiqueta",
                    orientation="h",
                    text_auto=".2s",
                    color_discrete_sequence=[COLOR_ROJO],
                )
                fig_ia.update_layout(
                    height=420, plot_bgcolor="white", yaxis=dict(autorange="reversed")
                )
                st.plotly_chart(fig_ia, use_container_width=True)
            else:
                st.success("✅ No se detectaron cuentas atípicas con los umbrales actuales.")

        with col_g2:
            st.subheader("🗂️ Alertas por Cuenta")
            if len(df_alertas_ia) > 0:
                resumen_cuenta = (
                    df_alertas_ia.groupby("Cuenta Contable")["Saldo"]
                    .agg(["sum", "count"])
                    .reset_index()
                    .sort_values("sum", key=lambda s: s.abs(), ascending=False)
                )
                resumen_cuenta.columns = ["Cuenta Contable", "Monto (S/)", "N° Alertas"]
                fig_resumen = px.pie(
                    resumen_cuenta,
                    names="Cuenta Contable",
                    values="Monto (S/)",
                    hole=0.45,
                    color_discrete_sequence=PALETA_ARMONICA,
                )
                fig_resumen.update_layout(height=420)
                st.plotly_chart(fig_resumen, use_container_width=True)

        st.markdown("---")
        st.subheader("📋 Detalle Preciso de Alertas (Monto, Mes, Motivo)")

        if len(df_alertas_ia) > 0:
            filtro_cuenta_ia = st.selectbox(
                "Filtrar por Cuenta Contable:",
                options=["Todas"] + sorted(df_alertas_ia["Cuenta Contable"].unique().tolist()),
            )
            df_tabla_ia = (
                df_alertas_ia
                if filtro_cuenta_ia == "Todas"
                else df_alertas_ia[df_alertas_ia["Cuenta Contable"] == filtro_cuenta_ia]
            )
            cols_finales = [
                "Mes", "Centro de Costos", "Cuenta Contable", "Saldo",
                "Promedio_Historico", "Z_Score", "Motivo_Alerta_IA",
            ]
            st.dataframe(
                df_tabla_ia[cols_finales].rename(columns={
                    "Saldo": "Monto del Mes (S/)",
                    "Promedio_Historico": "Promedio Histórico (S/)",
                    "Z_Score": "Z-Score",
                    "Motivo_Alerta_IA": "Motivo Preciso",
                }).style.format({
                    "Monto del Mes (S/)": "S/ {:,.2f}",
                    "Promedio Histórico (S/)": "S/ {:,.2f}",
                    "Z-Score": "{:.2f}",
                }),
                use_container_width=True,
                height=400,
                hide_index=True,
            )
        else:
            st.info("Ajusta los umbrales de arriba si esperabas ver alertas específicas.")

