import datetime
import io
import os
import textwrap
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

    /* --- Panel de Filtros (derecha) --- */
    div[data-testid="stMultiSelect"] > label {
        font-size: 12px; font-weight: 700; color: #475569;
        text-transform: uppercase; letter-spacing: 0.4px;
    }
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] {
        border-radius: 8px;
    }
    /* Chips de selección: gris/azulado neutro en vez de rojo saturado */
    div[data-testid="stMultiSelect"] span[data-baseweb="tag"],
    div[data-testid="stMultiSelect"] div[data-baseweb="tag"],
    [data-testid="stPopover"] span[data-baseweb="tag"],
    [data-testid="stPopover"] div[data-baseweb="tag"] {
        background-color: #e2e8f0 !important;
        background: #e2e8f0 !important;
        border: 1px solid #cbd5e1 !important;
        color: #1e293b !important;
        border-radius: 6px !important;
        font-size: 11.5px !important;
        font-weight: 600 !important;
    }
    div[data-testid="stMultiSelect"] span[data-baseweb="tag"] *,
    [data-testid="stPopover"] span[data-baseweb="tag"] * {
        color: #1e293b !important;
        fill: #64748b !important;
    }
    /* Contenedor de chips con altura máxima y scroll (evita que "Sucursales" se desborde) */
    div[data-testid="stMultiSelect"] div[data-baseweb="select"] > div {
        max-height: 130px;
        overflow-y: auto;
    }
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

# Paleta exclusiva del Dashboard Financiero: cálida y oscura (carmesí, navy, ámbar,
# terracota, teal profundo, gris piedra) — deliberadamente SIN celestes/azules brillantes
# para que contraste con el rojo carmesí y el navy en vez de competir con ellos.
PALETA_FINANCIERO = [
    "#a81121",  # carmesí (principal)
    "#1e293b",  # navy
    "#d97706",  # ámbar
    "#7c2d12",  # terracota oscuro
    "#0f766e",  # teal profundo
    "#b45309",  # ámbar oscuro
    "#57534e",  # gris piedra
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


_MESES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}


def _a_int(v):
    """np.int64(2025) -> 2025 (evita que el PDF imprima 'np.int64(...)')."""
    try:
        return int(v)
    except (TypeError, ValueError):
        return v


def _resumir_anios(anios):
    vals = sorted({_a_int(a) for a in anios})
    return ", ".join(str(a) for a in vals) if vals else "Todos"


def _resumir_meses(meses):
    """[1,2,3,4,5,6,7,8,9,12] -> 'Ene-Sep, Dic'."""
    ms = sorted({_a_int(m) for m in meses})
    if not ms:
        return "Todos"
    rangos, ini, prev = [], ms[0], ms[0]
    for m in ms[1:]:
        if isinstance(m, int) and isinstance(prev, int) and m == prev + 1:
            prev = m
            continue
        rangos.append((ini, prev))
        ini = prev = m
    rangos.append((ini, prev))
    partes = []
    for a, b in rangos:
        na, nb = _MESES_ABR.get(a, str(a)), _MESES_ABR.get(b, str(b))
        if a == b:
            partes.append(na)
        elif isinstance(a, int) and isinstance(b, int) and b == a + 1:
            partes.append(f"{na}, {nb}")
        else:
            partes.append(f"{na}-{nb}")
    return ", ".join(partes)


def _wrap(txt, n=20):
    """Parte una etiqueta larga en varias lineas (<br>) para los ejes de Plotly."""
    return "<br>".join(textwrap.wrap(str(txt), n)) or str(txt)


def _acortar(txt, n=30):
    txt = str(txt)
    return txt if len(txt) <= n else txt[: n - 1] + "…"


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

    def tabla(headers, rows, col_widths, x=None, alto_fila=5.0):
        x0 = MARGEN if x is None else x

        def encabezado():
            pdf.set_x(x0)
            pdf.set_font("helvetica", "B", 8)
            pdf.set_fill_color(226, 232, 240)
            for h, w in zip(headers, col_widths):
                pdf.cell(w, 6, limpiar_txt(h), border=1, align="C", fill=True)
            pdf.ln()
            pdf.set_font("helvetica", "", 8)

        encabezado()
        for row in rows:
            if pdf.will_page_break(alto_fila):
                pdf.add_page()
                pdf.set_y(MARGEN)
                encabezado()
            pdf.set_x(x0)
            for val, w in zip(row, col_widths):
                es_num = isinstance(val, (int, float, np.integer, np.floating))
                align = "R" if es_num else "L"
                texto = f"{val:,.2f}" if es_num else str(val)
                pdf.cell(w, alto_fila, limpiar_txt(texto), border=1, align=align)
            pdf.ln()
        pdf.ln(2)

    def fig_a_png(fig, w_mm, h_mm):
        """Renderiza una figura Plotly a PNG (bytes) via kaleido, con el mismo
        estilo de la app, lista para insertarse en el PDF."""
        fig2 = go.Figure(fig)
        fig2.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            margin=dict(l=10, r=40, t=50, b=10),
            font=dict(size=16),
            title_font=dict(size=20),
        )
        # automargin: el margen crece solo segun el largo de las etiquetas,
        # asi no se truncan ("ativos", "porte", "igicos"...)
        fig2.update_xaxes(automargin=True)
        fig2.update_yaxes(automargin=True)
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
                "Gráfico no disponible en este equipo. Instala el motor de "
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
    if len(suc_sel) <= 5:
        suc_txt = ", ".join(str(x) for x in sorted(suc_sel)) or "Ninguna"
    else:
        suc_txt = f"{len(suc_sel)} seleccionadas"
    filtro_txt = (
        f"Filtros aplicados  ->  Años: {_resumir_anios(anios_sel)}   |   "
        f"Meses: {_resumir_meses(meses_sel)}   |   "
        f"Sucursales: {suc_txt}"
    )
    pdf.set_x(MARGEN)
    pdf.multi_cell(ANCHO_UTIL, 5, limpiar_txt(filtro_txt), new_x="LMARGIN", new_y="NEXT")
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
    h_charts = 115

    if "Mes" in df_f.columns:
        df_mes_pdf = df_f.groupby("Mes")["Saldo"].sum().reset_index().sort_values("Mes")
        # descarta meses sin movimiento real (evita la linea cayendo a 0 en el mes 12)
        if not df_mes_pdf.empty:
            tope = df_mes_pdf["Saldo"].abs().max()
            df_mes_pdf = df_mes_pdf[df_mes_pdf["Saldo"].abs() > tope * 0.002].copy()
        df_mes_pdf["Mes_txt"] = df_mes_pdf["Mes"].map(
            lambda m: _MESES_ABR.get(_a_int(m), str(m))
        )
        orden_meses = list(df_mes_pdf["Mes_txt"])
        fig_m = px.bar(
            df_mes_pdf, x="Mes_txt", y="Saldo", text_auto=".2s",
            color_discrete_sequence=[COLOR_ROJO], title="Tendencia y Evolución Mensual",
        )
        fig_m.update_traces(textposition="outside", cliponaxis=False)
        fig_m.add_trace(go.Scatter(
            x=df_mes_pdf["Mes_txt"], y=df_mes_pdf["Saldo"], mode="lines+markers",
            line=dict(color=COLOR_NAVY, width=2), marker=dict(size=6), hoverinfo="skip",
        ))
        fig_m.update_xaxes(
            title_text="Mes", type="category",
            categoryorder="array", categoryarray=orden_meses,
        )
        fig_m.update_yaxes(title_text="Saldo (S/)")
        fig_m.update_layout(showlegend=False)
        insertar_grafico(fig_m, MARGEN, y_charts, w_izq, h_charts)

    df_comp_pdf = df_f.groupby("Comportamiento")["Saldo"].sum().reset_index()
    fig_p = px.pie(
        df_comp_pdf, names="Comportamiento", values="Saldo", hole=0.45,
        color_discrete_sequence=PALETA_ARMONICA, title="Comportamiento del Gasto",
    )
    insertar_grafico(fig_p, MARGEN + w_izq + 4, y_charts, w_der, h_charts)

    # ============ helpers de dibujo: replican los HTML de las pestañas 2 y 3 ============
    def _monto_ceco(nombre):
        """Igual que calc_val() de la pestaña 'Mapa de Procesos' (coincidencia EXACTA por CECO)."""
        val = df_f[df_f["Centro de Costos"] == nombre]["Saldo"].sum()
        return f"S/ {val:,.2f}"

    def _tiene_datos(nombre):
        """True si el Centro de Costos existe en los datos filtrados y su monto no es 0.
        Evita tarjetas en blanco (ej. ELUC, Gastos Financieros) cuando el CECO no aplica
        a los filtros actuales o no tiene movimiento."""
        return bool((df_f["Centro de Costos"] == nombre).any()) and _monto_ceco(nombre) != "S/ 0.00"

    def _monto_area(*areas):
        """Igual que get_monto_area() de la pestaña 'Organigrama' (contiene, sin mayúsculas)."""
        val = df_f[
            df_f["Subproceso / Área"].astype(str).str.contains("|".join(areas), case=False, na=False)
        ]["Saldo"].sum()
        return f"S/ {val:,.2f}"

    def texto_ajustado(x, y, w, h, texto, size, style="", align="L", color=(0, 0, 0)):
        """Escribe una linea de texto y reduce la fuente si no cabe en el ancho w."""
        texto = limpiar_txt(texto)
        pdf.set_font("helvetica", style, size)
        while size > 5 and pdf.get_string_width(texto) > w - 0.5:
            size -= 0.25
            pdf.set_font("helvetica", style, size)
        pdf.set_text_color(*color)
        pdf.set_xy(x, y)
        pdf.cell(w, h, texto, align=align)

    def pill(x, y, w, h, titulo, valor, borde=(37, 99, 235)):
        """Tarjeta 'pill-card' del Mapa de Procesos."""
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(203, 213, 225)
        pdf.set_line_width(0.2)
        pdf.rect(x, y, w, h, style="DF")
        pdf.set_fill_color(*borde)
        pdf.rect(x, y, 1.4, h, style="F")
        texto_ajustado(x + 3.5, y + 1.5, w - 5, 4, titulo, 7.5, "B", color=(71, 85, 105))
        texto_ajustado(x + 3.5, y + 5.6, w - 5, 4.5, valor, 10.5, "B", color=(15, 23, 42))

    # ================= 2. MAPA DE PROCESOS (mismo diseño que la pestaña) =================
    pdf.add_page()
    banner("Mapa de Procesos de Negocio", "Diseño corporativo estructurado por áreas tácticas y operativas.")

    PAD, GAP, H_CARD, G_V, H_HEAD, G_CONT = 3, 4, 11, 1.5, 6, 2.5
    X0 = MARGEN + PAD
    W_IN = ANCHO_UTIL - 2 * PAD

    def w_col(n):
        return (W_IN - (n - 1) * GAP) / n

    def cabecera_proceso(y, titulo, color):
        pdf.set_fill_color(*color)
        pdf.rect(MARGEN, y, ANCHO_UTIL, H_HEAD, style="F")
        texto_ajustado(MARGEN + 4, y + 0.8, ANCHO_UTIL - 8, 4.5, titulo, 9, "B", color=(255, 255, 255))

    def marco(y, alto):
        pdf.set_draw_color(226, 232, 240)
        pdf.set_line_width(0.3)
        pdf.rect(MARGEN, y, ANCHO_UTIL, alto, style="D")

    y = pdf.get_y()

    # --- PROCESOS ESTRATEGICOS (solo tarjetas con datos) ---
    est_cecos = [c for c in ["Gerencia General", "Nuevos Negocios"] if _tiene_datos(c)]
    est_labels = {"Gerencia General": "010101 | Gerencia General", "Nuevos Negocios": "010102 | Nuevos Negocios"}
    if est_cecos:
        alto = H_HEAD + PAD + H_CARD + PAD
        marco(y, alto)
        cabecera_proceso(y, "PROCESOS ESTRATÉGICOS", RGB_NAVY)
        w = w_col(len(est_cecos))
        yy = y + H_HEAD + PAD
        for i, ceco in enumerate(est_cecos):
            pill(X0 + i * (w + GAP), yy, w, H_CARD, est_labels[ceco], _monto_ceco(ceco), RGB_NAVY)
        y += alto + G_CONT

    # --- PROCESOS OPERATIVOS (columnas dinamicas + Marketing, solo con datos) ---
    columnas_op_full = [
        ("LOGÍSTICA", ["Logistica-Gastos Generales", "Logistica-Compras Locales", "Logistica-Importaciones"]),
        ("ALMACENES", ["Almacen-Gastos Generales", "Almacen-Gestión de Almacenes", "Almacen-Transporte Interno"]),
        ("VENTAS", ["Ventas-Gastos Generales", "Ventas-Ferreteria", "Ventas-Corporativo", "Ventas-Licitaciones"]),
        ("DISTRIBUCIÓN", ["Distribución-Gastos Generales"]),
    ]
    columnas_op = [(nom, [c for c in cecos if _tiene_datos(c)]) for nom, cecos in columnas_op_full]
    columnas_op = [(nom, cecos) for nom, cecos in columnas_op if cecos]
    marketing_ok = _tiene_datos("Marketing-Gastos Generales")

    H_COLH = 5.5
    H_BAN = 5.5
    if columnas_op or marketing_ok:
        max_cards = max([len(c[1]) for c in columnas_op] + [0])
        alto_cols = (H_COLH + 2 + max_cards * H_CARD + (max_cards - 1) * G_V) if columnas_op else 0
        alto = H_HEAD + PAD
        if columnas_op:
            alto += alto_cols + 2.5
        if marketing_ok:
            alto += H_BAN + G_V + H_CARD
        alto += PAD
        marco(y, alto)
        cabecera_proceso(y, "PROCESOS OPERATIVOS", RGB_ROJO)
        yy = y + H_HEAD + PAD
        if columnas_op:
            w = w_col(len(columnas_op))
            for i, (nom_col, cecos) in enumerate(columnas_op):
                xc = X0 + i * (w + GAP)
                pdf.set_fill_color(*RGB_ROJO)
                pdf.rect(xc, yy, w, H_COLH, style="F")
                texto_ajustado(xc, yy + 0.7, w, 4, nom_col, 8, "B", "C", (255, 255, 255))
                yc = yy + H_COLH + 2
                for ceco in cecos:
                    pill(xc, yc, w, H_CARD, ceco, _monto_ceco(ceco))
                    yc += H_CARD + G_V
            yy += alto_cols + 2.5
        if marketing_ok:
            pdf.set_fill_color(51, 65, 85)
            pdf.rect(X0, yy, W_IN, H_BAN, style="F")
            texto_ajustado(X0, yy + 0.7, W_IN, 4, "MARKETING", 8, "B", "C", (255, 255, 255))
            w_mk = 90
            pill(MARGEN + (ANCHO_UTIL - w_mk) / 2, yy + H_BAN + G_V, w_mk, H_CARD,
                 "Marketing-Gastos Generales", _monto_ceco("Marketing-Gastos Generales"))
        y += alto + G_CONT

    # --- PROCESOS DE SOPORTE (filas dinamicas, solo tarjetas con datos) ---
    filas_soporte_full = [
        ["Soporte-Gastos Generales", "Soporte-Tesoreria", "Soporte-Administración", "Soporte-TI"],
        ["Soporte-Recursos Humanos", "Soporte-Contabilidad", "Control Interno / Finanzas", "Creditos y Cobranza"],
    ]
    filas_soporte = [[c for c in fila if _tiene_datos(c)] for fila in filas_soporte_full]
    filas_soporte = [fila for fila in filas_soporte if fila]
    if filas_soporte:
        alto = H_HEAD + PAD + len(filas_soporte) * H_CARD + (len(filas_soporte) - 1) * G_V + PAD
        marco(y, alto)
        cabecera_proceso(y, "PROCESOS DE SOPORTE", RGB_NAVY)
        yy = y + H_HEAD + PAD
        for fila in filas_soporte:
            w = w_col(len(fila))
            for j, ceco in enumerate(fila):
                pill(X0 + j * (w + GAP), yy, w, H_CARD, ceco, _monto_ceco(ceco))
            yy += H_CARD + G_V

    # ================= 3. ORGANIGRAMA POR AREAS (mismo diseño que la pestaña) =================
    pdf.add_page()
    banner("Organigrama Estructural por Áreas Orgánicas",
           "Jerarquía organizacional con montos acumulados por departamento.")

    cx = ANCHO / 2
    y = 32

    # Nodo raiz (naranja)
    rw, rh = 84, 26
    pdf.set_fill_color(217, 119, 6)
    pdf.rect(cx - rw / 2, y, rw, rh, style="F")
    texto_ajustado(cx - rw / 2, y + 3, rw, 5, "DIRECCIÓN GENERAL", 12, "B", "C", (255, 255, 255))
    texto_ajustado(cx - rw / 2, y + 9.5, rw, 4, "Gerencia & Administración", 8.5, "B", "C", (255, 237, 213))
    monto_total_org = f"S/ {df_f['Saldo'].sum():,.2f}"
    bw = 54
    pdf.set_fill_color(168, 90, 5)
    pdf.rect(cx - bw / 2, y + 16, bw, 7, style="F")
    texto_ajustado(cx - bw / 2, y + 16.8, bw, 5, monto_total_org, 10.5, "B", "C", (255, 255, 255))

    # Conectores
    col_w = ANCHO_UTIL / 4
    centros = [MARGEN + col_w * (i + 0.5) for i in range(4)]
    y_h = y + rh + 9
    pdf.set_draw_color(148, 163, 184)
    pdf.set_line_width(0.7)
    pdf.line(cx, y + rh, cx, y_h)
    pdf.line(centros[0], y_h, centros[-1], y_h)
    y_card = y_h + 9
    for c in centros:
        pdf.line(c, y_h, c, y_card)

    ramas = [
        dict(
            titulo="GERENCIA ESTRATÉGICA", sub="Planeación & Expansión", rojo=False,
            monto=_monto_area("Gerencia General"),
            hijos=[
                ("Gerencia General", _monto_area("Gerencia General")),
            ],
        ),
        dict(
            titulo="GERENCIA OPERACIONES", sub="Cadena de Suministro", rojo=True,
            monto=_monto_area("Logística", "Compras Locales", "Importaciones", "Gestión de Almacenes", "Transporte Interno", "Distribución"),
            hijos=[
                ("Logística & Compras", _monto_area("Logística", "Compras Locales", "Importaciones")),
                ("Gestión Almacenes", _monto_area("Gestión de Almacenes", "Transporte Interno")),
                ("Distribución y Flota", _monto_area("Distribución")),
            ],
        ),
        dict(
            titulo="GERENCIA COMERCIAL", sub="Ventas & Marketing", rojo=False,
            monto=_monto_area("Ventas", "Ferretería", "Corporativo", "Licitaciones", "Marketing"),
            hijos=[
                ("Ferretería & Corp.", _monto_area("Ventas Ferretería", "Ventas Corporativo")),
                ("Licitaciones", _monto_area("Licitaciones")),
                ("Marketing Digital", _monto_area("Marketing")),
            ],
        ),
        dict(
            titulo="G. ADMIN. & FINANZAS", sub="Soporte Corporativo", rojo=False,
            monto=_monto_area("Recursos Humanos", "Tesorería", "Contabilidad", "Control Interno & Finanzas", "TI & Sistemas", "Créditos y Cobranzas", "Administración", "Soporte General"),
            hijos=[
                ("Recursos Humanos", _monto_area("Recursos Humanos")),
                ("Finanzas & Contabilidad", _monto_area("Tesorería", "Contabilidad", "Control Interno & Finanzas")),
                ("TI, Sistemas & Créditos", _monto_area("TI & Sistemas", "Créditos y Cobranzas", "Administración")),
            ],
        ),
    ]

    aw, ah = col_w * 0.92, 30
    ch, cg = 15, 5
    for c, rama in zip(centros, ramas):
        xa = c - aw / 2
        pdf.set_fill_color(*(RGB_ROJO if rama["rojo"] else RGB_NAVY))
        pdf.rect(xa, y_card, aw, ah, style="F")
        texto_ajustado(xa, y_card + 3, aw, 4.5, rama["titulo"], 9.5, "B", "C", (255, 255, 255))
        texto_ajustado(xa, y_card + 9, aw, 4, rama["sub"], 8, "", "C", (203, 213, 225))
        pdf.set_draw_color(120, 135, 155)
        pdf.set_line_width(0.2)
        pdf.line(xa + 4, y_card + 16, xa + aw - 4, y_card + 16)
        texto_ajustado(xa, y_card + 19, aw, 5, rama["monto"], 11, "B", "C", (248, 250, 252))

        # conector tarjeta -> hijos
        y_hijo = y_card + ah + 9
        pdf.set_draw_color(203, 213, 225)
        pdf.set_line_width(0.7)
        pdf.line(c, y_card + ah, c, y_hijo)

        for nombre, valor in rama["hijos"]:
            pdf.set_fill_color(255, 255, 255)
            pdf.set_draw_color(226, 232, 240)
            pdf.set_line_width(0.2)
            pdf.rect(xa, y_hijo, aw, ch, style="DF")
            pdf.set_fill_color(*RGB_ROJO)
            pdf.rect(xa, y_hijo, 1.3, ch, style="F")
            pdf.set_font("helvetica", "B", 8.5)
            bv = min(pdf.get_string_width(limpiar_txt(valor)) + 4, aw * 0.5)
            pdf.set_fill_color(241, 245, 249)
            pdf.rect(xa + aw - bv - 2, y_hijo + 3.5, bv, 8, style="F")
            texto_ajustado(xa + aw - bv - 2, y_hijo + 4.5, bv, 6, valor, 8.5, "B", "C", (15, 23, 42))
            texto_ajustado(xa + 3.5, y_hijo + 4.5, aw - bv - 8, 6, nombre, 8, "B", "L", (51, 65, 85))
            y_hijo += ch + cg

    pdf.set_text_color(0, 0, 0)
    pdf.set_draw_color(0, 0, 0)
    pdf.set_line_width(0.2)

    # ================= 4. MAPA SUCURSALES PERU =================
    pdf.add_page()
    banner("Mapa de Perú y Distribución por Sucursales", "Geolocalización de sedes con montos ejecutados")
    y0 = pdf.get_y()
    fig_map = _fig_mapa_peru(df_f)
    if fig_map is not None:
        insertar_grafico(fig_map, MARGEN, y0, ANCHO_UTIL * 0.50, 150)

    df_suc_bar = df_f.groupby("Sucursal")["Saldo"].sum().reset_index().sort_values("Saldo")
    fig_suc = px.bar(
        df_suc_bar, x="Saldo", y="Sucursal", orientation="h", text_auto=".2s",
        color="Saldo", color_continuous_scale=[COLOR_NAVY, "#7c2d12", COLOR_ROJO],
        title="Gasto por Sucursal",
    )
    fig_suc.update_traces(textposition="outside", cliponaxis=False)
    fig_suc.update_layout(coloraxis_showscale=False)
    insertar_grafico(fig_suc, MARGEN + ANCHO_UTIL * 0.52, y0, ANCHO_UTIL * 0.48, 150)

    # ================= 5. TOP 15 ASIENTOS =================
    pdf.add_page()
    banner("Top 15 Asientos por Monto Absoluto")
    titulo_seccion("Detalle de las transacciones de mayor impacto")
    df_top_asientos = df_f.reindex(df_f["Saldo"].abs().sort_values(ascending=False).index).head(15)
    cols_disp = [
        (
            str(r.get("Mes", "")),
            str(r.get("Centro de Costos", ""))[:40],
            str(r.get("Glosa / Comentario", ""))[:65],
            r.get("Saldo", 0),
        )
        for _, r in df_top_asientos.iterrows()
    ]
    tabla(["Mes", "Centro de Costo", "Glosa", "Monto (S/)"], cols_disp, [18, 75, 105, 40])

    # ================= 6. CENTRO DE ALERTAS =================
    pdf.add_page()
    banner("Centro de Alertas", "Por motivo del asiento y por comportamiento estadístico de la cuenta (IA)")
    titulo_seccion("Alertas por Motivo del Asiento")

    df_f_pdf = df_f.copy()
    df_f_pdf["Motivo_Alerta"] = clasificar_motivo_vectorizado(df_f_pdf)
    df_motivo_pdf = df_f_pdf.groupby("Motivo_Alerta")["Saldo"].agg(["sum", "count"]).reset_index()
    df_motivo_pdf.columns = ["Motivo", "Monto", "Cantidad"]
    df_motivo_pdf["Motivo"] = df_motivo_pdf["Motivo"].map(lambda t: _wrap(t, 16))

    y0 = pdf.get_y()
    fig_motivo_monto = px.bar(
        df_motivo_pdf, x="Motivo", y="Monto", color="Motivo",
        text_auto=".2s", color_discrete_sequence=PALETA_ARMONICA, title="Monto por Motivo (S/)",
    )
    fig_motivo_monto.update_layout(showlegend=False)
    fig_motivo_monto.update_traces(textposition="outside", cliponaxis=False)
    insertar_grafico(fig_motivo_monto, MARGEN, y0, ANCHO_UTIL * 0.48, 110)

    fig_motivo_cant = px.bar(
        df_motivo_pdf, x="Motivo", y="Cantidad", color="Motivo",
        text_auto=True, color_discrete_sequence=PALETA_ARMONICA, title="Cantidad de Asientos",
    )
    fig_motivo_cant.update_layout(showlegend=False)
    fig_motivo_cant.update_traces(textposition="outside", cliponaxis=False)
    insertar_grafico(fig_motivo_cant, MARGEN + ANCHO_UTIL * 0.52, y0, ANCHO_UTIL * 0.48, 110)

    # --- Alertas IA (pagina nueva, incluye tabla de detalle) ---
    pdf.add_page()
    banner("Centro de Alertas (continuación)", "Cuentas con comportamiento atípico (motor IA, umbrales por defecto)")
    titulo_seccion("Alertas IA - Cuentas Atípicas")
    df_alertas_ia_pdf = detectar_alertas_ia(df_f, umbral_residual=1000, umbral_z=2.0)

    if len(df_alertas_ia_pdf) > 0:
        y0 = pdf.get_y()
        top_ia = df_alertas_ia_pdf.head(12).copy()
        top_ia["Etiqueta"] = top_ia["Cuenta Contable"].map(lambda t: _acortar(t, 32)) + " (Mes " + top_ia["Mes"].astype(str) + ")"
        fig_ia = px.bar(
            top_ia.sort_values("Saldo"), x="Saldo", y="Etiqueta", orientation="h",
            text_auto=".2s", color_discrete_sequence=[COLOR_ROJO], title="Top Cuentas Atípicas",
        )
        fig_ia.update_traces(textposition="outside", cliponaxis=False)
        insertar_grafico(fig_ia, MARGEN, y0, ANCHO_UTIL * 0.55, 66)

        resumen_cuenta = (
            df_alertas_ia_pdf.groupby("Cuenta Contable")["Saldo"].sum().reset_index()
            .sort_values("Saldo", key=lambda s: s.abs(), ascending=False).head(8)
        )
        resumen_cuenta["Cuenta Contable"] = resumen_cuenta["Cuenta Contable"].map(lambda t: _acortar(t, 34))
        fig_resumen = px.pie(
            resumen_cuenta, names="Cuenta Contable", values="Saldo", hole=0.45,
            color_discrete_sequence=PALETA_ARMONICA, title="Distribución por Cuenta",
        )
        insertar_grafico(fig_resumen, MARGEN + ANCHO_UTIL * 0.56, y0, ANCHO_UTIL * 0.44, 66)
        pdf.set_y(y0 + 70)

        titulo_seccion("Detalle preciso (Top 15)")
        filas_ia = [
            (str(row["Mes"]), str(row["Cuenta Contable"])[:70], row["Saldo"], row["Z_Score"])
            for _, row in df_alertas_ia_pdf.head(15).iterrows()
        ]
        tabla(["Mes", "Cuenta Contable", "Monto (S/)", "Z-Score"], filas_ia, [20, 140, 45, 30], alto_fila=4.8)
    else:
        pdf.set_font("helvetica", "", 9)
        pdf.cell(0, 6, limpiar_txt("No se detectaron cuentas atípicas con los umbrales por defecto."), ln=True)

    return bytes(pdf.output())

def _monto_ceco(df_x, nombre):
    """Monto por Centro de Costos con coincidencia EXACTA (igual que calc_val de la pestaña Mapa de Procesos)."""
    val = df_x[df_x["Centro de Costos"] == nombre]["Saldo"].sum()
    return f"S/ {val:,.2f}"


def _tiene_datos_ceco(df_x, nombre):
    """True si el Centro de Costos existe en los datos filtrados y su monto no es 0.
    Evita tarjetas en blanco (ej. ELUC, Gastos Financieros) cuando el CECO no aplica
    a los filtros actuales o no tiene movimiento."""
    return bool((df_x["Centro de Costos"] == nombre).any()) and _monto_ceco(df_x, nombre) != "S/ 0.00"


def _monto_area(df_x, *areas):
    """Monto por área (contiene, sin mayúsculas). Igual que get_monto_area de la pestaña Organigrama."""
    val = df_x[
        df_x["Subproceso / Área"].astype(str).str.contains("|".join(areas), case=False, na=False)
    ]["Saldo"].sum()
    return f"S/ {val:,.2f}"


def _fig_mapa_peru(df_f):
    """Figura Plotly del mapa de Perú (compartida por el PDF y el PowerPoint)."""
    # Se agrupa por CIUDAD (MTRU y TRU comparten coordenadas: antes sus etiquetas se encimaban)
    # y se excluye "Sin Sucursal" (monto negativo, no tiene ubicacion real).
    df_geo_suc = (
        df_f[df_f["Sucursal"] != "Sin Sucursal"]
        .groupby(["Ciudad_Nombre", "Latitud", "Longitud"])["Saldo"].sum().reset_index()
    )
    df_geo_suc = df_geo_suc.dropna(subset=["Latitud", "Longitud"])
    df_geo_suc = df_geo_suc[df_geo_suc["Saldo"] > 0].copy()
    if df_geo_suc.empty:
        return None
    df_geo_suc["size_mapa"] = df_geo_suc["Saldo"]
    df_geo_suc["Etiqueta"] = (
        df_geo_suc["Ciudad_Nombre"] + "<br>S/ " + (df_geo_suc["Saldo"] / 1e6).map("{:.2f}M".format)
    )
    # Alterna la posicion del texto (ordenado de norte a sur) para que no se pisen
    df_geo_suc = df_geo_suc.sort_values("Latitud", ascending=False).reset_index(drop=True)
    _pos = ["middle right", "middle left"]
    df_geo_suc["pos_texto"] = [_pos[i % 2] for i in range(len(df_geo_suc))]

    fig_map = px.scatter_geo(
        df_geo_suc, lat="Latitud", lon="Longitud", size="size_mapa",
        hover_name="Ciudad_Nombre", text="Etiqueta", color="Saldo",
        color_continuous_scale="Reds", projection="mercator", size_max=40,
    )
    # En imagen estatica (kaleido) projection_scale/center no aplican igual que en el
    # navegador: se fijan los rangos de longitud/latitud de Perú.
    fig_map.update_geos(
        lonaxis_range=[-83, -67], lataxis_range=[-19, 0.5],
        resolution=50,
        showland=True, landcolor="#f1f5f9", showocean=True, oceancolor="#e2e8f0",
        showcountries=True, countrycolor="#94a3b8",
        showcoastlines=True, coastlinecolor="#94a3b8",
    )
    fig_map.update_traces(textposition=list(df_geo_suc["pos_texto"]), marker=dict(sizemin=8))
    fig_map.update_layout(margin=dict(l=0, r=0, t=10, b=0), coloraxis_showscale=False)
    return fig_map


def generar_reporte_pptx(df_f, anios_sel, meses_sel, suc_sel):
    """Genera la presentación PowerPoint (.pptx) con el mismo contenido del dashboard.
    - Gráficos NATIVOS de PowerPoint (editables), no imágenes.
    - Mapa de Procesos y Organigrama dibujados con formas editables.
    - Solo el mapa de Perú es imagen (PowerPoint no tiene mapa nativo)."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
    from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import (
        XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION, XL_TICK_LABEL_POSITION,
    )
    from pptx.oxml.ns import qn

    FUENTE = "Calibri"
    SW, SH = 13.333, 7.5
    MX = 0.5
    W_UTIL = SW - 2 * MX

    def C(hexa):
        return RGBColor(*hex_a_rgb(hexa))

    NAVY, ROJO, AZUL = C(COLOR_NAVY), C(COLOR_ROJO), C(COLOR_AZUL)
    BLANCO = RGBColor(255, 255, 255)
    TXT, TXT2, GRIS = C("#0f172a"), C("#475569"), C("#64748b")
    BORDE, FONDO, CLARO = C("#cbd5e1"), C("#f8fafc"), C("#f1f5f9")
    PALETA = [C(c) for c in PALETA_ARMONICA]

    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    contador = {"n": 0}

    # ---------------- helpers de formato ----------------
    def fmt_corto(v):
        a = abs(v)
        if a >= 1e6:
            t = f"{a / 1e6:.1f}M"
        elif a >= 1e3:
            t = f"{a / 1e3:.0f}k"
        else:
            t = f"{a:,.0f}"
        return ("-" if v < 0 else "") + t

    def recorta(txt, n):
        txt = str(txt)
        return txt if len(txt) <= n else txt[: n - 1] + "…"

    # ---------------- helpers de dibujo ----------------
    def rect(s, x, y, w, h, fill=None, line=None, lw=0.75):
        sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
        sh.shadow.inherit = False
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = fill
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = line
            sh.line.width = Pt(lw)
        return sh

    def poner_texto(sh, lineas, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, m=(0.1, 0.03, 0.1, 0.03)):
        tf = sh.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.NONE
        tf.margin_left, tf.margin_top, tf.margin_right, tf.margin_bottom = [Inches(v) for v in m]
        tf.vertical_anchor = anchor
        for i, (t, sz, neg, col) in enumerate(lineas):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            r = p.add_run()
            r.text = t
            r.font.name = FUENTE
            r.font.size = Pt(sz)
            r.font.bold = neg
            r.font.color.rgb = col

    def texto(s, x, y, w, h, t, sz, neg=False, col=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        poner_texto(tb, [(t, sz, neg, col or TXT)], align, anchor, m=(0, 0, 0, 0))
        return tb

    def linea(s, x1, y1, x2, y2, col=None, w=2.25):
        cn = s.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
        )
        cn.line.color.rgb = col or C("#94a3b8")
        cn.line.width = Pt(w)
        estilo = cn._element.find(qn("p:style"))
        if estilo is not None:      # quita sombra/efectos heredados del tema
            cn._element.remove(estilo)

    def nueva_slide(titulo=None, subtitulo=""):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        contador["n"] += 1
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = FONDO
        if titulo:
            rect(s, 0, 0, SW, 1.05, NAVY)
            rect(s, 0, 0, 0.14, 1.05, ROJO)
            texto(s, MX, 0.14, W_UTIL, 0.5, titulo, 28, True, BLANCO)
            if subtitulo:
                texto(s, MX, 0.66, W_UTIL, 0.3, subtitulo, 14, False, C("#cbd5e1"))
            texto(s, MX, 7.14, 8, 0.25, "Control Presupuestal FP&A  |  Junior FP&A", 10, False, GRIS)
            texto(s, SW - MX - 1, 7.14, 1, 0.25, str(contador["n"]), 10, False, GRIS, PP_ALIGN.RIGHT)
        return s

    def pill(s, x, y, w, h, titulo, valor, borde=None, sz_t=11.5, sz_v=15.5):
        card = rect(s, x, y, w, h, BLANCO, BORDE)
        poner_texto(card, [(titulo, sz_t, True, TXT2), (valor, sz_v, True, TXT)], m=(0.22, 0.03, 0.08, 0.03))
        rect(s, x, y, 0.08, h, borde or ROJO)

    # ---------------- helpers de gráficos nativos ----------------
    def estilo_chart(ch, titulo, size=12):
        ch.font.name = FUENTE
        ch.font.size = Pt(size)
        ch.font.color.rgb = TXT2
        ch.has_title = True
        ch.chart_title.text_frame.text = titulo
        r = ch.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size, r.font.bold, r.font.name, r.font.color.rgb = Pt(15), True, FUENTE, TXT
        ch.chart_title.include_in_layout = False

    def barras(s, x, y, w, h, cats, vals, titulo, color, horizontal=True, size=12):
        cd = CategoryChartData()
        cd.categories = [str(c) for c in cats]
        cd.add_series("Monto", [float(v) for v in vals])
        tipo = XL_CHART_TYPE.BAR_CLUSTERED if horizontal else XL_CHART_TYPE.COLUMN_CLUSTERED
        ch = s.shapes.add_chart(tipo, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
        estilo_chart(ch, titulo, size)
        ch.has_legend = False
        plot = ch.plots[0]
        plot.gap_width = 45
        plot.vary_by_categories = False
        ser = plot.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = color
        ser.invert_if_negative = False
        vmax = max([float(v) for v in vals] + [0.0])
        vmin = min([float(v) for v in vals] + [0.0])
        if vmax <= 0 and vmin >= 0:
            vmax = 1.0
        va = ch.value_axis
        va.visible = False
        va.has_major_gridlines = False
        va.maximum_scale = vmax * 1.22 if vmax > 0 else 0
        va.minimum_scale = vmin * 1.45 if vmin < 0 else 0
        ca = ch.category_axis
        ca.has_major_gridlines = False
        ca.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        ca.tick_labels.font.size = Pt(size)
        ca.tick_labels.font.color.rgb = TXT2
        ca.format.line.color.rgb = BORDE
        if horizontal:
            ca.reverse_order = True   # el primer dato queda ARRIBA
        for i, v in enumerate(vals):
            dl = ser.points[i].data_label
            dl.position = XL_LABEL_POSITION.OUTSIDE_END
            dl.text_frame.text = fmt_corto(float(v))
            r = dl.text_frame.paragraphs[0].runs[0]
            r.font.size, r.font.bold, r.font.name, r.font.color.rgb = Pt(size), True, FUENTE, TXT
        return ch

    def dona(s, x, y, w, h, cats, vals, titulo, size=12):
        cd = CategoryChartData()
        cd.categories = [str(c) for c in cats]
        cd.add_series("Monto", [abs(float(v)) for v in vals])
        ch = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y), Inches(w), Inches(h), cd).chart
        estilo_chart(ch, titulo, size)
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(size)
        ch.legend.font.name = FUENTE
        plot = ch.plots[0]
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.show_percentage = True
        dl.show_value = False
        dl.show_category_name = False
        dl.number_format = "0.0%"
        dl.number_format_is_linked = False
        dl.font.size, dl.font.bold, dl.font.color.rgb = Pt(size), True, BLANCO
        for i in range(len(vals)):
            pt = plot.series[0].points[i]
            pt.format.fill.solid()
            pt.format.fill.fore_color.rgb = PALETA[i % len(PALETA)]
        hs = plot._element.find(qn("c:holeSize"))
        if hs is not None:
            hs.set("val", "55")
        return ch

    def tabla(s, x, y, anchos, headers, filas, aligns, sz=12, alto=0.34):
        n_f, n_c = len(filas) + 1, len(headers)
        gf = s.shapes.add_table(n_f, n_c, Inches(x), Inches(y), Inches(sum(anchos)), Inches(alto * n_f))
        tb = gf.table
        tb.horz_banding = False
        tb.first_row = True
        for j, w in enumerate(anchos):
            tb.columns[j].width = Inches(w)
        for i in range(n_f):
            tb.rows[i].height = Inches(alto)

        def celda(i, j, txt, neg, col, fill):
            c = tb.cell(i, j)
            c.fill.solid()
            c.fill.fore_color.rgb = fill
            c.margin_left = c.margin_right = Inches(0.1)
            c.margin_top = c.margin_bottom = Inches(0.02)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.alignment = aligns[j] if i > 0 else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = txt
            r.font.name, r.font.size, r.font.bold, r.font.color.rgb = FUENTE, Pt(sz), neg, col

        for j, h in enumerate(headers):
            celda(0, j, h, True, BLANCO, NAVY)
        for i, fila in enumerate(filas, start=1):
            fondo = BLANCO if i % 2 else CLARO
            for j, v in enumerate(fila):
                negativo = isinstance(v, str) and v.strip().startswith("-") and aligns[j] == PP_ALIGN.RIGHT
                celda(i, j, str(v), False, ROJO if negativo else TXT, fondo)
        return tb

    # ================= 1. PORTADA =================
    if len(suc_sel) <= 5:
        suc_txt = ", ".join(str(x) for x in sorted(suc_sel)) or "Ninguna"
    else:
        suc_txt = f"{len(suc_sel)} seleccionadas"
    filtro_txt = (
        f"Años: {_resumir_anios(anios_sel)}   |   Meses: {_resumir_meses(meses_sel)}   |   "
        f"Sucursales: {suc_txt}"
    )
    s = nueva_slide()
    s.background.fill.fore_color.rgb = NAVY
    rect(s, 0.9, 2.15, 0.14, 2.3, ROJO)
    texto(s, 1.35, 2.05, 11.5, 0.9, "Reporte Ejecutivo", 46, True, BLANCO)
    texto(s, 1.35, 3.0, 11.5, 0.6, "Control Presupuestal FP&A", 28, False, C("#e2e8f0"))
    texto(s, 1.35, 3.85, 11.5, 0.4,
          f"Generado: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}   |   Junior FP&A",
          16, False, C("#94a3b8"))
    texto(s, 1.35, 4.3, 11.0, 0.9, "Filtros aplicados:  " + filtro_txt, 14, False, C("#94a3b8"))

    # ================= 2. RESUMEN GENERAL =================
    s = nueva_slide("Resumen Ejecutivo", "Panel ejecutivo de control presupuestal y análisis financiero")
    gasto_total = df_f["Saldo"].sum()
    cecos_cant = df_f["Centro de Costos"].nunique()
    registros = len(df_f)
    promedio = gasto_total / max(registros, 1)
    kpis = [
        ("Gasto acumulado total", f"S/ {gasto_total:,.2f}"),
        ("Centros de costos", f"{cecos_cant}"),
        ("Transacciones", f"{registros:,}"),
        ("Gasto promedio por registro", f"S/ {promedio:,.2f}"),
    ]
    gap = 0.25
    kw = (W_UTIL - 3 * gap) / 4
    for i, (t, v) in enumerate(kpis):
        x = MX + i * (kw + gap)
        card = rect(s, x, 1.3, kw, 1.25, BLANCO, BORDE)
        poner_texto(card, [(t.upper(), 11.5, True, GRIS), (v, 24 if len(v) <= 12 else 20, True, TXT)],
                    m=(0.2, 0.12, 0.1, 0.05))
        rect(s, x, 1.3, kw, 0.08, ROJO)

    if "Mes" in df_f.columns:
        df_mes = df_f.groupby("Mes")["Saldo"].sum().reset_index().sort_values("Mes")
        if not df_mes.empty:
            tope = df_mes["Saldo"].abs().max()
            df_mes = df_mes[df_mes["Saldo"].abs() > tope * 0.002]
        meses_txt = [_MESES_ABR.get(_a_int(m), str(m)) for m in df_mes["Mes"]]
        if len(df_mes):
            barras(s, MX, 2.8, 7.7, 4.25, meses_txt, list(df_mes["Saldo"]),
                   "Tendencia y evolución mensual (S/)", ROJO, horizontal=False, size=13)
    df_comp = df_f.groupby("Comportamiento")["Saldo"].sum().reset_index()
    if len(df_comp):
        dona(s, 8.45, 2.8, SW - MX - 8.45, 4.25, list(df_comp["Comportamiento"]),
             list(df_comp["Saldo"]), "Comportamiento del gasto", size=13)

    # ================= 3. MAPA DE PROCESOS (1/2) =================
    PAD, GAP, HH = 0.12, 0.15, 0.34
    X0, WIN = MX + PAD, W_UTIL - 2 * PAD

    def wcol(n):
        return (WIN - (n - 1) * GAP) / n

    def contenedor(s, y, alto, titulo, color, hh=HH, sz=13):
        rect(s, MX, y, W_UTIL, alto, BLANCO, C("#e2e8f0"))
        cab = rect(s, MX, y, W_UTIL, hh, color)
        poner_texto(cab, [(titulo, sz, True, BLANCO)], m=(0.2, 0, 0.1, 0))

    s = nueva_slide("Mapa de Procesos de Negocio", "Procesos estratégicos y operativos · diseño corporativo por áreas tácticas")
    HC, GV = 0.58, 0.05
    y = 1.28

    est_cecos = [c for c in ["Gerencia General", "Nuevos Negocios"] if _tiene_datos_ceco(df_f, c)]
    est_labels = {"Gerencia General": "010101 | Gerencia General", "Nuevos Negocios": "010102 | Nuevos Negocios"}
    if est_cecos:
        alto = HH + PAD + HC + PAD
        contenedor(s, y, alto, "PROCESOS ESTRATÉGICOS", NAVY)
        w = wcol(len(est_cecos))
        for i, ceco in enumerate(est_cecos):
            pill(s, X0 + i * (w + GAP), y + HH + PAD, w, HC, est_labels[ceco], _monto_ceco(df_f, ceco), NAVY)
        y += alto + 0.1

    columnas_op_full = [
        ("LOGÍSTICA", ["Logistica-Gastos Generales", "Logistica-Compras Locales", "Logistica-Importaciones"]),
        ("ALMACENES", ["Almacen-Gastos Generales", "Almacen-Gestión de Almacenes", "Almacen-Transporte Interno"]),
        ("VENTAS", ["Ventas-Gastos Generales", "Ventas-Ferreteria", "Ventas-Corporativo", "Ventas-Licitaciones"]),
        ("DISTRIBUCIÓN", ["Distribución-Gastos Generales"]),
    ]
    columnas_op = [(nom, [c for c in cecos if _tiene_datos_ceco(df_f, c)]) for nom, cecos in columnas_op_full]
    columnas_op = [(nom, cecos) for nom, cecos in columnas_op if cecos]
    marketing_ok = _tiene_datos_ceco(df_f, "Marketing-Gastos Generales")

    HCOL, HBAN = 0.28, 0.26
    if columnas_op or marketing_ok:
        max_cards = max([len(c[1]) for c in columnas_op] + [0])
        alto_cols = (HCOL + 0.06 + max_cards * HC + (max_cards - 1) * GV) if columnas_op else 0
        alto = HH + PAD
        if columnas_op:
            alto += alto_cols + 0.14
        if marketing_ok:
            alto += HBAN + 0.05 + HC
        alto += PAD
        contenedor(s, y, alto, "PROCESOS OPERATIVOS", ROJO)
        yy = y + HH + PAD
        if columnas_op:
            w = wcol(len(columnas_op))
            for i, (nom, cecos) in enumerate(columnas_op):
                xc = X0 + i * (w + GAP)
                cab = rect(s, xc, yy, w, HCOL, ROJO)
                poner_texto(cab, [(nom, 11.5, True, BLANCO)], PP_ALIGN.CENTER, m=(0.05, 0, 0.05, 0))
                yc = yy + HCOL + 0.06
                for ceco in cecos:
                    pill(s, xc, yc, w, HC, ceco, _monto_ceco(df_f, ceco))
                    yc += HC + GV
            yy += alto_cols + 0.14
        if marketing_ok:
            ban = rect(s, X0, yy, WIN, HBAN, C("#334155"))
            poner_texto(ban, [("MARKETING", 11.5, True, BLANCO)], PP_ALIGN.CENTER, m=(0.05, 0, 0.05, 0))
            wmk = 4.2
            pill(s, MX + (W_UTIL - wmk) / 2, yy + HBAN + 0.05, wmk, HC,
                 "Marketing-Gastos Generales", _monto_ceco(df_f, "Marketing-Gastos Generales"))

    # ================= 4. MAPA DE PROCESOS (2/2: Soporte) =================
    s = nueva_slide("Mapa de Procesos de Negocio", "Procesos de soporte")
    filas_soporte_full = [
        ["Soporte-Gastos Generales", "Soporte-Tesoreria", "Soporte-Administración", "Soporte-TI"],
        ["Soporte-Recursos Humanos", "Soporte-Contabilidad", "Control Interno / Finanzas", "Creditos y Cobranza"],
    ]
    filas_soporte = [[c for c in fila if _tiene_datos_ceco(df_f, c)] for fila in filas_soporte_full]
    filas_soporte = [fila for fila in filas_soporte if fila]
    HC2, GV2, PAD2, HH2 = 1.0, 0.2, 0.25, 0.42
    if filas_soporte:
        alto = HH2 + PAD2 + len(filas_soporte) * HC2 + (len(filas_soporte) - 1) * GV2 + PAD2
        y = 1.5
        contenedor(s, y, alto, "PROCESOS DE SOPORTE", NAVY, hh=HH2, sz=15)
        yy = y + HH2 + PAD2
        xin, win2 = MX + PAD2, W_UTIL - 2 * PAD2
        for fila in filas_soporte:
            n = len(fila)
            w = (win2 - (n - 1) * 0.25) / n
            for j, ceco in enumerate(fila):
                pill(s, xin + j * (w + 0.25), yy, w, HC2, ceco, _monto_ceco(df_f, ceco), sz_t=14, sz_v=22)
            yy += HC2 + GV2

    # ================= 5. ORGANIGRAMA =================
    s = nueva_slide("Organigrama Estructural por Áreas Orgánicas",
                    "Jerarquía organizacional con montos acumulados por departamento")
    col_w = W_UTIL / 4
    centros = [MX + col_w * (i + 0.5) for i in range(4)]
    cx = SW / 2
    rw, rh, ry = 4.8, 0.92, 1.28
    root = rect(s, cx - rw / 2, ry, rw, rh, C("#d97706"))
    poner_texto(root, [
        ("DIRECCIÓN GENERAL", 16, True, BLANCO),
        ("Gerencia & Administración", 11.5, False, C("#ffedd5")),
        (f"S/ {df_f['Saldo'].sum():,.2f}", 15, True, BLANCO),
    ], PP_ALIGN.CENTER)
    y_h = ry + rh + 0.22
    y_card = y_h + 0.22
    linea(s, cx, ry + rh, cx, y_h)
    linea(s, centros[0], y_h, centros[-1], y_h)
    for c in centros:
        linea(s, c, y_h, c, y_card)

    ramas = [
        ("GERENCIA ESTRATÉGICA", "Planeación & Expansión", False,
         _monto_area(df_f, "Gerencia General"),
         [("Gerencia General", _monto_area(df_f, "Gerencia General"))]),
        ("GERENCIA OPERACIONES", "Cadena de Suministro", True,
         _monto_area(df_f, "Logística", "Compras Locales", "Importaciones", "Gestión de Almacenes", "Transporte Interno", "Distribución"),
         [("Logística & Compras", _monto_area(df_f, "Logística", "Compras Locales", "Importaciones")),
          ("Gestión Almacenes", _monto_area(df_f, "Gestión de Almacenes", "Transporte Interno")),
          ("Distribución y Flota", _monto_area(df_f, "Distribución"))]),
        ("GERENCIA COMERCIAL", "Ventas & Marketing", False,
         _monto_area(df_f, "Ventas", "Ferretería", "Corporativo", "Licitaciones", "Marketing"),
         [("Ferretería & Corp.", _monto_area(df_f, "Ventas Ferretería", "Ventas Corporativo")),
          ("Licitaciones", _monto_area(df_f, "Licitaciones")),
          ("Marketing Digital", _monto_area(df_f, "Marketing"))]),
        ("G. ADMIN. & FINANZAS", "Soporte Corporativo", False,
         _monto_area(df_f, "Recursos Humanos", "Tesorería", "Contabilidad", "Control Interno & Finanzas", "TI & Sistemas", "Créditos y Cobranzas", "Administración", "Soporte General"),
         [("Recursos Humanos", _monto_area(df_f, "Recursos Humanos")),
          ("Finanzas & Contabilidad", _monto_area(df_f, "Tesorería", "Contabilidad", "Control Interno & Finanzas")),
          ("TI, Sistemas & Créditos", _monto_area(df_f, "TI & Sistemas", "Créditos y Cobranzas", "Administración"))]),
    ]
    aw, ah, ch_h, ch_g = col_w * 0.92, 1.1, 0.86, 0.12
    for c, (tit, sub, rojo, monto, hijos) in zip(centros, ramas):
        xa = c - aw / 2
        card = rect(s, xa, y_card, aw, ah, ROJO if rojo else NAVY)
        poner_texto(card, [(tit, 14, True, BLANCO), (sub, 11, False, C("#cbd5e1")), (monto, 16, True, BLANCO)],
                    PP_ALIGN.CENTER)
        y_hijo = y_card + ah + 0.22
        linea(s, c, y_card + ah, c, y_hijo, col=C("#cbd5e1"))
        for nombre, valor in hijos:
            hc = rect(s, xa, y_hijo, aw, ch_h, BLANCO, C("#e2e8f0"))
            poner_texto(hc, [(nombre, 13, True, TXT2), (valor, 16, True, TXT)], m=(0.22, 0.03, 0.08, 0.03))
            rect(s, xa, y_hijo, 0.08, ch_h, ROJO)
            y_hijo += ch_h + ch_g

    # ================= 6. SUCURSALES =================
    s = nueva_slide("Mapa de Perú y Distribución por Sucursales", "Geolocalización de sedes con montos ejecutados")
    png = None
    try:
        fig_map = _fig_mapa_peru(df_f)
        if fig_map is not None:
            fig_map.update_layout(paper_bgcolor="white", font=dict(size=18), margin=dict(l=0, r=0, t=10, b=0))
            png = fig_map.to_image(format="png", width=1200, height=1140, scale=1)
    except Exception:
        png = None
    if png:
        s.shapes.add_picture(io.BytesIO(png), Inches(MX), Inches(1.3), height=Inches(5.7))
    else:
        caja = rect(s, MX, 1.3, 5.95, 5.7, BLANCO, BORDE)
        poner_texto(caja, [("Mapa no disponible en este equipo.", 16, True, GRIS),
                           ("Instala el motor de renderizado:  pip install -U kaleido", 13, False, GRIS)],
                    PP_ALIGN.CENTER)
    df_suc = df_f.groupby("Sucursal")["Saldo"].sum().reset_index().sort_values("Saldo", ascending=False)
    if len(df_suc):
        barras(s, 6.75, 1.3, SW - MX - 6.75, 5.7, list(df_suc["Sucursal"]), list(df_suc["Saldo"]),
               "Gasto por sucursal (S/)", ROJO, size=13)

    # ================= 7. TOP 15 ASIENTOS =================
    s = nueva_slide("Top 15 Asientos por Monto Absoluto", "Detalle de las transacciones de mayor impacto")
    df_top = df_f.reindex(df_f["Saldo"].abs().sort_values(ascending=False).index).head(15)
    filas = [
        (str(r.get("Mes", "")), recorta(r.get("Centro de Costos", ""), 36),
         recorta(r.get("Glosa / Comentario", ""), 58), f"{r.get('Saldo', 0):,.2f}")
        for _, r in df_top.iterrows()
    ]
    tabla(s, MX, 1.3, [0.8, 3.6, 5.6, 2.33], ["Mes", "Centro de Costo", "Glosa", "Monto (S/)"], filas,
          [PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.RIGHT], sz=12, alto=0.34)

    # ================= 8. CENTRO DE ALERTAS: POR MOTIVO =================
    s = nueva_slide("Centro de Alertas", "Alertas por motivo del asiento")
    df_mot = df_f.copy()
    df_mot["Motivo_Alerta"] = clasificar_motivo_vectorizado(df_mot)
    df_mot = df_mot.groupby("Motivo_Alerta")["Saldo"].agg(["sum", "count"]).reset_index().sort_values("Motivo_Alerta")
    if len(df_mot):
        mitad = (W_UTIL - 0.3) / 2
        barras(s, MX, 1.3, mitad, 5.7, list(df_mot["Motivo_Alerta"]), list(df_mot["sum"]),
               "Monto por motivo (S/)", ROJO, size=12)
        barras(s, MX + mitad + 0.3, 1.3, mitad, 5.7, list(df_mot["Motivo_Alerta"]), list(df_mot["count"]),
               "Cantidad de asientos", NAVY, size=12)

    # ================= 9-10. CENTRO DE ALERTAS: IA =================
    df_ia = detectar_alertas_ia(df_f, umbral_residual=1000, umbral_z=2.0)
    s = nueva_slide("Centro de Alertas: cuentas atípicas",
                    "Comportamiento atípico de la cuenta (motor IA, umbrales por defecto)")
    if len(df_ia) > 0:
        top_ia = df_ia.head(12)
        etiquetas = [f"{recorta(c, 30)} (Mes {int(m)})" for c, m in zip(top_ia["Cuenta Contable"], top_ia["Mes"])]
        barras(s, MX, 1.3, 7.3, 5.75, etiquetas, list(top_ia["Saldo"]), "Top cuentas atípicas (S/)", ROJO, size=12)
        res = (
            df_ia.groupby("Cuenta Contable")["Saldo"].sum().reset_index()
            .sort_values("Saldo", key=lambda x: x.abs(), ascending=False).head(8)
        )
        dona(s, 8.0, 1.3, SW - MX - 8.0, 5.75, [recorta(c, 34) for c in res["Cuenta Contable"]],
             list(res["Saldo"]), "Distribución por cuenta", size=11)

        s = nueva_slide("Centro de Alertas: detalle preciso (Top 15)",
                        "Cuentas con mayor desviación estadística o naturaleza residual")
        filas = [
            (str(int(r["Mes"])), recorta(r["Cuenta Contable"], 80), f"{r['Saldo']:,.2f}", f"{r['Z_Score']:.2f}")
            for _, r in df_ia.head(15).iterrows()
        ]
        tabla(s, MX, 1.3, [0.8, 7.5, 2.4, 1.63], ["Mes", "Cuenta contable", "Monto (S/)", "Z-Score"], filas,
              [PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER], sz=12, alto=0.34)
    else:
        texto(s, MX, 1.6, W_UTIL, 0.5, "No se detectaron cuentas atípicas con los umbrales por defecto.", 18, False, GRIS)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# 3. SIDEBAR: NAVEGACIÓN (antes eran pestañas arriba)
PAGINAS = [
    ("📊", "Resumen General"),
    ("🗺️", "Mapa de Procesos"),
    ("🏛️", "Organigrama por Áreas"),
    ("🌎", "Mapa Sucursales Perú"),
    ("🔍", "Detalle por CECO / Asientos"),
    ("🚨", "Centro de Alertas"),
    ("💰", "Dashboard Financiero"),
]

if "pagina_activa" not in st.session_state:
    st.session_state.pagina_activa = PAGINAS[0][1]

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

    st.markdown("#### 🧭 Navegación")
    for icono, nombre in PAGINAS:
        activo = st.session_state.pagina_activa == nombre
        if st.button(
            f"{icono}  {nombre}",
            key=f"nav_{nombre}",
            use_container_width=True,
            type="primary" if activo else "secondary",
        ):
            st.session_state.pagina_activa = nombre
            st.rerun()

    st.markdown("---")

    generar_pdf_click = st.button("📄 Generar Reporte PDF Ejecutivo", type="primary")
    generar_pptx_click = st.button("📊 Generar Presentación PowerPoint", type="primary")
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

pagina = st.session_state.pagina_activa

# 4. LAYOUT PRINCIPAL: contenido a la izquierda, filtros a la derecha
# 4. BARRA DE FILTROS (arriba, horizontal — aplica a todas las secciones)
anios_tot = sorted(list(df["Año"].dropna().unique())) if "Año" in df.columns else []
meses_totales = sorted(list(df["Mes"].unique())) if "Mes" in df.columns else []
suc_opt = sorted(list(df["Sucursal"].unique()))

for _k, _default in (
    ("filtro_anio", anios_tot),
    ("filtro_meses", meses_totales),
    ("filtro_sucursales", suc_opt),
):
    if _k not in st.session_state:
        st.session_state[_k] = _default

# Si se pidió "Limpiar filtros" en el clic anterior, se aplica AQUÍ,
# antes de crear los widgets (Streamlit no permite tocar session_state
# de una key después de que su widget ya fue instanciado en ese mismo run).
if st.session_state.get("_reset_filtros", False):
    st.session_state.filtro_anio = anios_tot
    st.session_state.filtro_meses = meses_totales
    st.session_state.filtro_sucursales = suc_opt
    st.session_state._reset_filtros = False

barra_filtros = st.container(border=True)
with barra_filtros:
    fc0, fc1, fc2, fc3, fc4 = st.columns([0.9, 1.3, 1.3, 1.7, 1])
    with fc0:
        st.markdown(
            "<div style='padding-top:8px; font-weight:700; color:#475569;'>🎛️ Filtros</div>",
            unsafe_allow_html=True,
        )
    with fc1:
        with st.popover(
            f"📅 Año ({len(st.session_state.filtro_anio)}/{len(anios_tot)})",
            use_container_width=True,
        ):
            anios_sel = st.multiselect(
                "Año", options=anios_tot, key="filtro_anio", label_visibility="collapsed"
            )
    with fc2:
        with st.popover(
            f"📅 Meses ({len(st.session_state.filtro_meses)}/{len(meses_totales)})",
            use_container_width=True,
        ):
            meses_sel = st.multiselect(
                "Meses", options=meses_totales, key="filtro_meses", label_visibility="collapsed"
            )
    with fc3:
        with st.popover(
            f"🏢 Sucursales ({len(st.session_state.filtro_sucursales)}/{len(suc_opt)})",
            use_container_width=True,
        ):
            suc_sel = st.multiselect(
                "Sucursales", options=suc_opt, key="filtro_sucursales", label_visibility="collapsed"
            )
    with fc4:
        st.markdown("<div style='padding-top:2px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Limpiar", use_container_width=True):
            st.session_state._reset_filtros = True
            st.rerun()

anios_sel = st.session_state.filtro_anio
meses_sel = st.session_state.filtro_meses
suc_sel = st.session_state.filtro_sucursales


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

# Generación de la presentación PowerPoint (gráficos nativos y editables)
if generar_pptx_click:
    _t1 = datetime.datetime.now()
    try:
        with st.spinner(f"Generando presentación PowerPoint ({len(df_f):,} registros a procesar)..."):
            st.session_state["pptx_bytes"] = generar_reporte_pptx(df_f, anios_sel, meses_sel, suc_sel)
        _seg = (datetime.datetime.now() - _t1).total_seconds()
        with st.sidebar:
            st.success(f"✅ PowerPoint generado en {_seg:.1f} segundos. Descárgalo abajo 👇")
    except ModuleNotFoundError as e:
        with st.sidebar:
            st.error(
                f"❌ Falta una librería para generar el PowerPoint: {e.name}. "
                f"Instálala en tu entorno con:\n\n"
                f"pip install -U python-pptx kaleido"
            )

if "pptx_bytes" in st.session_state:
    with st.sidebar:
        st.download_button(
            label="⬇️ Descargar Presentación PowerPoint",
            data=st.session_state["pptx_bytes"],
            file_name=f"Presentacion_FPA_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )

if "pdf_bytes" in st.session_state:
    with st.sidebar:
        st.download_button(
            label="⬇️ Descargar Reporte PDF",
            data=st.session_state["pdf_bytes"],
            file_name=f"Reporte_Ejecutivo_FPA_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
            mime="application/pdf",
        )

# TAB 1: RESUMEN GENERAL
def render_tab1():
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
def render_tab2():
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
        </div>
    </div>

    </body>
    </html>
    """
    components.html(html_mapa, height=820, scrolling=True)

# TAB 3: ORGANIGRAMA
def render_tab3():
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
                    <div class="monto-head">{get_monto_area('Gerencia General')}</div>
                </div>
                <div class="sub-stem"></div>
                <div class="sub-cards">
                    <div class="child-card"><span class="dept-name">Gerencia General</span><span class="dept-val">{get_monto_area('Gerencia General')}</span></div>
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
def render_tab4():
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
            color="Saldo",
            color_continuous_scale=[COLOR_NAVY, "#7c2d12", COLOR_ROJO],
        )
        fig_suc.update_traces(
            marker_line_color="white", marker_line_width=1.5,
            textfont=dict(size=11, color="#1e293b"),
            hovertemplate="<b>%{y}</b><br>S/ %{x:,.0f}<extra></extra>",
        )
        fig_suc.update_layout(
            height=480, plot_bgcolor="white", coloraxis_showscale=False,
        )
        st.plotly_chart(fig_suc, use_container_width=True)

# TAB 5: DETALLE CECO Y ASIENTOS
def render_tab5():
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
def render_tab6():
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


# ============================================================
# TAB 7: DASHBOARD FINANCIERO (KPIs de Gasto / OPEX consolidados)
# ============================================================
def render_tab7():
    st.markdown(
        '<div class="header-banner"><h1>💰 Dashboard Financiero</h1>'
        '<p>Monitoreo integral de indicadores de Gasto y OPEX. '
        'Basado en el Reporte de Centro de Costos (no incluye Ingresos, ya que esta '
        'base solo registra Gastos). | <b>Junior FP&A</b></p></div>',
        unsafe_allow_html=True,
    )

    if len(df_f) == 0:
        st.warning("⚠️ No hay datos para los filtros seleccionados.")
    else:
        # ---------- Preparación de periodos (Año-Mes) ----------
        df_f7 = df_f.copy()
        if "Año" in df_f7.columns and "Mes" in df_f7.columns:
            df_f7["_Periodo"] = df_f7["Año"].astype(int) * 100 + df_f7["Mes"].astype(int)
            periodos = sorted(df_f7["_Periodo"].unique())
        else:
            df_f7["_Periodo"] = 0
            periodos = [0]

        periodo_actual = periodos[-1] if periodos else None
        periodo_anterior = periodos[-2] if len(periodos) >= 2 else None

        gasto_total_acum = df_f7["Saldo"].sum()
        gasto_mes_actual = (
            df_f7.loc[df_f7["_Periodo"] == periodo_actual, "Saldo"].sum()
            if periodo_actual is not None else gasto_total_acum
        )
        gasto_mes_anterior = (
            df_f7.loc[df_f7["_Periodo"] == periodo_anterior, "Saldo"].sum()
            if periodo_anterior is not None else None
        )
        var_mes_pct = (
            (gasto_mes_actual - gasto_mes_anterior) / abs(gasto_mes_anterior) * 100
            if gasto_mes_anterior not in (None, 0) else None
        )

        gasto_fijos = df_f7.loc[df_f7["Comportamiento"] == "Fijos", "Saldo"].sum()
        gasto_variables = df_f7.loc[df_f7["Comportamiento"] == "Variables", "Saldo"].sum()
        pct_fijos = (gasto_fijos / gasto_total_acum * 100) if gasto_total_acum else 0.0
        pct_variables = (gasto_variables / gasto_total_acum * 100) if gasto_total_acum else 0.0

        n_meses = df_f7["_Periodo"].nunique()
        promedio_mensual = gasto_total_acum / max(n_meses, 1)

        # ---------- FILA 1: KPIs ----------
        k1, k2, k3, k4, k5 = st.columns(5)
        with k1:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Gasto Total Acumulado</div>'
                f'<div class="kpi-value">S/ {gasto_total_acum:,.0f}</div></div>',
                unsafe_allow_html=True,
            )
        with k2:
            if var_mes_pct is not None:
                flecha = "▲" if var_mes_pct >= 0 else "▼"
                badge = "badge-red" if var_mes_pct >= 0 else "badge-green"
                var_html = f'<span class="{badge}">{flecha} {var_mes_pct:.1f}% vs mes ant.</span>'
            else:
                var_html = ""
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Gasto Mes Actual</div>'
                f'<div class="kpi-value">S/ {gasto_mes_actual:,.0f}</div>{var_html}</div>',
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Gastos Fijos</div>'
                f'<div class="kpi-value">S/ {gasto_fijos:,.0f}</div>'
                f'<span class="badge-green">{pct_fijos:.1f}% del total</span></div>',
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Gastos Variables</div>'
                f'<div class="kpi-value">S/ {gasto_variables:,.0f}</div>'
                f'<span class="badge-red">{pct_variables:.1f}% del total</span></div>',
                unsafe_allow_html=True,
            )
        with k5:
            st.markdown(
                f'<div class="kpi-card"><div class="kpi-title">Promedio Mensual</div>'
                f'<div class="kpi-value">S/ {promedio_mensual:,.0f}</div>'
                f'<span class="badge-green">{n_meses} mes(es) con datos</span></div>',
                unsafe_allow_html=True,
            )

        st.markdown("---")

        # ---------- FILA 2: Evolución mensual + Top CECO ----------
        col1, col2 = st.columns([1.4, 1])
        df_evol = (
            df_f7.groupby(["Año", "Mes"])["Saldo"].sum().reset_index().sort_values(["Año", "Mes"])
            if "Año" in df_f7.columns else pd.DataFrame(columns=["Año", "Mes", "Saldo"])
        )
        if len(df_evol):
            df_evol["Periodo"] = (
                df_evol["Año"].astype(int).astype(str) + "-" + df_evol["Mes"].astype(int).astype(str).str.zfill(2)
            )

        with col1:
            chart_box1 = st.container(border=True)
            with chart_box1:
                st.subheader("📈 Evolución Mensual del Gasto")
                if len(df_evol) >= 1:
                    promedio_linea = df_evol["Saldo"].mean()
                    max_idx = df_evol["Saldo"].idxmax()
                    fig_evol = go.Figure()
                    # Área con degradé carmesí bajo la línea de tendencia — mucho más
                    # dramático que una barra plana, y no usa ningún azul.
                    fig_evol.add_trace(go.Scatter(
                        x=df_evol["Periodo"], y=df_evol["Saldo"],
                        mode="lines+markers", name="Gasto Mensual",
                        line=dict(color=COLOR_ROJO, width=4, shape="spline"),
                        marker=dict(size=9, color=COLOR_NAVY, line=dict(color="white", width=2)),
                        fill="tozeroy", fillcolor="rgba(168, 17, 33, 0.18)",
                        hovertemplate="<b>%{x}</b><br>S/ %{y:,.0f}<extra></extra>",
                    ))
                    fig_evol.add_hline(
                        y=promedio_linea, line_dash="dot", line_color="#d97706", line_width=2,
                        annotation_text=f"Promedio S/ {promedio_linea:,.0f}",
                        annotation_font=dict(size=11, color="#b45309"),
                    )
                    # Resalta el mes pico con una anotación llamativa
                    fig_evol.add_annotation(
                        x=df_evol.loc[max_idx, "Periodo"], y=df_evol.loc[max_idx, "Saldo"],
                        text="Pico ▲", showarrow=True, arrowhead=2, arrowcolor=COLOR_NAVY,
                        font=dict(color=COLOR_NAVY, size=11, family="Arial Black"),
                        ay=-35,
                    )
                    fig_evol.update_layout(
                        height=380, plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                        showlegend=False,
                        font=dict(family="Arial", size=12, color="#1e293b"),
                        margin=dict(t=30, b=10, l=10, r=10),
                        xaxis=dict(showgrid=False, tickfont=dict(size=11)),
                        yaxis=dict(showgrid=True, gridcolor="#f1f5f9", tickfont=dict(size=11)),
                    )
                    st.plotly_chart(fig_evol, use_container_width=True)
                else:
                    st.info("No hay suficientes periodos (Año/Mes) para graficar la evolución.")

        with col2:
            chart_box2 = st.container(border=True)
            with chart_box2:
                st.subheader("🏢 Gasto por Centro de Costo (Top 8)")
                df_top_ceco = (
                    df_f7.groupby("Centro de Costos")["Saldo"].sum()
                    .reset_index().sort_values("Saldo", ascending=False).head(8)
                )
                df_top_ceco_plot = df_top_ceco.sort_values("Saldo")
                fig_ceco = px.bar(
                    df_top_ceco_plot,
                    x="Saldo", y="Centro de Costos", orientation="h",
                    text_auto=".2s",
                    color="Saldo",
                    color_continuous_scale=[COLOR_NAVY, "#7c2d12", COLOR_ROJO],
                )
                fig_ceco.update_traces(
                    marker_line_color="white", marker_line_width=1.5,
                    textfont=dict(size=11, color="#1e293b"),
                    hovertemplate="<b>%{y}</b><br>S/ %{x:,.0f}<extra></extra>",
                )
                fig_ceco.update_layout(
                    height=380, plot_bgcolor="white", paper_bgcolor="rgba(0,0,0,0)",
                    showlegend=False, coloraxis_showscale=False,
                    font=dict(family="Arial", size=12, color="#1e293b"),
                    margin=dict(t=20, b=10, l=10, r=10),
                    xaxis=dict(showgrid=True, gridcolor="#f1f5f9"),
                    yaxis=dict(showgrid=False),
                )
                st.plotly_chart(fig_ceco, use_container_width=True)

        st.markdown("---")

        # ---------- FILA 3: Variaciones + Composición + Top Desviaciones ----------
        col3, col4, col5 = st.columns([1.2, 1, 1.1])

        with col3:
            chart_box3 = st.container(border=True)
            with chart_box3:
                st.subheader("📋 Variación por Clasificación")
                if periodo_anterior is not None:
                    df_act = df_f7[df_f7["_Periodo"] == periodo_actual].groupby("Clasificación")["Saldo"].sum()
                    df_ant = df_f7[df_f7["_Periodo"] == periodo_anterior].groupby("Clasificación")["Saldo"].sum()
                    df_var = pd.DataFrame({"Mes Actual": df_act, "Mes Anterior": df_ant}).fillna(0)
                    df_var["Var S/"] = df_var["Mes Actual"] - df_var["Mes Anterior"]
                    df_var["Var %"] = np.where(
                        df_var["Mes Anterior"] != 0,
                        df_var["Var S/"] / df_var["Mes Anterior"].abs() * 100,
                        0.0,
                    )
                    df_var = df_var.reset_index().rename(columns={"index": "Clasificación"})
                    df_var = df_var.sort_values("Var S/", key=lambda s: s.abs(), ascending=False)
                    st.dataframe(
                        df_var,
                        use_container_width=True, height=320, hide_index=True,
                        column_config={
                            "Mes Actual": st.column_config.NumberColumn(format="S/ %.0f"),
                            "Mes Anterior": st.column_config.NumberColumn(format="S/ %.0f"),
                            "Var S/": st.column_config.NumberColumn(format="S/ %.0f"),
                            "Var %": st.column_config.ProgressColumn(
                                format="%+.1f%%",
                                min_value=float(df_var["Var %"].min()) if len(df_var) else -100,
                                max_value=float(df_var["Var %"].max()) if len(df_var) else 100,
                            ),
                        },
                    )
                else:
                    st.info("Se necesitan al menos 2 periodos (Año-Mes) para comparar variación.")

        with col4:
            chart_box4 = st.container(border=True)
            with chart_box4:
                st.subheader("🍩 Composición del Gasto")
                df_comp7 = (
                    df_f7.groupby("Clasificación")["Saldo"].sum()
                    .reset_index().sort_values("Saldo", ascending=False)
                )
                pulls = [0.06 if i == 0 else 0 for i in range(len(df_comp7))]
                fig_comp7 = go.Figure(data=[go.Pie(
                    labels=df_comp7["Clasificación"], values=df_comp7["Saldo"],
                    hole=0.55, pull=pulls,
                    marker=dict(colors=PALETA_FINANCIERO, line=dict(color="white", width=2)),
                    textinfo="percent", textfont=dict(size=12, color="white"),
                    hovertemplate="<b>%{label}</b><br>S/ %{value:,.0f} (%{percent})<extra></extra>",
                )])
                fig_comp7.update_layout(
                    height=320, paper_bgcolor="rgba(0,0,0,0)",
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="top", y=-0.05, font=dict(size=10)),
                    margin=dict(t=20, b=10, l=10, r=10),
                    annotations=[dict(
                        text=f"<b>S/ {gasto_total_acum:,.0f}</b>", x=0.5, y=0.5,
                        font=dict(size=13, color=COLOR_NAVY), showarrow=False,
                    )],
                )
                st.plotly_chart(fig_comp7, use_container_width=True)

        with col5:
            chart_box5 = st.container(border=True)
            with chart_box5:
                st.subheader("🚨 Top 5 Desviaciones del Mes")
                if periodo_anterior is not None:
                    df_cta_act = df_f7[df_f7["_Periodo"] == periodo_actual].groupby("Cuenta Contable")["Saldo"].sum()
                    df_cta_ant = df_f7[df_f7["_Periodo"] == periodo_anterior].groupby("Cuenta Contable")["Saldo"].sum()
                    df_dev = pd.DataFrame({"Actual": df_cta_act, "Anterior": df_cta_ant}).fillna(0)
                    df_dev["Var"] = df_dev["Actual"] - df_dev["Anterior"]
                    top5 = df_dev.reset_index().rename(columns={"index": "Cuenta Contable"})
                    top5 = top5.sort_values("Var", key=lambda s: s.abs(), ascending=False).head(5)
                    if len(top5):
                        max_abs = top5["Var"].abs().max() or 1
                        for _, r in top5.iterrows():
                            es_alza = r["Var"] >= 0
                            flecha = "▲" if es_alza else "▼"
                            color = COLOR_ROJO if es_alza else "#0f766e"
                            ancho_barra = max(6, abs(r["Var"]) / max_abs * 100)
                            st.markdown(
                                f'<div style="padding:7px 2px;border-bottom:1px solid #e2e8f0;">'
                                f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                                f'<span style="font-size:12.5px;color:#334155;">{_acortar(r["Cuenta Contable"], 26)}</span>'
                                f'<span style="font-weight:800;color:{color};font-size:12.5px;">{flecha} S/ {abs(r["Var"]):,.0f}</span>'
                                f'</div>'
                                f'<div style="background:#f1f5f9;border-radius:6px;height:6px;margin-top:4px;">'
                                f'<div style="width:{ancho_barra:.0f}%;background:{color};height:6px;border-radius:6px;"></div>'
                                f'</div></div>',
                                unsafe_allow_html=True,
                            )
                    else:
                        st.info("Sin variaciones para mostrar.")
                else:
                    st.info("Se necesita un mes anterior para calcular desviaciones.")

        st.markdown("---")

        # ---------- FILA 4: Proyección de cierre + Semáforo ----------
        col6, col7 = st.columns([1.3, 1])

        with col6:
            chart_box6 = st.container(border=True)
            with chart_box6:
                st.subheader("📅 Proyección de Cierre (Run-Rate)")
                if len(df_evol) >= 2:
                    promedio_reciente = df_evol.tail(3)["Saldo"].mean()
                    var_vs_prom = (
                        (gasto_mes_actual - promedio_mensual) / abs(promedio_mensual) * 100
                        if promedio_mensual else 0.0
                    )
                    pc1, pc2, pc3 = st.columns(3)
                    with pc1:
                        st.metric("Gasto Mes Actual", f"S/ {gasto_mes_actual:,.0f}")
                    with pc2:
                        st.metric("Promedio Últimos 3 Meses", f"S/ {promedio_reciente:,.0f}")
                    with pc3:
                        st.metric(
                            "Proyección Cierre (Run-Rate)",
                            f"S/ {promedio_reciente:,.0f}",
                            delta=f"{var_vs_prom:+.1f}% vs prom. histórico",
                            delta_color="inverse",
                        )
                    st.caption(
                        "Proyección basada en el promedio de gasto de los últimos 3 meses con datos "
                        "(run-rate simple). No sustituye un forecast presupuestal formal."
                    )
                else:
                    st.info("Se necesitan al menos 2 meses con datos para proyectar el cierre.")

        with col7:
            chart_box7 = st.container(border=True)
            with chart_box7:
                st.subheader("🚦 Semáforo de la Operación")

                def _semaforo(cond_verde, cond_amarillo):
                    if cond_verde:
                        return "🟢"
                    elif cond_amarillo:
                        return "🟡"
                    return "🔴"

                n_alertas_ia = len(detectar_alertas_ia(df_f7, umbral_residual=1000, umbral_z=2.0))
                top_ceco_pct = (
                    (df_top_ceco["Saldo"].max() / gasto_total_acum * 100)
                    if gasto_total_acum and len(df_top_ceco) else 0.0
                )

                items_semaforo = [
                    ("Gasto vs Mes Anterior", _semaforo(
                        var_mes_pct is not None and var_mes_pct <= 0,
                        var_mes_pct is not None and var_mes_pct <= 10,
                    )),
                    ("Gasto Mes Actual vs Promedio", _semaforo(
                        gasto_mes_actual <= promedio_mensual,
                        gasto_mes_actual <= promedio_mensual * 1.1,
                    )),
                    ("Peso de Gastos Fijos", _semaforo(pct_fijos <= 60, pct_fijos <= 75)),
                    ("Alertas IA Activas", _semaforo(n_alertas_ia == 0, n_alertas_ia <= 5)),
                    ("Concentración en Top CECO", _semaforo(top_ceco_pct <= 30, top_ceco_pct <= 50)),
                ]
                for nombre, icono in items_semaforo:
                    st.markdown(
                        f'<div style="display:flex;justify-content:space-between;align-items:center;'
                        f'padding:10px 4px;border-bottom:1px solid #e2e8f0;">'
                        f'<span style="font-size:13px;color:#334155;font-weight:600;">{nombre}</span>'
                        f'<span style="font-size:18px;">{icono}</span></div>',
                        unsafe_allow_html=True,
                    )


# 5. RENDER DE LA PÁGINA SELECCIONADA
if pagina == "Resumen General":
    render_tab1()
elif pagina == "Mapa de Procesos":
    render_tab2()
elif pagina == "Organigrama por Áreas":
    render_tab3()
elif pagina == "Mapa Sucursales Perú":
    render_tab4()
elif pagina == "Detalle por CECO / Asientos":
    render_tab5()
elif pagina == "Centro de Alertas":
    render_tab6()
elif pagina == "Dashboard Financiero":
    render_tab7()
