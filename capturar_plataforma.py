# -*- coding: utf-8 -*-
"""
capturar_plataforma.py
-----------------------
Recorre tu app de Streamlit ya desplegada (la abre en un navegador real),
toma una captura de pantalla de cada sección del menú, y además hace clic
en "Generar PDF" y "Generar PowerPoint" para descargar esos archivos.

Todo queda guardado en una carpeta "capturas_automaticas/" lista para
usarse con el script generar_video.py que ya tienes.

INSTALACIÓN (una sola vez, en tu PC — requiere internet):
    pip install playwright
    playwright install chromium

USO:
    python capturar_plataforma.py

Si tu app pide un usuario/contraseña de Streamlit Cloud (apps privadas),
avísame y le agrego el login; si es pública como la que compartiste, no
hace falta nada más.
"""

import os
import time
from playwright.sync_api import sync_playwright

# ============================================================
# CONFIGURACIÓN — edita esto si algo cambia
# ============================================================
URL_APP = "https://finanzas-oun7ivsfcju26mjgrgxwaj.streamlit.app/"
CARPETA_SALIDA = "capturas_automaticas"

# Debe coincidir EXACTO con el texto de los botones del menú en app.py
SECCIONES = [
    "📊  Resumen General",
    "🗺️  Mapa de Procesos",
    "🏛️  Organigrama por Áreas",
    "🌎  Mapa Sucursales Perú",
    "🔍  Detalle por CECO / Asientos",
    "🚨  Centro de Alertas",
    "💰  Dashboard Financiero",
]

BOTON_PDF = "📄 Generar Reporte PDF Ejecutivo"
BOTON_PPTX = "📊 Generar Presentación PowerPoint"
DESCARGA_PDF = "⬇️ Descargar Reporte PDF"
DESCARGA_PPTX = "⬇️ Descargar Presentación PowerPoint"

ESPERA_RENDER_SEG = 3       # tiempo a esperar tras cambiar de sección
ESPERA_GENERACION_SEG = 30  # tiempo máximo a esperar mientras se genera el PDF/PPTX


def nombre_archivo(texto: str) -> str:
    limpio = "".join(c for c in texto if c.isalnum() or c in " _-").strip()
    return limpio.replace(" ", "_").lower()


def main():
    os.makedirs(CARPETA_SALIDA, exist_ok=True)

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=False)  # headless=True si no quieres ver la ventana
        pagina = navegador.new_page(viewport={"width": 1600, "height": 1000})

        print(f"Abriendo {URL_APP} ...")
        pagina.goto(URL_APP, wait_until="networkidle", timeout=60000)
        time.sleep(4)  # Streamlit Cloud a veces tarda en "despertar" la app dormida
        pagina.screenshot(path=os.path.join(CARPETA_SALIDA, "00_carga_inicial.png"), full_page=True)

        # --- 1) Recorrer cada sección del menú y capturar ---
        for i, nombre_boton in enumerate(SECCIONES, start=1):
            try:
                boton = pagina.get_by_role("button", name=nombre_boton, exact=False)
                boton.click(timeout=10000)
                time.sleep(ESPERA_RENDER_SEG)
                ruta = os.path.join(CARPETA_SALIDA, f"{i:02d}_{nombre_archivo(nombre_boton)}.png")
                pagina.screenshot(path=ruta, full_page=True)
                print(f"  ✅ Captura: {ruta}")
            except Exception as e:
                print(f"  ⚠️ No se pudo capturar '{nombre_boton}': {e}")

        # Volver a la primera sección antes de generar reportes
        try:
            pagina.get_by_role("button", name=SECCIONES[0], exact=False).click(timeout=10000)
            time.sleep(ESPERA_RENDER_SEG)
        except Exception:
            pass

        # --- 2) Generar y descargar el PDF ---
        try:
            print("Generando PDF...")
            pagina.get_by_role("button", name=BOTON_PDF, exact=False).click(timeout=10000)
            with pagina.expect_download(timeout=ESPERA_GENERACION_SEG * 1000) as descarga_info:
                pagina.get_by_role("button", name=DESCARGA_PDF, exact=False).click(timeout=ESPERA_GENERACION_SEG * 1000)
            descarga = descarga_info.value
            ruta_pdf = os.path.join(CARPETA_SALIDA, "reporte_ejecutivo.pdf")
            descarga.save_as(ruta_pdf)
            print(f"  ✅ PDF descargado: {ruta_pdf}")
        except Exception as e:
            print(f"  ⚠️ No se pudo generar/descargar el PDF: {e}")

        # --- 3) Generar y descargar el PowerPoint ---
        try:
            print("Generando PowerPoint...")
            pagina.get_by_role("button", name=BOTON_PPTX, exact=False).click(timeout=10000)
            with pagina.expect_download(timeout=ESPERA_GENERACION_SEG * 1000) as descarga_info:
                pagina.get_by_role("button", name=DESCARGA_PPTX, exact=False).click(timeout=ESPERA_GENERACION_SEG * 1000)
            descarga = descarga_info.value
            ruta_pptx = os.path.join(CARPETA_SALIDA, "presentacion.pptx")
            descarga.save_as(ruta_pptx)
            print(f"  ✅ PowerPoint descargado: {ruta_pptx}")
        except Exception as e:
            print(f"  ⚠️ No se pudo generar/descargar el PowerPoint: {e}")

        navegador.close()

    print(f"\nListo. Revisa la carpeta: {CARPETA_SALIDA}/")


if __name__ == "__main__":
    main()
