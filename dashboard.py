import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from PIL import Image
from fpdf import FPDF
import tempfile
import os
from pathlib import Path

# --- CONFIGURACIÓN Y CONSTANTES ---
pio.templates.default = "plotly"
COLOR_PALETTE = px.colors.qualitative.Plotly
CURRENT_DIR = Path(__file__).parent
BANNER_PATH = str(CURRENT_DIR / "image.png")

LOGOS = {
    "COSEDUCAM S A": str(CURRENT_DIR / "coseducam.png"),
    "M&Q SPA": str(CURRENT_DIR / "mq.png"),
    "M S & D SPA": str(CURRENT_DIR / "msd.png"),
    "JORQUERA TRANSPORTE S. A.": str(CURRENT_DIR / "jorquera.png"),
    "AG SERVICES SPA": str(CURRENT_DIR / "ag.png"),
}

# --- FUNCIONES DE APOYO ---
def normalizar_nombre_empresa(nombre):
    nombre = str(nombre).strip().upper()
    nombre = nombre.replace('.', '').replace('&', 'AND')
    nombre = ' '.join(nombre.split())
    equivalencias = {
        "JORQUERA TRANSPORTE S A": "JORQUERA TRANSPORTE S. A.",
        "MINING SERVICES AND DERIVATES": "M S & D SPA",
        "MINING SERVICES AND DERIVATES SPA": "M S & D SPA",
        "M S AND D": "M S & D SPA",
        "M S AND D SPA": "M S & D SPA",
        "MSANDD SPA": "M S & D SPA",
        "MSANDD": "M S & D SPA",
        "M S D": "M S & D SPA",
        "M S D SPA": "M S & D SPA",
        "M S & D": "M S & D SPA",
        "M S & D SPA": "M S & D SPA",
        "MS&D SPA": "M S & D SPA",
        "M AND Q SPA": "M&Q SPA",
        "M AND Q": "M&Q SPA",
        "M Q SPA": "M&Q SPA",
        "MQ SPA": "M&Q SPA",
        "M&Q SPA": "M&Q SPA",
        "MANDQ SPA": "M&Q SPA",
        "MANDQ": "M&Q SPA",
        "MINING AND QUARRYING SPA": "M&Q SPA",
        "MINING AND QUARRYNG SPA": "M&Q SPA",
        "AG SERVICE SPA": "AG SERVICES SPA",
        "COSEDUCAM S A": "COSEDUCAM S A",
        "COSEDUCAM": "COSEDUCAM S A"
    }
    return equivalencias.get(nombre, nombre)

def clean_text(text):
    """Limpia texto para evitar errores de encoding en FPDF (latin-1)"""
    return str(text).encode('latin-1', 'replace').decode('latin-1')

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Dashboard Equipos", layout="wide")
st.title("Dashboard: Equipos por Hora y Empresa")

uploaded_file = st.file_uploader("Carga tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        required_columns = {'fecha_col': 0, 'destino_col': 3, 'empresa_col': 11, 'hora_col': 14}
        
        if len(df.columns) < 15:
            st.error("El archivo no tiene suficientes columnas.")
            st.stop()

        fecha_col, destino_col, empresa_col, hora_col = [df.columns[i] for i in [0, 3, 11, 14]]
        df = df.dropna(subset=[fecha_col, destino_col, empresa_col, hora_col])
        
        df[fecha_col] = pd.to_datetime(df[fecha_col], errors='coerce', dayfirst=True)
        df[hora_col] = pd.to_datetime(df[hora_col], format='%H:%M:%S', errors='coerce').dt.hour
        df[empresa_col] = df[empresa_col].apply(normalizar_nombre_empresa)

        # Filtros
        fechas = df[fecha_col].dropna().dt.date.unique()
        fecha_sel = st.date_input("Fecha:", value=min(fechas))
        df_filtered = df[df[fecha_col].dt.date == fecha_sel]
        
        empresas = sorted(df_filtered[empresa_col].unique())
        empresas_sel = st.multiselect("Empresas:", empresas, default=empresas)
        
        destinos = sorted(df_filtered[destino_col].unique())
        destinos_sel = st.multiselect("Destinos:", destinos, default=destinos)

        # Labels de horas fijos
        horas_labels = [f"{str(h).zfill(2)}:00 - {str(h).zfill(2)}:59" for h in range(24)]

        for empresa in empresas_sel:
            df_empresa = df_filtered[(df_filtered[empresa_col] == empresa) & (df_filtered[destino_col].isin(destinos_sel))]
            st.markdown(f"## Empresa: {empresa}")
            
            if df_empresa.empty:
                st.info(f"Sin datos para {empresa}")
                continue

            col1, col2 = st.columns(2)
            
            # Gráfico Plotly
            resumen = df_empresa.groupby([hora_col, destino_col]).size().reset_index(name='Cantidad')
            fig = px.line(resumen, x=hora_col, y="Cantidad", color=destino_col, markers=True, title=f"Equipos por Hora - {empresa}")
            fig.update_layout(xaxis=dict(dtick=1))
            col1.plotly_chart(fig, use_container_width=True)

            # Tabla resumen
            df_empresa['Hora Intervalo'] = df_empresa[hora_col].apply(lambda h: f"{str(int(h)).zfill(2)}:00 - {str(int(h)).zfill(2)}:59")
            tabla = pd.pivot_table(df_empresa, index='Hora Intervalo', columns=destino_col, values=empresa_col, aggfunc='count', fill_value=0)
            tabla = tabla.reindex(horas_labels, fill_value=0)
            tabla_final = pd.concat([tabla, pd.DataFrame(tabla.sum(axis=0), columns=['TOTAL']).T])
            col2.dataframe(tabla_final)

            # --- GENERACIÓN DE PDF ---
            with tempfile.TemporaryDirectory() as tmpdir:
                # 1. Imagen del gráfico
                img_path = os.path.join(tmpdir, "chart.png")
                fig.write_image(img_path, scale=2)

                # 2. Combinar Imágenes (Banner + Logo + Gráfico)
                imgs = []
                if os.path.exists(BANNER_PATH): imgs.append(Image.open(BANNER_PATH).convert("RGB"))
                
                logo_p = LOGOS.get(empresa)
                if logo_p and os.path.exists(logo_p):
                    l_img = Image.open(logo_p).convert("RGBA")
                    # Redimensionar logo a ancho 150px
                    l_w = 150
                    l_h = int(l_img.size[1] * (l_w / l_img.size[0]))
                    l_res = l_img.resize((l_w, l_h), Image.LANCZOS)
                    # Fondo blanco para el logo
                    bg = Image.new("RGB", (800, l_h), (255, 255, 255))
                    bg.paste(l_res, ((800 - l_w)//2, 0), l_res)
                    imgs.append(bg)
                
                imgs.append(Image.open(img_path).convert("RGB"))

                # Stack vertical
                w = max(i.width for i in imgs)
                h_total = sum(i.height for i in imgs)
                combined = Image.new("RGB", (w, h_total), (255, 255, 255))
                curr_h = 0
                for i in imgs:
                    combined.paste(i, (0, curr_h))
                    curr_h += i.height
                
                combined_path = os.path.join(tmpdir, "final.png")
                combined.save(combined_path)

                # 3. PDF
                pdf = FPDF(orientation='L', unit='mm', format='A4')
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, clean_text(f"Empresa: {empresa} - Fecha: {fecha_sel}"), ln=1, align="C")
                pdf.image(combined_path, x=10, y=25, w=275)

                # Tabla en pag 2
                pdf.add_page(orientation='P')
                pdf.set_font("Arial", "B", 12)
                pdf.cell(0, 10, "Resumen de Cantidades", ln=1, align="C")
                pdf.set_font("Arial", "", 7)
                
                # Encabezados de tabla
                c_w = 190 / (len(tabla_final.columns) + 1)
                pdf.cell(c_w, 8, "Hora", 1, 0, "C")
                for c in tabla_final.columns: pdf.cell(c_w, 8, clean_text(c[:15]), 1, 0, "C")
                pdf.ln()
                
                # Filas de tabla
                for idx, row in tabla_final.iterrows():
                    pdf.cell(c_w, 6, clean_text(idx), 1, 0, "C")
                    for val in row: pdf.cell(c_w, 6, str(int(val)), 1, 0, "C")
                    pdf.ln()

                pdf_output = pdf.output(dest='S').encode('latin-1')
                
                # BOTÓN ÚNICO
                st.download_button(
                    label=f"📥 Descargar Reporte {empresa}",
                    data=pdf_output,
                    file_name=f"Reporte_{empresa.replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    key=f"dl_btn_{empresa.replace('&', 'AND').replace(' ', '_')}"
                )

    except Exception as e:
        st.error(f"Error general: {e}")
else:
    st.info("Por favor, carga un archivo Excel.")
