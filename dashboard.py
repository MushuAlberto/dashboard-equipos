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
    """Limpia texto para evitar errores de encoding en FPDF"""
    try:
        return str(text).encode('latin-1', 'replace').decode('latin-1')
    except:
        return str(text)

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Dashboard Equipos", layout="wide")
st.title("Dashboard: Equipos por Hora y Empresa")

uploaded_file = st.file_uploader("Carga tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        
        # Identificar columnas por posición
        fecha_col = df.columns[0]
        destino_col = df.columns[3]
        empresa_col = df.columns[11]
        hora_col = df.columns[14]
        
        df = df.dropna(subset=[fecha_col, destino_col, empresa_col, hora_col])
        
        df[fecha_col] = pd.to_datetime(df[fecha_col], errors='coerce', dayfirst=True)
        df[hora_col] = pd.to_datetime(df[hora_col], format='%H:%M:%S', errors='coerce').dt.hour
        df[empresa_col] = df[empresa_col].apply(normalizar_nombre_empresa)

        # Filtros principales
        fechas = sorted(df[fecha_col].dropna().dt.date.unique())
        fecha_sel = st.date_input("Selecciona Fecha:", value=fechas[0] if fechas else None)
        
        df_filtered = df[df[fecha_col].dt.date == fecha_sel]
        
        empresas = sorted(df_filtered[empresa_col].unique())
        empresas_sel = st.multiselect("Filtrar Empresas:", empresas, default=empresas)
        
        destinos = sorted(df_filtered[destino_col].unique())
        destinos_sel = st.multiselect("Filtrar Destinos:", destinos, default=destinos)

        horas_labels = [f"{str(h).zfill(2)}:00 - {str(h).zfill(2)}:59" for h in range(24)]

        # BUCLE POR CADA EMPRESA
        for empresa in empresas_sel:
            # Filtrar datos de la empresa actual
            df_empresa = df_filtered[(df_filtered[empresa_col] == empresa) & (df_filtered[destino_col].isin(destinos_sel))]
            
            if df_empresa.empty:
                continue

            st.markdown(f"---")
            st.header(f"Empresa: {empresa}")
            
            col1, col2 = st.columns([3, 2])
            
            # 1. Gráfico
            resumen = df_empresa.groupby([hora_col, destino_col]).size().reset_index(name='Cantidad')
            fig = px.line(resumen, x=hora_col, y="Cantidad", color=destino_col, markers=True, 
                         title=f"Equipos por Hora - {empresa}",
                         labels={hora_col: "HR ENTRADA", "Cantidad": "Cantidad"})
            fig.update_layout(xaxis=dict(dtick=1, range=[0, 23]))
            col1.plotly_chart(fig, use_container_width=True)

            # 2. Tabla
            df_empresa['Hora Intervalo'] = df_empresa[hora_col].apply(lambda h: f"{str(int(h)).zfill(2)}:00 - {str(int(h)).zfill(2)}:59")
            tabla = pd.pivot_table(df_empresa, index='Hora Intervalo', columns=destino_col, values=empresa_col, aggfunc='count', fill_value=0)
            tabla = tabla.reindex(horas_labels, fill_value=0)
            
            # Sumatoria TOTAL
            sumatoria = pd.DataFrame(tabla.sum(axis=0)).T
            sumatoria.index = ['TOTAL']
            tabla_final = pd.concat([tabla, sumatoria])
            col2.write("Resumen Numérico")
            col2.dataframe(tabla_final, use_container_width=True)

            # 3. Preparación de PDF (dentro de un try para que si falla una empresa no rompa el resto)
            try:
                with tempfile.TemporaryDirectory() as tmpdir:
                    # Guardar gráfico como imagen
                    img_path = os.path.join(tmpdir, "chart.png")
                    fig.write_image(img_path, scale=2)

                    # Crear PDF
                    pdf = FPDF(orientation='L', unit='mm', format='A4')
                    pdf.add_page()
                    pdf.set_font("Arial", "B", 16)
                    pdf.cell(0, 10, clean_text(f"Reporte: {empresa} | Fecha: {fecha_sel}"), ln=1, align="C")
                    
                    # Insertar gráfico
                    pdf.image(img_path, x=15, y=30, w=260)

                    # Nueva página para la tabla
                    pdf.add_page(orientation='P')
                    pdf.set_font("Arial", "B", 12)
                    pdf.cell(0, 10, "Detalle de Cantidades por Destino", ln=1, align="C")
                    pdf.ln(5)
                    pdf.set_font("Arial", "", 8)

                    # Dibujar tabla en PDF
                    ancho_col = 190 / (len(tabla_final.columns) + 1)
                    pdf.cell(ancho_col, 8, "Hora", 1, 0, "C")
                    for c in tabla_final.columns:
                        pdf.cell(ancho_col, 8, clean_text(str(c)[:15]), 1, 0, "C")
                    pdf.ln()

                    for idx, row in tabla_final.iterrows():
                        pdf.cell(ancho_col, 7, clean_text(idx), 1, 0, "C")
                        for val in row:
                            pdf.cell(ancho_col, 7, str(int(val)), 1, 0, "C")
                        pdf.ln()

                    # Generar salida de bytes de forma segura
                    pdf_bytes = pdf.output(dest='S')
                    if isinstance(pdf_bytes, str): # Por si devuelve string en versiones viejas
                        pdf_bytes = pdf_bytes.encode('latin-1')

                    st.download_button(
                        label=f"📥 Descargar PDF {empresa}",
                        data=pdf_bytes,
                        file_name=f"Reporte_{empresa.replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        key=f"btn_{empresa.replace('&', 'AND').replace(' ', '_')}"
                    )
            except Exception as pdf_err:
                st.warning(f"No se pudo generar el PDF para {empresa}: {pdf_err}")

    except Exception as e:
        st.error(f"Error general en el proceso: {e}")
else:
    st.info("Carga un archivo Excel para comenzar.")
