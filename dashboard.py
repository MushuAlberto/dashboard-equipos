import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from PIL import Image
from fpdf import FPDF
import tempfile
import os
from pathlib import Path

# --- CONFIGURACIÓN ---
pio.templates.default = "plotly"
CURRENT_DIR = Path(__file__).parent

def normalizar_nombre_empresa(nombre):
    nombre = str(nombre).strip().upper().replace('.', '').replace('&', 'AND')
    nombre = ' '.join(nombre.split())
    equivalencias = {
        "JORQUERA TRANSPORTE S A": "JORQUERA TRANSPORTE S. A.",
        "MINING SERVICES AND DERIVATES SPA": "M S & D SPA",
        "M AND Q SPA": "M&Q SPA",
        "M&Q SPA": "M&Q SPA",
        "AG SERVICE SPA": "AG SERVICES SPA",
        "COSEDUCAM": "COSEDUCAM S A"
    }
    return equivalencias.get(nombre, nombre)

def clean_text(text):
    return str(text).encode('latin-1', 'replace').decode('latin-1')

# FUNCIÓN MAESTRA PARA GENERAR PDF
def crear_pdf_empresa(empresa, df_empresa, fig, fecha_sel, tabla_final):
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    
    # Usar un archivo temporal para la imagen del gráfico
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
        fig.write_image(tmp_img.name, format="png", width=1000, height=500, scale=2)
        img_path = tmp_img.name

    try:
        # PÁGINA 1: GRÁFICO
        pdf.add_page()
        pdf.set_font("Arial", "B", 16)
        pdf.cell(0, 10, clean_text(f"REPORTE: {empresa}"), ln=1, align="C")
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 10, clean_text(f"Fecha: {fecha_sel}"), ln=1, align="C")
        pdf.image(img_path, x=10, y=35, w=275)

        # PÁGINA 2: TABLA
        pdf.add_page(orientation='P')
        pdf.set_font("Arial", "B", 14)
        pdf.cell(0, 10, "Resumen de Cantidades", ln=1, align="C")
        pdf.ln(5)
        
        pdf.set_font("Arial", "", 8)
        ancho_col = 190 / (len(tabla_final.columns) + 1)
        
        # Encabezados
        pdf.cell(ancho_col, 8, "Hora", 1, 0, "C")
        for c in tabla_final.columns:
            pdf.cell(ancho_col, 8, clean_text(str(c)[:12]), 1, 0, "C")
        pdf.ln()

        # Datos
        for idx, row in tabla_final.iterrows():
            pdf.cell(ancho_col, 7, clean_text(idx), 1, 0, "C")
            for val in row:
                pdf.cell(ancho_col, 7, str(int(val)), 1, 0, "C")
            pdf.ln()

        return bytes(pdf.output(dest='S'))
    finally:
        if os.path.exists(img_path):
            os.remove(img_path)

# --- APP PRINCIPAL ---
st.set_page_config(page_title="Dashboard Logística", layout="wide")
st.title("Generador de Reportes de Equipos")

uploaded_file = st.file_uploader("Cargar Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # Ajuste de columnas según tu descripción (A=0, D=3, L=11, O=14)
    f_col, d_col, e_col, h_col = df.columns[0], df.columns[3], df.columns[11], df.columns[14]
    
    df = df.dropna(subset=[f_col, d_col, e_col, h_col])
    df[f_col] = pd.to_datetime(df[f_col], errors='coerce', dayfirst=True)
    df[h_col] = pd.to_datetime(df[h_col], format='%H:%M:%S', errors='coerce').dt.hour
    df[e_col] = df[e_col].apply(normalizar_nombre_empresa)

    fechas = sorted(df[f_col].dropna().dt.date.unique())
    fecha_sel = st.date_input("Fecha", value=fechas[0] if fechas else None)
    
    df_filtered = df[df[f_col].dt.date == fecha_sel]
    empresas_list = sorted(df_filtered[e_col].unique())
    
    for empresa in empresas_list:
        st.write(f"### {empresa}")
        df_emp = df_filtered[df_filtered[e_col] == empresa]
        
        # Preparar datos
        resumen = df_emp.groupby([h_col, d_col]).size().reset_index(name='Cant')
        fig = px.line(resumen, x=h_col, y="Cant", color=d_col, markers=True)
        
        # Preparar Tabla
        labels = [f"{str(h).zfill(2)}:00 - {str(h).zfill(2)}:59" for h in range(24)]
        df_emp['Intervalo'] = df_emp[h_col].apply(lambda x: labels[int(x)])
        tabla = pd.pivot_table(df_emp, index='Intervalo', columns=d_col, values=e_col, aggfunc='count', fill_value=0)
        tabla = tabla.reindex(labels, fill_value=0)
        tabla_final = pd.concat([tabla, pd.DataFrame(tabla.sum(axis=0)).T.rename(index={0:'TOTAL'})])

        col1, col2 = st.columns([2, 1])
        col1.plotly_chart(fig, use_container_width=True)
        col2.dataframe(tabla_final)

        # BOTÓN DE DESCARGA (Genera los bytes solo al ser necesario)
        try:
            pdf_data = crear_pdf_empresa(empresa, df_emp, fig, fecha_sel, tabla_final)
            
            st.download_button(
                label=f"⬇️ Descargar PDF {empresa}",
                data=pdf_data,
                file_name=f"Reporte_{empresa.replace(' ', '_')}.pdf",
                mime="application/pdf",
                key=f"btn_{empresa.replace(' ', '_')}_{fecha_sel}"
            )
        except Exception as e:
            st.error(f"Error preparando descarga para {empresa}: {e}")
