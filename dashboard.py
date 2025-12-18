import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import io
import tempfile
import os

# --- FUNCIONES DE APOYO ---
def normalizar_nombre(nombre):
    nombre = str(nombre).strip().upper().replace('.', '').replace('&', 'AND')
    return ' '.join(nombre.split())

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Dashboard M&Q", layout="wide")
st.title("Reportes Logística - M&Q SPA y Empresas")

uploaded_file = st.file_uploader("Cargar Archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    
    # Asignación de columnas por posición
    f_col, d_col, e_col, h_col = df.columns[0], df.columns[3], df.columns[11], df.columns[14]
    
    # Limpieza básica
    df = df.dropna(subset=[f_col, d_col, e_col, h_col])
    df[f_col] = pd.to_datetime(df[f_col], errors='coerce', dayfirst=True)
    df[h_col] = pd.to_datetime(df[h_col], format='%H:%M:%S', errors='coerce').dt.hour
    df[e_col] = df[e_col].apply(normalizar_nombre)

    fechas = sorted(df[f_col].dropna().dt.date.unique())
    fecha_sel = st.date_input("Seleccionar Fecha", value=fechas[0] if fechas else None)
    
    df_filtered = df[df[f_col].dt.date == fecha_sel]
    empresas_disponibles = sorted(df_filtered[e_col].unique())

    for empresa in empresas_disponibles:
        st.markdown(f"---")
        st.header(f"Empresa: {empresa}")
        
        df_emp = df_filtered[df_filtered[e_col] == empresa]
        
        # 1. Gráfico
        resumen = df_emp.groupby([h_col, d_col]).size().reset_index(name='Cant')
        fig = px.line(resumen, x=h_col, y="Cant", color=d_col, markers=True, title=f"Flujo {empresa}")
        
        # 2. Tabla
        labels = [f"{str(h).zfill(2)}:00" for h in range(24)]
        df_emp['HR'] = df_emp[h_col].apply(lambda x: f"{str(int(x)).zfill(2)}:00")
        tabla = pd.pivot_table(df_emp, index='HR', columns=d_col, values=e_col, aggfunc='count', fill_value=0)
        tabla = tabla.reindex(labels, fill_value=0)
        
        c1, c2 = st.columns([2, 1])
        c1.plotly_chart(fig, use_container_width=True)
        c2.dataframe(tabla)

        # 3. GENERACIÓN DE PDF PROFESIONAL
        try:
            # Crear el PDF
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, f"REPORTE: {empresa}", ln=1, align="C")
            pdf.cell(0, 10, f"FECHA: {fecha_sel}", ln=1, align="C")

            # Guardar gráfico como imagen temporal para el PDF
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
                fig.write_image(tmp.name, format="png")
                pdf.image(tmp.name, x=10, y=40, w=275)
                img_path = tmp.name

            # Generar los bytes del PDF de forma segura
            pdf_bytes = pdf.output() 
            
            # Asegurar que siempre entreguemos bytes a Streamlit
            if not isinstance(pdf_bytes, (bytes, bytearray)):
                pdf_bytes = str(pdf_bytes).encode('latin-1')

            st.download_button(
                label=f"📥 Descargar PDF - {empresa}",
                data=pdf_bytes,
                file_name=f"Reporte_{empresa.replace(' ', '_')}.pdf",
                mime="application/pdf",
                key=f"btn_{empresa}_{fecha_sel}"
            )
            
            # Limpiar imagen temporal
            if os.path.exists(img_path):
                os.remove(img_path)

        except Exception as e:
            st.error(f"Error preparando el PDF de {empresa}: {e}")

else:
    st.info("👋 Por favor, carga el archivo Excel para procesar los datos.")
