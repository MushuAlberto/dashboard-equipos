import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import tempfile
import os

# --- FUNCIONES DE LIMPIEZA ---
def normalizar_empresa(nombre):
    return str(nombre).strip().upper().replace('.', '').replace('&', 'AND')

def safe_text(text):
    # FPDF2 maneja mejor los caracteres, pero esto asegura compatibilidad
    return str(text).encode('latin-1', 'replace').decode('latin-1')

# --- APP PRINCIPAL ---
st.set_page_config(page_title="Reportes M&Q", layout="wide")
st.title("Generador de Reportes de Equipos")

uploaded_file = st.file_uploader("Cargar Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    # Columnas: A=0 (Fecha), D=3 (Destino), L=11 (Empresa), O=14 (Hora)
    f_col, d_col, e_col, h_col = df.columns[0], df.columns[3], df.columns[11], df.columns[14]
    
    df = df.dropna(subset=[f_col, d_col, e_col, h_col])
    df[f_col] = pd.to_datetime(df[f_col], errors='coerce', dayfirst=True)
    df[h_col] = pd.to_datetime(df[h_col], format='%H:%M:%S', errors='coerce').dt.hour
    df[e_col] = df[e_col].apply(normalizar_empresa)

    # Filtros
    fecha_sel = st.date_input("Selecciona Fecha", value=df[f_col].min())
    df_filtered = df[df[f_col].dt.date == fecha_sel]
    empresas_encontradas = sorted(df_filtered[e_col].unique())

    for empresa in empresas_encontradas:
        st.markdown(f"---")
        st.subheader(f"Empresa: {empresa}")
        
        df_emp = df_filtered[df_filtered[e_col] == empresa]
        
        # Gráfico y Tabla
        resumen = df_emp.groupby([h_col, d_col]).size().reset_index(name='Cant')
        fig = px.line(resumen, x=h_col, y="Cant", color=d_col, markers=True)
        
        c1, c2 = st.columns([2, 1])
        c1.plotly_chart(fig, use_container_width=True)
        
        # Preparar tabla para PDF
        labels = [f"{str(h).zfill(2)}:00" for h in range(24)]
        df_emp['HR'] = df_emp[h_col].apply(lambda x: f"{str(int(x)).zfill(2)}:00")
        tabla = pd.pivot_table(df_emp, index='HR', columns=d_col, values=e_col, aggfunc='count', fill_value=0)
        tabla = tabla.reindex(labels, fill_value=0)
        c2.dataframe(tabla)

        # --- GENERACIÓN DE PDF ---
        try:
            pdf = FPDF(orientation='L', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_font("helvetica", "B", 16)
            pdf.cell(0, 10, safe_text(f"Reporte de Equipos: {empresa}"), ln=1, align="C")
            
            # Gráfico a Imagen
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img:
                fig.write_image(tmp_img.name, format="png", scale=2)
                pdf.image(tmp_img.name, x=10, y=30, w=275)
                img_name = tmp_img.name

            # Generar los bytes del PDF de forma segura
            # .output() en fpdf2 devuelve bytes por defecto
            pdf_output = pdf.output()
            
            # TRUCO DE COMPATIBILIDAD: Forzar conversión a bytes puros
            if isinstance(pdf_output, (bytearray, str)):
                final_pdf_bytes = bytes(pdf_output) if isinstance(pdf_output, bytearray) else pdf_output.encode('latin-1')
            else:
                final_pdf_bytes = pdf_output

            st.download_button(
                label=f"📥 Descargar PDF {empresa}",
                data=final_pdf_bytes,
                file_name=f"Reporte_{empresa.replace(' ', '_')}.pdf",
                mime="application/pdf",
                key=f"dl_btn_{empresa}_{fecha_sel}"
            )
            
            # Limpiar temporal
            if os.path.exists(img_name):
                os.remove(img_name)

        except Exception as e:
            st.error(f"Error procesando {empresa}: {e}")
