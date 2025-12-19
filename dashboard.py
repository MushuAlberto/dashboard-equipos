import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from PIL import Image
from fpdf import FPDF
import tempfile
import os
from pathlib import Path

# Forzar tema de color en Plotly
pio.templates.default = "plotly"

# Paleta de colores para destinos
COLOR_PALETTE = px.colors.qualitative.Plotly

# Función de normalización robusta de nombres de empresa
def normalizar_nombre_empresa(nombre):
    """Normaliza nombres de empresa para estandarizar variantes."""
    nombre = str(nombre).strip().upper()
    nombre = nombre.replace('.', '').replace('&', 'AND')
    nombre = ' '.join(nombre.split())  # Normaliza espacios múltiples
    equivalencias = {
        "JORQUERA TRANSPORTE S A": "JORQUERA TRANSPORTE S. A.",
        "JORQUERA TRANSPORTE SA": "JORQUERA TRANSPORTE S. A.",
        "MINING SERVICES AND DERIVATES": "M S & D SPA",
        "MINING SERVICES AND DERIVATES SPA": "M S & D SPA",
        "M S AND D": "M S & D SPA",
        "M S AND D SPA": "M S & D SPA",
        "MSANDD SPA": "M S & D SPA",
        "M S D": "M S & D SPA",
        "M S D SPA": "M S & D SPA",
        "M S & D": "M S & D SPA",
        "M S & D SPA": "M S & D SPA",
        "MS&D SPA": "M S & D SPA",
        "M AND Q SPA": "M&Q SPA",
        "M AND Q": "M&Q SPA",
        "M Q SPA": "M&Q SPA",
        "M & Q": "M&Q SPA",
        "MQ SPA": "M&Q SPA",
        "M&Q SPA": "M&Q SPA",
        "MANDQ SPA": "M&Q SPA",
        "MINING AND QUARRYING SPA": "M&Q SPA",
        "MINING AND QUARRYNG SPA": "M&Q SPA",
        "AG SERVICE SPA": "AG SERVICES SPA",
        "AG SERVICES SPA": "AG SERVICES SPA",
        "AG SERVICES": "AG SERVICES SPA",
        "COSEDUCAM": "COSEDUCAM S A"
    }
    return equivalencias.get(nombre, nombre)

# Configuración de la página
st.set_page_config(page_title="Dashboard Equipos por Hora", layout="wide")

# Definir rutas de archivos relativas al script
CURRENT_DIR = Path(__file__).parent
LOGOS = {
    "COSEDUCAM S A": str(CURRENT_DIR / "coseducam.png"),
    "M&Q SPA": str(CURRENT_DIR / "mq.png"),
    "M S & D SPA": str(CURRENT_DIR / "msd.png"),
    "AGRETOC": str(CURRENT_DIR / "agretoc.png"),
    "JORQUERA TRANSPORTE S. A.": str(CURRENT_DIR / "jorquera.png"),
    "AG SERVICES SPA": str(CURRENT_DIR / "ag.png")
}
BANNER_PATH = str(CURRENT_DIR / "image.png")

st.title("Dashboard: Equipos por Hora, Empresa, Fecha y Destino")

uploaded_file = st.file_uploader("Carga tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        required_columns_idx = {'fecha_col': 0, 'destino_col': 3, 'empresa_col': 11, 'hora_col': 14}

        max_idx = max(required_columns_idx.values())
        if len(df.columns) < max_idx + 1:
            st.error(f"El archivo Excel debe tener al menos {max_idx + 1} columnas.")
            st.stop()

        fecha_col_name = df.columns[required_columns_idx['fecha_col']]
        destino_col_name = df.columns[required_columns_idx['destino_col']]
        empresa_col_name = df.columns[required_columns_idx['empresa_col']]
        hora_col_name = df.columns[required_columns_idx['hora_col']]

        df = df.dropna(subset=[fecha_col_name, destino_col_name, empresa_col_name, hora_col_name])

        try:
            df[fecha_col_name] = pd.to_datetime(df[fecha_col_name], errors='coerce', dayfirst=True)
            df[hora_col_name] = pd.to_datetime(df[hora_col_name], format='%H:%M:%S', errors='coerce').dt.hour
        except Exception as e:
            st.error(f"Error al procesar fechas u horas: {str(e)}")
            st.stop()

        df = df.dropna(subset=[fecha_col_name, hora_col_name])
        df[empresa_col_name] = df[empresa_col_name].apply(normalizar_nombre_empresa)

        fechas_disponibles = df[fecha_col_name].dt.date.unique()
        fecha_sel = st.date_input("Selecciona la fecha:", min_value=min(fechas_disponibles), max_value=max(fechas_disponibles), value=min(fechas_disponibles))

        df_filtrado = df[df[fecha_col_name].dt.date == fecha_sel].copy()
        destinos_disponibles = sorted(df_filtrado[destino_col_name].dropna().unique())
        empresas_disponibles = sorted(df_filtrado[empresa_col_name].dropna().unique())

        destinos_sel = st.multiselect("Selecciona destino(s):", destinos_disponibles, default=list(destinos_disponibles))
        empresas_sel = st.multiselect("Selecciona empresa(s):", empresas_disponibles, default=list(empresas_disponibles))

        df_filtrado = df_filtrado[df_filtrado[destino_col_name].isin(destinos_sel) & df_filtrado[empresa_col_name].isin(empresas_sel)]

        horas_disponibles_filtradas = df_filtrado[hora_col_name].dropna().unique()
        if len(horas_disponibles_filtradas) > 0:
            hora_rango = st.slider("Selecciona el rango de horas:", 0, 23, (int(min(horas_disponibles_filtradas)), int(max(horas_disponibles_filtradas))), format="%d:00")
            df_filtrado = df_filtrado[(df_filtrado[hora_col_name] >= hora_rango[0]) & (df_filtrado[hora_col_name] <= hora_rango[1])]

        df_filtrado['Hora Intervalo'] = df_filtrado[hora_col_name].apply(lambda h: f"{str(int(h)).zfill(2)}:00 - {str(int(h)).zfill(2)}:59")

        for empresa in empresas_sel:
            empresa_normalizada = normalizar_nombre_empresa(empresa)
            df_empresa = df_filtrado[df_filtrado[empresa_col_name] == empresa_normalizada].copy()

            st.markdown(f"---\n## Empresa: {empresa}")
            col1, col2 = st.columns([2, 2])

            with col1:
                if os.path.exists(BANNER_PATH): st.image(BANNER_PATH, use_container_width=True)
                logo_path = LOGOS.get(empresa_normalizada)
                if logo_path and os.path.exists(logo_path): st.image(logo_path, width=100)

                resumen_grafico = df_empresa.groupby([hora_col_name, destino_col_name]).size().reset_index(name='Cantidad')
                if not resumen_grafico.empty:
                    fig = px.line(resumen_grafico, x=hora_col_name, y="Cantidad", color=destino_col_name, markers=True, title=f"Equipos por hora - {empresa}")
                    st.plotly_chart(fig, use_container_width=True)

            with col2:
                if not df_empresa.empty:
                    tabla = pd.pivot_table(df_empresa, index='Hora Intervalo', columns=destino_col_name, values=empresa_col_name, aggfunc='count', fill_value=0)
                    horas_labels = [f"{str(h).zfill(2)}:00 - {str(h).zfill(2)}:59" for h in range(24)]
                    tabla = tabla.reindex(horas_labels, fill_value=0)
                    sumatoria = pd.DataFrame(tabla.sum(axis=0)).T
                    sumatoria.index = ['TOTAL']
                    tabla_final = pd.concat([tabla, sumatoria])
                    st.dataframe(tabla_final.style.format(precision=0))

            if st.button(f"Generar PDF para {empresa}", key=f"btn_{empresa_normalizada}"):
                with st.spinner("Generando..."):
                    try:
                        with tempfile.TemporaryDirectory() as tmpdir:
                            # 1. Generar Imagen del Gráfico
                            grafico_path = os.path.join(tmpdir, "graf.png")
                            fig.update_layout(width=900, height=400)
                            fig.write_image(grafico_path, scale=2)

                            # 2. Procesar Imágenes (AJUSTE DE LOGO)
                            images_to_stack = []
                            # Base width definido por el gráfico
                            base_width = 1800 

                            # Banner
                            if os.path.exists(BANNER_PATH):
                                b_img = Image.open(BANNER_PATH).convert('RGB')
                                w_perc = base_width / float(b_img.size[0])
                                b_img = b_img.resize((base_width, int(b_img.size[1] * w_perc)), Image.Resampling.LANCZOS)
                                images_to_stack.append(b_img)

                            # Logo Pequeño (Corrección Solicitada)
                            if logo_path and os.path.exists(logo_path):
                                l_img = Image.open(logo_path).convert('RGBA')
                                # Definir tamaño pequeño para el logo
                                l_small_w = 200 
                                l_perc = l_small_w / float(l_img.size[0])
                                l_img = l_img.resize((l_small_w, int(l_img.size[1] * l_perc)), Image.Resampling.LANCZOS)
                                
                                # Crear un lienzo del ancho base para que el logo NO se estire
                                l_canvas = Image.new('RGB', (base_width, l_img.height + 40), (255, 255, 255))
                                # Pegar logo centrado (o a la izquierda si prefieres)
                                l_canvas.paste(l_img, (40, 20), l_img) 
                                images_to_stack.append(l_canvas)

                            # Gráfico
                            g_img = Image.open(grafico_path).convert('RGB')
                            g_perc = base_width / float(g_img.size[0])
                            g_img = g_img.resize((base_width, int(g_img.size[1] * g_perc)), Image.Resampling.LANCZOS)
                            images_to_stack.append(g_img)

                            # Combinar
                            total_h = sum(i.height for i in images_to_stack)
                            combined = Image.new('RGB', (base_width, total_h), (255, 255, 255))
                            y_off = 0
                            for i in images_to_stack:
                                combined.paste(i, (0, y_off))
                                y_off += i.height
                            
                            combined_path = os.path.join(tmpdir, "comb.png")
                            combined.save(combined_path)

                            # 3. Crear PDF
                            pdf = FPDF(orientation='L', unit='mm', format='A4')
                            pdf.add_page()
                            pdf.set_font("Arial", "B", 16)
                            pdf.cell(0, 10, f"Reporte - {empresa}", ln=1, align="C")
                            pdf.image(combined_path, x=10, y=25, w=277)
                            
                            # Página 2: Tabla
                            pdf.add_page(orientation='P')
                            pdf.set_font("Arial", "B", 12)
                            pdf.cell(0, 10, "Tabla de equipos", ln=1, align="C")
                            pdf.set_font("Arial", "", 7)
                            
                            # Dibujar tabla simple
                            col_w = 190 / (len(tabla_final.columns) + 1)
                            pdf.cell(col_w, 8, "Hora", 1)
                            for c in tabla_final.columns: pdf.cell(col_w, 8, str(c), 1)
                            pdf.ln()
                            for idx, row in tabla_final.iterrows():
                                pdf.cell(col_w, 8, str(idx), 1)
                                for val in row: pdf.cell(col_w, 8, str(int(val)), 1)
                                pdf.ln()

                            pdf_path = os.path.join(tmpdir, "report.pdf")
                            pdf.output(pdf_path)
                            with open(pdf_path, "rb") as f:
                                st.download_button("Descargar PDF", f, file_name=f"{empresa}.pdf")

                    except Exception as e:
                        st.error(f"Error: {e}")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")
else:
    st.info("Carga un archivo Excel.")
