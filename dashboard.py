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

# Paleta de colores para destinos (puedes personalizarla)
COLOR_PALETTE = px.colors.qualitative.Plotly

# Configuración de la página
st.set_page_config(page_title="Dashboard Equipos por Hora", layout="wide")

def normalizar_nombre_empresa(nombre):
    equivalencias = {
        "AG SERVICE SPA": "AG SERVICES SPA",
        "AG SERVICES SPA": "AG SERVICES SPA",
    }
    return equivalencias.get(nombre, nombre)

CURRENT_DIR = Path(__file__).parent
LOGOS = {
    "COSEDUCAM S A": str(CURRENT_DIR / "coseducam.png"),
    "M&Q SPA": str(CURRENT_DIR / "mq.png"),
    "M S & D SPA": str(CURRENT_DIR / "msd.png"),
    "JORQUERA TRANSPORTE S. A.": str(CURRENT_DIR / "jorquera.png"),
    "AG SERVICES SPA": str(CURRENT_DIR / "ag.png")
}
BANNER_PATH = str(CURRENT_DIR / "image.png")

st.title("Dashboard: Equipos por Hora, Empresa, Fecha y Destino")

uploaded_file = st.file_uploader("Carga tu archivo Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        required_columns = {
            'fecha_col': 0,     # Columna A
            'destino_col': 3,   # Columna D
            'empresa_col': 11,  # Columna L
            'hora_col': 14      # Columna O
        }

        if len(df.columns) < max(required_columns.values()) + 1:
            st.error("El archivo Excel no tiene el formato esperado. Por favor, verifica el archivo.")
            st.stop()

        fecha_col = df.columns[required_columns['fecha_col']]
        destino_col = df.columns[required_columns['destino_col']]
        empresa_col = df.columns[required_columns['empresa_col']]
        hora_col = df.columns[required_columns['hora_col']]

        df = df.dropna(subset=[fecha_col, destino_col, empresa_col, hora_col])

        try:
            df[fecha_col] = pd.to_datetime(df[fecha_col], errors='coerce', dayfirst=True)
            df[hora_col] = pd.to_datetime(df[hora_col], format='%H:%M:%S', errors='coerce').dt.hour
        except Exception as e:
            st.error(f"Error al procesar fechas u horas: {str(e)}")
            st.stop()

        df[empresa_col] = df[empresa_col].apply(normalizar_nombre_empresa)

        fechas_disponibles = df[fecha_col].dropna().dt.date.unique()
        if len(fechas_disponibles) == 0:
            st.warning("No hay fechas válidas en el archivo.")
            st.stop()

        fecha_sel = st.date_input(
            "Selecciona la fecha:",
            min_value=min(fechas_disponibles),
            max_value=max(fechas_disponibles),
            value=min(fechas_disponibles)
        )
        df = df[df[fecha_col].dt.date == fecha_sel]

        destinos = sorted(df[destino_col].dropna().unique())
        destinos_sel = st.multiselect("Selecciona destino(s):", destinos, default=list(destinos))

        empresas = sorted(df[empresa_col].dropna().unique())
        empresas_sel = st.multiselect("Selecciona empresa(s):", empresas, default=list(empresas))

        df = df[
            df[destino_col].isin(destinos_sel) &
            df[empresa_col].isin(empresas_sel)
        ]

        horas = df[hora_col].dropna().unique()
        if len(horas) > 0:
            min_hora, max_hora = int(min(horas)), int(max(horas))
            hora_rango = st.slider(
                "Selecciona el rango de horas de entrada:",
                min_value=min_hora,
                max_value=max_hora,
                value=(min_hora, max_hora),
                step=1
            )
            df = df[(df[hora_col] >= hora_rango[0]) & (df[hora_col] <= hora_rango[1])]

        horas_labels = [f"{str(h).zfill(2)}:00 - {str(h).zfill(2)}:59" for h in range(24)]
        df['Hora Intervalo'] = df[hora_col].apply(
            lambda h: f"{str(int(h)).zfill(2)}:00 - {str(int(h)).zfill(2)}:59"
        )

        for empresa in empresas_sel:
            empresa_normalizada = normalizar_nombre_empresa(empresa)
            st.markdown(f"---\n## Empresa: {empresa}")

            col1, col2 = st.columns([2, 2])

            with col1:
                try:
                    if os.path.exists(BANNER_PATH):
                        st.image(BANNER_PATH, use_container_width=True)
                    logo_path = LOGOS.get(empresa_normalizada)
                    if logo_path and os.path.exists(logo_path):
                        st.image(logo_path, width=120)
                    else:
                        st.info(f"No se encontró logo para {empresa}")
                except Exception as e:
                    st.warning(f"Error al cargar imágenes: {str(e)}")

                df_empresa = df[df[empresa_col] == empresa_normalizada]
                resumen = df_empresa.groupby([hora_col, destino_col]).size().reset_index(name='Cantidad')

                if not resumen.empty:
                    # Asignar colores a cada destino
                    destinos_unicos = resumen[destino_col].unique()
                    color_map = {dest: COLOR_PALETTE[i % len(COLOR_PALETTE)] for i, dest in enumerate(destinos_unicos)}
                    fig = px.line(
                        resumen,
                        x=hora_col,
                        y="Cantidad",
                        color=destino_col,
                        markers=True,
                        labels={
                            hora_col: "Hora de Entrada",
                            "Cantidad": "Cantidad de Equipos",
                            destino_col: "Destino"
                        },
                        color_discrete_map=color_map
                    )
                    fig.update_layout(
                        xaxis=dict(dtick=1),
                        title=f"Cantidad de equipos por hora - {empresa}"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("No hay datos para los filtros seleccionados.")

            with col2:
                tabla = pd.pivot_table(
                    df_empresa,
                    index='Hora Intervalo',
                    columns=destino_col,
                    values=empresa_col,
                    aggfunc='count',
                    fill_value=0
                )
                tabla = tabla.reindex(horas_labels, fill_value=0)
                sumatoria = pd.DataFrame(tabla.sum(axis=0)).T
                sumatoria.index = ['TOTAL']
                tabla_final = pd.concat([tabla, sumatoria])
                st.dataframe(tabla_final.style.format(na_rep="0", precision=0))

            st.markdown("---")
            st.subheader(f"Descargar PDF para {empresa}")

            if st.button(f"Generar y descargar PDF para {empresa}"):
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        # Asignar colores a cada destino para el gráfico del PDF
                        destinos_unicos = resumen[destino_col].unique()
                        color_map = {dest: COLOR_PALETTE[i % len(COLOR_PALETTE)] for i, dest in enumerate(destinos_unicos)}
                        fig2 = px.line(
                            resumen,
                            x=hora_col,
                            y="Cantidad",
                            color=destino_col,
                            markers=True,
                            labels={
                                hora_col: "Hora de Entrada",
                                "Cantidad": "Cantidad de Equipos",
                                destino_col: "Destino"
                            },
                            color_discrete_map=color_map
                        )
                        fig2.update_layout(
                            xaxis=dict(dtick=1),
                            title=f"Cantidad de equipos por hora - {empresa}"
                        )
                        grafico_path = os.path.join(tmpdir, f"grafico_{empresa}.png")
                        fig2.write_image(grafico_path, width=900, height=400, scale=2)

                        images_to_stack = []
                        opened_imgs = []

                        if os.path.exists(BANNER_PATH):
                            banner_img = Image.open(BANNER_PATH)
                            images_to_stack.append(banner_img)
                            opened_imgs.append(banner_img)

                        logo_path = LOGOS.get(empresa_normalizada)
                        if logo_path and os.path.exists(logo_path):
                            logo_img = Image.open(logo_path)
                            logo_width = 120
                            wpercent = (logo_width / float(logo_img.size[0]))
                            hsize = int((float(logo_img.size[1]) * float(wpercent)))
                            logo_img = logo_img.resize((logo_width, hsize), Image.LANCZOS)
                            grafico_img = Image.open(grafico_path)
                            logo_bg = Image.new('RGBA', (grafico_img.width, logo_img.height), (255,255,255,0))
                            logo_bg.paste(
                                logo_img,
                                ((grafico_img.width - logo_width)//2, 0),
                                logo_img if logo_img.mode=='RGBA' else None
                            )
                            images_to_stack.append(logo_bg.convert('RGB'))
                            opened_imgs.extend([logo_img, grafico_img])
                        else:
                            grafico_img = Image.open(grafico_path)
                            opened_imgs.append(grafico_img)

                        images_to_stack.append(grafico_img)

                        base_width = images_to_stack[-1].width
                        resized_imgs = []
                        for img in images_to_stack:
                            if img.width != base_width:
                                wpercent = (base_width / float(img.size[0]))
                                hsize = int((float(img.size[1]) * float(wpercent)))
                                img = img.resize((base_width, hsize), Image.LANCZOS)
                            resized_imgs.append(img)

                        total_height = sum(img.height for img in resized_imgs)
                        combined_img = Image.new('RGB', (base_width, total_height), (255, 255, 255))

                        y_offset = 0
                        for img in resized_imgs:
                            combined_img.paste(img, (0, y_offset))
                            y_offset += img.height

                        combined_path = os.path.join(tmpdir, f"combinado_{empresa}.png")
                        combined_img.save(combined_path)

                        for img in opened_imgs:
                            img.close()
                        combined_img.close()

                        pdf = FPDF(orientation='L', unit='mm', format='A4')
                        pdf.add_page()
                        pdf.set_font("Arial", "B", 16)
                        pdf.cell(0, 10, f"Empresa: {empresa}", ln=1, align="C")
                        pdf.ln(5)
                        pdf.image(combined_path, x=10, y=20, w=270)

                        pdf.add_page(orientation='P')
                        pdf.set_font("Arial", "B", 12)
                        pdf.cell(0, 10, "Tabla de equipos por hora y destino", ln=1, align="C")
                        pdf.set_font("Arial", "", 8)
                        col_width = max(20, int(180 / (len(tabla_final.columns)+1)))
                        tabla_reset = tabla_final.reset_index()
                        hora_col_name = tabla_reset.columns[0]
                        pdf.cell(col_width, 8, "Hora", border=1, align="C")
                        for col in tabla_final.columns:
                            pdf.cell(col_width, 8, str(col), border=1, align="C")
                        pdf.ln()
                        for idx, row in tabla_reset.iterrows():
                            hora_label = row[hora_col_name]
                            if pd.isnull(hora_label):
                                hora_label = "TOTAL"
                            pdf.cell(col_width, 8, str(hora_label), border=1, align="C")
                            for col in tabla_final.columns:
                                pdf.cell(col_width, 8, str(int(row[col])), border=1, align="C")
                            pdf.ln()

                        pdf_path = os.path.join(tmpdir, f"dashboard_{empresa}.pdf")
                        pdf.output(pdf_path)

                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()

                        st.download_button(
                            label=f"Descargar PDF para {empresa}",
                            data=pdf_bytes,
                            file_name=f"dashboard_{empresa}.pdf",
                            mime="application/pdf"
                        )

                except Exception as e:
                    st.error(f"Error al generar el PDF: {str(e)}")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")

else:
    st.info("Carga un archivo Excel para ver el dashboard.")