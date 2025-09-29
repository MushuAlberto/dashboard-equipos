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
    nombre = str(nombre).strip().upper()
    nombre = nombre.replace('.', '').replace('&', 'AND')
    nombre = ' '.join(nombre.split())  # Normaliza espacios múltiples
    equivalencias = {
        # JORQUERA TRANSPORTE S. A.
        "JORQUERA TRANSPORTE S A": "JORQUERA TRANSPORTE S. A.",
        # M S & D SPA y variantes
        "MINING SERVICES AND DERIVATES": "M S & D SPA",
        "MINING SERVICES AND DERIVATES SPA": "M S & D SPA",
        "M S AND D": "M S & D SPA",
        "M S AND D SPA": "M S & D SPA",
        "MSANDD SPA": "M S & D SPA",
        "MSANDD": "M S & D SPA",  # <-- NUEVO
        "M S D": "M S & D SPA",
        "M S D SPA": "M S & D SPA",
        "M S & D": "M S & D SPA",
        "M S & D SPA": "M S & D SPA",
        "MS&D SPA": "M S & D SPA",
        # M&Q SPA y variantes
        "M AND Q SPA": "M&Q SPA",
        "M AND Q": "M&Q SPA",
        "M Q SPA": "M&Q SPA",
        "MQ SPA": "M&Q SPA",
        "M&Q SPA": "M&Q SPA",
        "MANDQ SPA": "M&Q SPA",
        "MANDQ": "M&Q SPA",  # <-- NUEVO
        "MINING AND QUARRYING SPA": "M&Q SPA",
        "MINING AND QUARRYNG SPA": "M&Q SPA",
        # AG SERVICES SPA
        "AG SERVICE SPA": "AG SERVICES SPA",
        "AG SERVICES SPA": "AG SERVICES SPA",
        # COSEDUCAM S A
        "COSEDUCAM S A": "COSEDUCAM S A",
        "COSEDUCAM": "COSEDUCAM S A"
    }
    return equivalencias.get(nombre, nombre)

# Configuración de la página
st.set_page_config(page_title="Dashboard Equipos por Hora", layout="wide")

CURRENT_DIR = Path(__file__).parent
LOGOS = {
    "COSEDUCAM S A": str(CURRENT_DIR / "coseducam.png"),
    "M&Q SPA": str(CURRENT_DIR / "mq.png"),
    "M S & D SPA": str(CURRENT_DIR / "msd.png"),
    "JORQUERA TRANSPORTE S. A.": str(CURRENT_DIR / "jorquera.png"),
    "AG SERVICES SPA": str(CURRENT_DIR / "ag.png"),
    "AG SERVICES": str(CURRENT_DIR / "ag.png")
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

        # Normalizar nombres de empresa
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

            # ✅ Botón con key único para evitar conflictos en Streamlit
            if st.button(f"Generar y descargar PDF para {empresa}", key=f"pdf_btn_{empresa}"):
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        df_empresa = df[df[empresa_col] == empresa_normalizada]
                        resumen = df_empresa.groupby([hora_col, destino_col]).size().reset_index(name='Cantidad')

                        if resumen.empty:
                            st.warning("No hay datos para generar el PDF.")
                            continue

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

                        # Banner
                        if os.path.exists(BANNER_PATH):
                            with Image.open(BANNER_PATH) as banner_img:
                                banner_rgb = banner_img.convert("RGB")
                                images_to_stack.append(banner_rgb)

                        # Gráfico
                        with Image.open(grafico_path) as grafico_img:
                            grafico_rgb = grafico_img.convert("RGB")

                            # Logo (opcional)
                            logo_path = LOGOS.get(empresa_normalizada)
                            if logo_path and os.path.exists(logo_path):
                                with Image.open(logo_path) as logo_img:
                                    logo_rgb = logo_img.convert("RGBA")
                                    logo_width = 120
                                    wpercent = (logo_width / float(logo_rgb.size[0]))
                                    hsize = int((float(logo_rgb.size[1]) * float(wpercent)))
                                    logo_resized = logo_rgb.resize((logo_width, hsize), Image.LANCZOS)

                                    # Crear fondo blanco para el logo
                                    logo_bg = Image.new('RGB', (grafico_rgb.width, logo_resized.height), (255, 255, 255))
                                    logo_bg.paste(
                                        logo_resized,
                                        ((grafico_rgb.width - logo_width) // 2, 0),
                                        logo_resized if logo_resized.mode == 'RGBA' else None
                                    )
                                    images_to_stack.append(logo_bg)

                            images_to_stack.append(grafico_rgb)

                        # Redimensionar todas las imágenes al ancho del gráfico
                        base_width = images_to_stack[-1].width
                        resized_imgs = []
                        for img in images_to_stack:
                            if img.width != base_width:
                                wpercent = (base_width / float(img.size[0]))
                                hsize = int((float(img.size[1]) * float(wpercent)))
                                img = img.resize((base_width, hsize), Image.LANCZOS)
                            resized_imgs.append(img)

                        # Combinar verticalmente
                        total_height = sum(img.height for img in resized_imgs)
                        combined_img = Image.new('RGB', (base_width, total_height), (255, 255, 255))
                        y_offset = 0
                        for img in resized_imgs:
                            combined_img.paste(img, (0, y_offset))
                            y_offset += img.height

                        combined_path = os.path.join(tmpdir, f"combinado_{empresa}.png")
                        combined_img.save(combined_path)

                        # Generar PDF
                        pdf = FPDF(orientation='L', unit='mm', format='A4')
                        pdf.add_page()
                        pdf.set_font("Arial", "B", 16)
                        pdf.cell(0, 10, f"Empresa: {empresa}", ln=1, align="C")
                        pdf.ln(5)
                        pdf.image(combined_path, x=10, y=20, w=270)

                        # Segunda página: tabla
                        pdf.add_page(orientation='P')
                        pdf.set_font("Arial", "B", 12)
                        pdf.cell(0, 10, "Tabla de equipos por hora y destino", ln=1, align="C")
                        pdf.set_font("Arial", "", 8)

                        # Preparar tabla
                        df_empresa['Hora Intervalo'] = df_empresa[hora_col].apply(
                            lambda h: f"{str(int(h)).zfill(2)}:00 - {str(int(h)).zfill(2)}:59"
                        )
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

                        tabla_reset = tabla_final.reset_index()
                        hora_col_name = tabla_reset.columns[0]

                        col_width = max(20, int(180 / (len(tabla_final.columns) + 1)))
                        # Encabezados
                        pdf.cell(col_width, 8, "Hora", border=1, align="C")
                        for col in tabla_final.columns:
                            pdf.cell(col_width, 8, str(col)[:30], border=1, align="C")  # Truncar si es muy largo
                        pdf.ln()
                        # Filas
                        for idx, row in tabla_reset.iterrows():
                            hora_label = row[hora_col_name]
                            if pd.isnull(hora_label) or hora_label == "TOTAL":
                                pass
                            pdf.cell(col_width, 8, str(hora_label), border=1, align="C")
                            for col in tabla_final.columns:
                                valor = int(row[col]) if not pd.isna(row[col]) else 0
                                pdf.cell(col_width, 8, str(valor), border=1, align="C")
                            pdf.ln()

                        pdf_path = os.path.join(tmpdir, f"dashboard_{empresa}.pdf")
                        pdf.output(pdf_path)

                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()

                        st.download_button(
                            label=f"Descargar PDF para {empresa}",
                            data=pdf_bytes,
                            file_name=f"dashboard_{empresa}.pdf",
                            mime="application/pdf",
                            key=f"download_pdf_{empresa}"
                        )

                except Exception as e:
                    st.error(f"Error al generar el PDF: {str(e)}")

    except Exception as e:
        st.error(f"Error al procesar el archivo: {str(e)}")

else:
    st.info("Carga un archivo Excel para ver el dashboard.")

