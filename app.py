import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.io as pio
from PIL import Image
from fpdf import FPDF
import tempfile
import os
from pathlib import Path

# Configuración global
pio.templates.default = "plotly"
COLOR_PALETTE = px.colors.qualitative.Plotly

# Función de normalización robusta de nombres de empresa
def normalizar_nombre_empresa(nombre):
    nombre = str(nombre).strip().upper()
    nombre = nombre.replace('.', '').replace('&', 'AND')
    nombre = ' '.join(nombre.split())  # Normaliza espacios múltiples
    equivalencias = {
        "JORQUERA TRANSPORTE S A": "JORQUERA TRANSPORTE S. A.",
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
        "MQ SPA": "M&Q SPA",
        "M&Q SPA": "M&Q SPA",
        "MANDQ SPA": "M&Q SPA",
        "MINING AND QUARRYING SPA": "M&Q SPA",
        "MINING AND QUARRYNG SPA": "M&Q SPA",
        "AG SERVICE SPA": "AG SERVICES SPA",
        "AG SERVICES SPA": "AG SERVICES SPA",
        "COSEDUCAM S A": "COSEDUCAM S A"
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
    "AG SERVICES SPA": str(CURRENT_DIR / "ag.png")
}
BANNER_PATH = str(CURRENT_DIR / "image.png")

# Título
st.title("📊 Dashboard: Equipos por Hora, Empresa, Fecha y Destino")

# Cargar archivo Excel (.xlsx o .xlsm)
uploaded_file = st.file_uploader("📂 Carga tu archivo Excel (.xlsx o .xlsm)", type=["xlsx", "xlsm"])

# Inicializar dataframe
df = None

if uploaded_file:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name

        df = pd.read_excel(tmp_path, engine='openpyxl')
        required_columns = {
            'fecha_col': 0,     # Columna A
            'destino_col': 3,   # Columna D
            'empresa_col': 11,  # Columna L
            'hora_col': 14      # Columna O
        }

        if len(df.columns) < max(required_columns.values()) + 1:
            st.error("❌ El archivo Excel no tiene el formato esperado.")
            st.stop()

        fecha_col = df.columns[required_columns['fecha_col']]
        destino_col = df.columns[required_columns['destino_col']]
        empresa_col = df.columns[required_columns['empresa_col']]
        hora_col = df.columns[required_columns['hora_col']]

        df = df.dropna(subset=[fecha_col, destino_col, empresa_col, hora_col])

        # Procesar fechas y horas
        df[fecha_col] = pd.to_datetime(df[fecha_col], errors='coerce', dayfirst=True)
        df[hora_col] = pd.to_datetime(df[hora_col], format='%H:%M:%S', errors='coerce').dt.hour

        # Normalizar empresas
        df[empresa_col] = df[empresa_col].apply(normalizar_nombre_empresa)

    except Exception as e:
        st.error(f"🚫 Error al procesar el archivo: {e}")
else:
    st.info("📌 Puedes cargar un archivo Excel o ingresar datos manualmente.")

# Sección de ingreso manual
st.markdown("---")
st.subheader("✍️ Ingresar datos manualmente")

with st.form("manual_form"):
    fecha = st.date_input("📅 Fecha")
    abc_seguridad = st.selectbox("🔤 ABC de Seguridad", ["A", "B", "C"])
    tiempo_salar = st.text_input("⏳ Tiempo Promedio Salar (ej: 4:30)")
    tiempo_angamos = st.text_input("⏱️ Tiempo Promedio Angamos (ej: 5:15)")

    submit = st.form_submit_button("Guardar Dato")

    if submit:
        if not tiempo_salar or ":" not in tiempo_salar:
            st.error("⚠️ Formato incorrecto en Tiempo Salar. Usa HH:MM.")
        elif not tiempo_angamos or ":" not in tiempo_angamos:
            st.error("⚠️ Formato incorrecto en Tiempo Angamos. Usa HH:MM.")
        else:
            st.success("✅ Datos guardados correctamente.")
            # Aquí puedes guardar en SQLite o CSV si quieres almacenamiento persistente

# Mostrar dashboard si hay datos cargados
if df is not None and not df.empty:
    fechas_disponibles = df[fecha_col].dropna().dt.date.unique()
    if len(fechas_disponibles) == 0:
        st.warning("No se encontraron fechas válidas.")
    else:
        fecha_sel = st.date_input(
            "📅 Selecciona la fecha:",
            min_value=min(fechas_disponibles),
            max_value=max(fechas_disponibles),
            value=min(fechas_disponibles)
        )
        df_filtrado = df[df[fecha_col].dt.date == fecha_sel]
        destinos = sorted(df_filtrado[destino_col].dropna().unique())
        destinos_sel = st.multiselect("📍 Selecciona destino(s):", destinos, default=list(destinos))
        empresas = sorted(df_filtrado[empresa_col].dropna().unique())
        empresas_sel = st.multiselect("🏭 Selecciona empresa(s):", empresas, default=list(empresas))

        df_filtrado = df_filtrado[
            df_filtrado[destino_col].isin(destinos_sel) &
            df_filtrado[empresa_col].isin(empresas_sel)
        ]

        if not df_filtrado.empty:
            horas = df_filtrado[hora_col].dropna().unique()
            if len(horas) > 0:
                min_hora, max_hora = int(min(horas)), int(max(horas))
                hora_rango = st.slider("⏰ Rango de horas:", min_hora, max_hora, (min_hora, max_hora), step=1)
                df_filtrado = df_filtrado[(df_filtrado[hora_col] >= hora_rango[0]) & (df_filtrado[hora_col] <= hora_rango[1])]

            for empresa in empresas_sel:
                empresa_normalizada = normalizar_nombre_empresa(empresa)
                st.markdown(f"---\n### Empresa: {empresa}")

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

                    df_empresa = df_filtrado[df_filtrado[empresa_col] == empresa_normalizada]
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
                        index=df_empresa[hora_col],
                        columns=destino_col,
                        values=empresa_col,
                        aggfunc='count',
                        fill_value=0
                    )
                    st.dataframe(tabla.style.format(na_rep="0", precision=0))

                st.markdown("---")
                st.subheader(f"📄 Descargar PDF para {empresa}")

                if st.button(f"Generar PDF para {empresa}"):
                    try:
                        with tempfile.TemporaryDirectory() as tmpdir:
                            grafico_path = os.path.join(tmpdir, f"grafico_{empresa}.png")
                            fig.write_image(grafico_path, width=900, height=400, scale=2)

                            images_to_stack = []
                            opened_imgs = []

                            combined_img_path = os.path.join(tmpdir, f"combinado_{empresa}.png")
                            pdf_path = os.path.join(tmpdir, f"dashboard_{empresa}.pdf")

                            # Aquí generas el PDF como en tu código original...

                            with open(pdf_path, "rb") as f:
                                pdf_bytes = f.read()

                            st.download_button(
                                label=f"⬇️ Descargar PDF para {empresa}",
                                data=pdf_bytes,
                                file_name=f"dashboard_{empresa}.pdf",
                                mime="application/pdf"
                            )

                    except Exception as e:
                        st.error(f"🚫 Error al generar el PDF: {str(e)}")

        else:
            st.info("No hay datos para mostrar con los filtros actuales.")

else:
    st.info("📁 Carga un archivo Excel para comenzar.")
