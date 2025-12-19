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
        # JORQUERA TRANSPORTE S. A.
        "JORQUERA TRANSPORTE S A": "JORQUERA TRANSPORTE S. A.",
        "JORQUERA TRANSPORTE SA": "JORQUERA TRANSPORTE S. A.",
        # M S & D SPA y variantes
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
        # M&Q SPA y variantes
        "M AND Q SPA": "M&Q SPA",
        "M AND Q": "M&Q SPA",
        "M Q SPA": "M&Q SPA",
        "M & Q": "M&Q SPA",
        "MQ SPA": "M&Q SPA",
        "M&Q SPA": "M&Q SPA",
        "MANDQ SPA": "M&Q SPA",
        "MINING AND QUARRYING SPA": "M&Q SPA",
        "MINING AND QUARRYNG SPA": "M&Q SPA",
        # AG SERVICES SPA
        "AG SERVICE SPA": "AG SERVICES SPA",
        "AG SERVICES SPA": "AG SERVICES SPA",
        "AG SERVICES": "AG SERVICES SPA",
        # AGRETOC
        "AGRETOC": "AGRETOC",
        # COSEDUCAM S A
        "COSEDUCAM S A": "COSEDUCAM S A",
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

        # Mapeo de columnas esperadas por índice
        required_columns_idx = {
            'fecha_col': 0,     # Columna A
            'destino_col': 3,   # Columna D
            'empresa_col': 11,  # Columna L
            'hora_col': 14      # Columna O
        }

        # Validación básica de columnas por índice
        max_idx = max(required_columns_idx.values())
        if len(df.columns) < max_idx + 1:
            st.error(f"El archivo Excel debe tener al menos {max_idx + 1} columnas. Por favor, verifica el archivo.")
            st.stop()

        # Obtener nombres de columna usando los índices
        fecha_col_name = df.columns[required_columns_idx['fecha_col']]
        destino_col_name = df.columns[required_columns_idx['destino_col']]
        empresa_col_name = df.columns[required_columns_idx['empresa_col']]
        hora_col_name = df.columns[required_columns_idx['hora_col']]

        # Validar si los nombres de columna obtenidos son los esperados (opcional, pero bueno)
        # Puedes agregar aquí una verificación si sabes los nombres exactos esperados

        # Eliminar filas con valores faltantes en columnas clave
        df = df.dropna(subset=[fecha_col_name, destino_col_name, empresa_col_name, hora_col_name])

        # Procesar fechas y horas
        try:
            # Intentar varios formatos de fecha, si falla uno, prueba el siguiente
            # errors='coerce' convertirá los fallos a NaT
            df[fecha_col_name] = pd.to_datetime(df[fecha_col_name], errors='coerce', dayfirst=True)

            # Extraer solo la hora en formato 24h (0-23)
            # Usamos errors='coerce' aquí también
            df[hora_col_name] = pd.to_datetime(df[hora_col_name], format='%H:%M:%S', errors='coerce').dt.hour
        except Exception as e:
            st.error(f"Error al procesar fechas u horas. Asegúrate de que las columnas {fecha_col_name} y {hora_col_name} tengan formatos válidos. Detalles: {str(e)}")
            st.stop()

        # Filtrar filas donde el procesamiento de fecha/hora resultó en NaT/NaN
        df = df.dropna(subset=[fecha_col_name, hora_col_name])


        # Normalizar nombres de empresa
        df[empresa_col_name] = df[empresa_col_name].apply(normalizar_nombre_empresa)

        # --- Filtros de Usuario ---
        fechas_disponibles = df[fecha_col_name].dt.date.unique()
        if len(fechas_disponibles) == 0:
            st.warning("No hay fechas válidas en el archivo después del procesamiento.")
            st.stop()

        fecha_sel = st.date_input(
            "Selecciona la fecha:",
            min_value=min(fechas_disponibles),
            max_value=max(fechas_disponibles),
            # Intentar seleccionar la primera fecha disponible por defecto
            value=min(fechas_disponibles) if fechas_disponibles.any() else None
        )

        if fecha_sel is None:
             st.info("Selecciona una fecha para continuar.")
             st.stop()

        # Filtrar por fecha seleccionada
        df_filtrado = df[df[fecha_col_name].dt.date == fecha_sel].copy() # Usar .copy() para evitar SettingWithCopyWarning

        if df_filtrado.empty:
            st.info("No hay datos para la fecha seleccionada.")
            st.stop()

        # Asegurarse de que las opciones de selección provengan solo de los datos filtrados por fecha
        destinos_disponibles = sorted(df_filtrado[destino_col_name].dropna().unique())
        empresas_disponibles = sorted(df_filtrado[empresa_col_name].dropna().unique())

        if not destinos_disponibles:
             st.warning("No hay destinos válidos para la fecha seleccionada.")
             st.stop()

        if not empresas_disponibles:
             st.warning("No hay empresas válidas para la fecha seleccionada.")
             st.stop()


        destinos_sel = st.multiselect("Selecciona destino(s):", destinos_disponibles, default=list(destinos_disponibles))
        empresas_sel = st.multiselect("Selecciona empresa(s):", empresas_disponibles, default=list(empresas_disponibles))

        if not destinos_sel or not empresas_sel:
             st.info("Selecciona al menos un destino y una empresa.")
             st.stop()

        # Aplicar filtros de destino y empresa
        df_filtrado = df_filtrado[
            df_filtrado[destino_col_name].isin(destinos_sel) &
            df_filtrado[empresa_col_name].isin(empresas_sel)
        ]

        if df_filtrado.empty:
            st.info("No hay datos para la combinación de filtros seleccionada (fecha, destino, empresa).")
            st.stop()

        # Filtrar por rango de horas (basado en las horas *disponibles* después de otros filtros)
        horas_disponibles_filtradas = df_filtrado[hora_col_name].dropna().unique()
        if len(horas_disponibles_filtradas) > 0:
            min_hora, max_hora = int(min(horas_disponibles_filtradas)), int(max(horas_disponibles_filtradas))
            hora_rango = st.slider(
                "Selecciona el rango de horas de entrada:",
                min_value=0, # Rango completo de 0 a 23
                max_value=23,
                value=(min_hora, max_hora), # Valor inicial basado en datos filtrados
                step=1,
                format="%d:00"
            )
            df_filtrado = df_filtrado[(df_filtrado[hora_col_name] >= hora_rango[0]) & (df_filtrado[hora_col_name] <= hora_rango[1])]

        if df_filtrado.empty:
            st.info("No hay datos para el rango de horas seleccionado.")
            st.stop()

        # Preparar datos para visualización y tabla
        horas_labels = [f"{str(h).zfill(2)}:00 - {str(h).zfill(2)}:59" for h in range(24)] # Labels para la tabla
        df_filtrado['Hora Intervalo'] = df_filtrado[hora_col_name].apply(
            lambda h: f"{str(int(h)).zfill(2)}:00 - {str(int(h)).zfill(2)}:59" if pd.notnull(h) else "Desconocido"
        )


        # --- Mostrar Dashboard por Empresa ---
        # Usar las empresas seleccionadas por el usuario, no todas las disponibles
        for empresa in empresas_sel:
            empresa_normalizada = normalizar_nombre_empresa(empresa)
            # Filtrar el DF para la empresa actual
            df_empresa = df_filtrado[df_filtrado[empresa_col_name] == empresa_normalizada].copy()

            st.markdown(f"---\n## Empresa: {empresa}")

            col1, col2 = st.columns([2, 2])

            with col1:
                # --- Cargar Imágenes ---
                try:
                    if os.path.exists(BANNER_PATH):
                        st.image(BANNER_PATH, use_container_width=True)
                    logo_path = LOGOS.get(empresa_normalizada)
                    if logo_path and os.path.exists(logo_path):
                        st.image(logo_path, width=120)
                    # else:
                        # st.info(f"No se encontró logo para {empresa}") # Opcional: mostrar si no hay logo
                except Exception as e:
                    st.warning(f"Error al cargar imágenes (banner o logo): {str(e)}")

                # --- Gráfico Plotly ---
                resumen_grafico = df_empresa.groupby([hora_col_name, destino_col_name]).size().reset_index(name='Cantidad')

                if not resumen_grafico.empty:
                    destinos_unicos_grafico = sorted(resumen_grafico[destino_col_name].unique())
                    # Asegurar que el mapa de colores se crea solo para los destinos presentes en el gráfico
                    color_map_grafico = {
                        dest: COLOR_PALETTE[i % len(COLOR_PALETTE)]
                        for i, dest in enumerate(destinos_unicos_grafico)
                    }

                    fig = px.line(
                        resumen_grafico,
                        x=hora_col_name,
                        y="Cantidad",
                        color=destino_col_name,
                        markers=True,
                        labels={
                            hora_col_name: "Hora de Entrada",
                            "Cantidad": "Cantidad de Equipos",
                            destino_col_name: "Destino"
                        },
                        color_discrete_map=color_map_grafico
                    )
                    fig.update_layout(
                        xaxis=dict(dtick=1),
                        title=f"Cantidad de equipos por hora - {empresa}"
                    )
                    st.plotly_chart(fig, use_container_width=True) # Mostrar el gráfico en Streamlit
                else:
                    st.info(f"No hay datos para generar el gráfico de {empresa} con los filtros seleccionados.")
                    fig = None # Asegurarse de que fig sea None si no hay datos para el gráfico

            with col2:
                # --- Tabla Resumen ---
                if not df_empresa.empty:
                    tabla = pd.pivot_table(
                        df_empresa,
                        index='Hora Intervalo',
                        columns=destino_col_name, # Usar el nombre real de la columna
                        values=empresa_col_name, # Usar el nombre real de la columna
                        aggfunc='count',
                        fill_value=0
                    )
                    # Asegurar que todas las horas de 0 a 23 están en el índice, llenando con 0 si faltan
                    tabla = tabla.reindex(horas_labels, fill_value=0)

                    # Calcular la sumatoria total por columna (destino)
                    sumatoria = pd.DataFrame(tabla.sum(axis=0)).T
                    sumatoria.index = ['TOTAL']

                    # Concatenar tabla y sumatoria
                    tabla_final = pd.concat([tabla, sumatoria])

                    st.dataframe(tabla_final.style.format(na_rep="0", precision=0))
                else:
                    st.info(f"No hay datos para generar la tabla de {empresa} con los filtros seleccionados.")
                    tabla_final = pd.DataFrame() # Asegurarse de que tabla_final sea DataFrame vacío si no hay datos


            st.markdown("---")
            st.subheader(f"Descargar PDF para {empresa}")

            # Botón de descarga de PDF dentro del bucle de empresas
            # Se genera un botón por cada empresa seleccionada
            if st.button(f"Generar y descargar PDF para {empresa}", key=f"pdf_button_{empresa_normalizada}"):
                 # Mostrar un mensaje de estado mientras se genera el PDF
                 with st.spinner(f"Generando PDF para {empresa}..."):
                    try:
                        # Usar un directorio temporal para archivos intermedios
                        with tempfile.TemporaryDirectory() as tmpdir:

                            # --- Generar Imagen del Gráfico (solo si hay datos y gráfico generado) ---
                            # Re-generar el gráfico para guardarlo como imagen para el PDF
                            # Esto es necesario si fig original fue modificado o si no se generó
                            resumen_pdf_grafico = df_empresa.groupby([hora_col_name, destino_col_name]).size().reset_index(name='Cantidad')

                            if not resumen_pdf_grafico.empty:
                                destinos_unicos_pdf_grafico = sorted(resumen_pdf_grafico[destino_col_name].unique())
                                color_map_pdf_grafico = {
                                    dest: COLOR_PALETTE[i % len(COLOR_PALETTE)]
                                    for i, dest in enumerate(destinos_unicos_pdf_grafico)
                                }
                                fig_pdf = px.line(
                                    resumen_pdf_grafico,
                                    x=hora_col_name,
                                    y="Cantidad",
                                    color=destino_col_name,
                                    markers=True,
                                    labels={
                                        hora_col_name: "Hora de Entrada",
                                        "Cantidad": "Cantidad de Equipos",
                                        destino_col_name: "Destino"
                                    },
                                    color_discrete_map=color_map_pdf_grafico
                                )
                                fig_pdf.update_layout(
                                    xaxis=dict(dtick=1),
                                    title=f"Cantidad de equipos por hora - {empresa}",
                                    width=900, # Especificar ancho y alto para la exportación
                                    height=400
                                )

                                grafico_path = os.path.join(tmpdir, f"grafico_{empresa_normalizada}.png")
                                # AQUÍ ES DONDE OCURRE EL ERROR SIN CHROME/KALEIDO
                                fig_pdf.write_image(grafico_path, scale=2) # Usar scale para mayor resolución
                                has_grafico_img = True
                            else:
                                has_grafico_img = False
                                grafico_path = None
                                st.warning(f"No se pudo generar el gráfico para el PDF de {empresa} (sin datos).")


                            # --- Apilar Imágenes (Banner, Logo, Gráfico) ---
                            images_to_stack = []
                            # Lista para mantener las imágenes abiertas y cerrarlas después
                            opened_imgs = []

                            try:
                                if os.path.exists(BANNER_PATH):
                                    banner_img = Image.open(BANNER_PATH).convert('RGB') # Convertir a RGB para compatibilidad
                                    images_to_stack.append(banner_img)
                                    opened_imgs.append(banner_img)

                                logo_path = LOGOS.get(empresa_normalizada)
                                if logo_path and os.path.exists(logo_path):
                                    logo_img = Image.open(logo_path)
                                    # Redimensionar el logo
                                    logo_width = 150 # Un poco más grande quizás? Ajusta si es necesario
                                    wpercent = (logo_width / float(logo_img.size[0]))
                                    hsize = int((float(logo_img.size[1]) * float(wpercent)))
                                    logo_img = logo_img.resize((logo_width, hsize), Image.Resampling.LANCZOS) # Usar Resampling.LANCZOS
                                    images_to_stack.append(logo_img.convert('RGB')) # Convertir a RGB
                                    opened_imgs.append(logo_img)

                                if has_grafico_img and os.path.exists(grafico_path):
                                     grafico_img = Image.open(grafico_path).convert('RGB') # Convertir a RGB
                                     images_to_stack.append(grafico_img)
                                     opened_imgs.append(grafico_img)

                                if not images_to_stack:
                                    st.warning("No hay imágenes (banner, logo, gráfico) para combinar en el PDF.")
                                    continue # Saltar a la siguiente empresa o salir del bucle

                                # Determinar el ancho base (usamos el más ancho o un valor fijo)
                                # Podríamos usar el ancho del gráfico si existe, o un valor por defecto
                                base_width = grafico_img.width if has_grafico_img else 1000 # Ejemplo: 1000px si no hay gráfico

                                # Redimensionar todas las imágenes al ancho base manteniendo la proporción
                                resized_imgs = []
                                for img in images_to_stack:
                                    if img.width != base_width:
                                        wpercent = (base_width / float(img.size[0]))
                                        hsize = int((float(img.size[1]) * float(wpercent)))
                                        # Asegurarse de que las dimensiones sean válidas (positivas)
                                        if hsize > 0:
                                            img = img.resize((base_width, hsize), Image.Resampling.LANCZOS)
                                        resized_imgs.append(img)
                                    else:
                                         resized_imgs.append(img)

                                # Combinar las imágenes verticalmente
                                total_height = sum(img.height for img in resized_imgs)
                                # Crear una nueva imagen blanca con el ancho base y la altura total
                                combined_img = Image.new('RGB', (base_width, total_height), (255, 255, 255))

                                y_offset = 0
                                for img in resized_imgs:
                                    combined_img.paste(img, (0, y_offset))
                                    y_offset += img.height

                                combined_path = os.path.join(tmpdir, f"combinado_{empresa_normalizada}.png")
                                combined_img.save(combined_path)

                            except Exception as e:
                                st.error(f"Error al combinar imágenes para el PDF de {empresa}: {str(e)}")
                                combined_path = None # Asegurarse de que no se intente añadir una imagen combinada si falló
                                # Asegurarse de cerrar las imágenes aunque falle la combinación
                                for img in opened_imgs:
                                    try:
                                        img.close()
                                    except:
                                        pass # Ignorar errores al cerrar
                                # Limpiar lista ya que intentamos cerrarlas
                                opened_imgs = []

                            finally:
                                # Cerrar todas las imágenes abiertas *después* de usarlas
                                for img in opened_imgs:
                                    try:
                                        img.close()
                                    except:
                                        pass # Ignorar errores al cerrar


                            # --- Generar PDF con FPDF ---
                            pdf = FPDF(orientation='L', unit='mm', format='A4') # Primera página horizontal
                            pdf.add_page()
                            pdf.set_font("Arial", "B", 16)
                            pdf.cell(0, 10, f"Reporte Diario - {empresa}", ln=1, align="C")
                            pdf.set_font("Arial", "", 12)
                            pdf.cell(0, 10, f"Fecha: {fecha_sel.strftime('%d/%m/%Y')}", ln=1, align="C")
                            pdf.ln(5)

                            # Añadir la imagen combinada si se generó
                            if combined_path and os.path.exists(combined_path):
                                # Calcular las dimensiones para que la imagen quepa en la página horizontal (297x210 mm)
                                # Asumimos un margen de 10mm en cada lado
                                page_width_mm = 297 - 20
                                img_width_px, img_height_px = Image.open(combined_path).size
                                img_aspect_ratio = img_height_px / img_width_px
                                img_width_mm = page_width_mm # Intentar usar todo el ancho disponible
                                img_height_mm = img_width_mm * img_aspect_ratio

                                # Ajustar si la altura calculada excede la altura de la página (210 - 20 mm)
                                page_height_mm = 210 - 30 # Dejar espacio para título y quizás pie de página
                                if img_height_mm > page_height_mm:
                                     img_height_mm = page_height_mm
                                     img_width_mm = img_height_mm / img_aspect_ratio # Re-calcular ancho basado en nueva altura

                                pdf.image(combined_path, x=(297-img_width_mm)/2, y=25, w=img_width_mm, h=img_height_mm) # Centrar imagen
                                pdf.ln(img_height_mm + 10) # Dejar espacio después de la imagen

                            # Añadir la tabla si se generó
                            if not tabla_final.empty:
                                pdf.add_page(orientation='P') # Segunda página vertical para la tabla
                                pdf.set_font("Arial", "B", 12)
                                pdf.cell(0, 10, "Tabla de equipos por hora y destino", ln=1, align="C")
                                pdf.ln(5)

                                pdf.set_font("Arial", "", 7) # Fuente más pequeña para la tabla
                                # Calcular ancho de columna dinámicamente
                                # Ancho total disponible en página vertical (210 - 20 mm)
                                page_width_table_mm = 210 - 20
                                num_cols_table = len(tabla_final.columns) + 1 # Horas + Destinos
                                col_width = page_width_table_mm / num_cols_table
                                col_width = max(15, col_width) # Asegurar un ancho mínimo para evitar columnas muy estrechas

                                # Cabecera de la tabla
                                tabla_reset = tabla_final.reset_index()
                                hora_col_name_tabla = tabla_reset.columns[0] # Normalmente 'index' o 'Hora Intervalo'
                                pdf.cell(col_width, 8, "Hora", border=1, align="C")
                                # Ordenar las columnas de destino alfabéticamente para consistencia, excepto 'TOTAL' si existe
                                table_cols_sorted = sorted([col for col in tabla_final.columns if col != 'TOTAL']) + ([ 'TOTAL'] if 'TOTAL' in tabla_final.columns else [])

                                for col in table_cols_sorted:
                                    pdf.cell(col_width, 8, str(col), border=1, align="C")
                                pdf.ln()

                                # Filas de datos
                                for idx, row in tabla_reset.iterrows():
                                    hora_label = row[hora_col_name_tabla]
                                    # Manejar el caso de la fila 'TOTAL'
                                    if pd.isnull(hora_label) or hora_label == 'TOTAL':
                                        hora_label_str = "TOTAL"
                                        pdf.set_font("Arial", "B", 7) # Fuente negrita para TOTAL
                                    else:
                                         hora_label_str = str(hora_label)
                                         pdf.set_font("Arial", "", 7) # Fuente normal

                                    pdf.cell(col_width, 8, hora_label_str, border=1, align="C")

                                    for col in table_cols_sorted:
                                        cell_value = row.get(col, 0) # Usar .get para evitar KeyError si una columna falta inesperadamente
                                        # Asegurarse de que el valor es numérico para formatear
                                        if pd.notnull(cell_value):
                                             pdf.cell(col_width, 8, str(int(cell_value)), border=1, align="C")
                                        else:
                                             pdf.cell(col_width, 8, "0", border=1, align="C") # Mostrar 0 si es NaN
                                    pdf.ln()
                                pdf.set_font("Arial", "", 8) # Volver a la fuente normal si es necesario

                            else:
                                 st.warning(f"No se pudo generar la tabla para el PDF de {empresa} (sin datos).")
                                 # Si no hay tabla, quizás añadir un mensaje en el PDF
                                 if combined_path: # Si al menos hay imagen, añadir un mensaje en la segunda página
                                     pdf.add_page(orientation='P')
                                     pdf.set_font("Arial", "", 12)
                                     pdf.cell(0, 10, "No hay datos de tabla disponibles para estos filtros.", ln=1, align="C")


                            # --- Guardar y Descargar PDF ---
                            pdf_path = os.path.join(tmpdir, f"dashboard_{empresa_normalizada}.pdf")
                            pdf.output(pdf_path)

                            with open(pdf_path, "rb") as f:
                                pdf_bytes = f.read()

                            # Ofrecer el botón de descarga después de generar el archivo
                            st.download_button(
                                label=f"Descargar PDF para {empresa}",
                                data=pdf_bytes,
                                file_name=f"dashboard_{empresa_normalizada}_{fecha_sel.strftime('%Y%m%d')}.pdf",
                                mime="application/pdf",
                                key=f"download_button_{empresa_normalizada}" # Clave única para el download button
                            )
                            st.success(f"PDF para {empresa} generado con éxito.")

                    except Exception as e:
                        # Capturar cualquier error durante el proceso de generación de PDF
                        st.error(f"Error al generar el PDF para {empresa}: {str(e)}")


    except Exception as e:
        # Capturar errores generales al cargar o procesar el archivo Excel
        st.error(f"Error general al procesar el archivo: {str(e)}")
        st.info("Asegúrate de que el archivo sea un .xlsx válido y tenga las columnas esperadas en el orden correcto.")

else:
    st.info("Carga un archivo Excel para ver el dashboard.")



