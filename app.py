import streamlit as st
import pandas as pd
import plotly.express as px

# ---------------------------
# CONFIGURACIÓN INICIAL
# ---------------------------
st.set_page_config(page_title="Dashboard IPN", layout="wide")


@st.cache_data
def load_data(file_path):
    """Carga todas las hojas de un archivo Excel en un diccionario de DataFrames."""
    try:
        return pd.read_excel(file_path, sheet_name=None)
    except FileNotFoundError:
        st.error(
            f"Error: No se encontró el archivo '{file_path}'. Asegúrate de que esté en la misma carpeta.")
        return None


FILE_PATH = "EquipoComputo.xlsx"
sheets = load_data(FILE_PATH)

if sheets is None:
    st.stop()

# ---------------------------
# Nombres de tabs y creación de tabs
# ---------------------------
st.title("📊 Dashboard de Equipo de Cómputo")
tab_names = list(sheets.keys())
tabs = st.tabs(tab_names)

# ---------------------------
# PESTAÑA 1: Resumen General (con 2 tablas y 2 gráficas)
# ---------------------------
with tabs[0]:
    st.subheader(f"Resumen General - {tab_names[0]}")
    # Usar .copy() para evitar modificar el original
    df_ipn = sheets[tab_names[0]].copy()

    # --- Definir las columnas para cada tabla ---
    col_categoria = "IPN"

    columnas_tabla_equipo = [
        col_categoria,
        "Computadoras y Laptops",
        "Servidores",
        "Equipo con Conexión a Internet",
        "Equipo sin Conexión a Internet",
    ]

    columnas_tabla_usuarios = [
        col_categoria,
        "Alumnos",
        "Docentes",
        "Docentes en Labores de Investigación",
        "Personal Directivo y de Mando",
        "PAAE"
    ]

    # --- Crear layout con 2 columnas para las tablas ---
    col_tabla1, col_tabla2 = st.columns(2)

    with col_tabla1:
        st.write("#### Resumen de Equipo")
        cols_mostrar_equipo = [
            col for col in columnas_tabla_equipo if col in df_ipn.columns]
        df_equipo = df_ipn[cols_mostrar_equipo]

        # --- INICIO DEL CAMBIO: Calcular y añadir fila de total ---
        if not df_equipo.empty:
            # Seleccionar solo columnas numéricas para sumar
            numeric_cols_equipo = df_equipo.select_dtypes(
                include='number').columns
            total_equipo = df_equipo[numeric_cols_equipo].sum().to_frame().T
            total_equipo[col_categoria] = "TOTAL"  # Añadir etiqueta
            # Combinar la tabla original con la fila de total
            df_equipo_con_total = pd.concat(
                [df_equipo, total_equipo], ignore_index=True)
            st.dataframe(df_equipo_con_total, width='stretch', hide_index=True)
        else:
            st.dataframe(df_equipo, width='stretch', hide_index=True)
        # --- FIN DEL CAMBIO ---

    with col_tabla2:
        st.write("#### Resumen de Usuarios")
        cols_mostrar_usuarios = [
            col for col in columnas_tabla_usuarios if col in df_ipn.columns]
        df_usuarios = df_ipn[cols_mostrar_usuarios]

        # --- INICIO DEL CAMBIO: Calcular y añadir fila de total ---
        if not df_usuarios.empty:
            numeric_cols_usuarios = df_usuarios.select_dtypes(
                include='number').columns
            total_usuarios = df_usuarios[numeric_cols_usuarios].sum(
            ).to_frame().T
            total_usuarios[col_categoria] = "TOTAL"
            df_usuarios_con_total = pd.concat(
                [df_usuarios, total_usuarios], ignore_index=True)
            st.dataframe(df_usuarios_con_total,
                         width='stretch', hide_index=True)
        else:
            st.dataframe(df_usuarios, width='stretch', hide_index=True)
        # --- FIN DEL CAMBIO ---

    st.markdown("---")  # Separador visual

    # --- SECCIÓN DE GRÁFICAS (SIN CAMBIOS) ---
    columnas_grafica_equipo = [
        "Computadoras y Laptops", "Servidores", "Equipo con Conexión a Internet", "Equipo sin Conexión a Internet"
    ]
    columnas_grafica_usuarios = [
        "Alumnos", "Docentes", "Docentes en Labores de Investigación",
        "Personal Directivo y de Mando", "PAAE"
    ]
    col_graf1, col_graf2 = st.columns(2)
    with col_graf1:
        st.subheader("📊 Distribución de Equipo")
        equipo_elegido = st.selectbox(
            "Selecciona el tipo de equipo:", options=columnas_grafica_equipo)

        if col_categoria in df_ipn.columns and equipo_elegido in df_ipn.columns:
            # Filtrar la fila TOTAL para que no salga en la gráfica
            df_grafica_equipo = df_ipn[df_ipn[col_categoria] != "TOTAL"]
            fig_equipo = px.pie(
                df_grafica_equipo, names=col_categoria, values=equipo_elegido,
                title=f"Distribución de '{equipo_elegido}' por Nivel"
            )
            fig_equipo.update_traces(
                textposition='inside', textinfo='percent+value+label')
            # --- CAMBIO AQUÍ ---
            st.plotly_chart(fig_equipo, use_container_width=True)
            # --- FIN DEL CAMBIO ---

    with col_graf2:
        st.subheader("👤 Distribución de Usuarios")
        usuario_elegido = st.selectbox(
            "Selecciona el tipo de usuario:", options=columnas_grafica_usuarios)

        if col_categoria in df_ipn.columns and usuario_elegido in df_ipn.columns:
            # Filtrar la fila TOTAL para que no salga en la gráfica
            df_grafica_usuarios = df_ipn[df_ipn[col_categoria] != "TOTAL"]
            fig_usuarios = px.pie(
                df_grafica_usuarios, names=col_categoria, values=usuario_elegido,
                title=f"Distribución de '{usuario_elegido}' por Nivel"
            )
            fig_usuarios.update_traces(
                textposition='inside', textinfo='percent+value+label')
            # --- CAMBIO AQUÍ ---
            st.plotly_chart(fig_usuarios, use_container_width=True)
            # --- FIN DEL CAMBIO ---


# ---------------------------
# FUNCIÓN GENERAL PARA RENDERIZAR PESTAÑAS (CON FILTROS OCULTOS)
# ---------------------------
def render_filtered_tab(df, sheet_name):
    """Crea el layout con filtros, datos y gráficas para una pestaña específica."""
    st.subheader(f"Análisis de: {sheet_name}")

    col_rama = next((c for c in df.columns if "Rama" in str(c)), None)
    col_unidad = next(
        (c for c in df.columns if "UNIDAD" in str(c) or "Acad" in str(c)), None)
    rama_labels = {1: "ICFM", 2: "CMB", 3: "CSA", 4: "Interdisciplinaria"}

    if not col_rama or not col_unidad:
        st.warning(
            f"⚠️ No se encontraron las columnas 'Rama' o 'Unidad' en la hoja '{sheet_name}'.")
        st.dataframe(df)
        return

    # --- INICIO DE CAMBIOS: FILTROS OCULTOS ---
    # ... (toda la sección de filtros permanece comentada) ...

    # Dataframe final ahora usa el df completo...
    df_final = df.copy()  # Usamos .copy() para evitar advertencias

    # --- ¡NUEVA LÍNEA CLAVE! ---
    # Excluimos todas las filas que contengan "TOTAL" en la columna de unidad
    if col_unidad in df_final.columns:
        df_final = df_final[~df_final[col_unidad].astype(
            str).str.contains("TOTAL", case=False, na=False)]
    # --- FIN DE LA LÍNEA CLAVE ---

    # --- FIN DE CAMBIOS: FILTROS ---

    st.info(
        f"Mostrando **{len(df_final)}** de **{len(df)}** registros totales.")

    # --- INICIO DE CAMBIOS: TABLAS SEPARADAS ---
    columnas_equipo_tabla = [
        col_unidad,
        "Computadoras y Laptops",
        "Servidores"
    ]
    columnas_usuarios_tabla = [
        col_unidad,
        "Alumnos",
        "Docentes",
        "Docentes en Labores de Investigación",
        "Personal Directivo y de Mando",
        "PAAE"
    ]

    col_tabla1, col_tabla2 = st.columns(2)

    with col_tabla1:
        st.write("##### Resumen de Equipo")
        cols_mostrar_equipo = [
            col for col in columnas_equipo_tabla if col in df_final.columns]
        if col_unidad in cols_mostrar_equipo:
            st.dataframe(df_final[cols_mostrar_equipo],
                         width='stretch', hide_index=True)
        else:
            st.warning(
                "No se encontró la columna de Unidad Académica para la tabla de equipo.")

    with col_tabla2:
        st.write("##### Resumen de Usuarios")
        cols_mostrar_usuarios = [
            col for col in columnas_usuarios_tabla if col in df_final.columns]
        if col_unidad in cols_mostrar_usuarios:
            st.dataframe(df_final[cols_mostrar_usuarios],
                         width='stretch', hide_index=True)
        else:
            st.warning(
                "No se encontró la columna de Unidad Académica para la tabla de usuarios.")
    # --- FIN DE CAMBIOS: TABLAS SEPARADAS ---

    # --- INICIO DE SECCIÓN DE GRÁFICAS (CON CORRECCIÓN DE ERROR) ---
    st.markdown("---")
    st.subheader("📊 Porcentaje de Cómputo por Rama")
    col_computo = "Computadoras y Laptops"
    if col_computo not in df_final.columns:
        st.warning(
            "⚠️ No se encontró la columna 'Computadoras y Laptops' para graficar.")
    elif len(df_final) == 0:
        st.warning(
            "⚠️ No hay datos seleccionados para mostrar en la gráfica de pastel.")
    else:
        # --- CORRECCIÓN DE ERROR (LOCAL) ---
        df_final[col_computo] = pd.to_numeric(
            df_final[col_computo], errors='coerce').fillna(0)

        df_agrupado_rama = df_final.groupby(
            col_rama)[col_computo].sum().reset_index()
        df_agrupado_rama['Rama_Label'] = df_agrupado_rama[col_rama].map(
            rama_labels).fillna("Desconocida")
        fig_pie_rama = px.pie(
            df_agrupado_rama,
            names='Rama_Label',
            values=col_computo,
            title=f"Distribución de '{col_computo}' por Rama"
        )
        fig_pie_rama.update_traces(
            textposition='inside', textinfo='percent+value+label')
        st.plotly_chart(fig_pie_rama, use_container_width=True)

    st.markdown("---")
    st.subheader("Distribución de Tipos de Equipo por Unidad Académica")

    columnas_equipo_grafica = [
        "Computadoras y Laptops",
        "Servidores"
    ]

    columnas_equipo_existentes = [
        col for col in columnas_equipo_grafica if col in df_final.columns]
    if not columnas_equipo_existentes:
        st.warning("⚠️ No se encontraron columnas de equipo para graficar.")
    else:
        df_equipo_melted = df_final.melt(
            id_vars=col_unidad,
            value_vars=columnas_equipo_existentes,
            var_name="Tipo de Equipo",
            value_name="Cantidad"
        )

        # --- CORRECCIÓN DE ERROR (LOCAL) ---
        df_equipo_melted['Cantidad'] = pd.to_numeric(
            df_equipo_melted['Cantidad'], errors='coerce').fillna(0)

        # --- BLOQUE DE CÓDIGO ELIMINADO ---
        # Ya no pre-calculamos el orden aquí
        # df_equipo_totals = ...
        # sorted_unidades_equipo = ...
        # --- FIN DE BLOQUE ELIMINADO ---

        fig_equipo = px.bar(
            df_equipo_melted,
            x=col_unidad,
            y="Cantidad",
            color="Tipo de Equipo",
            title="Composición del Equipo de Cómputo por Unidad Académica",
            labels={col_unidad: "Unidad Académica", "Cantidad": "Unidades"}
            # Se eliminó el parámetro 'category_orders'
        )
        # --- CAMBIO AQUÍ: Añadimos el orden dinámico de Plotly ---
        fig_equipo.update_xaxes(tickangle=45, categoryorder="total descending")
        # --- FIN DEL CAMBIO ---
        st.plotly_chart(fig_equipo, use_container_width=True)

    st.markdown("---")
    st.subheader("Distribución de Usuarios por Unidad Académica")
    columnas_usuarios_grafica = [
        "Alumnos", "Docentes", "Docentes en Labores de Investigación", "Personal Directivo y de Mando", "PAAE"
    ]
    columnas_usuarios_existentes = [
        col for col in columnas_usuarios_grafica if col in df_final.columns]
    if not columnas_usuarios_existentes:
        st.warning("⚠️ No se encontraron columnas de usuarios para graficar.")
    else:
        df_usuarios_melted = df_final.melt(
            id_vars=col_unidad,
            value_vars=columnas_usuarios_existentes,
            var_name="Tipo de Usuario",
            value_name="Cantidad"
        )

        # --- CORRECCIÓN DE ERROR (LOCAL) ---
        df_usuarios_melted['Cantidad'] = pd.to_numeric(
            df_usuarios_melted['Cantidad'], errors='coerce').fillna(0)

        # --- BLOQUE DE CÓDIGO ELIMINADO ---
        # Ya no pre-calculamos el orden aquí
        # df_usuarios_totals = ...
        # sorted_unidades_usuarios = ...
        # --- FIN DE BLOQUE ELIMINADO ---

        fig_barras_usuarios = px.bar(
            df_usuarios_melted,
            x=col_unidad,
            y="Cantidad",
            color="Tipo de Usuario",
            title="Composición de Usuarios por Unidad Académica",
            labels={col_unidad: "Unidad Académica",
                    "Cantidad": "Número de Usuarios"}
            # Se eliminó el parámetro 'category_orders'
        )
        # --- CAMBIO AQUÍ: Añadimos el orden dinámico de Plotly ---
        fig_barras_usuarios.update_xaxes(
            tickangle=45, categoryorder="total descending")
        # --- FIN DEL CAMBIO ---
        st.plotly_chart(fig_barras_usuarios, use_container_width=True)


# ---------------------------
# Bucle para crear el resto de las pestañas
# ---------------------------
for i, tab in enumerate(tabs[1:]):  # Empezar desde la segunda pestaña
    with tab:
        sheet_name = tab_names[i+1]  # El índice correcto de la hoja
        df_active = sheets[sheet_name].copy()
        render_filtered_tab(df_active, sheet_name)
