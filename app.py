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
# PESTAÑA 1: Resumen General (con 2 gráficas dinámicas)
# ---------------------------
with tabs[0]:
    st.subheader(f"Resumen General - {tab_names[0]}")
    df_ipn = sheets[tab_names[0]]

    # Mostrar la tabla de datos principal en la parte superior
    st.dataframe(df_ipn.reset_index(drop=True), use_container_width=True)
    st.markdown("---")  # Separador visual

    # --- Definir las columnas para cada categoría ---
    col_categoria = "IPN"  # Esta es la columna para las "rebanadas" del pastel

    columnas_equipo = [
        "Computadoras y Laptops",
        "Servidores",
        "Equipo con Conexión a Internet"
    ]

    columnas_usuarios = [
        "Alumnos",
        "Docentes",
        "Docentes en Labores de Investigación",
        "Personal Directivo y de Mando",
        "Personal de Apoyo y Asistencia a la Educación"
    ]

    # --- Crear el layout con dos columnas para las gráficas ---
    col_equipo, col_usuarios = st.columns(2)

    # --- COLUMNA 1: GRÁFICA DE EQUIPO ---
    with col_equipo:
        st.subheader("📊 Distribución de Equipo")

        # Selector para el tipo de equipo
        equipo_elegido = st.selectbox(
            "Selecciona el tipo de equipo:",
            options=columnas_equipo
        )

        # Generar la gráfica de pastel para el equipo
        if col_categoria in df_ipn.columns and equipo_elegido in df_ipn.columns:
            fig_equipo = px.pie(
                df_ipn,
                names=col_categoria,
                values=equipo_elegido,
                title=f"Distribución de '{equipo_elegido}' por Nivel"
            )
            fig_equipo.update_traces(
                textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_equipo, use_container_width=True)
        else:
            st.warning("⚠️ Columnas para la gráfica de equipo no encontradas.")

    # --- COLUMNA 2: GRÁFICA DE USUARIOS ---
    with col_usuarios:
        st.subheader("👤 Distribución de Usuarios")

        # Selector para el tipo de usuario
        usuario_elegido = st.selectbox(
            "Selecciona el tipo de usuario:",
            options=columnas_usuarios
        )

        # Generar la gráfica de pastel para los usuarios
        if col_categoria in df_ipn.columns and usuario_elegido in df_ipn.columns:
            fig_usuarios = px.pie(
                df_ipn,
                names=col_categoria,
                values=usuario_elegido,
                title=f"Distribución de '{usuario_elegido}' por Nivel"
            )
            fig_usuarios.update_traces(
                textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_usuarios, use_container_width=True)
        else:
            st.warning(
                "⚠️ Columnas para la gráfica de usuarios no encontradas.")


# ---------------------------
# FUNCIÓN GENERAL PARA RENDERIZAR PESTAÑAS CON FILTROS
# ---------------------------


# ---------------------------
# FUNCIÓN GENERAL PARA RENDERIZAR PESTAÑAS CON FILTROS Y GRÁFICAS
# ---------------------------
def render_filtered_tab(df, sheet_name):
    """Crea el layout con filtros, datos y gráficas para una pestaña específica."""
    st.subheader(f"Análisis de: {sheet_name}")

    # Detectar dinámicamente las columnas de interés
    col_rama = next((c for c in df.columns if "Rama" in str(c)), None)
    col_unidad = next(
        (c for c in df.columns if "UNIDAD" in str(c) or "Acad" in str(c)), None)

    if not col_rama or not col_unidad:
        st.warning(
            f"⚠️ No se encontraron las columnas 'Rama' o 'Unidad' en la hoja '{sheet_name}'.")
        st.dataframe(df)
        return

    # --- Layout de la pestaña: Columna de filtros y columna de datos ---
    filter_col, data_col = st.columns([1, 3])

    # --- COLUMNA DE FILTROS ---
    with filter_col:
        st.header("Filtros")

        # Mapeo de Rama
        rama_labels = {1: "ICFM", 2: "CMB", 3: "CSA", 4: "Interdisciplinaria"}

        # Filtro de RAMA
        ramas_en_hoja = df[col_rama].dropna().unique()
        opciones_rama_labels = [rama_labels.get(
            r, r) for r in sorted(ramas_en_hoja) if r in rama_labels]
        rama_sel_labels = st.multiselect(
            "Selecciona Rama(s)",
            options=opciones_rama_labels,
            default=opciones_rama_labels,
            key=f"rama_{sheet_name}"
        )

        rama_values_filtrar = [
            k for k, v in rama_labels.items() if v in rama_sel_labels]
        df_filtrado_rama = df[df[col_rama].isin(rama_values_filtrar)]

        # Filtro de ESCUELA
        opciones_escuela = sorted(
            df_filtrado_rama[col_unidad].dropna().unique())
        unidad_sel = st.multiselect(
            "Selecciona Escuela(s)",
            options=opciones_escuela,
            default=opciones_escuela,
            key=f"unidad_{sheet_name}"
        )

        df_final = df_filtrado_rama[df_filtrado_rama[col_unidad].isin(
            unidad_sel)]

    # --- COLUMNA DE DATOS Y GRÁFICAS ---
    with data_col:
        st.info(
            f"Mostrando **{len(df_final)}** de **{len(df)}** registros totales.")
        df_a_mostrar = df_final.drop(columns=[col_rama], errors='ignore')
        st.dataframe(df_a_mostrar.reset_index(
            drop=True), use_container_width=True)

        # --- GRÁFICA 1: DISTRIBUCIÓN DE EQUIPO ---
        st.markdown("---")
        st.subheader("Distribución de Tipos de Equipo por Unidad Académica")

        columnas_equipo = [
            "Computadoras y Laptops",
            "Servidores",
            "Equipo con Conexión a Internet"
        ]
        columnas_equipo_existentes = [
            col for col in columnas_equipo if col in df_final.columns]

        if not columnas_equipo_existentes:
            st.warning("⚠️ No se encontraron columnas de equipo para graficar.")
        else:
            df_equipo_melted = df_final.melt(
                id_vars=col_unidad,
                value_vars=columnas_equipo_existentes,
                var_name="Tipo de Equipo",
                value_name="Cantidad"
            )
            fig_equipo = px.bar(
                df_equipo_melted,
                x=col_unidad,
                y="Cantidad",
                color="Tipo de Equipo",
                title="Composición del Equipo de Cómputo por Unidad Académica",
                labels={col_unidad: "Unidad Académica", "Cantidad": "Unidades"}
            )
            fig_equipo.update_xaxes(tickangle=45)
            st.plotly_chart(fig_equipo, use_container_width=True)

        # --- GRÁFICA 2: DISTRIBUCIÓN DE USUARIOS ---
        st.markdown("---")
        st.subheader("Distribución de Usuarios por Unidad Académica")

        columnas_usuarios = [
            "Alumnos", "Docentes", "Personal Directivo y de Mando",
            "Personal de Apoyo y Asistencia a la Educación"
        ]
        columnas_usuarios_existentes = [
            col for col in columnas_usuarios if col in df_final.columns]

        if not columnas_usuarios_existentes:
            st.warning(
                "⚠️ No se encontraron columnas de usuarios para graficar.")
        else:
            df_usuarios_melted = df_final.melt(
                id_vars=col_unidad,
                value_vars=columnas_usuarios_existentes,
                var_name="Tipo de Usuario",
                value_name="Cantidad"
            )
            fig_barras_usuarios = px.bar(
                df_usuarios_melted,
                x=col_unidad,
                y="Cantidad",
                color="Tipo de Usuario",
                title="Composición de Usuarios por Unidad Académica",
                labels={col_unidad: "Unidad Académica",
                        "Cantidad": "Número de Usuarios"}
            )
            fig_barras_usuarios.update_xaxes(tickangle=45)
            st.plotly_chart(fig_barras_usuarios, use_container_width=True)


# ---------------------------
# Bucle para crear el resto de las pestañas
# ---------------------------
for i, tab in enumerate(tabs[1:]):  # Empezar desde la segunda pestaña
    with tab:
        sheet_name = tab_names[i+1]  # El índice correcto de la hoja
        df_active = sheets[sheet_name].copy()
        render_filtered_tab(df_active, sheet_name)
