import streamlit as st
import pandas as pd
import os
import io

FILE_PERSISTENCIA = "datos_alumnos.csv"

def leer_dataframe_robusto(fuente):
    nombre = getattr(fuente, 'name', str(fuente)).lower()

    if nombre.endswith(('.xlsx', '.xls')):
        try:
            return pd.read_excel(fuente)
        except Exception:
            pass

    encodings = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    separators = [None, ';', ',', '\t']

    for enc in encodings:
        for sep in separators:
            try:
                if hasattr(fuente, 'seek'):
                    fuente.seek(0)
                if sep is None:
                    df = pd.read_csv(fuente, encoding=enc, sep=sep, engine='python')
                else:
                    df = pd.read_csv(fuente, encoding=enc, sep=sep)
                
                if len(df.columns) >= 1 and not df.empty:
                    return df
            except Exception:
                continue

    if hasattr(fuente, 'seek'):
        fuente.seek(0)
    return pd.read_excel(fuente)


def cargar_datos_disco():
    if os.path.exists(FILE_PERSISTENCIA):
        return leer_dataframe_robusto(FILE_PERSISTENCIA)
    return None


def guardar_datos_disco(df):
    if df is not None and not df.empty:
        df.to_csv(FILE_PERSISTENCIA, index=False, encoding='utf-8-sig')


# Configuracion de pagina
st.set_page_config(page_title="Facturacion Servicios Escolares", layout="wide")

st.title("Aplicacion de Facturacion de Servicios Escolares")

# 1. Cargar datos iniciales en la sesion si no existen
if 'df_master' not in st.session_state or st.session_state['df_master'] is None:
    st.session_state['df_master'] = cargar_datos_disco()

# --- BARRA LATERAL ---
st.sidebar.header("Gestion de Datos")

uploaded_file = st.sidebar.file_uploader(
    "Cargar/Reemplazar archivo de Alumnos (Excel o CSV)", 
    type=["xlsx", "xls", "csv"]
)

if uploaded_file is not None:
    try:
        df_nuevo = leer_dataframe_robusto(uploaded_file)
        df_nuevo.columns = [str(c).strip() for c in df_nuevo.columns]
        
        cols_requeridas = ['IdAlumno', 'Apellido1', 'Apellido2', 'Nombre', 'Clase', 'IdFamilia']
        missing_cols = [c for c in cols_requeridas if c not in df_nuevo.columns]
        
        if missing_cols:
            st.sidebar.error(f"Faltan columnas requeridas: {', '.join(missing_cols)}")
        else:
            columnas_dias = ['Dias_Comedor', 'Dias_Mad_Antes815', 'Dias_Mad_Despues815', 'Dias_Continuadores']
            for col in columnas_dias:
                if col not in df_nuevo.columns:
                    df_nuevo[col] = 0
                else:
                    df_nuevo[col] = df_nuevo[col].fillna(0).astype(int)

            st.session_state['df_master'] = df_nuevo
            guardar_datos_disco(df_nuevo)
            st.sidebar.success("Archivo cargado y guardado correctamente.")
    except Exception as e:
        st.sidebar.error(f"Error al leer el archivo: {e}")

if st.sidebar.button("Borrar todos los datos guardados"):
    if os.path.exists(FILE_PERSISTENCIA):
        os.remove(FILE_PERSISTENCIA)
    st.session_state['df_master'] = None
    st.sidebar.warning("Datos eliminados correctamente.")
    st.rerun()

# --- CONTENIDO PRINCIPAL ---
if st.session_state['df_master'] is None or st.session_state['df_master'].empty:
    st.info("Carga tu archivo de alumnos (Excel o CSV) desde la barra lateral izquierda para empezar.")
else:
    # 2. Selector de filtro por clase
    lista_clases = ["Todas las clases"] + sorted([str(c) for c in st.session_state['df_master']['Clase'].dropna().unique()])
    
    clase_seleccionada = st.selectbox("Filtrar por Clase:", lista_clases)

    if clase_seleccionada == "Todas las clases":
        df_mostrar = st.session_state['df_master'].copy()
    else:
        df_mostrar = st.session_state['df_master'][st.session_state['df_master']['Clase'].astype(str) == clase_seleccionada].copy()

    st.subheader(f"Introducir dias ({len(df_mostrar)} alumnos)")

    # 3. Formulario para editar la tabla sin recargar la pagina al pulsar Enter
    columnas_dias = ['Dias_Comedor', 'Dias_Mad_Antes815', 'Dias_Mad_Despues815', 'Dias_Continuadores']

    with st.form(key=f"form_tabla_{clase_seleccionada}"):
        df_editado = st.data_editor(
            df_mostrar,
            num_rows="dynamic",
            column_config={
                "Dias_Comedor": st.column_config.NumberColumn("Comedor (9,80 EUR/d | max 133 EUR)", min_value=0, step=1),
                "Dias_Mad_Antes815": st.column_config.NumberColumn("Madrug. <08:15 (5 EUR/d | max 58 EUR)", min_value=0, step=1),
                "Dias_Mad_Despues815": st.column_config.NumberColumn("Madrug. >08:15 (3 EUR/d | max 25 EUR)", min_value=0, step=1),
                "Dias_Continuadores": st.column_config.NumberColumn("Continuadores (3 EUR/d | max 25 EUR)", min_value=0, step=1),
            },
            use_container_width=True,
            key=f"editor_{clase_seleccionada}"
        )

        boton_guardar = st.form_submit_button("?? Guardar Cambios de esta Tabla", type="primary")

        if boton_guardar:
            for col in columnas_dias:
                if col in df_editado.columns:
                    st.session_state['df_master'].loc[df_editado.index, col] = df_editado[col]
            
            guardar_datos_disco(st.session_state['df_master'])
            st.success("Cambios guardados con exito.")

    st.markdown("---")

    # 4. Calculo global y exportacion
    if st.button("Calcular Facturacion Total y Generar Excel", type="primary"):
        df_calc = st.session_state['df_master'].copy()
        df_calc['Orden_Hermano'] = df_calc.groupby('IdFamilia').cumcount() + 1

        importes_finales = []

        for _, row in df_calc.iterrows():
            d_com = float(row.get('Dias_Comedor', 0) or 0)
            d_m_ant = float(row.get('Dias_Mad_Antes815', 0) or 0)
            d_m_des = float(row.get('Dias_Mad_Despues815', 0) or 0)
            d_cont = float(row.get('Dias_Continuadores', 0) or 0)
            orden_hermano = row['Orden_Hermano']

            c_com = min(d_com * 9.80, 133.0)
            c_m_ant = min(d_m_ant * 5.00, 58.0)
            c_m_des = min(d_m_des * 3.00, 25.0)
            c_cont = min(d_cont * 3.00, 25.0)

            if d_m_ant > 0 or d_m_des > 0:
                c_cont = 0.0

            if (c_com + c_m_ant) > 175.0:
                factor = 175.0 / (c_com + c_m_ant)
                c_com *= factor
                c_m_ant *= factor

            if (c_com + c_m_des) > 150.0:
                factor = 150.0 / (c_com + c_m_des)
                c_com *= factor
                c_m_des *= factor

            total_base = c_com + c_m_ant + c_m_des + c_cont

            if orden_hermano == 1:
                factor_desc = 1.00
            elif orden_hermano in [2, 3]:
                factor_desc = 0.80
            else:
                factor_desc = 0.70

            total_final = round(total_base * factor_desc, 2)
            importes_finales.append(total_final)

        df_calc['Importe'] = importes_finales

        cols_exportar = ['Apellido1', 'Apellido2', 'Nombre', 'Clase', 'IdAlumno', 'Importe']
        df_resultado = df_calc[cols_exportar]

        st.subheader("Resultado Final de Facturacion")
        st.dataframe(df_resultado, use_container_width=True)

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False, sheet_name='Facturacion')

        st.download_button(
            label="Descargar Excel de Facturacion Final",
            data=buffer.getvalue(),
            file_name="Facturacion_Servicios.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )