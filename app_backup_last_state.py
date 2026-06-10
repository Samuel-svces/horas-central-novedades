# -*- coding: utf-8 -*-
"""
Interfaz de Usuario en Streamlit para el Dashboard de Auditoría de Médicos Supernumerarios.
"""

import streamlit as st
import pandas as pd
import os
import io
from datetime import datetime
import data_processor as dp

# Configuración de página
st.set_page_config(
    page_title="Control de horas Central de novedades",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para imitar el panel premium del usuario
custom_css = """
<style>
    /* Tipografía y fondo */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Reducir espacio superior e inferior del contenedor principal */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 1rem !important;
    }
    
    /* Ocultar la cabecera predeterminada de Streamlit para ganar espacio */
    [data-testid="stHeader"] {
        display: none !important;
    }
    
    /* Contenedor de KPIs */
    .kpi-container {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin: 15px 0px 25px 0px;
    }
    
    /* Cajas de KPI individuales estilo M365 */
    .kpi-card {
        padding: 10px 16px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        flex: 1;
        min-width: 180px;
        border-top: 1px solid rgba(0,0,0,0.05);
        border-bottom: 1px solid rgba(0,0,0,0.05);
        border-right: 1px solid rgba(0,0,0,0.05);
        background-color: #ffffff;
    }
    
    .kpi-blue { border-left: 6px solid #1e5cc8; background-color: #f0f4fc; }
    .kpi-green { border-left: 6px solid #1e8e3e; background-color: #f4faf6; }
    .kpi-yellow { border-left: 6px solid #f9ab00; background-color: #fefcf3; }
    .kpi-red { border-left: 6px solid #d93025; background-color: #fdf5f4; }
    
    .kpi-title {
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .kpi-blue .kpi-title { color: #1e5cc8; }
    .kpi-green .kpi-title { color: #1e8e3e; }
    .kpi-yellow .kpi-title { color: #b06000; }
    .kpi-red .kpi-title { color: #d93025; }
    
    .kpi-value {
        font-size: 22px;
        font-weight: 700;
        color: #202124;
    }
    
    /* Contenedor de filtros */
    .filter-panel {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #e0e0e0;
        margin-bottom: 20px;
    }
    
    /* Estilo de botones personalizados */
    div.stButton > button {
        border-radius: 4px;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    
    /* Botón de limpiar */
    .clear-btn button {
        background-color: #fff5f5 !important;
        color: #e03131 !important;
        border: 1px solid #ffc9c9 !important;
        width: 100%;
        height: 40px;
        margin-top: 28px;
    }
    .clear-btn button:hover {
        background-color: #fae0e0 !important;
        border-color: #fa5252 !important;
    }
    
    /* Botón de exportar */
    .export-btn button {
        background-color: #f4faf6 !important;
        color: #1e8e3e !important;
        border: 1px solid #a3e2bc !important;
        width: 100%;
        height: 40px;
        margin-top: 28px;
    }
    .export-btn button:hover {
        background-color: #e6f6ec !important;
        border-color: #1e8e3e !important;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# Inicializar variables en st.session_state
if 'mes_sel' not in st.session_state:
    st.session_state.mes_sel = []
if 'cedula_sel' not in st.session_state:
    st.session_state.cedula_sel = "Todas"
if 'nombre_sel' not in st.session_state:
    st.session_state.nombre_sel = "Todos"

if 'file_path_input' not in st.session_state:
    st.session_state.file_path_input = r"C:\Users\JuanJoseOsorioMolina\OneDrive - U.T SAN VICENTE CES\CONSOLIDADO 2026.xlsx"
if 'data_loaded' not in st.session_state:
    st.session_state.df_raw = None
    st.session_state.load_error = None

# Función para limpiar filtros
def reset_filters():
    st.session_state.mes_sel = []
    st.session_state.cedula_sel = "Todas"
    st.session_state.nombre_sel = "Todos"


# Cabecera de la aplicación
st.markdown("<h1 style='font-family: \"Segoe UI\", -apple-system, BlinkMacSystemFont, Roboto, sans-serif; font-weight: 400; color: #0b3c5d; font-size: 38px; margin-top: 0px; margin-bottom: 15px;'>Control de horas Central de novedades</h1>", unsafe_allow_html=True)

# Expander de configuración de datos
with st.sidebar:
    st.header("⚙️ Configuración")
    
    data_source_mode = st.radio(
        "Origen del Archivo Excel:",
        ["Ruta Local / OneDrive", "Cargar Archivo Manualmente"]
    )
    
    df_loaded = None
    
    if data_source_mode == "Ruta Local / OneDrive":
        file_path = st.text_input(
            "Ruta del archivo local (.xlsx):",
            value=st.session_state.file_path_input,
            key="file_path_input"
        )
        
        # Botón para cargar/recargar desde ruta local
        if st.button("🔄 Cargar desde Ruta Local", use_container_width=True) or st.session_state.df_raw is None:
            if os.path.exists(file_path):
                with st.spinner("Leyendo archivo de Excel local..."):
                    try:
                        st.session_state.df_raw = dp.load_and_clean_data(file_path)
                        # Cargar metas de calendario
                        m_targets, d_targets = dp.load_calendar_targets(file_path)
                        st.session_state.monthly_targets = m_targets
                        st.session_state.daily_targets = d_targets
                        st.session_state.load_error = None
                        st.toast("¡Datos y metas cargados con éxito!", icon="✅")
                    except Exception as e:
                        st.session_state.df_raw = None
                        st.session_state.load_error = str(e)
            else:
                st.session_state.df_raw = None
                st.session_state.load_error = f"No se encontró el archivo en la ruta especificada:\n{file_path}\n\nPor favor, verifica que el archivo exista o súbelo manualmente con la opción 'Cargar Archivo Manualmente'."
    
    else:
        uploaded_file = st.file_uploader(
            "Selecciona el archivo Excel consolidado:",
            type=["xlsx", "xls"],
            help="Sube el archivo 'CONSOLIDADO 2026.xlsx' directamente"
        )
        if uploaded_file is not None:
            with st.spinner("Procesando archivo subido..."):
                try:
                    # Usar BytesIO para leer en memoria
                    file_bytes = io.BytesIO(uploaded_file.read())
                    st.session_state.df_raw = dp.load_and_clean_data(file_bytes)
                    # Reiniciar puntero para leer metas
                    file_bytes.seek(0)
                    m_targets, d_targets = dp.load_calendar_targets(file_bytes)
                    st.session_state.monthly_targets = m_targets
                    st.session_state.daily_targets = d_targets
                    st.session_state.load_error = None
                except Exception as e:
                    st.session_state.df_raw = None
                    st.session_state.load_error = str(e)
        else:
            st.session_state.df_raw = None
            st.session_state.load_error = "Por favor, sube un archivo de Excel para comenzar."


    st.markdown("---")
    st.markdown("""
    **Instrucciones corporativas:**
    El archivo Excel se maneja en la nube de SharePoint. Si utilizas la sincronización de OneDrive/SharePoint, puedes ingresar la ruta de tu carpeta local sincronizada para que los datos siempre estén actualizados en tiempo real al abrir la aplicación.
    """)

# Mostrar alerta si hay error de carga
if st.session_state.load_error:
    st.error(f"⚠️ Error de Carga de Datos: {st.session_state.load_error}")
    st.stop()

# Si no hay datos cargados, parar ejecución
if st.session_state.df_raw is None:
    st.info("ℹ️ Esperando la carga de datos. Por favor configura el archivo de datos en el panel lateral.")
    st.stop()

# Datos cargados con éxito
df_raw = st.session_state.df_raw



# === 2. APLICAR FILTROS A LA DATA ===
df_filtrado = df_raw.copy()

# A) Filtrar por Mes
if st.session_state.mes_sel:
    df_filtrado = df_filtrado[df_filtrado['MES'].isin(st.session_state.mes_sel)]

# B) Filtrar por Cédula
if st.session_state.cedula_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado['CEDULA_FINAL'] == st.session_state.cedula_sel]

# C) Filtrar por Nombre
if st.session_state.nombre_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado['NOMBRE SUPER VALIDADO'] == st.session_state.nombre_sel]


# === 3. SECCIÓN DE MÉTRICAS (KPI BAR) ===
total_horas_periodo = df_filtrado['HORAS TOTALES DECIMAL'].sum()
total_medicos_activos = df_filtrado['NOMBRE SUPER VALIDADO'].nunique()
total_registros_novedades = len(df_filtrado)
promedio_horas_medico = total_horas_periodo / total_medicos_activos if total_medicos_activos > 0 else 0.0

# Renderizar KPI Bar utilizando HTML/CSS personalizados para aproximarse a la visual solicitada
st.markdown(f"""
<div class="kpi-container">
    <div class="kpi-card kpi-blue">
        <div class="kpi-title">MÉDICOS ACTIVOS</div>
        <div class="kpi-value">{total_medicos_activos}</div>
    </div>
    <div class="kpi-card kpi-green">
        <div class="kpi-title">TOTAL HORAS REGISTRADAS</div>
        <div class="kpi-value">{total_horas_periodo:,.2f} hrs</div>
    </div>
    <div class="kpi-card kpi-yellow">
        <div class="kpi-title">PROMEDIO HORAS POR MÉDICO</div>
        <div class="kpi-value">{promedio_horas_medico:,.2f} hrs</div>
    </div>
    <div class="kpi-card kpi-red">
        <div class="kpi-title">CANTIDAD DE NOVEDADES</div>
        <div class="kpi-value">{total_registros_novedades}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# === 4. TABLAS E INTERFAZ PRINCIPAL ===
tab1, tab2, tab3 = st.tabs([
    "📊 Resumen Consolidado (Día / Mes)",
    "📑 Detalle de Novedades (Auditoría Completa)",
    "📈 Gráficos Analíticos"
])

# Obtener tabla pivote agrupada (mensual para gráficos y operaciones globales)
tabla_consolidada = dp.get_consolidated_hours(df_filtrado)

with tab1:
    # --- PANEL DE FILTROS HORIZONTALES (Movido aquí) ---
    meses_disponibles = sorted(df_raw['MES'].dropna().unique().tolist(), key=lambda m: list(dp.MESES_MAP.values()).index(m) if m in dp.MESES_MAP.values() else 99)

    f_col1, f_col2, f_col3, f_col4, f_col5 = st.columns([2.5, 2.5, 3.5, 1.5, 1.5])

    with f_col1:
        meses_sel = st.multiselect(
            "📅 Mes(es):",
            options=meses_disponibles,
            default=st.session_state.mes_sel,
            key="mes_sel_widget",
            help="Selecciona uno o más meses. Si queda vacío, se asumen TODOS los meses."
        )
        st.session_state.mes_sel = meses_sel

    df_para_filtros = df_raw.copy()
    if st.session_state.mes_sel:
        df_para_filtros = df_para_filtros[df_para_filtros['MES'].isin(st.session_state.mes_sel)]

    cedulas_disponibles = ["Todas"] + sorted(df_para_filtros['CEDULA_FINAL'].dropna().unique().tolist())
    nombres_disponibles = ["Todos"] + sorted(df_para_filtros['NOMBRE SUPER VALIDADO'].dropna().unique().tolist())

    if st.session_state.cedula_sel not in cedulas_disponibles:
        st.session_state.cedula_sel = "Todas"
    if st.session_state.nombre_sel not in nombres_disponibles:
        st.session_state.nombre_sel = "Todos"

    with f_col2:
        cedula_idx = 0
        if st.session_state.cedula_sel in cedulas_disponibles:
            cedula_idx = cedulas_disponibles.index(st.session_state.cedula_sel)
            
        cedula_sel = st.selectbox(
            "🆔 Cédula Médica:",
            options=cedulas_disponibles,
            index=cedula_idx,
            key="cedula_sel_widget"
        )
        st.session_state.cedula_sel = cedula_sel

    with f_col3:
        nombre_idx = 0
        if st.session_state.nombre_sel in nombres_disponibles:
            nombre_idx = nombres_disponibles.index(st.session_state.nombre_sel)
            
        nombre_sel = st.selectbox(
            "👨‍⚕️ Nombre del Médico:",
            options=nombres_disponibles,
            index=nombre_idx,
            key="nombre_sel_widget"
        )
        st.session_state.nombre_sel = nombre_sel

    with f_col4:
        st.markdown('<div class="clear-btn">', unsafe_allow_html=True)
        if st.button("🗑️ Borrar Filtros", key="btn_clear"):
            reset_filters()
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with f_col5:
        st.markdown('<div class="export-btn">', unsafe_allow_html=True)
        
        detalle_cols_base = [
            'REVISION POR CENTRAL DE NOVEDADES', 'FECHA_CLEAN', 'NOMBRE SUPER VALIDADO', 'CEDULA_FINAL',
            'MEDICOS', 'DOCUMENTO', 'CIS', 'ZONA', 'TIPO DE NOVEDAD', 'HORAS TOTALES DECIMAL', 'RECARGO NOCTURNO ORDINARIO'
        ]
        cols_to_export_det = [c for c in detalle_cols_base if c in df_filtrado.columns]
        
        output = io.BytesIO()
        
        # 1. Consolidado por Día
        df_export_dia = dp.get_consolidated_hours_by_date(df_filtrado)
        daily_targets = st.session_state.get('daily_targets', {})
        df_export_dia['HORAS_A_LABORAR'] = df_export_dia['FECHA_STR'].map(daily_targets).fillna(0)
        df_export_dia['TOTAL'] = df_export_dia['HORAS_TOTALES'] - df_export_dia['HORAS_A_LABORAR']
        df_export_dia_rename = df_export_dia.rename(columns={
            'FECHA_STR': 'Fecha',
            'CEDULA_FINAL': 'Cédula',
            'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
            'HORAS_A_LABORAR': 'Horas a laborar',
            'HORAS_TOTALES': 'Horas Laboradas',
            'TOTAL': 'Total',
            'CANTIDAD_NOVEDADES': 'Novedades Reportadas'
        })
        df_export_dia_rename = df_export_dia_rename[['Cédula', 'Médico Supernumerario', 'Fecha', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Reportadas']]
        if not df_export_dia_rename.empty:
            totales_dict_dia = {}
            for col in df_export_dia_rename.columns:
                if col in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Reportadas']:
                    totales_dict_dia[col] = [df_export_dia_rename[col].sum()]
                elif col == 'Médico Supernumerario':
                    totales_dict_dia[col] = ['TOTAL GENERAL']
                else:
                    totales_dict_dia[col] = ['']
            df_export_dia_rename = pd.concat([df_export_dia_rename, pd.DataFrame(totales_dict_dia)], ignore_index=True)
        
        # 2. Consolidado por Mes
        df_export_mes = dp.get_consolidated_hours(df_filtrado)
        monthly_targets = st.session_state.get('monthly_targets', {})
        df_export_mes['HORAS_A_LABORAR'] = df_export_mes['MES_NUM'].map(monthly_targets).fillna(0)
        df_export_mes['TOTAL'] = df_export_mes['HORAS_TOTALES'] - df_export_mes['HORAS_A_LABORAR']
        df_export_mes_rename = df_export_mes.rename(columns={
            'CEDULA_FINAL': 'Cédula',
            'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
            'MES': 'Mes',
            'HORAS_A_LABORAR': 'Horas a laborar',
            'HORAS_TOTALES': 'Horas Laboradas',
            'TOTAL': 'Total',
            'CANTIDAD_NOVEDADES': 'Novedades Reportadas'
        })
        df_export_mes_rename = df_export_mes_rename[['Cédula', 'Médico Supernumerario', 'Mes', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Reportadas']]
        if not df_export_mes_rename.empty:
            totales_dict_mes = {}
            for col in df_export_mes_rename.columns:
                if col in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Reportadas']:
                    totales_dict_mes[col] = [df_export_mes_rename[col].sum()]
                elif col == 'Médico Supernumerario':
                    totales_dict_mes[col] = ['TOTAL GENERAL']
                else:
                    totales_dict_mes[col] = ['']
            df_export_mes_rename = pd.concat([df_export_mes_rename, pd.DataFrame(totales_dict_mes)], ignore_index=True)
        
        # 3. Detalle completo de Novedades
        df_export_det = df_filtrado[cols_to_export_det].copy()
        if 'FECHA_CLEAN' in df_export_det.columns:
            df_export_det['Fecha Novedad'] = df_export_det['FECHA_CLEAN'].dt.strftime('%Y-%m-%d')
            df_export_det = df_export_det.drop(columns=['FECHA_CLEAN'])
        df_export_det_rename = df_export_det.rename(columns={
            'REVISION POR CENTRAL DE NOVEDADES': 'Fecha Revisión',
            'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
            'CEDULA_FINAL': 'Cédula Supernumerario',
            'MEDICOS': 'Médico Reemplazado',
            'DOCUMENTO': 'Cédula Reemplazado',
            'CIS': 'Sede CIS',
            'ZONA': 'Zona',
            'TIPO DE NOVEDAD': 'Novedad',
            'HORAS TOTALES DECIMAL': 'Horas',
            'RECARGO NOCTURNO ORDINARIO': 'Recargo Nocturno'
        })

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_export_dia_rename.to_excel(writer, sheet_name='Consolidado por Día', index=False)
            df_export_mes_rename.to_excel(writer, sheet_name='Consolidado por Mes', index=False)
            df_export_det_rename.to_excel(writer, sheet_name='Detalle Completo', index=False)
            
            for sheet_name in ['Consolidado por Día', 'Consolidado por Mes', 'Detalle Completo']:
                worksheet = writer.sheets[sheet_name]
                df_temp = df_export_dia_rename if sheet_name == 'Consolidado por Día' else (df_export_mes_rename if sheet_name == 'Consolidado por Mes' else df_export_det_rename)
                for idx, col in enumerate(df_temp.columns):
                    val_lengths = [len(str(val)) for val in df_temp[col].dropna()]
                    max_val_len = max(val_lengths) if val_lengths else 0
                    max_len = max(max_val_len, len(str(col))) + 2
                    worksheet.set_column(idx, idx, min(max_len, 40))
                
        st.download_button(
            label="📥 Exportar Excel",
            data=output.getvalue(),
            file_name=f"CONSOLIDADO_HORAS_SUPERNUMERARIOS_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_export"
        )
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("---")

    # Selector de tipo de consolidación para cumplir con la visualización solicitada
    agrupacion_vista = st.radio(
        "Agrupar horas por:",
        ["Día (Fecha específica)", "Mes (Consolidado mensual)"],
        index=0, # Día como predeterminado
        horizontal=True,
        help="Permite cambiar entre el acumulado diario detallado por fecha y el acumulado mensual total."
    )
    
    if agrupacion_vista == "Día (Fecha específica)":
        tabla_consolidada_vista = dp.get_consolidated_hours_by_date(df_filtrado)
        
        # Mapear horas metas de días hábiles/festivos
        daily_targets = st.session_state.get('daily_targets', {})
        tabla_consolidada_vista['HORAS_A_LABORAR'] = tabla_consolidada_vista['FECHA_STR'].map(daily_targets).fillna(0)
        # La resta de horas laboradas - horas a laborar (horas_totales es la columna agrupada sumada en Pandas)
        tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
    else:
        tabla_consolidada_vista = tabla_consolidada.copy()
        
        # Mapear horas metas mensuales
        monthly_targets = st.session_state.get('monthly_targets', {})
        tabla_consolidada_vista['HORAS_A_LABORAR'] = tabla_consolidada_vista['MES_NUM'].map(monthly_targets).fillna(0)
        # La resta de horas laboradas - horas a laborar
        tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
        
    st.caption(f"Mostrando {len(tabla_consolidada_vista)} filas en el resumen consolidado.")

    # Formatear columnas según la vista seleccionada
    if agrupacion_vista == "Día (Fecha específica)":
        tabla_display = tabla_consolidada_vista.rename(columns={
            'FECHA_STR': 'Fecha',
            'CEDULA_FINAL': 'Cédula',
            'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
            'HORAS_A_LABORAR': 'Horas a laborar',
            'HORAS_TOTALES': 'Horas Laboradas',
            'TOTAL': 'Total',
            'CANTIDAD_NOVEDADES': 'Novedades Reportadas'
        })
        cols_show = ['Cédula', 'Médico Supernumerario', 'Fecha', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Reportadas']
    else:
        tabla_display = tabla_consolidada_vista.rename(columns={
            'CEDULA_FINAL': 'Cédula',
            'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
            'MES': 'Mes',
            'HORAS_A_LABORAR': 'Horas a laborar',
            'HORAS_TOTALES': 'Horas Laboradas',
            'TOTAL': 'Total',
            'CANTIDAD_NOVEDADES': 'Novedades Reportadas'
        })
        cols_show = ['Cédula', 'Médico Supernumerario', 'Mes', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Reportadas']
    
    # 1. Tabla de Datos Principal (con scroll interno si hay muchos registros)
    st.dataframe(
        tabla_display[cols_show],
        width='stretch',
        height=350,
        hide_index=True
    )
    
    # 2. Resumen de Totales en formato de tarjetas de métricas abajo de la tabla
    if not tabla_display.empty:
        tot_horas_a_laborar = tabla_display['Horas a laborar'].sum()
        tot_horas_laboradas = tabla_display['Horas Laboradas'].sum()
        tot_diferencia = tabla_display['Total'].sum()
        tot_novedades = tabla_display['Novedades Reportadas'].sum()
        
        # Determinar clase de color para la diferencia de forma dinámica
        if tot_diferencia > 0:
            dif_color_class = "kpi-green"
        elif tot_diferencia < 0:
            dif_color_class = "kpi-red"
        else:
            dif_color_class = "kpi-yellow"
            
        st.markdown(f"""
        <div class="kpi-container" style="margin-top: -10px; margin-bottom: 10px;">
            <div class="kpi-card kpi-blue">
                <div class="kpi-title">TOTAL HORAS A LABORAR</div>
                <div class="kpi-value">{tot_horas_a_laborar:,.1f} hrs</div>
            </div>
            <div class="kpi-card kpi-green">
                <div class="kpi-title">TOTAL HORAS LABORADAS</div>
                <div class="kpi-value">{tot_horas_laboradas:,.1f} hrs</div>
            </div>
            <div class="kpi-card {dif_color_class}">
                <div class="kpi-title">DIFERENCIA TOTAL</div>
                <div class="kpi-value">{tot_diferencia:+,.1f} hrs</div>
            </div>
            <div class="kpi-card kpi-red">
                <div class="kpi-title">TOTAL NOVEDADES</div>
                <div class="kpi-value">{tot_novedades:,}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

with tab2:
    st.markdown("### Detalle Individual de Novedades")
    st.caption("Esta vista contiene el detalle fila por fila registrado en la base de datos para auditorías individuales de turnos.")
    
    # Utilizar las columnas predefinidas
    cols_to_show = [c for c in detalle_cols_base if c in df_filtrado.columns]
    
    df_detalle_display = df_filtrado[cols_to_show].copy()
    
    # Formatear la columna fecha para que se vea bonita
    if 'FECHA_CLEAN' in df_detalle_display.columns:
        df_detalle_display['Fecha Novedad'] = df_detalle_display['FECHA_CLEAN'].dt.strftime('%Y-%m-%d')
        # Quitar la columna interna
        df_detalle_display = df_detalle_display.drop(columns=['FECHA_CLEAN'])
        
    df_detalle_display = df_detalle_display.rename(columns={
        'REVISION POR CENTRAL DE NOVEDADES': 'Fecha Revisión',
        'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'CEDULA_FINAL': 'Cédula Supernumerario',
        'MEDICOS': 'Médico Reemplazado',
        'DOCUMENTO': 'Cédula Reemplazado',
        'CIS': 'Sede CIS',
        'ZONA': 'Zona',
        'TIPO DE NOVEDAD': 'Novedad',
        'HORAS TOTALES DECIMAL': 'Horas',
        'RECARGO NOCTURNO ORDINARIO': 'Recargo Nocturno'
    })
    
    st.dataframe(
        df_detalle_display,
        width='stretch',
        hide_index=True
    )

with tab3:
    st.markdown("### Visualizaciones y Análisis de Tendencias")
    
    if not tabla_consolidada.empty:
        # Gráfico 1: Horas totales por Mes
        st.markdown("#### Horas Totales Registradas por Mes")
        
        # Agrupar datos por mes
        horas_por_mes = tabla_consolidada.groupby(['MES', 'MES_NUM'])['HORAS_TOTALES'].sum().reset_index()
        horas_por_mes = horas_por_mes.sort_values(by='MES_NUM')
        
        st.bar_chart(
            data=horas_por_mes,
            x='MES',
            y='HORAS_TOTALES',
            color='#1e5cc8',
            use_container_width=True
        )
        
        # Gráfico 2: Top 10 Médicos con más horas en el periodo seleccionado
        st.markdown("#### Top 10 Médicos Supernumerarios por Horas Totales Acumuladas")
        
        top_medicos = tabla_consolidada.groupby('NOMBRE SUPER VALIDADO')['HORAS_TOTALES'].sum().reset_index()
        top_medicos = top_medicos.sort_values(by='HORAS_TOTALES', ascending=False).head(10)
        
        st.bar_chart(
            data=top_medicos,
            x='NOMBRE SUPER VALIDADO',
            y='HORAS_TOTALES',
            color='#1e8e3e',
            use_container_width=True
        )
    else:
        st.warning("No hay suficientes datos aplicando los filtros actuales para renderizar los gráficos.")
