# -*- coding: utf-8 -*-
"""
Interfaz de Usuario en Streamlit para el Dashboard de Auditoría de Médicos Supernumerarios.
"""

import streamlit as st
import pandas as pd
import os
import io
import platform
from datetime import datetime, timedelta, timezone

def get_local_now():
    return datetime.now(timezone(timedelta(hours=-5)))
import data_processor as dp
import base64
import re

# ── Helpers ───────────────────────────────────────────────────────────────────

def save_to_downloads(data):
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders")
        path, _ = winreg.QueryValueEx(key, "{374DE290-123F-4565-9164-39C4925E467B}")
        downloads_path = os.path.expandvars(path)
    except Exception:
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
    try:
        filename = f"CONSOLIDADO_HORAS_SUPERNUMERARIOS_{get_local_now().strftime('%Y%m%d')}.xlsx"
        full_path = os.path.join(downloads_path, filename)
        with open(full_path, "wb") as f:
            f.write(data)
        st.toast(f"¡Excel guardado en Descargas como: {filename}!", icon="📥")
    except Exception as e:
        st.toast(f"No se pudo guardar en Descargas: {e}", icon="⚠️")


def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        return f"data:image/png;base64,{encoded_string}"
    return ""


def get_onedrive_config():
    """
    Lee las credenciales desde st.secrets (Streamlit Cloud) o variables de entorno (local).
    Retorna None si no están configuradas → la app cae al modo local/manual.
    """
    try:
        return {
            "tenant_id":     st.secrets["AZURE_TENANT_ID"],
            "client_id":     st.secrets["AZURE_CLIENT_ID"],
            "client_secret": st.secrets["AZURE_CLIENT_SECRET"],
            "drive_id":      st.secrets["ONEDRIVE_DRIVE_ID"],
            "file_id":       st.secrets["ONEDRIVE_FILE_ID"],
        }
    except Exception:
        return None


def cargar_desde_onedrive():
    """Descarga el Excel desde OneDrive y lo guarda en session_state."""
    config = get_onedrive_config()
    if not config:
        st.session_state.load_error = (
            "No se encontraron credenciales de OneDrive. "
            "Configura los Secrets en Streamlit Cloud o usa el modo manual."
        )
        return

    try:
        with st.spinner("Conectando con OneDrive..."):
            file_bytes = dp.download_excel_from_onedrive(
                tenant_id=config["tenant_id"],
                client_id=config["client_id"],
                client_secret=config["client_secret"],
                drive_id=config["drive_id"],
                file_id=config["file_id"],
            )
            st.session_state.df_raw = dp.load_and_clean_data(file_bytes)
            file_bytes.seek(0)
            m_targets, d_targets = dp.load_calendar_targets(file_bytes)
            st.session_state.monthly_targets = m_targets
            st.session_state.daily_targets = d_targets
            st.session_state.load_error = None
            st.session_state.last_refresh = get_local_now().strftime('%d/%m/%Y %H:%M:%S')
    except Exception as e:
        st.session_state.df_raw = None
        st.session_state.load_error = str(e)


def calculate_doctor_target_hours(df_grouped, df_raw_filtered, daily_targets, monthly_targets):
    targets = []
    for idx, row in df_grouped.iterrows():
        doc_name = row['NOMBRE SUPER VALIDADO']
        month_num = row['MES_NUM']
        doc_entries = df_raw_filtered[
            (df_raw_filtered['NOMBRE SUPER VALIDADO'] == doc_name) &
            (df_raw_filtered['MES_NUM'] == month_num)
        ]
        max_date = doc_entries['FECHA_CLEAN'].max()
        min_date = doc_entries['FECHA_CLEAN'].min()
        if pd.notna(max_date) and pd.notna(min_date):
            target_sum = 0
            curr = min_date
            while curr <= max_date:
                date_str = curr.strftime('%d/%m/%Y')
                val = daily_targets.get(date_str, 0)
                if doc_name == 'SEBASTIAN GIL GALLEGO' and val == 7:
                    target_sum += 7.33
                else:
                    target_sum += val
                curr += pd.Timedelta(days=1)
            targets.append(int(round(target_sum)))
        else:
            m_target = monthly_targets.get(month_num, 0)
            if doc_name == 'SEBASTIAN GIL GALLEGO':
                targets.append(int(round((m_target / 7.0) * 7.33)))
            else:
                targets.append(int(round(m_target)))
    return targets


def calculate_weekly_target_hours(df_weekly, daily_targets):
    """Calcula las horas a laborar por semana sumando daily_targets de lunes a domingo."""
    targets = []
    for idx, row in df_weekly.iterrows():
        doc_name = row['NOMBRE SUPER VALIDADO']
        inicio = row.get('SEMANA_INICIO')
        fin = row.get('SEMANA_FIN')
        if pd.notna(inicio) and pd.notna(fin):
            target_sum = 0
            curr = pd.Timestamp(inicio)
            end = pd.Timestamp(fin)
            while curr <= end:
                date_str = curr.strftime('%d/%m/%Y')
                val = daily_targets.get(date_str, 0)
                if doc_name == 'SEBASTIAN GIL GALLEGO' and val == 7:
                    target_sum += 7.33
                else:
                    target_sum += val
                curr += pd.Timedelta(days=1)
            targets.append(int(round(target_sum)))
        else:
            targets.append(0)
    return targets


@st.cache_data
def generate_excel_data(df, daily_targets, monthly_targets, cols_to_export_det):
    output = io.BytesIO()

    df_export_dia = dp.get_consolidated_hours_by_date(df)
    df_export_dia['HORAS_A_LABORAR'] = df_export_dia.apply(
        lambda r: 7.33 if (r['NOMBRE SUPER VALIDADO'] == 'SEBASTIAN GIL GALLEGO' and daily_targets.get(r['FECHA_STR'], 0) == 7)
                  else daily_targets.get(r['FECHA_STR'], 0),
        axis=1
    )
    df_export_dia['TOTAL'] = df_export_dia['HORAS_TOTALES'] - df_export_dia['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        df_export_dia[col] = df_export_dia[col].round(0).astype(int)
    df_export_dia_rename = df_export_dia.rename(columns={
        'FECHA_STR': 'Fecha', 'CEDULA_FINAL': 'Cédula',
        'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'HORAS_A_LABORAR': 'Horas a laborar', 'HORAS_TOTALES': 'Horas Laboradas',
        'TOTAL': 'Total', 'CANTIDAD_NOVEDADES': 'Novedades Cubiertas'
    })
    df_export_dia_rename = df_export_dia_rename[
        ['Cédula', 'Médico Supernumerario', 'Fecha', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
    ]
    if not df_export_dia_rename.empty:
        totales = {c: [df_export_dia_rename[c].sum()] if c in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
                   else (['TOTAL GENERAL'] if c == 'Médico Supernumerario' else [''])
                   for c in df_export_dia_rename.columns}
        df_export_dia_rename = pd.concat([df_export_dia_rename, pd.DataFrame(totales)], ignore_index=True)

    df_export_mes = dp.get_consolidated_hours(df)
    df_export_mes['HORAS_A_LABORAR'] = calculate_doctor_target_hours(df_export_mes, df, daily_targets, monthly_targets)
    df_export_mes['TOTAL'] = df_export_mes['HORAS_TOTALES'] - df_export_mes['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        df_export_mes[col] = df_export_mes[col].round(0).astype(int)
    df_export_mes_rename = df_export_mes.rename(columns={
        'CEDULA_FINAL': 'Cédula', 'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'MES': 'Mes', 'HORAS_A_LABORAR': 'Horas a laborar',
        'HORAS_TOTALES': 'Horas Laboradas', 'TOTAL': 'Total',
        'CANTIDAD_NOVEDADES': 'Novedades Cubiertas'
    })
    df_export_mes_rename = df_export_mes_rename[
        ['Cédula', 'Médico Supernumerario', 'Mes', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
    ]
    if not df_export_mes_rename.empty:
        totales = {c: [df_export_mes_rename[c].sum()] if c in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
                   else (['TOTAL GENERAL'] if c == 'Médico Supernumerario' else [''])
                   for c in df_export_mes_rename.columns}
        df_export_mes_rename = pd.concat([df_export_mes_rename, pd.DataFrame(totales)], ignore_index=True)

    df_export_det = df[cols_to_export_det].copy()
    if 'FECHA_CLEAN' in df_export_det.columns:
        df_export_det['Fecha Novedad'] = df_export_det['FECHA_CLEAN'].dt.strftime('%d/%m/%Y')
        df_export_det = df_export_det.drop(columns=['FECHA_CLEAN'])
    if 'REVISION POR CENTRAL DE NOVEDADES' in df_export_det.columns:
        temp_rev = pd.to_datetime(df_export_det['REVISION POR CENTRAL DE NOVEDADES'], errors='coerce')
        df_export_det['REVISION POR CENTRAL DE NOVEDADES'] = temp_rev.dt.strftime('%d/%m/%Y').fillna('')
    df_export_det_rename = df_export_det.rename(columns={
        'REVISION POR CENTRAL DE NOVEDADES': 'Fecha Revisión',
        'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'CEDULA_FINAL': 'Cédula Supernumerario', 'MEDICOS': 'Médico Reemplazado',
        'DOCUMENTO': 'Cédula Reemplazado', 'CIS': 'Sede CIS', 'ZONA': 'Zona',
        'TIPO DE NOVEDAD': 'Novedad', 'HORAS TOTALES DECIMAL': 'Horas',
        'RECARGO NOCTURNO ORDINARIO': 'Recargo Nocturno'
    })

    df_export_semana = dp.get_consolidated_hours_by_week(df)
    df_export_semana['HORAS_A_LABORAR'] = calculate_weekly_target_hours(df_export_semana, daily_targets)
    df_export_semana['TOTAL'] = df_export_semana['HORAS_TOTALES'] - df_export_semana['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        df_export_semana[col] = df_export_semana[col].round(0).astype(int)
    df_export_semana_rename = df_export_semana.rename(columns={
        'CEDULA_FINAL': 'Cédula', 'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'SEMANA': 'Semana', 'HORAS_A_LABORAR': 'Horas a laborar',
        'HORAS_TOTALES': 'Horas Laboradas', 'TOTAL': 'Total',
        'CANTIDAD_NOVEDADES': 'Novedades Cubiertas'
    })
    df_export_semana_rename = df_export_semana_rename[
        ['Cédula', 'Médico Supernumerario', 'Semana', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
    ]
    if not df_export_semana_rename.empty:
        totales = {c: [df_export_semana_rename[c].sum()] if c in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
                   else (['TOTAL GENERAL'] if c == 'Médico Supernumerario' else [''])
                   for c in df_export_semana_rename.columns}
        df_export_semana_rename = pd.concat([df_export_semana_rename, pd.DataFrame(totales)], ignore_index=True)

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export_dia_rename.to_excel(writer, sheet_name='Consolidado por Día', index=False)
        df_export_semana_rename.to_excel(writer, sheet_name='Consolidado por Semana', index=False)
        df_export_mes_rename.to_excel(writer, sheet_name='Consolidado por Mes', index=False)
        df_export_det_rename.to_excel(writer, sheet_name='Detalle Completo', index=False)
        all_sheets = {
            'Consolidado por Día': df_export_dia_rename,
            'Consolidado por Semana': df_export_semana_rename,
            'Consolidado por Mes': df_export_mes_rename,
            'Detalle Completo': df_export_det_rename
        }
        for sheet_name, df_temp in all_sheets.items():
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(df_temp.columns):
                val_lengths = [len(str(v)) for v in df_temp[col].dropna()]
                max_len = max((max(val_lengths) if val_lengths else 0), len(str(col))) + 2
                worksheet.set_column(idx, idx, min(max_len, 40))

    return output.getvalue()


# ── Configuración de página ───────────────────────────────────────────────────

st.set_page_config(
    page_title="Control de horas Central de novedades",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed"
)

custom_css = r"""
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="st-"], [class*="stWidget"], [class*="stSelectbox"], [class*="stMultiSelect"],
    [class*="stMarkdown"], label, p, h1, h2, h3, h4, h5, h6, [data-baseweb="select"] *,
    div[role="listbox"] *, button, .stButton button, [data-testid="stSidebar"] * {
        font-family: 'Inter', sans-serif; color: #202124 !important;
    }
    .block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
    [data-testid="stHeader"] { display: none !important; }
    .kpi-container { display: flex; flex-wrap: wrap; gap: 12px; margin: 15px 0px 25px 0px; }
    .kpi-card { padding: 10px 16px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        flex: 1; min-width: 180px; border-top: 1px solid rgba(0,0,0,0.05);
        border-bottom: 1px solid rgba(0,0,0,0.05); border-right: 1px solid rgba(0,0,0,0.05);
        background-color: #ffffff; }
    .kpi-blue { border-left: 6px solid #1e5cc8; background-color: #f0f4fc; }
    .kpi-green { border-left: 6px solid #1e8e3e; background-color: #f4faf6; }
    .kpi-yellow { border-left: 6px solid #f9ab00; background-color: #fefcf3; }
    .kpi-red { border-left: 6px solid #d93025; background-color: #fdf5f4; }
    .kpi-title { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px; }
    .kpi-blue .kpi-title { color: #1e5cc8 !important; }
    .kpi-green .kpi-title { color: #1e8e3e !important; }
    .kpi-yellow .kpi-title { color: #b06000 !important; }
    .kpi-red .kpi-title { color: #d93025 !important; }
    .kpi-value { font-size: 22px; font-weight: 700; color: #202124 !important; }
    div.stButton > button { border-radius: 4px; font-weight: 500; transition: all 0.2s ease; }
    .element-container:has(.search-btn), .element-container:has(.clear-btn),
    .element-container:has(.export-btn) { display: none !important; }
    .element-container:has(.clear-btn) + .element-container button,
    .element-container:has(.export-btn) + .element-container button,
    .element-container:has(.search-btn) + .element-container button {
        display: inline-flex !important; align-items: center !important; justify-content: center !important; }
    .element-container:has(.clear-btn) + .element-container button *,
    .element-container:has(.export-btn) + .element-container button *,
    .element-container:has(.search-btn) + .element-container button * { display: none !important; }
    .element-container:has(.clear-btn) + .element-container button::after {
        font-family: "bootstrap-icons" !important; content: "\F5DE" !important;
        font-size: 18px !important; visibility: visible !important; color: inherit !important; }
    .element-container:has(.export-btn) + .element-container button::after {
        font-family: "bootstrap-icons" !important; content: "\F368" !important;
        font-size: 18px !important; visibility: visible !important; color: inherit !important; }
    .element-container:has(.search-btn) + .element-container button::after {
        font-family: "bootstrap-icons" !important; content: "\F52A" !important;
        font-size: 18px !important; visibility: visible !important; color: inherit !important; }
    .element-container:has(.clear-btn) + .element-container button {
        background-color: #ffffff !important; color: #d93025 !important;
        border: 1.5px solid #d93025 !important; width: 100%; height: 40px;
        margin-top: 24px !important; font-weight: 600; transition: all 0.2s ease; }
    .element-container:has(.clear-btn) + .element-container button:hover {
        background-color: #d93025 !important; color: #ffffff !important; }
    .element-container:has(.export-btn) + .element-container button {
        background-color: #ffffff !important; color: #1e8e3e !important;
        border: 1.5px solid #1e8e3e !important; width: 100%; height: 40px;
        margin-top: 24px !important; font-weight: 600; transition: all 0.2s ease; }
    .element-container:has(.export-btn) + .element-container button:hover {
        background-color: #1e8e3e !important; color: #ffffff !important; }
    .element-container:has(.search-btn) + .element-container button {
        background-color: #ffffff !important; color: #1a73e8 !important;
        border: 1.5px solid #1a73e8 !important; width: 100%; height: 40px;
        margin-top: 24px !important; font-weight: 600; transition: all 0.2s ease; }
    .element-container:has(.search-btn) + .element-container button:hover {
        background-color: #1a73e8 !important; color: #ffffff !important; }
    .totals-inline-bar { display: flex !important; justify-content: center !important;
        align-items: center !important; flex-wrap: wrap !important;
        background-color: #ffffff !important; padding: 8px 12px !important;
        border-radius: 6px !important; border: 1px solid #e2e8f0 !important;
        margin-top: 10px !important; margin-bottom: 20px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important; gap: 8px 24px !important; width: 100% !important; }
    .totals-inline-item { display: flex !important; align-items: center !important; flex-shrink: 0 !important; }
    .totals-inline-label { font-size: 13px !important; font-weight: 700 !important;
        color: #202124 !important; white-space: nowrap !important; flex-shrink: 0 !important; }
    .totals-inline-badge { padding: 3px 8px !important; border-radius: 4px !important;
        color: #ffffff !important; font-weight: 700 !important; font-size: 13px !important;
        margin-left: 6px !important; white-space: nowrap !important; flex-shrink: 0 !important; }
    .badge-blue { background-color: #1a73e8 !important; }
    .badge-green { background-color: #137333 !important; }
    .badge-yellow { background-color: #f9ab00 !important; }
    .badge-red { background-color: #d93025 !important; }
    div[data-testid="stSelectbox"] > div > div, div[data-testid="stMultiSelect"] > div > div {
        border: 1.5px solid #cccccc !important; background-color: #ffffff !important;
        border-radius: 6px !important; transition: border-color 0.2s ease-in-out !important; }
    div[data-testid="stSelectbox"] > div > div:hover, div[data-testid="stMultiSelect"] > div > div:hover {
        border-color: #1a73e8 !important; }
    .filter-panel-marker { display: none !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) {
        border: 1.5px solid #cccccc !important; border-radius: 8px !important;
        background-color: #ffffff !important; padding: 16px 20px !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05) !important; margin-bottom: 20px !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important; gap: 12px !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) [data-testid="column"] { min-width: 0px !important; }
    .table-scroll-container { max-height: 420px; overflow-y: auto; overflow-x: auto;
        border: 1.5px solid #cccccc; border-radius: 8px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-top: 10px; margin-bottom: 15px; }
    .custom-table { width: 100%; border-collapse: collapse; font-size: 13.5px;
        color: #202124; background-color: #ffffff; }
    .custom-table th { background-color: #0f5ba6 !important; color: #ffffff !important;
        font-weight: 600; padding: 12px 14px; border: 1px solid #cbd5e1;
        position: sticky; top: 0; z-index: 10; }
    .custom-table th, .custom-table th * { color: #ffffff !important; }
    .custom-table td { padding: 10px 14px; border: 1px solid #cbd5e1;
        vertical-align: middle; color: #202124 !important; }
    .custom-table tr:nth-child(even) { background-color: #f8fafc; }
    .custom-table tr:hover td { background-color: #f1f5f9; }
    .custom-table th, .custom-table td { text-align: center !important; }
    button[data-testid="stPopoverButton"] { background-color: #ffffff !important;
        border: 1.5px solid #cccccc !important; color: #202124 !important;
        border-radius: 6px !important; height: 40px !important; width: 40px !important;
        display: inline-flex !important; align-items: center !important; justify-content: center !important;
        padding: 0 !important; transition: all 0.2s ease !important; font-size: 0 !important; }
    button[data-testid="stPopoverButton"]:hover { background-color: #f8f9fa !important;
        border-color: #1a73e8 !important; }
    button[data-testid="stPopoverButton"]::before { font-family: "bootstrap-icons" !important;
        content: "\F3E5" !important; font-size: 20px !important; visibility: visible !important;
        color: inherit !important; display: inline-block !important; transition: transform 0.4s ease !important; }
    button[data-testid="stPopoverButton"]:hover::before { transform: rotate(90deg) !important; }
    button[data-testid="stPopoverButton"] * { display: none !important; font-size: 0 !important;
        width: 0 !important; height: 0 !important; overflow: hidden !important; visibility: hidden !important; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────

IS_LOCAL = platform.system() == "Windows"
current_month_name = dp.MESES_MAP.get(get_local_now().month, "Enero")

defaults = {
    'mes_sel': [current_month_name],
    'cedula_sel': "Todas",
    'nombre_sel': "Todos",
    'agrupacion_sel': "Por Mes",
    'mes_sel_draft': [current_month_name],
    'cedula_sel_draft': "Todas",
    'nombre_sel_draft': "Todos",
    'agrupacion_sel_draft': "Por Mes",
    'file_path_input': r"C:\Users\JuanJoseOsorioMolina\OneDrive - U.T SAN VICENTE CES\CONSOLIDADO 2026.xlsx",
    'uploaded_file_name': None,
    'df_raw': None,
    'load_error': None,
    'last_refresh': None,
    'monthly_targets': {},
    'daily_targets': {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def reset_filters():
    st.session_state.mes_sel = []
    st.session_state.cedula_sel = "Todas"
    st.session_state.nombre_sel = "Todos"
    st.session_state.agrupacion_sel = "Por Mes"
    st.session_state.mes_sel_draft = []
    st.session_state.cedula_sel_draft = "Todas"
    st.session_state.nombre_sel_draft = "Todos"
    st.session_state.agrupacion_sel_draft = "Por Mes"
    for k in list(st.session_state.keys()):
        if k.startswith("mes_sel_draft_widget_"):
            del st.session_state[k]
    for k in ['cedula_sel_draft_widget', 'nombre_sel_draft_widget', 'agrupacion_sel_draft_widget']:
        if k in st.session_state:
            st.session_state[k] = defaults.get(k.replace('_widget', ''), "Todas")


# ── Cabecera ──────────────────────────────────────────────────────────────────

col_config, col_title, col_logo = st.columns([0.6, 6.4, 3.0], vertical_alignment="center")

with col_config:
    with st.popover("", help="Configuración de Origen de Datos"):
        st.markdown("<h3 style='margin:0 0 10px 0; font-family:Inter,sans-serif; color:#0b3c5d;'>⚙️ Configuración de Datos</h3>", unsafe_allow_html=True)

        # Mostrar última actualización si existe
        if st.session_state.last_refresh:
            st.info(f"🕐 Última carga: **{st.session_state.last_refresh}**")

        # Botón para recargar desde OneDrive (siempre visible)
        if get_onedrive_config():
            if st.button("🔄 Recargar desde OneDrive", use_container_width=True):
                st.session_state.df_raw = None
                st.session_state.load_error = None
                cargar_desde_onedrive()
                st.rerun()
        else:
            # Solo en local: mostrar opciones manuales
            st.warning("⚠️ Sin credenciales de OneDrive. Modo local activo.")
            file_path = st.text_input(
                "Ruta del archivo local (.xlsx):",
                key="file_path_input"
            )
            if os.path.exists(file_path):
                mtime = os.path.getmtime(file_path)
                last_updated = datetime.fromtimestamp(mtime).strftime('%d/%m/%Y %I:%M:%S %p')
                st.success(f"Archivo encontrado — Modificado: **{last_updated}**")
            else:
                st.error("❌ Archivo no encontrado.")

            if st.button("🔄 Cargar desde ruta local", use_container_width=True):
                if os.path.exists(file_path):
                    try:
                        st.session_state.df_raw = dp.load_and_clean_data(file_path)
                        m, d = dp.load_calendar_targets(file_path)
                        st.session_state.monthly_targets = m
                        st.session_state.daily_targets = d
                        st.session_state.load_error = None
                        st.session_state.last_refresh = get_local_now().strftime('%d/%m/%Y %H:%M:%S')
                        st.rerun()
                    except Exception as e:
                        st.session_state.load_error = str(e)

            st.markdown("---")
            uploaded_file = st.file_uploader("O sube el archivo manualmente:", type=["xlsx", "xls"])
            if uploaded_file is not None:
                try:
                    file_bytes = io.BytesIO(uploaded_file.read())
                    st.session_state.df_raw = dp.load_and_clean_data(file_bytes)
                    file_bytes.seek(0)
                    m, d = dp.load_calendar_targets(file_bytes)
                    st.session_state.monthly_targets = m
                    st.session_state.daily_targets = d
                    st.session_state.load_error = None
                    st.session_state.uploaded_file_name = uploaded_file.name
                    st.session_state.last_refresh = get_local_now().strftime('%d/%m/%Y %H:%M:%S')
                    st.rerun()
                except Exception as e:
                    st.session_state.load_error = str(e)

with col_title:
    st.markdown(
        "<h1 style='font-family:\"Segoe UI\",sans-serif; font-weight:400; color:#0b3c5d; "
        "font-size:38px; margin:0;'>Control de horas Central de novedades</h1>",
        unsafe_allow_html=True
    )
with col_logo:
    logo_b64 = get_base64_image("logo.png")
    if logo_b64:
        st.markdown(
            f'<div style="text-align:right;"><img src="{logo_b64}" style="max-height:55px;"></div>',
            unsafe_allow_html=True
        )

# ── Carga automática al arrancar ──────────────────────────────────────────────

if st.session_state.df_raw is None and st.session_state.load_error is None:
    config = get_onedrive_config()
    if config:
        # En Streamlit Cloud: carga automática desde OneDrive
        cargar_desde_onedrive()
    elif IS_LOCAL:
        # En local Windows: carga desde ruta sincronizada de OneDrive
        file_path = st.session_state.file_path_input
        if os.path.exists(file_path):
            try:
                st.session_state.df_raw = dp.load_and_clean_data(file_path)
                m, d = dp.load_calendar_targets(file_path)
                st.session_state.monthly_targets = m
                st.session_state.daily_targets = d
                st.session_state.last_refresh = get_local_now().strftime('%d/%m/%Y %H:%M:%S')
            except Exception as e:
                st.session_state.load_error = str(e)
        else:
            st.session_state.load_error = (
                f"Archivo no encontrado: {file_path}\n"
                "Usa el botón ⚙️ para cambiar la ruta o subir manualmente."
            )
    else:
        st.session_state.load_error = (
            "No se encontraron credenciales de OneDrive configuradas. "
            "Configura los Secrets en Streamlit Cloud."
        )

# ── Manejo de errores de carga ────────────────────────────────────────────────

if st.session_state.load_error:
    msg = st.session_state.load_error
    if "Permission denied" in msg or "Errno 13" in msg:
        st.error("⚠️ **Archivo bloqueado:** Cierra el Excel o espera que OneDrive termine de sincronizar, luego usa ⚙️ → Recargar.")
    else:
        st.error(f"⚠️ Error de carga: {msg}")
    st.info("💡 Usa el botón ⚙️ arriba a la izquierda para recargar o cambiar el origen de datos.")
    st.stop()

if st.session_state.df_raw is None:
    st.info("⏳ Cargando datos...")
    st.stop()

df_raw = st.session_state.df_raw

# ── Filtros ───────────────────────────────────────────────────────────────────

df_filtrado = df_raw.copy()
if st.session_state.mes_sel:
    df_filtrado = df_filtrado[df_filtrado['MES'].isin(st.session_state.mes_sel)]
if st.session_state.cedula_sel != "Todas":
    df_filtrado = df_filtrado[df_filtrado['CEDULA_FINAL'] == st.session_state.cedula_sel]
if st.session_state.nombre_sel != "Todos":
    df_filtrado = df_filtrado[df_filtrado['NOMBRE SUPER VALIDADO'] == st.session_state.nombre_sel]

detalle_cols_base = [
    'REVISION POR CENTRAL DE NOVEDADES', 'FECHA_CLEAN', 'NOMBRE SUPER VALIDADO', 'CEDULA_FINAL',
    'MEDICOS', 'DOCUMENTO', 'CIS', 'ZONA', 'TIPO DE NOVEDAD', 'HORAS TOTALES DECIMAL', 'RECARGO NOCTURNO ORDINARIO'
]

meses_disponibles = sorted(
    df_raw['MES'].dropna().unique().tolist(),
    key=lambda m: list(dp.MESES_MAP.values()).index(m) if m in dp.MESES_MAP.values() else 99
)

# ── Panel de filtros ──────────────────────────────────────────────────────────

with st.container(border=True):
    st.markdown('<div class="filter-panel-marker"></div>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6, c7 = st.columns([1.5, 2.0, 1.8, 2.8, 1.0, 1.0, 1.0])

    with c1:
        agrupacion_options = ["Por Día", "Por Semana", "Por Mes"]
        agrupacion_idx = agrupacion_options.index(st.session_state.agrupacion_sel_draft) if st.session_state.agrupacion_sel_draft in agrupacion_options else 0
        agrupacion_sel_draft = st.selectbox("Agrupar por:", options=agrupacion_options, index=agrupacion_idx, key="agrupacion_sel_draft_widget")
        st.session_state.agrupacion_sel_draft = agrupacion_sel_draft

    with c2:
        active_keys = [k for k in st.session_state.keys() if k.startswith("mes_sel_draft_widget_")]
        num_items = 0
        if active_keys:
            val = st.session_state[active_keys[0]]
            if isinstance(val, list):
                num_items = len(val)
                st.session_state.mes_sel_draft = val
        mes_key = f"mes_sel_draft_widget_{num_items}"
        default_meses = [m for m in st.session_state.mes_sel_draft if m in meses_disponibles]
        meses_sel_draft = st.multiselect("Mes:", options=meses_disponibles, default=default_meses, key=mes_key, placeholder="seleccione")
        st.session_state.mes_sel_draft = meses_sel_draft

    df_para_filtros = df_raw.copy()
    if st.session_state.mes_sel_draft:
        df_para_filtros = df_para_filtros[df_para_filtros['MES'].isin(st.session_state.mes_sel_draft)]
    cedulas_disponibles = ["Todas"] + sorted(df_para_filtros['CEDULA_FINAL'].dropna().unique().tolist())
    nombres_disponibles = ["Todos"] + sorted(df_para_filtros['NOMBRE SUPER VALIDADO'].dropna().unique().tolist())
    if st.session_state.cedula_sel_draft not in cedulas_disponibles:
        st.session_state.cedula_sel_draft = "Todas"
    if st.session_state.nombre_sel_draft not in nombres_disponibles:
        st.session_state.nombre_sel_draft = "Todos"

    with c3:
        cedula_idx = cedulas_disponibles.index(st.session_state.cedula_sel_draft) if st.session_state.cedula_sel_draft in cedulas_disponibles else 0
        cedula_sel_draft = st.selectbox("Cedula:", options=cedulas_disponibles, index=cedula_idx, key="cedula_sel_draft_widget")
        st.session_state.cedula_sel_draft = cedula_sel_draft

    with c4:
        nombre_idx = nombres_disponibles.index(st.session_state.nombre_sel_draft) if st.session_state.nombre_sel_draft in nombres_disponibles else 0
        nombre_sel_draft = st.selectbox("Supernumerario:", options=nombres_disponibles, index=nombre_idx, key="nombre_sel_draft_widget")
        st.session_state.nombre_sel_draft = nombre_sel_draft

    with c5:
        st.markdown('<div class="search-btn">', unsafe_allow_html=True)
        if st.button("Buscar", key="btn_search", use_container_width=True):
            st.session_state.mes_sel = st.session_state.mes_sel_draft
            st.session_state.cedula_sel = st.session_state.cedula_sel_draft
            st.session_state.nombre_sel = st.session_state.nombre_sel_draft
            st.session_state.agrupacion_sel = st.session_state.agrupacion_sel_draft
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with c6:
        st.markdown('<div class="clear-btn">', unsafe_allow_html=True)
        st.button("Borrar Filtros", key="btn_clear", use_container_width=True, on_click=reset_filters)
        st.markdown('</div>', unsafe_allow_html=True)

    with c7:
        st.markdown('<div class="export-btn">', unsafe_allow_html=True)
        cols_to_export_det = [c for c in detalle_cols_base if c in df_filtrado.columns]
        daily_targets = st.session_state.get('daily_targets', {})
        monthly_targets = st.session_state.get('monthly_targets', {})
        excel_data = generate_excel_data(df_filtrado, daily_targets, monthly_targets, cols_to_export_det)
        st.download_button(
            label="Exportar Excel",
            data=excel_data,
            file_name=f"CONSOLIDADO_HORAS_SUPERNUMERARIOS_{get_local_now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_export",
            use_container_width=True,
            on_click=save_to_downloads,
            args=(excel_data,)
        )
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<hr style='margin-top:-15px; margin-bottom:10px; border:0; border-top:1px solid #e0e0e0;'/>", unsafe_allow_html=True)

# ── Tabla consolidada ─────────────────────────────────────────────────────────

tabla_consolidada = dp.get_consolidated_hours(df_filtrado)
agrupacion_vista = st.session_state.agrupacion_sel

if agrupacion_vista == "Por Día":
    tabla_consolidada_vista = dp.get_consolidated_hours_by_date(df_filtrado)
    daily_targets = st.session_state.get('daily_targets', {})
    tabla_consolidada_vista['HORAS_A_LABORAR'] = tabla_consolidada_vista.apply(
        lambda r: 7.33 if (r['NOMBRE SUPER VALIDADO'] == 'SEBASTIAN GIL GALLEGO' and daily_targets.get(r['FECHA_STR'], 0) == 7)
                  else daily_targets.get(r['FECHA_STR'], 0),
        axis=1
    )
    tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        tabla_consolidada_vista[col] = tabla_consolidada_vista[col].round(0).astype(int)
elif agrupacion_vista == "Por Semana":
    tabla_consolidada_vista = dp.get_consolidated_hours_by_week(df_filtrado)
    daily_targets = st.session_state.get('daily_targets', {})
    tabla_consolidada_vista['HORAS_A_LABORAR'] = calculate_weekly_target_hours(
        tabla_consolidada_vista, daily_targets
    )
    tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        tabla_consolidada_vista[col] = tabla_consolidada_vista[col].round(0).astype(int)
else:
    tabla_consolidada_vista = tabla_consolidada.copy()
    monthly_targets = st.session_state.get('monthly_targets', {})
    daily_targets = st.session_state.get('daily_targets', {})
    tabla_consolidada_vista['HORAS_A_LABORAR'] = calculate_doctor_target_hours(
        tabla_consolidada_vista, df_filtrado, daily_targets, monthly_targets
    )
    tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        tabla_consolidada_vista[col] = tabla_consolidada_vista[col].round(0).astype(int)

if not tabla_consolidada_vista.empty:
    tot_horas_a_laborar = tabla_consolidada_vista['HORAS_A_LABORAR'].sum()
    tot_horas_laboradas = tabla_consolidada_vista['HORAS_TOTALES'].sum()
    tot_diferencia = tabla_consolidada_vista['TOTAL'].sum()
    tot_novedades = tabla_consolidada_vista['CANTIDAD_NOVEDADES'].sum()
    tot_medicos_activos = df_filtrado['NOMBRE SUPER VALIDADO'].nunique()
    dif_badge_class = "badge-green" if tot_diferencia > 0 else ("badge-red" if tot_diferencia < 0 else "badge-yellow")
    dif_sign = "+" if tot_diferencia > 0 else ""
    st.markdown(
        f'<div class="totals-inline-bar">'
        f'<div class="totals-inline-item"><span class="totals-inline-label">MÉDICOS ACTIVOS:</span><span class="totals-inline-badge badge-blue">{tot_medicos_activos:,}</span></div>'
        f'<div class="totals-inline-item"><span class="totals-inline-label">HORAS A LABORAR:</span><span class="totals-inline-badge badge-blue">{tot_horas_a_laborar:,.0f} hrs</span></div>'
        f'<div class="totals-inline-item"><span class="totals-inline-label">HORAS LABORADAS:</span><span class="totals-inline-badge badge-green">{tot_horas_laboradas:,.0f} hrs</span></div>'
        f'<div class="totals-inline-item"><span class="totals-inline-label">DIFERENCIA:</span><span class="totals-inline-badge {dif_badge_class}">{dif_sign}{tot_diferencia:,.0f} hrs</span></div>'
        f'<div class="totals-inline-item"><span class="totals-inline-label">NOVEDADES CUBIERTAS:</span><span class="totals-inline-badge badge-red">{tot_novedades:,}</span></div>'
        f'</div>',
        unsafe_allow_html=True
    )

st.caption(f"Mostrando {len(tabla_consolidada_vista)} filas en el resumen consolidado.")

if agrupacion_vista == "Por Día":
    tabla_display = tabla_consolidada_vista.rename(columns={
        'FECHA_STR': 'Fecha', 'CEDULA_FINAL': 'Cédula',
        'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'HORAS_A_LABORAR': 'Horas a laborar', 'HORAS_TOTALES': 'Horas Laboradas',
        'TOTAL': 'Total', 'CANTIDAD_NOVEDADES': 'Novedades Cubiertas'
    })
    cols_show = ['Cédula', 'Médico Supernumerario', 'Fecha', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
elif agrupacion_vista == "Por Semana":
    tabla_display = tabla_consolidada_vista.rename(columns={
        'CEDULA_FINAL': 'Cédula', 'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'SEMANA': 'Semana', 'HORAS_A_LABORAR': 'Horas a laborar',
        'HORAS_TOTALES': 'Horas Laboradas', 'TOTAL': 'Total',
        'CANTIDAD_NOVEDADES': 'Novedades Cubiertas'
    })
    cols_show = ['Cédula', 'Médico Supernumerario', 'Semana', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']
else:
    tabla_display = tabla_consolidada_vista.rename(columns={
        'CEDULA_FINAL': 'Cédula', 'NOMBRE SUPER VALIDADO': 'Médico Supernumerario',
        'MES': 'Mes', 'HORAS_A_LABORAR': 'Horas a laborar',
        'HORAS_TOTALES': 'Horas Laboradas', 'TOTAL': 'Total',
        'CANTIDAD_NOVEDADES': 'Novedades Cubiertas'
    })
    cols_show = ['Cédula', 'Médico Supernumerario', 'Mes', 'Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']

tabla_display_formatted = tabla_display[cols_show].copy()
for col in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Novedades Cubiertas']:
    if col in tabla_display_formatted.columns:
        tabla_display_formatted[col] = pd.to_numeric(tabla_display_formatted[col], errors='coerce').fillna(0).round(0).astype(int)

html_table = tabla_display_formatted.to_html(index=False, classes='custom-table', escape=False)

th_index = 0
def add_onclick_to_th(match):
    global th_index
    content = match.group(1)
    res = f'<th onclick="sortTable({th_index})" style="cursor:pointer;user-select:none;" title="Haz clic para ordenar">{content}</th>'
    th_index += 1
    return res

th_index = 0
html_table = re.sub(r'<th>(.*?)</th>', add_onclick_to_th, html_table)

iframe_template = r"""<!DOCTYPE html><html><head>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
body{margin:0;padding:0;font-family:'Inter',sans-serif;background-color:transparent;}
.table-scroll-container{max-height:420px;overflow-y:auto;overflow-x:auto;border:1.5px solid #cccccc;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,0.05);}
.custom-table{width:100%;border-collapse:collapse;font-size:13.5px;color:#202124;background-color:#ffffff;}
.custom-table th{background-color:#0f5ba6!important;color:#ffffff!important;font-weight:600;padding:12px 14px;border:1px solid #cbd5e1;position:sticky;top:0;z-index:10;cursor:pointer;user-select:none;text-align:center!important;}
.custom-table td{padding:10px 14px;border:1px solid #cbd5e1;vertical-align:middle;color:#202124!important;text-align:center!important;}
.custom-table tr:nth-child(even){background-color:#f8fafc;}
.custom-table tr:hover td{background-color:#f1f5f9;}
</style>
<script>
function sortTable(columnIndex){
    const table=document.querySelector(".custom-table");if(!table)return;
    const tbody=table.querySelector("tbody")||table;
    const rows=Array.from(tbody.querySelectorAll("tr"));
    let dir=table.getAttribute("data-sort-dir")==="asc"?"desc":"asc";
    let lastCol=parseInt(table.getAttribute("data-sort-col"));
    if(lastCol!==columnIndex){dir="asc";}
    table.setAttribute("data-sort-dir",dir);table.setAttribute("data-sort-col",columnIndex);
    rows.sort((a,b)=>{
        let vA=a.cells[columnIndex].innerText.trim();let vB=b.cells[columnIndex].innerText.trim();
        let cA=vA.replace(/\./g,"").replace(/,/g,"").replace(/%/g,"").replace(/ hrs/g,"").trim();
        let cB=vB.replace(/\./g,"").replace(/,/g,"").replace(/%/g,"").replace(/ hrs/g,"").trim();
        let nA=parseFloat(cA);let nB=parseFloat(cB);
        if(!isNaN(nA)&&!isNaN(nB)){return dir==="asc"?nA-nB:nB-nA;}
        return dir==="asc"?vA.localeCompare(vB):vB.localeCompare(vA);
    });
    rows.forEach(row=>tbody.appendChild(row));
    const headers=table.querySelectorAll("th");
    headers.forEach((th,idx)=>{
        th.innerHTML=th.innerHTML.replace(/ <span style="font-size: 11px;">(▲|▼)<\/span>/g,"");
        if(idx===columnIndex){th.innerHTML+=` <span style="font-size: 11px;">${dir==='asc'?'▲':'▼'}</span>`;}
    });
}
</script>
</head><body>
<div class="table-scroll-container">__TABLE_HTML__</div>
</body></html>"""

iframe_content = iframe_template.replace("__TABLE_HTML__", html_table)
st.components.v1.html(iframe_content, height=450, scrolling=False)

st.markdown("<div style='margin-top:25px;font-size:13px;color:#888888;'>© 2026 - San Vicente CES</div>", unsafe_allow_html=True)