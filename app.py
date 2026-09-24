# -*- coding: utf-8 -*-
"""
Interfaz de Usuario en Streamlit para el Dashboard de Auditoría de Médicos Supernumerarios.
"""

import streamlit as st
import pandas as pd
import os
import io
import platform
import requests
from datetime import datetime, timedelta, timezone

def get_local_now():
    return datetime.now(timezone(timedelta(hours=-5)))

import data_processor as dp
import base64
import re

# ── Helpers ───────────────────────────────────────────────────────────────────

def get_base64_image(image_path):
    if os.path.exists(image_path):
        with open(image_path, "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        return f"data:image/png;base64,{encoded_string}"
    return ""


def get_onedrive_config():
    """
    Lee las credenciales de SharePoint desde st.secrets.
    Retorna None si no están configuradas → la app cae al modo local/manual.
    """
    try:
        config = {
            "tenant_id": st.secrets["AZURE_TENANT_ID"],
            "client_id": st.secrets["AZURE_CLIENT_ID"],
            "client_secret": st.secrets["AZURE_CLIENT_SECRET"],
            "mode": "sharepoint",
            "sharepoint_host": st.secrets["SHAREPOINT_HOST"],
            "site_path": st.secrets["SHAREPOINT_SITE_PATH"],
            "file_path": st.secrets["SHAREPOINT_FILE_PATH"],
        }
        if "SHAREPOINT_FILE_PATH_HISTORIC" in st.secrets:
            config["file_path_historic"] = st.secrets["SHAREPOINT_FILE_PATH_HISTORIC"]
        else:
            config["file_path_historic"] = config["file_path"].replace("CONSOLIDADO 2026.xlsx", "CONSOLIDADO 2026 HISTORICO.xlsx")
        return config
    except Exception:
        return None


@st.cache_data(ttl=900, show_spinner=False)
def fetch_sharepoint_dataset(tenant_id, client_id, client_secret, sharepoint_host, site_path, file_path, file_path_historic):
    """
    Descarga y procesa los archivos de SharePoint de forma optimizada con caché en Streamlit (15 min).
    Reutiliza la sesión HTTP y procesa los datos con el motor calamine.
    """
    session = requests.Session()
    file_bytes_actual = dp.download_excel_from_sharepoint(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        sharepoint_host=sharepoint_host,
        site_path=site_path,
        file_drive_path=file_path,
        session=session,
    )
    xl_actual = pd.ExcelFile(file_bytes_actual, engine='calamine')
    df_raw_actual = dp.load_and_clean_data(xl_actual, preferred_sheet='CONSOLIDADO 2026 NOMINA')

    df_raw_historico = None
    xl_hist = None
    hist_loaded = False
    if file_path_historic:
        try:
            file_bytes_hist = dp.download_excel_from_sharepoint(
                tenant_id=tenant_id,
                client_id=client_id,
                client_secret=client_secret,
                sharepoint_host=sharepoint_host,
                site_path=site_path,
                file_drive_path=file_path_historic,
                session=session,
            )
            xl_hist = pd.ExcelFile(file_bytes_hist, engine='calamine')
            df_raw_historico = dp.load_and_clean_data(xl_hist, preferred_sheet='CONSOLIDADO 2026 NOMINA HISTORI')
        except Exception:
            pass

    if df_raw_historico is not None and not df_raw_historico.empty:
        meses_en_actual = set(df_raw_actual['MES_NUM'].dropna().unique())
        df_raw_historico = df_raw_historico[~df_raw_historico['MES_NUM'].isin(meses_en_actual)]
        df_raw = pd.concat([df_raw_historico, df_raw_actual], ignore_index=True)
        hist_loaded = True
    else:
        df_raw = df_raw_actual

    df_super_actual = dp.load_supernumerario_sheets(xl_actual)
    df_super_hist = dp.load_supernumerario_sheets(xl_hist) if (df_raw_historico is not None and xl_hist is not None) else None
    if df_super_hist is not None and not df_super_hist.empty:
        if df_super_actual is not None and not df_super_actual.empty:
            df_super = pd.concat([df_super_hist, df_super_actual], ignore_index=True).drop_duplicates()
        else:
            df_super = df_super_hist
    else:
        df_super = df_super_actual

    plaza_actual = dp.load_plaza_fija_dates(xl_actual)
    plaza_hist = dp.load_plaza_fija_dates(xl_hist) if (df_raw_historico is not None and xl_hist is not None) else {}
    if plaza_hist:
        plaza_hist.update(plaza_actual)
        plaza_fija_dates = plaza_hist
    else:
        plaza_fija_dates = plaza_actual

    m_targets, d_targets = dp.load_calendar_targets(xl_actual)

    return df_raw, df_super, plaza_fija_dates, m_targets, d_targets, hist_loaded


def cargar_desde_onedrive(force_refresh=False):
    """Carga los datos desde SharePoint con caché optimizado y los guarda en session_state."""
    config = get_onedrive_config()
    if not config:
        st.session_state.load_error = (
            "No se encontraron credenciales de SharePoint. "
            "Configura los Secrets en Streamlit Cloud o usa el modo manual."
        )
        return

    if force_refresh:
        fetch_sharepoint_dataset.clear()

    try:
        with st.spinner("Conectando con SharePoint..."):
            (
                df_raw,
                df_super,
                plaza_fija_dates,
                m_targets,
                d_targets,
                hist_loaded
            ) = fetch_sharepoint_dataset(
                config["tenant_id"],
                config["client_id"],
                config["client_secret"],
                config["sharepoint_host"],
                config["site_path"],
                config["file_path"],
                config.get("file_path_historic", "")
            )

            st.session_state.df_raw = df_raw
            st.session_state.df_super = df_super
            st.session_state.plaza_fija_dates = plaza_fija_dates
            st.session_state.monthly_targets = m_targets
            st.session_state.daily_targets = d_targets
            st.session_state.hist_loaded = hist_loaded
            st.session_state.load_error = None
            st.session_state.last_refresh = get_local_now().strftime('%d/%m/%Y %H:%M:%S')
    except Exception as e:
        st.session_state.df_raw = None
        st.session_state.df_super = None
        st.session_state.load_error = str(e)



def calculate_doctor_target_hours(df_grouped, df_raw_filtered, daily_targets, monthly_targets, df_super=None):
    if 'HORAS_A_LABORAR' in df_grouped.columns:
        return df_grouped['HORAS_A_LABORAR'].tolist()
    targets = []
    restr_dict = dp.get_restriction_dict(df_ref=df_raw_filtered, df_super=df_super)

    for idx, row in df_grouped.iterrows():
        doc_name = row['NOMBRE SUPER VALIDADO']
        month_num = row['MES_NUM']
        if pd.isna(month_num):
            targets.append(0)
            continue
        month_num = int(month_num)
        doc_entries = df_raw_filtered[
            (df_raw_filtered['NOMBRE SUPER VALIDADO'] == doc_name) &
            (df_raw_filtered['MES_NUM'] == month_num)
        ]

        # Intentar determinar el primer y último turno basándonos en la hoja de supernumerarios
        found_in_super = False
        if df_super is not None and not df_super.empty:
            doc_norm = dp.normalize_name(doc_name)
            super_entries = df_super[
                (df_super['NOMBRE_NORM'] == doc_norm) &
                (df_super['FECHA_CLEAN'].dt.month == month_num)
            ]
            if not super_entries.empty:
                max_date = super_entries['FECHA_CLEAN'].max()
                min_date = super_entries['FECHA_CLEAN'].min()
                found_in_super = True

        if not found_in_super:
            max_date = doc_entries['FECHA_CLEAN'].max()
            min_date = doc_entries['FECHA_CLEAN'].min()

        if pd.notna(max_date) and pd.notna(min_date):
            doc_norm = dp.normalize_name(doc_name)
            # Verificar si tiene días de licencia/permiso en el mes
            has_leave_days = any(
                dp.is_unpaid_leave_or_permission(v)
                for (d_norm, f_date), v in restr_dict.items()
                if d_norm == doc_norm and f_date.month == month_num
            )

            # Si el médico estuvo activo todo el mes Y NO tiene licencias no remuneradas/permisos,
            # se le asigna la meta mensual completa.
            if min_date.day <= 7 and max_date.day >= (max_date.days_in_month - 6) and not has_leave_days:
                m_target = monthly_targets.get(month_num, 0)
                if doc_name == 'SEBASTIAN GIL GALLEGO':
                    targets.append(int(round((m_target / 7.0) * 7.33)))
                else:
                    targets.append(int(round(m_target)))
                continue

            # Suma de metas día a día excluyendo licencias/permisos
            min_date_aligned = min_date
            max_date_aligned = max_date

            target_sum = 0
            curr = min_date_aligned
            today_date = get_local_now().date()
            while curr <= max_date_aligned:
                if curr.month == month_num:
                    if curr.date() > today_date:
                        day_entries = doc_entries[doc_entries['FECHA_CLEAN'].dt.date == curr.date()]
                        if day_entries.empty:
                            curr += pd.Timedelta(days=1)
                            continue

                    date_str = curr.strftime('%d/%m/%Y')
                    val = daily_targets.get(date_str, 0)

                    restr_text = restr_dict.get((doc_norm, curr.date()))
                    if restr_text and dp.is_unpaid_leave_or_permission(restr_text):
                        val = 0.0

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


def calculate_weekly_target_hours(df_weekly, daily_targets, df_raw=None):
    """Calcula las horas a laborar por semana sumando daily_targets de lunes a domingo."""
    if 'HORAS_A_LABORAR' in df_weekly.columns:
        return df_weekly['HORAS_A_LABORAR'].tolist()
    targets = []
    today_date = get_local_now().date()
    for idx, row in df_weekly.iterrows():
        doc_name = row['NOMBRE SUPER VALIDADO']
        inicio = row.get('SEMANA_INICIO')
        fin = row.get('SEMANA_FIN')
        if pd.notna(inicio) and pd.notna(fin):
            target_sum = 0
            curr = pd.Timestamp(inicio)
            end = pd.Timestamp(fin)
            
            # Obtener registros del médico si tenemos df_raw para verificar días futuros trabajados
            doc_entries = pd.DataFrame()
            if df_raw is not None and not df_raw.empty:
                doc_entries = df_raw[df_raw['NOMBRE SUPER VALIDADO'] == doc_name]
                
            while curr <= end:
                # Omitir días futuros si no hay novedades registradas ese día
                if curr.date() > today_date:
                    has_worked = False
                    if not doc_entries.empty:
                        has_worked = not doc_entries[doc_entries['FECHA_CLEAN'].dt.date == curr.date()].empty
                    if not has_worked:
                        curr += pd.Timedelta(days=1)
                        continue
                        
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
def generate_excel_data(df, daily_targets, monthly_targets, cols_to_export_det, df_super=None, df_unfiltered=None, plaza_fija_dates=None):
    output = io.BytesIO()

    df_export_dia = dp.get_consolidated_hours_by_date(df, daily_targets, monthly_targets, df_super, df_unfiltered=df_unfiltered, plaza_fija_dates=plaza_fija_dates)
    if 'HORAS_A_LABORAR' not in df_export_dia.columns:
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
        'TOTAL': 'Total', 'RECARGO_NOCTURNO': 'Recargo Nocturno',
        'CANTIDAD_NOVEDADES': 'Novedades Cubiertas',
        'ESTADO': 'Estado'
    })
    df_export_dia_rename = df_export_dia_rename[
        ['Cédula', 'Médico Supernumerario', 'Fecha', 'Horas a laborar', 'Horas Laboradas', 'Recargo Nocturno', 'Total', 'Novedades Cubiertas', 'Estado']
    ]
    if not df_export_dia_rename.empty:
        totales = {c: [df_export_dia_rename[c].sum()] if c in ['Horas a laborar', 'Horas Laboradas', 'Total', 'Recargo Nocturno', 'Novedades Cubiertas']
                   else (['TOTAL GENERAL'] if c == 'Médico Supernumerario' else [''])
                   for c in df_export_dia_rename.columns}
        df_export_dia_rename = pd.concat([df_export_dia_rename, pd.DataFrame(totales)], ignore_index=True)

    df_export_mes = dp.get_consolidated_hours(df, daily_targets, monthly_targets, df_super, df_unfiltered=df_unfiltered, plaza_fija_dates=plaza_fija_dates)
    if 'HORAS_A_LABORAR' not in df_export_mes.columns:
        df_export_mes['HORAS_A_LABORAR'] = calculate_doctor_target_hours(df_export_mes, df, daily_targets, monthly_targets, df_super=df_super)
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

    df_export_semana = dp.get_consolidated_hours_by_week(df, daily_targets, monthly_targets, df_super, df_unfiltered=df_unfiltered, plaza_fija_dates=plaza_fija_dates)
    if 'HORAS_A_LABORAR' not in df_export_semana.columns:
        df_export_semana['HORAS_A_LABORAR'] = calculate_weekly_target_hours(df_export_semana, daily_targets, df)
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

favicon_path = os.path.join(os.path.dirname(__file__), "favicon.png")
if not os.path.exists(favicon_path):
    favicon_path = os.path.join(os.path.dirname(__file__), "favicon.svg")

st.set_page_config(
    page_title="Control de horas Central de novedades",
    page_icon=favicon_path if os.path.exists(favicon_path) else "⏳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

custom_css = r"""
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="stMarkdown"], label, p, [data-testid="stSidebar"] * {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    body, p, label {
        color: #202124 !important;
    }
    [data-baseweb="select"] { font-family: 'Plus Jakarta Sans', sans-serif; }
    
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Outfit', sans-serif !important;
        font-weight: 700 !important;
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
    div.stButton > button { border-radius: 10px !important; font-weight: 600 !important; transition: all 0.2s ease !important; }
    .element-container:has(.search-btn), .element-container:has(.clear-btn),
    .element-container:has(.export-btn) { display: none !important; }
    .element-container:has(.clear-btn) + .element-container button,
    .element-container:has(.export-btn) + .element-container button,
    .element-container:has(.search-btn) + .element-container button {
        display: inline-flex !important; align-items: center !important; justify-content: center !important;
        width: 100% !important; height: 38px !important; min-height: 38px !important; max-height: 38px !important;
        margin: 0 !important; padding: 0 !important; border-radius: 10px !important; font-weight: 600 !important;
        background-color: #ffffff !important; transition: all 0.2s ease !important; }
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
    /* 1. Botón Buscar: blanco con borde e icono turquesa, hover/click turquesa sólido */
    .element-container:has(.search-btn) + .element-container button {
        background-color: #ffffff !important; color: #199596 !important;
        border: 1.5px solid #199596 !important;
        box-shadow: 0 1px 3px rgba(25, 149, 150, 0.1) !important; }
    .element-container:has(.search-btn) + .element-container button:hover,
    .element-container:has(.search-btn) + .element-container button:active,
    .element-container:has(.search-btn) + .element-container button:focus {
        background-color: #199596 !important; border-color: #199596 !important; color: #ffffff !important;
        box-shadow: 0 4px 10px rgba(25, 149, 150, 0.35) !important; transform: translateY(-1px) !important; }
    /* 2. Botón Limpiar: blanco con borde e icono rojo, hover/click rojo sólido */
    .element-container:has(.clear-btn) + .element-container button {
        background-color: #ffffff !important; color: #ef4444 !important;
        border: 1.5px solid #fca5a5 !important;
        box-shadow: 0 1px 3px rgba(239, 68, 68, 0.08) !important; }
    .element-container:has(.clear-btn) + .element-container button:hover,
    .element-container:has(.clear-btn) + .element-container button:active,
    .element-container:has(.clear-btn) + .element-container button:focus {
        background-color: #ef4444 !important; border-color: #ef4444 !important; color: #ffffff !important;
        box-shadow: 0 4px 10px rgba(239, 68, 68, 0.3) !important; transform: translateY(-1px) !important; }
    /* 3. Botón Exportar: blanco con borde e icono verde lima, hover/click verde lima sólido */
    .element-container:has(.export-btn) + .element-container button {
        background-color: #ffffff !important; color: #85be26 !important;
        border: 1.5px solid #85be26 !important;
        box-shadow: 0 1px 3px rgba(133, 190, 38, 0.1) !important; }
    .element-container:has(.export-btn) + .element-container button:hover,
    .element-container:has(.export-btn) + .element-container button:active,
    .element-container:has(.export-btn) + .element-container button:focus {
        background-color: #85be26 !important; border-color: #85be26 !important; color: #ffffff !important;
        box-shadow: 0 4px 10px rgba(133, 190, 38, 0.35) !important; transform: translateY(-1px) !important; }
    /* ── Tarjetas de Métricas (Referencia Imagen 3) ─────────────────── */
    .kpi-cards-grid {
        display: flex !important;
        flex-wrap: wrap !important;
        gap: 12px !important;
        margin-top: 10px !important;
        margin-bottom: 18px !important;
        width: 100% !important;
    }
    .kpi-metric-card {
        flex: 1 1 calc(16.666% - 12px) !important;
        min-width: 150px !important;
        background: #ffffff !important;
        border-radius: 12px !important;
        padding: 10px 14px !important;
        display: flex !important;
        align-items: center !important;
        gap: 12px !important;
        position: relative !important;
        overflow: hidden !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04) !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease !important;
        box-sizing: border-box !important;
    }
    .kpi-metric-card:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.06) !important;
    }
    .kpi-metric-card::after {
        content: '' !important;
        position: absolute !important;
        right: -14px !important;
        top: 50% !important;
        transform: translateY(-50%) !important;
        width: 48px !important;
        height: 48px !important;
        border-radius: 50% !important;
        opacity: 0.12 !important;
        pointer-events: none !important;
    }
    .kpi-card-icon {
        width: 36px !important;
        height: 36px !important;
        border-radius: 9px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        color: #ffffff !important;
        font-size: 17px !important;
        flex-shrink: 0 !important;
    }
    .kpi-card-content {
        display: flex !important;
        flex-direction: column !important;
        justify-content: center !important;
        min-width: 0 !important;
        z-index: 1 !important;
    }
    .kpi-card-value {
        font-size: 19px !important;
        font-weight: 800 !important;
        line-height: 1.15 !important;
        white-space: nowrap !important;
    }
    .kpi-card-label {
        font-size: 9.5px !important;
        font-weight: 700 !important;
        letter-spacing: 0.5px !important;
        text-transform: uppercase !important;
        margin-top: 3px !important;
        white-space: nowrap !important;
    }
    .kpi-card-medicos { border: 1.5px solid #99f6e4 !important; }
    .kpi-card-medicos .kpi-card-icon { background-color: #199596 !important; }
    .kpi-card-medicos .kpi-card-value { color: #0f766e !important; }
    .kpi-card-medicos .kpi-card-label { color: #14b8a6 !important; }
    .kpi-card-medicos::after { background-color: #14b8a6 !important; }

    .kpi-card-laborar { border: 1.5px solid #cbd5e1 !important; }
    .kpi-card-laborar .kpi-card-icon { background-color: #1e293b !important; }
    .kpi-card-laborar .kpi-card-value { color: #0f172a !important; }
    .kpi-card-laborar .kpi-card-label { color: #64748b !important; }
    .kpi-card-laborar::after { background-color: #1e293b !important; }

    .kpi-card-laboradas { border: 1.5px solid #bef264 !important; }
    .kpi-card-laboradas .kpi-card-icon { background-color: #85be26 !important; }
    .kpi-card-laboradas .kpi-card-value { color: #4d7c0f !important; }
    .kpi-card-laboradas .kpi-card-label { color: #65a30d !important; }
    .kpi-card-laboradas::after { background-color: #85be26 !important; }

    .kpi-card-diferencia-pos { border: 1.5px solid #bbf7d0 !important; }
    .kpi-card-diferencia-pos .kpi-card-icon { background-color: #10b981 !important; }
    .kpi-card-diferencia-pos .kpi-card-value { color: #047857 !important; }
    .kpi-card-diferencia-pos .kpi-card-label { color: #059669 !important; }
    .kpi-card-diferencia-pos::after { background-color: #10b981 !important; }

    .kpi-card-diferencia-neg { border: 1.5px solid #fecaca !important; }
    .kpi-card-diferencia-neg .kpi-card-icon { background-color: #ef4444 !important; }
    .kpi-card-diferencia-neg .kpi-card-value { color: #b91c1c !important; }
    .kpi-card-diferencia-neg .kpi-card-label { color: #dc2626 !important; }
    .kpi-card-diferencia-neg::after { background-color: #ef4444 !important; }

    .kpi-card-diferencia-zero { border: 1.5px solid #fde68a !important; }
    .kpi-card-diferencia-zero .kpi-card-icon { background-color: #f59e0b !important; }
    .kpi-card-diferencia-zero .kpi-card-value { color: #b45309 !important; }
    .kpi-card-diferencia-zero .kpi-card-label { color: #d97706 !important; }
    .kpi-card-diferencia-zero::after { background-color: #f59e0b !important; }

    .kpi-card-recargo { border: 1.5px solid #fed7aa !important; }
    .kpi-card-recargo .kpi-card-icon { background-color: #f59e0b !important; }
    .kpi-card-recargo .kpi-card-value { color: #b45309 !important; }
    .kpi-card-recargo .kpi-card-label { color: #d97706 !important; }
    .kpi-card-recargo::after { background-color: #f59e0b !important; }

    .kpi-card-novedades { border: 1.5px solid #fca5a5 !important; }
    .kpi-card-novedades .kpi-card-icon { background-color: #ef4444 !important; }
    .kpi-card-novedades .kpi-card-value { color: #b91c1c !important; }
    .kpi-card-novedades .kpi-card-label { color: #ef4444 !important; }
    .kpi-card-novedades::after { background-color: #ef4444 !important; }

    /* ── Selectbox (dropdown simple) ─────────────────────────────────── */
    div[data-testid="stSelectbox"] > div > div {
        border: 1.5px solid #d1d5db !important; background-color: #ffffff !important;
        border-radius: 10px !important; transition: all 0.2s ease !important;
        height: 38px !important; min-height: 38px !important; max-height: 38px !important;
        font-size: 13px !important; padding: 0 10px !important;
        display: flex !important; align-items: center !important;
        overflow: hidden !important; }
    div[data-testid="stSelectbox"] > div > div:hover { border-color: #14b8a6 !important; }
    div[data-testid="stSelectbox"] [data-baseweb="select"],
    div[data-testid="stSelectbox"] [data-baseweb="select"] > div { background-color: #ffffff !important; }
    /* ── Multiselect: contenedor redondeado con scroll horizontal ────── */
    div[data-testid="stMultiSelect"] > div > div {
        border: 1.5px solid #d1d5db !important; background-color: #ffffff !important;
        border-radius: 10px !important; transition: all 0.2s ease !important;
        height: 38px !important; min-height: 38px !important; max-height: 38px !important;
        font-size: 13px !important; padding: 0 8px !important;
        display: flex !important; align-items: center !important;
        flex-wrap: nowrap !important; gap: 4px !important;
        overflow-x: auto !important; overflow-y: hidden !important;
        scrollbar-width: none !important; }
    div[data-testid="stMultiSelect"] > div > div::-webkit-scrollbar { display: none !important; }
    div[data-testid="stMultiSelect"] > div > div:hover { border-color: #14b8a6 !important; }
    div[data-testid="stMultiSelect"] > div > div:focus-within,
    div[data-testid="stMultiSelect"] > div > div:focus,
    div[data-testid="stMultiSelect"] > div > div[aria-expanded="true"] {
        border-color: #14b8a6 !important; box-shadow: 0 0 0 2px rgba(20,184,166,0.18) !important; outline: none !important; }
    div[data-testid="stSelectbox"] > div > div:focus-within,
    div[data-testid="stSelectbox"] > div > div:focus {
        border-color: #14b8a6 !important; box-shadow: 0 0 0 2px rgba(20,184,166,0.18) !important; outline: none !important; }
    div[data-testid="stMultiSelect"] * { outline: none !important; }
    div[data-testid="stMultiSelect"] [data-baseweb="select"] { border: none !important; box-shadow: none !important; }
    div[data-testid="stMultiSelect"] [data-baseweb="select"],
    div[data-testid="stMultiSelect"] [data-baseweb="select"] > div { background-color: #ffffff !important; }
    /* ── Tags de multiselect (Mes y Supernumerario) ──────────────────── */
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"],
    div[data-testid*="stMultiSelect" i] span[data-baseweb="tag"],
    div[data-testid*="stMultiSelect" i] div[data-baseweb="tag"],
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] span[role="button"],
    [data-baseweb="tag"] {
        background-color: #199596 !important;
        background: #199596 !important;
        border: 1px solid #147a7b !important;
        border-radius: 6px !important;
        flex-shrink: 0 !important;
        height: 26px !important; min-height: 26px !important; max-height: 26px !important;
        padding: 0 6px 0 8px !important; margin: 0 4px 0 0 !important;
        display: inline-flex !important; align-items: center !important;
        color: #ffffff !important;
    }
    /* Letra blanca nítida para TODO el texto de los tags seleccionados */
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"],
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] *,
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] span,
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] div,
    div[data-testid*="stMultiSelect" i] span[data-baseweb="tag"],
    div[data-testid*="stMultiSelect" i] span[data-baseweb="tag"] *,
    div[data-testid*="stMultiSelect" i] span[data-baseweb="tag"] span,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span *,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span span,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] span[role="button"],
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] span[role="button"] *,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] span[role="button"] span,
    [data-baseweb="tag"],
    [data-baseweb="tag"] *,
    [data-baseweb="tag"] span {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        font-size: 12.5px !important; font-weight: 600 !important;
        white-space: nowrap !important; line-height: 1.2 !important;
        max-width: none !important; overflow: visible !important; text-overflow: unset !important;
    }
    /* Icono de cerrar (x) en blanco */
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] svg,
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] svg *,
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] svg path,
    div[data-testid*="stMultiSelect" i] span[data-baseweb="tag"] svg,
    div[data-testid*="stMultiSelect" i] span[data-baseweb="tag"] svg path,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span svg,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span svg *,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span path,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] span[role="button"] svg,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] span[role="button"] path,
    [data-baseweb="tag"] svg,
    [data-baseweb="tag"] svg *,
    [data-baseweb="tag"] svg path {
        fill: #ffffff !important;
        color: #ffffff !important;
        stroke: #ffffff !important;
    }
    div[data-testid*="stMultiSelect" i] [data-baseweb="tag"]:hover,
    div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span:hover,
    [data-baseweb="tag"]:hover {
        background-color: #147a7b !important;
        background: #147a7b !important;
    }
    /* Opciones dentro del dropdown desplegable (fondo blanco, letra oscura legible) */
    ul[role="listbox"], ul[role="listbox"] li, ul[role="listbox"] *,
    div[data-baseweb="menu"], div[data-baseweb="menu"] * {
        color: #202124 !important;
        -webkit-text-fill-color: #202124 !important;
    }
    ul[role="listbox"] li:hover, ul[role="listbox"] li[aria-selected="true"],
    div[data-baseweb="menu"] li:hover {
        background-color: #f0fdfa !important;
        color: #147a7b !important;
        -webkit-text-fill-color: #147a7b !important;
    }
    ul[role="listbox"] li:hover *, ul[role="listbox"] li[aria-selected="true"] * {
        color: #147a7b !important;
        -webkit-text-fill-color: #147a7b !important;
    }
    /* ── Panel de filtros ──────────────────────────────────────────────── */
    .filter-panel-marker { display: none !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) {
        border: 1px solid #e2e8f0 !important; border-radius: 12px !important;
        background-color: #ffffff !important; padding: 14px 18px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important; margin-bottom: 16px !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important; gap: 12px !important; align-items: flex-end !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) [data-testid="column"] {
        min-width: 0px !important; display: flex !important; flex-direction: column !important; justify-content: flex-end !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) [data-testid="column"] > div {
        display: flex !important; flex-direction: column !important; justify-content: flex-end !important; }
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) div.stButton,
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) div.stDownloadButton,
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) div[data-testid="stSelectbox"],
    div[data-testid="stVerticalBlock"]:has(.filter-panel-marker) div[data-testid="stMultiSelect"] {
        margin: 0 !important; padding: 0 !important; }
    .filter-label {
        font-size: 13.5px !important; font-weight: 600 !important; color: #202124 !important;
        margin-bottom: 5px !important; line-height: 18px !important; height: 18px !important;
        display: block !important; overflow: hidden !important; white-space: nowrap !important; }
    .table-scroll-container { max-height: 440px; overflow-y: auto; overflow-x: auto;
        border: 1px solid #e2e8f0; border-radius: 12px; box-shadow: 0 1px 4px rgba(0,0,0,0.03);
        background: #ffffff; margin-top: 10px; margin-bottom: 15px; }
    .custom-table { width: 100%; border-collapse: collapse; font-size: 12.5px;
        color: #334155; background-color: #ffffff; }
    .custom-table th { background-color: #ffffff !important; color: #0f766e !important;
        font-weight: 700; font-size: 11.5px; text-transform: uppercase; letter-spacing: 0.5px;
        padding: 13px 14px; border-bottom: 2.5px solid #14b8a6 !important;
        border-top: none; border-left: none; border-right: none;
        position: sticky; top: 0; z-index: 10; white-space: nowrap; }
    .custom-table th, .custom-table th * { color: #0f766e !important; }
    .custom-table td { padding: 11px 16px; border-bottom: 1px solid #f1f5f9;
        border-top: none; border-left: none; border-right: none;
        vertical-align: middle; color: #1e293b !important; }
    .custom-table tr:nth-child(even) { background-color: #f8fafc; }
    .custom-table tr:hover td { background-color: #f0fdfa !important; }
    .custom-table th, .custom-table td { text-align: center !important; }
    button[data-testid="stPopoverButton"] { background-color: #ffffff !important;
        border: 1.5px solid #d1d5db !important; color: #1e293b !important;
        border-radius: 10px !important; height: 38px !important; width: 38px !important;
        display: inline-flex !important; align-items: center !important; justify-content: center !important;
        padding: 0 !important; transition: all 0.2s ease !important; font-size: 0 !important; }
    button[data-testid="stPopoverButton"]:hover { background-color: #f0fdfa !important;
        border-color: #14b8a6 !important; color: #0f766e !important; }
    button[data-testid="stPopoverButton"]::before { font-family: "bootstrap-icons" !important;
        content: "\F3E5" !important; font-size: 20px !important; visibility: visible !important;
        color: inherit !important; display: inline-block !important; transition: transform 0.4s ease !important; }
    button[data-testid="stPopoverButton"]:hover::before { transform: rotate(90deg) !important; }
    button[data-testid="stPopoverButton"] * { display: none !important; font-size: 0 !important;
        width: 0 !important; height: 0 !important; overflow: hidden !important; visibility: hidden !important; }
    /* Banner de Cabecera (Imagen 1 style) - Aplicado solo al contenedor stVerticalBlock específico de la cabecera */
    .header-banner-marker { display: none !important; }
    div.element-container:has(.header-banner-marker) { display: none !important; }
    
    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) {
        background: linear-gradient(90deg, #091a32 0%, #102e4d 40%, #1e456e 75%, #3b6b94 100%) !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 16px 24px !important;
        box-shadow: 0 4px 14px rgba(9, 26, 50, 0.16) !important;
        margin-bottom: 20px !important;
    }

    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) h1,
    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) h1 * {
        color: #ffffff !important;
        -webkit-text-fill-color: #ffffff !important;
        font-family: 'Outfit', sans-serif !important;
        font-weight: 700 !important;
        font-size: 26px !important;
        margin: 0 !important;
        line-height: 1.2 !important;
    }

    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) button[data-testid="stPopoverButton"] {
        background-color: rgba(255, 255, 255, 0.12) !important;
        border: 1px solid rgba(255, 255, 255, 0.25) !important;
        color: #ffffff !important;
        border-radius: 10px !important;
        height: 38px !important;
        width: 38px !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.15) !important;
    }
    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) button[data-testid="stPopoverButton"]:hover {
        background-color: rgba(255, 255, 255, 0.22) !important;
        border-color: rgba(255, 255, 255, 0.4) !important;
    }
    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) button[data-testid="stPopoverButton"]::before {
        color: #ffffff !important;
    }

    /* Imagen del logo sin fondo blanco, transparente para integrarse con el degradado */
    div[data-testid="stVerticalBlock"]:has(> div.element-container .header-banner-marker) .header-logo-container img {
        background-color: transparent !important;
        padding: 0 !important;
        border-radius: 0 !important;
        box-shadow: none !important;
        display: inline-block !important;
    }
    /* Ocultar el botón 'Gestionar la aplicación' (viewer badge de Streamlit Cloud) y botones del desarrollador */
    div[data-testid="stConnectionStatus"],
    div[data-testid="stStatusWidget"],
    div[data-testid="stDeveloperTools"],
    div[data-testid="stAppToolbar"],
    *[class*="viewerBadge"],
    *[class*="viewer-badge"],
    *[class*="ViewerBadge"],
    *[class*="styles_viewerBadge"],
    *[class*="styles_viewerBadge_"],
    *[class*="styles_viewer-badge"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
        height: 0px !important;
        width: 0px !important;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────

IS_LOCAL = platform.system() == "Windows"
current_month_name = dp.MESES_MAP.get(get_local_now().month, "Enero")

defaults = {
    'mes_sel': [current_month_name],
    'nombre_sel': [],
    'agrupacion_sel': "Por Mes",
    'mes_sel_draft': [current_month_name],
    'nombre_sel_draft': [],
    'agrupacion_sel_draft': "Por Mes",
    'file_path_input': r"C:\Users\JuanJoseOsorioMolina\OneDrive - U.T SAN VICENTE CES\CONSOLIDADOS\CONSOLIDADO 2026\CONSOLIDADO 2026.xlsx",
    'uploaded_file_name': None,
    'df_raw': None,
    'df_super': None,
    'load_error': None,
    'last_refresh': None,
    'plaza_fija_dates': {},
    'monthly_targets': {},
    'daily_targets': {},
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def reset_filters():
    st.session_state.mes_sel = []
    st.session_state.nombre_sel = []
    st.session_state.agrupacion_sel = "Por Mes"
    st.session_state.mes_sel_draft = []
    st.session_state.nombre_sel_draft = []
    st.session_state.agrupacion_sel_draft = "Por Mes"
    for k in list(st.session_state.keys()):
        if k.startswith("mes_sel_draft_widget_") or k.startswith("nombre_sel_draft_widget_"):
            del st.session_state[k]
    if 'agrupacion_sel_draft_widget' in st.session_state:
        st.session_state.agrupacion_sel_draft_widget = "Por Mes"

def clear_agrupacion():
    st.session_state.agrupacion_sel = "Por Mes"
    st.session_state.agrupacion_sel_draft = "Por Mes"
    if 'agrupacion_sel_draft_widget' in st.session_state:
        st.session_state.agrupacion_sel_draft_widget = "Por Mes"

def clear_mes():
    st.session_state.mes_sel = []
    st.session_state.mes_sel_draft = []
    for k in list(st.session_state.keys()):
        if k.startswith("mes_sel_draft_widget_"):
            del st.session_state[k]

def clear_nombre():
    st.session_state.nombre_sel = []
    st.session_state.nombre_sel_draft = []
    for k in list(st.session_state.keys()):
        if k.startswith("nombre_sel_draft_widget_"):
            del st.session_state[k]


def on_change_nombre():
    active_keys = [k for k in st.session_state.keys() if k.startswith("nombre_sel_draft_widget_")]
    if active_keys:
        val = st.session_state[active_keys[0]]
        if isinstance(val, list):
            if len(val) > 20:  # Si se seleccionaron casi todos / Select all
                val = []
                st.session_state[active_keys[0]] = []
            st.session_state.nombre_sel_draft = val
        else:
            st.session_state.nombre_sel_draft = []

    st.session_state.nombre_sel = st.session_state.nombre_sel_draft
    st.session_state.mes_sel = st.session_state.mes_sel_draft
    st.session_state.agrupacion_sel = st.session_state.agrupacion_sel_draft


def cargar_desde_ruta_local(file_path):
    """Carga los datos desde la ruta del archivo local y busca el histórico en la misma carpeta si existe."""
    safe_path, cleanup = dp.get_safe_file_source(file_path)
    safe_hist_path = None
    cleanup_hist = lambda: None
    try:
        xl_actual = pd.ExcelFile(safe_path, engine='calamine')
        df_raw_actual = dp.load_and_clean_data(xl_actual, preferred_sheet='CONSOLIDADO 2026 NOMINA')

        # Buscar si existe el archivo histórico en la misma carpeta
        dir_name = os.path.dirname(file_path)
        hist_path = os.path.join(dir_name, "CONSOLIDADO 2026 HISTORICO.xlsx")
        df_raw_historico = None
        xl_hist = None
        if os.path.exists(hist_path):
            safe_hist_path, cleanup_hist = dp.get_safe_file_source(hist_path)
            try:
                xl_hist = pd.ExcelFile(safe_hist_path, engine='calamine')
                df_raw_historico = dp.load_and_clean_data(xl_hist, preferred_sheet='CONSOLIDADO 2026 NOMINA HISTORI')
            except Exception:
                pass

        if df_raw_historico is not None and not df_raw_historico.empty:
            meses_en_actual = set(df_raw_actual['MES_NUM'].dropna().unique())
            df_raw_historico = df_raw_historico[
                ~df_raw_historico['MES_NUM'].isin(meses_en_actual)
            ]
            st.session_state.df_raw = pd.concat(
                [df_raw_historico, df_raw_actual], ignore_index=True
            )
            st.session_state.hist_loaded = True
        else:
            st.session_state.df_raw = df_raw_actual
            st.session_state.hist_loaded = False

        df_super_actual = dp.load_supernumerario_sheets(xl_actual)
        df_super_hist = dp.load_supernumerario_sheets(xl_hist) if xl_hist is not None else None
        if df_super_hist is not None and not df_super_hist.empty:
            if df_super_actual is not None and not df_super_actual.empty:
                st.session_state.df_super = pd.concat([df_super_hist, df_super_actual], ignore_index=True).drop_duplicates()
            else:
                st.session_state.df_super = df_super_hist
        else:
            st.session_state.df_super = df_super_actual

        plaza_actual = dp.load_plaza_fija_dates(xl_actual)
        plaza_hist = dp.load_plaza_fija_dates(xl_hist) if xl_hist is not None else {}
        if plaza_hist:
            plaza_hist.update(plaza_actual)
            st.session_state.plaza_fija_dates = plaza_hist
        else:
            st.session_state.plaza_fija_dates = plaza_actual

        m, d = dp.load_calendar_targets(xl_actual)
        st.session_state.monthly_targets = m
        st.session_state.daily_targets = d
        st.session_state.load_error = None
        st.session_state.last_refresh = get_local_now().strftime('%d/%m/%Y %H:%M:%S')
    finally:
        cleanup()
        cleanup_hist()


# ── Popover de Configuración en Cabecera ──────────────────────────────────────

with st.container():
    st.markdown('<div class="header-banner-marker"></div>', unsafe_allow_html=True)
    col_config, col_title = st.columns([0.6, 9.4], vertical_alignment="center")

    with col_config:
        with st.popover("", help="Configuración de Origen de Datos"):
            st.markdown("<h3 style='margin:0 0 10px 0; font-family:Outfit,sans-serif; font-weight:700; color:#0b3c5d;'>⚙️ Configuración de Datos</h3>", unsafe_allow_html=True)

            if 'hist_loaded' in st.session_state:
                if st.session_state.hist_loaded:
                    st.success("✅ Archivo Histórico Cargado")
                else:
                    st.error("❌ Archivo Histórico NO Cargado")

            # Mostrar última actualización si existe
            if st.session_state.last_refresh:
                st.info(f"🕐 Última carga: **{st.session_state.last_refresh}**")

            # Botón para recargar desde SharePoint (siempre visible)
            if get_onedrive_config():
                if st.button("🔄 Recargar desde SharePoint", use_container_width=True):
                    st.session_state.df_raw = None
                    st.session_state.load_error = None
                    cargar_desde_onedrive(force_refresh=True)
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
                            cargar_desde_ruta_local(file_path)
                            st.rerun()
                        except Exception as e:
                            st.session_state.load_error = str(e)

                uploaded_file = st.file_uploader("O sube el archivo manualmente:", type=["xlsx", "xls"])
                if uploaded_file is not None:
                    try:
                        file_bytes = io.BytesIO(uploaded_file.read())
                        xl = pd.ExcelFile(file_bytes, engine='calamine')
                        st.session_state.df_raw = dp.load_and_clean_data(xl)
                        st.session_state.df_super = dp.load_supernumerario_sheets(xl)
                        st.session_state.plaza_fija_dates = dp.load_plaza_fija_dates(xl)
                        m, d = dp.load_calendar_targets(xl)
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
            "<h1 style='font-family:\"Outfit\",sans-serif; font-weight:700; color:#ffffff !important; -webkit-text-fill-color:#ffffff !important; font-size:28px; margin:0;'><span style='color:#ffffff !important; -webkit-text-fill-color:#ffffff !important;'>Control de horas Central de novedades</span></h1>",
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
                cargar_desde_ruta_local(file_path)
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

if isinstance(st.session_state.nombre_sel, list) and st.session_state.nombre_sel:
    df_filtrado = df_filtrado[df_filtrado['NOMBRE SUPER VALIDADO'].isin(st.session_state.nombre_sel)]
elif isinstance(st.session_state.nombre_sel, str) and st.session_state.nombre_sel != "Todos":
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
    st.markdown('''<div class="filter-panel-marker"></div>
    <style>
        div[data-testid*="stMultiSelect" i] [data-baseweb="tag"],
        div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] *,
        div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span,
        div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span * {
            color: #ffffff !important;
            -webkit-text-fill-color: #ffffff !important;
            fill: #ffffff !important;
            font-weight: 600 !important;
        }
        div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] svg,
        div[data-testid*="stMultiSelect" i] [data-baseweb="tag"] svg *,
        div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span svg,
        div[data-testid*="stMultiSelect" i] div[data-baseweb="select"] > div > div > span svg * {
            color: #ffffff !important;
            fill: #ffffff !important;
            stroke: #ffffff !important;
        }
    </style>''', unsafe_allow_html=True)
    c1, c2, c3, c4, c5, c6 = st.columns([3, 3, 3, 0.6, 0.6, 0.6], gap="small", vertical_alignment="bottom")

    with c1:
        st.markdown("<div class='filter-label'>Agrupar por:</div>", unsafe_allow_html=True)
        agrupacion_options = ["Por Día", "Por Semana", "Por Mes"]
        agrupacion_idx = agrupacion_options.index(st.session_state.agrupacion_sel_draft) if st.session_state.agrupacion_sel_draft in agrupacion_options else 0
        agrupacion_sel_draft = st.selectbox("Agrupar por:", options=agrupacion_options, index=agrupacion_idx, key="agrupacion_sel_draft_widget", label_visibility="collapsed")
        st.session_state.agrupacion_sel_draft = agrupacion_sel_draft

    with c2:
        st.markdown("<div class='filter-label'>Mes:</div>", unsafe_allow_html=True)
        active_keys = [k for k in st.session_state.keys() if k.startswith("mes_sel_draft_widget_")]
        num_items = 0
        if active_keys:
            val = st.session_state[active_keys[0]]
            if isinstance(val, list):
                num_items = len(val)
                st.session_state.mes_sel_draft = val
        mes_key = f"mes_sel_draft_widget_{num_items}"
        default_meses = [m for m in st.session_state.mes_sel_draft if m in meses_disponibles]
        meses_sel_draft = st.multiselect("Mes:", options=meses_disponibles, default=default_meses, key=mes_key, placeholder="Seleccionar...", label_visibility="collapsed")
        st.session_state.mes_sel_draft = meses_sel_draft

    df_para_filtros = df_raw.copy()
    if st.session_state.mes_sel_draft:
        df_para_filtros = df_para_filtros[df_para_filtros['MES'].isin(st.session_state.mes_sel_draft)]
    nombres_disponibles = sorted(df_para_filtros['NOMBRE SUPER VALIDADO'].dropna().unique().tolist())

    with c3:
        st.markdown("<div class='filter-label'>Supernumerario:</div>", unsafe_allow_html=True)
        active_nom_keys = [k for k in st.session_state.keys() if k.startswith("nombre_sel_draft_widget_")]
        num_nom_items = 0
        if active_nom_keys:
            val = st.session_state[active_nom_keys[0]]
            if isinstance(val, list):
                if len(nombres_disponibles) > 0 and len(val) >= len(nombres_disponibles):
                    val = []
                    st.session_state[active_nom_keys[0]] = []
                num_nom_items = len(val)
                st.session_state.nombre_sel_draft = val
        nom_key = f"nombre_sel_draft_widget_{num_nom_items}"
        cur_nom_draft = st.session_state.nombre_sel_draft if isinstance(st.session_state.nombre_sel_draft, list) else []
        default_nombres = [n for n in cur_nom_draft if n in nombres_disponibles]
        nombres_sel_draft = st.multiselect("Supernumerario:", options=nombres_disponibles, default=default_nombres, key=nom_key, placeholder="Todos", label_visibility="collapsed", on_change=on_change_nombre)
        st.session_state.nombre_sel_draft = nombres_sel_draft

    with c4:
        st.markdown("<div class='filter-label' style='visibility:hidden;'>&nbsp;</div>", unsafe_allow_html=True)
        st.markdown('<div class="search-btn">', unsafe_allow_html=True)
        if st.button("Buscar", key="btn_search", use_container_width=True):
            st.session_state.mes_sel = st.session_state.mes_sel_draft
            st.session_state.nombre_sel = st.session_state.nombre_sel_draft
            st.session_state.agrupacion_sel = st.session_state.agrupacion_sel_draft
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with c5:
        st.markdown("<div class='filter-label' style='visibility:hidden;'>&nbsp;</div>", unsafe_allow_html=True)
        st.markdown('<div class="clear-btn">', unsafe_allow_html=True)
        st.button("Borrar Filtro Supernumerario", key="btn_clear", help="Borrar filtro Supernumerario", use_container_width=True, on_click=clear_nombre)
        st.markdown('</div>', unsafe_allow_html=True)

    with c6:
        st.markdown("<div class='filter-label' style='visibility:hidden;'>&nbsp;</div>", unsafe_allow_html=True)
        st.markdown('<div class="export-btn">', unsafe_allow_html=True)
        cols_to_export_det = [c for c in detalle_cols_base if c in df_filtrado.columns]
        daily_targets = st.session_state.get('daily_targets', {})
        monthly_targets = st.session_state.get('monthly_targets', {})
        excel_data = generate_excel_data(df_filtrado, daily_targets, monthly_targets, cols_to_export_det, df_super=st.session_state.get('df_super'), df_unfiltered=df_raw, plaza_fija_dates=st.session_state.get('plaza_fija_dates'))
        st.download_button(
            label="Exportar Excel",
            data=excel_data,
            file_name=f"CONSOLIDADO_HORAS_SUPERNUMERARIOS_{get_local_now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_export",
            use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<hr style='margin-top:-15px; margin-bottom:10px; border:0; border-top:1px solid #e0e0e0;'/>", unsafe_allow_html=True)

# ── Tabla consolidada ─────────────────────────────────────────────────────────

daily_targets = st.session_state.get('daily_targets', {})
monthly_targets = st.session_state.get('monthly_targets', {})
df_super = st.session_state.get('df_super')

agrupacion_vista = st.session_state.agrupacion_sel

if agrupacion_vista == "Por Día":
    tabla_consolidada_vista = dp.get_consolidated_hours_by_date(df_filtrado, daily_targets, monthly_targets, df_super, df_unfiltered=df_raw, plaza_fija_dates=st.session_state.get('plaza_fija_dates'))
    if 'HORAS_A_LABORAR' not in tabla_consolidada_vista.columns:
        tabla_consolidada_vista['HORAS_A_LABORAR'] = tabla_consolidada_vista.apply(
            lambda r: 7.33 if (r['NOMBRE SUPER VALIDADO'] == 'SEBASTIAN GIL GALLEGO' and daily_targets.get(r['FECHA_STR'], 0) == 7)
                      else daily_targets.get(r['FECHA_STR'], 0),
            axis=1
        )
    tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        tabla_consolidada_vista[col] = tabla_consolidada_vista[col].round(0).astype(int)
elif agrupacion_vista == "Por Semana":
    tabla_consolidada_vista = dp.get_consolidated_hours_by_week(df_filtrado, daily_targets, monthly_targets, df_super, df_unfiltered=df_raw, plaza_fija_dates=st.session_state.get('plaza_fija_dates'))
    if 'HORAS_A_LABORAR' not in tabla_consolidada_vista.columns:
        tabla_consolidada_vista['HORAS_A_LABORAR'] = calculate_weekly_target_hours(
            tabla_consolidada_vista, daily_targets, df_filtrado
        )
    tabla_consolidada_vista['TOTAL'] = tabla_consolidada_vista['HORAS_TOTALES'] - tabla_consolidada_vista['HORAS_A_LABORAR']
    for col in ['HORAS_TOTALES', 'HORAS_A_LABORAR', 'TOTAL']:
        tabla_consolidada_vista[col] = tabla_consolidada_vista[col].round(0).astype(int)
else:
    tabla_consolidada_vista = dp.get_consolidated_hours(df_filtrado, daily_targets, monthly_targets, df_super, df_unfiltered=df_raw, plaza_fija_dates=st.session_state.get('plaza_fija_dates'))
    if 'HORAS_A_LABORAR' not in tabla_consolidada_vista.columns:
        tabla_consolidada_vista['HORAS_A_LABORAR'] = calculate_doctor_target_hours(
            tabla_consolidada_vista, df_filtrado, daily_targets, monthly_targets, df_super=df_super
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
    if 'RECARGO_NOCTURNO' in tabla_consolidada_vista.columns:
        tot_recargo = tabla_consolidada_vista['RECARGO_NOCTURNO'].sum()
    elif 'RECARGO NOCTURNO ORDINARIO' in df_filtrado.columns:
        tot_recargo = df_filtrado['RECARGO NOCTURNO ORDINARIO'].sum()
    else:
        tot_recargo = 0.0

    dif_card_class = "kpi-card-diferencia-pos" if tot_diferencia > 0 else ("kpi-card-diferencia-neg" if tot_diferencia < 0 else "kpi-card-diferencia-zero")
    dif_icon = "bi-arrow-up-right" if tot_diferencia > 0 else ("bi-arrow-down-right" if tot_diferencia < 0 else "bi-dash")
    dif_sign = "+" if tot_diferencia > 0 else ""
    recargo_fmt = f"{tot_recargo:,.0f} hrs" if (tot_recargo % 1 == 0) else f"{tot_recargo:,.1f} hrs"

    col_res1, col_res2 = st.columns([1, 1])
    with col_res1:
        st.markdown("<div style='font-size:15px; font-weight:700; color:#0f766e; display:flex; align-items:center; gap:8px;'><i class='bi bi-table'></i> Resultados</div>", unsafe_allow_html=True)
    with col_res2:
        st.markdown(f"<div style='text-align:right; font-size:13px; color:#64748b; font-weight:500;'>{len(tabla_consolidada_vista):,} registros</div>", unsafe_allow_html=True)

    st.markdown(
        f'''
        <div class="kpi-cards-grid">
            <div class="kpi-metric-card kpi-card-medicos">
                <div class="kpi-card-icon"><i class="bi bi-people-fill"></i></div>
                <div class="kpi-card-content">
                    <span class="kpi-card-value">{tot_medicos_activos:,}</span>
                    <span class="kpi-card-label">MÉDICOS ACTIVOS</span>
                </div>
            </div>
            <div class="kpi-metric-card kpi-card-laborar">
                <div class="kpi-card-icon"><i class="bi bi-calendar-check-fill"></i></div>
                <div class="kpi-card-content">
                    <span class="kpi-card-value">{tot_horas_a_laborar:,.0f} hrs</span>
                    <span class="kpi-card-label">HORAS A LABORAR</span>
                </div>
            </div>
            <div class="kpi-metric-card kpi-card-laboradas">
                <div class="kpi-card-icon"><i class="bi bi-clock-fill"></i></div>
                <div class="kpi-card-content">
                    <span class="kpi-card-value">{tot_horas_laboradas:,.0f} hrs</span>
                    <span class="kpi-card-label">HORAS LABORADAS</span>
                </div>
            </div>
            <div class="kpi-metric-card {dif_card_class}">
                <div class="kpi-card-icon"><i class="bi {dif_icon}"></i></div>
                <div class="kpi-card-content">
                    <span class="kpi-card-value">{dif_sign}{tot_diferencia:,.0f} hrs</span>
                    <span class="kpi-card-label">DIFERENCIA</span>
                </div>
            </div>
            <div class="kpi-metric-card kpi-card-recargo">
                <div class="kpi-card-icon"><i class="bi bi-hourglass-split"></i></div>
                <div class="kpi-card-content">
                    <span class="kpi-card-value">{recargo_fmt}</span>
                    <span class="kpi-card-label">RECARGO</span>
                </div>
            </div>
            <div class="kpi-metric-card kpi-card-novedades">
                <div class="kpi-card-icon"><i class="bi bi-shield-fill-check"></i></div>
                <div class="kpi-card-content">
                    <span class="kpi-card-value">{tot_novedades:,}</span>
                    <span class="kpi-card-label">NOVEDADES CUBIERTAS</span>
                </div>
            </div>
        </div>
        ''',
        unsafe_allow_html=True
    )

if agrupacion_vista == "Por Día":
    tabla_display = tabla_consolidada_vista.rename(columns={
        'FECHA_STR': 'FECHA', 'CEDULA_FINAL': 'CÉDULA',
        'NOMBRE SUPER VALIDADO': 'MÉDICO SUPERNUMERARIO',
        'HORAS_A_LABORAR': 'HORAS A LABORAR', 'HORAS_TOTALES': 'HORAS LABORADAS',
        'TOTAL': 'TOTAL', 'RECARGO_NOCTURNO': 'RECARGO NOCTURNO',
        'CANTIDAD_NOVEDADES': 'NOVEDADES CUBIERTAS',
        'ESTADO': 'ESTADO'
    })
    cols_show = ['CÉDULA', 'MÉDICO SUPERNUMERARIO', 'FECHA', 'HORAS A LABORAR', 'HORAS LABORADAS', 'RECARGO NOCTURNO', 'TOTAL', 'NOVEDADES CUBIERTAS', 'ESTADO']
elif agrupacion_vista == "Por Semana":
    tabla_display = tabla_consolidada_vista.rename(columns={
        'CEDULA_FINAL': 'CÉDULA', 'NOMBRE SUPER VALIDADO': 'MÉDICO SUPERNUMERARIO',
        'SEMANA': 'SEMANA', 'HORAS_A_LABORAR': 'HORAS A LABORAR',
        'HORAS_TOTALES': 'HORAS LABORADAS', 'TOTAL': 'TOTAL',
        'RECARGO_NOCTURNO': 'RECARGO NOCTURNO',
        'CANTIDAD_NOVEDADES': 'NOVEDADES CUBIERTAS'
    })
    cols_show = ['CÉDULA', 'MÉDICO SUPERNUMERARIO', 'SEMANA', 'HORAS A LABORAR', 'HORAS LABORADAS', 'RECARGO NOCTURNO', 'TOTAL', 'NOVEDADES CUBIERTAS']
else:
    tabla_display = tabla_consolidada_vista.rename(columns={
        'CEDULA_FINAL': 'CÉDULA', 'NOMBRE SUPER VALIDADO': 'MÉDICO SUPERNUMERARIO',
        'MES': 'MES', 'HORAS_A_LABORAR': 'HORAS A LABORAR',
        'HORAS_TOTALES': 'HORAS LABORADAS', 'TOTAL': 'TOTAL',
        'RECARGO_NOCTURNO': 'RECARGO NOCTURNO',
        'CANTIDAD_NOVEDADES': 'NOVEDADES CUBIERTAS'
    })
    cols_show = ['CÉDULA', 'MÉDICO SUPERNUMERARIO', 'MES', 'HORAS A LABORAR', 'HORAS LABORADAS', 'RECARGO NOCTURNO', 'TOTAL', 'NOVEDADES CUBIERTAS']

tabla_display_formatted = tabla_display[cols_show].copy()
for col in ['HORAS A LABORAR', 'HORAS LABORADAS', 'TOTAL', 'NOVEDADES CUBIERTAS']:
    if col in tabla_display_formatted.columns:
        tabla_display_formatted[col] = pd.to_numeric(tabla_display_formatted[col], errors='coerce').fillna(0).round(0).astype(int)

if 'RECARGO NOCTURNO' in tabla_display_formatted.columns:
    rec_vals = pd.to_numeric(tabla_display_formatted['RECARGO NOCTURNO'], errors='coerce').fillna(0)
    if (rec_vals % 1 == 0).all():
        tabla_display_formatted['RECARGO NOCTURNO'] = rec_vals.astype(int)
    else:
        tabla_display_formatted['RECARGO NOCTURNO'] = rec_vals.round(1)

html_table = tabla_display_formatted.to_html(index=False, classes='custom-table', escape=False)

th_index = 0
def add_onclick_to_th(match):
    global th_index
    content = match.group(1)
    res = f'<th onclick="sortTable({th_index})" style="cursor:pointer;user-select:none;" title="Haz clic para ordenar">{content} <span style="font-size:10px; opacity:0.35; color:#14b8a6;">▲▼</span></th>'
    th_index += 1
    return res

html_table = re.sub(r'<th>(.*?)</th>', add_onclick_to_th, html_table)

iframe_template = r"""<!DOCTYPE html><html><head>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.3/font/bootstrap-icons.min.css">
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
body{margin:0;padding:0;font-family:'Plus Jakarta Sans',sans-serif;background-color:transparent;}
.table-scroll-container{max-height:440px;overflow-y:auto;overflow-x:auto;border:1px solid #e2e8f0;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,0.03);background:#ffffff;}
.custom-table{width:100%;border-collapse:collapse;font-size:12.5px;color:#334155;background-color:#ffffff;}
.custom-table th{background-color:#ffffff!important;color:#0f766e!important;font-weight:700;font-size:11.5px;text-transform:uppercase;letter-spacing:0.5px;padding:13px 14px;border-bottom:2.5px solid #14b8a6!important;border-top:none;border-left:none;border-right:none;position:sticky;top:0;z-index:10;cursor:pointer;user-select:none;text-align:center!important;white-space:nowrap;}
.custom-table th *{color:#0f766e!important;}
.custom-table th:hover{background-color:#f8fafc!important;}
.custom-table td{padding:12px 14px;border-bottom:1px solid #f1f5f9;border-top:none;border-left:none;border-right:none;vertical-align:middle;color:#334155!important;text-align:center!important;line-height:1.4;}
.custom-table tr:hover td{background-color:#f0fdfa!important;}
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
        th.innerHTML=th.innerHTML.replace(/ <span style="font-size: 10px; color: #14b8a6;">(▲|▼)<\/span>/g,"");
        th.innerHTML=th.innerHTML.replace(/ <span style="font-size:10px; opacity:0.35; color:#14b8a6;">▲▼<\/span>/g,"");
        if(idx===columnIndex){th.innerHTML+=` <span style="font-size: 10px; color: #14b8a6;">${dir==='asc'?'▲':'▼'}</span>`;}
        else{th.innerHTML+=` <span style="font-size:10px; opacity:0.35; color:#14b8a6;">▲▼</span>`;}
    });
}
</script>
</head><body>
<div class="table-scroll-container">__TABLE_HTML__</div>
</body></html>"""

iframe_content = iframe_template.replace("__TABLE_HTML__", html_table)
st.components.v1.html(iframe_content, height=450, scrolling=False)

st.markdown("<div style='margin-top:25px;font-size:13px;color:#888888;'>© 2026 - Unión para la Salud y la Vida </div>", unsafe_allow_html=True)