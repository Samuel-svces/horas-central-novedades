# -*- coding: utf-8 -*-
"""
Módulo de Procesamiento de Datos para la Auditoría de Horas de Médicos Supernumerarios.
Este script contiene toda la lógica de backend para cargar, limpiar y transformar los datos.
"""

import pandas as pd
import numpy as np
import os
import io
import shutil
import tempfile
import requests
import msal
import unicodedata
import re


# Mapeo de meses de inglés/número a español
MESES_MAP = {
    1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril', 5: 'Mayo', 6: 'Junio',
    7: 'Julio', 8: 'Agosto', 9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
}

def clean_cedula_val(val):
    """
    Limpia el valor de la cédula para retornar un string entero o None si es inválido.
    """
    if pd.isna(val):
        return None
    
    # Quitar saltos de línea, retornos de carro y espacios en blanco
    val_str = str(val).replace('\r', '').replace('\n', '').strip()
    if not val_str or val_str.lower() in ['nan', 'null', 'none', '0', '0.0']:
        return None
        
    try:
        # Si tiene formato float (ej. 1001456294.0), convertir a int primero para quitar el decimal
        float_val = float(val_str)
        if float_val.is_integer():
            return str(int(float_val))
        return str(float_val)
    except ValueError:
        return val_str


def get_safe_file_source(file_source):
    """
    Si el origen es una ruta de archivo local y está bloqueada (ej. abierta en Excel),
    hace una copia temporal para permitir la lectura sin lanzar PermissionError.
    """
    temp_path = None
    safe_source = file_source

    if isinstance(file_source, str) and os.path.exists(file_source):
        try:
            # Probar si se puede abrir para lectura
            with open(file_source, 'rb'):
                pass
        except PermissionError:
            try:
                # Conservar la extensión original (.xlsx, .csv, etc.)
                _, ext = os.path.splitext(file_source)
                fd, temp_path = tempfile.mkstemp(suffix=ext)
                os.close(fd)
                shutil.copy2(file_source, temp_path)
                safe_source = temp_path
            except Exception:
                pass

    def cleanup():
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass

    return safe_source, cleanup


def get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """
    Obtiene un access token de Azure AD usando client credentials flow.
    No requiere login del usuario — usa las credenciales de la app registrada.
    """
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        error_desc = result.get("error_description", "Sin descripción")
        raise ValueError(f"No se pudo obtener el token de Azure AD: {error_desc}")
    return result["access_token"]


def download_excel_from_onedrive(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    drive_id: str,
    file_id: str
) -> io.BytesIO:
    """
    Descarga el archivo Excel desde OneDrive/SharePoint usando Microsoft Graph API.
    Retorna un BytesIO listo para pasarle a load_and_clean_data().
    """
    token = get_access_token(tenant_id, client_id, client_secret)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.get(url, headers=headers, timeout=30)
    
    if response.status_code != 200:
        raise ValueError(
            f"Error al descargar el archivo desde OneDrive. "
            f"Status: {response.status_code} — {response.text[:200]}"
        )
    return io.BytesIO(response.content)

def load_and_clean_data(file_source):
    """
    Carga el archivo (CSV o Excel) y realiza todo el pipeline de limpieza y transformación.
    :param file_source: Ruta al archivo (str) o un objeto tipo File de Streamlit.
    :return: DataFrame limpio
    """
    safe_source, cleanup = get_safe_file_source(file_source)
    try:
        # 1. Cargar el archivo según su tipo
        if isinstance(safe_source, str) and safe_source.endswith(('.csv', '.CSV')):
            # Leer CSV detectando el separador (punto y coma o coma)
            try:
                df = pd.read_csv(safe_source, sep=';', encoding='utf-8')
            except Exception:
                df = pd.read_csv(safe_source, sep=',', encoding='utf-8')
        else:
            # Por defecto tratar como Excel
            try:
                # Leer específicamente la hoja solicitada por el usuario
                df = pd.read_excel(safe_source, sheet_name='CONSOLIDADO 2026 NOMINA')
            except Exception as e:
                # Fallback en caso de que no exista la pestaña específica o sea otro tipo de Excel
                try:
                    xl = pd.ExcelFile(safe_source)
                    # Seleccionar la primera hoja que contenga "NOMINA" o en su defecto la primera disponible
                    sheets = xl.sheet_names
                    nomina_sheets = [s for s in sheets if 'NOMINA' in s.upper()]
                    sheet_to_load = nomina_sheets[0] if nomina_sheets else sheets[0]
                    df = pd.read_excel(safe_source, sheet_name=sheet_to_load)
                except Exception as ex:
                    raise ValueError(f"Error al leer el archivo Excel: {str(ex)}")
    finally:
        cleanup()

    # 2. Renombrar columnas para estandarizar en caso de ligeras variaciones de mayúsculas/minúsculas
    df.columns = [str(col).strip() for col in df.columns]

    # Auto-recuperación: Si la primera columna (columna A) viene vacía o sin nombre en OneDrive,
    # la renombramos a 'REVISION POR CENTRAL DE NOVEDADES' que es donde están las fechas de revisión.
    if len(df.columns) > 0:
        first_col = df.columns[0]
        first_col_norm = str(first_col).strip().upper()
        if first_col_norm in ['', 'NAN'] or first_col_norm.startswith('UNNAMED'):
            cols = list(df.columns)
            cols[0] = 'REVISION POR CENTRAL DE NOVEDADES'
            df.columns = cols
    
    # 3. Validar columnas clave requeridas
    col_name = 'NOMBRE SUPER VALIDADO'
    col_horas = 'HORAS TOTALES DECIMAL'
    
    # Buscar fecha en REVISION POR CENTRAL DE NOVEDADES o en F inic Novedad (insensible a mayúsculas y acentos)
    col_fecha = None
    
    # 1. Buscar coincidencia para REVISION POR CENTRAL DE NOVEDADES
    for col in df.columns:
        col_norm = str(col).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        if col_norm == 'REVISION POR CENTRAL DE NOVEDADES':
            col_fecha = col
            break
            
    # 2. Si no se encontró, buscar F INIC NOVEDAD o FECHA
    if not col_fecha:
        for col in df.columns:
            col_norm = str(col).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
            if col_norm in ['F INIC NOVEDAD', 'FECHA']:
                col_fecha = col
                break
            
    if col_name not in df.columns or col_horas not in df.columns:
        raise ValueError(
            f"El archivo no contiene las columnas requeridas. "
            f"Columnas encontradas: {df.columns.tolist()}"
        )

    # 4. Filtrar y limpiar registros vacíos o no válidos de Médicos
    df = df[df[col_name].notnull()]
    df[col_name] = df[col_name].astype(str).str.strip()
    df = df[df[col_name] != '']

    
    # Excluir registros temporales, de eliminación o sin supernumerario
    excluir_patterns = ['ELIMINAR', 'SIN SUPERNUMERARIO', 'SIN PROCESAR']
    pattern_regex = '|'.join(excluir_patterns)
    df = df[~df[col_name].str.upper().str.contains(pattern_regex, na=False)]
    # Adicionalmente, evitar registros que solo digan "SIN"
    df = df[df[col_name].str.upper() != 'SIN']

    # 5. Transformar y limpiar HORAS TOTALES DECIMAL
    df[col_horas] = pd.to_numeric(df[col_horas], errors='coerce').fillna(0.0)

    # 6. Procesar Fechas y Meses
    if col_fecha:
        df['FECHA_CLEAN'] = pd.to_datetime(df[col_fecha], errors='coerce')
    else:
        df['FECHA_CLEAN'] = pd.NaT

    # Filtrar solo registros del año 2026
    df = df[df['FECHA_CLEAN'].dt.year == 2026]
        
    # Extraer mes numérico y luego mapearlo al nombre en español
    df['MES_NUM'] = df['FECHA_CLEAN'].dt.month
    df['MES'] = df['MES_NUM'].map(MESES_MAP).fillna('Sin Mes')

    # 7. Limpieza e Inteligencia de Cédulas (Mapear Cédulas Faltantes)
    col_cedula_orig = None
    for col in df.columns:
        col_norm = str(col).strip().upper()
        # Normalizar acentos comunes
        col_norm = col_norm.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        if col_norm in [
            'CEDULA SUPERNUMERARIO', 'CEDULA SUPERNUMERARIOS', 'CEDULAS SUPERNUMERARIOS', 
            'CEDULA SUPER', 'CEDULAS SUPER', 'CEDULA_SUPERNUMERARIO', 'CEDULA_SUPERNUMERARIOS',
            'IDENTIFICACION SUPERNUMERARIO', 'IDENTIFICACION_SUPERNUMERARIO'
        ]:
            col_cedula_orig = col
            break

    if col_cedula_orig:
        df['CEDULA_PROCESADA'] = df[col_cedula_orig].apply(clean_cedula_val)
        
        # Crear base de conocimientos: Mapear Nombre -> Cédula a partir de registros que sí tienen cédula
        valid_mappings = df[df['CEDULA_PROCESADA'].notnull()]
        name_to_cedula = valid_mappings.groupby(col_name)['CEDULA_PROCESADA'].first().to_dict()
        
        # Autocompletar cédulas faltantes basándose en el nombre
        df['CEDULA_FINAL'] = df.apply(
            lambda row: name_to_cedula.get(row[col_name], row['CEDULA_PROCESADA']),
            axis=1
        )
        # Rellenar las que queden completamente vacías con 'SIN CÉDULA'
        df['CEDULA_FINAL'] = df['CEDULA_FINAL'].fillna('SIN CÉDULA')
    else:
        df['CEDULA_FINAL'] = 'SIN CÉDULA'

    # 8. Limpiar columna DOCUMENTO para evitar problemas de tipos de datos en la visualización
    if 'DOCUMENTO' in df.columns:
        df['DOCUMENTO'] = df['DOCUMENTO'].apply(clean_cedula_val)

    return df


def get_active_daily_df(df, daily_targets, monthly_targets, df_super=None, df_unfiltered=None, plaza_fija_dates=None):
    """
    Genera un DataFrame a nivel diario para cada médico activo en el mes,
    cubriendo todo su periodo activo.
    """
    if df.empty or not daily_targets:
        return pd.DataFrame(columns=[
            'FECHA_CLEAN', 'FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO',
            'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'MES', 'MES_NUM', 'HORAS_A_LABORAR'
        ])
        
    df_ref = df_unfiltered if df_unfiltered is not None else df
        
    df_worked = df.groupby(['FECHA_CLEAN', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False).agg(
        HORAS_TRABAJADAS=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    
    # Crear un diccionario para búsquedas rápidas de días trabajados (O(1))
    worked_dict = {
        (row['FECHA_CLEAN'], row['NOMBRE SUPER VALIDADO']): (row['HORAS_TRABAJADAS'], row['CANTIDAD_NOVEDADES'])
        for _, row in df_worked.iterrows()
    }
    
    medicos_meses = df.groupby(['CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'MES', 'MES_NUM'], as_index=False).size()
    daily_rows = []
    
    for idx, row in medicos_meses.iterrows():
        cedula = row['CEDULA_FINAL']
        doc_name = row['NOMBRE SUPER VALIDADO']
        month_name = row['MES']
        month_num = row['MES_NUM']
        
        if pd.isna(month_num):
            continue
        month_num = int(month_num)
        
        # 1. Buscar todas las fechas de novedades del médico en todo el año (df_ref)
        all_doc_dates = df_ref[df_ref['NOMBRE SUPER VALIDADO'] == doc_name]['FECHA_CLEAN'].dropna().tolist()
        
        # 2. Buscar todas las fechas en las hojas de supernumerarios (df_super) en todo el año
        if df_super is not None and not df_super.empty:
            doc_norm = normalize_name(doc_name)
            super_dates = df_super[df_super['NOMBRE_NORM'] == doc_norm]['FECHA_CLEAN'].dropna().tolist()
            all_doc_dates.extend(super_dates)
            
        # 3. Determinar min_date y max_date del mes, pero expandiendo si hay actividad en otros meses
        if all_doc_dates:
            all_doc_dates = pd.to_datetime(all_doc_dates)
            abs_min = all_doc_dates.min()
            abs_max = all_doc_dates.max()
            
            # Fechas por defecto para este mes específico
            doc_entries_month = df[(df['NOMBRE SUPER VALIDADO'] == doc_name) & (df['MES_NUM'] == month_num)]
            month_min = doc_entries_month['FECHA_CLEAN'].min()
            month_max = doc_entries_month['FECHA_CLEAN'].max()
            
            if df_super is not None and not df_super.empty:
                doc_norm = normalize_name(doc_name)
                super_entries_month = df_super[(df_super['NOMBRE_NORM'] == doc_norm) & (df_super['FECHA_CLEAN'].dt.month == month_num)]
                if not super_entries_month.empty:
                    month_min = min(month_min, super_entries_month['FECHA_CLEAN'].min()) if pd.notna(month_min) else super_entries_month['FECHA_CLEAN'].min()
                    month_max = max(month_max, super_entries_month['FECHA_CLEAN'].max()) if pd.notna(month_max) else super_entries_month['FECHA_CLEAN'].max()
            
            # Si no hay registros en este mes, no tiene periodo activo este mes
            if pd.isna(month_min) or pd.isna(month_max):
                min_date_aligned = pd.NaT
                max_date_aligned = pd.NaT
            else:
                # Si tiene registros en meses ANTERIORES, asumimos que está activo desde el día 1 de este mes
                if abs_min.month < month_num or abs_min.year < 2026:
                    min_date_aligned = pd.Timestamp(year=2026, month=month_num, day=1)
                else:
                    min_date_aligned = month_min
                    
                # Si tiene registros en meses POSTERIORES, asumimos que está activo hasta el último día de este mes
                days_in_month = pd.Period(f"2026-{month_num:02d}").days_in_month
                if abs_max.month > month_num or abs_max.year > 2026:
                    max_date_aligned = pd.Timestamp(year=2026, month=month_num, day=days_in_month)
                else:
                    # Si su último registro en este mes está en la última semana, también consideramos hasta fin de mes
                    if month_max.day >= (days_in_month - 6):
                        max_date_aligned = pd.Timestamp(year=2026, month=month_num, day=days_in_month)
                    else:
                        max_date_aligned = month_max
        else:
            min_date_aligned = pd.NaT
            max_date_aligned = pd.NaT
            
        # 4. Ajustar por fecha de traslado a plaza fija (si existe)
        if plaza_fija_dates:
            doc_norm = normalize_name(doc_name)
            limit_date = plaza_fija_dates.get(doc_norm)
            if pd.notna(limit_date):
                limit_date = pd.to_datetime(limit_date)
                # La fecha de traslado (limit_date) ya es plaza fija, por lo que el límite como supernumerario es limit_date - 1 día
                target_limit = limit_date - pd.Timedelta(days=1)
                if pd.notna(max_date_aligned):
                    max_date_aligned = min(max_date_aligned, target_limit)
                else:
                    max_date_aligned = target_limit
            
        if pd.notna(min_date_aligned) and pd.notna(max_date_aligned):
            # Obtener fecha actual en Colombia (UTC-5)
            from datetime import datetime, timezone, timedelta
            today_dt = datetime.now(timezone(timedelta(hours=-5))).date()
            
            curr = min_date_aligned
            while curr <= max_date_aligned:
                if curr.month == month_num:
                    date_str = curr.strftime('%d/%m/%Y')
                    worked_info = worked_dict.get((curr, doc_name))
                    if worked_info:
                        horas_trabajadas, novedades = worked_info
                    else:
                        horas_trabajadas = 0.0
                        novedades = 0
                        
                    # Omitir días futuros si no hay novedades registradas ese día
                    if curr.date() > today_dt and horas_trabajadas == 0:
                        curr += pd.Timedelta(days=1)
                        continue
                        
                    val = daily_targets.get(date_str, 0)
                    if doc_name == 'SEBASTIAN GIL GALLEGO' and val == 7:
                        horas_a_laborar = 7.33
                    else:
                        horas_a_laborar = val
                        
                    daily_rows.append({
                        'FECHA_CLEAN': curr,
                        'FECHA_STR': date_str,
                        'CEDULA_FINAL': cedula,
                        'NOMBRE SUPER VALIDADO': doc_name,
                        'HORAS_TOTALES': horas_trabajadas,
                        'CANTIDAD_NOVEDADES': novedades,
                        'MES': month_name,
                        'MES_NUM': month_num,
                        'HORAS_A_LABORAR': horas_a_laborar
                    })
                curr += pd.Timedelta(days=1)
                
    df_res = pd.DataFrame(daily_rows)
    if not df_res.empty:
        # Filtrar días no laborados que no tienen horas a laborar (ej. domingos y festivos sin novedades)
        df_res = df_res[~((df_res['HORAS_TOTALES'] == 0) & (df_res['HORAS_A_LABORAR'] == 0))]
    return df_res


def get_consolidated_hours(df, daily_targets=None, monthly_targets=None, df_super=None, df_unfiltered=None, plaza_fija_dates=None):
    """
    Agrupa y sumariza las horas totales trabajadas por Cédula, Nombre y Mes.
    """
    if daily_targets is not None and monthly_targets is not None:
        df_daily = get_active_daily_df(df, daily_targets, monthly_targets, df_super, df_unfiltered=df_unfiltered, plaza_fija_dates=plaza_fija_dates)
        if df_daily.empty:
            return pd.DataFrame(columns=[
                'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'MES', 'MES_NUM',
                'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'HORAS_A_LABORAR'
            ])
        grouped = df_daily.groupby(
            ['CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'MES', 'MES_NUM'], as_index=False
        ).agg(
            HORAS_TOTALES=('HORAS_TOTALES', 'sum'),
            CANTIDAD_NOVEDADES=('CANTIDAD_NOVEDADES', 'sum'),
            HORAS_A_LABORAR=('HORAS_A_LABORAR', 'sum')
        )
        grouped = grouped.sort_values(by=['MES_NUM', 'NOMBRE SUPER VALIDADO']).reset_index(drop=True)
        return grouped

    grouped = df.groupby(['CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'MES', 'MES_NUM'], as_index=False).agg(
        HORAS_TOTALES=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    grouped = grouped.sort_values(by=['MES_NUM', 'NOMBRE SUPER VALIDADO']).reset_index(drop=True)
    return grouped

def get_consolidated_hours_by_date(df, daily_targets=None, monthly_targets=None, df_super=None, df_unfiltered=None, plaza_fija_dates=None):
    """
    Agrupa y sumariza las horas totales trabajadas por Fecha (YYYY-MM-DD), Cédula y Nombre.
    """
    if daily_targets is not None and monthly_targets is not None:
        df_daily = get_active_daily_df(df, daily_targets, monthly_targets, df_super, df_unfiltered=df_unfiltered, plaza_fija_dates=plaza_fija_dates)
        if df_daily.empty:
            return pd.DataFrame(columns=[
                'FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO',
                'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'HORAS_A_LABORAR'
            ])
        df_daily = df_daily.sort_values(by=['FECHA_CLEAN', 'NOMBRE SUPER VALIDADO'], ascending=[False, True]).reset_index(drop=True)
        return df_daily[['FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'HORAS_A_LABORAR']]

    df_copy = df.copy()
    df_copy['FECHA_STR'] = df_copy['FECHA_CLEAN'].dt.strftime('%Y-%m-%d')
    df_copy['FECHA_STR'] = df_copy['FECHA_STR'].fillna('Sin Fecha')
    grouped = df_copy.groupby(['FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False).agg(
        HORAS_TOTALES=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    grouped = grouped.sort_values(by=['FECHA_STR', 'NOMBRE SUPER VALIDADO'], ascending=[False, True]).reset_index(drop=True)
    temp_date = pd.to_datetime(grouped['FECHA_STR'], format='%Y-%m-%d', errors='coerce')
    grouped['FECHA_STR'] = temp_date.dt.strftime('%d/%m/%Y').fillna(grouped['FECHA_STR'])
    return grouped

def load_calendar_targets(file_source):
    """
    Carga la pestaña FESTIVOS y genera dinámicamente las horas a laborar (metas)
    por día y por mes para el año 2026, imitando la consulta CALENDARIO_METAS.
    Lunes-Sábado hábil = 7 horas, Domingo o Festivo = 0 horas.
    """
    safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            # Intentar leer la pestaña FESTIVOS
            if isinstance(safe_source, str) and safe_source.endswith(('.csv', '.CSV')):
                fest_df = pd.DataFrame(columns=['FECHA'])
            else:
                fest_df = pd.read_excel(safe_source, sheet_name='FESTIVOS')
        except Exception:
            # Fallback en caso de no encontrarse la pestaña
            fest_df = pd.DataFrame(columns=['FECHA'])
    finally:
        cleanup()
        
    # Limpiar fechas de festivos
    festivos = []
    if not fest_df.empty and 'FECHA' in fest_df.columns:
        fest_df['FECHA_CLEAN'] = pd.to_datetime(fest_df['FECHA'], errors='coerce')
        festivos = fest_df['FECHA_CLEAN'].dt.date.dropna().tolist()
        
    # Generar todos los días del año 2026 (365 días)
    dates = pd.date_range('2026-01-01', '2026-12-31')
    df_cal = pd.DataFrame({'Fecha': dates})
    df_cal['DiaSemana'] = df_cal['Fecha'].dt.dayofweek
    df_cal['EsFestivo'] = df_cal['Fecha'].dt.date.isin(festivos)
    
    # CÁLCULO DE META: 0 horas si es Domingo (6) o Festivo, 7 horas si es día hábil
    df_cal['HORAS_META'] = df_cal.apply(
        lambda r: 0 if r['DiaSemana'] == 6 or r['EsFestivo'] else 7, axis=1
    )
    
    # Agrupación por mes numérico para metas mensuales
    df_cal['MES_NUM'] = df_cal['Fecha'].dt.month
    monthly_targets = df_cal.groupby('MES_NUM')['HORAS_META'].sum().to_dict()
    
    # Metas diarias por fecha formateada string (Dia/mes/año)
    df_cal['FECHA_STR'] = df_cal['Fecha'].dt.strftime('%d/%m/%Y')
    daily_targets = df_cal.set_index('FECHA_STR')['HORAS_META'].to_dict()
    
    return monthly_targets, daily_targets


def get_consolidated_hours_by_week(df, daily_targets=None, monthly_targets=None, df_super=None, df_unfiltered=None, plaza_fija_dates=None):
    """
    Agrupa y sumariza las horas totales trabajadas por Semana, Cédula y Nombre.
    """
    if daily_targets is not None and monthly_targets is not None:
        df_daily = get_active_daily_df(df, daily_targets, monthly_targets, df_super, df_unfiltered=df_unfiltered, plaza_fija_dates=plaza_fija_dates)
        if df_daily.empty:
            return pd.DataFrame(columns=[
                'SEMANA_INICIO', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO',
                'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'MES_NUM',
                'SEMANA_FIN', 'SEMANA', 'FECHA_INICIO_STR', 'FECHA_FIN_STR', 'HORAS_A_LABORAR'
            ])
        df_copy = df_daily.copy()
        df_copy['SEMANA_INICIO'] = df_copy['FECHA_CLEAN'] - pd.to_timedelta(
            df_copy['FECHA_CLEAN'].dt.dayofweek, unit='D'
        )
        df_copy['SEMANA_INICIO'] = df_copy['SEMANA_INICIO'].dt.normalize()
        grouped = df_copy.groupby(
            ['SEMANA_INICIO', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False
        ).agg(
            HORAS_TOTALES=('HORAS_TOTALES', 'sum'),
            CANTIDAD_NOVEDADES=('CANTIDAD_NOVEDADES', 'sum'),
            HORAS_A_LABORAR=('HORAS_A_LABORAR', 'sum'),
            MES_NUM=('MES_NUM', 'first')
        )
        grouped['SEMANA_FIN'] = grouped['SEMANA_INICIO'] + pd.Timedelta(days=5)
        grouped['SEMANA'] = grouped.apply(
            lambda r: f"{r['SEMANA_INICIO'].strftime('%d/%m')} - {r['SEMANA_FIN'].strftime('%d/%m/%Y')}"
            if pd.notna(r['SEMANA_INICIO']) else 'Sin Fecha',
            axis=1
        )
        grouped['FECHA_INICIO_STR'] = grouped['SEMANA_INICIO'].dt.strftime('%d/%m/%Y').fillna('')
        grouped['FECHA_FIN_STR'] = grouped['SEMANA_FIN'].dt.strftime('%d/%m/%Y').fillna('')
        grouped = grouped.sort_values(
            by=['SEMANA_INICIO', 'NOMBRE SUPER VALIDADO'], ascending=[False, True]
        ).reset_index(drop=True)
        return grouped

    df_copy = df.copy()
    df_copy = df_copy[df_copy['FECHA_CLEAN'].notna()].copy()
    if df_copy.empty:
        return pd.DataFrame(columns=[
            'SEMANA_INICIO', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO',
            'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'MES_NUM',
            'SEMANA_FIN', 'SEMANA', 'FECHA_INICIO_STR', 'FECHA_FIN_STR'
        ])
    df_copy['SEMANA_INICIO'] = df_copy['FECHA_CLEAN'] - pd.to_timedelta(
        df_copy['FECHA_CLEAN'].dt.dayofweek, unit='D'
    )
    df_copy['SEMANA_INICIO'] = df_copy['SEMANA_INICIO'].dt.normalize()
    grouped = df_copy.groupby(
        ['SEMANA_INICIO', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False
    ).agg(
        HORAS_TOTALES=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count'),
        MES_NUM=('MES_NUM', 'first')
    )
    grouped['SEMANA_FIN'] = grouped['SEMANA_INICIO'] + pd.Timedelta(days=5)
    grouped['SEMANA'] = grouped.apply(
        lambda r: f"{r['SEMANA_INICIO'].strftime('%d/%m')} - {r['SEMANA_FIN'].strftime('%d/%m/%Y')}"
        if pd.notna(r['SEMANA_INICIO']) else 'Sin Fecha',
        axis=1
    )
    grouped['FECHA_INICIO_STR'] = grouped['SEMANA_INICIO'].dt.strftime('%d/%m/%Y').fillna('')
    grouped['FECHA_FIN_STR'] = grouped['SEMANA_FIN'].dt.strftime('%d/%m/%Y').fillna('')
    grouped = grouped.sort_values(
        by=['SEMANA_INICIO', 'NOMBRE SUPER VALIDADO'], ascending=[False, True]
    ).reset_index(drop=True)
    return grouped


def normalize_name(name):
    if pd.isna(name):
        return ""
    # Quitar acentos y eñes
    name_str = str(name)
    normalized = unicodedata.normalize('NFKD', name_str).encode('ASCII', 'ignore').decode('ASCII')
    # Convertir a mayúsculas y quitar caracteres que no sean letras, números o espacios
    normalized = re.sub(r'[^A-Z0-9\s]', '', normalized.upper())
    # Colapsar múltiples espacios
    return ' '.join(normalized.split())


def load_supernumerario_sheets(file_source):
    """
    Carga las hojas mensuales de supernumerarios (SUPERNUMERARIOS ENERO,
    SUPERNUMERARIOS FEBRERO, etc.) y las combina en un solo DataFrame.
    Cada hoja tiene columnas: Cedula, Nombre, Fecha, Zona, Cis.
    Retorna un DataFrame con columnas [FECHA_CLEAN, NOMBRE_NORM, MES_NUM] o None.
    """
    safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            xl = pd.ExcelFile(safe_source)
            sheet_names = xl.sheet_names

            # Mapeo de nombre de mes en español a número
            mes_a_num = {
                'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4,
                'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8,
                'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12
            }

            all_dfs = []
            for sheet in sheet_names:
                sheet_upper = sheet.strip().upper()
                # Buscar hojas con patrón "SUPERNUMERARIOS {MES}"
                if not sheet_upper.startswith('SUPERNUMERARIOS '):
                    continue
                mes_name = sheet_upper.replace('SUPERNUMERARIOS ', '').strip()
                month_num = mes_a_num.get(mes_name)
                if month_num is None:
                    continue

                df = pd.read_excel(xl, sheet_name=sheet)
                df.columns = [str(col).strip() for col in df.columns]

                # Buscar columnas Fecha y Nombre
                col_fecha = None
                col_nombre = None
                for col in df.columns:
                    col_norm = col.upper()
                    if col_norm == 'FECHA':
                        col_fecha = col
                    elif col_norm in ['NOMBRE', 'NOMBRE SUPERNUMERARIO', 'MEDICO']:
                        col_nombre = col

                if col_fecha and col_nombre:
                    df['FECHA_CLEAN'] = pd.to_datetime(df[col_fecha], errors='coerce')
                    df['NOMBRE_NORM'] = df[col_nombre].apply(normalize_name)
                    df['MES_NUM'] = month_num
                    all_dfs.append(df[['FECHA_CLEAN', 'NOMBRE_NORM', 'MES_NUM']].dropna(subset=['FECHA_CLEAN', 'NOMBRE_NORM']))

            if all_dfs:
                return pd.concat(all_dfs, ignore_index=True)
        except Exception:
            pass
    finally:
        cleanup()
    return None


def load_plaza_fija_dates(file_source):
    """
    Carga la hoja 'PLAZA FIJA' o 'CONSOLIDADO PLAZAS FIJAS' (si existe en el Excel)
    y retorna un diccionario {NOMBRE_NORMALIZADO: FECHA_TRASLADO} para cada médico.
    """
    safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            xl = pd.ExcelFile(safe_source)
            sheet_names = xl.sheet_names
            
            # Buscar una hoja que contenga "PLAZA FIJA" o "PLAZAS FIJAS"
            target_sheet = None
            for s in sheet_names:
                s_upper = s.strip().upper()
                if 'PLAZA FIJA' in s_upper or 'PLAZAS FIJAS' in s_upper:
                    target_sheet = s
                    break
                    
            if not target_sheet:
                return {}
                
            df = pd.read_excel(xl, sheet_name=target_sheet)
            df.columns = [str(col).strip() for col in df.columns]
            
            # Buscar la columna del supernumerario (SUPER ASIGNADO)
            col_super = None
            for col in df.columns:
                col_norm = col.upper()
                if 'SUPER ASIGNADO' in col_norm:
                    col_super = col
                    break
            
            # Buscar la columna de la fecha de traslado (FECHA DE TRASLADO)
            col_fecha = None
            for col in df.columns:
                col_norm = col.upper()
                if 'FECHA DE TRASLADO' in col_norm or 'FECHA TRASLADO' in col_norm or 'TRASLADO' in col_norm:
                    col_fecha = col
                    break
            
            # Si no se encuentra con nombre, usar índices por defecto si el layout coincide
            if not col_super and len(df.columns) > 8:
                col_super = df.columns[8]
            if not col_fecha and len(df.columns) > 9:
                col_fecha = df.columns[9]
                
            if col_super and col_fecha:
                # Limpiar y normalizar los nombres y parsear las fechas
                df['FECHA_PARSED'] = pd.to_datetime(df[col_fecha], errors='coerce')
                df = df.dropna(subset=[col_super, 'FECHA_PARSED'])
                
                # Crear mapeo normalizado
                mapping = {}
                for _, r in df.iterrows():
                    name_norm = normalize_name(r[col_super])
                    if name_norm:
                        mapping[name_norm] = r['FECHA_PARSED']
                return mapping
        except Exception:
            pass
    finally:
        cleanup()
    return {}


