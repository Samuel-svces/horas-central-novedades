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
    
    # 3. Validar columnas clave requeridas
    col_name = 'NOMBRE SUPER VALIDADO'
    col_horas = 'HORAS TOTALES DECIMAL'
    
    # Buscar fecha en REVISION POR CENTRAL DE NOVEDADES o en F inic Novedad
    col_fecha = None
    for option in ['REVISION POR CENTRAL DE NOVEDADES', 'F inic Novedad', 'FECHA']:
        if option in df.columns:
            col_fecha = option
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


def get_active_daily_df(df, daily_targets, monthly_targets, df_super=None):
    """
    Genera un DataFrame a nivel diario para cada médico activo en el mes,
    cubriendo todo su periodo activo (alineado a semanas y acotado al mes).
    """
    if df.empty or not daily_targets:
        return pd.DataFrame(columns=[
            'FECHA_CLEAN', 'FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO',
            'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'MES', 'MES_NUM', 'HORAS_A_LABORAR'
        ])
        
    df_worked = df.groupby(['FECHA_CLEAN', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False).agg(
        HORAS_TRABAJADAS=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    
    medicos_meses = df.groupby(['CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'MES', 'MES_NUM'], as_index=False).size()
    daily_rows = []
    
    for idx, row in medicos_meses.iterrows():
        cedula = row['CEDULA_FINAL']
        doc_name = row['NOMBRE SUPER VALIDADO']
        month_name = row['MES']
        month_num = row['MES_NUM']
        
        found_in_super = False
        if df_super is not None and not df_super.empty:
            doc_norm = normalize_name(doc_name)
            super_entries = df_super[
                (df_super['NOMBRE_NORM'] == doc_norm) &
                (df_super['FECHA_CLEAN'].dt.month == month_num)
            ]
            if not super_entries.empty:
                max_date = super_entries['FECHA_CLEAN'].max()
                min_date = super_entries['FECHA_CLEAN'].min()
                found_in_super = True
                
        if not found_in_super:
            doc_entries = df[
                (df['NOMBRE SUPER VALIDADO'] == doc_name) &
                (df['MES_NUM'] == month_num)
            ]
            max_date = doc_entries['FECHA_CLEAN'].max()
            min_date = doc_entries['FECHA_CLEAN'].min()
            
        if pd.notna(min_date) and pd.notna(max_date):
            if min_date.day <= 7 and max_date.day >= (max_date.days_in_month - 6):
                min_date_aligned = pd.Timestamp(year=min_date.year, month=month_num, day=1)
                max_date_aligned = pd.Timestamp(year=max_date.year, month=month_num, day=max_date.days_in_month)
            else:
                min_date_aligned = min_date - pd.to_timedelta(min_date.dayofweek, unit='D')
                max_date_aligned = max_date + pd.to_timedelta(6 - max_date.dayofweek, unit='D')
            
            curr = min_date_aligned
            while curr <= max_date_aligned:
                if curr.month == month_num:
                    date_str = curr.strftime('%d/%m/%Y')
                    worked_day = df_worked[
                        (df_worked['FECHA_CLEAN'] == curr) &
                        (df_worked['NOMBRE SUPER VALIDADO'] == doc_name)
                    ]
                    if not worked_day.empty:
                        horas_trabajadas = worked_day.iloc[0]['HORAS_TRABAJADAS']
                        novedades = worked_day.iloc[0]['CANTIDAD_NOVEDADES']
                    else:
                        horas_trabajadas = 0.0
                        novedades = 0
                        
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
                
    return pd.DataFrame(daily_rows)


def get_consolidated_hours(df, daily_targets=None, monthly_targets=None, df_super=None):
    """
    Agrupa y sumariza las horas totales trabajadas por Cédula, Nombre y Mes.
    """
    if daily_targets is not None and monthly_targets is not None:
        df_daily = get_active_daily_df(df, daily_targets, monthly_targets, df_super)
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

def get_consolidated_hours_by_date(df, daily_targets=None, monthly_targets=None, df_super=None):
    """
    Agrupa y sumariza las horas totales trabajadas por Fecha (YYYY-MM-DD), Cédula y Nombre.
    """
    if daily_targets is not None and monthly_targets is not None:
        df_daily = get_active_daily_df(df, daily_targets, monthly_targets, df_super)
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


def get_consolidated_hours_by_week(df, daily_targets=None, monthly_targets=None, df_super=None):
    """
    Agrupa y sumariza las horas totales trabajadas por Semana, Cédula y Nombre.
    """
    if daily_targets is not None and monthly_targets is not None:
        df_daily = get_active_daily_df(df, daily_targets, monthly_targets, df_super)
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


def load_supernumerario_sheet(file_source):
    """
    Carga la hoja SUPERNUMERARIO si existe en el archivo.
    Retorna un DataFrame limpio o None si no existe.
    """
    safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            xl = pd.ExcelFile(safe_source)
            sheet_names = xl.sheet_names
            target_sheet = None
            for s in sheet_names:
                if s.strip().upper() in ['SUPERNUMERARIO', 'SUPERNUMERARIOS']:
                    target_sheet = s
                    break
            
            if target_sheet:
                df = pd.read_excel(xl, sheet_name=target_sheet)
                df.columns = [str(col).strip() for col in df.columns]
                
                # Normalizar columnas Fecha, Nombre
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
                    return df[['FECHA_CLEAN', 'NOMBRE_NORM']].dropna()
        except Exception:
            pass
    finally:
        cleanup()
    return None

