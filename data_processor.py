# -*- coding: utf-8 -*-
"""
Módulo de Procesamiento de Datos para la Auditoría de Horas de Médicos Supernumerarios.
Este script contiene toda la lógica de backend para cargar, limpiar y transformar los datos.
"""

import pandas as pd
import numpy as np
import os
import shutil
import tempfile
import requests
import msal

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
    # Omitir filas donde la columna original SUPERNUMERARIOS esté vacía,
    # ya que indica que no se asignó un médico supernumerario real para el turno.
    if 'SUPERNUMERARIOS' in df.columns:
        df = df[df['SUPERNUMERARIOS'].notnull()]
        df = df[df['SUPERNUMERARIOS'].astype(str).str.strip() != '']
        
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
    col_cedula_orig = 'CEDULA SUPERNUMERARIO'
    if col_cedula_orig in df.columns:
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


def get_consolidated_hours(df):
    """
    Agrupa y sumariza las horas totales trabajadas por Cédula, Nombre y Mes.
    :param df: DataFrame limpio
    :return: DataFrame agrupado
    """
    # Agrupar por las columnas solicitadas
    grouped = df.groupby(['CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'MES', 'MES_NUM'], as_index=False).agg(
        HORAS_TOTALES=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    
    # Ordenar por mes (numérico) y luego por nombre para mejor legibilidad
    grouped = grouped.sort_values(by=['MES_NUM', 'NOMBRE SUPER VALIDADO']).reset_index(drop=True)
    # Eliminar columna auxiliar de ordenamiento de mes si se desea, pero la conservamos para ordenar la visualización
    return grouped

def get_consolidated_hours_by_date(df):
    """
    Agrupa y sumariza las horas totales trabajadas por Fecha (YYYY-MM-DD), Cédula y Nombre.
    :param df: DataFrame limpio
    :return: DataFrame agrupado por fecha
    """
    df_copy = df.copy()
    
    # Formatear la fecha como string YYYY-MM-DD para agrupar y ordenar cronológicamente
    df_copy['FECHA_STR'] = df_copy['FECHA_CLEAN'].dt.strftime('%Y-%m-%d')
    df_copy['FECHA_STR'] = df_copy['FECHA_STR'].fillna('Sin Fecha')
    
    # Agrupar por fecha, cédula y nombre
    grouped = df_copy.groupby(['FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False).agg(
        HORAS_TOTALES=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    
    # Ordenar por fecha (descendente, la más reciente primero) y luego por nombre
    grouped = grouped.sort_values(by=['FECHA_STR', 'NOMBRE SUPER VALIDADO'], ascending=[False, True]).reset_index(drop=True)
    
    # Convertir a formato Dia/mes/año (DD/MM/YYYY) para la visualización y exportación
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


