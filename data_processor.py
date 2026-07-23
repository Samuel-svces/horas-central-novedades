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

MESES_ABR = {
    1: 'Ene', 2: 'Feb', 3: 'Mar', 4: 'Abr', 5: 'May', 6: 'Jun',
    7: 'Jul', 8: 'Ago', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dic'
}

def format_semana_str(inicio, fin):
    if pd.isna(inicio) or pd.isna(fin):
        return 'Sin Fecha'
    m_ini = MESES_ABR.get(inicio.month, '')
    m_fin = MESES_ABR.get(fin.month, '')
    if inicio.month == fin.month:
        return f"{inicio.strftime('%d')} - {fin.strftime('%d')} {m_ini}"
    else:
        return f"{inicio.strftime('%d')} {m_ini} - {fin.strftime('%d')} {m_fin}"


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


def download_excel_from_sharepoint(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    sharepoint_host: str,
    site_path: str,
    file_drive_path: str
) -> io.BytesIO:
    """
    Descarga el archivo Excel desde SharePoint usando Microsoft Graph API.

    Parámetros en Streamlit Secrets:
        SHAREPOINT_HOST      = "sanvicenteces2.sharepoint.com"
        SHAREPOINT_SITE_PATH = "/sites/CENTRALDENOVEDADESCONSOLIDADOS"
        SHAREPOINT_FILE_PATH = "/CONSOLIDADOS/CONSOLIDADO 2026/CONSOLIDADO 2026.xlsx"
                               (ruta dentro de la biblioteca, sin prefijo de sitio ni de biblioteca)

    Retorna un BytesIO listo para pasarle a load_and_clean_data().
    """
    token = get_access_token(tenant_id, client_id, client_secret)
    headers = {"Authorization": f"Bearer {token}"}

    # 1. Obtener el ID del sitio de SharePoint
    site_url = f"https://graph.microsoft.com/v1.0/sites/{sharepoint_host}:{site_path}"
    site_resp = requests.get(site_url, headers=headers, timeout=30)
    if site_resp.status_code != 200:
        raise ValueError(
            f"No se pudo obtener el sitio de SharePoint. "
            f"Status: {site_resp.status_code} — {site_resp.text[:300]}"
        )
    site_id = site_resp.json()["id"]

    # 2. Obtener el drive raíz (Documentos compartidos) del sitio
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives_resp = requests.get(drives_url, headers=headers, timeout=30)
    if drives_resp.status_code != 200:
        raise ValueError(
            f"No se pudo listar las bibliotecas del sitio. "
            f"Status: {drives_resp.status_code} — {drives_resp.text[:300]}"
        )
    drives = drives_resp.json().get("value", [])

    # Buscar el drive "Documentos compartidos" / "Documents"; si no, usar el primero
    target_drive = None
    for d in drives:
        drive_name = d.get("name", "").upper()
        if drive_name in ("DOCUMENTOS COMPARTIDOS", "DOCUMENTS", "SHARED DOCUMENTS"):
            target_drive = d
            break
    if target_drive is None and drives:
        target_drive = drives[0]
    if target_drive is None:
        raise ValueError("No se encontró ninguna biblioteca de documentos en el sitio de SharePoint.")

    drive_id = target_drive["id"]

    # 3. Usar directamente la ruta dentro del drive (sin prefijo de sitio ni de biblioteca)
    # Ejemplo: "/CONSOLIDADOS/CONSOLIDADO 2026/CONSOLIDADO 2026.xlsx"
    path_in_drive = file_drive_path.lstrip("/")

    # 4. Descargar el archivo por su ruta dentro del drive
    import urllib.parse
    encoded_path = urllib.parse.quote(path_in_drive)
    content_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/content"
    response = requests.get(content_url, headers=headers, timeout=60)

    if response.status_code != 200:
        raise ValueError(
            f"Error al descargar el archivo desde SharePoint. "
            f"Ruta buscada: '{path_in_drive}' — "
            f"Status: {response.status_code} — {response.text[:300]}"
        )
    return io.BytesIO(response.content)



def download_excel_from_onedrive(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    drive_id: str,
    file_id: str
) -> io.BytesIO:
    """
    Descarga el archivo Excel desde OneDrive/SharePoint usando Microsoft Graph API
    mediante drive_id y file_id (método clásico).
    Retorna un BytesIO listo para pasarle a load_and_clean_data().
    """
    token = get_access_token(tenant_id, client_id, client_secret)
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content"
    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(url, headers=headers, timeout=60)

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
    if isinstance(file_source, pd.ExcelFile):
        safe_source = file_source
        cleanup = lambda: None
    else:
        safe_source, cleanup = get_safe_file_source(file_source)
    try:
        # 1. Cargar el archivo según su tipo
        if isinstance(safe_source, pd.ExcelFile):
            try:
                df = pd.read_excel(safe_source, sheet_name='CONSOLIDADO 2026 NOMINA')
            except Exception as e:
                try:
                    sheets = safe_source.sheet_names
                    nomina_sheets = [s for s in sheets if 'NOMINA' in s.upper()]
                    sheet_to_load = nomina_sheets[0] if nomina_sheets else sheets[0]
                    df = pd.read_excel(safe_source, sheet_name=sheet_to_load)
                except Exception as ex:
                    raise ValueError(f"Error al leer el archivo Excel: {str(ex)}")
        elif isinstance(safe_source, str) and safe_source.endswith(('.csv', '.CSV')):
            # Leer CSV detectando el separador (punto y coma o coma)
            try:
                df = pd.read_csv(safe_source, sep=';', encoding='utf-8')
            except Exception:
                df = pd.read_csv(safe_source, sep=',', encoding='utf-8')
        else:
            # Por defecto tratar como Excel
            try:
                df = pd.read_excel(safe_source, sheet_name='CONSOLIDADO 2026 NOMINA', engine='calamine')
            except Exception as e:
                try:
                    xl = pd.ExcelFile(safe_source, engine='calamine')
                    sheets = xl.sheet_names
                    nomina_sheets = [s for s in sheets if 'NOMINA' in s.upper()]
                    sheet_to_load = nomina_sheets[0] if nomina_sheets else sheets[0]
                    df = pd.read_excel(xl, sheet_name=sheet_to_load)
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

    # Auto-completar NOMBRE SUPER VALIDADO
    if col_name in df.columns:
        known_supers = set(df[col_name].dropna().astype(str).str.strip().str.upper().unique())
        try:
            if isinstance(safe_source, pd.ExcelFile):
                if 'SUPERNUMERARIOS' in safe_source.sheet_names:
                    super_df = pd.read_excel(safe_source, sheet_name='SUPERNUMERARIOS')
                    super_df.columns = [str(c).strip().upper() for c in super_df.columns]
                    if 'NOMBRE' in super_df.columns:
                        known_supers.update(super_df['NOMBRE'].dropna().astype(str).str.strip().str.upper().unique())
            else:
                if hasattr(safe_source, 'seek'):
                    safe_source.seek(0)
                if not isinstance(safe_source, str) or not safe_source.endswith(('.csv', '.CSV')):
                    super_df = pd.read_excel(safe_source, sheet_name='SUPERNUMERARIOS', engine='calamine')
                    super_df.columns = [str(c).strip().upper() for c in super_df.columns]
                    if 'NOMBRE' in super_df.columns:
                        known_supers.update(super_df['NOMBRE'].dropna().astype(str).str.strip().str.upper().unique())
        except Exception:
            pass
        
        known_supers = {n for n in known_supers if n and n.strip() != '' and n.upper() != 'NAN'}
        known_supers_norm = set(normalize_name(n) for n in known_supers)
        
        # 1. Preferir la columna SUPERNUMERARIOS (el médico supernumerario que cubre la novedad)
        if 'SUPERNUMERARIOS' in df.columns:
            def fill_missing_from_supernumerarios(row):
                val_super_val = row.get(col_name)
                if pd.isna(val_super_val) or str(val_super_val).strip() == '' or str(val_super_val).strip().upper() == 'NAN':
                    val_super = row.get('SUPERNUMERARIOS')
                    if pd.notna(val_super) and str(val_super).strip() != '' and str(val_super).strip().upper() != 'NAN':
                        super_norm = normalize_name(val_super)
                        if super_norm in known_supers_norm:
                            for name in known_supers:
                                if normalize_name(name) == super_norm:
                                    return name
                            return str(val_super).strip().upper()
                return val_super_val
            df[col_name] = df.apply(fill_missing_from_supernumerarios, axis=1)

        # 2. Como fallback, usar la columna MEDICOS (para novedades propias del supernumerario, ej. incapacidades)
        if 'MEDICOS' in df.columns:
            def fill_missing_super_validated(row):
                val_super = row.get(col_name)
                if pd.isna(val_super) or str(val_super).strip() == '' or str(val_super).strip().upper() == 'NAN':
                    val_medicos = row.get('MEDICOS')
                    if pd.notna(val_medicos) and str(val_medicos).strip() != '' and str(val_medicos).strip().upper() != 'NAN':
                        medicos_norm = normalize_name(val_medicos)
                        if medicos_norm in known_supers_norm:
                            for name in known_supers:
                                if normalize_name(name) == medicos_norm:
                                        return name
                            return str(val_medicos).strip().upper()
                return val_super
            df[col_name] = df.apply(fill_missing_super_validated, axis=1)

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

    # 9. Extraer columna de restricciones / licencias (Hora Fin Restriccion)
    col_restr = None
    for col in df.columns:
        col_norm = str(col).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
        if any(k in col_norm for k in ['HORA FIN RESTRICCION', 'RESTRICCION', 'RESTRICION', 'FIN RESTRICCION']):
            col_restr = col
            break

    if col_restr:
        df['RESTRICCION'] = df[col_restr].fillna('').astype(str).str.strip()
    else:
        df['RESTRICCION'] = ''

    return df


def is_unpaid_leave_or_permission(val):
    """
    Verifica si el valor de 'Hora Fin Restriccion' contiene alguna de las palabras clave:
    'LICENCIA NO REMUNERADA', 'PERMISO', 'LICENCIA' (excluyendo 'LICENCIA REMUNERADA'), 'VACACIONES'.
    """
    if pd.isna(val):
        return False
    val_str = str(val).strip().upper()
    val_str = val_str.replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
    if not val_str or val_str in ['NAN', 'NONE', 'NULL', '']:
        return False
    if 'LICENCIA REMUNERADA' in val_str and 'NO REMUNERADA' not in val_str:
        return False
    keywords = ['LICENCIA NO REMUNERADA', 'PERMISO', 'LICENCIA', 'VACACIONES', 'VACACION']
    return any(kw in val_str for kw in keywords)


def get_restriction_dict(df_ref=None, df_super=None):
    """
    Construye un diccionario {(doc_norm, fecha_date): text_restriccion}
    para rápida búsqueda de licencias/permisos no remunerados por médico y fecha.
    """
    restr_dict = {}

    if df_super is not None and not df_super.empty and 'RESTRICCION' in df_super.columns:
        for _, r in df_super.iterrows():
            f_clean = r.get('FECHA_CLEAN')
            doc_norm = r.get('NOMBRE_NORM')
            restr_val = r.get('RESTRICCION')
            if pd.notna(f_clean) and pd.notna(doc_norm) and is_unpaid_leave_or_permission(restr_val):
                restr_dict[(doc_norm, f_clean.date())] = str(restr_val).strip()

    if df_ref is not None and not df_ref.empty and 'RESTRICCION' in df_ref.columns:
        for _, r in df_ref.iterrows():
            f_clean = r.get('FECHA_CLEAN')
            doc_name = r.get('NOMBRE SUPER VALIDADO')
            restr_val = r.get('RESTRICCION')
            if pd.notna(f_clean) and pd.notna(doc_name) and is_unpaid_leave_or_permission(restr_val):
                doc_norm = normalize_name(doc_name)
                restr_dict[(doc_norm, f_clean.date())] = str(restr_val).strip()

    return restr_dict


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
    
    # Determinar los meses permitidos según el filtro de Streamlit
    allowed_months = None
    try:
        import streamlit as st
        if 'mes_sel' in st.session_state and st.session_state.mes_sel:
            inv_map = {v.upper(): k for k, v in MESES_MAP.items()}
            allowed_months = [inv_map[m.upper()] for m in st.session_state.mes_sel if m.upper() in inv_map]
    except Exception:
        pass
        
    if allowed_months is None:
        allowed_months = list(MESES_MAP.keys())

    # Obtener médicos activos en el df filtrado
    active_medicos = df[['CEDULA_FINAL', 'NOMBRE SUPER VALIDADO']].drop_duplicates()
    
    active_combos = []
    for _, m_row in active_medicos.iterrows():
        cedula = m_row['CEDULA_FINAL']
        nombre = m_row['NOMBRE SUPER VALIDADO']
        doc_norm = normalize_name(nombre)
        
        meses_activos = set()
        # 1. Meses de df_ref (histórico completo de novedades)
        months_df = df_ref[df_ref['NOMBRE SUPER VALIDADO'] == nombre]['MES_NUM'].dropna().unique()
        meses_activos.update(int(m) for m in months_df)
        
        # 2. Meses de df_super (programación)
        if df_super is not None and not df_super.empty:
            months_super = df_super[df_super['NOMBRE_NORM'] == doc_norm]['MES_NUM'].dropna().unique()
            meses_activos.update(int(m) for m in months_super)
            
        # Filtrar por meses permitidos
        for m_num in sorted(meses_activos):
            if m_num in allowed_months:
                active_combos.append({
                    'CEDULA_FINAL': cedula,
                    'NOMBRE SUPER VALIDADO': nombre,
                    'MES': MESES_MAP.get(m_num, 'Sin Mes'),
                    'MES_NUM': m_num
                })
                
    medicos_meses = pd.DataFrame(active_combos)
    daily_rows = []
    
    restr_dict = get_restriction_dict(df_ref=df_ref, df_super=df_super)

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

                    # Verificar restricciones (Licencia no remunerada, Permiso)
                    doc_norm = normalize_name(doc_name)
                    restr_text = restr_dict.get((doc_norm, curr.date()))

                    if restr_text and is_unpaid_leave_or_permission(restr_text):
                        horas_a_laborar = 0.0
                        estado = restr_text.upper()
                    elif curr.dayofweek == 5 and horas_trabajadas == 0.0:
                        estado = "Descanso"
                    else:
                        estado = ""

                    daily_rows.append({
                        'FECHA_CLEAN': curr,
                        'FECHA_STR': date_str,
                        'CEDULA_FINAL': cedula,
                        'NOMBRE SUPER VALIDADO': doc_name,
                        'HORAS_TOTALES': horas_trabajadas,
                        'CANTIDAD_NOVEDADES': novedades,
                        'MES': month_name,
                        'MES_NUM': month_num,
                        'HORAS_A_LABORAR': horas_a_laborar,
                        'ESTADO': estado
                    })
                curr += pd.Timedelta(days=1)

    df_res = pd.DataFrame(daily_rows)
    if not df_res.empty:
        # Filtrar días no laborados sin horas a laborar (ej. domingos/festivos sin novedades),
        # pero conservar aquellos con estado específico (ej. Descanso, Licencia No Remunerada, Permiso).
        df_res = df_res[~((df_res['HORAS_TOTALES'] == 0) & (df_res['HORAS_A_LABORAR'] == 0) & (df_res['ESTADO'] == ''))]
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
                'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'HORAS_A_LABORAR', 'ESTADO'
            ])
        df_daily = df_daily.sort_values(by=['FECHA_CLEAN', 'NOMBRE SUPER VALIDADO'], ascending=[False, True]).reset_index(drop=True)
        return df_daily[['FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO', 'HORAS_TOTALES', 'CANTIDAD_NOVEDADES', 'HORAS_A_LABORAR', 'ESTADO']]

    df_copy = df.copy()
    df_copy['FECHA_STR'] = df_copy['FECHA_CLEAN'].dt.strftime('%Y-%m-%d')
    df_copy['FECHA_STR'] = df_copy['FECHA_STR'].fillna('Sin Fecha')
    grouped = df_copy.groupby(['FECHA_STR', 'CEDULA_FINAL', 'NOMBRE SUPER VALIDADO'], as_index=False).agg(
        HORAS_TOTALES=('HORAS TOTALES DECIMAL', 'sum'),
        CANTIDAD_NOVEDADES=('HORAS TOTALES DECIMAL', 'count')
    )
    grouped['ESTADO'] = ""
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
    if isinstance(file_source, pd.ExcelFile):
        safe_source = file_source
        cleanup = lambda: None
    else:
        safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            # Intentar leer la pestaña FESTIVOS
            if isinstance(safe_source, pd.ExcelFile):
                if 'FESTIVOS' in safe_source.sheet_names:
                    fest_df = pd.read_excel(safe_source, sheet_name='FESTIVOS')
                else:
                    fest_df = pd.DataFrame(columns=['FECHA'])
            elif isinstance(safe_source, str) and safe_source.endswith(('.csv', '.CSV')):
                fest_df = pd.DataFrame(columns=['FECHA'])
            else:
                fest_df = pd.read_excel(safe_source, sheet_name='FESTIVOS', engine='calamine')
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
        
    # Asegurar que el 13 de julio de 2026 sea considerado festivo (en Colombia fue festivo)
    julio_13 = pd.Timestamp('2026-07-13').date()
    if julio_13 not in festivos:
        festivos.append(julio_13)
        
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
            lambda r: format_semana_str(r['SEMANA_INICIO'], r['SEMANA_FIN']),
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
        lambda r: format_semana_str(r['SEMANA_INICIO'], r['SEMANA_FIN']),
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
    SUPERNUMERARIOS FEBRERO, etc.) o una hoja consolidada única "SUPERNUMERARIOS",
    y las combina en un solo DataFrame.
    Cada hoja tiene columnas: Cedula, Nombre, Fecha, Zona, Cis.
    Retorna un DataFrame con columnas [FECHA_CLEAN, NOMBRE_NORM, MES_NUM] o None.
    """
    if isinstance(file_source, pd.ExcelFile):
        safe_source = file_source
        cleanup = lambda: None
    else:
        safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            if isinstance(safe_source, pd.ExcelFile):
                xl = safe_source
            else:
                xl = pd.ExcelFile(safe_source, engine='calamine')
            sheet_names = xl.sheet_names

            # Mapeo de nombre de mes en español a número
            mes_a_num = {
                'ENERO': 1, 'FEBRERO': 2, 'MARZO': 3, 'ABRIL': 4,
                'MAYO': 5, 'JUNIO': 6, 'JULIO': 7, 'AGOSTO': 8,
                'SEPTIEMBRE': 9, 'OCTUBRE': 10, 'NOVIEMBRE': 11, 'DICIEMBRE': 12
            }

            all_dfs = []

            # 1. Buscar si existe una hoja única consolidada llamada "SUPERNUMERARIOS" o "SUPERNUMERARIO_2026_HISTORICO"
            super_single_sheets = [s for s in sheet_names if s.strip().upper() in ['SUPERNUMERARIOS', 'SUPERNUMERARIO_2026_HISTORICO']]
            for sheet in super_single_sheets:
                df = pd.read_excel(xl, sheet_name=sheet)
                df.columns = [str(col).strip() for col in df.columns]

                # Buscar columnas Fecha, Nombre y Restricción
                col_fecha = None
                col_nombre = None
                col_restr = None
                for col in df.columns:
                    col_norm = str(col).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
                    if col_norm == 'FECHA':
                        col_fecha = col
                    elif col_norm in ['NOMBRE', 'NOMBRE SUPERNUMERARIO', 'MEDICO']:
                        col_nombre = col
                    elif any(k in col_norm for k in ['HORA FIN RESTRICCION', 'RESTRICCION', 'RESTRICION', 'FIN RESTRICCION']):
                        col_restr = col

                if col_fecha and col_nombre:
                    df['FECHA_CLEAN'] = pd.to_datetime(df[col_fecha], errors='coerce')
                    df['NOMBRE_NORM'] = df[col_nombre].apply(normalize_name)
                    df['MES_NUM'] = df['FECHA_CLEAN'].dt.month
                    if col_restr:
                        df['RESTRICCION'] = df[col_restr].fillna('').astype(str).str.strip()
                    else:
                        df['RESTRICCION'] = ''
                    all_dfs.append(df[['FECHA_CLEAN', 'NOMBRE_NORM', 'MES_NUM', 'RESTRICCION']].dropna(subset=['FECHA_CLEAN', 'NOMBRE_NORM']))

            # 2. Buscar hojas por mes individual (SUPERNUMERARIOS {MES})
            for sheet in sheet_names:
                sheet_upper = sheet.strip().upper()
                if not sheet_upper.startswith('SUPERNUMERARIOS '):
                    continue
                mes_name = sheet_upper.replace('SUPERNUMERARIOS ', '').strip()
                month_num = mes_a_num.get(mes_name)
                if month_num is None:
                    continue

                df = pd.read_excel(xl, sheet_name=sheet)
                df.columns = [str(col).strip() for col in df.columns]

                # Buscar columnas Fecha, Nombre y Restricción
                col_fecha = None
                col_nombre = None
                col_restr = None
                for col in df.columns:
                    col_norm = str(col).strip().upper().replace('Á', 'A').replace('É', 'E').replace('Í', 'I').replace('Ó', 'O').replace('Ú', 'U')
                    if col_norm == 'FECHA':
                        col_fecha = col
                    elif col_norm in ['NOMBRE', 'NOMBRE SUPERNUMERARIO', 'MEDICO']:
                        col_nombre = col
                    elif any(k in col_norm for k in ['HORA FIN RESTRICCION', 'RESTRICCION', 'RESTRICION', 'FIN RESTRICCION']):
                        col_restr = col

                if col_fecha and col_nombre:
                    df['FECHA_CLEAN'] = pd.to_datetime(df[col_fecha], errors='coerce')
                    df['NOMBRE_NORM'] = df[col_nombre].apply(normalize_name)
                    df['MES_NUM'] = month_num
                    if col_restr:
                        df['RESTRICCION'] = df[col_restr].fillna('').astype(str).str.strip()
                    else:
                        df['RESTRICCION'] = ''
                    all_dfs.append(df[['FECHA_CLEAN', 'NOMBRE_NORM', 'MES_NUM', 'RESTRICCION']].dropna(subset=['FECHA_CLEAN', 'NOMBRE_NORM']))

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
    if isinstance(file_source, pd.ExcelFile):
        safe_source = file_source
        cleanup = lambda: None
    else:
        safe_source, cleanup = get_safe_file_source(file_source)
    try:
        try:
            if isinstance(safe_source, pd.ExcelFile):
                xl = safe_source
            else:
                xl = pd.ExcelFile(safe_source, engine='calamine')
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


