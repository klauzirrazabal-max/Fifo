# -*- coding: utf-8 -*-
"""
Script para verificación y análisis FIFO de inventario (FIJA y MOVIL)
Segunda parte del proceso de automatización
Usa Excel COM para preservar valores y crear tabla dinámica correcta

USO:
    python verificar_fifo.py
    python verificar_fifo.py "Fifo.zip" "DICIEMBRE" "FIJA"
    python verificar_fifo.py "Fifo.zip" "DICIEMBRE" "MOVIL"
"""

import win32com.client as win32
from datetime import datetime, timedelta
from collections import defaultdict
import os
import sys
import zipfile
import shutil

# Directorio de trabajo
DIRECTORIO = r"C:\FIFO"

# Diccionario de meses en español
MESES_ESP = {
    1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
    5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
    9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
}

MESES_ESP_INV = {v: k for k, v in MESES_ESP.items()}

# Configuración por tipo de archivo
CONFIG_TIPO = {
    'FIJA': {
        'carpeta_zip': 'Fija',
        'filtro_ce': 'P008',
        'filtro_alm': 'P000',
        'grupos_filtrar': ['ALMACENES'],  # Grupos cuyas series se validan
        'archivo_procesado': 'FIFO_PROCESADO_FIJA.xlsb',
        'archivo_verificar': 'VERIFICAR SALIDA CDVES FIJA {MES}.xlsx',
        'archivo_salida': 'Fifo_Fija_{MES}.xlsb',
        'hoja_series': 'FIJA'
    },
    'MOVIL': {
        'carpeta_zip': 'Movil',
        'filtro_ce': 'P008',
        'filtro_alm': 'H000',
        'grupos_filtrar': ['PDV', 'Televentas VES', 'Tienda Virtual VES', 'Corporativo Ves'],
        'archivo_procesado': 'FIFO_PROCESADO_MOVIL.xlsb',
        'archivo_verificar': 'VERIFICAR SALIDA CDVES MOVIL {MES}.xlsx',
        'archivo_salida': 'Fifo_Movil_{MES}.xlsb',
        'hoja_series': 'MOVIL'
    }
}


def normalizar_serie_para_comparacion(serie):
    """
    Normaliza una serie para comparación:
    - Convierte a string
    - Quita sufijo .0 (de números Excel)
    - Quita ceros iniciales
    """
    if serie is None or serie == '':
        return ''
    serie_str = str(serie).strip()

    # Quitar sufijo .0 de números flotantes
    if serie_str.endswith('.0'):
        serie_str = serie_str[:-2]

    # Quitar ceros iniciales
    return serie_str.lstrip('0') or '0'


def quitar_ceros_iniciales(serie):
    """Alias para compatibilidad - usa normalizar_serie_para_comparacion"""
    return normalizar_serie_para_comparacion(serie)


def parsear_fecha_contable(fecha_str):
    """Parsea fecha en formato dd.MM.yyyy y retorna (año, mes)"""
    if fecha_str is None or fecha_str == '':
        return None, None
    try:
        fecha_str = str(fecha_str).strip()
        partes = fecha_str.split('.')
        if len(partes) == 3:
            dia, mes, año = int(partes[0]), int(partes[1]), int(partes[2])
            return año, mes
    except:
        pass
    return None, None


def detectar_tipo_archivo(zip_path):
    """Detecta si el ZIP contiene datos FIJA o MOVIL"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as z:
            nombres = z.namelist()
            tiene_fija = any('Fija/' in n or 'Fija\\' in n or n.startswith('Fija') for n in nombres)
            tiene_movil = any('Movil/' in n or 'Movil\\' in n or n.startswith('Movil') for n in nombres)

            if tiene_fija and not tiene_movil:
                return 'FIJA'
            elif tiene_movil and not tiene_fija:
                return 'MOVIL'
            elif tiene_fija and tiene_movil:
                return None  # Ambos, necesita selección manual
    except:
        pass
    return None


def obtener_parametros():
    """Obtiene parámetros del script"""
    fecha_actual = datetime.now()
    primer_dia = fecha_actual.replace(day=1)
    ultimo_dia_mes_ant = primer_dia - timedelta(days=1)
    mes_num = ultimo_dia_mes_ant.month
    año_num = ultimo_dia_mes_ant.year
    nombre_mes = MESES_ESP[mes_num]
    tipo_archivo = None

    if len(sys.argv) > 3:
        zip_file = sys.argv[1]
        nombre_mes = sys.argv[2].upper()
        tipo_archivo = sys.argv[3].upper()
        if nombre_mes in MESES_ESP_INV:
            mes_num = MESES_ESP_INV[nombre_mes]
    elif len(sys.argv) > 2:
        zip_file = sys.argv[1]
        nombre_mes = sys.argv[2].upper()
        if nombre_mes in MESES_ESP_INV:
            mes_num = MESES_ESP_INV[nombre_mes]
    elif len(sys.argv) > 1:
        zip_file = sys.argv[1]
    else:
        zips = [f for f in os.listdir(DIRECTORIO) if f.lower().endswith('.zip')]
        if zips:
            print("\nArchivos ZIP disponibles:")
            for i, z in enumerate(zips, 1):
                print(f"  {i}. {z}")
            seleccion = input("Seleccione ZIP (número o nombre): ").strip()
            if seleccion.isdigit():
                zip_file = zips[int(seleccion) - 1]
            else:
                zip_file = seleccion
        else:
            print("ERROR: No se encontraron archivos ZIP")
            sys.exit(1)

    # Detectar tipo si no se especificó
    if tipo_archivo is None:
        zip_path = os.path.join(DIRECTORIO, zip_file)
        tipo_archivo = detectar_tipo_archivo(zip_path)

        if tipo_archivo is None:
            print("\nNo se pudo detectar el tipo automáticamente.")
            print("Seleccione el tipo:")
            print("  1. FIJA")
            print("  2. MOVIL")
            opcion = input("Opción: ").strip()
            tipo_archivo = 'FIJA' if opcion == '1' else 'MOVIL'

    return zip_file, nombre_mes, mes_num, año_num, tipo_archivo


def extraer_zip(zip_path, destino):
    """Extrae el contenido del ZIP"""
    print(f"    Extrayendo {zip_path}...")
    with zipfile.ZipFile(zip_path, 'r') as z:
        z.extractall(destino)
    return destino


def leer_archivos_movimientos(carpeta_base, config, mes_filtro, año_filtro):
    """
    Lee todos los archivos .XLS (TSV) de la carpeta correspondiente
    y retorna un set de series que cumplen los filtros
    """
    series_encontradas = set()
    carpeta_nombre = config['carpeta_zip']
    filtro_ce = config['filtro_ce']
    filtro_alm = config['filtro_alm']

    ruta_carpeta = os.path.join(carpeta_base, carpeta_nombre)
    if not os.path.exists(ruta_carpeta):
        print(f"    Carpeta {carpeta_nombre} no encontrada, saltando...")
        return series_encontradas

    archivos = [f for f in os.listdir(ruta_carpeta) if f.upper().endswith('.XLS')]
    print(f"    Procesando {len(archivos)} archivos en {carpeta_nombre}/...")
    print(f"    Filtros: Ce.={filtro_ce}, Alm.={filtro_alm}, Fecha={mes_filtro}/{año_filtro}")

    for archivo in archivos:
        ruta_archivo = os.path.join(ruta_carpeta, archivo)
        try:
            with open(ruta_archivo, 'r', encoding='utf-16') as f:
                lineas = f.readlines()

            header_idx = None
            for i, linea in enumerate(lineas):
                if 'Ce.' in linea and 'Alm.' in linea and 'serie' in linea.lower():
                    header_idx = i
                    break

            if header_idx is None:
                continue

            headers = lineas[header_idx].strip().split('\t')

            col_ce = col_alm = col_serie = col_fecha = None
            for j, h in enumerate(headers):
                h_lower = h.lower().strip()
                if h_lower == 'ce.':
                    col_ce = j
                elif h_lower == 'alm.':
                    col_alm = j
                elif 'serie' in h_lower:
                    col_serie = j
                elif 'fe.contab' in h_lower:
                    col_fecha = j

            if None in [col_ce, col_alm, col_serie, col_fecha]:
                continue

            for i in range(header_idx + 1, len(lineas)):
                linea = lineas[i].strip()
                if not linea:
                    continue

                campos = linea.split('\t')
                if len(campos) <= max(col_ce, col_alm, col_serie, col_fecha):
                    continue

                ce = campos[col_ce].strip()
                alm = campos[col_alm].strip()
                serie = campos[col_serie].strip()
                fecha = campos[col_fecha].strip()

                if ce != filtro_ce or alm != filtro_alm:
                    continue

                año, mes = parsear_fecha_contable(fecha)
                if año == año_filtro and mes == mes_filtro:
                    serie_norm = quitar_ceros_iniciales(serie)
                    if serie_norm:
                        series_encontradas.add(serie_norm)

        except Exception as e:
            print(f"      Error leyendo {archivo}: {e}")
            continue

    return series_encontradas


def procesar_reporte(excel, archivo_procesado, archivo_salida, series_no_encontradas,
                     grupos_filtrar, tipo_archivo, batch_size=50000, label=""):
    """
    Procesa FIFO_PROCESADO, filtra series no encontradas, crea tabla dinámica y análisis FIFO.
    Genera el archivo de salida final (.xlsb).

    Retorna: registros_finales (int)
    """
    print(f"\n--- Procesando{label} {os.path.basename(archivo_procesado)} ---")
    wb_proc = excel.Workbooks.Open(archivo_procesado)
    ws_proc = wb_proc.ActiveSheet

    ultima_fila_proc = ws_proc.UsedRange.Rows.Count
    ultima_col_proc = ws_proc.UsedRange.Columns.Count
    print(f"    Registros iniciales: {ultima_fila_proc - 1:,}")

    # Leer encabezados
    headers_range = ws_proc.Range(ws_proc.Cells(1, 1), ws_proc.Cells(1, ultima_col_proc))
    headers_vals = headers_range.Value[0]

    # Encontrar índices de columnas (0-indexed para Python)
    idx_serie = idx_almacen_agrupado = None
    nombre_col_serie = None
    for i, h in enumerate(headers_vals):
        if h in ('Serie', 'Número de serie'):
            idx_serie = i
            nombre_col_serie = h
        elif h == 'Almacen Agrupado':
            idx_almacen_agrupado = i

    print(f"    Grupos a filtrar: {grupos_filtrar}")

    # Leer y filtrar datos por lotes
    print(f"    Leyendo y filtrando por lotes...")
    filas_salida = []
    eliminados = 0

    fila = 2
    while fila <= ultima_fila_proc:
        fila_fin = min(fila + batch_size - 1, ultima_fila_proc)
        print(f"    Lote: filas {fila:,} - {fila_fin:,} de {ultima_fila_proc:,}")

        data_range = ws_proc.Range(ws_proc.Cells(fila, 1), ws_proc.Cells(fila_fin, ultima_col_proc))
        batch_data = data_range.Value

        if fila_fin == fila:
            batch_data = [batch_data]

        for row in batch_data:
            serie = row[idx_serie]
            almacen_agrupado = row[idx_almacen_agrupado]

            if almacen_agrupado in grupos_filtrar:
                serie_norm = quitar_ceros_iniciales(serie)
                if serie_norm in series_no_encontradas:
                    eliminados += 1
                    continue

            filas_salida.append(list(row))

        fila = fila_fin + 1

    wb_proc.Close(SaveChanges=False)

    registros_finales = len(filas_salida)
    print(f"    Registros eliminados: {eliminados:,}")
    print(f"    Registros finales: {registros_finales:,}")

    # Crear nuevo workbook para salida
    print(f"    Creando archivo de salida...")
    wb_salida = excel.Workbooks.Add()
    ws_data = wb_salida.ActiveSheet
    ws_data.Name = "Data"

    for i, h in enumerate(headers_vals, 1):
        ws_data.Cells(1, i).Value = h

    print(f"    Escribiendo datos...")
    fila_salida = 2
    for i in range(0, registros_finales, batch_size):
        fin = min(i + batch_size, registros_finales)
        print(f"      Escribiendo registros {i+1:,} - {fin:,}")

        batch = filas_salida[i:fin]
        rango = ws_data.Range(
            ws_data.Cells(fila_salida, 1),
            ws_data.Cells(fila_salida + len(batch) - 1, ultima_col_proc)
        )
        rango.Value = batch
        fila_salida += len(batch)

    # Crear tabla dinámica
    print(f"    Creando tabla dinámica...")

    ws_td = wb_salida.Worksheets.Add(After=wb_salida.Worksheets(wb_salida.Worksheets.Count))
    ws_td.Name = "TD"

    rango_datos = ws_data.Range(ws_data.Cells(1, 1), ws_data.Cells(fila_salida - 1, ultima_col_proc))

    pivot_cache = wb_salida.PivotCaches().Create(
        SourceType=1,
        SourceData=rango_datos
    )

    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=ws_td.Range("A3"),
        TableName="TablaDinamicaFIFO"
    )

    pf_material = pivot_table.PivotFields("Material")
    pf_material.Orientation = 1
    pf_material.Position = 1

    pf_desc = pivot_table.PivotFields("Descripción Material SAP")
    pf_desc.Orientation = 1
    pf_desc.Position = 2

    pf_mes = pivot_table.PivotFields("Mes/Año")
    pf_mes.Orientation = 1
    pf_mes.Position = 3

    pf_almacen = pivot_table.PivotFields("Almacen Agrupado")
    pf_almacen.Orientation = 2
    pf_almacen.Position = 1

    try:
        pi_cdves = pf_almacen.PivotItems("CD VES")
        pi_cdves.Position = 1
    except:
        pass

    pf_serie = pivot_table.PivotFields(nombre_col_serie)
    pivot_table.AddDataField(pf_serie, "Cuenta de Serie", -4112)

    pivot_table.RowAxisLayout(1)
    pivot_table.RepeatAllLabels(1)

    pf_material.Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)
    pf_desc.Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)
    pf_mes.Subtotals = (False, False, False, False, False, False, False, False, False, False, False, False)

    pivot_table.RefreshTable()

    pivot_range = pivot_table.TableRange2
    col_inicio_td = pivot_range.Column
    fila_inicio_td = pivot_range.Row
    num_filas_td = pivot_range.Rows.Count
    num_cols_td = pivot_range.Columns.Count

    for col in range(col_inicio_td, col_inicio_td + num_cols_td):
        header = ws_td.Cells(fila_inicio_td, col).Value
        if header is not None and "Mes" in str(header):
            rango_mes = ws_td.Range(
                ws_td.Cells(fila_inicio_td + 1, col),
                ws_td.Cells(fila_inicio_td + num_filas_td - 1, col)
            )
            rango_mes.NumberFormat = "mmm-yy"
            break

    col_no_desp = col_inicio_td + num_cols_td
    col_no_fifo = col_no_desp + 1

    ws_td.Cells(fila_inicio_td, col_no_desp).Value = "NO DESPACHÓ"
    ws_td.Cells(fila_inicio_td, col_no_fifo).Value = "NO CUMPLIÓ FIFO"

    # === ANÁLISIS FIFO ===
    print(f"    Ejecutando análisis FIFO...")

    col_material = col_mes = col_cdves = None
    cols_grupos = {}
    fila_headers = None

    for fila_busqueda in range(fila_inicio_td, min(fila_inicio_td + 3, fila_inicio_td + num_filas_td)):
        for col in range(col_inicio_td, col_inicio_td + num_cols_td):
            header = ws_td.Cells(fila_busqueda, col).Value
            if header is None:
                continue
            header_str = str(header).strip()
            if header_str == "Material":
                col_material = col
                fila_headers = fila_busqueda
            elif "Mes" in header_str:
                col_mes = col
            elif header_str == "CD VES":
                col_cdves = col
            for grupo in grupos_filtrar:
                if header_str == grupo:
                    cols_grupos[grupo] = col

    print(f"      Columnas detectadas: Material={col_material}, Mes={col_mes}, CD VES={col_cdves}")
    print(f"      Grupos detectados: {cols_grupos}")

    if fila_headers:
        fila_inicio_td = fila_headers

    cumplimiento = 0.0
    total_despacho = 0
    total_no_despacho = 0

    if col_material and col_mes and col_cdves and cols_grupos:
        datos_td = []
        ultimo_material = None

        for r in range(fila_inicio_td + 1, fila_inicio_td + num_filas_td):
            try:
                material_val = ws_td.Cells(r, col_material).Value
                mes_año = ws_td.Cells(r, col_mes).Value
                cdves_val = ws_td.Cells(r, col_cdves).Value

                grupos_val = 0
                for grupo, col_grupo in cols_grupos.items():
                    val = ws_td.Cells(r, col_grupo).Value
                    if val:
                        grupos_val += int(val)
            except:
                continue

            if material_val is not None and str(material_val).strip() != "":
                if "Total" in str(material_val):
                    ultimo_material = None
                    continue
                ultimo_material = str(material_val).strip()
                material = ultimo_material
            elif ultimo_material is not None and mes_año is not None:
                material = ultimo_material
            else:
                continue

            if mes_año is None:
                continue

            datos_td.append({
                'fila': r,
                'material': material,
                'mes_año': mes_año,
                'cdves': int(cdves_val) if cdves_val else 0,
                'grupos': grupos_val,
                'no_desp': 0,
                'no_fifo': 0
            })

        print(f"      Registros TD: {len(datos_td):,}")

        materiales = defaultdict(list)
        for d in datos_td:
            materiales[d['material']].append(d)

        def convertir_fecha(fecha_val):
            if fecha_val is None:
                return datetime(1900, 1, 1)
            if isinstance(fecha_val, datetime):
                return fecha_val
            if isinstance(fecha_val, (int, float)):
                try:
                    return datetime(1899, 12, 30) + timedelta(days=int(fecha_val))
                except:
                    return datetime(1900, 1, 1)
            if isinstance(fecha_val, str):
                try:
                    partes = fecha_val.split('/')
                    if len(partes) == 3:
                        return datetime(int(partes[2]), int(partes[1]), int(partes[0]))
                except:
                    pass
            return datetime(1900, 1, 1)

        for material, filas in materiales.items():
            filas.sort(key=lambda x: convertir_fecha(x['mes_año']))

            n = len(filas)
            arr_cdves = [f['cdves'] for f in filas]
            arr_grupos = [f['grupos'] for f in filas]
            arr_remain = arr_cdves.copy()
            arr_no_desp = [0] * n
            arr_no_fifo = [0] * n

            for k in range(n):
                total_despacho += arr_grupos[k]

                if arr_grupos[k] > 0:
                    prior_sum = sum(arr_remain[:k])
                    if prior_sum > 0:
                        violation = min(prior_sum, arr_grupos[k])
                        arr_no_fifo[k] = violation

                        need = violation
                        for j in range(k):
                            if arr_remain[j] > 0 and need > 0:
                                take = min(arr_remain[j], need)
                                arr_no_desp[j] += take
                                arr_remain[j] -= take
                                need -= take

            for i, f in enumerate(filas):
                f['no_desp'] = arr_no_desp[i]
                f['no_fifo'] = arr_no_fifo[i]
                total_no_despacho += arr_no_desp[i]

        for d in datos_td:
            if d['no_desp'] > 0:
                ws_td.Cells(d['fila'], col_no_desp).Value = d['no_desp']
            if d['no_fifo'] > 0:
                ws_td.Cells(d['fila'], col_no_fifo).Value = d['no_fifo']

        col_resumen = col_no_fifo + 2
        ws_td.Cells(1, col_resumen).Value = f"RESUMEN FIFO {tipo_archivo}{label}"
        ws_td.Cells(1, col_resumen).Font.Bold = True

        ws_td.Cells(2, col_resumen).Value = "Total Despacho:"
        ws_td.Cells(2, col_resumen + 1).Value = total_despacho
        ws_td.Cells(2, col_resumen).Font.Bold = True

        ws_td.Cells(3, col_resumen).Value = "Total No Despachó:"
        ws_td.Cells(3, col_resumen + 1).Value = total_no_despacho
        ws_td.Cells(3, col_resumen).Font.Bold = True

        if total_despacho > 0:
            cumplimiento = ((total_despacho - total_no_despacho) / total_despacho) * 100
            ws_td.Cells(4, col_resumen).Value = "Cumplimiento FIFO:"
            ws_td.Cells(4, col_resumen + 1).Value = f"{cumplimiento:.2f}%"
            ws_td.Cells(4, col_resumen).Font.Bold = True

        print(f"    Análisis FIFO completado:")
        print(f"      - Total Despacho: {total_despacho:,}")
        print(f"      - Total No Despachó: {total_no_despacho:,}")
        if total_despacho > 0:
            print(f"      - Cumplimiento FIFO: {cumplimiento:.2f}%")
    else:
        print(f"    ADVERTENCIA: No se encontraron todas las columnas para análisis FIFO")

    wb_salida.SaveAs(archivo_salida, FileFormat=50)
    print(f"    Archivo guardado: {archivo_salida}")
    wb_salida.Close(SaveChanges=False)

    return registros_finales


def main():
    print("=" * 60)
    print("=== VERIFICACIÓN Y ANÁLISIS FIFO (Excel COM) ===")
    print("=" * 60)

    # Obtener parámetros
    zip_file, nombre_mes, mes_num, año_num, tipo_archivo = obtener_parametros()
    config = CONFIG_TIPO[tipo_archivo]

    zip_path = os.path.join(DIRECTORIO, zip_file)
    if not os.path.exists(zip_path):
        print(f"ERROR: No se encontró el archivo: {zip_path}")
        sys.exit(1)

    print(f"\nParámetros:")
    print(f"  - Archivo ZIP: {zip_file}")
    print(f"  - Tipo: {tipo_archivo}")
    print(f"  - Mes a evaluar: {nombre_mes} {año_num}")

    archivo_verificar = os.path.join(DIRECTORIO, config['archivo_verificar'].format(MES=nombre_mes))
    archivo_procesado = os.path.join(DIRECTORIO, config['archivo_procesado'])
    archivo_salida = os.path.join(DIRECTORIO, config['archivo_salida'].format(MES=nombre_mes))

    if not os.path.exists(archivo_verificar):
        print(f"ERROR: No se encontró: {archivo_verificar}")
        sys.exit(1)

    if not os.path.exists(archivo_procesado):
        print(f"ERROR: No se encontró: {archivo_procesado}")
        sys.exit(1)

    # Paso 1: Extraer ZIP
    print(f"\n[1/6] Extrayendo ZIP...")
    carpeta_temp = os.path.join(DIRECTORIO, "temp_movimientos")
    if os.path.exists(carpeta_temp):
        shutil.rmtree(carpeta_temp)
    extraer_zip(zip_path, carpeta_temp)

    # Paso 2: Leer series a verificar usando Excel COM
    print(f"\n[2/6] Leyendo series a verificar...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    BATCH_SIZE = 50000

    try:
        wb_verif = excel.Workbooks.Open(archivo_verificar)
        ws_verif = wb_verif.Worksheets(config['hoja_series'])
        ultima_fila_verif = ws_verif.UsedRange.Rows.Count

        series_verificar = set()
        # Leer series por lotes
        fila = 2
        while fila <= ultima_fila_verif:
            fila_fin = min(fila + BATCH_SIZE - 1, ultima_fila_verif)
            rango = ws_verif.Range(ws_verif.Cells(fila, 1), ws_verif.Cells(fila_fin, 1))
            valores = rango.Value

            if valores:
                if isinstance(valores, tuple):
                    for v in valores:
                        if v and v[0]:
                            serie_norm = quitar_ceros_iniciales(v[0])
                            if serie_norm:
                                series_verificar.add(serie_norm)
                else:
                    serie_norm = quitar_ceros_iniciales(valores)
                    if serie_norm:
                        series_verificar.add(serie_norm)

            fila = fila_fin + 1

        wb_verif.Close(SaveChanges=False)
        print(f"    Series a verificar: {len(series_verificar):,}")

        # Paso 3: Buscar series en archivos de movimientos
        print(f"\n[3/6] Buscando series en archivos de movimientos...")
        series_encontradas = leer_archivos_movimientos(carpeta_temp, config, mes_num, año_num)
        print(f"    Series encontradas en movimientos: {len(series_encontradas):,}")

        # Paso 4: Identificar series no encontradas
        print(f"\n[4/6] Identificando series no encontradas...")
        series_no_encontradas = series_verificar - series_encontradas
        print(f"    Series NO encontradas: {len(series_no_encontradas):,}")

        # Paso 5-6: Procesar reporte original
        grupos_filtrar = config['grupos_filtrar']
        print(f"\n[5/6] Procesando reporte original...")
        registros_finales = procesar_reporte(
            excel, archivo_procesado, archivo_salida, series_no_encontradas,
            grupos_filtrar, tipo_archivo, BATCH_SIZE
        )

        # Paso 7: Procesar reporte alternativo (ALT) si existe
        archivo_procesado_alt = archivo_procesado.replace('.xlsb', '_ALT.xlsb')
        archivo_salida_alt = archivo_salida.replace('.xlsb', '_ALT.xlsb')

        registros_finales_alt = None
        if os.path.exists(archivo_procesado_alt):
            print(f"\n[6/6] Procesando reporte alternativo (ALT)...")
            print(f"    CD VES: Mes/Año basado en Fecha Modificación")
            registros_finales_alt = procesar_reporte(
                excel, archivo_procesado_alt, archivo_salida_alt, series_no_encontradas,
                grupos_filtrar, tipo_archivo, BATCH_SIZE, label=" ALT"
            )
        else:
            print(f"\n[6/6] No se encontró archivo ALT ({os.path.basename(archivo_procesado_alt)}), omitiendo reporte alternativo.")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            excel.Quit()
        except:
            pass

        # Limpiar carpeta temporal
        if os.path.exists(carpeta_temp):
            shutil.rmtree(carpeta_temp)

    print(f"\n{'=' * 60}")
    print(f"=== PROCESO COMPLETADO ({tipo_archivo}) ===")
    print(f"{'=' * 60}")
    print(f"\nArchivo generado: {archivo_salida}")
    print(f"  - Hoja 'Data': {registros_finales:,} registros")
    print(f"  - Hoja 'TD': Tabla dinámica con análisis FIFO")
    if registros_finales_alt is not None:
        print(f"\nArchivo alternativo: {archivo_salida_alt}")
        print(f"  - Hoja 'Data': {registros_finales_alt:,} registros")
        print(f"  - Hoja 'TD': Tabla dinámica con análisis FIFO (Fecha Modificación)")


if __name__ == "__main__":
    main()
