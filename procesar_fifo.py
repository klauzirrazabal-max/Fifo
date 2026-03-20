# -*- coding: utf-8 -*-
"""
Script para procesamiento FIFO de inventario (FIJA y MOVIL)
Soporta dos formatos de entrada:
  - ZIP con TXT pipe-delimited (formato nuevo)
  - XLSB via Excel COM (formato legado)

USO:
    python procesar_fifo.py archivos_base.zip FIJA
    python procesar_fifo.py archivos_base.zip MOVIL
    python procesar_fifo.py "Base_Fija_enero.xlsb"
    python procesar_fifo.py  (solicita el nombre del archivo)
"""

import win32com.client as win32
from datetime import datetime, timedelta
import os
import sys
import zipfile
from functools import partial

# Directorio de trabajo
DIRECTORIO = r"C:\FIFO"
DIRECTORIO_ENTRADA = os.path.join(DIRECTORIO, "archivos_base")

# Diccionario de meses en español
MESES_ESP = {
    1: 'ENERO', 2: 'FEBRERO', 3: 'MARZO', 4: 'ABRIL',
    5: 'MAYO', 6: 'JUNIO', 7: 'JULIO', 8: 'AGOSTO',
    9: 'SEPTIEMBRE', 10: 'OCTUBRE', 11: 'NOVIEMBRE', 12: 'DICIEMBRE'
}

# ===== CONSTANTES FIJA =====
EXCLUIDOS_FIJA = [
    'CD VES',
    'BAJAS DE INVENTARIO',
    'BODEGA FALTANTE DE INVENTARIO',
    'LOGISTICA INVERSA INFRAESTRUCTURA',
    'REFURBISHED'
]

# ===== CONSTANTES MOVIL =====
ALMACENES_TIPO1_PDV = {'CAV', 'CAC', 'OTROS CADS'}
ESTADOS_CADENA = {'ALMA ENCL', 'ENCL ALMA'}

# Columnas que deben formatearse como texto (formato xlsb)
COLUMNAS_TEXTO = ['Serie', 'EQUIPO', 'Material', 'Lote', 'Cliente', 'CEAL Cliente']

# Tamaño del lote para procesamiento
BATCH_SIZE = 50000

# ===== CONSTANTES FORMATO ZIP (TXT pipe-delimited) =====

# Mapeo de nombres: clave interna → nombre de columna en los .txt
TXT_COL_MAP = {
    'almacen_tipo1': 'Tipo 1',
    'centro': 'Centro',
    'almacen': 'Almacén',
    'tipo_stock': 'Tipo Stocks Movimiento mercancía (Status)',
    'fecha_ing': 'Fecha de ingreso',
    'fecha_mod': 'Modificado el',
    'serie': 'Número de serie',
    'marca': 'Marca Claro',
    'antiguedad': 'Antiguedad',
}
# 'estado' → segunda columna "Status" (col 37), manejo especial por duplicado

# Archivos dentro del zip según tipo
ZIP_FILES = {
    'FIJA': 'Fija/Descarga_inventario_Infra_Valorado.txt',
    'MOVIL': 'Movil/Descarga_inventario_MOVIL.txt',
}

# Columnas que deben formatearse como texto (formato txt)
COLUMNAS_TEXTO_TXT = ['Número de serie', 'Número de equipo', 'Material', 'Lote', 'Nº cliente']


def detectar_tipo_archivo(nombre_archivo):
    """Detecta si el archivo es FIJA o MOVIL por el nombre"""
    nombre_upper = nombre_archivo.upper()
    if 'FIJA' in nombre_upper:
        return 'FIJA'
    elif 'MOVIL' in nombre_upper:
        return 'MOVIL'
    return None


def obtener_archivo_entrada():
    """
    Obtiene el nombre del archivo de entrada desde argumentos o input del usuario.

    Formatos soportados:
    - .zip: archivo_base.zip con TXT pipe-delimited (requiere tipo FIJA/MOVIL)
    - .xlsb: archivo Excel binario (formato legado)

    Retorna: (ruta_completa, tipo_archivo, formato)
      formato: 'zip' o 'xlsb'
    """
    if len(sys.argv) > 1:
        nombre_archivo = sys.argv[1]
    else:
        print("\n=== PROCESAMIENTO FIFO ===")
        print(f"Archivos disponibles en {DIRECTORIO_ENTRADA}:")

        if not os.path.exists(DIRECTORIO_ENTRADA):
            print(f"ERROR: No existe el directorio {DIRECTORIO_ENTRADA}")
            sys.exit(1)

        archivos = [f for f in os.listdir(DIRECTORIO_ENTRADA)
                     if f.endswith('.xlsb') or f.endswith('.zip')]
        for i, archivo in enumerate(archivos, 1):
            print(f"  {i}. {archivo}")
        print()
        nombre_archivo = input("Ingrese el nombre del archivo (o número): ").strip()

        if nombre_archivo.isdigit():
            idx = int(nombre_archivo) - 1
            if 0 <= idx < len(archivos):
                nombre_archivo = archivos[idx]

    # Determinar formato
    if nombre_archivo.lower().endswith('.zip'):
        formato = 'zip'
        # Para zip, buscar primero en DIRECTORIO, luego en DIRECTORIO_ENTRADA
        if os.path.isabs(nombre_archivo) and os.path.exists(nombre_archivo):
            ruta_completa = nombre_archivo
        elif os.path.exists(os.path.join(DIRECTORIO, nombre_archivo)):
            ruta_completa = os.path.join(DIRECTORIO, nombre_archivo)
        elif os.path.exists(os.path.join(DIRECTORIO_ENTRADA, nombre_archivo)):
            ruta_completa = os.path.join(DIRECTORIO_ENTRADA, nombre_archivo)
        else:
            print(f"ERROR: No se encontró el archivo: {nombre_archivo}")
            sys.exit(1)

        # Tipo FIJA/MOVIL desde segundo argumento o preguntar
        if len(sys.argv) > 2:
            tipo_archivo = sys.argv[2].upper()
        else:
            print("\nSeleccione el tipo de procesamiento:")
            print("  1. FIJA")
            print("  2. MOVIL")
            opcion = input("Opción: ").strip()
            tipo_archivo = 'FIJA' if opcion == '1' else 'MOVIL'

        if tipo_archivo not in ZIP_FILES:
            print(f"ERROR: Tipo '{tipo_archivo}' no válido. Use FIJA o MOVIL.")
            sys.exit(1)

        return ruta_completa, tipo_archivo, formato

    else:
        formato = 'xlsb'
        if not nombre_archivo.endswith('.xlsb'):
            nombre_archivo += '.xlsb'

        ruta_completa = os.path.join(DIRECTORIO_ENTRADA, nombre_archivo)

        if not os.path.exists(ruta_completa):
            print(f"ERROR: No se encontró el archivo: {ruta_completa}")
            sys.exit(1)

        # Detectar tipo de archivo
        tipo_archivo = detectar_tipo_archivo(nombre_archivo)
        if tipo_archivo is None:
            print("\nNo se pudo detectar el tipo de archivo (FIJA/MOVIL)")
            print("Seleccione el tipo:")
            print("  1. FIJA")
            print("  2. MOVIL")
            opcion = input("Opción: ").strip()
            tipo_archivo = 'FIJA' if opcion == '1' else 'MOVIL'

        return ruta_completa, tipo_archivo, formato


def convertir_fecha_mes_año(fecha_valor):
    """Convierte fecha de formato yyyyMMdd a datetime con día 01"""
    if fecha_valor is None or fecha_valor == 0 or str(fecha_valor).strip() in ['', '0', '00000000']:
        return None
    try:
        fecha_str = str(int(float(fecha_valor)))
        if len(fecha_str) == 8:
            año = int(fecha_str[:4])
            mes = int(fecha_str[4:6])
            return datetime(año, mes, 1)
    except:
        pass
    return None


def clasificar_fila_fija(fila, idx, fecha_inicio, fecha_fin, alt_mode=False):
    """
    Aplica reglas de clasificación FIJA

    Retorna: (almacen_agrupado, mes_año) o (None, None) si no aplica

    Reglas:
    - CD VES: Amacén Tipo 1 = 'CD VES', Centro = 'P008', Almacén = 'P000',
              Estado = 'ALMA', Tipo Stock = '01', sin filtro de fecha
    - ALMACENES: Amacén Tipo 1 no en excluidos, Fecha Modificación en mes anterior

    alt_mode: Si True, CD VES usa Fecha Modificación como Mes/Año (si válida), sino Fecha Ingreso
    """
    almacen_tipo1 = fila[idx['almacen_tipo1']]
    centro = fila[idx['centro']]
    almacen_val = fila[idx['almacen']]
    estado = fila[idx['estado']]
    tipo_stock = fila[idx['tipo_stock']]
    fecha_ing = fila[idx['fecha_ing']]
    fecha_mod = fila[idx['fecha_mod']]

    tipo_stock_str = str(tipo_stock).strip() if tipo_stock else ''

    # CD VES
    if (almacen_tipo1 == 'CD VES' and
        centro == 'P008' and
        almacen_val == 'P000' and
        estado == 'ALMA' and
        tipo_stock_str in ['1', '01']):

        if alt_mode:
            mes_año = convertir_fecha_mes_año(fecha_mod)
            if mes_año is None:
                mes_año = convertir_fecha_mes_año(fecha_ing)
        else:
            mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('CD VES', mes_año)

    # ALMACENES (filtrar por Fecha Modificación en mes anterior, Mes/Año de Fecha Ingreso)
    if almacen_tipo1 not in EXCLUIDOS_FIJA:
        try:
            fecha_mod_int = int(float(fecha_mod)) if fecha_mod else 0
        except:
            fecha_mod_int = 0

        if fecha_inicio <= fecha_mod_int <= fecha_fin:
            mes_año = convertir_fecha_mes_año(fecha_ing)
            return ('ALMACENES', mes_año)

    return (None, None)


def clasificar_fila_movil(fila, idx, fecha_inicio, fecha_fin, alt_mode=False):
    """
    Aplica reglas de clasificación MOVIL

    Retorna: (almacen_agrupado, mes_año) o (None, None) si no aplica

    Reglas:
    - CD VES: Amacén Tipo 1 = 'CD VES', Centro = 'P008', Almacén = 'H000',
              Estado = 'ALMA', Tipo Stock = '01', sin filtro de fecha
    - PDV (CADENA): Amacén Tipo 1 = 'CADENA', Estado in {'ALMA ENCL', 'ENCL ALMA'},
                    Tipo Stock = '02', Fecha Modificación en mes anterior
    - PDV (CAV/CAC): Amacén Tipo 1 in {'CAV', 'CAC', 'OTROS CADS'},
                     Almacén in {'H000', 'H001'}, Estado = 'ALMA', Tipo Stock = '01',
                     Fecha Modificación en mes anterior
    - Televentas VES: Centro = 'P150', Almacén = 'H000', Estado = 'ALMA',
                      Tipo Stock in {'01', '07'}, Fecha Modificación en mes anterior
    - Tienda Virtual VES: Centro = 'P102', Almacén = 'H000', Estado = 'ALMA',
                          Tipo Stock in {'01', '07'}, Fecha Modificación en mes anterior
    - Corporativo Ves: Centro = 'P101', Almacén in {'H001', 'H002'}, Estado = 'ALMA',
                       Tipo Stock in {'01', '07'}, Fecha Modificación en mes anterior

    alt_mode: Si True, CD VES usa Fecha Modificación como Mes/Año (si válida), sino Fecha Ingreso
    """
    almacen_tipo1 = fila[idx['almacen_tipo1']]
    centro = fila[idx['centro']]
    almacen_val = fila[idx['almacen']]
    estado = fila[idx['estado']]
    tipo_stock = fila[idx['tipo_stock']]
    fecha_ing = fila[idx['fecha_ing']]
    fecha_mod = fila[idx['fecha_mod']]

    tipo_stock_str = str(tipo_stock).strip() if tipo_stock else ''
    estado_str = str(estado).strip() if estado else ''

    # Obtener fecha de modificación como entero para filtros
    try:
        fecha_mod_int = int(float(fecha_mod)) if fecha_mod else 0
    except:
        fecha_mod_int = 0

    en_mes_anterior = fecha_inicio <= fecha_mod_int <= fecha_fin

    # CD VES (sin filtro de fecha)
    if (almacen_tipo1 == 'CD VES' and
        centro == 'P008' and
        almacen_val == 'H000' and
        estado == 'ALMA' and
        tipo_stock_str in ['1', '01']):

        if alt_mode:
            mes_año = convertir_fecha_mes_año(fecha_mod)
            if mes_año is None:
                mes_año = convertir_fecha_mes_año(fecha_ing)
        else:
            mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('CD VES', mes_año)

    # PDV (CADENA) - filtro por Fecha Modificación
    if (almacen_tipo1 == 'CADENA' and
        estado_str in ESTADOS_CADENA and
        tipo_stock_str in ['2', '02'] and
        en_mes_anterior):

        mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('PDV', mes_año)

    # PDV (CAV/CAC/OTROS CADS) - filtro por Fecha Modificación
    if (almacen_tipo1 in ALMACENES_TIPO1_PDV and
        almacen_val in ['H000', 'H001'] and
        estado == 'ALMA' and
        tipo_stock_str in ['1', '01'] and
        en_mes_anterior):

        mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('PDV', mes_año)

    # Televentas VES - filtro por Fecha Modificación
    if (centro == 'P150' and
        almacen_val == 'H000' and
        estado == 'ALMA' and
        tipo_stock_str in ['1', '01', '7', '07'] and
        en_mes_anterior):

        mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('Televentas VES', mes_año)

    # Tienda Virtual VES - filtro por Fecha Modificación
    if (centro == 'P102' and
        almacen_val == 'H000' and
        estado == 'ALMA' and
        tipo_stock_str in ['1', '01', '7', '07'] and
        en_mes_anterior):

        mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('Tienda Virtual VES', mes_año)

    # Corporativo Ves - filtro por Fecha Modificación
    if (centro == 'P101' and
        almacen_val in ['H001', 'H002'] and
        estado == 'ALMA' and
        tipo_stock_str in ['1', '01', '7', '07'] and
        en_mes_anterior):

        mes_año = convertir_fecha_mes_año(fecha_ing)
        return ('Corporativo Ves', mes_año)

    return (None, None)


def leer_datos_desde_zip(ruta_zip, tipo_archivo, fecha_inicio, fecha_fin, clasificar_fila):
    """
    Lee datos desde archivos_base.zip (TXT pipe-delimited) y aplica clasificación.

    Retorna: (header_cols, filas_salida, series_verificar, contadores, count_no_seriado, count_sin_marca)
    """
    archivo_patron = ZIP_FILES[tipo_archivo]

    # Inicializar contadores
    if tipo_archivo == 'FIJA':
        contadores = {'CD VES': 0, 'ALMACENES': 0}
    else:
        contadores = {
            'CD VES': 0, 'PDV': 0, 'Televentas VES': 0,
            'Tienda Virtual VES': 0, 'Corporativo Ves': 0
        }

    filas_salida = []
    series_verificar = []
    count_no_seriado = 0
    count_sin_marca = 0

    print(f"\n[1/6] Abriendo archivo zip...")

    with zipfile.ZipFile(ruta_zip) as z:
        # Buscar archivo: primero exacto, luego por nombre parcial
        nombres = z.namelist()
        base_nombre = os.path.splitext(os.path.basename(archivo_patron))[0]  # ej: 'Descarga_inventario_Infra_Valorado'
        carpeta_patron = archivo_patron.split('/')[0]  # ej: 'Fija'

        archivo_interno = None
        if archivo_patron in nombres:
            archivo_interno = archivo_patron
        else:
            for nombre in nombres:
                if (nombre.endswith('.txt') and
                    carpeta_patron in nombre and
                    base_nombre in nombre):
                    archivo_interno = nombre
                    break

        if archivo_interno is None:
            print(f"ERROR: No se encontró archivo que coincida con '{archivo_patron}' en el ZIP")
            print(f"    Archivos disponibles: {[n for n in nombres if n.endswith('.txt')]}")
            sys.exit(1)

        print(f"    Archivo interno: {archivo_interno}")

        with z.open(archivo_interno) as f:
            # Leer header
            header_line = f.readline().decode('latin-1').strip()
            header_cols = header_line.split('|')

            # Construir diccionario de índices
            idx = {}
            for key, col_name in TXT_COL_MAP.items():
                try:
                    idx[key] = header_cols.index(col_name)
                except ValueError:
                    print(f"ERROR: No se encontró la columna '{col_name}' en el archivo TXT")
                    print(f"    Columnas disponibles: {header_cols[:10]}...")
                    sys.exit(1)

            # Manejar duplicado "Status" → usar la segunda ocurrencia (col 37 = Estado)
            status_positions = [i for i, h in enumerate(header_cols) if h == 'Status']
            if len(status_positions) < 2:
                print(f"ERROR: Se esperaban 2 columnas 'Status', se encontraron {len(status_positions)}")
                sys.exit(1)
            idx['estado'] = status_positions[1]

            # Índices de columnas texto (para preservar formato en salida)
            indices_texto = []
            for col_name in COLUMNAS_TEXTO_TXT:
                if col_name in header_cols:
                    indices_texto.append(header_cols.index(col_name))

            print(f"\n[2/6] Encabezados leídos: {len(header_cols)} columnas")
            print(f"\n[3/6] Procesando datos...")
            print(f"    Rango fechas filtro: {fecha_inicio} - {fecha_fin}")

            count_lineas = 0
            for line_bytes in f:
                line = line_bytes.decode('latin-1').strip()
                if not line:
                    continue

                fila = line.split('|')
                count_lineas += 1

                if count_lineas % 100000 == 0:
                    print(f"    Procesadas {count_lineas:,} filas...")

                # Filtrar No Seriado
                antiguedad = fila[idx['antiguedad']] if idx['antiguedad'] < len(fila) else ''
                if antiguedad == 'No Seriado':
                    count_no_seriado += 1
                    continue

                # MOVIL: Filtrar registros sin Marca
                if tipo_archivo == 'MOVIL' and idx.get('marca') is not None:
                    marca = fila[idx['marca']] if idx['marca'] < len(fila) else ''
                    if marca is None or str(marca).strip() == '':
                        count_sin_marca += 1
                        continue

                # Clasificar
                almacen_agrupado, mes_año = clasificar_fila(fila, idx, fecha_inicio, fecha_fin)

                if almacen_agrupado:
                    # Convertir columnas texto (asegurar que sean strings)
                    nueva_fila = list(fila)
                    for idx_txt in indices_texto:
                        if idx_txt < len(nueva_fila):
                            val = nueva_fila[idx_txt]
                            if val is not None and not isinstance(val, str):
                                if isinstance(val, float) and val == int(val):
                                    nueva_fila[idx_txt] = str(int(val))
                                else:
                                    nueva_fila[idx_txt] = str(val)
                    nueva_fila = nueva_fila + [almacen_agrupado, mes_año]
                    filas_salida.append(nueva_fila)
                    contadores[almacen_agrupado] = contadores.get(almacen_agrupado, 0) + 1

                    # Recolectar series para verificación
                    serie = fila[idx['serie']] if idx['serie'] < len(fila) else None
                    if serie is not None and str(serie).strip() != '':
                        serie = str(serie).strip()
                        if tipo_archivo == 'FIJA':
                            if almacen_agrupado == 'ALMACENES':
                                series_verificar.append(serie)
                        else:
                            if almacen_agrupado in ['PDV', 'Televentas VES', 'Tienda Virtual VES', 'Corporativo Ves']:
                                series_verificar.append(serie)

            print(f"    Total filas leídas: {count_lineas:,}")

    return header_cols, filas_salida, series_verificar, contadores, count_no_seriado, count_sin_marca


def escribir_procesado_xlsb(excel, headers_vals, headers, filas_salida, columnas_texto_salida, archivo_salida, batch_size=50000):
    """Escribe datos procesados FIFO a archivo xlsb via Excel COM."""
    print(f"    Creando archivo: {os.path.basename(archivo_salida)}...")
    wb = excel.Workbooks.Add()
    ws = wb.ActiveSheet
    ws.Name = "Data"

    nuevos_headers = list(headers_vals) + ['Almacen Agrupado', 'Mes/Año']
    for i, h in enumerate(nuevos_headers, 1):
        ws.Cells(1, i).Value = h

    if filas_salida:
        num_cols = len(filas_salida[0])
        total = len(filas_salida)
        ultima_fila = total + 1

        print(f"    Pre-formateando columnas de texto...")
        for col_name in columnas_texto_salida:
            if col_name in headers:
                col_idx = headers[col_name] + 1
                rango = ws.Range(ws.Cells(2, col_idx), ws.Cells(ultima_fila, col_idx))
                rango.NumberFormat = "@"

        print(f"    Escribiendo {total:,} registros...")
        fila = 2
        for i in range(0, total, batch_size):
            fin = min(i + batch_size, total)
            print(f"      Registros {i+1:,} - {fin:,}")
            batch = filas_salida[i:fin]
            rango = ws.Range(ws.Cells(fila, 1), ws.Cells(fila + len(batch) - 1, num_cols))
            rango.Value = batch
            fila += len(batch)

    col_mes_año = len(nuevos_headers)
    uf = len(filas_salida) + 1 if filas_salida else 1
    ws.Range(ws.Cells(2, col_mes_año), ws.Cells(uf, col_mes_año)).NumberFormat = "mmm-yy"

    wb.SaveAs(archivo_salida, FileFormat=50)
    print(f"    Guardado: {archivo_salida}")
    wb.Close(SaveChanges=False)


def main():
    archivo_entrada, tipo_archivo, formato = obtener_archivo_entrada()

    # Calcular mes anterior
    fecha_actual = datetime.now()
    primer_dia_mes_actual = fecha_actual.replace(day=1)
    ultimo_dia_mes_anterior = primer_dia_mes_actual - timedelta(days=1)
    año_mes_anterior = ultimo_dia_mes_anterior.year
    mes_num_anterior = ultimo_dia_mes_anterior.month

    nombre_mes_esp = MESES_ESP[mes_num_anterior]

    # Nombres de archivos según tipo
    archivo_salida_principal = os.path.join(DIRECTORIO, f"FIFO_PROCESADO_{tipo_archivo}.xlsb")
    archivo_salida_series = os.path.join(DIRECTORIO, f"VERIFICAR SALIDA CDVES {tipo_archivo} {nombre_mes_esp}.xlsx")

    # Nombre de la hoja de series según tipo
    nombre_hoja_series = tipo_archivo

    print(f"\n=== PROCESAMIENTO FIFO {tipo_archivo} ===")
    print(f"Formato entrada: {formato.upper()}")
    print(f"Fecha actual: {fecha_actual.strftime('%d/%m/%Y')}")
    print(f"Mes anterior: {nombre_mes_esp} {año_mes_anterior}")
    print(f"Archivo entrada: {archivo_entrada}")
    print(f"Tipo archivo: {tipo_archivo}")

    # Rango de fechas para filtros (mes anterior)
    fecha_inicio = int(f"{año_mes_anterior}{mes_num_anterior:02d}01")
    ultimo_dia = (datetime(año_mes_anterior, mes_num_anterior, 1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    fecha_fin = int(f"{año_mes_anterior}{mes_num_anterior:02d}{ultimo_dia.day:02d}")

    # Seleccionar función de clasificación según tipo
    clasificar_fila = clasificar_fila_fija if tipo_archivo == 'FIJA' else clasificar_fila_movil

    # ===== LECTURA DE DATOS: bifurcar según formato =====

    if formato == 'zip':
        # --- Formato nuevo: ZIP con TXT pipe-delimited ---
        headers_vals, filas_salida, series_verificar, contadores, count_no_seriado, count_sin_marca = \
            leer_datos_desde_zip(archivo_entrada, tipo_archivo, fecha_inicio, fecha_fin, clasificar_fila)

        # Construir dict headers para la escritura (índice por nombre)
        headers = {}
        for i, h in enumerate(headers_vals):
            if h:
                headers[str(h).strip()] = i

        # Usar COLUMNAS_TEXTO_TXT para formato texto en salida
        columnas_texto_salida = COLUMNAS_TEXTO_TXT

    else:
        # --- Formato legado: XLSB via COM ---
        headers = None  # se asigna abajo
        filas_salida = []
        series_verificar = []

        if tipo_archivo == 'FIJA':
            contadores = {'CD VES': 0, 'ALMACENES': 0}
        else:
            contadores = {
                'CD VES': 0, 'PDV': 0, 'Televentas VES': 0,
                'Tienda Virtual VES': 0, 'Corporativo Ves': 0
            }
        count_no_seriado = 0
        count_sin_marca = 0

        columnas_texto_salida = COLUMNAS_TEXTO

    # ===== ESCRITURA: siempre via Excel COM =====

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        if formato == 'xlsb':
            # Leer datos desde XLSB via COM
            print(f"\n[1/6] Abriendo archivo...")
            wb_origen = excel.Workbooks.Open(archivo_entrada)
            ws_origen = wb_origen.ActiveSheet

            ultima_fila = ws_origen.UsedRange.Rows.Count
            ultima_col = ws_origen.UsedRange.Columns.Count
            print(f"    Filas: {ultima_fila:,}, Columnas: {ultima_col}")

            print(f"\n[2/6] Leyendo encabezados...")
            headers_range = ws_origen.Range(ws_origen.Cells(1, 1), ws_origen.Cells(1, ultima_col))
            headers_vals = headers_range.Value[0]

            headers = {}
            for i, h in enumerate(headers_vals):
                if h:
                    headers[str(h).strip()] = i

            columnas_requeridas = ['Antiguedad', 'Amacén Tipo 1', 'Centro', 'Almacén', 'Estado',
                                   'Tipo Stock', 'Fecha de ingreso', 'Fecha Modificación', 'Serie']
            for col in columnas_requeridas:
                if col not in headers:
                    print(f"ERROR: No se encontró la columna '{col}'")
                    wb_origen.Close(SaveChanges=False)
                    excel.Quit()
                    sys.exit(1)

            idx = {
                'antiguedad': headers['Antiguedad'],
                'almacen_tipo1': headers['Amacén Tipo 1'],
                'centro': headers['Centro'],
                'almacen': headers['Almacén'],
                'estado': headers['Estado'],
                'tipo_stock': headers['Tipo Stock'],
                'fecha_ing': headers['Fecha de ingreso'],
                'fecha_mod': headers['Fecha Modificación'],
                'serie': headers['Serie'],
                'marca': headers.get('Marca')
            }

            indices_texto = []
            for col_name in COLUMNAS_TEXTO:
                if col_name in headers:
                    indices_texto.append(headers[col_name])

            print(f"\n[3/6] Procesando datos por lotes...")
            print(f"    Rango fechas filtro: {fecha_inicio} - {fecha_fin}")

            fila_actual = 2
            while fila_actual <= ultima_fila:
                fila_fin_lote = min(fila_actual + BATCH_SIZE - 1, ultima_fila)
                print(f"    Lote: filas {fila_actual:,} - {fila_fin_lote:,} de {ultima_fila:,}")

                data_range = ws_origen.Range(ws_origen.Cells(fila_actual, 1), ws_origen.Cells(fila_fin_lote, ultima_col))
                batch_data = data_range.Value

                if fila_fin_lote == fila_actual:
                    batch_data = [batch_data]

                for fila in batch_data:
                    antiguedad = fila[idx['antiguedad']]
                    if antiguedad == 'No Seriado':
                        count_no_seriado += 1
                        continue

                    if tipo_archivo == 'MOVIL' and idx['marca'] is not None:
                        marca = fila[idx['marca']]
                        if marca is None or str(marca).strip() == '':
                            count_sin_marca += 1
                            continue

                    almacen_agrupado, mes_año = clasificar_fila(fila, idx, fecha_inicio, fecha_fin)

                    if almacen_agrupado:
                        nueva_fila = list(fila)
                        for idx_txt in indices_texto:
                            val = nueva_fila[idx_txt]
                            if val is not None and not isinstance(val, str):
                                if isinstance(val, float) and val == int(val):
                                    nueva_fila[idx_txt] = str(int(val))
                                else:
                                    nueva_fila[idx_txt] = str(val)
                        nueva_fila = nueva_fila + [almacen_agrupado, mes_año]
                        filas_salida.append(nueva_fila)
                        contadores[almacen_agrupado] = contadores.get(almacen_agrupado, 0) + 1

                        serie = fila[idx['serie']]
                        if serie is not None:
                            if not isinstance(serie, str):
                                if isinstance(serie, float) and serie == int(serie):
                                    serie = str(int(serie))
                                else:
                                    serie = str(serie)

                            if tipo_archivo == 'FIJA':
                                if almacen_agrupado == 'ALMACENES':
                                    series_verificar.append(serie)
                            else:
                                if almacen_agrupado in ['PDV', 'Televentas VES', 'Tienda Virtual VES', 'Corporativo Ves']:
                                    series_verificar.append(serie)

                fila_actual = fila_fin_lote + 1

            wb_origen.Close(SaveChanges=False)

        # ===== Resumen (ambos formatos) =====
        print(f"\n    Resumen:")
        print(f"    - No Seriado eliminados: {count_no_seriado:,}")
        if tipo_archivo == 'MOVIL':
            print(f"    - Sin Marca eliminados: {count_sin_marca:,}")
        for grupo, cantidad in contadores.items():
            print(f"    - {grupo}: {cantidad:,}")
        print(f"    - Total registros: {len(filas_salida):,}")

        # ===== Crear archivo de salida (ambos formatos) =====
        print(f"\n[4/6] Creando archivo de salida...")
        escribir_procesado_xlsb(excel, headers_vals, headers, filas_salida, columnas_texto_salida, archivo_salida_principal)

        # Archivo de series para verificación
        print(f"\n[5/6] Creando archivo de series...")
        wb_series = excel.Workbooks.Add()
        ws_series = wb_series.ActiveSheet
        ws_series.Name = nombre_hoja_series
        ws_series.Cells(1, 1).Value = "Serie"

        if series_verificar:
            # Pre-formatear columna como texto ANTES de escribir
            ws_series.Range(ws_series.Cells(2, 1), ws_series.Cells(len(series_verificar) + 1, 1)).NumberFormat = "@"

            series_data = [[s] for s in series_verificar]
            for i in range(0, len(series_data), BATCH_SIZE):
                fin = min(i + BATCH_SIZE, len(series_data))
                batch = series_data[i:fin]
                rango = ws_series.Range(
                    ws_series.Cells(i + 2, 1),
                    ws_series.Cells(i + 2 + len(batch) - 1, 1)
                )
                rango.Value = batch

        wb_series.SaveAs(archivo_salida_series, FileFormat=51)
        print(f"    Guardado: {archivo_salida_series}")
        wb_series.Close(SaveChanges=False)

        # ===== Reporte alternativo (ALT) =====
        # CD VES usa Fecha Modificación como Mes/Año (si válida), sino Fecha Ingreso
        if formato == 'zip':
            print(f"\n[6/6] Generando reporte alternativo (ALT)...")
            print(f"    CD VES: Mes/Año = Fecha Modificación (si válida), sino Fecha Ingreso")
            archivo_salida_alt = os.path.join(DIRECTORIO, f"FIFO_PROCESADO_{tipo_archivo}_ALT.xlsb")
            clasificar_fila_alt = partial(clasificar_fila, alt_mode=True)
            _, filas_salida_alt, _, contadores_alt, _, _ = \
                leer_datos_desde_zip(archivo_entrada, tipo_archivo, fecha_inicio, fecha_fin, clasificar_fila_alt)

            print(f"    Resumen ALT:")
            for grupo, cantidad in contadores_alt.items():
                print(f"    - {grupo}: {cantidad:,}")
            print(f"    - Total registros: {len(filas_salida_alt):,}")

            escribir_procesado_xlsb(excel, headers_vals, headers, filas_salida_alt, columnas_texto_salida, archivo_salida_alt)

    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            excel.Quit()
        except:
            pass

    print(f"\n{'='*50}")
    print(f"=== PROCESO COMPLETADO ({tipo_archivo}) ===")
    print(f"{'='*50}")
    print(f"\nArchivo principal: {archivo_salida_principal}")
    print(f"  - Total registros: {len(filas_salida):,}")
    for grupo, cantidad in contadores.items():
        print(f"  - {grupo}: {cantidad:,}")
    print(f"\nArchivo de series: {archivo_salida_series}")
    print(f"  - Series: {len(series_verificar):,}")
    if formato == 'zip':
        print(f"\nArchivo alternativo: {os.path.join(DIRECTORIO, f'FIFO_PROCESADO_{tipo_archivo}_ALT.xlsb')}")
        print(f"  - CD VES Mes/Año basado en Fecha Modificación")


if __name__ == "__main__":
    main()
