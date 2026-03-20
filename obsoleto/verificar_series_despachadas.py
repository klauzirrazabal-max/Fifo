import os
import pandas as pd
from pathlib import Path

def procesar_carpeta(carpeta, filtro_alm, nombre_tipo):
    """
    Procesa archivos XLS (UTF-16) de una carpeta
    Filtra por Ce.=P008, Alm.=filtro_alm, Fe.contab.=noviembre 2025
    Retorna registros únicos con serie, Ce., Alm., Fe.contab.
    """
    registros = {}  # Usar dict para mantener series únicas con sus datos

    for archivo in carpeta.glob("*.XLS"):
        print(f"  Procesando: {archivo.name}")

        try:
            # Leer archivo con codificación UTF-16
            with open(archivo, 'r', encoding='utf-16') as f:
                contenido = f.read()

            # Procesar línea por línea
            lineas = contenido.split('\n')
            idx_ce = None
            idx_alm = None
            idx_fecha = None
            idx_serie = None

            for linea in lineas:
                linea = linea.strip()
                if not linea:
                    continue

                columnas = linea.split('\t')

                # Detectar si es una cabecera
                if 'Material' in columnas and 'Ce.' in columnas:
                    # Encontrar índices de las columnas
                    for i, col in enumerate(columnas):
                        col_limpio = col.strip()
                        if col_limpio == 'Ce.':
                            idx_ce = i
                        elif col_limpio == 'Alm.':
                            idx_alm = i
                        elif col_limpio == 'Fe.contab.':
                            idx_fecha = i
                        elif 'mero de serie' in col_limpio or col_limpio == 'Número de serie':
                            idx_serie = i
                    continue

                # Procesar líneas de datos
                if idx_ce is None or idx_alm is None or idx_fecha is None or idx_serie is None:
                    continue

                if len(columnas) <= max(idx_ce, idx_alm, idx_fecha, idx_serie):
                    continue

                ce = columnas[idx_ce].strip()
                alm = columnas[idx_alm].strip()
                fecha = columnas[idx_fecha].strip()
                serie = columnas[idx_serie].strip()

                # Aplicar filtros
                if ce != 'P008':
                    continue
                if alm != filtro_alm:
                    continue

                # Verificar fecha de noviembre 2025 (formato dd.11.2025)
                if '.11.2025' in fecha:
                    if serie and serie not in registros:  # Solo agregar si hay serie y es única
                        registros[serie] = {
                            'Numero_de_Serie': serie,
                            'Ce': ce,
                            'Alm': alm,
                            'Fe_contab': fecha
                        }

        except Exception as e:
            print(f"  Error procesando {archivo.name}: {e}")

    print(f"  Series únicas encontradas en {nombre_tipo}: {len(registros)}")
    return registros


def guardar_excel(registros, output_path, sheet_name='Series'):
    """Guarda los registros en un archivo Excel con formato texto"""
    # Crear DataFrame
    df = pd.DataFrame(list(registros.values()))

    # Ordenar por serie
    df = df.sort_values('Numero_de_Serie').reset_index(drop=True)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Formatear columnas como texto
        worksheet = writer.sheets[sheet_name]
        for col_idx in range(1, 5):  # 4 columnas
            for row in range(2, len(df) + 2):
                cell = worksheet.cell(row=row, column=col_idx)
                cell.number_format = '@'  # Formato texto

    print(f"Archivo guardado: {output_path}")
    return df


def main():
    base_path = Path("C:/FIFO/Fifo_extracted")

    # Procesar carpeta Fija (Ce.=P008, Alm.=P000)
    print("\n=== Procesando carpeta FIJA ===")
    print("Filtros: Ce.=P008, Alm.=P000, Fe.contab.=xx.11.2025")
    carpeta_fija = base_path / "Fija"
    registros_fija = procesar_carpeta(carpeta_fija, 'P000', 'Fija')

    # Procesar carpeta Movil (Ce.=P008, Alm.=H000)
    print("\n=== Procesando carpeta MOVIL ===")
    print("Filtros: Ce.=P008, Alm.=H000, Fe.contab.=xx.11.2025")
    carpeta_movil = base_path / "Movil"
    registros_movil = procesar_carpeta(carpeta_movil, 'H000', 'Movil')

    # Guardar archivos Excel
    print("\n=== Guardando archivos Excel ===")
    df_fija = guardar_excel(registros_fija, "C:/FIFO/series_validadas_fija.xlsx", 'Series_Fija')
    df_movil = guardar_excel(registros_movil, "C:/FIFO/series_validadas_movil.xlsx", 'Series_Movil')

    # Resumen
    print("\n" + "="*50)
    print("RESUMEN")
    print("="*50)
    print(f"Series únicas FIJA:  {len(df_fija)}")
    print(f"Series únicas MOVIL: {len(df_movil)}")
    print(f"TOTAL:               {len(df_fija) + len(df_movil)}")

    return df_fija, df_movil


if __name__ == "__main__":
    df_fija, df_movil = main()

    print("\n--- Primeras 5 series FIJA ---")
    print(df_fija.head())

    print("\n--- Primeras 5 series MOVIL ---")
    print(df_movil.head())
