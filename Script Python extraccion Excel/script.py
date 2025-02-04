
import pandas as pd
import argparse
import os
import sys
from datetime import datetime

def validar_archivo(ruta):
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"Archivo no encontrado: {ruta}")
    if not ruta.lower().endswith(('.xls', '.xlsx')):
        raise ValueError("Formato de archivo no válido. Debe ser .xls o .xlsx")

def procesar_excel(ruta_entrada, hoja=None, n_filas=None, formato_salida='csv'):
    try:
        # Leer archivo Excel
        df = pd.read_excel(
            ruta_entrada,
            sheet_name=hoja,
            nrows=n_filas,
            engine='openpyxl' if ruta_entrada.endswith('.xlsx') else None
        )

        # Generar nombre de archivo de salida
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        nombre_base = os.path.splitext(os.path.basename(ruta_entrada))[0]
        archivo_salida = f"{nombre_base}_export_{timestamp}.{formato_salida}"

        # Exportar datos
        if formato_salida == 'csv':
            df.to_csv(archivo_salida, index=False)
        elif formato_salida == 'json':
            df.to_json(archivo_salida, indent=2, orient='records')
        else:
            df.to_excel(archivo_salida, index=False)

        return archivo_salida

    except Exception as e:
        print(f"\n[ERROR] Fallo al procesar el archivo: {str(e)}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description='Extractor de Datos de Excel')
    parser.add_argument('-i', '--input', required=True, help='Ruta del archivo Excel')
    parser.add_argument('-s', '--sheet', help='Nombre de la hoja a procesar')
    parser.add_argument('-n', '--rows', type=int, help='Número de filas a extraer')
    parser.add_argument('-f', '--format', choices=['csv', 'json', 'excel'], default='csv',
                       help='Formato de salida (default: csv)')

    args = parser.parse_args()

    try:
        print("\n=== Validadando archivo ===")
        validar_archivo(args.input)
        
        print("\n=== Procesando datos ===")
        archivo_salida = procesar_excel(
            args.input,
            hoja=args.sheet,
            n_filas=args.rows,
            formato_salida=args.format
        )

        print(f"\n[ÉXITO] Datos exportados correctamente a: {archivo_salida}")
        print(f"Ubicación completa: {os.path.abspath(archivo_salida)}")

    except Exception as e:
        print(f"\n[ERROR] {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
    
    #Dependecias
    #pip install pandas openpyxl xlrd
    #pip install pyinstaller
    #pyinstaller --onefile extractor_excel.py
