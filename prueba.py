import win32com.client
import time
import pandas as pd
import os
from pathlib import Path

def exportar_desde_odc():
    try:
        print("Iniciando proceso de extracción de datos...")
        
        # Obtener la ruta del directorio actual
        script_dir = Path(__file__).parent
        odc_file = "CLNALMISOTPRD rwExport report_Peeling_Production query.ode"
        odc_path = script_dir / odc_file
        
        # Verificar que el archivo .odc existe
        if not odc_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo {odc_file} en la misma carpeta que el script")
        
        print(f"Archivo .odc encontrado: {odc_path}")

        # Iniciar Excel
        print("Abriendo Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Cambiar a False después de probar
        
        # Abrir el archivo .odc
        print(f"Abriendo el archivo {odc_file}...")
        workbook = excel.Workbooks.Open(str(odc_path))
        
        # Esperar a que cargue (ajusta este tiempo según necesites)
        print("Esperando a que carguen los datos...")
        time.sleep(15)
        
        # Guardar como Excel normal
        output_path = script_dir / "datos_actualizados.xlsx"
        print(f"Guardando datos en {output_path}...")
        workbook.SaveAs(str(output_path), FileFormat=51)  # 51 = xlsx
        workbook.Close()
        excel.Quit()
        
        # Leer los datos del Excel guardado
        print("Leyendo los datos exportados...")
        datos = pd.read_excel(output_path)
        
        # Mostrar información básica de los datos
        print("\n¡Proceso completado con éxito!")
        print(f"\nResumen de datos obtenidos ({len(datos)} filas):")
        print(datos.head(5))  # Muestra las primeras 5 filas
        print("\nColumnas disponibles:", list(datos.columns))
        
        return datos
    
    except Exception as e:
        print(f"\nError durante el proceso: {str(e)}")
        # Asegurarse de cerrar Excel si hay error
        if 'excel' in locals():
            excel.Quit()
        return None

# Ejecutar la función
if __name__ == "__main__":
    datos_actualizados = exportar_desde_odc()
    
    # Mantener la consola abierta para ver los resultados
    if datos_actualizados is not None:
        input("\nPresiona Enter para salir...")
