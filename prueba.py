import logging
from logging.handlers import RotatingFileHandler
import tempfile
import sys
import win32com.client
import time
import pandas as pd
import os
from pathlib import Path

# Configura el directorio base
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA = "logs"  # Carpeta para guardar logs

def configurar_logging():
    """Configura un sistema de logging robusto"""
    posibles_rutas = [
        os.path.join(BASE_DIR, CARPETA, "log_tornos.log"),
        os.path.join(tempfile.gettempdir(), "log_tornos.log"),
        os.path.join(os.path.expanduser("~"), "log_tornos.log")  # Nueva opción en carpeta de usuario
    ]
    
    logger = logging.getLogger('TornosLogger')
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    for ruta in posibles_rutas:
        try:
            # Crear directorio si no existe
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            
            # Verificar permisos de escritura
            if not os.access(os.path.dirname(ruta), os.W_OK):
                continue
                
            handler = RotatingFileHandler(
                ruta,
                maxBytes=5*1024*1024,  # 5 MB
                backupCount=3,
                encoding='utf-8'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            
            # Verificar que realmente se puede escribir
            logger.info(f"Iniciando logging en: {ruta}")
            return logger
        except Exception as e:
            print(f"No se pudo configurar log en {ruta}: {str(e)}", file=sys.stderr)
    
    # Si fallan todas las rutas, crear logger de consola
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.warning("No se pudo crear archivo de log en ninguna ubicación. Usando consola.")
    return logger

logger = configurar_logging()

def escribir_log(mensaje, nivel="info"):
    """Escribe en el log de manera segura"""
    try:
        if nivel.lower() == "info":
            logger.info(mensaje)
        elif nivel.lower() == "warning":
            logger.warning(mensaje)
        elif nivel.lower() == "error":
            logger.error(mensaje)
        else:
            logger.debug(mensaje)
    except Exception as e:
        print(f"Error al escribir en log: {e}", file=sys.stderr)

# El resto de tu código permanece igual...
        
def exportar_desde_odc():
    try:
        escribir_log("Iniciando proceso de extracción de datos...")
        
        # Obtener la ruta del directorio actual
        script_dir = Path(__file__).parent
        odc_file = "CLNALMISOTPRD rwExport report_Peeling_Production query.ode"
        odc_path = script_dir / odc_file
        
        # Verificar que el archivo .odc existe
        if not odc_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo {odc_file} en la misma carpeta que el script")
        
        escribir_log(f"Archivo .odc encontrado: {odc_path}")

        # Iniciar Excel
        escribir_log("Abriendo Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True  # Cambiar a False después de probar
        
        # Abrir el archivo .odc
        escribir_log(f"Abriendo el archivo {odc_file}...")
        workbook = excel.Workbooks.Open(str(odc_path))
        
        # Esperar a que cargue (ajusta este tiempo según necesites)
        print("Esperando a que carguen los datos...")
        time.sleep(15)
        
        # Guardar como Excel normal
        output_path = script_dir / "datos_actualizados.xlsx"
        escribir_log(f"Guardando datos en {output_path}...")
        workbook.SaveAs(str(output_path), FileFormat=51)  # 51 = xlsx
        workbook.Close()
        excel.Quit()
        
        # Leer los datos del Excel guardado
        escribir_log("Leyendo los datos exportados...")
        datos = pd.read_excel(output_path)
        
        # Mostrar información básica de los datos
        escribir_log("\n¡Proceso completado con éxito!")
        escribir_log(f"\nResumen de datos obtenidos ({len(datos)} filas):")
        escribir_log(datos.head(5))  # Muestra las primeras 5 filas
        escribir_log("\nColumnas disponibles:", list(datos.columns))
        
        return datos
    
    except Exception as e:
        escribir_log(f"\nError durante el proceso: {str(e)}")
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
