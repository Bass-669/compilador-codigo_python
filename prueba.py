import logging
import os
from pathlib import Path
import sys
import win32com.client
import time
import pandas as pd
from logging.handlers import RotatingFileHandler

def configurar_logging():
    """Configura logging para escribir en un archivo en la misma carpeta del ejecutable"""
    try:
        # Obtener la ruta del directorio donde está el ejecutable
        if getattr(sys, 'frozen', False):
            # Si está empaquetado como .exe
            base_path = Path(sys.executable).parent
        else:
            # Si se ejecuta como script
            base_path = Path(__file__).parent
        
        log_path = base_path / "log_tornos.txt"
        
        # Configuración del logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
            handlers=[
                RotatingFileHandler(
                    log_path,
                    maxBytes=5*1024*1024,  # 5 MB
                    backupCount=3,
                    encoding='utf-8'
                )
            ]
        )
        
        logger = logging.getLogger('TornosLogger')
        logger.info(f"Iniciando aplicación. Log guardado en: {log_path}")
        return logger
        
    except Exception as e:
        # Fallback a logging básico si hay error
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger('TornosLogger')
        logger.error(f"No se pudo configurar el archivo de log: {str(e)}")
        return logger

logger = configurar_logging()

def escribir_log(mensaje, nivel="info"):
    """Escribe en el log convirtiendo cualquier objeto a string"""
    try:
        if not isinstance(mensaje, str):
            mensaje = str(mensaje)
            
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

def exportar_desde_odc():
    try:
        escribir_log("Iniciando proceso de extracción de datos...")
        
        # Obtener ruta base (funciona tanto para .exe como para script)
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        
        odc_file = "CLNALMISOTPRD rwExport report_Peeling_Production query.ode"
        odc_path = base_path / odc_file
        
        if not odc_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo {odc_file}")
        
        escribir_log(f"Archivo .odc encontrado: {odc_path}")

        # Resto del código de procesamiento...
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        
        workbook = excel.Workbooks.Open(str(odc_path))
        time.sleep(15)
        
        output_path = base_path / "datos_actualizados.xlsx"
        workbook.SaveAs(str(output_path), FileFormat=51)
        workbook.Close()
        excel.Quit()
        
        datos = pd.read_excel(output_path)
        escribir_log(f"Datos obtenidos correctamente. Filas: {len(datos)}")
        
        return datos
    
    except Exception as e:
        escribir_log(f"Error durante el proceso: {str(e)}", nivel="error")
        if 'excel' in locals():
            excel.Quit()
        return None

if __name__ == "__main__":
    datos_actualizados = exportar_desde_odc()
    escribir_log("Proceso finalizado")
