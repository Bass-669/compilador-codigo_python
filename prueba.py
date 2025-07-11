import logging
import os
import sys
import time
import win32com.client
import pandas as pd
from pathlib import Path
from logging.handlers import RotatingFileHandler

def configurar_logging():
    """Configura el sistema de logging en la carpeta del ejecutable"""
    try:
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        log_path = base_path / "log_tornos.txt"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
            handlers=[
                RotatingFileHandler(
                    log_path,
                    maxBytes=5*1024*1024,
                    backupCount=3,
                    encoding='utf-8'
                )
            ]
        )
        
        logger = logging.getLogger('TornosLogger')
        logger.info(f"Directorio base: {base_path}")
        logger.info(f"Archivos en el directorio: {[f.name for f in base_path.glob('*') if f.is_file()]}")
        return logger
        
    except Exception as e:
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger('TornosLogger')
        logger.error(f"Error configurando log: {str(e)}")
        return logger

logger = configurar_logging()

def encontrar_archivo_ode():
    """Busca el archivo ODC con flexibilidad en el nombre"""
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    
    # Posibles variantes del nombre del archivo
    posibles_patrones = [
        "*CLNALMISOTPRD*Peeling*Production*.ode",
        "*rwExport*Peeling*Production*.ode",
        "*Peeling*Production*.ode",
        "*.ode"  # Último recurso: cualquier archivo ODE
    ]
    
    for patron in posibles_patrones:
        archivos = list(base_path.glob(patron))
        if archivos:
            logger.info(f"Archivo encontrado con patrón '{patron}': {archivos[0]}")
            return archivos[0]
    
    logger.error(f"No se encontró archivo ODE. Directorio: {base_path}")
    logger.error(f"Archivos presentes: {[f.name for f in base_path.iterdir()]}")
    return None

def exportar_desde_odc():
    try:
        logger.info("Iniciando proceso de extracción de datos...")
        
        # Buscar el archivo ODC
        odc_path = encontrar_archivo_ode()
        if not odc_path:
            raise FileNotFoundError("No se pudo encontrar el archivo ODC")
        
        # Convertir a ruta absoluta y asegurar formato para COM
        ruta_absoluta = str(odc_path.absolute())
        logger.info(f"Ruta absoluta del archivo: {ruta_absoluta}")

        # Iniciar Excel
        logger.info("Iniciando Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Cambiar a True para depuración
        
        # Abrir archivo ODC
        logger.info("Abriendo archivo ODC...")
        workbook = excel.Workbooks.Open(ruta_absoluta)
        
        # Esperar a que cargue
        logger.info("Esperando carga de datos...")
        time.sleep(15)
        
        # Guardar como Excel
        output_path = odc_path.parent / "datos_actualizados.xlsx"
        logger.info(f"Guardando como: {output_path}")
        workbook.SaveAs(str(output_path.absolute()), FileFormat=51)
        workbook.Close()
        excel.Quit()
        
        # Leer datos
        logger.info("Leyendo datos exportados...")
        datos = pd.read_excel(output_path)
        logger.info(f"Datos obtenidos. Filas: {len(datos)}. Columnas: {list(datos.columns)}")
        
        return datos
    
    except Exception as e:
        logger.error(f"Error en el proceso: {str(e)}", exc_info=True)
        if 'excel' in locals():
            excel.Quit()
        return None

if __name__ == "__main__":
    logger.info("=== INICIO DE EJECUCIÓN ===")
    datos = exportar_desde_odc()
    logger.info("=== FIN DE EJECUCIÓN ===")
    
    # Mantener consola abierta solo en desarrollo
    if not getattr(sys, 'frozen', False):
        input("Presione Enter para salir...")
