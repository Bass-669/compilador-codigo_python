import win32com.client as win32
import pythoncom
import os
import logging
from logging.handlers import RotatingFileHandler
import tempfile
import sys

# Configuración de paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA = "reportes"  # Carpeta donde se guardarán los logs

def configurar_logging():
    """Configura el sistema de logging con archivo rotativo"""
    posibles_rutas = [
        os.path.join(BASE_DIR, "prueba.log"),  # Archivo en el mismo directorio
        os.path.join(tempfile.gettempdir(), "prueba.log")  # Archivo en temp
    ]
    
    logger = logging.getLogger('ExcelLogger')
    logger.setLevel(logging.INFO)
    
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    for ruta in posibles_rutas:
        try:
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            handler = RotatingFileHandler(
                ruta,
                maxBytes=5*1024*1024,  # 5 MB
                backupCount=3,
                encoding='utf-8'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            return logger
        except Exception as e:
            print(f"No se pudo configurar log en {ruta}: {e}", file=sys.stderr)
    
    # Fallback a consola si no se pudo crear archivo
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.warning("No se pudo crear archivo de log. Usando consola.")
    return logger

logger = configurar_logging()

def escribir_log(mensaje, nivel="info"):
    """Escribe un mensaje en el log con el nivel especificado"""
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

def copiar_hoja_con_graficos(origen_path, destino_path, nombre_hoja):
    """
    Copia una hoja específica con todos sus gráficos de un archivo Excel a otro.
    
    Args:
        origen_path (str): Ruta del archivo Excel de origen (plantilla)
        destino_path (str): Ruta del archivo Excel de destino (pruebas)
        nombre_hoja (str): Nombre de la hoja a copiar
    """
    escribir_log(f"Iniciando copia de hoja '{nombre_hoja}' de {origen_path} a {destino_path}")
    
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        # Abrir archivos
        escribir_log(f"Abriendo archivo origen: {origen_path}")
        wb_origen = excel.Workbooks.Open(os.path.abspath(origen_path))
        
        escribir_log(f"Abriendo archivo destino: {destino_path}")
        wb_destino = excel.Workbooks.Open(os.path.abspath(destino_path))
        
        # Buscar hoja en origen
        escribir_log(f"Buscando hoja '{nombre_hoja}' en archivo origen")
        hoja_origen = None
        for sheet in wb_origen.Sheets:
            if sheet.Name == nombre_hoja:
                hoja_origen = sheet
                break
        
        if not hoja_origen:
            error_msg = f"No se encontró la hoja '{nombre_hoja}' en el archivo de origen"
            escribir_log(error_msg, "error")
            raise Exception(error_msg)
        
        # Copiar hoja al destino
        escribir_log(f"Copiando hoja '{nombre_hoja}' al archivo destino")
        hoja_origen.Copy(Before=wb_destino.Sheets(wb_destino.Sheets.Count))
        nueva_hoja = wb_destino.ActiveSheet
        
        # Guardar cambios
        escribir_log("Guardando cambios en archivo destino")
        wb_destino.Save()
        
        escribir_log(f"Operación completada exitosamente. Hoja '{nombre_hoja}' copiada")
        
    except Exception as e:
        error_msg = f"Error durante la copia: {str(e)}"
        escribir_log(error_msg, "error")
        raise
    finally:
        # Cerrar siempre los archivos y Excel
        try:
            escribir_log("Cerrando archivos Excel")
            wb_origen.Close(SaveChanges=False)
            wb_destino.Close(SaveChanges=True)
            excel.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            escribir_log(f"Error al cerrar recursos: {str(e)}", "warning")

if __name__ == "__main__":
    try:
        # Configurar rutas
        directorio_actual = os.path.dirname(os.path.abspath(__file__))
        ARCHIVO_PLANTILLA = os.path.join(directorio_actual, "plantilla.xlsx")
        ARCHIVO_PRUEBAS = os.path.join(directorio_actual, "pruebas.xlsx")
        NOMBRE_HOJA = "IR Julio 2025"
        
        escribir_log(f"Directorios configurados: Origen={ARCHIVO_PLANTILLA}, Destino={ARCHIVO_PRUEBAS}")
        
        # Verificar archivos
        if not os.path.exists(ARCHIVO_PLANTILLA):
            error_msg = f"Archivo plantilla no encontrado: {ARCHIVO_PLANTILLA}"
            escribir_log(error_msg, "error")
            raise FileNotFoundError(error_msg)
            
        if not os.path.exists(ARCHIVO_PRUEBAS):
            escribir_log(f"Creando archivo de pruebas vacío: {ARCHIVO_PRUEBAS}", "warning")
            open(ARCHIVO_PRUEBAS, 'a').close()  # Crear archivo vacío si no existe
            
        # Ejecutar copia
        copiar_hoja_con_graficos(ARCHIVO_PLANTILLA, ARCHIVO_PRUEBAS, NOMBRE_HOJA)
        
    except Exception as e:
        escribir_log(f"Error en ejecución principal: {str(e)}", "error")
        raise
