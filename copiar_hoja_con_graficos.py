import win32com.client as win32
import pythoncom
import os
import logging
from logging.handlers import RotatingFileHandler
import tempfile
import sys

# Configuración básica
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CARPETA = "reportes"  # Carpeta para logs (se creará si no existe)

def configurar_logging():
    """Configura el sistema de logging con archivo rotativo"""
    log_file = os.path.join(BASE_DIR, "prueba.log")
    
    logger = logging.getLogger('ExcelCopyLogger')
    logger.setLevel(logging.INFO)
    
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    try:
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        handler = RotatingFileHandler(
            log_file,
            maxBytes=5*1024*1024,  # 5 MB
            backupCount=3,
            encoding='utf-8'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    except Exception as e:
        print(f"No se pudo configurar archivo de log: {e}", file=sys.stderr)
        # Fallback a consola
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    return logger

logger = configurar_logging()

def escribir_log(mensaje, nivel="info"):
    """Escribe un mensaje en el log"""
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

def verificar_archivos(plantilla, destino):
    """Verifica que los archivos existan y sean accesibles"""
    if not os.path.exists(plantilla):
        escribir_log(f"Archivo plantilla no encontrado: {plantilla}", "error")
        raise FileNotFoundError(f"No se encontró el archivo plantilla: {plantilla}")
    
    if not os.path.exists(destino):
        escribir_log(f"Creando archivo de destino vacío: {destino}", "warning")
        try:
            # Crear un archivo Excel vacío
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Add()
            wb.SaveAs(destino)
            wb.Close()
            excel.Quit()
            escribir_log(f"Archivo de destino creado exitosamente: {destino}")
        except Exception as e:
            escribir_log(f"Error al crear archivo de destino: {str(e)}", "error")
            raise

def copiar_hoja_con_graficos(origen_path, destino_path, nombre_hoja):
    """Copia una hoja con gráficos entre archivos Excel"""
    escribir_log(f"Iniciando proceso de copia de hoja '{nombre_hoja}'")
    
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        # Verificar y abrir archivos
        verificar_archivos(origen_path, destino_path)
        
        escribir_log(f"Abriendo archivo origen: {origen_path}")
        wb_origen = excel.Workbooks.Open(os.path.abspath(origen_path))
        
        escribir_log(f"Abriendo archivo destino: {destino_path}")
        wb_destino = excel.Workbooks.Open(os.path.abspath(destino_path))
        
        # Buscar hoja en origen
        escribir_log(f"Buscando hoja '{nombre_hoja}' en origen")
        hoja_origen = None
        for sheet in wb_origen.Sheets:
            if sheet.Name == nombre_hoja:
                hoja_origen = sheet
                break
        
        if not hoja_origen:
            error_msg = f"Hoja '{nombre_hoja}' no encontrada en {origen_path}"
            escribir_log(error_msg, "error")
            raise Exception(error_msg)
        
        # Copiar hoja
        escribir_log(f"Copiando hoja '{nombre_hoja}' a destino")
        hoja_origen.Copy(Before=wb_destino.Sheets(1))  # Copiar al inicio
        
        # Guardar cambios
        escribir_log("Guardando cambios en archivo destino")
        wb_destino.Save()
        
        escribir_log("Proceso completado exitosamente")
        
    except Exception as e:
        escribir_log(f"Error durante el proceso: {str(e)}", "error")
        raise
    finally:
        try:
            escribir_log("Cerrando archivos y liberando recursos")
            if 'wb_origen' in locals():
                wb_origen.Close(False)
            if 'wb_destino' in locals():
                wb_destino.Close(True)
            excel.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            escribir_log(f"Error al cerrar recursos: {str(e)}", "warning")

if __name__ == "__main__":
    try:
        # Configurar rutas en el mismo directorio del script
        directorio_script = os.path.dirname(os.path.abspath(__file__))
        ARCHIVO_PLANTILLA = os.path.join(directorio_script, "plantilla.xlsx")
        ARCHIVO_PRUEBAS = os.path.join(directorio_script, "pruebas.xlsx")
        NOMBRE_HOJA = "IR Julio 2025"
        
        escribir_log(f"Script iniciado desde: {directorio_script}")
        escribir_log(f"Archivo plantilla: {ARCHIVO_PLANTILLA}")
        escribir_log(f"Archivo pruebas: {ARCHIVO_PRUEBAS}")
        
        # Mostrar contenido del directorio para diagnóstico
        escribir_log(f"Contenido del directorio: {os.listdir(directorio_script)}")
        
        # Ejecutar copia
        copiar_hoja_con_graficos(ARCHIVO_PLANTILLA, ARCHIVO_PRUEBAS, NOMBRE_HOJA)
        
    except Exception as e:
        escribir_log(f"Error fatal: {str(e)}", "error")
        sys.exit(1)
