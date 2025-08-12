import win32com.client as win32
import pythoncom
import os
import logging
import sys
import ctypes
import time

# ========== CONFIGURACIÓN MODIFICABLE ==========
ARCHIVO_PLANTILLA = "plantilla.xlsx"      # Nombre del archivo origen
ARCHIVO_DESTINO = "Reporte IR Tornos.xlsx"           # Nombre del archivo destino
NOMBRE_HOJA_ORIGEN = "PLANTILLA"      # Nombre de la hoja a copiar (en plantilla.xlsx)
NOMBRE_HOJA_DESTINO = "IR Agosto 2025"   # Nombre que tendrá la hoja copiada (en prueba.xlsx)
# ==============================================

def get_script_dir():
    """Obtiene el directorio del ejecutable o script"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def mostrar_mensaje(mensaje, titulo="Excel Copier", es_error=False):
    """Muestra mensaje en cuadro de diálogo"""
    estilo = 0x10 if es_error else 0x40
    ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, estilo)

def configurar_logging():
    """Configura el sistema de logging"""
    logger = logging.getLogger('ExcelCopyLogger')
    logger.setLevel(logging.DEBUG)
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    try:
        log_file = os.path.join(get_script_dir(), 'excel_copy.log')
        file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        mostrar_mensaje(f"No se pudo crear archivo log: {str(e)}", "Error", True)
    
    return logger

logger = configurar_logging()

def encontrar_archivo(nombre_archivo, directorio):
    """Busca archivo ignorando mayúsculas/minúsculas"""
    try:
        for f in os.listdir(directorio):
            if f.lower() == nombre_archivo.lower():
                return os.path.join(directorio, f)
        return None
    except Exception as e:
        logger.error(f"Error buscando archivo: {str(e)}")
        return None

def verificar_archivos():
    """Verifica que los archivos existan"""
    directorio = get_script_dir()
    
    logger.info(f"Buscando archivos en: {directorio}")
    logger.info(f"Archivos presentes: {os.listdir(directorio)}")
    
    # Buscar archivo plantilla
    ruta_plantilla = encontrar_archivo(ARCHIVO_PLANTILLA, directorio)
    if not ruta_plantilla:
        error_msg = f"No se encontró {ARCHIVO_PLANTILLA} en el directorio"
        logger.error(error_msg)
        mostrar_mensaje(error_msg, "Error", True)
        raise FileNotFoundError(error_msg)
    
    # Buscar o crear archivo destino
    ruta_destino = os.path.join(directorio, ARCHIVO_DESTINO)
    if not os.path.exists(ruta_destino):
        try:
            logger.info(f"Creando archivo destino: {ruta_destino}")
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Add()
            wb.SaveAs(ruta_destino)
            wb.Close()
            excel.Quit()
        except Exception as e:
            error_msg = f"No se pudo crear {ARCHIVO_DESTINO}: {str(e)}"
            logger.error(error_msg)
            mostrar_mensaje(error_msg, "Error", True)
            raise
    
    return ruta_plantilla, ruta_destino

def copiar_hoja():
    """Copia la hoja especificada y le asigna el nombre deseado"""
    try:
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AskToUpdateLinks = False

        wb_origen = None
        wb_destino = None
        hoja_origen = None
        nueva_hoja = None
        
        ruta_plantilla, ruta_destino = verificar_archivos()
        
        # Abrir archivos
        logger.info(f"Abriendo archivo origen: {ruta_plantilla}")
        wb_origen = excel.Workbooks.Open(ruta_plantilla)
        
        logger.info(f"Abriendo archivo destino: {ruta_destino}")
        wb_destino = excel.Workbooks.Open(
                os.path.abspath(RUTA_ENTRADA),
                UpdateLinks=0,
                IgnoreReadOnlyRecommended=True,
                ReadOnly=False
            )
        
        # Buscar hoja en origen
        logger.info(f"Buscando hoja origen: {NOMBRE_HOJA_ORIGEN}")
        hoja_origen = None
        for sheet in wb_origen.Sheets:
            if sheet.Name == NOMBRE_HOJA_ORIGEN:
                hoja_origen = sheet
                break
        
        if not hoja_origen:
            error_msg = f"No se encontró la hoja '{NOMBRE_HOJA_ORIGEN}' en {ARCHIVO_PLANTILLA}"
            logger.error(error_msg)
            mostrar_mensaje(error_msg, "Error", True)
            raise Exception(error_msg)
        
        # Copiar hoja con verificación
        logger.info(f"Copiando hoja '{NOMBRE_HOJA_ORIGEN}'")
        try:
            # Primero copiar
            hoja_origen.Copy(Before=wb_destino.Sheets(1))
            
            # Luego obtener referencia a la hoja copiada
            hoja_copiada = wb_destino.Sheets(1)  # La hoja recién copiada estará en la primera posición
            
            # Verificar que se copió correctamente
            if hoja_copiada is None:
                raise Exception("No se pudo copiar la hoja")
                
            # Renombrar
            hoja_copiada.Name = NOMBRE_HOJA_DESTINO
            logger.info(f"Hoja renombrada a: '{NOMBRE_HOJA_DESTINO}'")
            
        except Exception as e:
            error_msg = f"Error al copiar/renombrar hoja: {str(e)}"
            logger.error(error_msg)
            mostrar_mensaje(error_msg, "Error", True)
            raise
        
        # Guardar cambios
        wb_destino.Save()
        logger.info("Proceso completado exitosamente")
        mostrar_mensaje(
            f"Hoja copiada y renombrada exitosamente:\n"
            f"Origen: '{NOMBRE_HOJA_ORIGEN}'\n"
            f"Destino: '{NOMBRE_HOJA_DESTINO}'",
            "Proceso Completado"
        )
        
    except Exception as e:
        logger.error(f"Error durante el proceso: {str(e)}")
        mostrar_mensaje(f"Error: {str(e)}", "Error", True)
        raise
    finally:
        try:
            if 'wb_origen' in locals():
                wb_origen.Close(False)
            if 'wb_destino' in locals():
                wb_destino.Close(True)
            excel.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.warning(f"Error al cerrar recursos: {str(e)}")

if __name__ == "__main__":
    try:
        logger.info("=== INICIO DEL PROCESO ===")
        logger.info(f"Configuración:")
        logger.info(f" - Archivo origen: {ARCHIVO_PLANTILLA}")
        logger.info(f" - Archivo destino: {ARCHIVO_DESTINO}")
        logger.info(f" - Hoja origen: {NOMBRE_HOJA_ORIGEN}")
        logger.info(f" - Hoja destino: {NOMBRE_HOJA_DESTINO}")
        
        copiar_hoja()
        
    except Exception as e:
        logger.error(f"Error fatal: {str(e)}")
        sys.exit(1)
    finally:
        logger.info("=== FIN DEL PROCESO ===")
        time.sleep(1)  # Pausa para leer logs si se ejecuta desde consola
