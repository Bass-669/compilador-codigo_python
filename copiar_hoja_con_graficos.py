import win32com.client as win32
import pythoncom
import os
import logging
import sys
import ctypes

# ========== CONFIGURACIÓN ==========
ARCHIVO_PLANTILLA = "plantilla.xlsx"
ARCHIVO_DESTINO = "prueba.xlsx"
NOMBRE_HOJA_A_COPIAR = "IR Julio 2025"
# ===================================

def get_script_dir():
    """Obtiene el directorio correcto tanto para .exe como para .py"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def mostrar_mensaje(mensaje, titulo="Excel Copier", es_error=False):
    estilo = 0x10 if es_error else 0x40
    ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, estilo)

def configurar_logging():
    logger = logging.getLogger('ExcelCopyLogger')
    logger.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    try:
        log_file = os.path.join(get_script_dir(), 'excel_copy.log')
        file_handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        logger.error(f"No se pudo crear archivo log: {str(e)}")
    
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
    directorio = get_script_dir()
    
    logger.debug(f"Buscando en directorio: {directorio}")
    logger.debug(f"Archivos presentes: {os.listdir(directorio)}")
    
    ruta_plantilla = encontrar_archivo(ARCHIVO_PLANTILLA, directorio)
    if not ruta_plantilla:
        error_msg = (
            f"No se encontró {ARCHIVO_PLANTILLA} en:\n{directorio}\n\n"
            f"Archivos presentes:\n{chr(10).join(os.listdir(directorio))}"
        )
        logger.error(error_msg)
        mostrar_mensaje(error_msg, "Error", True)
        raise FileNotFoundError(error_msg)
    
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
    """Copia la hoja especificada de plantilla.xlsx a prueba.xlsx"""
    try:
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        ruta_plantilla, ruta_destino = verificar_archivos()
        
        logger.info(f"Abriendo archivo origen: {ruta_plantilla}")
        wb_origen = excel.Workbooks.Open(ruta_plantilla)
        
        logger.info(f"Abriendo archivo destino: {ruta_destino}")
        wb_destino = excel.Workbooks.Open(ruta_destino)
        
        # Buscar hoja en origen
        logger.info(f"Buscando hoja: {NOMBRE_HOJA_A_COPIAR}")
        hoja_origen = None
        for sheet in wb_origen.Sheets:
            if sheet.Name == NOMBRE_HOJA_A_COPIAR:
                hoja_origen = sheet
                break
        
        if not hoja_origen:
            error_msg = f"No se encontró la hoja '{NOMBRE_HOJA_A_COPIAR}' en {ARCHIVO_PLANTILLA}"
            logger.error(error_msg)
            mostrar_mensaje(error_msg, "Error", True)
            raise Exception(error_msg)
        
        # Copiar hoja al inicio del archivo destino
        logger.info(f"Copiando hoja '{NOMBRE_HOJA_A_COPIAR}'")
        hoja_origen.Copy(Before=wb_destino.Sheets(1))
        
        # Guardar cambios
        wb_destino.Save()
        logger.info("Proceso completado exitosamente")
        mostrar_mensaje(
            f"Hoja '{NOMBRE_HOJA_A_COPIAR}' copiada exitosamente\n"
            f"de {ARCHIVO_PLANTILLA} a {ARCHIVO_DESTINO}",
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
        logger.info(f"Hoja a copiar: {NOMBRE_HOJA_A_COPIAR}")
        logger.info(f"Archivo origen: {ARCHIVO_PLANTILLA}")
        logger.info(f"Archivo destino: {ARCHIVO_DESTINO}")
        
        copiar_hoja()
        
    except Exception as e:
        logger.error(f"Error fatal: {str(e)}")
        sys.exit(1)
    finally:
        logger.info("=== FIN DEL PROCESO ===")
        input("Presiona Enter para salir...")  # Para que no se cierre la ventana inmediatamente
