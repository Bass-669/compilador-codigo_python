import win32com.client as win32
import pythoncom
import os
import logging
import tempfile
import sys
import ctypes

# ========== CONFIGURACIÓN ==========
ARCHIVO_PLANTILLA = "plantilla.xlsx"  # Archivo de origen (no modificar)
ARCHIVO_DESTINO = "prueba.xlsx"       # Archivo de destino (no modificar)
NOMBRE_HOJA_A_COPIAR = "IR Julio 2025"      # Cambiar por el nombre exacto de la hoja a copiar
# ===================================

def mostrar_mensaje(mensaje, titulo="Excel Copier", es_error=False):
    """Muestra mensaje en cuadro de diálogo"""
    estilo = 0x10 if es_error else 0x40
    ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, estilo)

def configurar_logging():
    """Configura logging para el script"""
    logger = logging.getLogger('ExcelCopyLogger')
    logger.setLevel(logging.DEBUG)
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # Log en el directorio del script
    log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'excel_copy.log')
    try:
        handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    except Exception as e:
        mostrar_mensaje(f"No se pudo crear archivo log: {str(e)}", "Error", True)
    
    return logger

logger = configurar_logging()

def verificar_archivos():
    """Verifica que los archivos existan"""
    directorio = os.path.dirname(os.path.abspath(__file__))
    ruta_plantilla = os.path.join(directorio, ARCHIVO_PLANTILLA)
    ruta_destino = os.path.join(directorio, ARCHIVO_DESTINO)
    
    if not os.path.exists(ruta_plantilla):
        error_msg = f"No se encontró {ARCHIVO_PLANTILLA} en el directorio"
        logger.error(error_msg)
        mostrar_mensaje(error_msg, "Error", True)
        raise FileNotFoundError(error_msg)
    
    if not os.path.exists(ruta_destino):
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Add()
            wb.SaveAs(ruta_destino)
            wb.Close()
            excel.Quit()
            logger.info(f"Se creó {ARCHIVO_DESTINO} porque no existía")
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
        
        logger.info(f"Abriendo {ARCHIVO_PLANTILLA}")
        wb_origen = excel.Workbooks.Open(ruta_plantilla)
        
        logger.info(f"Abriendo {ARCHIVO_DESTINO}")
        wb_destino = excel.Workbooks.Open(ruta_destino)
        
        # Buscar hoja en origen
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
