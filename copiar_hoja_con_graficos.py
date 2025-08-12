import win32com.client as win32
import pythoncom
import os
import logging
import tempfile
import sys
import ctypes

# ========== CONFIGURACIÓN MODIFICABLE ==========
ARCHIVO_PLANTILLA_NOMBRE = "plantilla.xlsx"  # Nombre del archivo plantilla
ARCHIVO_DESTINO_NOMBRE = "prueba.xlsx"       # Nombre del archivo destino
NOMBRE_HOJA_A_COPIAR = "Informe Mensual"     # Nombre de la hoja a copiar (cambiar según necesidad)
CREAR_DESTINO_SI_NO_EXISTE = True            # Si True, crea el archivo destino si no existe
# ==============================================

def mostrar_mensaje(mensaje, titulo="Excel Copier", es_error=False):
    """Muestra mensaje en cuadro de diálogo"""
    estilo = 0x10 if es_error else 0x40  # 0x10: icono error, 0x40: icono info
    ctypes.windll.user32.MessageBoxW(0, mensaje, titulo, estilo)

def configurar_logging():
    """Configura logging para .exe"""
    logger = logging.getLogger('ExcelCopyLogger')
    logger.setLevel(logging.DEBUG)
    
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # Log en tempdir y escritorio
    log_locations = [
        os.path.join(tempfile.gettempdir(), 'excel_copy.log'),
        os.path.join(os.path.expanduser('~'), 'Desktop', 'excel_copy.log')
    ]
    
    for log_file in log_locations:
        try:
            handler = logging.FileHandler(log_file, mode='w', encoding='utf-8')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            break
        except Exception:
            continue
    
    return logger

logger = configurar_logging()

def escribir_log(mensaje, nivel="info"):
    """Escribe en log y muestra mensaje si es error"""
    try:
        getattr(logger, nivel.lower())(mensaje)
        if nivel.lower() == "error":
            mostrar_mensaje(mensaje, "Error", es_error=True)
    except Exception:
        pass

def verificar_archivos(plantilla, destino):
    """Verifica que los archivos existan y sean accesibles"""
    if not os.path.exists(plantilla):
        error_msg = f"Archivo plantilla no encontrado: {plantilla}"
        escribir_log(error_msg, "error")
        raise FileNotFoundError(error_msg)
    
    if not os.path.exists(destino):
        if CREAR_DESTINO_SI_NO_EXISTE:
            escribir_log(f"Creando archivo de destino: {destino}", "info")
            try:
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
        else:
            error_msg = f"Archivo destino no encontrado: {destino}"
            escribir_log(error_msg, "error")
            raise FileNotFoundError(error_msg)

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
        mostrar_mensaje(f"Hoja '{nombre_hoja}' copiada exitosamente a {ARCHIVO_DESTINO_NOMBRE}")
        
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
        # Obtener directorio del ejecutable (no __file__ en .exe)
        if getattr(sys, 'frozen', False):
            directorio = os.path.dirname(sys.executable)
        else:
            directorio = os.path.dirname(__file__)
            
        ARCHIVO_PLANTILLA = os.path.join(directorio, ARCHIVO_PLANTILLA_NOMBRE)
        ARCHIVO_DESTINO = os.path.join(directorio, ARCHIVO_DESTINO_NOMBRE)
        
        mostrar_mensaje(
            f"Se copiará la hoja:\n'{NOMBRE_HOJA_A_COPIAR}'\n\n"
            f"Desde:\n{ARCHIVO_PLANTILLA_NOMBRE}\n\n"
            f"Hacia:\n{ARCHIVO_DESTINO_NOMBRE}",
            "Confirmación de Proceso"
        )
        
        escribir_log(f"Iniciando proceso con hoja: {NOMBRE_HOJA_A_COPIAR}")
        escribir_log(f"Archivo plantilla: {ARCHIVO_PLANTILLA}")
        escribir_log(f"Archivo destino: {ARCHIVO_DESTINO}")
        
        copiar_hoja_con_graficos(ARCHIVO_PLANTILLA, ARCHIVO_DESTINO, NOMBRE_HOJA_A_COPIAR)
        
    except Exception as e:
        escribir_log(f"Error fatal: {str(e)}", "error")
        sys.exit(1)
