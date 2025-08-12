import win32com.client as win32
import pythoncom
import os
import sys
import ctypes
import time

# ========== CONFIGURACIÓN MODIFICABLE ==========
ARCHIVO_PLANTILLA = "plantilla.xlsx"              # Nombre del archivo origen
ARCHIVO_DESTINO = "Reporte IR Tornos.xlsx"        # Nombre del archivo destino
NOMBRE_HOJA_ORIGEN = "PLANTILLA"                  # Nombre de la hoja a copiar
NOMBRE_HOJA_DESTINO = "IR plantilla"            # Nombre de la hoja copiada
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

def encontrar_archivo(nombre_archivo, directorio):
    """Busca archivo ignorando mayúsculas/minúsculas"""
    try:
        for f in os.listdir(directorio):
            if f.lower() == nombre_archivo.lower():
                return os.path.join(directorio, f)
        return None
    except Exception as e:
        mostrar_mensaje(f"Error buscando archivo: {str(e)}", "Error", True)
        return None

def verificar_archivos():
    """Verifica que los archivos existan"""
    directorio = get_script_dir()
    
    # Buscar archivo plantilla
    ruta_plantilla = encontrar_archivo(ARCHIVO_PLANTILLA, directorio)
    if not ruta_plantilla:
        error_msg = f"No se encontró {ARCHIVO_PLANTILLA} en el directorio"
        mostrar_mensaje(error_msg, "Error", True)
        raise FileNotFoundError(error_msg)
    
    # Buscar o crear archivo destino
    ruta_destino = os.path.join(directorio, ARCHIVO_DESTINO)
    if not os.path.exists(ruta_destino):
        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Add()
            wb.SaveAs(ruta_destino)
            wb.Close()
            excel.Quit()
        except Exception as e:
            error_msg = f"No se pudo crear {ARCHIVO_DESTINO}: {str(e)}"
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
        
        ruta_plantilla, ruta_destino = verificar_archivos()
        
        # Abrir archivos
        wb_origen = excel.Workbooks.Open(ruta_plantilla)
        wb_destino = excel.Workbooks.Open(
            os.path.abspath(ruta_destino),
            UpdateLinks=0,
            IgnoreReadOnlyRecommended=True,
            ReadOnly=False
        )
        
        # Buscar hoja en origen
        hoja_origen = None
        for sheet in wb_origen.Sheets:
            if sheet.Name == NOMBRE_HOJA_ORIGEN:
                hoja_origen = sheet
                break
        
        if not hoja_origen:
            error_msg = f"No se encontró la hoja '{NOMBRE_HOJA_ORIGEN}' en {ARCHIVO_PLANTILLA}"
            mostrar_mensaje(error_msg, "Error", True)
            raise Exception(error_msg)
        
        # Copiar hoja
        hoja_origen.Copy(Before=wb_destino.Sheets(1))
        hoja_copiada = wb_destino.Sheets(1)
        hoja_copiada.Name = NOMBRE_HOJA_DESTINO
        
        # Guardar cambios
        wb_destino.Save()
        mostrar_mensaje(
            f"Hoja copiada y renombrada exitosamente:\n"
            f"Origen: '{NOMBRE_HOJA_ORIGEN}'\n"
            f"Destino: '{NOMBRE_HOJA_DESTINO}'",
            "Proceso Completado"
        )
        
    except Exception as e:
        mostrar_mensaje(f"Error: {str(e)}", "Error", True)
        raise
    finally:
        try:
            if 'wb_origen' in locals() and wb_origen is not None:
                wb_origen.Close(False)
            if 'wb_destino' in locals() and wb_destino is not None:
                wb_destino.Close(True)
            if 'excel' in locals():
                excel.Quit()
        except Exception as e:
            pass
        finally:
            pythoncom.CoUninitialize()

if __name__ == "__main__":
    try:
        copiar_hoja()
    except Exception as e:
        sys.exit(1)
