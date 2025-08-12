import win32com.client as win32
import pythoncom
import os

def copiar_hoja_con_graficos(origen_path, destino_path, nombre_hoja):
    """
    Copia una hoja específica con todos sus gráficos de un archivo Excel a otro.
    
    Args:
        origen_path (str): Ruta del archivo Excel de origen (plantilla)
        destino_path (str): Ruta del archivo Excel de destino (pruebas)
        nombre_hoja (str): Nombre de la hoja a copiar
    """
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Ocultar Excel durante la operación
    excel.DisplayAlerts = False  # Deshabilitar alertas
    
    try:
        # Abrir archivos
        wb_origen = excel.Workbooks.Open(os.path.abspath(origen_path))
        wb_destino = excel.Workbooks.Open(os.path.abspath(destino_path))
        
        # Buscar hoja en origen
        hoja_origen = None
        for sheet in wb_origen.Sheets:
            if sheet.Name == nombre_hoja:
                hoja_origen = sheet
                break
        
        if not hoja_origen:
            raise Exception(f"No se encontró la hoja '{nombre_hoja}' en el archivo de origen")
        
        # Copiar hoja al destino (se colocará al final)
        hoja_origen.Copy(Before=wb_destino.Sheets(wb_destino.Sheets.Count))
        nueva_hoja = wb_destino.ActiveSheet
        
        # Guardar cambios
        wb_destino.Save()
        print(f"Hoja '{nombre_hoja}' copiada exitosamente a {destino_path}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
    finally:
        # Cerrar siempre los archivos y Excel
        wb_origen.Close(SaveChanges=False)
        wb_destino.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()

# Ejemplo de uso
if __name__ == "__main__":
    # Obtener la ruta del directorio donde está este script
    directorio_actual = os.path.dirname(os.path.abspath(__file__))
    
    # Configurar rutas relativas al directorio actual
    ARCHIVO_PLANTILLA = os.path.join(directorio_actual, "plantilla.xlsx")
    ARCHIVO_PRUEBAS = os.path.join(directorio_actual, "pruebas.xlsx")
    NOMBRE_HOJA = "IR Julio 2025"  # Nombre exacto de la hoja a copiar
    
    # Verificar que existan los archivos
    if not os.path.exists(ARCHIVO_PLANTILLA):
        print(f"Error: No se encontró el archivo plantilla en {ARCHIVO_PLANTILLA}")
    elif not os.path.exists(ARCHIVO_PRUEBAS):
        print(f"Error: No se encontró el archivo de pruebas en {ARCHIVO_PRUEBAS}")
    else:
        copiar_hoja_con_graficos(ARCHIVO_PLANTILLA, ARCHIVO_PRUEBAS, NOMBRE_HOJA)
