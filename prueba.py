# import logging
# import os
# import sys
# import time
# import win32com.client
# import pythoncom
# import pandas as pd
# from pathlib import Path
# from logging.handlers import RotatingFileHandler
# from datetime import datetime
# import shutil

# def configurar_logging():
#     """Configura el sistema de logging"""
#     try:
#         base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
#         log_path = base_path / "log.log"
        
#         logging.basicConfig(
#             level=logging.INFO,
#             format='%(asctime)s - %(levelname)s - %(message)s',
#             datefmt='%Y-%m-%d %H:%M:%S',
#             handlers=[
#                 RotatingFileHandler(
#                     str(log_path),
#                     maxBytes=5*1024*1024,
#                     backupCount=3,
#                     encoding='utf-8'
#                 ),
#                 logging.StreamHandler()
#             ]
#         )
        
#         logger = logging.getLogger('TornosLogger')
#         return logger
        
#     except Exception as e:
#         logging.basicConfig(level=logging.INFO)
#         logger = logging.getLogger('TornosLogger')
#         logger.error(f"Error configurando log: {str(e)}")
#         return logger

# def verificar_archivo_no_bloqueado(ruta_archivo):
#     """Verifica que el archivo no esté bloqueado"""
#     try:
#         with open(ruta_archivo, 'a+b'):
#             pass
#         return True
#     except PermissionError:
#         logger.error(f"Archivo bloqueado: {ruta_archivo}")
#         return False
#     except Exception as e:
#         logger.error(f"Error accediendo al archivo: {str(e)}")
#         return False

# def encontrar_archivo_odc():
#     """Busca el archivo ODC con flexibilidad en el nombre"""
#     base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
#     posibles_patrones = [
#         "*CLNALMISOTPRD*Peeling*Production*.odc",
#         "*rwExport*Peeling*Production*.odc",
#         "*Peeling*Production*.odc",
#         "*.odc"
#     ]
#     for patron in posibles_patrones:
#         archivos = list(base_path.glob(patron))
#         if archivos:
#             archivo = archivos[0]
#             if verificar_archivo_no_bloqueado(archivo):
#                 logger.info(f"Archivo encontrado y accesible")
#                 return archivo
#             else:
#                 logger.error(f"Archivo encontrado pero bloqueado: {archivo}")

#     logger.error(f"No se encontró archivo ODC accesible. Directorio: {base_path}")
#     logger.error(f"Archivos presentes: {[f.name for f in base_path.iterdir()]}")
#     return None

# def exportar_desde_odc():
#     """Exporta datos desde archivo ODC a Excel y registra últimos 5 días en log"""
#     excel = None
#     workbook = None
#     try:
#         logger.info("=== INICIANDO EXPORTACIÓN DESDE ODC === \n")
#         odc_path = encontrar_archivo_odc()
#         if not odc_path:
#             raise FileNotFoundError("No se pudo encontrar el archivo ODC accesible")
        
#         ruta_absoluta = str(odc_path.absolute())
#         pythoncom.CoInitialize()
#         excel = win32com.client.DispatchEx("Excel.Application")
#         excel.Visible = False
#         excel.DisplayAlerts = False
        
#         workbook = excel.Workbooks.Open(ruta_absoluta, UpdateLinks=0)
        
#         logger.info("Esperando carga de datos...")
#         tiempo_inicio = time.time()
#         while (time.time() - tiempo_inicio) < 10:
#             try:
#                 if workbook.ReadOnly:
#                     time.sleep(2)
#                     continue
#                 if workbook.Application.CalculationState == 0:
#                     break
#                 time.sleep(2)
#             except:
#                 time.sleep(2)
#         else:
#             raise TimeoutError("Tiempo de espera agotado para carga de datos")
        
#         output_path = odc_path.parent / "datos_actualizados.xlsx"
#         workbook.SaveAs(str(output_path.absolute()), FileFormat=51)
        
#         workbook.Close(False)
#         excel.Quit()
        
#         # Leer datos con pandas
#         datos = pd.read_excel(output_path)
        
#         # Procesar y registrar últimos 5 días
#         try:
#             # Convertir columna Fecha a datetime si no lo está
#             datos['Fecha'] = pd.to_datetime(datos['Fecha'])
            
#             # Ordenar por fecha descendente
#             datos = datos.sort_values('Fecha', ascending=False)
            
#             # Obtener las 5 fechas más recientes (únicas)
#             ultimas_5_fechas = datos['Fecha'].drop_duplicates().head(5)
            
#             # Filtrar solo los datos de esas fechas
#             datos_recientes = datos[datos['Fecha'].isin(ultimas_5_fechas)]
            
#             # Registrar en log
#             logger.info("=== ÚLTIMOS 5 DÍAS - RENDIMIENTO POR TORNO ===")
            
#             for fecha in ultimas_5_fechas:
#                 # Filtrar datos para esta fecha
#                 datos_fecha = datos_recientes[datos_recientes['Fecha'] == fecha]
                
#                 # Obtener datos para cada torno
#                 torno1 = datos_fecha[datos_fecha['WorkId'] == 3011]
#                 torno2 = datos_fecha[datos_fecha['WorkId'] == 3012]
                
#                 # Formatear mensaje
#                 mensaje = f"\n"
                
#                 if not torno1.empty:
#                     mensaje += (
#                         f" Fecha: {fecha.strftime('%Y-%m-%d')} TORNO 1 - Rendimiento: {torno1.iloc[0].get('Rendimiento', 'N/A')} | "
#                         f"Rendimiento Acumulado: {torno1.iloc[0].get('Rendimiento_Acumulado', 'N/A')}\n"
#                     )
#                 else:
#                     mensaje += "  TORNO 1 - Sin datos\n"
                
#                 if not torno2.empty:
#                     mensaje += (
#                         f" Fecha: {fecha.strftime('%Y-%m-%d')} TORNO 2 - Rendimiento: {torno2.iloc[0].get('Rendimiento', 'N/A')} | "
#                         f"Rendimiento Acumulado: {torno2.iloc[0].get('Rendimiento_Acumulado', 'N/A')}\n"
#                     )
#                 else:
#                     mensaje += "  TORNO 2 - Sin datos\n"
                
#                 logger.info(mensaje)
#         except Exception as e:
#             logger.error(f"Error al procesar últimos 5 días: {str(e)}")
        
#         return datos
    
#     except Exception as e:
#         logger.error(f"Error en exportar_desde_odc: {str(e)}", exc_info=True)
#         if workbook is not None:
#             try:
#                 workbook.Close(False)
#             except:
#                 pass
#         if excel is not None:
#             try:
#                 excel.Quit()
#             except:
#                 pass
#         return None
#     finally:
#         pythoncom.CoUninitialize()

# # Configuración inicial
# logger = configurar_logging()

# # Punto de entrada principal
# if __name__ == "__main__":
#     datos = exportar_desde_odc()
#     logger.info("=== FIN DE EJECUCIÓN === \n")



# -*- coding: utf-8 -*-
import logging
import sys
import time
from pathlib import Path

## ------------------------------------------------------------------
## 1. CONFIGURACIÓN INICIAL (INICIO RÁPIDO)
## ------------------------------------------------------------------

def configurar_logging_rapido():
    """Primera configuración para mostrar mensajes inmediatos"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(message)s',
        handlers=[logging.StreamHandler()]
    )
    return logging.getLogger('TornosLogger')

logger = configurar_logging_rapido()
logger.info("Iniciando proceso automatizado...")  # Mensaje inmediato

## ------------------------------------------------------------------
## 2. CARGA DIFERIDA DE DEPENDENCIAS (OPTIMIZACIÓN)
## ------------------------------------------------------------------

def cargar_dependencias():
    """Importa módulos pesados solo cuando se necesiten"""
    import pythoncom
    import win32com.client
    import pandas as pd
    from logging.handlers import RotatingFileHandler
    return pythoncom, win32com.client, pd, RotatingFileHandler

## ------------------------------------------------------------------
## 3. FUNCIONES PRINCIPALES (COMPROBADAS)
## ------------------------------------------------------------------

def configurar_logging_completo():
    """Configura el sistema de logging completo"""
    try:
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        log_path = base_path / "tornos.log"
        
        file_handler = RotatingFileHandler(
            str(log_path),
            maxBytes=5*1024*1024,
            backupCount=3,
            encoding='utf-8'
        )
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        logger.addHandler(file_handler)
    except Exception as e:
        logger.error(f"Error configurando log: {str(e)}")

def encontrar_archivo_odc():
    """Busca automáticamente el archivo ODC"""
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    patrones = ["*.odc", "*Production*.odc", "*Peeling*.odc"]
    
    for patron in patrones:
        try:
            for archivo in base_path.glob(patron):
                try:
                    with open(archivo, 'a+b'):
                        return archivo
                except:
                    continue
        except:
            continue
    
    logger.error("No se encontró archivo ODC")
    return None

def exportar_datos():
    """Proceso completo automatizado"""
    pythoncom, win32com, pd, _ = cargar_dependencias()
    configurar_logging_completo()
    
    excel = None
    workbook = None
    
    try:
        # 1. Buscar archivo
        odc_path = encontrar_archivo_odc()
        if not odc_path:
            raise FileNotFoundError()
        
        # 2. Configurar Excel
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        # 3. Procesar archivo
        workbook = excel.Workbooks.Open(str(odc_path), UpdateLinks=0, ReadOnly=True)
        
        # Espera máxima 15 segundos
        for _ in range(15):
            if workbook.Application.Ready:
                break
            time.sleep(1)
        
        # 4. Guardar resultados
        output_path = odc_path.parent / "datos_actualizados.xlsx"
        workbook.SaveAs(str(output_path), FileFormat=51)
        
        # 5. Generar reporte
        datos = pd.read_excel(output_path)
        
        if not datos.empty:
            datos['Fecha'] = pd.to_datetime(datos['Fecha'])
            datos.sort_values('Fecha', ascending=False, inplace=True)
            
            logger.info("\n=== RESUMEN AUTOMÁTICO ===")
            for fecha in datos['Fecha'].unique()[:5]:
                for torno in [3011, 3012]:
                    filtro = (datos['Fecha'] == fecha) & (datos['WorkId'] == torno)
                    if filtro.any():
                        row = datos.loc[filtro].iloc[0]
                        logger.info(
                            f"{fecha.date()} - Torno {torno-3010}: "
                            f"Rend: {row.get('Rendimiento', 'N/A')} | "
                            f"Acum: {row.get('Rendimiento_Acumulado', 'N/A')}"
                        )
        
        return True
        
    except Exception as e:
        logger.error(f"Error automático: {str(e)}")
        return False
        
    finally:
        try:
            if workbook: workbook.Close(False)
            if excel: excel.Quit()
            pythoncom.CoUninitialize()
        except: pass

## ------------------------------------------------------------------
## 4. EJECUCIÓN PRINCIPAL (AUTOMÁTICA)
## ------------------------------------------------------------------

if __name__ == "__main__":
    logger.info("Inicio del proceso automatizado")
    resultado = exportar_datos()
    
    if resultado:
        logger.info("Proceso completado exitosamente")
    else:
        logger.error("El proceso encontró errores")
    
    # Cierre automático en .exe (sin input())
    if not getattr(sys, 'frozen', False):
        input("Presione Enter para salir (solo en modo desarrollo)...")
