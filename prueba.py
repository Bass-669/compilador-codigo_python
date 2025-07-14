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

## ---------------------------------------------------------------
## 1. CONFIGURACIÓN INICIAL
## ---------------------------------------------------------------

# Configuración rápida inicial
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger('TornosLogger')
logger.info("Iniciando proceso...")

## ---------------------------------------------------------------
## 2. CARGA SEGURA DE DEPENDENCIAS
## ---------------------------------------------------------------

try:
    import pythoncom
    import win32com.client
    import pandas as pd
    from logging.handlers import RotatingFileHandler
except ImportError as e:
    logger.error(f"Error crítico: Falta dependencia - {str(e)}")
    if not getattr(sys, 'frozen', False):
        input("Presione Enter para salir...")
    sys.exit(1)

## ---------------------------------------------------------------
## 3. CONFIGURACIÓN COMPLETA DEL LOG
## ---------------------------------------------------------------

def configurar_log_completo():
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
        logger.info(f"Log configurado en: {log_path}")
    except Exception as e:
        logger.error(f"No se pudo configurar archivo de log: {str(e)}")

configurar_log_completo()

## ---------------------------------------------------------------
## 4. FUNCIONES PRINCIPALES (CORREGIDAS)
## ---------------------------------------------------------------

def encontrar_archivo_odc_especifico():
    """Búsqueda específica para el archivo ODC"""
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    
    patrones_exactos = [
        "CLNALMISOTPRD rwExport report_Peeling_Production query.odc",
        "CLNALMISOTPRD*.odc",
        "*report_Peeling_Production*.odc",
        "*.odc"
    ]
    
    for patron in patrones_exactos:
        try:
            for archivo in base_path.glob(patron):
                logger.info(f"Archivo encontrado: {archivo.name}")
                return archivo
        except Exception as e:
            logger.warning(f"Error buscando {patron}: {str(e)}")
    
    logger.error("No se encontró ningún archivo ODC")
    return None

def procesar_archivo_odc():
    """Procesamiento completo corregido"""
    excel = None
    workbook = None
    
    try:
        # 1. LOCALIZAR ARCHIVO
        odc_path = encontrar_archivo_odc_especifico()
        if not odc_path:
            raise FileNotFoundError("Archivo ODC no encontrado")
        
        logger.info(f"Procesando archivo: {odc_path.name}")
        
        # 2. CONFIGURAR EXCEL
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        # 3. ABRIR ARCHIVO (VERSIÓN CORREGIDA)
        logger.info("Abriendo archivo ODC...")
        workbook = excel.Workbooks.Open(
            str(odc_path.absolute()),  # CORRECCIÓN: Sin parámetro FileName
            UpdateLinks=0,
            ReadOnly=True
        )
        
        # 4. ESPERA CON CONTROL
        logger.info("Cargando datos...")
        start_time = time.time()
        while (time.time() - start_time) < 15:
            try:
                if workbook.Application.Ready:
                    logger.info("Datos cargados correctamente")
                    break
                time.sleep(1)
            except:
                time.sleep(1)
        else:
            logger.warning("Tiempo de espera agotado, continuando...")
        
        # 5. GUARDAR RESULTADOS
        output_path = odc_path.parent / "datos_actualizados.xlsx"
        if output_path.exists():
            logger.warning("El archivo de salida ya existe, se sobrescribirá")
        
        workbook.SaveAs(
            str(output_path),  # CORRECCIÓN: Sin parámetro FileName
            FileFormat=51,
            ConflictResolution=2
        )
        logger.info(f"Datos exportados a: {output_path.name}")
        
        # 6. GENERAR REPORTE
        datos = pd.read_excel(output_path)
        
        if not datos.empty:
            logger.info("\n=== RESUMEN DE DATOS ===")
            logger.info(f"Total registros: {len(datos)}")
            
            if 'Fecha' in datos.columns:
                try:
                    datos['Fecha'] = pd.to_datetime(datos['Fecha'])
                    datos.sort_values('Fecha', ascending=False, inplace=True)
                    
                    for fecha in datos['Fecha'].unique()[:5]:
                        logger.info(f"\nFecha: {fecha.strftime('%Y-%m-%d')}")
                        
                        for workid in [3011, 3012]:
                            filtro = (datos['Fecha'] == fecha) & (datos['WorkId'] == workid)
                            if any(filtro):
                                row = datos.loc[filtro].iloc[0]
                                logger.info(
                                    f"Torno {workid-3010}: "
                                    f"Rendimiento: {row.get('Rendimiento', 0):.2f} | "
                                    f"Acumulado: {row.get('Rendimiento_Acumulado', 0):.2f}"
                                )
                except Exception as e:
                    logger.error(f"Error procesando fechas: {str(e)}")
            
            logger.info("\n" + "="*40)
        
        return True
        
    except Exception as e:
        logger.error(f"ERROR: {str(e)}", exc_info=True)
        return False
        
    finally:
        try:
            if workbook: workbook.Close(False)
            if excel: excel.Quit()
            pythoncom.CoUninitialize()
        except: pass

## ---------------------------------------------------------------
## 5. EJECUCIÓN PRINCIPAL
## ---------------------------------------------------------------

if __name__ == "__main__":
    logger.info("=== INICIO DEL PROCESO ===")
    
    if procesar_archivo_odc():
        logger.info("Proceso completado con ÉXITO")
    else:
        logger.error("Proceso completado con ERRORES")
    
    if not getattr(sys, 'frozen', False):
        input("\nPresione Enter para salir...")
    
    logger.info("=== FIN DEL PROCESO ===\n")
