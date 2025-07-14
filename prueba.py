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
import os
from pathlib import Path
import time

# 1. Configuración INMEDIATA del logging (primer mensaje en <1 segundo)
def configurar_logging_rapido():
    """Configuración mínima inicial para feedback inmediato"""
    try:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[logging.StreamHandler()]
        )
        logger = logging.getLogger('TornosLogger')
        logger.info("Iniciando aplicación...")  # Primer mensaje visible inmediatamente
        return logger
    except Exception:
        logging.basicConfig(level=logging.INFO)
        return logging.getLogger('TornosLogger')

logger = configurar_logging_rapido()

# 2. Carga DIFERIDA de módulos pesados
def cargar_dependencias():
    """Importa los módulos pesados solo cuando se necesiten"""
    import pythoncom
    import win32com.client
    import pandas as pd
    from logging.handlers import RotatingFileHandler
    import shutil
    return pythoncom, win32com.client, pd, RotatingFileHandler, shutil

# 3. Configuración COMPLETA del logging (se ejecuta después del inicio rápido)
def configurar_logging_completo(base_path):
    """Configura el sistema de logging completo con rotación"""
    try:
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
        logger.info(f"Log completo configurado en: {log_path}")
    except Exception as e:
        logger.error(f"No se pudo configurar log completo: {str(e)}")

# 4. Funciones principales
def verificar_archivo_no_bloqueado(ruta_archivo):
    try:
        with open(ruta_archivo, 'a+b'):
            pass
        return True
    except Exception as e:
        logger.error(f"Archivo bloqueado/inaccesible: {ruta_archivo} - {str(e)}")
        return False

def encontrar_archivo_odc(base_path):
    patrones = [
        "*CLNALMISOTPRD*Peeling*Production*.odc",
        "*rwExport*Peeling*Production*.odc",
        "*Peeling*Production*.odc",
        "*.odc"
    ]
    
    for patron in patrones:
        try:
            for archivo in base_path.glob(patron):
                if verificar_archivo_no_bloqueado(archivo):
                    return archivo
        except Exception as e:
            logger.warning(f"Error buscando {patron}: {str(e)}")
    
    logger.error(f"No se encontró archivo ODC. Directorio: {base_path}")
    return None

def exportar_desde_odc():
    """Función principal optimizada"""
    # Carga diferida de dependencias
    pythoncom, win32com, pd, _, _ = cargar_dependencias()
    
    # Obtener ruta base
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    
    # Configurar logging completo
    configurar_logging_completo(base_path)
    
    excel = None
    workbook = None
    
    try:
        # Buscar archivo ODC
        odc_path = encontrar_archivo_odc(base_path)
        if not odc_path:
            raise FileNotFoundError("Archivo ODC no encontrado")
        
        # Iniciar Excel optimizado
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        # Abrir archivo
        workbook = excel.Workbooks.Open(
            str(odc_path.absolute()),
            UpdateLinks=0,
            ReadOnly=True
        )
        
        # Espera optimizada (15 segundos máximo)
        timeout = time.time() + 15
        while time.time() < timeout:
            try:
                if workbook.Application.Ready:
                    break
                time.sleep(1)
            except:
                time.sleep(1)
        
        # Guardar como Excel
        output_path = base_path / "datos_actualizados.xlsx"
        workbook.SaveAs(str(output_path), FileFormat=51)
        
        # Procesar datos
        datos = pd.read_excel(output_path)
        
        # Generar reporte de rendimiento
        if not datos.empty:
            datos['Fecha'] = pd.to_datetime(datos['Fecha'])
            datos = datos.sort_values('Fecha', ascending=False)
            
            logger.info("\n=== RESUMEN DE RENDIMIENTO ===")
            for fecha in datos['Fecha'].unique()[:5]:  # Últimos 5 días
                for torno in [3011, 3012]:
                    filtro = (datos['Fecha'] == fecha) & (datos['WorkId'] == torno)
                    if any(filtro):
                        rend = datos.loc[filtro, 'Rendimiento'].values[0]
                        acum = datos.loc[filtro, 'Rendimiento_Acumulado'].values[0]
                        logger.info(
                            f"{fecha.strftime('%Y-%m-%d')} - Torno {torno-3010}: "
                            f"Rendimiento: {rend:.2f} | Acumulado: {acum:.2f}"
                        )
            logger.info("="*40)
        
        return datos
        
    except Exception as e:
        logger.error(f"ERROR: {str(e)}", exc_info=True)
        return None
        
    finally:
        try:
            if workbook: workbook.Close(False)
            if excel: excel.Quit()
            pythoncom.CoUninitialize()
        except: pass

# Punto de entrada principal
if __name__ == "__main__":
    try:
        logger.info("=== INICIANDO PROCESO ===")
        datos = exportar_desde_odc()
        
        if datos is not None:
            logger.info(f"Proceso completado. Registros procesados: {len(datos)}")
        else:
            logger.error("El proceso no se completó correctamente")
            
    except Exception as e:
        logger.critical(f"Error crítico: {str(e)}", exc_info=True)
        
    finally:
        if not getattr(sys, 'frozen', False):
            input("Presione Enter para salir...")
        logger.info("=== FINALIZADO ===")
