import logging
import os
import sys
import time
import win32com.client
import pythoncom
import pandas as pd
from pathlib import Path
from logging.handlers import RotatingFileHandler
from datetime import datetime
import shutil

def configurar_logging():
    """Configura el sistema de logging"""
    try:
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        log_path = base_path / "log.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S',
            handlers=[
                RotatingFileHandler(
                    str(log_path),
                    maxBytes=5*1024*1024,
                    backupCount=3,
                    encoding='utf-8'
                ),
                logging.StreamHandler()
            ]
        )
        
        logger = logging.getLogger('TornosLogger')
        return logger
        
    except Exception as e:
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger('TornosLogger')
        logger.error(f"Error configurando log: {str(e)}")
        return logger

def verificar_archivo_no_bloqueado(ruta_archivo):
    """Verifica que el archivo no esté bloqueado"""
    try:
        with open(ruta_archivo, 'a+b'):
            pass
        return True
    except PermissionError:
        logger.error(f"Archivo bloqueado: {ruta_archivo}")
        return False
    except Exception as e:
        logger.error(f"Error accediendo al archivo: {str(e)}")
        return False

def encontrar_archivo_odc():
    """Busca el archivo ODC con flexibilidad en el nombre"""
    base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
    posibles_patrones = [
        "*CLNALMISOTPRD*Peeling*Production*.odc",
        "*rwExport*Peeling*Production*.odc",
        "*Peeling*Production*.odc",
        "*.odc"
    ]
    for patron in posibles_patrones:
        archivos = list(base_path.glob(patron))
        if archivos:
            archivo = archivos[0]
            if verificar_archivo_no_bloqueado(archivo):
                logger.info(f"Archivo encontrado y accesible")
                return archivo
            else:
                logger.error(f"Archivo encontrado pero bloqueado: {archivo}")

    logger.error(f"No se encontró archivo ODC accesible. Directorio: {base_path}")
    logger.error(f"Archivos presentes: {[f.name for f in base_path.iterdir()]}")
    return None

def exportar_desde_odc():
    """Exporta datos desde archivo ODC a Excel y registra últimos 5 días en log"""
    excel = None
    workbook = None
    try:
        logger.info("=== INICIANDO EXPORTACIÓN DESDE ODC ===")
        odc_path = encontrar_archivo_odc()
        if not odc_path:
            raise FileNotFoundError("No se pudo encontrar el archivo ODC accesible")
        
        ruta_absoluta = str(odc_path.absolute())
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        workbook = excel.Workbooks.Open(ruta_absoluta, UpdateLinks=0)
        
        logger.info("Esperando carga de datos...")
        tiempo_inicio = time.time()
        while (time.time() - tiempo_inicio) < 30:
            try:
                if workbook.ReadOnly:
                    time.sleep(2)
                    continue
                if workbook.Application.CalculationState == 0:
                    break
                time.sleep(2)
            except:
                time.sleep(2)
        else:
            raise TimeoutError("Tiempo de espera agotado para carga de datos")
        
        output_path = odc_path.parent / "datos_actualizados.xlsx"
        logger.info(f"Guardando como: {output_path}")
        workbook.SaveAs(str(output_path.absolute()), FileFormat=51)
        
        workbook.Close(False)
        excel.Quit()
        
        # Leer datos con pandas
        datos = pd.read_excel(output_path)
        
        # Procesar y registrar últimos 5 días
        try:
            # Convertir columna Fecha a datetime si no lo está
            datos['Fecha'] = pd.to_datetime(datos['Fecha'])
            
            # Ordenar por fecha descendente
            datos = datos.sort_values('Fecha', ascending=False)
            
            # Obtener las 5 fechas más recientes (únicas)
            ultimas_5_fechas = datos['Fecha'].drop_duplicates().head(5)
            
            # Filtrar solo los datos de esas fechas
            datos_recientes = datos[datos['Fecha'].isin(ultimas_5_fechas)]
            
            # Registrar en log
            logger.info("=== ÚLTIMOS 5 DÍAS - RENDIMIENTO POR TORNO ===")
            
            for fecha in ultimas_5_fechas:
                # Filtrar datos para esta fecha
                datos_fecha = datos_recientes[datos_recientes['Fecha'] == fecha]
                
                # Obtener datos para cada torno
                torno1 = datos_fecha[datos_fecha['WorkId'] == 3011]
                torno2 = datos_fecha[datos_fecha['WorkId'] == 3012]
                
                # Formatear mensaje
                mensaje = f"Fecha: {fecha.strftime('%Y-%m-%d')}\n"
                
                if not torno1.empty:
                    mensaje += (
                        f" Fecha: {fecha.strftime('%Y-%m-%d')} TORNO 1 - Rendimiento: {torno1.iloc[0].get('Rendimiento', 'N/A')} | "
                        f"Rendimiento Acumulado: {torno1.iloc[0].get('Rendimiento_Acumulado', 'N/A')}\n"
                    )
                else:
                    mensaje += "  TORNO 1 - Sin datos\n"
                
                if not torno2.empty:
                    mensaje += (
                        f" Fecha: {fecha.strftime('%Y-%m-%d')} TORNO 2 - Rendimiento: {torno2.iloc[0].get('Rendimiento', 'N/A')} | "
                        f"Rendimiento Acumulado: {torno2.iloc[0].get('Rendimiento_Acumulado', 'N/A')}"
                    )
                else:
                    mensaje += "  TORNO 2 - Sin datos"
                
                logger.info(mensaje)
        except Exception as e:
            logger.error(f"Error al procesar últimos 5 días: {str(e)}")
        
        return datos
    
    except Exception as e:
        logger.error(f"Error en exportar_desde_odc: {str(e)}", exc_info=True)
        if workbook is not None:
            try:
                workbook.Close(False)
            except:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except:
                pass
        return None
    finally:
        pythoncom.CoUninitialize()
        logger.info("=== FINALIZADO MANEJO DE EXCEL ===")

# Configuración inicial
logger = configurar_logging()

# Punto de entrada principal
if __name__ == "__main__":
    datos = exportar_desde_odc()
    logger.info("=== FIN DE EJECUCIÓN === \n")
