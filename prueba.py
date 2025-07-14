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
    """Configura el sistema de logging de manera robusta"""
    try:
        # Determinar el directorio base
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        # Posibles ubicaciones para el log (ordenadas por prioridad)
        posibles_rutas = [
            base_path / "log.log",  # Primero intentar en el directorio de la app
            Path.home() / "log.log",  # Luego en el home del usuario
            Path(tempfile.gettempdir()) / "log.log"  # Finalmente en temp
        ]
        logger = logging.getLogger('TornosLogger')
        logger.setLevel(logging.INFO)
        # Limpiar handlers existentes para evitar duplicados
        if logger.hasHandlers():
            logger.handlers.clear()
        
        formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        handler_exitoso = False
        # Probar distintas ubicaciones
        for ruta in posibles_rutas:
            try:
                ruta.parent.mkdir(parents=True, exist_ok=True)
                handler = RotatingFileHandler(
                    str(ruta),  # Convertir Path a string
                    maxBytes=5*1024*1024,  # 5MB
                    backupCount=3,
                    encoding='utf-8'
                )
                handler.setFormatter(formatter)
                logger.addHandler(handler)
                logger.info(f"Log configurado exitosamente en: {ruta}")
                handler_exitoso = True
                break  # Salir al encontrar la primera ubicación válida
            except Exception as e:
                continue
        # Si no se pudo crear archivo de log, usar consola
        if not handler_exitoso:
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)
            logger.warning("No se pudo crear archivo de log. Usando consola.")
        return logger
    except Exception as e:
        # Fallback básico si todo falla
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger('TornosLogger')
        logger.error(f"Error crítico configurando log: {str(e)}")
        return logger

logger = configurar_logging()

def verificar_archivo_no_bloqueado(ruta_archivo):
    """Verifica que el archivo no esté bloqueado (como en el código original)"""
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
                logger.info(f"Archivo encontrado y accesible: {archivo}")
                return archivo
            else:
                logger.error(f"Archivo encontrado pero bloqueado: {archivo}")

    logger.error(f"No se encontró archivo ODC accesible. Directorio: {base_path}")
    logger.error(f"Archivos presentes: {[f.name for f in base_path.iterdir()]}")
    return None

def exportar_desde_odc():
    """Exporta datos desde archivo ODC a Excel con manejo robusto"""
    excel = None
    workbook = None
    try:
        logger.info("=== INICIANDO EXPORTACIÓN DESDE ODC ===")
        # Buscar y verificar archivo ODC
        odc_path = encontrar_archivo_odc()
        if not odc_path:
            raise FileNotFoundError("No se pudo encontrar el archivo ODC accesible")
        
        ruta_absoluta = str(odc_path.absolute())
        logger.info(f"Ruta absoluta del archivo: {ruta_absoluta}")
        # Inicializar COM
        pythoncom.CoInitialize()
        # Configurar Excel
        logger.info("Iniciando Excel...")
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        # Abrir archivo ODC
        logger.info("Abriendo archivo ODC...")
        workbook = excel.Workbooks.Open(ruta_absoluta, UpdateLinks=0)
        # Esperar a que cargue
        logger.info("Esperando carga de datos...")
        tiempo_inicio = time.time()
        while (time.time() - tiempo_inicio) < 30:
            try:
                if workbook.ReadOnly:
                    time.sleep(2)
                    continue
                if workbook.Application.CalculationState == 0:  # xlDone
                    break
                time.sleep(2)
            except:
                time.sleep(2)
        else:
            raise TimeoutError("Tiempo de espera agotado para carga de datos")
        # Guardar como Excel (sin preguntar al usuario)
        output_path = odc_path.parent / "datos_actualizados.xlsx"
        # Crear copia de seguridad si el archivo ya existe
        if output_path.exists():
            # Nombre de backup con timestamp para evitar conflictos
            backup_path = output_path.parent / f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{output_path.name}"
            try:
                shutil.copy(output_path, backup_path)
                logger.info(f"Copia de seguridad creada: {backup_path}")
            except Exception as backup_error:
                logger.error(f"No se pudo crear copia de seguridad: {backup_error}")
                # Continuar de todos modos
        logger.info(f"Guardando como: {output_path}")
        workbook.SaveAs(str(output_path.absolute()), FileFormat=51)
        # Cerrar recursos
        workbook.Close(False)
        excel.Quit()
        # Leer datos con pandas
        logger.info("Leyendo datos exportados...")
        datos = pd.read_excel(output_path)
        logger.info(f"Datos obtenidos. Filas: {len(datos)}. Columnas: {list(datos.columns)}")
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

if __name__ == "__main__":
    logger.info("=== INICIO DE EJECUCIÓN ===")
    datos = exportar_desde_odc()
    logger.info("=== FIN DE EJECUCIÓN ===")
    
    if not getattr(sys, 'frozen', False):
        input("Presione Enter para salir...")
