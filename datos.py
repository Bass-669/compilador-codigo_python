import logging
import sys
import time
from pathlib import Path

# configuracion de inicio
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger('TornosLogger')
logger.info("Iniciando proceso...")

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

def configurar_log_completo():
    # Configuracion del documento log
    try:
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        log_path = base_path / "tornos.log"
        file_handler = logging.FileHandler(str(log_path), encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        ))
        logger.addHandler(file_handler)
    except Exception as e:
        logger.error(f"No se pudo configurar archivo de log: {str(e)}")

configurar_log_completo()

def reintentos(func, max_intentos=3, espera=5, mensaje_reintento=None):
    # Sistema de reintentos en caso de fallos
    intento = 0
    while intento < max_intentos:
        intento += 1
        try:
            return func()
        except Exception as e:
            if "The file is locked for editing" in str(e) or "El archivo está bloqueado" in str(e):
                if intento < max_intentos:
                    msg = mensaje_reintento or f"Archivo bloqueado (intento {intento}/{max_intentos}). Reintentando..."
                    logger.info(msg)
                    time.sleep(espera)
                    continue
                else:
                    logger.error("No se pudo acceder al archivo después de varios intentos.\n")
            logger.error(f"Error en operación: {str(e)}", exc_info=True)
            return False

def procesar_archivo_odc():
    # Procesamiento del archivo odc
    max_intentos = 20
    espera_entre_intentos = 10
    intento = 0
    resultado = False

    while intento < max_intentos and not resultado:
        intento += 1
        excel = None
        workbook = None
        
        try:
            logger.info(f"Intento {intento}/{max_intentos}")
            
            # 1. Encontar el archivo odc
            base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
            nombre_archivo = "CLNALMISOTPRD rwExport report_Peeling_Production query.odc"
            odc_path = base_path / nombre_archivo
            
            if not odc_path.exists():
                raise FileNotFoundError(f"No se encontró el archivo ODC: {nombre_archivo}")

            # 2. Configuracion del Excel
            pythoncom.CoInitialize()
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            
            # 3. Abrir archivo
            logger.info("Abriendo archivo ODC...")
            workbook = excel.Workbooks.Open(
                str(odc_path.absolute()),
                UpdateLinks=0,
                ReadOnly=True
            )
            datos_cargados = False
            start_time = time.time()
            while (time.time() - start_time) < 15:
                try:
                    if workbook.Application.Ready:
                        logger.info("Datos cargados correctamente")
                        datos_cargados = True
                        break
                    time.sleep(1)
                except:
                    time.sleep(1)
            if not datos_cargados:
                logger.warning("Tiempo de espera agotado, continuando...")

            # 5. Guardar datos
            output_path = odc_path.parent / "datos_actualizados.xlsx"
            if output_path.exists():
                logger.warning("El archivo de salida ya existe, se sobrescribirá")
            try:
                workbook.SaveAs(
                    str(output_path),
                    FileFormat=51,
                    ConflictResolution=2
                )
                logger.info(f"Datos exportados correctamente a: {output_path.name}")
                
                # 6. Generar reporte si no hay errores
                try:
                    datos = pd.read_excel(output_path)
                    
                    if not datos.empty:
                        datos['Fecha'] = pd.to_datetime(datos['Fecha'])
                        datos = datos.sort_values('Fecha', ascending=False)
                        fechas_unicas = datos['Fecha'].unique()
                        ultimas_fechas = fechas_unicas[:15]
                        ultima_fecha = datos['Fecha'].max()
                        
                        mensaje = "\n=== RESUMEN DE DATOS ===\n"
                        for fecha in ultimas_fechas:
                            if fecha == ultima_fecha:
                                continue
                                
                            datos_fecha = datos[datos['Fecha'] == fecha]
                            mensaje += f"\nFecha: {fecha.strftime('%Y-%m-%d')}\n"
                            
                            # Torno 1
                            torno1 = datos_fecha[datos_fecha['WorkId'] == 3011]
                            if not torno1.empty:
                                mensaje += f"Torno 1: Rendimiento: {torno1.iloc[0].get('Rendimiento', 0):.2f} | Acumulado: {torno1.iloc[0].get('Rendimiento_Acumulado', 0):.2f}\n"
                            else:
                                mensaje += "Torno 1: Sin datos\n"
                            
                            # Torno 2
                            torno2 = datos_fecha[datos_fecha['WorkId'] == 3012]
                            if not torno2.empty:
                                mensaje += f"Torno 2: Rendimiento: {torno2.iloc[0].get('Rendimiento', 0):.2f} | Acumulado: {torno2.iloc[0].get('Rendimiento_Acumulado', 0):.2f}\n"
                            else:
                                mensaje += "Torno 2: Sin datos\n"
                        
                        logger.info(mensaje)
                # En caso de errores
                except Exception as e:
                    logger.error(f"Error generando reporte: {str(e)}")
                
                resultado = True
                
            except Exception as e:
                if "locked" in str(e).lower() or "bloqueado" in str(e).lower() or "acceso" in str(e).lower():
                    logger.warning(f"Archivo bloqueado durante guardado (intento {intento}/{max_intentos})")
                    raise
                logger.error(f"Error al guardar: {str(e)}")
                resultado = False
                
        except Exception as e:
            if "locked" in str(e).lower() or "bloqueado" in str(e).lower() or "acceso" in str(e).lower():
                if intento < max_intentos:
                    logger.info(f"Archivo bloqueado detectado. Reintentando en {espera_entre_intentos} segundos...")
                    time.sleep(espera_entre_intentos)
                else:
                    logger.error("Máximo de intentos alcanzado. El archivo sigue bloqueado.")
            else:
                logger.error(f"Error inesperado: {str(e)}", exc_info=True)
                resultado = False
                
        finally:

            try:
                if workbook is not None:
                    workbook.Close(False)
                if excel is not None:
                    excel.Quit()
                pythoncom.CoUninitialize()
            except Exception as e:
                logger.warning(f"Error limpiando recursos Excel: {str(e)}")

    return resultado

# Ejecucion del codigo
if __name__ == "__main__":
    logger.info("=== INICIO DEL PROCESO ===")
    
    if procesar_archivo_odc():
        logger.info("=== FIN DEL PROCESO ===\n")
    else:
        logger.error("Proceso completado con ERRORES")
