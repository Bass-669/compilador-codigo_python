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
    except Exception as e:
        logger.error(f"No se pudo configurar archivo de log: {str(e)}")

configurar_log_completo()

## ---------------------------------------------------------------
## 4. FUNCIONES PRINCIPALES (CORREGIDAS)
## ---------------------------------------------------------------

def procesar_archivo_odc():
    """Procesamiento completo corregido con manejo adecuado de excepciones"""
    excel = None
    workbook = None
    
    try:
        # 1. LOCALIZAR ARCHIVO
        base_path = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
        nombre_archivo = "CLNALMISOTPRD rwExport report_Peeling_Production query.odc"
        odc_path = base_path / nombre_archivo
        
        if not odc_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo ODC en {base_path}: {nombre_archivo}")

        # 2. CONFIGURAR EXCEL
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        
        # 3. ABRIR ARCHIVO
        logger.info("Abriendo archivo ODC...")
        workbook = excel.Workbooks.Open(
            str(odc_path.absolute()),
            UpdateLinks=0,
            ReadOnly=True
        )
        
        # 4. ESPERA CON CONTROL
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
            str(output_path),
            FileFormat=51,
            ConflictResolution=2
        )
        logger.info(f"Datos exportados a: {output_path.name}")
        
        # 6. GENERAR REPORTE
        datos = pd.read_excel(output_path)
        
        if not datos.empty:
            
            try:
                datos['Fecha'] = pd.to_datetime(datos['Fecha'])
                datos = datos.sort_values('Fecha', ascending=False)
                ultimas_5_fechas = datos['Fecha'].unique()[:15]
                mensaje = "=== RESUMEN DE DATOS ===\n"
                
                for fecha in ultimas_5_fechas:
                    datos_fecha = datos[datos['Fecha'] == fecha]
                    
                    # Torno 1
                    torno1 = datos_fecha[datos_fecha['WorkId'] == 3011]
                    if not torno1.empty:
                        mensaje += ("\n"
                            f"Fecha: {fecha.strftime('%Y-%m-%d')} Torno 1: Rendimiento: {torno1.iloc[0].get('Rendimiento', 0):.2f} | "
                            f"Acumulado: {torno1.iloc[0].get('Rendimiento_Acumulado', 0):.2f}\n"
                        )

                    else:
                        mensaje += ("Torno 1: Sin datos\n")

                    # Torno 2
                    torno2 = datos_fecha[datos_fecha['WorkId'] == 3012]
                    if not torno2.empty:
                        mensaje += (
                            f"Fecha: {fecha.strftime('%Y-%m-%d')} Torno 2: Rendimiento: {torno2.iloc[0].get('Rendimiento', 0):.2f} | "
                            f"Acumulado: {torno2.iloc[0].get('Rendimiento_Acumulado', 0):.2f}\n"
                        )

                    else:
                        mensaje += ("Torno 2: Sin datos\n")

                logger.info(mensaje)
                
            except Exception as e:
                logger.error(f"Error procesando fechas: {str(e)}")
        
        return True
        
    except Exception as e:
        logger.error(f"Error en procesar_archivo_odc: {str(e)}", exc_info=True)
        return False
        
    finally:
        # LIMPIEZA DE RECURSOS
        try:
            if workbook: 
                workbook.Close(False)
            if excel: 
                excel.Quit()
            pythoncom.CoUninitialize()
        except Exception as e:
            logger.warning(f"Error al limpiar recursos: {str(e)}")
## ---------------------------------------------------------------
## 5. EJECUCIÓN PRINCIPAL
## ---------------------------------------------------------------

if __name__ == "__main__":
    logger.info("=== INICIO DEL PROCESO ===")
    
    if procesar_archivo_odc():
        logger.info("=== FIN DEL PROCESO ===\n")
    else:
        logger.error("Proceso completado con ERRORES")
