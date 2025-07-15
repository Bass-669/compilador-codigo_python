import openpyxl, re, shutil, time, os, sys, tkinter as tk
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from datetime import datetime
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
import threading
import tempfile
import logging
from logging.handlers import RotatingFileHandler

BASE_DIR = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
CARPETA, ARCHIVO = "reportes", "Reporte IR Tornos.xlsx"
RUTA_ENTRADA = os.path.join(BASE_DIR, CARPETA, ARCHIVO)
MESES = dict(zip(
    ["January","February","March","April","May","June","July","August","September","October","November","December"],
    ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
))
MESES_NUM = {
    "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4,
    "Mayo": 5, "Junio": 6, "Julio": 7, "Agosto": 8,
    "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12
}
BORDER = Border(*(Side(style='thin'),)*4)
ALIGN_R = Alignment(horizontal='right')
FILL_AMARILLO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def configurar_logging():
    """Configura un sistema de logging robusto con rotación de archivos"""
    posibles_rutas = [
        os.path.join(BASE_DIR, CARPETA, "log_tornos.log"),
        os.path.join(tempfile.gettempdir(), "log_tornos.log")
    ]
    # Configuración básica del logger
    logger = logging.getLogger('TornosLogger')
    logger.setLevel(logging.INFO)
    # Formato del log
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    # Probar distintas ubicaciones
    for ruta in posibles_rutas:
        try:
            os.makedirs(os.path.dirname(ruta), exist_ok=True)
            handler = RotatingFileHandler(
                ruta,
                maxBytes=5*1024*1024,
                backupCount=3,
                encoding='utf-8'
            )
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            return logger
        except Exception as e:
            escribir_log(f" \n No se pudo configurar log en {ruta}: {e}", file=sys.stderr)
    # Si fallan todas las rutas, crear logger de consola
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    logger.warning("No se pudo crear archivo de log. Usando consola.")
    return logger
logger = configurar_logging()

def escribir_log(mensaje, nivel="info"):
    """Función para escribir en el log de manera segura"""
    try:
        if nivel.lower() == "info":
            logger.info(mensaje)
        elif nivel.lower() == "warning":
            logger.warning(mensaje)
        elif nivel.lower() == "error":
            logger.error(mensaje)
        else:
            logger.debug(mensaje)
    except Exception as e:
        escribir_log(f"Error al escribir en log: {e}", file=sys.stderr)

def obtener_datos():
    datos = entrada_texto.get("1.0", tk.END).strip()
    if not datos: return messagebox.showwarning("Advertencia", "Ingresa los datos.")
    pedir_torno(lambda t: pedir_fecha(lambda m,d,a: iniciar(datos, t, m, d, a)))

def pedir_torno(callback):
    def confirmar():
        val = ent.get().strip()
        if not val:
            messagebox.showwarning("Advertencia", "Ingresa el número de torno.")
            return
        try:
            num_torno = int(val)
            if num_torno not in [1, 2]:
                messagebox.showwarning("Valor incorrecto", "El número de torno debe ser 1 o 2.")
                ent.delete(0, tk.END)
                ent.focus_set()
                return
            callback(num_torno)
            ventana.destroy()
        except ValueError:
            messagebox.showerror("Error", "Debe ingresar un número válido.")
            # Limpiar el campo y mantener la ventana abierta
            ent.delete(0, tk.END)
            ent.focus_set()
    ventana = tk.Toplevel()
    ventana.title("Número de torno")
    ventana.geometry("300x150")
    ventana.resizable(False, False)
    tk.Label(ventana, 
             text="Número de torno (1 o 2):",
             font=("Arial", 12)).pack(pady=10)
    ent = tk.Entry(ventana, font=("Arial", 12))
    ent.pack(pady=5)
    tk.Button(ventana, 
              text="Aceptar", 
              command=confirmar).pack(pady=5)
    ent.focus_set()
    ventana.grab_set()

def pedir_fecha(callback):
    ventana = tk.Toplevel()
    ventana.title("Fecha del reporte")
    ventana.geometry("300x200")
    ventana.resizable(False, False)
    tk.Label(ventana, text="Selecciona la fecha:").pack(pady=10)
    ent_fecha = DateEntry(ventana, date_pattern='dd/MM/yyyy')
    ent_fecha.pack(pady=10); ent_fecha.set_date(datetime.now())
    def confirmar():
        f = ent_fecha.get_date()
        callback(MESES[f.strftime("%B")], f.day, f.year)
        ventana.destroy()
    tk.Button(ventana, text="Aceptar", command=confirmar).pack(pady=10)
    ventana.grab_set()

def iniciar(texto, torno, mes, dia, anio):
    mostrar_carga()
    threading.Thread(target=lambda: ejecutar(texto, torno, mes, dia, anio), daemon=True).start()

def mostrar_carga():
    global ventana_carga, barra
    ventana_carga = tk.Toplevel()
    ventana_carga.title("Procesando...")
    ventana_carga.geometry("300x100")
    ventana_carga.resizable(False, False)
    tk.Label(ventana_carga, text="Procesando...", font=("Arial", 12)).pack(pady=10)
    barra = ttk.Progressbar(ventana_carga, mode='determinate', maximum=100)
    barra.pack(fill='x', padx=20, pady=5)
    ventana_carga.grab_set()

def cerrar_carga():
    if ventana_carga: ventana_carga.destroy()

def ejecutar(txt, torno, mes, dia, anio):
    escribir_log("Inicio de ejecutar")
    try:
        # Obtener rendimientos del log si existen
        fecha_actual = datetime(anio, MESES_NUM[mes], dia).date()
        rendimiento_log = obtener_rendimientos_de_log(fecha_actual)
        if rendimiento_log:
            escribir_log(f"Rendimientos encontrados en log: Torno 1: {rendimiento_log['torno1']}%, Torno 2: {rendimiento_log['torno2']}%")
        # Configuración inicial de la barra
        barra['value'] = 0
        ventana_carga.update_idletasks()
        # Función para incremento fluido de la barra
        def incrementar_barra(hasta, paso=1):
            actual = barra['value']
            for i in range(actual, hasta + 1, paso):
                barra['value'] = i
                ventana_carga.update_idletasks()
                time.sleep(0.01)
        # Paso 1: Preparar hoja del mes (0-25%)
        incrementar_barra(25)
        if not preparar_hoja_mes(mes, dia, anio):
            incrementar_barra(100)
            return
        # Paso 2: Procesar datos (25-75%)
        incrementar_barra(50)
        bloques, porcentajes = procesar_datos(txt, torno, mes, dia, anio)
        # Si hay error de permisos, terminamos
        if bloques is None or porcentajes is None:
            incrementar_barra(100)
            return False
        # Paso intermedio (50-75%)
        incrementar_barra(75)
        # Paso 3: Escribir en hoja mensual (75-100%)
        if bloques is not None and porcentajes is not None:
            fecha(mes, dia, anio, torno, bloques, porcentajes, incrementar_barra)
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error en ejecutar():\n{e}")
        escribir_log("Error", f"Ocurrió un error en ejecutar():\n{e}")
    finally:
        cerrar_carga()
        ventana.destroy()

def procesar_datos(entrada, torno, mes, dia, anio):
    """Procesa los datos y escribe en el archivo Excel con manejo de errores mejorado"""
    escribir_log(f"Inicio de procesar_datos - Torno: {torno}, Fecha: {dia}/{mes}/{anio}")
    bloques_detectados = []
    sumas_ad_por_bloque = []
    # 1. Verificación inicial del archivo
    if not os.path.exists(RUTA_ENTRADA):
        error_msg = f"No se encontró el archivo Excel en:\n{RUTA_ENTRADA}"
        messagebox.showerror("Error", error_msg)
        escribir_log("ERROR - Archivo no encontrado", nivel="error")
        return None, None
    # 2. Verificación de permisos de escritura
    try:
        # Intento de apertura para verificar permisos
        with open(RUTA_ENTRADA, 'a+b') as test_file:
            pass
    except PermissionError:
        error_msg = f"El archivo está abierto en Excel. Por favor cierre:\n{RUTA_ENTRADA}"
        messagebox.showerror("Error", error_msg)
        escribir_log("ERROR - El archivo esta abierto \n", nivel="error")
        return None, None
    except Exception as e:
        error_msg = f"No se puede acceder al archivo:\n{str(e)}"
        messagebox.showerror("Error", error_msg)
        escribir_log(f"ERROR - Acceso al archivo: {str(e)}", nivel="error")
        return None, None
    # 3. Procesamiento principal
    wb = None
    try:
        # Intentar abrir el workbook
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        # Verificar si existe la hoja "IR diario "
        if "IR diario " not in wb.sheetnames:
            error_msg = 'No se encontró la hoja "IR diario " en el archivo Excel'
            messagebox.showerror("Error", error_msg)
            escribir_log("ERROR - Hoja 'IR diario ' no encontrada", nivel="error")
            return None, None
        hoja = wb["IR diario "]
        ultima_fila = None
        # Buscar última fila con patrón "* * ..."
        for fila in hoja.iter_rows():
            if [str(c.value).strip() if c.value else "" for c in fila[:3]] == ["*", "*", "..."]:
                ultima_fila = fila[0].row
        if not ultima_fila:
            raise ValueError("No se encontró '* * ...'")
        fila = ultima_fila + 1
        for b in extraer_bloques(entrada):
            try:
                f_ini = fila
                subs = sub_bloques(b)
                filas_validas = []
                # Procesar cada subbloque
                for sub in subs:
                    txt = sub[0] if not re.match(r'^\d', sub[0]) else ""
                    datos = sub[1:] if txt else sub
                    # Construir datos de columnas
                    p = txt.split()
                    col_txt = (
                        [p[0], p[1], p[2], p[3], "", p[4]] if "*" in txt and len(p) >= 5 and p[0] == "*" else
                        ["*", "*", "...", "", "", ""] if "*" in txt else
                        [p[0], p[1], p[2], p[3], "", p[4]] if len(p) >= 5 else
                        ["", p[0], p[1], p[2], "", p[3]] if len(p) == 4 else
                        [""] * 6
                    )
                    col_nums = [val for l in datos for val in l.strip().split()]
                    fila_vals = col_txt + col_nums
                    # Escribir valores en las celdas (columnas 1-24)
                    for col, val in enumerate(fila_vals[:24], 1):
                        try:
                            n = float(val.replace(",", ".")) if 3 <= col <= 24 and val else val
                            escribir(hoja, fila, col, n, isinstance(n, float))
                        except:
                            escribir(hoja, fila, col, val)
                    # Escribir metadatos (columnas 25-28)
                    for col, val in zip(range(25, 29), [torno, mes, dia, anio]):
                        hoja.cell(row=fila, column=col, value=val).alignment = ALIGN_R
                    fila += 1
                f_fin = fila - 1
                tipo_bloque = "PODADO" if "PODADO" in txt.upper() else "REGULAR"
                bloques_detectados.append((tipo_bloque, f_fin))
                # Insertar fórmulas proporcionales (columna AD)
                if len(subs) > 1:
                    for f in range(f_ini, f_fin):
                        hoja.cell(row=f, column=30, value=f"=IFERROR(AC{f}*D{f}/D{f_fin}, 0)")
                # Configurar celda de autosuma
                for col in range(25, 30):
                    hoja.cell(row=f_fin, column=col, value="")
                celda_autosuma = hoja.cell(row=f_fin, column=30)
                celda_autosuma.value = f"=SUM(AD{f_ini}:AD{f_fin-1})"
                celda_autosuma.fill = FILL_AMARILLO
                bloque_texto = " ".join(b).upper()
                tipo_bloque = "PODADO" if "PODADO" in bloque_texto else "REGULAR"
                valor_d = hoja.cell(row=f_fin, column=4).value
                try:
                    valor_d = float(str(valor_d).replace(",", ".")) if valor_d else 0
                except:
                    valor_d = 0
                bloques_detectados.append((tipo_bloque, valor_d))
                
                if tipo_bloque != "PODADO":
                    for col in range(25, 30):
                        hoja.cell(row=f_fin, column=col, value="")
                try:
                    valor_ae = Pasar_referencia(f"AD{f_fin}")
                    sumas_ad_por_bloque.append(valor_ae)
                except Exception as e:
                    sumas_ad_por_bloque.append(0.0)
                    escribir_log(f"Error al obtener referencia AD{f_fin}: {str(e)}", nivel="warning")
                # Guardar cambios después de cada bloque
                try:
                    wb.save(RUTA_ENTRADA)
                except PermissionError:
                    error_msg = "El archivo Excel fue bloqueado durante la ejecución. Por favor ciérrelo."
                    messagebox.showerror("Error", error_msg)
                    escribir_log("ERROR - Archivo bloqueado durante la ejecución", nivel="error")
                    return None, None
            except Exception as e:
                escribir_log(f"Error procesando bloque: {str(e)}", nivel="error")
                continue
        # Copia de seguridad
        try:
            backup_path = os.path.join(BASE_DIR, CARPETA, "Reporte IR Tornos copia_de_seguridad.xlsx")
            shutil.copy(RUTA_ENTRADA, backup_path)
            escribir_log(f"Copia de seguridad creada")
        except Exception as e:
            escribir_log(f"No se pudo crear copia de seguridad: {str(e)}", nivel="warning")
        return bloques_detectados, sumas_ad_por_bloque
    except PermissionError:
        error_msg = "El archivo Excel fue abierto durante la ejecución. Operación cancelada."
        messagebox.showerror("Error", error_msg)
        escribir_log("ERROR - Archivo abierto durante la ejecución", nivel="error")
        return None, None
    except Exception as e:
        error_msg = f"Error inesperado: {str(e)}"
        messagebox.showerror("Error", error_msg)
        escribir_log(f"ERROR - {str(e)}", nivel="error")
        return None, None
    finally:
        if wb is not None:
            try:
                wb.close()
            except:
                pass
        escribir_log("Procesamiento de datos completado")

def escribir(hoja, fila, col, valor, es_numero=False):
    """Escribe un valor en la celda con formato adecuado"""
    celda = hoja.cell(row=fila, column=col, value=valor)
    celda.border = BORDER
    celda.alignment = ALIGN_R
    if es_numero:
        celda.number_format = '0.00'

def Pasar_referencia(celda_origen):
    """Retorna la referencia CORRECTAMENTE formateada"""
    escribir_log("Inicio de pasar_referencia")
    if not re.match(r'^AD\d+$', celda_origen):
        messagebox.showerror("Error", f"Formato de celda inválido: {celda_origen}")
        escribir_log("Error", f"Formato de celda inválido: {celda_origen}")
        raise ValueError(f"Formato de celda inválido: {celda_origen}")
    # FORMATO REQUERIDO POR EXCEL
    referencia = f"='IR diario '!{celda_origen}"
    return referencia

def extraer_bloques(txt):
    escribir_log("Inicio de extraer_bloques")
    lineas = [l.strip() for l in txt.strip().split("\n") if l.strip()]
    bloques, b, i = [], [], 0
    while i < len(lineas):
        l = lineas[i]
        if re.match(r'^\* \* \.\.\.', l):
            b.append(l)
            if i+1 < len(lineas) and re.match(r'^\d', lineas[i+1]):
                b.append(lineas[i+1])
                i += 1
            bloques.append(b)
            b = []
        else: b.append(l)
        i += 1
    if b: bloques.append(b)
    return bloques

def sub_bloques(b):
    escribir_log("Inicio de sub_bloques")
    subs, tmp = [], []
    for l in b:
        if re.match(r'^\D', l) or '*' in l:
            if tmp: subs.append(tmp)
            tmp = [l]
        else: tmp.append(l)
    if tmp: subs.append(tmp)
    return subs

def escribir_valor_bloque(hoja, col_dia, torno, valor, tipo_bloque):
    escribir_log("Inicio de escribir_valor_bloque")
    tipo_bloque = tipo_bloque.strip().upper()
    if tipo_bloque == "PODADO":
        fila_valor = 3 if torno == 1 else 4
    elif tipo_bloque == "REGULAR":
        fila_valor = 8 if torno == 1 else 9
    else:
        messagebox.showwarning("Advertencia", f"Tipo de bloque no reconocido: '{tipo_bloque}'")
        escribir_log("Advertencia", f"Tipo de bloque no reconocido: '{tipo_bloque}'")
        return
    try:
        if valor is None:
            valor_final = 0.0
        elif isinstance(valor, (int, float)):
            valor_final = float(valor)
        else:
            valor_final = float(str(valor).replace(",", "."))
    except ValueError:
        valor_final = 0.0
    celda = hoja.cell(row=fila_valor, column=col_dia) # Escribir el valor en la celda
    celda.value = valor_final
    celda.number_format = '0'

def escribir_valores_resumen_bloques(hoja, col_dia, torno, valores_ae_por_bloque, tipos_bloque, rendimiento_log):
    """Escribe referencias y fórmulas en las celdas correspondientes"""
    escribir_log("Inicio de escribir_valores_resumen_bloques")
    letra_actual = openpyxl.utils.get_column_letter(col_dia)
    try:
        # 1. Escribir fórmulas para los cálculos
        hoja.cell(row=34, column=col_dia, 
                 value=f"=IFERROR(AVERAGE({letra_actual}32:{letra_actual}33), 0)").number_format = '0.00%'
        hoja.cell(row=38, column=col_dia, 
                 value=f"=IFERROR({letra_actual}32/{letra_actual}23, 0)").font = Font(color='000000')
        hoja.cell(row=39, column=col_dia, 
                 value=f"=IFERROR({letra_actual}33/{letra_actual}24, 0)").font = Font(color='000000')
        hoja.cell(row=40, column=col_dia, 
                 value=f"=IFERROR({letra_actual}34/{letra_actual}28, 0)").number_format = '0.00%'
        hoja.cell(row=40, column=col_dia).font = Font(color='000000')
        # 2. Escribir referencias de bloques
        for i, (tipo_bloque, referencia) in enumerate(zip(tipos_bloque, valores_ae_por_bloque)):
            tipo_bloque = tipo_bloque.strip().upper()
            if tipo_bloque == "PODADO":
                fila_valor = 13 if torno == 1 else 14
            elif tipo_bloque == "REGULAR":
                fila_valor = 18 if torno == 1 else 19
            else:
                continue
            celda = hoja.cell(row=fila_valor, column=col_dia)
            celda.value = referencia
            celda.alignment = ALIGN_R
            escribir_log(f"Bloque {i} ({tipo_bloque}) | Torno {torno} | Fila {fila_valor}")
            escribir_log(f"Referencia escrita: {referencia}")
         # Escribir rendimientos del log si existen
        if rendimiento_log != None:
            filas_rendimiento = {
                1: 32,  # Torno 1 - fila de rendimiento
                2: 33   # Torno 2 - fila de rendimiento
            }
            for torno_num, fila in filas_rendimiento.items():
                hoja.cell(
                    row=fila, 
                    column=col_dia, 
                    value=rendimiento_log[f'torno{torno_num}']/100
                ).number_format = '0.00%'
                escribir_log(f"Rendimiento del Torno {torno_num} ({rendimiento_log[f'torno{torno_num}']}%) escrito en {fila},{col_dia}")
    except Exception as e:
        escribir_log(f"Error en escribir_valores_resumen_bloques: {str(e)}", nivel="error")
        raise

def fecha(mes, dia, anio, torno, bloques_detectados, sumas_ad_por_bloque, incrementar_barra, rendimiento_log=None):
    escribir_log("Inicio de fecha")
    """Escribe los datos en la hoja del mes incluyendo las fechas"""
    nombre_hoja = f"IR {mes} {anio}"
    col_dia = dia + 1
    wb = None
    exito = False
    try:
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        if nombre_hoja not in wb.sheetnames:
            messagebox.showerror("Error", f"No se encontró la hoja '{nombre_hoja}'")
            escribir_log(f"Error - Hoja '{nombre_hoja}' no encontrada", nivel="error")
            return False
        hoja_mes = wb[nombre_hoja]
        nueva_fecha = f"{dia:02d}/{MESES_NUM[mes]:02d}/{anio}"
        # 2. Escribir la fecha en las celdas correspondientes
        filas_fecha = [2, 7, 12, 17, 22, 27, 31, 37]
        for fila in filas_fecha:
            try:
                celda = hoja_mes.cell(row=fila, column=col_dia)
                celda.value = nueva_fecha
            except Exception as e:
                messagebox.showwarning("Advertencia", f"Error escribiendo fecha en {fila},{col_dia}: {str(e)}")
                escribir_log("Advertencia", f"Error escribiendo fecha en {fila},{col_dia}: {str(e)}")
        # 3. Escribir valores de bloques
        valores_para_escribir = [val for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        tipos_para_escribir = [tipo for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        for (tipo_bloque, valor), valor_ae in zip(zip(tipos_para_escribir, valores_para_escribir), sumas_ad_por_bloque):
            escribir_valor_bloque(hoja_mes, col_dia, torno, valor, tipo_bloque)
            escribir_valores_resumen_bloques(hoja_mes, col_dia, torno, [valor_ae], [tipo_bloque], rendimiento_log)
        # 4. Guardar cambios y crear copia
        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        wb.save(RUTA_ENTRADA)
        # 5. Copia de seguridad
        try:
            shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        except Exception as e:
            messagebox.showwarning("Advertencia", f"No se pudo crear copia de seguridad:\n{str(e)}")
        exito = True
        return exito
    except Exception as e:
        messagebox.showerror("Error", 
            f"No se pudo escribir en hoja:\n{str(e)}\n\n"
            "Verifique que el archivo no esté abierto\n"
                )
        error_msg = (
            f"El documento esta abierto cierrelo e intentelo nuevamente \n"
                )
        escribir_log(error_msg, "error")
        return False
    finally:
        # 6. Cerrar el workbook si está abierto
        if wb is not None:
            try:
                wb.close()
            except:
                messagebox.showwarning("Advertencia", "Error al cerrar el workbook")
                escribir_log("Advertencia", "Error al cerrar el workbook")
        # 7. Actualizar barra de progreso
        incrementar_barra(100)
        # 8. Mostrar mensaje de éxito solo si todo salió bien
        if exito:
            messagebox.showinfo("Éxito", "✅ Valores actualizados correctamente.")
            escribir_log(f"Éxito ✅ Valores actualizados correctamente")
            escribir_log(f"Fin de la ejecucucion \n")

def preparar_hoja_mes(mes, dia, anio):
    """Crea la hoja del mes si no existe y la configura con fórmulas iniciales."""
    escribir_log("Inicio de preparar_hoja_mes")
    nombre_hoja = f"IR {mes} {anio}"
    col_dia = dia + 1
    try:
        # Paso 1: Verificar si la hoja ya existe con openpyxl
        wb_check = openpyxl.load_workbook(RUTA_ENTRADA)
        if nombre_hoja in wb_check.sheetnames:
            hoja_existente = wb_check[nombre_hoja]
            celdas_clave = [
                hoja_existente.cell(row=3, column=col_dia).value,
                hoja_existente.cell(row=4, column=col_dia).value,
                hoja_existente.cell(row=8, column=col_dia).value,
                hoja_existente.cell(row=9, column=col_dia).value
            ]
            if any(cell is not None and str(cell).strip() != "" for cell in celdas_clave):
                escribir_log(f"El día {dia} ya tiene datos en {nombre_hoja}")
                wb_check.close()
                return True
            else:
                escribir_log(f"La hoja {nombre_hoja} ya existe y se usará tal cual.")
                wb_check.close()
                return True
        wb_check.close()
        # Paso 2: Crear hoja nueva si no existe
        import win32com.client as win32, pythoncom
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(RUTA_ENTRADA), UpdateLinks=0)
        hojas = [h.Name for h in wb.Sheets]
        if nombre_hoja not in hojas:
            # Buscar hoja anterior para copiar
            hojas_ir = [h for h in hojas if h.startswith("IR ") and len(h.split()) == 3]
            def total_meses(nombre):
                try:
                    _, mes_str, anio_str = nombre.split()
                    return int(anio_str) * 12 + MESES_NUM[mes_str]
                except:
                    return -1
            hojas_ordenadas = sorted(hojas_ir, key=total_meses)
            total_nueva = int(anio) * 12 + MESES_NUM[mes]
            hoja_anterior = None
            for h in hojas_ordenadas:
                if total_meses(h) < total_nueva:
                    hoja_anterior = h
                else:
                    break
            if not hoja_anterior:
                messagebox.showwarning("Orden inválido", f"No se encontró hoja anterior para '{nombre_hoja}'")
                wb.Close(SaveChanges=False)
                excel.Quit()
                pythoncom.CoUninitialize()
                return False
            idx_anterior = hojas.index(hoja_anterior)
            insert_idx = min(idx_anterior + 2, wb.Sheets.Count)
            wb.Sheets(hoja_anterior).Copy(After=wb.Sheets(insert_idx - 1))
            nueva_hoja = wb.ActiveSheet
            nueva_hoja.Name = nombre_hoja
            shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
            wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()
        # Paso 3: Configurar hoja con openpyxl
        wb2 = openpyxl.load_workbook(RUTA_ENTRADA)
        hoja = wb2[nombre_hoja]
        filas_a_limpiar = [2,3,4,7,8,9,12,13,14,17,18,19,22,23,24,27,28,31,32,33,34,37,38,39,40]
        for fila in filas_a_limpiar:
            for col in range(2, 35):
                celda = hoja.cell(row=fila, column=col)
                if not isinstance(celda, openpyxl.cell.cell.MergedCell):
                    celda.value = ""
        dias_mes = dias_en_mes(mes, anio)
        for col in range(2, 2 + dias_mes):
            dia_mes = col - 1
            fecha = f"{dia_mes:02d}/{MESES_NUM[mes]:02d}/{anio}"
            for fila in [2,7,12,17,22,27,31,37]:
                hoja.cell(row=fila, column=col, value=fecha)
        for col in range(2, 2 + dias_mes):
            letra = openpyxl.utils.get_column_letter(col)
            hoja.cell(row=23, column=col, value=f"=IFERROR(({letra}3*{letra}13+{letra}8*{letra}18)/({letra}3+{letra}8), 0)")
            hoja.cell(row=24, column=col, value=f"=IFERROR(({letra}4*{letra}14+{letra}9*{letra}19)/({letra}4+{letra}9), 0)")
            hoja.cell(row=28, column=col, value=f"=IFERROR(({letra}23*({letra}3+{letra}8)+{letra}24*({letra}4+{letra}9))/({letra}3+{letra}4+{letra}8+{letra}9), 0)")
        hoja.cell(row=32, column=34, value="R%").font = Font(bold=True)
        hoja.cell(row=38, column=34, value="IR%").font = Font(bold=True)
        hoja.cell(row=32, column=34).alignment = hoja.cell(row=38, column=34).alignment = Alignment(horizontal='center', vertical='center')
        for fila in [49,50,51]:
            for col_limpiar in [27,28,29,30]:
                hoja.cell(row=fila, column=col_limpiar, value=" ")
        hoja.cell(row=2, column=34, value=int(anio))
        for fila in [3,4,8,9]:
            hoja.cell(row=fila, column=34, value=f"=SUM(B{fila}:AG{fila})")
        hoja.cell(row=23, column=34, value="=(AH3*AH13+AH8*AH18)/(AH3+AH8)")
        hoja.cell(row=24, column=34, value="=(AH4*AH14+AH9*AH19)/(AH4+AH9)")
        hoja.cell(row=28, column=34, value="=(AH23*(AH3+AH8)+AH24*(AH4+AH9))/(AH3+AH4+AH8+AH9)")
        hoja.cell(row=39, column=34, value="=AH33/AH28").number_format = '0.00%'
        hoja.cell(row=40, column=34, value="=AH34/AH28").number_format = '0.00%'
        wb2.save(RUTA_ENTRADA)
        wb2.close()
        # Recalcular fórmulas
        pythoncom.CoInitialize()
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb_calc = excel.Workbooks.Open(os.path.abspath(RUTA_ENTRADA), UpdateLinks=0)
        excel.CalculateFull()
        wb_calc.Save()
        wb_calc.Close()
        excel.Quit()
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        messagebox.showerror("Error crítico", f"No se pudo completar la operación:\n{str(e)}")
        return False

def dias_en_mes(mes, anio):
    escribir_log("Inicio de dias_en_mes")
    """Devuelve el número de días en un mes, considerando años bisiestos para febrero"""
    if mes == "Febrero":
        # Año bisiesto si es divisible por 4, pero no por 100, a menos que también sea divisible por 400
        if (anio % 4 == 0 and anio % 100 != 0) or (anio % 400 == 0):
            return 29
        return 28
    meses_31_dias = ["Enero", "Marzo", "Mayo", "Julio", "Agosto", "Octubre", "Diciembre"]
    return 31 if mes in meses_31_dias else 30

def obtener_rendimientos_de_log(fecha_ingresada):
    escribir_log("Inicio de obtener_rendimientos_de_log")
    log_path = os.path.join(BASE_DIR, "tornos.log")
    fecha_str = fecha_ingresada.strftime("%Y-%m-%d")
    rendimientos = {'torno1': None, 'torno2': None}
    if not os.path.exists(log_path):
        escribir_log(f"Archivo de log no encontrado: {log_path}", nivel="warning")
        return None
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            lineas = f.readlines()
        # Buscar las líneas que contienen la fecha
        patron = re.compile(
            r"Fecha: " + re.escape(fecha_str) + 
            r" Torno (\d): Rendimiento: (\d+\.\d+)"
        )
        for linea in reversed(lineas[-20:]):  # Buscar en las últimas 20 líneas
            coincidencia = patron.search(linea)
            if coincidencia:
                torno = coincidencia.group(1)
                rendimiento = float(coincidencia.group(2))
                rendimientos[f'torno{torno}'] = rendimiento
        # Verificar que tengamos ambos tornos
        if None in rendimientos.values():
            return None
        return rendimientos
    except Exception as e:
        escribir_log(f"Error al leer el archivo de log: {str(e)}", nivel="error")
        return None

ventana = tk.Tk()
ventana.title("Ingresar datos")
entrada_texto = tk.Text(ventana, width=100, height=30)
entrada_texto.pack(padx=10, pady=10)
tk.Button(ventana, text="Procesar", command=obtener_datos).pack(pady=10)
ventana.mainloop()
