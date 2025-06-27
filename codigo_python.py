import openpyxl, re, shutil, time, os, sys, tkinter as tk
from openpyxl.styles import Border, Side, Alignment, PatternFill
from datetime import datetime
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
import win32com.client as win32, pythoncom
import threading

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
bloques_detectados = []
sumas_ad_por_bloque = []

def obtener_datos():
    datos = entrada_texto.get("1.0", tk.END).strip()
    if not datos: return messagebox.showwarning("Advertencia", "Ingresa los datos.")
    pedir_torno(lambda t: pedir_fecha(lambda m,d,a: iniciar(datos, t, m, d, a)))

def pedir_torno(callback):
    def confirmar():
        val = ent.get().strip()
        if not val: return messagebox.showwarning("Advertencia", "Ingresa el número de torno.")
        try: callback(int(val)); ventana.destroy()
        except: messagebox.showerror("Error", "Debe ser un número.")
    ventana = tk.Toplevel()
    ventana.title("Número de torno")
    ventana.geometry("300x120")
    ventana.resizable(False, False)
    tk.Label(ventana, text="Número de torno:", font=("Arial", 12)).pack(pady=10)
    ent = tk.Entry(ventana, font=("Arial", 12)); ent.pack(pady=5)
    tk.Button(ventana, text="Aceptar", command=confirmar).pack(pady=5)
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
    try:
        # Paso 1: Procesar datos de entrada
        barra['value'] = 30
        ventana_carga.update_idletasks()
        bloques, porcentajes = procesar_datos(txt, torno, mes, dia, anio)
        
        if not bloques or not porcentajes:
            messagebox.showerror("Error", "No se obtuvieron datos válidos")
            return

        # Paso 2: Asignar rendimiento con manejo de espera
        barra['value'] = 50
        ventana_carga.update_idletasks()
        
        if not asignar_rendimiento_a_ir(dia, mes, anio, torno):
            messagebox.showwarning("Advertencia", "No se pudo asignar rendimiento")
        
        # Espera adicional para asegurar cierre de archivos
        time.sleep(1)

        # Paso 3: Procesar fecha
        barra['value'] = 70
        ventana_carga.update_idletasks()
        fecha(mes, dia, anio, torno, bloques, porcentajes)

        barra['value'] = 100
        ventana_carga.update_idletasks()
        time.sleep(0.5)
        
        messagebox.showinfo("Éxito", "Proceso completado")

    except Exception as e:
        messagebox.showerror("Error Crítico", f"Error en ejecutar():\n{str(e)}")
        
    finally:
        cerrar_carga()
        try:
            ventana.destroy()
        except:
            pass  # En caso de que la ventana ya esté cerrada

def procesar_datos(entrada, torno, mes, dia, anio):
    bloques_detectados = []
    sumas_ad_por_bloque = []
    if not os.path.exists(RUTA_ENTRADA):
        messagebox.showerror("Error", f"No se encontró:\n{RUTA_ENTRADA}")
        return None, None
    try:
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        hoja = wb["IR diario "]
        ultima_fila = None
        for fila in hoja.iter_rows():
            if [str(c.value).strip() if c.value else "" for c in fila[:3]] == ["*", "*", "..."]:
                ultima_fila = fila[0].row
        if not ultima_fila:
            raise ValueError("No se encontró '* * ...'")
        fila = ultima_fila + 1
        for b in extraer_bloques(entrada):
            f_ini = fila
            subs = sub_bloques(b)
            for sub in subs:
                txt = sub[0] if not re.match(r'^\d', sub[0]) else ""
                datos = sub[1:] if txt else sub
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
                for col, val in enumerate(fila_vals[:24], 1):
                    try:
                        n = float(val.replace(",", ".")) if 3 <= col <= 24 and val else val
                        escribir(hoja, fila, col, n, isinstance(n, float))
                    except:
                        escribir(hoja, fila, col, val)
                for col, val in zip(range(25, 29), [torno, mes, dia, anio]):
                    hoja.cell(row=fila, column=col, value=val).alignment = ALIGN_R
                fila += 1
            f_fin = fila - 1
            tipo_bloque = "PODADO" if "PODADO" in txt.upper() else "REGULAR"
            bloques_detectados.append((tipo_bloque, f_fin))
            if len(subs) > 1:
                for f in range(f_ini, f_fin + 1):
                    hoja.cell(row=f, column=30, value=f"=AC{f}*D{f}/D{f_fin}")
            fila_autosuma = fila - 1  # porque después del último subbloque, fila ya se incrementó una más
            for col in range(25, 30):
                hoja.cell(row=fila_autosuma, column=col, value="")
            celda_autosuma = hoja.cell(row=fila_autosuma, column=30)
            celda_autosuma.value = f"=SUM(AD{f_ini}:AD{fila_autosuma - 1})"
            celda_autosuma.fill = FILL_AMARILLO
            celda_origen = f"AD{fila_autosuma}" # Guarda el valor que había antes en la celda de autosuma
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
            wb.save(RUTA_ENTRADA)
            shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO)) # asegurar que temp tenga los datos
            temp_path, valor_ae = crear_archivo_temporal_con_ae(celda_origen)
            if not temp_path:
                return None, None
            sumas_ad_por_bloque.append(valor_ae)  # Guardar el valor real desde AE
        backup_path = os.path.join(CARPETA, "Reporte IR Tornos copia_de_seguridad.xlsx")
        shutil.copy(RUTA_ENTRADA, backup_path)
        wb.save(RUTA_ENTRADA)
        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        return bloques_detectados, sumas_ad_por_bloque
    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar datos:\n{e}")
        return None, None
    finally:
        if 'wb' in locals():
            wb.close()

def crear_archivo_temporal_con_ae(celda_origen):
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    temp_path = os.path.join(BASE_DIR, CARPETA, "temp_report.xlsx")
    
    try:
        # 1. Abrir el archivo
        wb = excel.Workbooks.Open(RUTA_ENTRADA)
        hoja = wb.Sheets("IR diario ")

        # 2. Extraer número de fila (ej: "AD43448" -> 43448)
        fila = int(''.join(filter(str.isdigit, celda_origen)))

        # 3. Obtener el valor de AD43448 y asignarlo directamente a AE43448
        valor_ad = hoja.Range(celda_origen).Value
        if valor_ad is not None:
            try:
                # Convertir a float (manejando comas/puntos)
                valor_ae = float(str(valor_ad).replace(",", "."))
                hoja.Range(f"AE{fila}").Value = valor_ae  # ¡Asignación directa aquí!
            except ValueError as e:
                messagebox.showerror("Error", f"Valor no numérico en {celda_origen}: {valor_ad}")
                return None, 0.0

        # 4. Guardar el archivo temporal (asegura que los cambios persistan)
        wb.SaveAs(temp_path, FileFormat=51)  # FileFormat=51 es .xlsx
        wb.Close(False)

        # 5. Leer el valor desde el archivo temporal
        wb_temp = openpyxl.load_workbook(temp_path, data_only=True)
        hoja_temp = wb_temp["IR diario "]
        valor_final = hoja_temp.cell(row=fila, column=31).value  # Columna AE=31
        wb_temp.close()

        return temp_path, float(valor_final) if valor_final else 0.0

    except Exception as e:
        messagebox.showerror("Error", f"Fallo al guardar en AE{fila}:\n{str(e)}")
        return None, 0.0
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

def escribir(hoja, f, c, v, num=False):
    celda = hoja.cell(row=f, column=c, value=v)
    celda.border, celda.alignment = BORDER, ALIGN_R
    if num: celda.number_format = '0.00'

def extraer_bloques(txt):
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
    subs, tmp = [], []
    for l in b:
        if re.match(r'^\D', l) or '*' in l:
            if tmp: subs.append(tmp)
            tmp = [l]
        else: tmp.append(l)
    if tmp: subs.append(tmp)
    return subs

def escribir_valor_bloque(hoja, col_dia, torno, valor, tipo_bloque):
    tipo_bloque = tipo_bloque.strip().upper()
    if tipo_bloque == "PODADO":
        fila_valor = 3 if torno == 1 else 4
    elif tipo_bloque == "REGULAR":
        fila_valor = 8 if torno == 1 else 9
    else:
        messagebox.showwarning("Advertencia", f"Tipo de bloque no reconocido: '{tipo_bloque}'")
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
    celda.number_format = '0' # pequeño cambio de 0.00 a 0

# def escribir_valores_resumen_bloques(hoja, col_dia, torno, valores_ae_por_bloque, tipos_bloque):
#     for i, (tipo_bloque, valor_ae) in enumerate(zip(tipos_bloque, valores_ae_por_bloque)):
#         tipo_bloque = tipo_bloque.strip().upper()
#         if tipo_bloque == "PODADO":
#             fila_valor = 13 if torno == 1 else 14
#         elif tipo_bloque == "REGULAR":
#             fila_valor = 18 if torno == 1 else 19
#         else:
#             continue # ignorar bloques con tipo desconocido
#         celda = hoja.cell(row=fila_valor, column=col_dia)
#         celda.value = valor_ae / 100 if valor_ae > 1 else valor_ae  # Porcentaje en decimal
#         celda.number_format = '0.00%'

def escribir_valores_resumen_bloques(hoja, col_dia, torno, valores_ae_por_bloque, tipos_bloque, rendimiento=None):
    # Escribir valores de bloques
    for i, (tipo_bloque, valor_ae) in enumerate(zip(tipos_bloque, valores_ae_por_bloque)):
        tipo_bloque = tipo_bloque.strip().upper()
        if tipo_bloque == "PODADO":
            fila_valor = 13 if torno == 1 else 14
        elif tipo_bloque == "REGULAR":
            fila_valor = 18 if torno == 1 else 19
        else:
            continue  # ignorar bloques con tipo desconocido
        
        celda = hoja.cell(row=fila_valor, column=col_dia)
        celda.value = valor_ae / 100 if valor_ae > 1 else valor_ae  # Porcentaje en decimal
        celda.number_format = '0.00%'

    # Escribir rendimiento si se proporciona
    if rendimiento is not None:
        fila_rendimiento = 32 if torno == 1 else 33
        celda_rendimiento = hoja.cell(row=fila_rendimiento, column=col_dia)
        
        # Asegurar que el rendimiento sea un valor numérico válido
        try:
            valor_rendimiento = float(rendimiento)
            celda_rendimiento.value = valor_rendimiento / 100 if valor_rendimiento > 1 else valor_rendimiento
            celda_rendimiento.number_format = '0.00%'
        except (ValueError, TypeError):
            # Manejar caso de rendimiento no válido
            celda_rendimiento.value = None
            print(f"Advertencia: Valor de rendimiento no válido: {rendimiento}")


def fecha(mes, dia, anio, torno, bloques_detectados, sumas_ad_por_bloque):
    try:
        # Validación básica de parámetros (original)
        if not bloques_detectados or not sumas_ad_por_bloque:
            messagebox.showerror("Error", "Datos de bloques o sumas AD están vacíos")
            return

        # Convertir mes a número (original)
        try:
            mes_num = MESES_NUM[mes]
        except KeyError:
            messagebox.showerror("Error", f"Mes no válido: {mes}")
            return

        # --- Cambio clave: Formato de fecha solo para búsqueda en Peeling ---
        fecha_peeling = f"{anio}/{mes_num:02d}/{dia:02d}"  # YYYY/MM/DD (solo para Peeling)
        fecha_principal = f"{dia:02d}/{mes_num:02d}/{anio}"  # DD/MM/YYYY (para Excel principal)

        # Procesamiento con Excel (original)
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # --- Resto del código ORIGINAL sin cambios ---
        nueva = f"IR {mes} {anio}"
        hoja_anterior = None
        hoja_nueva_existia = False
        
        wb = excel.Workbooks.Open(RUTA_ENTRADA, UpdateLinks=0)
        nombres_hojas = [h.Name for h in wb.Sheets]
        hoja_nueva_existia = nueva in nombres_hojas
        
        if not hoja_nueva_existia:
            # ... (código original para buscar hoja anterior)
            hojas_ir = [h for h in nombres_hojas if h.startswith("IR ") and len(h.split()) == 3]
            def total_meses(nombre):
                try:
                    _, mes_str, anio_str = nombre.split()
                    return int(anio_str) * 12 + MESES_NUM[mes_str]
                except:
                    return -1
            hojas_ir_ordenadas = sorted(hojas_ir, key=total_meses)
            total_nueva = int(anio) * 12 + MESES_NUM[mes]
            for h in hojas_ir_ordenadas:
                if total_meses(h) < total_nueva:
                    hoja_anterior = h
                else:
                    break
            
            if not hoja_anterior:
                messagebox.showwarning("Orden inválido", f"No se encontró hoja anterior para insertar '{nueva}'")
                return
            
            idx_anterior = [h.Name for h in wb.Sheets].index(hoja_anterior)
            insert_idx = min(idx_anterior + 2, wb.Sheets.Count)
            wb.Sheets(hoja_anterior).Copy(After=wb.Sheets(insert_idx - 1))
            wb.ActiveSheet.Name = nueva
            wb.Save()

        # --- Continuación del código original ---
        wb2 = openpyxl.load_workbook(RUTA_ENTRADA)
        hoja_nueva = wb2[nueva]
        col_dia = dia + 1
        
        if not hoja_nueva_existia:
            filas_fechas = [2, 7, 12, 17, 22, 27, 31, 37]
            for fila in filas_fechas:
                for col in range(2, 37):
                    hoja_nueva.cell(row=fila, column=col, value="")
        
        # Usar formato DD/MM/YYYY en Excel principal (original)
        for fila in [2, 7, 12, 17, 22, 27, 31, 37]:
            hoja_nueva.cell(row=fila, column=col_dia, value=f"{dia:02d}/{mes_num:02d}/{anio}")
        
        # ... (resto del código original para escribir valores)
        valores_para_escribir = [val for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        tipos_para_escribir = [tipo for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        
        for (tipo_bloque, valor), valor_ae in zip(zip(tipos_para_escribir, valores_para_escribir), sumas_ad_por_bloque):
            escribir_valor_bloque(hoja_nueva, col_dia, torno, valor, tipo_bloque)
            escribir_valores_resumen_bloques(hoja_nueva, col_dia, torno, sumas_ad_por_bloque, tipos_para_escribir)
        
        wb2.save(RUTA_ENTRADA)
        wb2.close()
        
        if not hoja_nueva_existia:
            rotar_etiquetas_graficos(RUTA_ENTRADA, nueva)  # Función original conservada
        
        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        
        mensaje = "✅ Valores actualizados" if hoja_nueva_existia else f"✅ Hoja '{nueva}' creada"
        messagebox.showinfo("Éxito", mensaje)

    except Exception as e:
        messagebox.showerror("Error", f"Error en fecha():\n{str(e)}")
    finally:
        try:
            if 'wb' in locals():
                wb.Close(SaveChanges=True)
            if 'excel' in locals():
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

#-----------------------------------------------------------------------

def obtener_rendimientos_peeling(dia, mes, anio, torno_1=3011, torno_2=3012):
    try:
        ruta_peeling = os.path.join(BASE_DIR, "Nueva_Peeling_Query_202050501_arauco.xlsx")
        if not os.path.exists(ruta_peeling):
            messagebox.showerror("Error", f"Archivo no encontrado:\n{ruta_peeling}")
            return None

        # Usar with statement para manejo automático del archivo
        wb = None
        try:
            wb = openpyxl.load_workbook(ruta_peeling, data_only=True, read_only=True)  # Modo read-only
            if "Sheet2" not in wb.sheetnames:
                messagebox.showerror("Error", "La hoja 'Sheet2' no existe")
                return None

            hoja = wb["Sheet2"]
            mes_num = MESES_NUM.get(mes, 0)
            
            formatos_fecha = [
                f"{anio}/{mes_num:02d}/{dia:02d}",
                f"{dia}/{mes_num}/{anio}",
                f"{dia}-{mes_num}-{anio}",
                f"{anio}-{mes_num:02d}-{dia:02d}"
            ]

            rendimientos = {}
            
            for fila in range(2, hoja.max_row + 1):
                fecha_celda = str(hoja.cell(row=fila, column=1).value or "").strip()
                
                if fecha_celda in formatos_fecha:
                    try:
                        work_id = int(hoja.cell(row=fila, column=2).value)
                        rendimiento = float(hoja.cell(row=fila, column=12).value)
                        
                        if work_id == torno_1:
                            rendimientos['torno_1'] = rendimiento
                        elif work_id == torno_2:
                            rendimientos['torno_2'] = rendimiento
                            
                        if len(rendimientos) == 2:
                            break
                    except (ValueError, TypeError):
                        continue

            return rendimientos if rendimientos else None

        finally:
            if wb:
                wb.close()  # Cierre explícito

    except Exception as e:
        messagebox.showerror("Error", f"Error al leer archivo:\n{str(e)}")
        return None


def asignar_rendimiento_a_ir(dia, mes, anio, torno):
    try:
        # Validación básica
        mes_num = MESES_NUM.get(mes)
        if not mes_num:
            return None

        # Obtener rendimientos (sin mensajes interactivos)
        rendimientos = obtener_rendimientos_peeling(dia, mes, anio)
        if not rendimientos or f'torno_{torno}' not in rendimientos:
            return None

        rendimiento = rendimientos[f'torno_{torno}']

        # Guardar directamente en Excel
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        nombre_hoja = f"IR {mes} {anio}"
        
        if nombre_hoja not in wb.sheetnames:
            wb.create_sheet(nombre_hoja)

        hoja = wb[nombre_hoja]
        col_dia = dia + 1
        fila_rendimiento = 32 if torno == 1 else 33

        hoja.cell(row=fila_rendimiento, column=col_dia, value=rendimiento/100).number_format = '0.00%'
        wb.save(RUTA_ENTRADA)
        wb.close()

        # Pequeña pausa para asegurar el cierre
        time.sleep(0.5)
        
        return True
    except PermissionError:
        messagebox.showerror("Error", "No se pudo guardar - archivo en uso o sin permisos")
        return False
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar rendimiento: {str(e)}")
        return False

except Exception as e:
    messagebox.showerror("Error", f"Error general: {str(e)}")
    return None


ventana = tk.Tk()
ventana.title("Ingresar datos")
entrada_texto = tk.Text(ventana, width=100, height=30)
entrada_texto.pack(padx=10, pady=10)
tk.Button(ventana, text="Procesar", command=obtener_datos).pack(pady=10)
ventana.mainloop()
