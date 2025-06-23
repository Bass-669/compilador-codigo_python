import openpyxl, re, shutil, time, os, sys, tkinter as tk
from openpyxl.styles import Border, Side, Alignment, PatternFill
from datetime import datetime
from tkinter import messagebox, ttk
from tkcalendar import DateEntry
import win32com.client as win32, pythoncom

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
    ventana.after(100, lambda: ejecutar(texto, torno, mes, dia, anio))

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
        for i in range(1, 101, 20):
            barra['value'] = i
            ventana_carga.update_idletasks()
            time.sleep(1)
        bloques, porcentajes = procesar_datos(txt, torno, mes, dia, anio)
        if bloques is not None and porcentajes is not None:
            fecha(mes, dia, anio, torno, bloques, porcentajes)
        else:
            messagebox.showwarning("Advertencia", "No se pudo procesar los datos.")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error en ejecutar():\n{e}")
    finally:
        cerrar_carga()
        ventana.destroy()

def obtener_valor_numerico(hoja, fila, columna):
    valor = hoja.cell(row=fila, column=columna).value
    if valor is None:
        return 0.0
    if isinstance(valor, str):
        try:
            return float(valor.replace(",", "."))
        except ValueError:
            return 0.0
    return float(valor)

def procesar_datos(entrada, torno, mes, dia, anio):
    bloques_detectados = []
    sumas_ad_por_bloque = []
    valores_para_resumen = []
    
    if not os.path.exists(RUTA_ENTRADA):
        messagebox.showerror("Error", f"No se encontró:\n{RUTA_ENTRADA}")
        return None, None

    # Diccionario para guardar valores {fila: valor_AC}
    valores_ac_dict = {}
    wb = None
    excel = wb_com = None

    try:
        # --- PASO 1: Leer valores de AC con win32com ---
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        try:
            wb_com = excel.Workbooks.Open(RUTA_ENTRADA)
            hoja_com = wb_com.Sheets("IR diario ")
            
            # Verificar existencia de la hoja
            if hoja_com is None:
                raise ValueError("Hoja 'IR diario ' no encontrada")
            
            # Leer todos los valores de AC
            ultima_fila = hoja_com.UsedRange.Rows.Count
            for fila in range(1, ultima_fila + 1):
                cell_value = hoja_com.Range(f"AC{fila}").Value
                valores_ac_dict[fila] = float(cell_value) if cell_value is not None else 0.0
            
            # Cerrar Excel COM
            wb_com.Close(False)
            excel.Quit()
            pythoncom.CoUninitialize()

        except Exception as e:
            if wb_com is not None:
                wb_com.Close(False)
            if excel is not None:
                excel.Quit()
            pythoncom.CoUninitialize()
            raise ValueError(f"Error al leer valores AC: {str(e)}")

        # --- PASO 2: Procesar con openpyxl ---
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        hoja = wb["IR diario "]
        
        # Encontrar fila con '* * ...' (búsqueda segura)
        ultima_fila = None
        for fila in hoja.iter_rows(min_row=1, max_row=10000, max_col=3):
            valores = [
                str(fila[0].value).strip() if fila[0].value else "",
                str(fila[1].value).strip() if fila[1].value else "",
                str(fila[2].value).strip() if fila[2].value else ""
            ]
            if valores == ["*", "*", "..."]:
                ultima_fila = fila[0].row
                break

        if not ultima_fila:
            raise ValueError("No se encontró el patrón '* * ...'")

        fila = ultima_fila + 1
        bloques = extraer_bloques(entrada)
        
        if not bloques:
            raise ValueError("No se detectaron bloques válidos en los datos")

        for b in bloques:
            if not b:
                continue
                
            f_ini = fila
            subs = sub_bloques(b)
            valores_d = []
            valores_ae = []

            for sub in subs:
                if not sub:
                    continue
                
                # Leer valor D desde Excel
                try:
                    valor_d = obtener_valor_numerico(hoja, fila, 4)
                    valores_d.append(valor_d)
                    
                    # Obtener valor AC desde memoria
                    valor_ae = valores_ac_dict.get(fila, 0.0)
                    valores_ae.append(valor_ae)
                    
                    # Escribir metadatos
                    for col, val in zip(range(25, 29), [torno, mes, dia, anio]):
                        hoja.cell(row=fila, column=col, value=val).alignment = ALIGN_R
                    
                    fila += 1
                
                except Exception as e:
                    raise ValueError(f"Error en fila {fila}: {str(e)}")

            # Calcular suma AD
            try:
                d_fin = valores_d[-1] if valores_d else 0.0
                suma_ad_manual = sum(
                    (valores_ae[i] * valores_d[i]) / d_fin 
                    for i in range(len(valores_d) - 1) 
                    if d_fin != 0
                )
                
                # Clasificar bloque
                tipo_bloque = "PODADO" if "PODADO" in " ".join(b).upper() else "REGULAR"
                bloques_detectados.append((tipo_bloque, d_fin))
                valores_para_resumen.append(suma_ad_manual)
            
            except Exception as e:
                raise ValueError(f"Error en cálculo AD: {str(e)}")

        # Guardar resultados
        wb.save(RUTA_ENTRADA)
        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        
        return bloques_detectados, valores_para_resumen

    except Exception as e:
        messagebox.showerror("Error", f"Error al procesar datos:\n{str(e)}")
        return None, None
    
    finally:
        # Limpieza segura
        if wb is not None:
            wb.close()
        if excel is not None:
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

def escribir_valores_resumen_bloques(hoja, col_dia, torno, porcentajes_por_bloque, tipos_bloque):
    for i, (tipo_bloque, porcentaje) in enumerate(zip(tipos_bloque, porcentajes_por_bloque)):
        tipo_bloque = tipo_bloque.strip().upper()
        if tipo_bloque == "PODADO":
            fila_valor = 13 if torno == 1 else 14
        elif tipo_bloque == "REGULAR":
            fila_valor = 18 if torno == 1 else 19
        else:
            continue
        celda = hoja.cell(row=fila_valor, column=col_dia)
        celda.value = porcentaje / 100 if porcentaje > 1 else porcentaje  # Asegurar valor decimal
        celda.number_format = '0.00%'

def fecha(mes, dia, anio, torno, bloques_detectados, sumas_ad_por_bloque ):
    """Función principal para crear/modificar la hoja de reporte"""
    pythoncom.CoInitialize()
    excel = wb = None
    nueva = f"IR {mes} {anio}"
    hoja_anterior = None
    hoja_nueva_existia = False
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(RUTA_ENTRADA, UpdateLinks=0)
        nombres_hojas = [h.Name for h in wb.Sheets]
        hoja_nueva_existia = nueva in nombres_hojas
        if not hoja_nueva_existia:
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
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo crear hoja:\n{e}")
        return
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=True)
        except:
            pass
        try:
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

    try:
        wb2 = openpyxl.load_workbook(RUTA_ENTRADA)
        hoja_nueva = wb2[nueva]
        col_dia = dia + 1  # columna B es 2, día 1 → columna 2

        if not hoja_nueva_existia:
            filas_fechas = [2, 3, 4, 7, 8, 9, 12, 13, 14, 17, 18, 19, 22, 27, 31, 37]
            for fila in filas_fechas:
                for col in range(2, 33):
                    hoja_nueva.cell(row=fila, column=col, value="")

        nueva_fecha = f"{dia:02d}/{MESES_NUM[mes]:02d}/{anio}"
        for fila in [2, 7, 12, 17, 22, 27, 31, 37]:
            hoja_nueva.cell(row=fila, column=col_dia, value=nueva_fecha)

        for tipo_bloque, f_fin in bloques_detectados:
            escribir_valor_bloque(hoja_nueva, col_dia, torno, f_fin, tipo_bloque)

        # Extraer lista de tipos de bloque para resumen
        tipos_bloque = [tipo for tipo, _ in bloques_detectados]
        escribir_valores_resumen_bloques(hoja_nueva, col_dia, torno, sumas_ad_por_bloque, tipos_bloque)

        wb2.save(RUTA_ENTRADA)
        wb2.close()
        time.sleep(1)

        excel_app = None
        wb_excel = None
        try:
            pythoncom.CoInitialize()
            excel_app = win32.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            wb_excel = excel_app.Workbooks.Open(RUTA_ENTRADA)
            sheet_excel = wb_excel.Sheets(nueva)
            for chart_obj in sheet_excel.ChartObjects():
                try:
                    chart = chart_obj.Chart
                    x_axis = chart.Axes(1)
                    x_axis.TickLabels.Orientation = 45
                except Exception as e:
                    print(f"Error en gráfico: {str(e)}")
                    continue
            wb_excel.Save()
            wb_excel.Close(True)
        except Exception as e:
            messagebox.showwarning("Advertencia", f"Error al rotar etiquetas: {str(e)}")
        finally:
            try:
                if wb_excel and wb_excel.ReadOnly == False:
                    wb_excel.Close(False)
            except:
                pass
            try:
                if excel_app:
                    excel_app.Quit()
            except:
                pass
            wb_excel = None
            excel_app = None
            pythoncom.CoUninitialize()

        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        mensaje = "✅ Valores actualizados correctamente." if hoja_nueva_existia else f"✅ Hoja '{nueva}' creada correctamente."
        messagebox.showinfo("Éxito", mensaje)

    except Exception as e:
        messagebox.showwarning("Advertencia", f"No se pudo ajustar hoja:\n{e}")

ventana = tk.Tk()
ventana.title("Ingresar datos")
entrada_texto = tk.Text(ventana, width=100, height=30)
entrada_texto.pack(padx=10, pady=10)
tk.Button(ventana, text="Procesar", command=obtener_datos).pack(pady=10)
ventana.mainloop()
