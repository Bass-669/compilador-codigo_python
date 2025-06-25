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
        barra['value'] = 30 # Actualizar barra al inicio del procesamiento
        ventana_carga.update_idletasks()
        bloques, porcentajes = procesar_datos(txt, torno, mes, dia, anio)
        barra['value'] = 70 # Actualizar barra después de procesar datos
        ventana_carga.update_idletasks()
        if bloques is not None and porcentajes is not None:
            fecha(mes, dia, anio, torno, bloques, porcentajes)
        else:
            messagebox.showwarning("Advertencia", "No se pudo procesar los datos.")
        barra['value'] = 100 # Completar barra al final
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error en ejecutar():\n{e}")
    finally:
        cerrar_carga()
        ventana.destroy()

def procesar_datos(entrada, torno, mes, dia, anio):
    bloques_detectados = []
    celdas_autosuma = []  # Almacenar referencias a celdas AD para procesar después
    if not os.path.exists(RUTA_ENTRADA):
        messagebox.showerror("Error", f"No se encontró:\n{RUTA_ENTRADA}")
        return None, None

    # Inicialización única de win32com
    pythoncom.CoInitialize()
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    try:
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        hoja = wb["IR diario "]
        
        # Búsqueda eficiente de última fila con '* * ...'
        ultima_fila = None
        for fila in range(hoja.max_row, 0, -1):
            valores = [
                str(hoja.cell(fila, 1).value or "").strip(),
                str(hoja.cell(fila, 2).value or "").strip(),
                str(hoja.cell(fila, 3).value or "").strip()
            ]
            if valores == ["*", "*", "..."]:
                ultima_fila = fila
                break

        if not ultima_fila:
            raise ValueError("No se encontró patrón '* * ...'")

        fila = ultima_fila + 1
        bloques = list(extraer_bloques(entrada))

        # Precompile regex para mejor rendimiento
        regex_numero = re.compile(r'^\d')
        regex_asterisco = re.compile(r'^\*')
        
        for b in bloques:
            f_ini = fila
            subs = sub_bloques(b)
            
            for sub in subs:
                txt = sub[0] if not regex_numero.match(sub[0]) else ""
                datos = sub[1:] if txt else sub
                
                # Procesamiento eficiente de texto
                if txt:
                    p = txt.split()
                    if regex_asterisco.match(txt):
                        col_txt = ["*", "*", "...", "", "", ""] if len(p) < 5 or p[0] != "*" else [p[0], p[1], p[2], p[3], "", p[4]]
                    else:
                        col_txt = (
                            [p[0], p[1], p[2], p[3], "", p[4]] if len(p) >= 5 else
                            ["", p[0], p[1], p[2], "", p[3]] if len(p) == 4 else
                            [""] * 6
                        )
                else:
                    col_txt = [""] * 6
                
                # Procesamiento eficiente de números
                col_nums = []
                for l in datos:
                    col_nums.extend(l.strip().split())
                
                # Escritura optimizada en batch
                for col, val in enumerate(col_txt + col_nums, 1):
                    if col > 24: break
                    if 3 <= col <= 24 and val:
                        try:
                            n = float(val.replace(",", "."))
                            escribir(hoja, fila, col, n, True)
                            continue
                        except ValueError:
                            pass
                    escribir(hoja, fila, col, val, False)
                
                # Valores fijos
                for offset, val in enumerate([torno, mes, dia, anio], 25):
                    celda = hoja.cell(fila, offset, val)
                    celda.alignment = ALIGN_R
                
                fila += 1
            
            f_fin = fila - 1
            bloque_texto = " ".join(b).upper()
            tipo_bloque = "PODADO" if "PODADO" in bloque_texto else "REGULAR"
            
            # Configurar fórmulas y celdas especiales
            if len(subs) > 1:
                for f in range(f_ini, f_fin + 1):
                    hoja.cell(f, 30, f"=AC{f}*D{f}/D{f_fin}")
            
            # Configurar celda de autosuma
            celda_autosuma = hoja.cell(f_fin, 30)
            celda_autosuma.value = f"=SUM(AD{f_ini}:AD{f_fin-1})"
            celda_autosuma.fill = FILL_AMARILLO
            celdas_autosuma.append(f"AD{f_fin}")  # Guardar referencia
            
            # Obtener valor D para reporte
            valor_d = 0.0
            d_cell = hoja.cell(f_fin, 4).value
            if d_cell and isinstance(d_cell, (int, float, str)):
                try:
                    valor_d = float(str(d_cell).replace(",", "."))
                except ValueError:
                    pass
            bloques_detectados.append((tipo_bloque, valor_d))
            
            # Limpiar celdas para bloques regulares
            if tipo_bloque != "PODADO":
                for col in range(25, 30):
                    hoja.cell(f_fin, col, "")

        # Guardar cambios antes de leer con win32com
        wb.save(RUTA_ENTRADA)
        wb.close()
        
        # Obtener valores AE usando la función original
        sumas_ad_por_bloque = []
        for celda_ref in celdas_autosuma:
            valor_ae = obtener_valor_ae_directo(excel_app, RUTA_ENTRADA, celda_ref)
            if valor_ae is None:
                return None, None
            sumas_ad_por_bloque.append(valor_ae)
        
        # Crear copias de seguridad
        backup_path = os.path.join(BASE_DIR, CARPETA, "Reporte IR Tornos copia_de_seguridad.xlsx")
        shutil.copy(RUTA_ENTRADA, backup_path)
        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        
        return bloques_detectados, sumas_ad_por_bloque

    except Exception as e:
        messagebox.showerror("Error", f"Error en procesamiento:\n{str(e)}")
        return None, None
    finally:
        if 'wb' in locals() and wb: wb.close()
        excel_app.Quit()
        pythoncom.CoUninitialize()

def obtener_valor_ae_directo(excel_app, ruta_archivo, celda_origen):
    try:
        wb = excel_app.Workbooks.Open(ruta_archivo)
        hoja = wb.Sheets("IR diario ")
        fila = int(''.join(filter(str.isdigit, celda_origen)))
        celda_ae = f"AE{fila}"
        
        origen = hoja.Range(celda_origen)
        origen.Copy()
        
        destino = hoja.Range(celda_ae)
        destino.PasteSpecial(Paste=-4163)  # xlPasteValues
        excel_app.CutCopyMode = False
        
        # Forzar cálculo antes de leer el valor
        excel_app.Calculate()
        valor_ae = destino.Value
        
        wb.Close(False)
        return float(valor_ae) if valor_ae else 0.0
        
    except Exception as e:
        messagebox.showerror("Error", f"Error obteniendo valor AE:\n{e}")
        return None

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

def escribir_valores_resumen_bloques(hoja, col_dia, torno, valores_ae_por_bloque, tipos_bloque):
    for i, (tipo_bloque, valor_ae) in enumerate(zip(tipos_bloque, valores_ae_por_bloque)):
        tipo_bloque = tipo_bloque.strip().upper()
        if tipo_bloque == "PODADO":
            fila_valor = 13 if torno == 1 else 14
        elif tipo_bloque == "REGULAR":
            fila_valor = 18 if torno == 1 else 19
        else:
            continue # ignorar bloques con tipo desconocido
        celda = hoja.cell(row=fila_valor, column=col_dia)
        celda.value = valor_ae / 100 if valor_ae > 1 else valor_ae  # Porcentaje en decimal
        celda.number_format = '0.00%'

def fecha(mes, dia, anio, torno, bloques_detectados, sumas_ad_por_bloque):
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
        valores_para_escribir = [val for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        tipos_para_escribir = [tipo for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        for (tipo_bloque, valor), valor_ae in zip(zip(tipos_para_escribir, valores_para_escribir), sumas_ad_por_bloque):
            escribir_valor_bloque(hoja_nueva, col_dia, torno, valor, tipo_bloque)
        escribir_valores_resumen_bloques(hoja_nueva, col_dia, torno, sumas_ad_por_bloque, tipos_para_escribir)
        wb2.save(RUTA_ENTRADA)
        wb2.close()
        if not hoja_nueva_existia: # Rotar etiquetas solo si es hoja nueva
            rotar_etiquetas_graficos(RUTA_ENTRADA, nueva)
        shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
        mensaje = "✅ Valores actualizados correctamente." if hoja_nueva_existia else f"✅ Hoja '{nueva}' creada correctamente."
        messagebox.showinfo("Éxito", mensaje)
    except Exception as e:
        messagebox.showwarning("Advertencia", f"No se pudo ajustar hoja:\n{e}")

def rotar_etiquetas_graficos(ruta_archivo, nombre_hoja):
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.Dispatch("Excel.Application") # Iniciar Excel en segundo plano
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False  # Optimizar rendimiento
        wb = excel.Workbooks.Open(ruta_archivo)
        sheet = wb.Sheets(nombre_hoja)
        for chart_obj in sheet.ChartObjects(): # Procesar todos los gráficos en la hoja
            try:
                chart = chart_obj.Chart
                x_axis = chart.Axes(1)
                x_axis.TickLabels.Orientation = 45 # Rotar etiquetas a 45 grados
            except Exception as e:
                print(f"Advertencia: Error en gráfico - {str(e)}")
                continue
        wb.Save()
        wb.Close(True)
    except Exception as e:
        print(f"Error crítico al rotar etiquetas: {str(e)}")
    finally:
        try:
            if 'wb' in locals():
                wb.Close(False)
        except:
            pass
        try:
            if excel:
                excel.Quit()
        except:
            pass
        pythoncom.CoUninitialize()

ventana = tk.Tk()
ventana.title("Ingresar datos")
entrada_texto = tk.Text(ventana, width=100, height=30)
entrada_texto.pack(padx=10, pady=10)
tk.Button(ventana, text="Procesar", command=obtener_datos).pack(pady=10)
ventana.mainloop()
