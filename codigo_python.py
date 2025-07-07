import openpyxl, re, shutil, time, os, sys, tkinter as tk
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
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
        barra['value'] = 25
        ventana_carga.update_idletasks()

        # Paso 1: Preparar hoja del mes antes de cualquier escritura
        if not preparar_hoja_mes(mes, dia, anio):
            return

        barra['value'] = 50
        ventana_carga.update_idletasks()

        # Paso 2: Procesar los datos y escribir en "IR diario"
        bloques, porcentajes = procesar_datos(txt, torno, mes, dia, anio)

        barra['value'] = 75
        ventana_carga.update_idletasks()

        # Paso 3: Escribir los valores procesados en hoja mensual
        if bloques is not None and porcentajes is not None:
            fecha(mes, dia, anio, torno, bloques, porcentajes)

        barra['value'] = 100

    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error en ejecutar():\n{e}")
    finally:
        cerrar_carga()
        ventana.destroy()


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

                # Configurar celda de autosuma (última fila del bloque)
                for col in range(25, 30):# 29
                    hoja.cell(row=f_fin, column=col, value="")

                celda_autosuma = hoja.cell(row=f_fin, column=30)
                celda_autosuma.value = f"=SUM(AD{f_ini}:AD{f_fin-1})"
                celda_autosuma.fill = FILL_AMARILLO
                # celda_origen = f"AD{fila_autosuma}" # Guarda el valor que había antes en la celda de autosuma
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
                # Guardar cambios ANTES de procesar archivo temporal
                wb.save(RUTA_ENTRADA)
                # Procesar archivo temporal para obtener valor AE
                try:
                    temp_path, valor_ae = crear_archivo_temporal_con_ae(f"AD{f_fin}")
                    sumas_ad_por_bloque.append(valor_ae) 
                except Exception as e:
                    sumas_ad_por_bloque.append(0.0)
                # Guardar cambios finales del bloque
                wb.save(RUTA_ENTRADA)
                shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))

            except Exception as e:
                print(f"Error en bloque: {e}")
                continue

        # Copia de seguridad final
        backup_path = os.path.join(CARPETA, "Reporte IR Tornos copia_de_seguridad.xlsx")
        shutil.copy(RUTA_ENTRADA, backup_path)
        return bloques_detectados, sumas_ad_por_bloque

    except Exception as e:
        messagebox.showerror("Error", f"Error general: {e}")
        return None, None

    finally:
        if 'wb' in locals():
            wb.close()

def escribir(hoja, fila, col, valor, es_numero=False):
    """Escribe un valor en la celda con formato adecuado"""
    celda = hoja.cell(row=fila, column=col, value=valor)
    celda.border = BORDER
    celda.alignment = ALIGN_R
    if es_numero:
        celda.number_format = '0.00'

def es_valor_valido(valor):
    """Verifica si un valor es numérico y válido"""
    if valor in (None, "#NIA", "#N/A", "#VALOR!", ""):
        return False
    try:
        float(str(valor).replace(",", "."))
        return True
    except:
        return False


def crear_archivo_temporal_con_ae(celda_origen):
    """Retorna la referencia CORRECTAMENTE formateada"""
    if not re.match(r'^AD\d+$', celda_origen):
        messagebox.showerror("Error", f"Formato de celda inválido: {celda_origen}")
        raise ValueError(f"Formato de celda inválido: {celda_origen}")
    
    # FORMATO EXACTO REQUERIDO POR EXCEL:
    referencia = f"='IR diario '!{celda_origen}"
    
    # messagebox.showinfo("DEBUG Referencia", 
    #                   f"Referencia generada:\n{referencia}\n"
    #                   f"Verificar que tenga:\n"
    #                   f"1. = al inicio\n"
    #                   f"2. 'IR diario ' entre comillas simples\n"
    #                   f"3. ! antes de AD\n"
    #                   f"4. Número válido")
    
    return None, referencia

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
    celda.number_format = '0' # cambio de 0.00 a 0

def escribir_valores_resumen_bloques(hoja, col_dia, torno, valores_ae_por_bloque, tipos_bloque):
    """Escribe las referencias de Excel directamente sin calcular valores."""
    for i, (tipo_bloque, referencia) in enumerate(zip(tipos_bloque, valores_ae_por_bloque)):
        try:
            tipo_bloque = tipo_bloque.strip().upper()
            
            # Determinar fila destino
            if tipo_bloque == "PODADO":
                fila_valor = 13 if torno == 1 else 14
            elif tipo_bloque == "REGULAR":
                fila_valor = 18 if torno == 1 else 19
            else:
                continue  # Saltar tipos no reconocidos

            # Escribir la referencia DIRECTAMENTE (sin Excel COM)
            celda = hoja.cell(row=fila_valor, column=col_dia)
            celda.value = referencia  # Ej: ='IR diario '!AD123
            celda.alignment = ALIGN_R  # Alineación derecha (opcional)

            # DEBUG (opcional)
            print(f"Bloque {i} ({tipo_bloque}) | Torno {torno} | Fila {fila_valor}")
            print(f"Referencia escrita: {referencia}")
            
        except Exception as e:
            print(f"Error en bloque {i}: {str(e)}")
            continue  # Continuar con el siguiente bloque

def fecha(mes, dia, anio, torno, bloques_detectados, sumas_ad_por_bloque):
    """Escribe los datos en la hoja del mes incluyendo las fechas"""
    nombre_hoja = f"IR {mes} {anio}"
    col_dia = dia + 1  # columna B es 2, día 1 → columna 2
    wb = None
    try:
        # messagebox.showinfo("Proceso", f"Iniciando escritura en hoja {nombre_hoja} para el día {dia}")
        # 1. Abrir el archivo principal
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        # Verificar si la hoja existe
        if nombre_hoja not in wb.sheetnames:
            messagebox.showerror("Error", f"No se encontró la hoja '{nombre_hoja}'")
            return False
        hoja_mes = wb[nombre_hoja]
        # 2. Escribir la fecha en las celdas correspondientes
        # messagebox.showinfo("Proceso", f"Escribiendo fecha {dia}/{MESES_NUM[mes]:02d}/{anio} en columna {col_dia}")
        nueva_fecha = f"{dia:02d}/{MESES_NUM[mes]:02d}/{anio}"

        # Celdas donde debe ir la fecha (filas basadas en tu estructura de Excel)
        filas_fecha = [2, 7, 12, 17, 22, 27, 31, 37]

        for fila in filas_fecha:
            try:
                celda = hoja_mes.cell(row=fila, column=col_dia)
                celda.value = nueva_fecha
                # messagebox.showinfo("Detalle", f"Fecha escrita en {fila},{col_dia}")
            except Exception as e:
                messagebox.showwarning("Advertencia", f"Error escribiendo fecha en {fila},{col_dia}: {str(e)}")

        # 3. Escribir valores de bloques
        # messagebox.showinfo("Proceso", "Escribiendo valores de bloques...")
        valores_para_escribir = [val for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]
        tipos_para_escribir = [tipo for i, (tipo, val) in enumerate(bloques_detectados) if i % 2 == 1]

        for (tipo_bloque, valor), valor_ae in zip(zip(tipos_para_escribir, valores_para_escribir), sumas_ad_por_bloque):
            escribir_valor_bloque(hoja_mes, col_dia, torno, valor, tipo_bloque)
            escribir_valores_resumen_bloques(hoja_mes, col_dia, torno, [valor_ae], [tipo_bloque])
            # messagebox.showinfo("Detalle", f"Bloque {tipo_bloque} escrito en columna {col_dia}")
        # 4. Guardar cambios
        # messagebox.showinfo("Proceso", "Guardando cambios...")
        wb.save(RUTA_ENTRADA)
        # 5. Copia de seguridad
        try:
            shutil.copy(RUTA_ENTRADA, os.path.join(BASE_DIR, ARCHIVO))
            # messagebox.showinfo("Proceso", "Copia de seguridad creada")
        except Exception as e:
            messagebox.showwarning("Advertencia", f"No se pudo crear copia de seguridad:\n{str(e)}")

        messagebox.showinfo("Éxito", "✅ Valores actualizados correctamente.")
        return True

    except Exception as e:
        messagebox.showerror("Error", 
            f"No se pudo escribir en hoja:\n{str(e)}\n\n"
            "Verifique que:\n"
            "1. El archivo no esté abierto en Excel\n"
            "2. Tenga permisos de escritura\n"
            "3. La hoja exista en el archivo"
        )
        return False

    finally:
        # 6. Cerrar el workbook si está abierto
        if wb is not None:
            try:
                wb.close()
                # messagebox.showinfo("Proceso", "Workbook cerrado correctamente")
            except:
                messagebox.showwarning("Advertencia", "Error al cerrar el workbook")
                pass

# def preparar_hoja_mes(mes, dia, anio):
#     """Crea la hoja del mes si no existe y configura fórmulas iniciales"""
#     nombre_hoja = f"IR {mes} {anio}"
#     col_dia = dia + 1  # columna B es 2, día 1 → columna 2
#     hoja_nueva_creada = False

#     try:
#         # 1. Verificar si la hoja ya existe
#         wb_check = openpyxl.load_workbook(RUTA_ENTRADA)

#         if nombre_hoja in wb_check.sheetnames:
#             wb_check.close()

#             # Verificar si el día ya tiene datos
#             wb_existente = openpyxl.load_workbook(RUTA_ENTRADA)
#             hoja_existente = wb_existente[nombre_hoja]
#             # Comprobar si las celdas clave para este día están vacías
#             celdas_clave = [
#                 hoja_existente.cell(row=3, column=col_dia).value,  # Torno 1 regular
#                 hoja_existente.cell(row=4, column=col_dia).value,  # Torno 1 podado
#                 hoja_existente.cell(row=8, column=col_dia).value,  # Torno 2 regular
#                 hoja_existente.cell(row=9, column=col_dia).value   # Torno 2 podado
#             ]

#             dia_con_datos = any(cell is not None and str(cell).strip() != "" for cell in celdas_clave)
#             wb_existente.close()
#             if dia_con_datos:
#                 print("Información", f"El día {dia} ya tiene datos en la hoja {nombre_hoja}. Se agregarán los nuevos datos.")
#             return True
#         else:
#             wb_check.close()

#             # 2. Crear nueva hoja usando Excel COM
#             pythoncom.CoInitialize()
#             excel = wb = None
#             try:
#                 excel = win32.gencache.EnsureDispatch('Excel.Application')
#                 excel.Visible = False
#                 excel.DisplayAlerts = False
#                 wb = excel.Workbooks.Open(RUTA_ENTRADA, UpdateLinks=0)
#                 # Encontrar hoja anterior adecuada para copiar
#                 hojas = [h.Name for h in wb.Sheets]
#                 hojas_ir = [h for h in hojas if h.startswith("IR ") and len(h.split()) == 3]

#                 def total_meses(nombre):
#                     try:
#                         _, mes_str, anio_str = nombre.split()
#                         return int(anio_str) * 12 + MESES_NUM[mes_str]
#                     except:
#                         return -1

#                 hojas_ordenadas = sorted(hojas_ir, key=total_meses)
#                 total_nueva = int(anio) * 12 + MESES_NUM[mes]
#                 hoja_anterior = None

#                 for h in hojas_ordenadas:
#                     if total_meses(h) < total_nueva:
#                         hoja_anterior = h
#                     else:
#                         break

#                 if not hoja_anterior:
#                     messagebox.showwarning("Orden inválido", f"No se encontró hoja anterior para insertar '{nombre_hoja}'")
#                     return False

#                 # Copiar hoja anterior
#                 idx_anterior = hojas.index(hoja_anterior)
#                 insert_idx = min(idx_anterior + 2, wb.Sheets.Count)
#                 wb.Sheets(hoja_anterior).Copy(After=wb.Sheets(insert_idx - 1))
#                 # Renombrar la nueva hoja
#                 nueva_hoja = wb.ActiveSheet
#                 nueva_hoja.Name = nombre_hoja
#                 wb.Save()
#                 hoja_nueva_creada = True

#             except Exception as e:
#                 messagebox.showerror("Error", f"No se pudo crear hoja nueva:\n{e}")
#                 return False
#             finally:
#                 try:
#                     if wb:
#                         wb.Close(SaveChanges=True)
#                 except:
#                     pass
#                 try:
#                     if excel:
#                         excel.Close() #cambio 
#                 except:
#                     pass
#                 pythoncom.CoUninitialize()

#             # Rotar etiquetas de gráficos en la nueva hoja
#             rotar_etiquetas_graficos(RUTA_ENTRADA, nombre_hoja)

#         # 3. Configurar solo para nueva hoja (no limpiar nada para días nuevos)
#         if hoja_nueva_creada:
#             wb2 = openpyxl.load_workbook(RUTA_ENTRADA)
#             hoja = wb2[nombre_hoja]

#             # LIMPIEZA COMPLETA SOLO PARA NUEVA HOJA
#             filas_a_limpiar = [2,3,4,7,8,9,12,13,14,17,18,19,22,23,24,27,28,31,32,33,34,37,38,39,40]
#             for fila in filas_a_limpiar:
#                 for col in range(2, 35):  # Columnas B a AH (2-34)
#                     try:
#                         celda = hoja.cell(row=fila, column=col)
#                         if not isinstance(celda, openpyxl.cell.cell.MergedCell):
#                             celda.value = ""
#                     except Exception as e:
#                         messagebox.showwarning("Advertencia", f"Error limpiando celda {fila},{col}: {str(e)}")
#                         continue

#             # Obtener número de días en el mes
#             dias_mes = dias_en_mes(mes, anio)
#             # Escribir fechas en todas las columnas para nueva hoja, solo hasta el último día del mes
#             for col in range(2, 2 + dias_mes):  # Columnas B hasta el último día del mes
#                 dia_mes = col - 1  # La columna 2 es día 1, columna 3 es día 2, etc.
#                 fecha_formato = f"{dia_mes:02d}/{MESES_NUM[mes]:02d}/{anio}"
#                 for fila_fecha in [2,7,12,17,22,27,31,37]:
#                     hoja.cell(row=fila_fecha, column=col, value=fecha_formato)

#             # Configuración inicial completa para nueva hoja
#             for col in range(2, 2 + dias_mes):  # Columnas B hasta el último día del mes
#                 letra = openpyxl.utils.get_column_letter(col)

#                 # Fórmulas básicas
#                 hoja.cell(row=40, column=col, value=f"=IFERROR({letra}34/{letra}28, 0)").number_format = '0.00%'
#                 hoja.cell(row=40, column=col).font = Font(color='000000')
#                 hoja.cell(row=34, column=col, value=f"=IFERROR(AVERAGE({letra}32:{letra}33), 0)").number_format = '0.00%'

#                 # Limpieza y colocación de otros datos
#                 celda = hoja.cell(row=32, column=34, value="R%")
#                 celda.font = Font(bold=True)
#                 celda.alignment = Alignment(horizontal='center', vertical='center')
#                 celda = hoja.cell(row=38, column=34, value="IR%")
#                 celda.font = Font(bold=True)
#                 celda.alignment = Alignment(horizontal='center', vertical='center')

#                 # Limpiar celdas específicas
#                 for fila in [49,50,51]:
#                     for col_limpiar in [27,28,29,30]:
#                         hoja.cell(row=fila, column=col_limpiar, value=" ")

#                 # Fórmulas especiales
#                 hoja.cell(row=23, column=34, value="=(AH3*AH13+AH8*AH18)/(AH3+AH8)")
#                 hoja.cell(row=24, column=34, value="=(AH4*AH14+AH9*AH19)/(AH4+AH9)")
#                 hoja.cell(row=28, column=34, value="=(AH23*(AH3+AH8)+AH24*(AH4+AH9))/(AH3+AH4+AH8+AH9)")
#                 # Fórmulas por columna
#                 hoja.cell(row=23, column=col, 
#                          value=f"=IFERROR(({letra}3*{letra}13+{letra}8*{letra}18)/({letra}3+{letra}8), 0)")
#                 hoja.cell(row=24, column=col, 
#                          value=f"=IFERROR(({letra}4*{letra}14+{letra}9*{letra}19)/({letra}4+{letra}9), 0)")
#                 hoja.cell(row=28, column=col, 
#                          value=f"=IFERROR(({letra}23*({letra}3+{letra}8)+{letra}24*({letra}4+{letra}9))/({letra}3+{letra}4+{letra}8+{letra}9), 0)")
#                 hoja.cell(row=38, column=col, 
#                          value=f"=IFERROR({letra}32/{letra}23, 0)")
#                 hoja.cell(row=38, column=col).font = Font(color='000000')
#                 hoja.cell(row=39, column=col, 
#                          value=f"=IFERROR({letra}33/{letra}24, 0)")
#                 hoja.cell(row=39, column=col).font = Font(color='000000')

#             # Configuración columna AH (columna 34)
#             hoja.cell(row=2, column=34, value=int(anio))
#             for fila in [3,4,8,9]:
#                 hoja.cell(row=fila, column=34, value=f"=SUM(B{fila}:AG{fila})")
#             for fila in [13,14,18,19]:
#                 hoja.cell(row=fila, column=34, value=" ")
#             hoja.cell(row=39, column=34, value="=AH33/AH28").number_format = '0.00%'
#             hoja.cell(row=40, column=34, value="=AH34/AH28").number_format = '0.00%'
#             # Guardar cambios
#             try:
#                 wb2.save(RUTA_ENTRADA)
#                 # Forzar actualización de fórmulas
#                 try:
#                     pythoncom.CoInitialize()
#                     excel = win32.Dispatch("Excel.Application")
#                     excel.Visible = False
#                     excel.DisplayAlerts = False
#                     excel_wb = excel.Workbooks.Open(RUTA_ENTRADA)
#                     excel.CalculateFull()
#                     excel_wb.Save()
#                     excel_wb.Close()
#                     excel.Quit()
#                     pythoncom.CoUninitialize()
#                 except Exception as com_error:
#                     messagebox.showwarning("Advertencia", f"Error al actualizar fórmulas: {str(com_error)}")
#                     pythoncom.CoUninitialize()

#             except Exception as save_error:
#                 messagebox.showerror("Error", f"No se pudo guardar el archivo: {str(save_error)}")
#                 return False
#             finally:
#                 wb2.close()

#         return True

#     except Exception as main_error:
#         messagebox.showerror(
#             "Error crítico",
#             f"No se pudo completar la operación:\n{str(main_error)}\n\n"
#             "Verifique que el archivo no esté abierto en Excel."
#         )
#         return False




def preparar_hoja_mes(mes, dia, anio):
    """Versión que usa exclusivamente OpenPyXL para todas las operaciones"""
    nombre_hoja = f"IR {mes} {anio}"
    col_dia = dia + 1  # columna B es 2, día 1 → columna 2
    
    try:
        # 1. Cargar el libro completo
        wb = openpyxl.load_workbook(RUTA_ENTRADA)
        
        # 2. Verificar si la hoja ya existe
        if nombre_hoja in wb.sheetnames:
            hoja = wb[nombre_hoja]
            
            # Verificar si el día ya tiene datos
            celdas_clave = [
                hoja.cell(row=3, column=col_dia).value,  # Torno 1 regular
                hoja.cell(row=4, column=col_dia).value,  # Torno 1 podado
                hoja.cell(row=8, column=col_dia).value,  # Torno 2 regular
                hoja.cell(row=9, column=col_dia).value   # Torno 2 podado
            ]
            
            if any(cell is not None and str(cell).strip() != "" for cell in celdas_clave):
                print(f"El día {dia} ya tiene datos en {nombre_hoja}")
            wb.close()
            return True
        
        # 3. Crear nueva hoja (copiando formato de hoja anterior)
        hojas_ir = [name for name in wb.sheetnames if name.startswith("IR ") and len(name.split()) == 3]
        
        def mes_a_num(nombre_hoja):
            try:
                _, mes_str, anio_str = nombre_hoja.split()
                return int(anio_str) * 12 + MESES_NUM[mes_str]
            except:
                return -1
        
        # Encontrar la hoja anterior más reciente
        hojas_ordenadas = sorted(hojas_ir, key=mes_a_num)
        hoja_anterior = None
        total_nueva = int(anio) * 12 + MESES_NUM[mes]
        
        for h in reversed(hojas_ordenadas):
            if mes_a_num(h) < total_nueva:
                hoja_anterior = h
                break
        
        if not hoja_anterior:
            print("No se encontró hoja anterior para copiar")
            wb.close()
            return False
        
        # Copiar hoja (OpenPyXL copia contenido pero no gráficos/completamente)
        source = wb[hoja_anterior]
        new_sheet = wb.copy_worksheet(source)
        new_sheet.title = nombre_hoja
        
        # 4. Limpieza de datos en la nueva hoja
        filas_a_limpiar = [2,3,4,7,8,9,12,13,14,17,18,19,22,23,24,27,28,31,32,33,34,37,38,39,40]
        for fila in filas_a_limpiar:
            for col in range(2, 35):  # Columnas B a AH (2-34)
                try:
                    celda = new_sheet.cell(row=fila, column=col)
                    if not isinstance(celda, openpyxl.cell.cell.MergedCell):
                        celda.value = None
                except:
                    continue
        
        # 5. Configuración inicial de la nueva hoja
        dias_mes = dias_en_mes(mes, anio)
        
        # Escribir fechas
        for col in range(2, 2 + dias_mes):
            dia_mes = col - 1
            fecha_formato = f"{dia_mes:02d}/{MESES_NUM[mes]:02d}/{anio}"
            for fila_fecha in [2,7,12,17,22,27,31,37]:
                new_sheet.cell(row=fila_fecha, column=col, value=fecha_formato)
        
        # Configurar fórmulas
        for col in range(2, 2 + dias_mes):
            letra = openpyxl.utils.get_column_letter(col)
            
            # Fórmulas básicas
            new_sheet.cell(row=40, column=col, value=f"=IFERROR({letra}34/{letra}28, 0)")
            new_sheet.cell(row=34, column=col, value=f"=IFERROR(AVERAGE({letra}32:{letra}33), 0)")
            
            # Fórmulas especiales
            new_sheet.cell(row=23, column=col, 
                         value=f"=IFERROR(({letra}3*{letra}13+{letra}8*{letra}18)/({letra}3+{letra}8), 0)")
            new_sheet.cell(row=24, column=col, 
                         value=f"=IFERROR(({letra}4*{letra}14+{letra}9*{letra}19)/({letra}4+{letra}9), 0)")
            new_sheet.cell(row=28, column=col, 
                         value=f"=IFERROR(({letra}23*({letra}3+{letra}8)+{letra}24*({letra}4+{letra}9))/({letra}3+{letra}4+{letra}8+{letra}9), 0)")
        
        # Configuración columna AH (34)
        new_sheet.cell(row=2, column=34, value=int(anio))
        for fila in [3,4,8,9]:
            new_sheet.cell(row=fila, column=34, value=f"=SUM(B{fila}:AG{fila})")
        
        # 6. Guardar cambios
        wb.save(RUTA_ENTRADA)
        wb.close()
        return True
        
    except Exception as e:
        print(f"Error crítico: {str(e)}")
        try:
            wb.close()
        except:
            pass
        return False




# def rotar_etiquetas_graficos(ruta_archivo, nombre_hoja):
#     """Versión final que maneja todos los casos de gráficos y versiones de Excel"""
#     pythoncom.CoInitialize()
#     excel = wb = None
#     resultado = False
#     try:
#         # 1. Configuración robusta de Excel
#         excel = win32.Dispatch("Excel.Application")
#         excel.Visible = False
#         excel.DisplayAlerts = False
#         excel.ScreenUpdating = False
#         # 2. Abrir archivo con manejo de errores
#         try:
#             wb = excel.Workbooks.Open(os.path.abspath(ruta_archivo))
#         except Exception as e:
#             messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{str(e)}")
#             return False
#         # 3. Verificar hoja
#         try:
#             if nombre_hoja not in [s.Name for s in wb.Sheets]:
#                 messagebox.showerror("Error", f"Hoja '{nombre_hoja}' no encontrada")
#                 return False
#             sheet = wb.Sheets(nombre_hoja)
#         except Exception as e:
#             messagebox.showerror("Error", f"Error al acceder a la hoja:\n{str(e)}")
#             return False

#         # 4. Procesamiento mejorado de gráficos
#         graficos = sheet.ChartObjects()
#         total_graficos = graficos.Count
#         rotados = 0
#         problemas = []

#         for i, chart_obj in enumerate(graficos, 1):
#             try:
#                 chart = chart_obj.Chart

#                 # Método universal para diferentes versiones de Excel
#                 try:
#                     # Versión compatible con todas las versiones de Excel
#                     x_axis = chart.Axes(1)  # 1 = xlCategory

#                     # Verificación alternativa para HasTickLabels
#                     try:
#                         if hasattr(x_axis, 'HasTickLabels') and not x_axis.HasTickLabels:
#                             problemas.append(f"Gráfico {i}: No tiene etiquetas visibles")
#                             continue
#                     except:
#                         # Si falla HasTickLabels, verificamos de otra manera
#                         try:
#                             x_axis.TickLabels  # Intento acceder a TickLabels directamente
#                         except:
#                             problemas.append(f"Gráfico {i}: No se pueden acceder a las etiquetas")
#                             continue
#                     # Rotación segura
#                     try:
#                         x_axis.TickLabels.Orientation = 45
#                         rotados += 1
#                     except Exception as e:
#                         problemas.append(f"Gráfico {i}: Error al rotar - {str(e)}")

#                 except Exception as e:
#                     problemas.append(f"Gráfico {i}: Tipo no soportado - {str(e)}")

#             except Exception as e:
#                 problemas.append(f"Gráfico {i}: Error grave - {str(e)}")
#         # 5. Manejo de resultados mejorado
#         if total_graficos == 0:
#             messagebox.showinfo("Información", "No se encontraron gráficos en la hoja")
#             resultado = True
#         # elif rotados > 0:
#         #     mensaje = f"Éxito: Se rotaron {rotados} de {total_graficos} gráficos"
#             if problemas:
#                 mensaje += "\n\nProblemas encontrados:\n" + "\n".join(problemas[:3])
#             messagebox.showinfo("Resultado", mensaje)
#             resultado = True
#         # else:
#         #     mensaje = (
#         #         "No se pudo rotar ningún gráfico.\n\n"
#         #         "Causas probables:\n"
#         #         "1. Versión de Excel no compatible\n"
#         #         "2. Gráficos de tipo especial (3D, combinados, etc.)\n"
#         #         "3. El archivo está protegido o dañado\n\n"
#         #         "Solución recomendada:\n"
#         #         "1. Guardar los gráficos como imágenes\n"
#         #         "2. Crear nuevos gráficos estándar\n"
#         #         "3. Actualizar Microsoft Excel"
#         #     )
#         #     messagebox.showerror("Error", mensaje)
#         #     resultado = False
#         # 6. Guardar cambios si hubo éxito
#         if rotados > 0:
#             try:
#                 wb.Save()
#             except Exception as e:
#                 messagebox.showwarning("Advertencia", f"Cambios realizados pero no guardados:\n{str(e)}")

#     except Exception as e:
#         messagebox.showerror("Error crítico", f"Error inesperado:\n{str(e)}")
#     finally:
#         # 7. Limpieza garantizada
#         try:
#             if wb:
#                 wb.Close(SaveChanges=False)
#         except:
#             pass
#         try:
#             if excel:
#                 excel.Quit()
#                 del excel
#         except:
#             pass
#         pythoncom.CoUninitialize()
#     return resultado


# def rotar_etiquetas_graficos(ruta_archivo, nombre_hoja):
#     """Versión que no cierra otros Excels y trabaja en segundo plano."""
#     pythoncom.CoInitialize()
#     excel = wb = None
#     resultado = False
#     try:
#         # 1. Configuración de Excel sin afectar otros archivos
#         excel = win32.Dispatch("Excel.Application")
#         excel.Visible = False  # <<< Asegurar que nunca sea visible
#         excel.DisplayAlerts = False
#         excel.ScreenUpdating = False

#         # 2. Abrir archivo (manejo de errores)
#         try:
#             wb = excel.Workbooks.Open(os.path.abspath(ruta_archivo))
#         except Exception as e:
#             messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{str(e)}")
#             return False

#         # 3. Procesar gráficos (código original)
#         sheet = wb.Sheets(nombre_hoja)
#         graficos = sheet.ChartObjects()
#         rotados = 0
#         problemas = []

#         for chart_obj in graficos:
#             try:
#                 chart = chart_obj.Chart
#                 x_axis = chart.Axes(1)  # Eje X
#                 if hasattr(x_axis, 'TickLabels'):
#                     x_axis.TickLabels.Orientation = 45
#                     rotados += 1
#             except Exception as e:
#                 problemas.append(str(e))

#         # 4. Guardar cambios si hubo rotaciones
#         if rotados > 0:
#             try:
#                 wb.Save()
#             except Exception as e:
#                 messagebox.showwarning("Advertencia", f"No se guardaron cambios:\n{str(e)}")

#         resultado = (rotados > 0)

#     except Exception as e:
#         messagebox.showerror("Error crítico", f"Error inesperado:\n{str(e)}")
#     finally:
#         # 5. Limpieza SIN cerrar otros Excels
#         try:
#             if wb:
#                 wb.Close(SaveChanges=False)  # Cierra solo este archivo
#         except:
#             pass
#         try:
#             if excel:
#                 excel.ScreenUpdating = True  # Restaurar configuración
#                 # No ejecutar excel.Quit() para no cerrar otros archivos
#         except:
#             pass
#         pythoncom.CoUninitialize()
#     return resultado


def rotar_etiquetas_graficos(ruta_archivo, nombre_hoja):
    print("entrando a rotar etiquetas ")
    print(ruta_archivo)
    print(nombre_hoja)
    resultado = 1
    return resultado


def escribir_division(hoja):
    """Versión con validación de existencia de hoja y columnas"""
    try:
        if not hoja:
            raise ValueError("El objeto hoja no es válido")

        # Columnas a procesar
        for col_num in range(2, 33):  # B (2) a AF (32)
            letra = openpyxl.utils.get_column_letter(col_num)

            # Verificar que las celdas referenciadas existen
            if hoja.max_row >= 40 and hoja.max_row >= 34 and hoja.max_row >= 28:
                formula = f"=SI.ERROR({letra}34/{letra}28, 0)"
                hoja.cell(row=40, column=col_num).value = formula
                hoja.cell(row=40, column=col_num).number_format = '0.00'
    
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo escribir fórmulas:\n{str(e)}")
        return False
    return True

def dias_en_mes(mes, anio):
    """Devuelve el número de días en un mes, considerando años bisiestos para febrero"""
    if mes == "Febrero":
        # Año bisiesto si es divisible por 4, pero no por 100, a menos que también sea divisible por 400
        if (anio % 4 == 0 and anio % 100 != 0) or (anio % 400 == 0):
            return 29
        return 28
    meses_31_dias = ["Enero", "Marzo", "Mayo", "Julio", "Agosto", "Octubre", "Diciembre"]
    return 31 if mes in meses_31_dias else 30

ventana = tk.Tk()
ventana.title("Ingresar datos")
entrada_texto = tk.Text(ventana, width=100, height=30)
entrada_texto.pack(padx=10, pady=10)
tk.Button(ventana, text="Procesar", command=obtener_datos).pack(pady=10)
ventana.mainloop()
