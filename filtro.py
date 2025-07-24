from bs4 import BeautifulSoup
import os
import sys

def format_data_row(row):
    cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
    if not cells or len(cells) < 19:  # Asegurarnos de tener todas las columnas necesarias
        return None
    
    # Procesar primera línea (columnas 0, 2, 3, 5)
    first_line_parts = []
    
    # Columna 0: Tipo madera (solo para filas de categoría)
    part0 = cells[0].get_text(strip=True) if cells[0].get_text(strip=True) not in ('', '&nbsp;') else ''
    
    # Columna 2: Diámetro clase JAS
    part2 = cells[2].get_text(strip=True) if len(cells) > 2 and cells[2].get_text(strip=True) not in ('', '&nbsp;') else ''
    
    # Columna 3: Trozos
    part3 = cells[3].get_text(strip=True) if len(cells) > 3 and cells[3].get_text(strip=True) not in ('', '&nbsp;') else ''
    
    # Columna 5: Distribución (puede tener tabla anidada)
    part5 = ' '
    if len(cells) > 5:
        dist_cell = cells[5]
        if dist_cell.find('table'):
            dist_text = dist_cell.find('table').find_all('td')[-1].get_text(strip=True)
            part5 = dist_text.split('&nbsp;')[0].strip()
        else:
            part5 = dist_cell.get_text(strip=True) if dist_cell.get_text(strip=True) not in ('', '&nbsp;') else ' '
    
    # Manejar caso especial de subtotal (* * ...)
    if part0 == '*':
        first_line = f"* * ... {part3}  {part5}\n"
    else:
        if part0:  # Es una fila de categoría
            first_line = f"{part0} {part2} {part3}  {part5}\n"
        else:  # Fila normal
            first_line = f" {part2} {part3}  {part5}\n"
    
    # Procesar segunda línea (columnas 4, 6-18)
    second_line_parts = []
    # Columna 4: Diámetro máximo
    if len(cells) > 4:
        part4 = cells[4].get_text(strip=True) if cells[4].get_text(strip=True) not in ('', '&nbsp;') else ' '
        second_line_parts.append(part4)
    
    # Columnas 6-18
    for i in range(6, 19):
        if i < len(cells):
            cell = cells[i]
            if cell.find('table'):
                continue
            text = cell.get_text(strip=True)
            if text not in ('', '&nbsp;'):
                second_line_parts.append(text)
            else:
                second_line_parts.append(' ')
        else:
            second_line_parts.append(' ')
    
    # Asegurar que tenemos exactamente 19 valores
    if len(second_line_parts) < 19:
        second_line_parts.extend([' '] * (19 - len(second_line_parts)))
    
    second_line = ' ' + ' '.join(second_line_parts) + ' \n'
    
    return first_line + second_line

def process_html_file(html_file, output_file):
    with open(html_file, 'r', encoding='utf-8') as file:
        html = file.read()
    
    soup = BeautifulSoup(html, 'html.parser')
    h4_diametro = soup.find('h4', string='Diámetro')
    if not h4_diametro:
        print(f"Advertencia: No se encontró la tabla de diámetros en {html_file}")
        return False
    
    diameter_table = h4_diametro.find_next('table')
    output_lines = []
    current_category = None
    
    rows = diameter_table.find_all('tr')
    if len(rows) < 2:
        return False
    
    # Procesar todas las filas excepto la cabecera y la última (total general)
    for row in rows[1:-1]:
        # Manejar cambio de categoría
        first_cell = row.find('td', class_=['RWReport', 'RWReportSUM'])
        if first_cell:
            cell_text = first_cell.get_text(strip=True)
            if cell_text and cell_text not in ('', '&nbsp;', '*'):
                if "PODADO" in cell_text.upper() or "REGULAR" in cell_text.upper():
                    current_category = cell_text
                    output_lines.append(f"{current_category} \n")
        
        formatted_row = format_data_row(row)
        if formatted_row:
            output_lines.append(formatted_row)
    
    # Procesar fila de subtotal (penúltima fila)
    if len(rows) > 1:
        subtotal_row = rows[-2]
        if '*' in subtotal_row.get_text():
            formatted_subtotal = format_data_row(subtotal_row)
            if formatted_subtotal:
                output_lines.append(formatted_subtotal)
    
    # Escribir en el archivo de salida
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(''.join(output_lines))

def main():
    if len(sys.argv) < 2:
        print("Uso: python script.py <carpeta_con_html>")
        sys.exit(1)

    carpeta_reportes = sys.argv[1]
    if not os.path.isdir(carpeta_reportes):
        print(f"Error: La carpeta {carpeta_reportes} no existe.")
        sys.exit(1)

    carpeta_datos = os.path.join(carpeta_reportes, 'datos')
    os.makedirs(carpeta_datos, exist_ok=True)

    archivos_html = [f for f in os.listdir(carpeta_reportes) if f.lower().endswith('.html')]
    if not archivos_html:
        print(f"No se encontraron archivos .html en {carpeta_reportes}")
        sys.exit(0)

    total_procesados = 0
    total_sin_tabla = 0

    for archivo_html in archivos_html:
        nombre_base = os.path.splitext(archivo_html)[0]
        ruta_salida = os.path.join(carpeta_datos, nombre_base + '.txt')

        if os.path.exists(ruta_salida):
            print(f"Ya existe: {ruta_salida}, omitiendo...")
            continue

        ruta_html = os.path.join(carpeta_reportes, archivo_html)

        if process_html_file(ruta_html, ruta_salida):
            print(f"Procesado: {archivo_html} ➜ {ruta_salida}")
            total_procesados += 1
        else:
            total_sin_tabla += 1

    print(f"\n✔ Procesamiento completado.")
    print(f"   Archivos procesados: {total_procesados}")
    print(f"   Archivos sin tabla 'Diámetro': {total_sin_tabla}")

if __name__ == "__main__":
    main()


# from bs4 import BeautifulSoup
# import os
# import sys

# def format_data_row(row):
#     cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
#     if not cells or len(cells) < 5:
#         return None
    
#     # Procesar primera línea (primeras 4 columnas)
#     first_line_parts = []
#     for i, cell in enumerate(cells[:4]):
#         text = cell.get_text(strip=True)
#         if text in ('', '&nbsp;'):
#             first_line_parts.append(' ')
#         elif text == '*':
#             first_line_parts.append('*')
#         else:
#             first_line_parts.append(text)
    
#     # Manejar distribución (4ta columna) que puede tener tabla anidada
#     if len(cells) > 3:
#         dist_cell = cells[3]
#         if dist_cell.find('table'):
#             # Extraer solo el valor numérico de la tabla anidada
#             dist_text = dist_cell.find('table').find_all('td')[-1].get_text(strip=True)
#             first_line_parts[3] = dist_text.split('&nbsp;')[0].strip()
    
#     # Construir primera línea con 2 espacios antes del porcentaje
#     if len(first_line_parts) > 3:
#         first_line = ' '.join(first_line_parts[:3]) + '  ' + first_line_parts[3]
#     else:
#         first_line = ' '.join(first_line_parts)
    
#     # Procesar segunda línea (resto de columnas)
#     second_line_parts = []
#     for cell in cells[4:]:
#         # Omitir celdas que contienen tablas (ya procesadas)
#         if cell.find('table'):
#             continue
#         text = cell.get_text(strip=True)
#         if text not in ('', '&nbsp;'):
#             second_line_parts.append(text)
    
#     # Formatear como en el ejemplo deseado
#     formatted = f"{first_line} \n {' '.join(second_line_parts)} \n"
#     return formatted

# def process_html_file(html_file, output_file):
#     with open(html_file, 'r', encoding='utf-8') as file:
#         html = file.read()
    
#     soup = BeautifulSoup(html, 'html.parser')
#     h4_diametro = soup.find('h4', string='Diámetro')
#     if not h4_diametro:
#         print(f"Advertencia: No se encontró la tabla de diámetros en {html_file}")
#         return False
    
#     diameter_table = h4_diametro.find_next('table')
#     output_lines = []
#     current_category = None
    
#     for row in diameter_table.find_all('tr')[1:-1]:  # Ignorar cabecera y total
#         # Manejar cambio de categoría
#         first_cell = row.find('td', class_=['RWReport', 'RWReportSUM'])
#         if first_cell:
#             cell_text = first_cell.get_text(strip=True)
#             if cell_text and cell_text not in ('', '&nbsp;', '*'):
#                 if "PODADO" in cell_text.upper() or "REGULAR" in cell_text.upper():
#                     current_category = cell_text
#                     output_lines.append(f"{current_category} \n")
        
#         formatted_row = format_data_row(row)
#         if formatted_row:
#             # Para filas que no son categorías, agregar sangría
#             if not any(x in formatted_row for x in ['RADIATA PODADO', 'RADIATA REGULAR']):
#                 formatted_row = ' ' + formatted_row.lstrip()
#             output_lines.append(formatted_row)
    
#     # Procesar fila de subtotal (* * ...)
#     subtotal_row = diameter_table.find_all('tr')[-2]
#     if '*' in subtotal_row.get_text():
#         formatted_subtotal = format_data_row(subtotal_row)
#         if formatted_subtotal:
#             output_lines.append(formatted_subtotal)
    
#     # Escribir en el archivo de salida
#     with open(output_file, 'w', encoding='utf-8') as f:
#         f.write(''.join(output_lines))

# def main():
#     if len(sys.argv) < 2:
#         print("Uso: python script.py <carpeta_con_html>")
#         sys.exit(1)

#     carpeta_reportes = sys.argv[1]
#     if not os.path.isdir(carpeta_reportes):
#         print(f"Error: La carpeta {carpeta_reportes} no existe.")
#         sys.exit(1)

#     carpeta_datos = os.path.join(carpeta_reportes, 'datos')
#     os.makedirs(carpeta_datos, exist_ok=True)

#     archivos_html = [f for f in os.listdir(carpeta_reportes) if f.lower().endswith('.html')]
#     if not archivos_html:
#         print(f"No se encontraron archivos .html en {carpeta_reportes}")
#         sys.exit(0)

#     total_procesados = 0
#     total_sin_tabla = 0

#     for archivo_html in archivos_html:
#         nombre_base = os.path.splitext(archivo_html)[0]
#         ruta_salida = os.path.join(carpeta_datos, nombre_base + '.txt')

#         if os.path.exists(ruta_salida):
#             print(f"Ya existe: {ruta_salida}, omitiendo...")
#             continue

#         ruta_html = os.path.join(carpeta_reportes, archivo_html)

#         if process_html_file(ruta_html, ruta_salida):
#             print(f"Procesado: {archivo_html} ➜ {ruta_salida}")
#             total_procesados += 1
#         else:
#             total_sin_tabla += 1

#     print(f"\n✔ Procesamiento completado.")
#     print(f"   Archivos procesados: {total_procesados}")
#     print(f"   Archivos sin tabla 'Diámetro': {total_sin_tabla}")

# if __name__ == "__main__":
#     main()

