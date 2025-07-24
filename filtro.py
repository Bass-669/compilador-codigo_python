from bs4 import BeautifulSoup
import os
import sys

def format_data_row(row, is_category=False):
    cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
    if not cells or len(cells) < 19:
        return None

    # Procesar primera línea
    if is_category:
        # Para líneas de categoría (RADIATA PODADO/REGULAR)
        tipo_madera = cells[0].get_text(strip=True)
        tipo = cells[1].get_text(strip=True)
        diametro = cells[2].get_text(strip=True)
        trozos = cells[3].get_text(strip=True)
        
        # Manejar distribución (puede tener tabla anidada)
        distribucion = ' '
        if len(cells) > 5:
            dist_cell = cells[5]
            if dist_cell.find('table'):
                dist_text = dist_cell.find('table').find_all('td')[-1].get_text(strip=True)
                distribucion = dist_text.split('&nbsp;')[0].strip()
            else:
                distribucion = dist_cell.get_text(strip=True) if dist_cell.get_text(strip=True) not in ('', '&nbsp;') else ' '
        
        first_line = f"{tipo_madera} {tipo} {diametro} {trozos}  {distribucion}\n"
    elif cells[0].get_text(strip=True) == '*':
        # Para línea de subtotal (* * ...)
        trozos = cells[3].get_text(strip=True)
        distribucion = ' '
        if len(cells) > 5:
            dist_cell = cells[5]
            if dist_cell.find('table'):
                dist_text = dist_cell.find('table').find_all('td')[-1].get_text(strip=True)
                distribucion = dist_text.split('&nbsp;')[0].strip()
            else:
                distribucion = dist_cell.get_text(strip=True) if dist_cell.get_text(strip=True) not in ('', '&nbsp;') else ' '
        
        first_line = f"* * ... {trozos}  {distribucion}\n"
    else:
        # Para líneas normales
        tipo = cells[1].get_text(strip=True)
        diametro = cells[2].get_text(strip=True)
        trozos = cells[3].get_text(strip=True)
        
        # Manejar distribución
        distribucion = ' '
        if len(cells) > 5:
            dist_cell = cells[5]
            if dist_cell.find('table'):
                dist_text = dist_cell.find('table').find_all('td')[-1].get_text(strip=True)
                distribucion = dist_text.split('&nbsp;')[0].strip()
            else:
                distribucion = dist_cell.get_text(strip=True) if dist_cell.get_text(strip=True) not in ('', '&nbsp;') else ' '
        
        first_line = f" {tipo} {diametro} {trozos}  {distribucion}\n"

    # Procesar segunda línea (19 valores)
    second_line_parts = []
    indices = [4] + list(range(6, 19))  # Columnas requeridas: 4,6,7,...,18
    
    for i in indices:
        if i < len(cells):
            cell = cells[i]
            text = cell.get_text(strip=True)
            if text not in ('', '&nbsp;'):
                second_line_parts.append(text)
            else:
                second_line_parts.append(' ')
        else:
            second_line_parts.append(' ')
    
    # Asegurar 19 valores exactos
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
    
    # Procesar filas (excepto cabecera y total)
    for row in rows[1:-1]:
        first_cell = row.find('td', class_=['RWReport', 'RWReportSUM'])
        if first_cell:
            cell_text = first_cell.get_text(strip=True)
            if cell_text and cell_text not in ('', '&nbsp;', '*'):
                if "PODADO" in cell_text.upper() or "REGULAR" in cell_text.upper():
                    current_category = cell_text
                    # Procesar como línea de categoría
                    formatted_row = format_data_row(row, is_category=True)
                    if formatted_row:
                        output_lines.append(formatted_row)
                    continue
        
        # Procesar como línea normal
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
    
    # Escribir archivo de salida
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

