from bs4 import BeautifulSoup
import os
import sys

def format_data_row(row, is_category_row=False):
    cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
    if not cells or len(cells) < 19:
        return None

    # --- First line formatting ---
    # Part 0: Wood type (HTML[0])
    part0 = cells[0].get_text(strip=True) if cells[0].get_text(strip=True) not in ('', '&nbsp;') else ' '
    
    # Handle special case for subtotal row (* * ...)
    if part0 == '*':
        subtotal_part3 = cells[3].get_text(strip=True) if len(cells) > 3 and cells[3].get_text(strip=True) not in ('', '&nbsp;') else ' '
        subtotal_part5 = cells[5].get_text(strip=True) if len(cells) > 5 and cells[5].get_text(strip=True) not in ('', '&nbsp;') else ' '
        first_line = f"* * ... {subtotal_part3}  {subtotal_part5}".strip()
    else:
        # Normal row formatting
        # Part 1-2: Diameter range (HTML[1]...HTML[2])
        min_d = cells[1].get_text(strip=True) if len(cells) > 1 else ''
        max_d = cells[2].get_text(strip=True) if len(cells) > 2 else ''
        
        if min_d not in ('', '&nbsp;') and max_d not in ('', '&nbsp;'):
            part1_2 = f"{min_d}...{max_d}"
        elif min_d not in ('', '&nbsp;'):
            part1_2 = min_d
        elif max_d not in ('', '&nbsp;'):
            part1_2 = max_d
        else:
            part1_2 = ''

        # Part 3: Distribution (HTML[3])
        part3 = ' '
        if len(cells) > 3:
            cell3 = cells[3]
            if cell3.find('table'):
                nested_td = cell3.find('table').find('td', class_='RWReport')
                part3 = nested_td.get_text(strip=True) if nested_td and nested_td.get_text(strip=True) not in ('', '&nbsp;') else ' '
            else:
                part3 = cell3.get_text(strip=True) if cell3.get_text(strip=True) not in ('', '&nbsp;') else ' '

        # Part 5: Dry weight % (HTML[5])
        part5 = cells[5].get_text(strip=True) if len(cells) > 5 and cells[5].get_text(strip=True) not in ('', '&nbsp;') else ' '

        # For category rows (PODADO/REGULAR), don't include part0 in the first line
        if is_category_row:
            first_line = f"{part1_2} {part3}  {part5}".strip()
        else:
            first_line = f"{part0} {part1_2} {part3}  {part5}".strip()

    # --- Second line formatting ---
    # We need all 19 values (indices 4,6-18)
    line2_parts = []
    indices_second_line = [4] + list(range(6, 19))
    
    for i in indices_second_line:
        text = ' '
        if i < len(cells):
            cell = cells[i]
            if cell.find('table'):
                continue
            cell_text = cell.get_text(strip=True)
            if cell_text not in ('', '&nbsp;'):
                text = cell_text
        line2_parts.append(text)
    
    # Ensure we have exactly 19 values (4 + 6-18)
    if len(line2_parts) < 19:
        line2_parts.extend([' '] * (19 - len(line2_parts)))
    
    second_line = ' ' + ' '.join(line2_parts).strip()

    # Combine the lines
    formatted = f"{first_line}\n{second_line}\n"
    
    if not first_line.strip() and not second_line.strip():
        return None
    
    return formatted

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
    
    # Process all rows except header
    rows_to_process = diameter_table.find_all('tr')[1:]
    
    # Check if last row is a grand total to be excluded
    if len(rows_to_process) > 1:
        last_row_cells = rows_to_process[-1].find_all('td', class_=['RWReport', 'RWReportSUM'])
        if last_row_cells and last_row_cells[0].get_text(strip=True) not in ('', '&nbsp;', '*'):
            rows_to_process = rows_to_process[:-1]
    
    for row in rows_to_process:
        # Handle category changes (PODADO/REGULAR)
        first_cell = row.find('td', class_=['RWReport', 'RWReportSUM'])
        if first_cell:
            cell_text = first_cell.get_text(strip=True)
            if cell_text and cell_text not in ('', '&nbsp;', '*'):
                if "PODADO" in cell_text.upper() or "REGULAR" in cell_text.upper():
                    current_category = cell_text
                    output_lines.append(f"{current_category}\n")
                    # Format the category row without repeating the category name
                    formatted_row = format_data_row(row, is_category_row=True)
                else:
                    formatted_row = format_data_row(row)
            else:
                formatted_row = format_data_row(row)
        else:
            formatted_row = format_data_row(row)
        
        if formatted_row:
            # Add indentation for non-category rows
            if current_category and not any(x in formatted_row for x in ['RADIATA PODADO', 'RADIATA REGULAR', '* * ...']):
                formatted_row = ' ' + formatted_row
            output_lines.append(formatted_row)
    
    # Write to output file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(''.join(output_lines).strip())

    return True

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

