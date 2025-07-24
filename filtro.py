from bs4 import BeautifulSoup
import os
import sys

def format_data_row(row, is_category=False):
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
        first_line = f"* * ... {subtotal_part3}  {subtotal_part5}\n"
    else:
        # Normal row formatting
        # Part 1-2: Diameter range (HTML[2])
        part1_2 = cells[2].get_text(strip=True) if len(cells) > 2 and cells[2].get_text(strip=True) not in ('', '&nbsp;') else ''

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

        if is_category:
            first_line = f"{part0} {part1_2} {part3}  {part5}\n"
        else:
            first_line = f" {part0} {part1_2} {part3}  {part5}\n"

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
    
    # Ensure we have exactly 19 values
    if len(line2_parts) < 19:
        line2_parts.extend([' '] * (19 - len(line2_parts)))
    
    second_line = ' ' + ' '.join(line2_parts) + ' \n'

    # Combine the lines
    formatted = f"{first_line}{second_line}"
    
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
    
    # Process all rows except header and total
    rows = diameter_table.find_all('tr')
    if len(rows) < 2:
        return False
        
    rows_to_process = rows[1:-1]  # Skip header and total row
    
    for row in rows_to_process:
        # Handle category changes (PODADO/REGULAR)
        first_cell = row.find('td', class_=['RWReport', 'RWReportSUM'])
        if first_cell:
            cell_text = first_cell.get_text(strip=True)
            if cell_text and cell_text not in ('', '&nbsp;', '*'):
                if "PODADO" in cell_text.upper() or "REGULAR" in cell_text.upper():
                    current_category = cell_text
                    # Format the category row
                    formatted_row = format_data_row(row, is_category=True)
                    if formatted_row:
                        output_lines.append(formatted_row)
                    continue
        
        formatted_row = format_data_row(row)
        if formatted_row:
            output_lines.append(formatted_row)
    
    # Process subtotal row (second to last)
    if len(rows) > 1:
        subtotal_row = rows[-2]
        if '*' in subtotal_row.get_text():
            formatted_subtotal = format_data_row(subtotal_row)
            if formatted_subtotal:
                output_lines.append(formatted_subtotal)
    
    # Write to output file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(''.join(output_lines))

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

