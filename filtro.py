from bs4 import BeautifulSoup
import os
import sys

# def format_data_row(row, is_first_in_category=False):
#     cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
#     if not cells or len(cells) < 5:
#         return None
    
#     # Extraer los valores de las celdas
#     values = []
#     for cell in cells:
#         # Manejar celdas con tablas anidadas (para distribución)
#         if cell.find('table'):
#             table = cell.find('table')
#             dist_value = table.find_all('td')[-1].get_text(strip=True)
#             values.append(dist_value.split('&nbsp;')[0].strip())
#         else:
#             text = cell.get_text(strip=True)
#             values.append(text if text not in ('', '&nbsp;') else ' ')
    
#     # Determinar si es una fila de categoría (PODADO/REGULAR) o de datos
#     if len(values) >= 3 and (values[1].upper() in ['PODADO', 'REGULAR'] or values[0].upper() in ['PODADO', 'REGULAR']):
#         # Es una fila de categoría
#         tipo_madera = values[0] if (values[0] != ' ' and is_first_in_category) else ' '
#         tipo = values[1]
#         diametro_clase = values[2]
#         trozos = values[3]
#         distribucion = values[4] if len(values) > 4 else ' '
        
#         # Construir primera línea
#         first_line = f"{tipo_madera} {tipo} {diametro_clase} {trozos}  {distribucion}"
        
#         # Construir segunda línea con los valores restantes
#         second_line = ' ' + ' '.join(values[5:])
        
#         return f"{first_line} \n{second_line} \n"
#     elif len(values) >= 3 and values[0] == '*' and values[1] == '*':
#         # Es una fila de subtotal (* * ...)
#         trozos = values[3]
#         distribucion = values[4] if len(values) > 4 else ' '
#         return f"* * ... {trozos}  {distribucion} \n {' '.join(values[5:])} \n"
#     else:
#         # Es una fila de datos normal
#         first_part = ' '.join(values[:4])
#         second_part = ' ' + ' '.join(values[4:])
#         return f" {first_part} \n{second_part} \n"

def format_data_row(row, is_first_in_category=False):
    cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
    if not cells or len(cells) < 5:
        return None
    
    # Extraer los valores de las celdas
    values = []
    distribution = ' '  # Valor por defecto para distribución
    for i, cell in enumerate(cells):
        # Manejar celdas con tablas anidadas (para distribución)
        if cell.find('table'):
            table = cell.find('table')
            dist_value = table.find_all('td')[-1].get_text(strip=True)
            distribution = dist_value.split('&nbsp;')[0].strip()
            values.append(' ')  # Marcador de posición
        else:
            text = cell.get_text(strip=True)
            values.append(text if text not in ('', '&nbsp;') else ' ')
    
    # Determinar si es una fila de categoría (PODADO/REGULAR) o de datos
    if len(values) >= 3 and (values[1].upper() in ['PODADO', 'REGULAR'] or values[0].upper() in ['PODADO', 'REGULAR']):
        # Es una fila de categoría
        tipo_madera = values[0] if (values[0] != ' ' and is_first_in_category) else ' '
        tipo = values[1]
        diametro_clase = values[2]
        trozos = values[3]
        
        # Construir primera línea
        first_line = f"{tipo_madera} {tipo} {diametro_clase} {trozos}  {distribution}"
        
        # Construir segunda línea con los valores restantes (comenzando desde el índice 4)
        second_line = ' ' + ' '.join(values[4:])
        
        return f"{first_line} \n{second_line} \n"
    elif len(values) >= 3 and values[0] == '*' and values[1] == '*':
        # Es una fila de subtotal (* * ...)
        trozos = values[3]
        return f"* * ... {trozos}  {distribution} \n {' '.join(values[4:])} \n"
    else:
        # Es una fila de datos normal
        first_part = ' '.join(values[:4])
        second_part = ' ' + ' '.join(values[4:])
        return f" {first_part} \n{second_part} \n"

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
    is_first_in_category = True
    
    # Procesar todas las filas excepto la última (TOTAL)
    for row in diameter_table.find_all('tr')[1:-1]:
        # Verificar si es un cambio de categoría
        first_cell = row.find('td', class_=['RWReport', 'RWReportSUM'])
        if first_cell:
            cell_text = first_cell.get_text(strip=True)
            if cell_text and cell_text.upper() in ['PODADO', 'REGULAR']:
                current_category = cell_text
                is_first_in_category = True
        
        formatted_row = format_data_row(row, is_first_in_category)
        if formatted_row:
            output_lines.append(formatted_row)
            is_first_in_category = False
    
    # Procesar fila de subtotal (* * ...)
    subtotal_row = diameter_table.find_all('tr')[-2]
    if '*' in subtotal_row.get_text():
        formatted_subtotal = format_data_row(subtotal_row, False)
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
