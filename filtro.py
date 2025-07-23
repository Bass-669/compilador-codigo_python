from bs4 import BeautifulSoup
import os
import sys

def format_data_row(row, is_subtotal=False):
    cells = row.find_all('td', class_='RWReport')
    if not cells or len(cells) < 5:
        return None

    # Procesar primera línea (cabecera)
    first_line_parts = []
    for i, cell in enumerate(cells[:4]):
        text = cell.get_text(strip=True)
        if i == 0 and text in ('', '&nbsp;', '*'):
            first_line_parts.append('*')
        else:
            first_line_parts.append(text if text not in ('&nbsp;', '') else ' ')
    
    # Para líneas normales (no subtotal)
    if not is_subtotal:
        first_line = ' '.join(first_line_parts)
        second_line = ' '.join([cell.get_text(strip=True) for cell in cells[4:] 
                             if cell.get_text(strip=True) not in ('&nbsp;', '')])
        return f"{first_line} \n {second_line} \n"
    # Para línea de subtotal
    else:
        first_line = ' '.join(first_line_parts)
        second_line = ' '.join([cell.get_text(strip=True) for cell in cells[4:] 
                             if cell.get_text(strip=True) not in ('&nbsp;', '')])
        return f"{first_line} \n {second_line} \n"

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

    for row in diameter_table.find_all('tr')[1:-1]:
        first_cell = row.find('td', class_='RWReport')
        if not first_cell:
            continue
            
        cell_text = first_cell.get_text(strip=True)
        
        # Detectar cambio de categoría (RADIATA PODADO/REGULAR)
        if cell_text and cell_text not in ('', '&nbsp;', '*'):
            if "PODADO" in cell_text.upper() or "REGULAR" in cell_text.upper():
                current_category = cell_text
                output_lines.append(f"{current_category} \n")
                continue
        
        formatted_row = format_data_row(row)
        if formatted_row:
            output_lines.append(formatted_row)

    # Procesar fila de subtotal
    subtotal_row = diameter_table.find_all('tr')[-2]
    if '*' in subtotal_row.get_text():
        formatted_subtotal = format_data_row(subtotal_row, is_subtotal=True)
        if formatted_subtotal:
            output_lines.append(formatted_subtotal)

    # Escribir al archivo con el formato específico
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(''.join(output_lines).replace('  ', ' ').replace(' \n ', '\n'))

    return True

def main():
    if len(sys.argv) < 2:
        print("Uso: filtro.exe <carpeta_con_html>")
        sys.exit(1)

    carpeta_reportes = sys.argv[1]
    if not os.path.isdir(carpeta_reportes):
        print(f"Error: La carpeta {carpeta_reportes} no existe.")
        sys.exit(1)

    carpeta_datos = os.path.join(carpeta_reportes, 'datos')
    os.makedirs(carpeta_datos, exist_ok=True)

    html_files = [f for f in os.listdir(carpeta_reportes) if f.lower().endswith('.html')]
    if not html_files:
        print(f"No se encontraron archivos .html en {carpeta_reportes}")
        sys.exit(0)

    archivos_procesados = 0
    archivos_sin_tabla = 0

    for html_file in html_files:
        nombre_base = os.path.splitext(html_file)[0]
        ruta_salida = os.path.join(carpeta_datos, nombre_base + '.txt')

        # Omitir si ya existe el archivo de salida
        if os.path.exists(ruta_salida):
            print(f"Ya existe: {ruta_salida}, omitiendo...")
            continue

        ruta_html = os.path.join(carpeta_reportes, html_file)

        if process_html_file(ruta_html, ruta_salida):
            print(f"Procesado: {html_file} ➜ {ruta_salida}")
            archivos_procesados += 1
        else:
            archivos_sin_tabla += 1

    print(f"\n✔ Procesamiento completado.")
    print(f"   Archivos procesados: {archivos_procesados}")
    print(f"   Archivos sin tabla 'Diámetro': {archivos_sin_tabla}")

if __name__ == "__main__":
    main()
