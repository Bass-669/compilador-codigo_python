from bs4 import BeautifulSoup
import os
import sys

def format_data_row(row):
    cells = row.find_all('td', class_='RWReport')
    if not cells or len(cells) < 5:
        return None

    first_line = []
    for i, cell in enumerate(cells[:4]):
        text = cell.get_text(strip=True)
        if i == 0 and text in ('', '&nbsp;', '*'):
            first_line.append(' ')
        else:
            first_line.append(text if text not in ('&nbsp;', '') else ' ')

    second_line = []
    for cell in cells[4:]:
        text = cell.get_text(strip=True)
        if cell.find('table'):
            continue
        second_line.append(text if text not in ('&nbsp;', '') else ' ')

    formatted = ' '.join(first_line) + ' \n ' + ' '.join(second_line) + ' \n'
    return formatted.replace('  ', ' ')

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

    for row in diameter_table.find_all('tr')[1:-1]:
        formatted_row = format_data_row(row)
        if formatted_row:
            output_lines.append(formatted_row)

    subtotal_row = diameter_table.find_all('tr')[-2]
    if '*' in subtotal_row.get_text():
        output_lines.append(format_data_row(subtotal_row))

    with open(output_file, 'a', encoding='utf-8') as f:
        f.write(''.join(output_lines))

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
