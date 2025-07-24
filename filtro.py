def format_data_row(row, is_first_in_category=False):
    cells = row.find_all('td', class_=['RWReport', 'RWReportSUM'])
    if not cells or len(cells) < 5:
        return None
    
    # Extraer los valores de las celdas
    values = []
    for cell in cells:
        # Manejar celdas con tablas anidadas (para distribución)
        if cell.find('table'):
            table = cell.find('table')
            dist_value = table.find_all('td')[-1].get_text(strip=True)
            values.append(dist_value.split('&nbsp;')[0].strip())
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
        distribucion = values[4] if len(values) > 4 else ' '
        
        # Construir primera línea
        first_line = f"{tipo_madera} {tipo} {diametro_clase} {trozos}  {distribucion}"
        
        # Construir segunda línea con los valores restantes (omitir distribución)
        second_line = ' ' + ' '.join(values[5:])
        
        return f"{first_line} \n{second_line} \n"
    elif len(values) >= 3 and values[0] == '*' and values[1] == '*':
        # Es una fila de subtotal (* * ...)
        trozos = values[3]
        distribucion = values[4] if len(values) > 4 else ' '
        return f"* * ... {trozos}  {distribucion} \n {' '.join(values[5:])} \n"
    else:
        # Es una fila de datos normal
        first_part = ' '.join(values[:4])
        second_part = ' ' + ' '.join(values[5:])  # Omitir el valor de distribución (values[4])
        return f" {first_part} \n{second_part} \n"
