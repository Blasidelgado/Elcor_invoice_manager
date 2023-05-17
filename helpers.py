import os
from datetime import datetime

def parse_afip(page):
    # Extract text from page
    text = page.extract_text(x_tolerance=2, y_tolerance=1)
    # Create a list using line breaks and declare a new dict
    rows = text.split('\n')
    temp_dict = {}

    products = []
    # Start iterating over rows to add info into dict
    for line in rows:
        # Format text lines into dicts
        if ':' in line:
            # Check for conflictive line
            if 'Apellido y Nombre / Razón Social' in line:
                index = line.find('Apellido y Nombre / Razón Social')
                line2 = line[index:]
                line = line[:index]
                pairs = line2.split(':')
                key, value = pairs[0].strip(), pairs[1].strip()
                temp_dict[key] = value

            # Split line into key-value pairs using ':' as a separator
            pairs = line.split(':')

            # Add first pair to the info_dict
            key, value = pairs[0].strip(), pairs[1].strip()
            temp_dict[key] = value
        # Find indexes where product line-items are probably listed
        if 'Código Producto / Servicio' in line:
            start_index = rows.index(line) + 2 # Index + 1 corresponds to a row that is not intended to be so
        if 'Importe Otros Tributos' in line:
            end_index = rows.index(line)

    # Search between indexes for line items
    for i in range(start_index, end_index):
        product = rows[i].strip()
        product = product[:25] if len(product) > 25 else product
        products.append(product)

    # Take required data
    date = temp_dict['Fecha de Emisión']
    # If owner company is the emittor, take the receiver, else take the emittor
    company = temp_dict['Razón Social'] if temp_dict['Razón Social'] != 'GRAINING SA' else temp_dict['Apellido y Nombre / Razón Social']
    concepts = '/'.join(products)
    # Clean total in order to convert it into a number
    total = temp_dict['Importe Total'].removeprefix('$').strip().replace(',', '.')
    # If owner company is the emittor, take the amount as positive, else account as negative balance
    total = float(total) if temp_dict['Razón Social'] == 'GRAINING SA' else float(total) * (-1)
    
    info = {'date': date, 'company': company, 'concepts': concepts, 'total': total}

    return info

    
def parse_veryfi(page, client):
    results = client.process_document(page)
    unformatted_date = results['date']
    date_object = datetime.strptime(unformatted_date, '%Y-%m-%d %H:%M:%S')
    date = date_object.strftime('%d/%m/%Y')
    company = results['vendor']['name']
    items = results['line_items']
    concepts = []
    for item in items:
        item_desc = item.get('description')
        concepts.append(item_desc) if len(item_desc) < 25 else concepts.append(item_desc[:25])
        
    concepts = '/'.join(concepts)

    total = float(results['total'])

    info = {'date': date, 'company': company, 'concepts': concepts, 'total': total}

    return info


def update_worksheet(worksheet, data):
    last_row = worksheet.max_row
    worksheet.insert_rows(last_row + 1)

    for col, value in enumerate(data, start=1):
        cell = worksheet.cell(row=last_row+1, column=col)
        prev_cell = worksheet.cell(row=last_row, column=col)
        cell.value = value
        cell.font = prev_cell.font.copy()
        cell.border = prev_cell.border.copy()
        cell.fill = prev_cell.fill.copy()
        cell.number_format = prev_cell.number_format
        cell.protection = prev_cell.protection.copy()
        cell.alignment = prev_cell.alignment.copy()


def manipulate_invoice(filepath, invoice, suf, type, data):
    # Get date from file
    date = datetime.strptime(data[0], '%d/%m/%Y')
    year, month = date.year, date.month

    # Create new filename and helper variable for duplicates
    filename = f"{filepath}/{type}/{year}/{month}/{data[1]} {data[0].replace('/', '-')}"
    count = 1

    if os.path.exists(f"{filepath}/{type}/{year}/{month}/"):
        try:
            os.rename(invoice, filename + suf)
        except FileExistsError:
            while os.path.exists(filename + suf):
                filename = filename.removesuffix(f' {count - 1}')
                filename = f"{filename} {count}"
                count += 1
        
            os.rename(invoice, filename + suf)

    else:
        os.makedirs(f"{filepath}/{type}/{year}/{month}")
        os.rename(invoice, filename + suf)

    return f'Renamed {filename + suf} and moved succesfully\n'
    