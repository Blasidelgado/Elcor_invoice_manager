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
        # Find concepts previous lines
        if 'Código Producto / Servicio' in line:
            index = rows.index(line) + 2 # Index + 1 corresponds to a strange pdfplumber behavior

    # Start searching for concepts
    while True:
        product = rows[index].strip()
        # Common text lines after concepts in AFIP invoices
        if 'Importe Otros Tributos' in product or 'Subtotal' in product:
            break
        product = product[:25] if len(product) > 25 else product
        products.append(product)

    # Take required data
    try:
        date = temp_dict['Fecha de Emisión']
        date = datetime.strptime(date, '%d/%m/%Y')
        # If owner company is the emittor, take the receiver, else take the emittor
        company = temp_dict['Razón Social'] if temp_dict['Razón Social'] != 'GRAINING SA' else temp_dict['Apellido y Nombre / Razón Social']
        concepts = '/'.join(products)
        # Clean total in order to convert it into a number
        total = temp_dict['Importe Total'].removeprefix('$').strip().replace(',', '.')
        # If owner company is the emittor, take the amount as positive, else account as negative balance
        total = float(total) if temp_dict['Razón Social'] == 'GRAINING SA' else float(total) * (-1)
    except:
        return None
    
    info = {'date': date, 'company': company, 'concepts': concepts, 'total': total}

    return info


def parse_bank(page):
    # Extract text from page
    text = page.extract_text(x_tolerance=3, y_tolerance=3)

    # Create a list using line breaks and declare a new dict
    rows = text.split('\n')
    temp_dict = {}

    # Start iterating over rows to add info into dict
    for line in rows:
        # Format text lines into dicts
        if ':' in line:
            # Split line into key-value pairs using ':' as a separator
            pairs = line.split(':')

            # Add first pair to the info_dict
            key, value = pairs[0].strip(), pairs[1].strip()
            temp_dict[key] = value

    # Take required information
    date = temp_dict.get('Fecha de Pago')
    try:
        date = datetime.strptime(date, '%d/%m/%Y')
        company = temp_dict.get('Beneficiario')
    except:
        return None

    info = {'date': date, 'company': company}

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

    return f'New data appended in row {last_row + 1}\n'

def manipulate_invoice(filepath, invoice, suf, type, data):
    # Get date from file
    date = data.get('date')
    day, year, month = date.day, date.year, date.month

    # Format month to be always 2 digits
    month = str(month) if month > 9 else "0"+str(month)

    # Create new filename and helper variable for duplicates
    filename = f"{filepath}/{type}/{year}/{month}/{data.get('company')} {day}-{month}-{year}"
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
