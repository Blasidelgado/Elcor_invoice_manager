import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from operator import itemgetter
from pathlib import Path

import dotenv
import pdfplumber
from openpyxl import load_workbook

from console import Console
from helpers import manipulate_invoice, parse_afip, parse_bank, update_worksheet
from Veryfi import Veryfi


class ElcorInvoiceManager:
    def __init__(self):
        self.root = tk.Tk()
        self.selected_directory = None
        self.selected_xlsx = None
        self.setup_ui()
        self.verify = Veryfi(self)
        dotenv.load_dotenv() # Load .env keys
        # Call setup verify to set up keys for further api requests
        self.client = self.verify.setup_veryfi(dotenv.get_key('.env', 'CLIENT_ID'), 
            dotenv.get_key('.env', 'CLIENT_SECRET'), dotenv.get_key('.env', 'USERNAME'), 
            dotenv.get_key('.env', 'API_KEY'))
        self.root.mainloop()
        
    def setup_ui(self):
        # Create a grid with two columns and two rows
        self.root.title("Elcor Invoice Manager")
        self.root.columnconfigure(0, weight=3)
        self.root.columnconfigure(1, weight=1)
        self.root.columnconfigure(2, weight=1)
        self.root.columnconfigure(3, weight=1)
        self.root.columnconfigure(4, weight=2)
        self.root.rowconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self.root.rowconfigure(2, weight=1)
        self.root.rowconfigure(3, weight=5)

        # Add widgets to the grid
        tk.Label(self.root, text="Please select a work directory", font=(16)).grid(row=0, column=0, columnspan=4, sticky="nsew")
        tk.Button(self.root, text="Select directory", font=(16), command=self.select_directory).grid(row=0, column=4)
        tk.Label(self.root, text="Please select a xlsx file to insert info", font=(16)).grid(row=1, column=0, columnspan=4, sticky="nsew")
        tk.Button(self.root, text="Select xlsx", font=(16), command=self.select_xlsx).grid(row=1, column=4)
        tk.Button(self.root, text="START", font=(16), command=self.start_processing).grid(row=2, column=1, columnspan=2, sticky="nsew")

        self.console = Console(self.root)
        self.console.grid(row=3, column=0, columnspan=5)

    def select_directory(self):
        self.selected_directory = filedialog.askdirectory()
        self.console.write(f'Selected directory: {self.selected_directory}\n')

    def select_xlsx(self):
        self.selected_xlsx = filedialog.askopenfilename()
        self.console.write(f'Selected xlsl file: {self.selected_xlsx}\n')

    def start_processing(self):
        self.console.flush()
        if self.selected_directory is None:
            self.console.write("Please select a directory\n")
            return

        if self.selected_xlsx is None:
            self.console.write("Please select an xlsx file\n")
            return
        
        # Instancing a path class to the selected directory
        p = Path(self.selected_directory)
        if not p.is_dir():
            self.console.write("Please select a valid folder\n")
            return

        # Instancing a path class to the xlsx file
        xlsx = Path(self.selected_xlsx)
        if not xlsx.suffix == '.xlsx':
            self.console.write("Please select a valid xlsx file\n")
            return

        # List all available files to scan
        extensions = ['pdf', 'jpg', 'jpeg', 'png']
        invoices = [file_path for ext in extensions for file_path in p.glob(f'*.{ext}') if file_path.is_file()]

        # Load the workbook and select worksheet
        wb = load_workbook(filename = xlsx)
        ws = wb.active

        data_list = []

        # Start scanning files
        for invoice in invoices:
            filename = invoice.name
            self.console.write(f'Processing file: {filename}\n')
            suffix = invoice.suffix

            if suffix == '.pdf':

                # Read pdf with pdfplumber
                reader = pdfplumber.open(invoice)

                # Take only the first pdf page
                page = reader.pages[0]

                # Read document metadata for further processing
                if reader.metadata.get('Creator') == 'AFIP': # Administración Federal de Ingresos Públicos   
                    data = parse_afip(page)

                    # Close file
                    reader.close()

                    if not data:
                        self.console.write(f'Error parsing file {filename}\n')
                        continue

                    # Append data to data_list
                    data_list.append(data)


                    # Manipulate file in respective inner directory
                    self.console.write(manipulate_invoice(p, invoice, suffix, 'facturas', data))
                
                # PDF engines used by our bank to generate pdfs
                elif 'Skia/PDF' in reader.metadata.get('Producer'):
                    data = parse_bank(page)

                    # Close file
                    reader.close()

                    if not data:
                        self.console.write(f'Error parsing file {filename}\n')
                        continue

                    # Manipulate file in respective inner directory
                    self.console.write(manipulate_invoice(p, invoice, suffix, 'comprobantes de pago', data))

                else: # PDF file is sent to our trusted API
                    data = self.verify.parse_veryfi(invoice, self.client)

                    # Close file
                    reader.close()

                    if not self.verify.check_response(filename, data):
                        continue

                    # Append data to data_list
                    data_list.append(data)

                    # Manipulate file in respective inner directory
                    self.console.write(manipulate_invoice(p, invoice, suffix, 'facturas', data))

                    
            else: # File is an image that will be sent to our trusted API
                data = self.verify.parse_veryfi(invoice, self.client)
                
                # Close file
                reader.close()

                if not self.verify.check_response(filename, data):
                    continue
                
                # Append data to data_list
                data_list.append(data)

                # Manipulate file in respective inner directory
                self.console.write(manipulate_invoice(p, invoice, suffix, 'facturas', data))

        # Sort the data list by date
        data_list = sorted(data_list, key=itemgetter('date'))

        # Iterate data from datalist and append each into a new row
        for data in data_list:
            # Format date to string in order to append to xlsx
            data['date'] = datetime.strftime(data['date'], '%d/%m/%Y')

            # Create list to append data correctly
            appendable_data = [data['date'], data['company'], data['concepts'], data['total']]
            
            # Append data and inform the user
            self.console.write(update_worksheet(ws, appendable_data))

            # Save worksheet    
            wb.save(xlsx)
        
        self.console.write("Exiting Program Succesfully")
        return
    
                